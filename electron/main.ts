import { app, BrowserWindow, ipcMain, dialog } from 'electron';
import path from 'path';
import fs from 'fs';
import { formatFsError } from '../src/core/fsError';
import { updateRecentFilesList, removeRecentFileItem } from '../src/core/recentFiles';

let mainWindow: BrowserWindow | null = null;
let isForcedClose = false;

// Extract .xlsx / .csv path from command line arguments
function extractFilePathFromArgs(args: string[]): string | null {
  for (const arg of args) {
    if (!arg || arg.startsWith('-')) continue;
    const lower = arg.toLowerCase();
    if (lower.endsWith('.xlsx') || lower.endsWith('.csv')) {
      try {
        if (fs.existsSync(arg)) {
          return path.resolve(arg);
        }
      } catch {}
    }
  }
  return null;
}

// Single Instance Lock
const gotTheLock = app.requestSingleInstanceLock();

if (!gotTheLock) {
  app.quit();
} else {
  app.on('second-instance', (_event, argv) => {
    // When a second instance attempts to run, focus current window and open file if passed
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();

      const filePath = extractFilePathFromArgs(argv);
      if (filePath) {
        mainWindow.webContents.send('app:openExternalFile', filePath);
      }
    }
  });
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 800,
    minHeight: 600,
    title: 'SheetEasy',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });

  const devServerUrl = process.env.VITE_DEV_SERVER_URL;
  if (devServerUrl) {
    mainWindow.loadURL(devServerUrl);
  } else {
    mainWindow.loadFile(path.join(__dirname, '../dist/index.html'));
  }

  // Intercept close to ask user about unsaved changes
  mainWindow.on('close', (e) => {
    if (!isForcedClose) {
      e.preventDefault();
      mainWindow?.webContents.send('app:requestClose');
    }
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  // When initial web contents finishes loading, check if file was provided via command line (Open with)
  mainWindow.webContents.on('did-finish-load', () => {
    const initialFile = extractFilePathFromArgs(process.argv);
    if (initialFile && mainWindow) {
      mainWindow.webContents.send('app:openExternalFile', initialFile);
    }
  });
}

// Recent Files Paths & Helpers
function getRecentFilesPath(): string {
  return path.join(app.getPath('userData'), 'recent-files.json');
}

function getRecentFilesSync(): string[] {
  try {
    const p = getRecentFilesPath();
    if (fs.existsSync(p)) {
      const content = fs.readFileSync(p, 'utf8');
      const parsed = JSON.parse(content);
      if (Array.isArray(parsed)) return parsed;
    }
  } catch {}
  return [];
}

function saveRecentFilesSync(list: string[]) {
  try {
    const p = getRecentFilesPath();
    const dir = path.dirname(p);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    const tempPath = path.join(dir, `.recent-files.tmp-${Date.now()}`);
    fs.writeFileSync(tempPath, JSON.stringify(list, null, 2), 'utf8');
    fs.renameSync(tempPath, p);
  } catch {}
}

// Auto-Recovery Directory & Helpers
function getRecoveryDir(): string {
  const dir = path.join(app.getPath('userData'), 'recovery');
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
  return dir;
}

// IPC Handlers - File Dialogs
ipcMain.handle('dialog:openFile', async () => {
  if (!mainWindow) return null;
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'เปิดไฟล์',
    properties: ['openFile'],
    filters: [
      { name: 'ตารางคำนวณ (*.xlsx, *.csv)', extensions: ['xlsx', 'csv'] },
      { name: 'Excel Worksheet (*.xlsx)', extensions: ['xlsx'] },
      { name: 'CSV File (*.csv)', extensions: ['csv'] },
      { name: 'ทุกไฟล์ (*.*)', extensions: ['*'] },
    ],
  });

  if (result.canceled || result.filePaths.length === 0) {
    return { canceled: true };
  }

  const filePath = result.filePaths[0];
  try {
    const buffer = fs.readFileSync(filePath);
    return {
      canceled: false,
      filePath,
      data: Array.from(buffer),
    };
  } catch (err: any) {
    return {
      canceled: false,
      filePath,
      error: formatFsError(err),
    };
  }
});

ipcMain.handle('dialog:saveFile', async (_event, defaultPath?: string) => {
  if (!mainWindow) return null;
  const result = await dialog.showSaveDialog(mainWindow, {
    title: 'บันทึกเป็น',
    defaultPath: defaultPath || 'ตารางรายรับรายจ่าย.xlsx',
    filters: [
      { name: 'Excel Worksheet (*.xlsx)', extensions: ['xlsx'] },
      { name: 'CSV File (*.csv)', extensions: ['csv'] },
    ],
  });

  if (result.canceled || !result.filePath) {
    return { canceled: true };
  }

  return {
    canceled: false,
    filePath: result.filePath,
  };
});

// Atomic File Writing with Specific Thai Error Mapping (SPEC 6.2)
ipcMain.handle('fs:writeFile', async (_event, filePath: string, bufferData: number[]) => {
  const dir = path.dirname(filePath);
  const tempPath = path.join(dir, `.${path.basename(filePath)}.tmp-${Date.now()}`);
  try {
    const buffer = Buffer.from(bufferData);
    // Write temporary file in the same directory first
    fs.writeFileSync(tempPath, buffer);
    // Rename atomically over target file
    fs.renameSync(tempPath, filePath);

    return { success: true };
  } catch (err: any) {
    // Clean up temporary file if write or rename failed, leaving original file completely intact
    try {
      if (fs.existsSync(tempPath)) {
        fs.unlinkSync(tempPath);
      }
    } catch {}
    return { success: false, error: formatFsError(err) };
  }
});

ipcMain.handle('fs:readFile', async (_event, filePath: string) => {
  try {
    const buffer = fs.readFileSync(filePath);
    return { success: true, data: Array.from(buffer) };
  } catch (err: any) {
    return { success: false, error: formatFsError(err) };
  }
});

ipcMain.handle('fs:checkExists', async (_event, filePath: string) => {
  try {
    return fs.existsSync(filePath);
  } catch {
    return false;
  }
});

// Recent Files IPC Handlers
ipcMain.handle('recent:get', async () => {
  return getRecentFilesSync();
});

ipcMain.handle('recent:add', async (_event, filePath: string) => {
  const current = getRecentFilesSync();
  const updated = updateRecentFilesList(current, filePath);
  saveRecentFilesSync(updated);
  return updated;
});

ipcMain.handle('recent:remove', async (_event, filePath: string) => {
  const current = getRecentFilesSync();
  const updated = removeRecentFileItem(current, filePath);
  saveRecentFilesSync(updated);
  return updated;
});

// Auto-Recovery IPC Handlers (SPEC 6.3)
ipcMain.handle('recovery:save', async (_event, data: any) => {
  try {
    const dir = getRecoveryDir();
    const id = data.id || 'default-session';
    const tempPath = path.join(dir, `.${id}.tmp-${Date.now()}`);
    const finalPath = path.join(dir, `${id}.json`);

    fs.writeFileSync(tempPath, JSON.stringify(data), 'utf8');
    fs.renameSync(tempPath, finalPath);

    return { success: true };
  } catch (err: any) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('recovery:list', async () => {
  try {
    const dir = getRecoveryDir();
    const files = fs.readdirSync(dir);
    const result: any[] = [];

    for (const f of files) {
      if (f.endsWith('.json') && !f.startsWith('.')) {
        try {
          const content = fs.readFileSync(path.join(dir, f), 'utf8');
          const parsed = JSON.parse(content);
          if (parsed && parsed.workbook) {
            result.push(parsed);
          }
        } catch {}
      }
    }

    // Sort newest first
    result.sort((a, b) => (b.savedAt || 0) - (a.savedAt || 0));
    return result;
  } catch {
    return [];
  }
});

ipcMain.handle('recovery:delete', async (_event, id: string) => {
  try {
    const dir = getRecoveryDir();
    const target = path.join(dir, `${id}.json`);
    if (fs.existsSync(target)) {
      fs.unlinkSync(target);
    }
    return { success: true };
  } catch {
    return { success: false };
  }
});

ipcMain.handle('recovery:clearAll', async () => {
  try {
    const dir = getRecoveryDir();
    const files = fs.readdirSync(dir);
    for (const f of files) {
      try {
        fs.unlinkSync(path.join(dir, f));
      } catch {}
    }
    return { success: true };
  } catch {
    return { success: false };
  }
});

// Window title & force close
ipcMain.on('window:setTitle', (_event, title: string) => {
  if (mainWindow) {
    mainWindow.setTitle(title);
  }
});

ipcMain.on('window:forceClose', () => {
  isForcedClose = true;
  if (mainWindow) {
    mainWindow.close();
  }
});

// App Lifecycle
app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

// macOS open-file event
app.on('open-file', (event, filePath) => {
  event.preventDefault();
  if (mainWindow) {
    mainWindow.webContents.send('app:openExternalFile', filePath);
  }
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
