import { app, BrowserWindow, ipcMain, dialog } from 'electron';
import path from 'path';
import fs from 'fs';

let mainWindow: BrowserWindow | null = null;

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

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// IPC Handlers
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
      error: err.message,
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

ipcMain.handle('fs:writeFile', async (_event, filePath: string, bufferData: number[]) => {
  try {
    const buffer = Buffer.from(bufferData);
    // Atomic write: write to temp file then rename
    const dir = path.dirname(filePath);
    const tempPath = path.join(dir, `.${path.basename(filePath)}.tmp-${Date.now()}`);

    fs.writeFileSync(tempPath, buffer);
    fs.renameSync(tempPath, filePath);

    return { success: true };
  } catch (err: any) {
    return { success: false, error: err.message };
  }
});

ipcMain.handle('fs:readFile', async (_event, filePath: string) => {
  try {
    const buffer = fs.readFileSync(filePath);
    return { success: true, data: Array.from(buffer) };
  } catch (err: any) {
    return { success: false, error: err.message };
  }
});

ipcMain.on('window:setTitle', (_event, title: string) => {
  if (mainWindow) {
    mainWindow.setTitle(title);
  }
});

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
