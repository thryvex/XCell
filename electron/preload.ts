import { contextBridge, ipcRenderer } from 'electron';

export interface RecoveryItem {
  id: string;
  originalPath: string | null;
  fileName: string;
  savedAt: number;
  workbook: any;
}

export interface ElectronAPI {
  openFile: () => Promise<{ canceled: boolean; filePath?: string; data?: number[]; error?: string } | null>;
  saveFile: (defaultPath?: string) => Promise<{ canceled: boolean; filePath?: string } | null>;
  exportCsvDialog: (defaultPath?: string) => Promise<{ canceled: boolean; filePath?: string } | null>;
  openCsvFile: () => Promise<{ canceled: boolean; filePath?: string; data?: number[]; error?: string } | null>;
  writeFile: (filePath: string, buffer: number[]) => Promise<{ success: boolean; error?: string }>;
  readFile: (filePath: string) => Promise<{ success: boolean; data?: number[]; error?: string }>;
  setTitle: (title: string) => void;
  // Recent files
  getRecentFiles: () => Promise<string[]>;
  addRecentFile: (filePath: string) => Promise<string[]>;
  removeRecentFile: (filePath: string) => Promise<string[]>;
  checkFileExists: (filePath: string) => Promise<boolean>;
  // Auto-recovery
  saveRecovery: (data: RecoveryItem) => Promise<{ success: boolean; error?: string }>;
  getRecoveryList: () => Promise<RecoveryItem[]>;
  deleteRecovery: (id: string) => Promise<{ success: boolean }>;
  clearAllRecovery: () => Promise<{ success: boolean }>;
  // Window & Life-cycle
  forceClose: () => void;
  onOpenExternalFile: (callback: (filePath: string) => void) => () => void;
  onRequestClose: (callback: () => void) => () => void;
  // Preferences
  getPreferences: () => Promise<any>;
  setPreferences: (prefs: any) => Promise<void>;
}

const electronAPI: ElectronAPI = {
  openFile: () => ipcRenderer.invoke('dialog:openFile'),
  saveFile: (defaultPath) => ipcRenderer.invoke('dialog:saveFile', defaultPath),
  exportCsvDialog: (defaultPath) => ipcRenderer.invoke('dialog:exportCsv', defaultPath),
  openCsvFile: () => ipcRenderer.invoke('dialog:openCsv'),
  writeFile: (filePath, buffer) => ipcRenderer.invoke('fs:writeFile', filePath, buffer),
  readFile: (filePath) => ipcRenderer.invoke('fs:readFile', filePath),
  setTitle: (title) => ipcRenderer.send('window:setTitle', title),
  getRecentFiles: () => ipcRenderer.invoke('recent:get'),
  addRecentFile: (filePath) => ipcRenderer.invoke('recent:add', filePath),
  removeRecentFile: (filePath) => ipcRenderer.invoke('recent:remove', filePath),
  checkFileExists: (filePath) => ipcRenderer.invoke('fs:checkExists', filePath),
  saveRecovery: (data) => ipcRenderer.invoke('recovery:save', data),
  getRecoveryList: () => ipcRenderer.invoke('recovery:list'),
  deleteRecovery: (id) => ipcRenderer.invoke('recovery:delete', id),
  clearAllRecovery: () => ipcRenderer.invoke('recovery:clearAll'),
  forceClose: () => ipcRenderer.send('window:forceClose'),
  onOpenExternalFile: (callback) => {
    const handler = (_event: any, filePath: string) => callback(filePath);
    ipcRenderer.on('app:openExternalFile', handler);
    return () => ipcRenderer.removeListener('app:openExternalFile', handler);
  },
  onRequestClose: (callback) => {
    const handler = () => callback();
    ipcRenderer.on('app:requestClose', handler);
    return () => ipcRenderer.removeListener('app:requestClose', handler);
  },
  getPreferences: () => ipcRenderer.invoke('prefs:get'),
  setPreferences: (prefs) => ipcRenderer.invoke('prefs:set', prefs),
};

contextBridge.exposeInMainWorld('electronAPI', electronAPI);
