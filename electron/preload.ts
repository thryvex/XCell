import { contextBridge, ipcRenderer } from 'electron';

export interface ElectronAPI {
  openFile: () => Promise<{ canceled: boolean; filePath?: string; data?: number[] } | null>;
  saveFile: (defaultPath?: string) => Promise<{ canceled: boolean; filePath?: string } | null>;
  writeFile: (filePath: string, buffer: number[]) => Promise<{ success: boolean; error?: string }>;
  readFile: (filePath: string) => Promise<{ success: boolean; data?: number[]; error?: string }>;
  setTitle: (title: string) => void;
}

const electronAPI: ElectronAPI = {
  openFile: () => ipcRenderer.invoke('dialog:openFile'),
  saveFile: (defaultPath) => ipcRenderer.invoke('dialog:saveFile', defaultPath),
  writeFile: (filePath, buffer) => ipcRenderer.invoke('fs:writeFile', filePath, buffer),
  readFile: (filePath) => ipcRenderer.invoke('fs:readFile', filePath),
  setTitle: (title) => ipcRenderer.send('window:setTitle', title),
};

contextBridge.exposeInMainWorld('electronAPI', electronAPI);
