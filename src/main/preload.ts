import { contextBridge, ipcRenderer } from 'electron';

contextBridge.exposeInMainWorld('electronAPI', {
  selectFiles: () => ipcRenderer.invoke('select-files'),
  processFiles: (filePaths: string[], options: any) => 
    ipcRenderer.invoke('process-files', filePaths, options),
  exportResults: (results: any[], format: 'xlsx' | 'csv') => 
    ipcRenderer.invoke('export-results', results, format),
  getExcelColumns: (filePath: string) => ipcRenderer.invoke('get-excel-columns', filePath)
});