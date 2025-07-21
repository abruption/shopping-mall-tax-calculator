import { app, BrowserWindow, ipcMain, dialog } from 'electron';
import * as path from 'path';
import { ExcelService } from '../services/ExcelService';
import { CalculationService } from '../services/CalculationService';
import { ProcessingOptions } from '../types';

let mainWindow: BrowserWindow | null = null;
const excelService = new ExcelService();
const calculationService = new CalculationService();

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  mainWindow.loadFile(path.join(__dirname, '..', 'renderer', 'index.html'));

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

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

// IPC Handlers
ipcMain.handle('select-files', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls'] },
      { name: 'All Files', extensions: ['*'] }
    ]
  });

  return result.filePaths;
});

ipcMain.handle('process-files', async (event, filePaths: string[], options: ProcessingOptions) => {
  try {
    const results = [];
    
    for (const filePath of filePaths) {
      const mallName = path.basename(filePath, path.extname(filePath));
      const data = await excelService.readExcelFile(filePath, options);
      const calculation = calculationService.calculateTotals({
        mallName,
        filePath,
        data
      });
      results.push(calculation);
    }
    
    return { success: true, data: results };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('export-results', async (event, results: any[], format: 'xlsx' | 'csv') => {
  const saveResult = await dialog.showSaveDialog({
    filters: format === 'xlsx' 
      ? [{ name: 'Excel Files', extensions: ['xlsx'] }]
      : [{ name: 'CSV Files', extensions: ['csv'] }]
  });

  if (!saveResult.canceled && saveResult.filePath) {
    await excelService.exportResults(results, saveResult.filePath, format);
    return { success: true, path: saveResult.filePath };
  }
  
  return { success: false };
});

ipcMain.handle('get-excel-columns', async (event, filePath: string) => {
  try {
    const columns = await excelService.getColumnHeaders(filePath);
    return { success: true, columns };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

