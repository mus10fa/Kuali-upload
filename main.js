const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const path = require('path');
const { runAutomation } = require('./kuali_automation');

function createWindow() {
  const win = new BrowserWindow({
    width: 600,
    height: 300,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  win.loadFile('index.html');
}

app.whenReady().then(() => {
  createWindow();
});

// Listen for file-open requests from renderer
ipcMain.handle('open-file-dialog', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xlsm'] }],
    properties: ['openFile']
  });
  if (canceled) return null;
  return filePaths[0];
});

// Listen for 'run-automation' request
ipcMain.handle('run-automation', async (event, filePath) => {
  try {
    await runAutomation(filePath);
    return 'Success';
  } catch (e) {
    return `Error: ${e.message}`;
  }
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});