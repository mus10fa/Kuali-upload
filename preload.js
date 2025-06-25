const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  openFile: () => ipcRenderer.invoke('open-file-dialog'),
  runAutomation: (filePath) => ipcRenderer.invoke('run-automation', filePath)
});
