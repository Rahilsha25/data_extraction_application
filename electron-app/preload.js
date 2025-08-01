const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  selectFolder: () => ipcRenderer.invoke('select-folder'),
  runExtraction: (folderPath) => ipcRenderer.invoke('run-python', folderPath),
  loadResults: (folderPath) => ipcRenderer.invoke('load-results', folderPath),
  showResultsPage: (folderPath) => ipcRenderer.invoke('show-results-page', folderPath),  // ✅ ADD THIS
  goHome: () => ipcRenderer.send('show-home-page'),  // already exists
  onSetFolderPath: (callback) => ipcRenderer.on('set-folder-path', callback)
});
