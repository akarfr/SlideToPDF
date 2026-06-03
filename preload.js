const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('api', {
  // Triggers
  selectFiles: () => ipcRenderer.send('select-files'),
  selectOutputDir: () => ipcRenderer.send('select-output-dir'),
  convertFiles: (files, outputDir) => ipcRenderer.send('convert-files', { files, outputDir }),

  // Callbacks
  onSelectedFiles: (callback) => {
    ipcRenderer.removeAllListeners('selected-files-reply');
    ipcRenderer.on('selected-files-reply', (event, files) => callback(files));
  },
  onSelectedOutputDir: (callback) => {
    ipcRenderer.removeAllListeners('selected-output-dir-reply');
    ipcRenderer.on('selected-output-dir-reply', (event, dir) => callback(dir));
  },
  onConversionResult: (callback) => {
    ipcRenderer.removeAllListeners('conversion-result');
    ipcRenderer.on('conversion-result', (event, message) => callback(message));
  },

  // Helper for webUtils drag/drop path extraction
  getPathForFile: (file) => {
    if (file instanceof File) {
      return webUtils.getPathForFile(file);
    }
    return null;
  }
});
