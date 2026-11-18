const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  loadPrice: () => ipcRenderer.invoke('load-price'),
  saveEstimate: (data) => ipcRenderer.invoke('save-estimate', data)
});
