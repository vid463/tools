const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('agentApi', {
  run: (command) => ipcRenderer.invoke('agent:run', command),
  fillExcel: () => ipcRenderer.invoke('agent:fill-excel'),
  cancel: () => ipcRenderer.invoke('agent:cancel'),
  getFillHistory: () => ipcRenderer.invoke('fill-history:list'),
  getAppVersion: () => ipcRenderer.invoke('app:get-version'),
  checkUpdate: () => ipcRenderer.invoke('app:check-update'),
  openPath: (filePath) => ipcRenderer.invoke('shell:openPath', filePath),
  onProgress: (callback) => {
    const listener = (_event, data) => callback(data);
    ipcRenderer.on('agent:progress', listener);
    return () => ipcRenderer.removeListener('agent:progress', listener);
  },
  openExternal: (url) => ipcRenderer.invoke('shell:openExternal', url),
});
