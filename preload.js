const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('agentApi', {
  run: (command) => ipcRenderer.invoke('agent:run', command),
  openExternal: (url) => ipcRenderer.invoke('shell:openExternal', url),
});
