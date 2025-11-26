const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  pickDirectory: () => ipcRenderer.invoke('pickDirectory'),
  runOfficeToPdfMissing: (payload) => ipcRenderer.invoke('runOfficeToPdfMissing', payload),
  onProgress: (cb) => {
    const listener = (_event, data) => {
      try { cb && cb(data); } catch {}
    };
    ipcRenderer.on('office2pdf-progress', listener);
    return () => ipcRenderer.removeListener('office2pdf-progress', listener);
  }
});

