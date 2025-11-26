const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const officeToPdfMissing = require('../src/commands/officeToPdfMissing');

function createWindow() {
  const win = new BrowserWindow({
    width: 900,
    height: 650,
    useContentSize: true,
    backgroundColor: '#f3f3f3',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js')
    }
  });
  win.loadFile(path.join(__dirname, 'index.html'));
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

ipcMain.handle('pickDirectory', async () => {
  const res = await dialog.showOpenDialog({ properties: ['openDirectory'] });
  if (res.canceled || !res.filePaths || res.filePaths.length === 0) return null;
  return res.filePaths[0];
});

ipcMain.handle('runOfficeToPdfMissing', async (event, args) => {
  const { dir, recursive, overwrite, dryRun } = args || {};
  const opts = { path: dir, recursive: !!recursive, overwrite: !!overwrite, dryRun: !!dryRun };
  opts.onProgress = (payload) => {
    event.sender.send('office2pdf-progress', payload);
  };
  try {
    const summary = await officeToPdfMissing(dir || '.', opts);
    return { ok: true, summary };
  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
});
