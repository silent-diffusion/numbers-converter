const path = require('path');
const { app, BrowserWindow, ipcMain, dialog, shell, Menu } = require('electron');
const { convertFile, previewFile, stopHelper } = require('./convert');
const settings = require('./settings');
const updater = require('./updater');

const AUTO_CHECK_INTERVAL = 6 * 60 * 60 * 1000; // 6 hours

let win = null;
let config = null;
let autoCheckTimer = null;
let pendingFiles = numbersArgs(process.argv);

function numbersArgs(argv) {
  return argv.slice(1).filter((a) => a.toLowerCase().endsWith('.numbers'));
}

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', (_e, argv) => {
    if (!win) return;
    if (win.isMinimized()) win.restore();
    win.focus();
    const files = numbersArgs(argv);
    if (files.length) win.webContents.send('files:opened', files);
  });
  app.whenReady().then(start);
}

function scheduleAutoUpdates() {
  clearInterval(autoCheckTimer);
  autoCheckTimer = null;
  if (!config.autoUpdate) return;
  updater.check();
  autoCheckTimer = setInterval(updater.check, AUTO_CHECK_INTERVAL);
}

function start() {
  config = settings.load();
  Menu.setApplicationMenu(null);

  win = new BrowserWindow({
    width: 960,
    height: 640,
    minWidth: 640,
    minHeight: 460,
    backgroundColor: '#f7f7f8',
    title: 'Numbers Converter',
    icon: path.join(__dirname, 'renderer', 'icon.png'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      sandbox: true
    }
  });
  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
  win.webContents.on('did-finish-load', () => {
    if (pendingFiles.length) win.webContents.send('files:opened', pendingFiles);
    pendingFiles = [];
  });

  updater.init((state) => win?.webContents.send('update:state', state));
  scheduleAutoUpdates();

  ipcMain.handle('app:info', () => ({ version: app.getVersion() }));

  ipcMain.handle('settings:get', () => config);
  ipcMain.handle('settings:set', (_e, patch) => {
    const autoChanged = 'autoUpdate' in patch && patch.autoUpdate !== config.autoUpdate;
    config = { ...config, ...patch };
    settings.save(config);
    if (autoChanged) scheduleAutoUpdates();
    return config;
  });

  ipcMain.handle('files:pick', async () => {
    const res = await dialog.showOpenDialog(win, {
      title: 'Choose Numbers files',
      properties: ['openFile', 'multiSelections'],
      filters: [{ name: 'Numbers spreadsheets', extensions: ['numbers'] }]
    });
    return res.canceled ? [] : res.filePaths;
  });

  ipcMain.handle('folder:pick', async () => {
    const res = await dialog.showOpenDialog(win, {
      title: 'Choose output folder',
      properties: ['openDirectory', 'createDirectory']
    });
    return res.canceled ? null : res.filePaths[0];
  });

  ipcMain.handle('files:preview', (_e, inputPath) => previewFile(inputPath));
  ipcMain.handle('files:convert', (_e, inputPath) => convertFile(inputPath, config.outputDir));

  ipcMain.handle('files:reveal', (_e, p) => shell.showItemInFolder(p));
  ipcMain.handle('files:open', (_e, p) => shell.openPath(p));

  ipcMain.handle('update:get', () => updater.getState());
  ipcMain.handle('update:check', () => updater.check());
  ipcMain.handle('update:install', () => updater.install());
}

app.on('window-all-closed', () => app.quit());
// Stop the helper before quitting so an update installer isn't blocked by its locked files.
app.on('before-quit', stopHelper);
