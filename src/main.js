const fs = require('fs');
const os = require('os');
const path = require('path');
const { app, BrowserWindow, ipcMain, dialog, shell, Menu, nativeTheme, clipboard, screen } = require('electron');
const { convertFile, previewFile, findFiles, targetPath, stopHelper, setLogger } = require('./convert');
const settings = require('./settings');
const updater = require('./updater');
const log = require('./log');

const AUTO_CHECK_INTERVAL = 6 * 60 * 60 * 1000; // 6 hours
const REPO = 'https://github.com/silent-diffusion/numbers-converter';
// The title bar's colours, matching the page's light and dark backgrounds (styles.css).
const CHROME = {
  light: { background: '#f3f3f3', symbols: '#1b1b1b' },
  dark: { background: '#202020', symbols: '#ffffff' }
};

let win = null;
let config = null;
let autoCheckTimer = null;
let pendingFiles = numbersArgs(process.argv);
const savedThisSession = new Set(); // Excel copies the app wrote, which it may write over when saving a file again

function numbersArgs(argv) {
  return argv.slice(1).filter((a) => a.toLowerCase().replace(/[\\/]+$/, '').endsWith('.numbers'));
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

function saveConfig(patch) {
  config = { ...config, ...patch };
  settings.save(config);
}

function scheduleAutoUpdates() {
  clearInterval(autoCheckTimer);
  autoCheckTimer = null;
  if (!config.autoUpdate) return;
  updater.check();
  autoCheckTimer = setInterval(updater.check, AUTO_CHECK_INTERVAL);
}

const chrome = () => CHROME[nativeTheme.shouldUseDarkColors ? 'dark' : 'light'];

// Where the window was when the app last closed, if that's still on a connected screen.
function windowBounds() {
  const saved = config.window || {};
  const bounds = { width: Math.max(760, saved.width || 1120), height: Math.max(520, saved.height || 740) };
  if (Number.isFinite(saved.x) && Number.isFinite(saved.y)) {
    const area = screen.getDisplayMatching({ x: saved.x, y: saved.y, ...bounds }).workArea;
    const visible = saved.x >= area.x - 40 && saved.y >= area.y - 40
      && saved.x + 200 <= area.x + area.width && saved.y + 100 <= area.y + area.height;
    if (visible) Object.assign(bounds, { x: saved.x, y: saved.y });
  }
  return bounds;
}

function windowsName() {
  if (process.platform !== 'win32') return `${os.type()} ${os.release()}`;
  const build = Number(os.release().split('.')[2] || 0);
  return `${build >= 22000 ? 'Windows 11' : 'Windows 10'} (build ${build})`;
}

// Only this project's own pages open in the browser (README sections, releases, new issues).
function openExternal(url) {
  if (typeof url === 'string' && (url === REPO || url.startsWith(`${REPO}/`) || url.startsWith(`${REPO}#`))) {
    shell.openExternal(url);
  }
}

function issueUrl(details) {
  const body = [
    '**What happened**',
    '',
    '<!-- What were you doing? Please don’t attach files with private data. -->',
    '',
    '',
    '**Details from the app**',
    '',
    '```',
    String(details || '').slice(0, 2000),
    '```'
  ].join('\n');
  return `${REPO}/issues/new?${new URLSearchParams({ title: 'Problem: ', body })}`;
}

function rememberFolder(file) {
  if (file) saveConfig({ lastFolder: path.dirname(file) });
}

// The numbers-converter command that comes with this copy: on the PATH when Scoop installed it, otherwise
// numbers-converter.cmd beside the app. None when the app runs from its source folder.
function commandLine() {
  if (!app.isPackaged) return null;
  if (updater.getState().managedBy === 'scoop') return { command: 'numbers-converter', onPath: true };
  const command = path.join(path.dirname(process.execPath), 'numbers-converter.cmd');
  return fs.existsSync(command) ? { command, onPath: false } : null;
}

function start() {
  config = settings.load();
  log.init(path.join(app.getPath('userData'), 'logs'));
  setLogger((text) => log.write('converter', text));
  process.on('uncaughtException', (err) => log.write('app', err?.stack || err));
  Menu.setApplicationMenu(null);

  win = new BrowserWindow({
    ...windowBounds(),
    minWidth: 760,
    minHeight: 520,
    show: false,
    backgroundColor: chrome().background,
    title: 'Numbers Converter',
    icon: path.join(__dirname, 'renderer', 'icon.png'),
    titleBarStyle: 'hidden',
    titleBarOverlay: { color: chrome().background, symbolColor: chrome().symbols, height: 48 },
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      sandbox: true
    }
  });
  if (config.window?.maximized) win.maximize();
  win.once('ready-to-show', () => win.show());
  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
  win.webContents.on('did-finish-load', () => {
    if (pendingFiles.length) win.webContents.send('files:opened', pendingFiles);
    pendingFiles = [];
  });

  nativeTheme.on('updated', () => {
    if (!win || win.isDestroyed()) return;
    win.setBackgroundColor(chrome().background);
    win.setTitleBarOverlay({ color: chrome().background, symbolColor: chrome().symbols });
  });

  // Remember the window's size and place (its normal size while maximized) for next time.
  let boundsTimer = null;
  const rememberBounds = (now) => {
    clearTimeout(boundsTimer);
    const save = () => {
      if (!win || win.isDestroyed() || win.isMinimized()) return;
      saveConfig({ window: { ...win.getNormalBounds(), maximized: win.isMaximized() } });
    };
    if (now === true) save();
    else boundsTimer = setTimeout(save, 600);
  };
  for (const event of ['resize', 'move', 'maximize', 'unmaximize']) win.on(event, rememberBounds);
  win.on('close', () => rememberBounds(true));

  // Links never open inside the app's window; this project's pages open in the browser.
  win.webContents.setWindowOpenHandler(({ url }) => {
    openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (event, url) => {
    if (!url.startsWith('file:')) {
      event.preventDefault();
      openExternal(url);
    }
  });

  // There's no menu bar, so the zoom shortcuts it would have provided are handled here.
  win.webContents.on('before-input-event', (event, input) => {
    if (input.type !== 'keyDown' || !(input.control || input.meta) || input.alt) return;
    const step = { '=': 0.5, '+': 0.5, '-': -0.5, '0': 0 }[input.key];
    if (step === undefined) return;
    const level = input.key === '0' ? 0 : Math.max(-3, Math.min(3, win.webContents.getZoomLevel() + step));
    win.webContents.setZoomLevel(level);
    event.preventDefault();
  });

  updater.init((state) => {
    if (state.status === 'error') log.write('update', state.message);
    win?.webContents.send('update:state', state);
  });
  scheduleAutoUpdates();

  ipcMain.handle('app:info', () => ({ version: require('../package.json').version, os: windowsName(), commandLine: commandLine() }));

  ipcMain.handle('settings:get', () => config);
  ipcMain.handle('settings:set', (_e, patch) => {
    const autoChanged = 'autoUpdate' in patch && patch.autoUpdate !== config.autoUpdate;
    saveConfig(patch);
    if (autoChanged) scheduleAutoUpdates();
    return config;
  });

  ipcMain.handle('files:pick', async () => {
    const res = await dialog.showOpenDialog(win, {
      title: 'Choose Numbers files',
      defaultPath: config.lastFolder || undefined,
      properties: ['openFile', 'multiSelections'],
      filters: [{ name: 'Numbers spreadsheets', extensions: ['numbers'] }]
    });
    if (res.canceled) return [];
    rememberFolder(res.filePaths[0]);
    return res.filePaths;
  });

  // Folders with Numbers files in them, or files a Mac saved as a "package": a folder whose name ends in .numbers.
  ipcMain.handle('files:pickFolders', async () => {
    const res = await dialog.showOpenDialog(win, {
      title: 'Choose a folder of Numbers files, or a Mac package folder',
      defaultPath: config.lastFolder || undefined,
      properties: ['openDirectory', 'multiSelections']
    });
    if (res.canceled) return [];
    rememberFolder(res.filePaths[0]);
    return res.filePaths;
  });
  ipcMain.handle('files:find', (_e, paths) => findFiles(paths));

  ipcMain.handle('folder:pick', async () => {
    const res = await dialog.showOpenDialog(win, {
      title: 'Choose where to save Excel files',
      defaultPath: config.outputDir || config.lastFolder || undefined,
      properties: ['openDirectory', 'createDirectory']
    });
    return res.canceled ? null : res.filePaths[0];
  });

  ipcMain.handle('files:preview', async (_e, inputPath, password, layout) => {
    const res = await previewFile(inputPath, password, layout);
    if (!res.ok && !res.needsPassword) log.write('open', [res.error, res.details].filter(Boolean).join(' — '));
    return res;
  });
  ipcMain.handle('files:convert', async (_e, inputPath, password, layout, options = {}) => {
    const replace = options.replace && savedThisSession.has(options.replace) ? options.replace : null;
    const res = await convertFile(inputPath, config.outputDir, password, layout, { contents: config.contentsSheet, replace });
    if (res.ok) savedThisSession.add(res.outputPath);
    else log.write('save', [res.error, res.details].filter(Boolean).join(' — '));
    return res;
  });
  ipcMain.handle('files:target', (_e, inputPath) => targetPath(inputPath, config.outputDir));

  ipcMain.handle('files:reveal', (_e, p) => shell.showItemInFolder(p));
  ipcMain.handle('files:open', (_e, p) => shell.openPath(p));

  ipcMain.handle('link:open', (_e, url) => openExternal(url));
  ipcMain.handle('log:open', () => {
    fs.mkdirSync(log.folder(), { recursive: true });
    return shell.openPath(log.folder());
  });
  ipcMain.handle('report:open', (_e, details) => shell.openExternal(issueUrl(details)));
  ipcMain.handle('clipboard:write', (_e, text) => clipboard.writeText(String(text)));

  ipcMain.handle('update:get', () => updater.getState());
  ipcMain.handle('update:check', () => updater.check());
  ipcMain.handle('update:install', () => updater.install());
}

app.on('window-all-closed', () => app.quit());
// Stop the helper before quitting so an update installer isn't blocked by its locked files.
app.on('before-quit', stopHelper);
