const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('api', {
  info: () => ipcRenderer.invoke('app:info'),

  getSettings: () => ipcRenderer.invoke('settings:get'),
  setSettings: (patch) => ipcRenderer.invoke('settings:set', patch),

  pathForFile: (file) => webUtils.getPathForFile(file),
  pickFiles: () => ipcRenderer.invoke('files:pick'),
  pickFolders: () => ipcRenderer.invoke('files:pickFolders'),
  find: (paths) => ipcRenderer.invoke('files:find', paths),
  pickFolder: () => ipcRenderer.invoke('folder:pick'),
  preview: (p, password, layout) => ipcRenderer.invoke('files:preview', p, password, layout),
  convert: (p, password, layout, options) => ipcRenderer.invoke('files:convert', p, password, layout, options),
  target: (p) => ipcRenderer.invoke('files:target', p),
  reveal: (p) => ipcRenderer.invoke('files:reveal', p),
  open: (p) => ipcRenderer.invoke('files:open', p),
  onFilesOpened: (cb) => ipcRenderer.on('files:opened', (_e, files) => cb(files)),

  openLink: (url) => ipcRenderer.invoke('link:open', url),
  openLogFolder: () => ipcRenderer.invoke('log:open'),
  reportProblem: (details) => ipcRenderer.invoke('report:open', details),
  copy: (text) => ipcRenderer.invoke('clipboard:write', text),

  getUpdateState: () => ipcRenderer.invoke('update:get'),
  checkForUpdates: () => ipcRenderer.invoke('update:check'),
  installUpdate: () => ipcRenderer.invoke('update:install'),
  onUpdateState: (cb) => ipcRenderer.on('update:state', (_e, s) => cb(s))
});
