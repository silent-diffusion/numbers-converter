const { app } = require('electron');
const { autoUpdater } = require('electron-updater');

// Update state pushed to the renderer. `status` is one of:
// idle | unsupported | checking | available | downloading | downloaded | up-to-date | error
let state = { status: 'idle' };
let notify = () => {};

function set(next) {
  state = { ...next, checkedAt: next.checkedAt ?? state.checkedAt };
  notify(state);
}

function init(onChange) {
  notify = onChange;
  if (!app.isPackaged) {
    set({ status: 'unsupported', message: 'Updates are only available in the installed app.' });
    return;
  }

  // Updates are always downloaded once found: automatically in the background,
  // or right after the user clicks "Check for updates".
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;

  autoUpdater.on('checking-for-update', () => set({ status: 'checking' }));
  autoUpdater.on('update-available', (info) => set({ status: 'downloading', version: info.version, percent: 0 }));
  autoUpdater.on('download-progress', (p) => set({ ...state, status: 'downloading', percent: Math.round(p.percent) }));
  autoUpdater.on('update-not-available', () => set({ status: 'up-to-date', checkedAt: Date.now() }));
  autoUpdater.on('update-downloaded', (info) => set({ status: 'downloaded', version: info.version, checkedAt: Date.now() }));
  autoUpdater.on('error', (err) => set({ status: 'error', message: friendlyError(err), checkedAt: Date.now() }));
}

function friendlyError(err) {
  const msg = String(err?.message || err);
  if (/ENOTFOUND|ENETUNREACH|ECONNREFUSED|ETIMEDOUT|net::ERR_/i.test(msg)) {
    return "Couldn't reach the update server. Check your internet connection.";
  }
  return 'Update check failed. Please try again later.';
}

function check() {
  if (!app.isPackaged || state.status === 'checking' || state.status === 'downloading') return;
  autoUpdater.checkForUpdates().catch(() => {}); // errors arrive via the 'error' event
}

function install() {
  if (state.status === 'downloaded') autoUpdater.quitAndInstall(true, true);
}

module.exports = { init, check, install, getState: () => state };
