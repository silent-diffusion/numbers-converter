const fs = require('fs');
const path = require('path');
const { app } = require('electron');
const { autoUpdater } = require('electron-updater');

// Update state pushed to the renderer. `status` is one of:
// idle | unsupported | checking | available | downloading | downloaded | up-to-date | error
// (`managedBy` says why it's unsupported when something else updates this copy: scoop | portable).
let state = { status: 'idle' };
let notify = () => {};

// What the app says when another way of updating looks after this copy.
const ELSEWHERE = {
  scoop: 'Scoop keeps this copy up to date. To update, run: scoop update numbers-converter',
  portable: 'This copy runs without being installed, so it doesn’t update itself. Download the new version from GitHub, or use the installer.'
};

function set(next) {
  state = { ...next, checkedAt: next.checkedAt ?? state.checkedAt };
  notify(state);
}

// Copies the installer didn't set up are updated some other way: Scoop updates its own copy, and a
// copy unzipped by hand is replaced by hand. Installing an update over either would leave a second
// copy, so they don't update themselves. The installer puts its uninstaller beside the app.
function updatedElsewhere() {
  const dir = path.dirname(process.execPath);
  if (/[\\/]scoop[\\/]apps[\\/]/i.test(dir)) return 'scoop';
  if (!fs.existsSync(path.join(dir, `Uninstall ${app.getName()}.exe`))) return 'portable';
  return null;
}

function init(onChange) {
  notify = onChange;
  if (!app.isPackaged) {
    set({ status: 'unsupported', message: 'Updates are only available in the installed app.' });
    return;
  }
  const managedBy = updatedElsewhere();
  if (managedBy) {
    set({ status: 'unsupported', managedBy, message: ELSEWHERE[managedBy] });
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
  if (!app.isPackaged || ['unsupported', 'checking', 'downloading'].includes(state.status)) return;
  autoUpdater.checkForUpdates().catch(() => {}); // errors arrive via the 'error' event
}

function install() {
  if (state.status === 'downloaded') autoUpdater.quitAndInstall(true, true);
}

module.exports = { init, check, install, getState: () => state };
