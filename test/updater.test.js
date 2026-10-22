// Checks that only copies set up by the installer update themselves: a Scoop copy or a portable
// (unzipped) copy says how it's updated instead, and never checks GitHub.
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const Module = require('module');

let checks = 0;
const fakes = {
  electron: { app: { isPackaged: true, getName: () => 'Numbers Converter' } },
  'electron-updater': { autoUpdater: { on() {}, checkForUpdates: async () => { checks++; } } }
};
const load = Module._load;
Module._load = (request, ...rest) => (request in fakes ? fakes[request] : load(request, ...rest));

function stateFor(appDir) {
  fs.mkdirSync(appDir, { recursive: true });
  const exe = path.join(appDir, 'Numbers Converter.exe');
  Object.defineProperty(process, 'execPath', { value: exe, configurable: true, writable: true });
  delete require.cache[require.resolve('../src/updater')];
  const updater = require('../src/updater');
  updater.init(() => {});
  updater.check();
  return updater.getState();
}

const root = fs.mkdtempSync(path.join(os.tmpdir(), 'nc-updater-'));
const realExecPath = process.execPath;
try {
  const installed = path.join(root, 'Programs', 'numbers-converter');
  fs.mkdirSync(installed, { recursive: true });
  fs.writeFileSync(path.join(installed, 'Uninstall Numbers Converter.exe'), '');
  assert.strictEqual(stateFor(installed).status, 'idle');
  assert.strictEqual(checks, 1, 'an installed copy checks for updates');

  const scoop = stateFor(path.join(root, 'scoop', 'apps', 'numbers-converter', '1.9.0'));
  assert.strictEqual(scoop.managedBy, 'scoop');
  assert.match(scoop.message, /scoop update numbers-converter/);

  const portable = stateFor(path.join(root, 'Downloads', 'NumbersConverter-1.9.0-win-x64'));
  assert.strictEqual(portable.managedBy, 'portable');
  assert.strictEqual(checks, 1, 'Scoop and portable copies never check GitHub');
} finally {
  Object.defineProperty(process, 'execPath', { value: realExecPath, configurable: true, writable: true });
  Module._load = load;
  fs.rmSync(root, { recursive: true, force: true });
}
console.log('All updater tests passed.');
