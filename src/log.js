const fs = require('fs');
const path = require('path');

// A small error log kept on this PC (in the app's data folder) to help with problem reports.
// It's never sent anywhere; Settings → Help → Open log folder shows it.
const MAX_BYTES = 1024 * 1024;
let dir = null;

function init(folder) {
  dir = folder;
}

function write(kind, message) {
  if (!dir) return;
  try {
    fs.mkdirSync(dir, { recursive: true });
    const file = path.join(dir, 'numbers-converter.log');
    if (fs.existsSync(file) && fs.statSync(file).size > MAX_BYTES) {
      fs.renameSync(file, path.join(dir, 'numbers-converter.old.log'));
    }
    const text = String(message).trimEnd().replace(/\r?\n/g, '\n    ');
    fs.appendFileSync(file, `${new Date().toISOString()} ${kind}: ${text}\n`);
  } catch {
    // logging must never get in the way of converting
  }
}

module.exports = { init, write, folder: () => dir };
