const fs = require('fs');
const path = require('path');
const { app } = require('electron');

const DEFAULTS = {
  autoUpdate: true,
  outputDir: null // null = save next to the original file
};

const file = () => path.join(app.getPath('userData'), 'settings.json');

function load() {
  try {
    return { ...DEFAULTS, ...JSON.parse(fs.readFileSync(file(), 'utf8')) };
  } catch {
    return { ...DEFAULTS };
  }
}

function save(settings) {
  fs.mkdirSync(path.dirname(file()), { recursive: true });
  fs.writeFileSync(file(), JSON.stringify(settings, null, 2));
}

module.exports = { load, save };
