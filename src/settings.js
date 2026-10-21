const fs = require('fs');
const path = require('path');
const { app } = require('electron');

const DEFAULTS = {
  autoUpdate: true,
  autoConvert: false, // save an .xlsx as soon as a file is opened, instead of waiting for "Convert & save"
  outputDir: null, // null = save next to the original file
  sheetLayout: 'tables', // Numbers sheets with several tables: 'tables' = a worksheet per table, 'sheet' = one, as in Numbers
  contentsSheet: false, // start each Excel file with a sheet linking to every table
  window: null, // { x, y, width, height, maximized } when the app last closed
  lastFolder: null, // where the file dialogs open
  fileLayouts: {} // { [file path]: { [Numbers sheet]: 'tables' | 'sheet' } } chosen in the preview, most recent last
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
