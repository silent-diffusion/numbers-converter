const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

XLSX.set_fs(fs);

const PREVIEW_MAX_ROWS = 1000;
const PREVIEW_MAX_COLS = 50;

// Returns a path in `dir` named after `base` that doesn't exist yet: "Budget.xlsx", "Budget (1).xlsx", ...
function uniquePath(dir, base, ext) {
  let candidate = path.join(dir, base + ext);
  for (let i = 1; fs.existsSync(candidate); i++) {
    candidate = path.join(dir, `${base} (${i})${ext}`);
  }
  return candidate;
}

function readNumbers(inputPath) {
  const stat = fs.statSync(inputPath);
  if (stat.isDirectory()) {
    throw new Error('This file was saved as a Numbers package. Re-save it in Numbers as a single file, or zip it first.');
  }
  if (path.extname(inputPath).toLowerCase() !== '.numbers') {
    throw new Error('Not a .numbers file.');
  }

  const data = fs.readFileSync(inputPath);
  // Numbers files are zip archives; without this check SheetJS would happily parse any text as CSV.
  if (data.length < 4 || data.readUInt32LE(0) !== 0x04034b50) {
    throw new Error('Could not read this Numbers file (it may be damaged or not a real Numbers file).');
  }

  let workbook;
  try {
    workbook = XLSX.read(data, { type: 'buffer', cellDates: true });
  } catch (err) {
    throw new Error(`Could not read this Numbers file (${err.message}).`);
  }
  if (!workbook.SheetNames.length) {
    throw new Error('No tables found in this file.');
  }
  return workbook;
}

// Formatted cell text for each sheet, capped so huge tables stay responsive.
function previewFile(inputPath) {
  const workbook = readNumbers(inputPath);
  const sheets = workbook.SheetNames.map((name) => {
    const ws = workbook.Sheets[name];
    if (!ws['!ref']) return { name, rows: [], totalRows: 0, totalCols: 0 };
    const range = XLSX.utils.decode_range(ws['!ref']);
    const totalRows = range.e.r - range.s.r + 1;
    const totalCols = range.e.c - range.s.c + 1;
    const shown = {
      s: range.s,
      e: {
        r: Math.min(range.e.r, range.s.r + PREVIEW_MAX_ROWS - 1),
        c: Math.min(range.e.c, range.s.c + PREVIEW_MAX_COLS - 1)
      }
    };
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '', blankrows: true, range: shown, dateNF: 'm/d/yyyy' });
    const numeric = rows.map((row, r) => row.map((_, c) => {
      const cell = ws[XLSX.utils.encode_cell({ r: shown.s.r + r, c: shown.s.c + c })];
      return cell?.t === 'n' || cell?.t === 'd';
    }));
    return { name, rows, numeric, totalRows, totalCols, startRow: range.s.r, startCol: range.s.c };
  });
  return { sheets };
}

function convertFile(inputPath, outputDir) {
  const workbook = readNumbers(inputPath);
  const dir = outputDir || path.dirname(inputPath);
  const base = path.basename(inputPath, path.extname(inputPath));
  const outputPath = uniquePath(dir, base, '.xlsx');
  XLSX.writeFile(workbook, outputPath, { bookType: 'xlsx', compression: true });
  return { outputPath, sheets: workbook.SheetNames.length };
}

module.exports = { convertFile, previewFile };
