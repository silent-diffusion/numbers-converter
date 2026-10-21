const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

XLSX.set_fs(fs);

// Returns a path in `dir` named after `base` that doesn't exist yet: "Budget.xlsx", "Budget (1).xlsx", ...
function uniquePath(dir, base, ext) {
  let candidate = path.join(dir, base + ext);
  for (let i = 1; fs.existsSync(candidate); i++) {
    candidate = path.join(dir, `${base} (${i})${ext}`);
  }
  return candidate;
}

function convertFile(inputPath, outputDir) {
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

  const dir = outputDir || path.dirname(inputPath);
  const base = path.basename(inputPath, path.extname(inputPath));
  const outputPath = uniquePath(dir, base, '.xlsx');
  XLSX.writeFile(workbook, outputPath, { bookType: 'xlsx', compression: true });
  return { outputPath, sheets: workbook.SheetNames.length };
}

module.exports = { convertFile };
