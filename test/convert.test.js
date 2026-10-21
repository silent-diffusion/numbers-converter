// Round-trip test: build a .numbers file with SheetJS, convert it, read the .xlsx back.
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const XLSX = require('xlsx');
const { convertFile, previewFile } = require('../src/convert');

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'numbers-converter-'));
const input = path.join(dir, 'Budget.numbers');

const rows = [
  ['Item', 'Amount', 'Paid'],
  ['Rent', 1200, true],
  ['Groceries', 342.5, false],
  ['Internet', 60, true]
];
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Expenses');
// SheetJS needs its Numbers template payload to write .numbers files.
XLSX.writeFile(wb, input, { bookType: 'numbers', numbers: require('xlsx/dist/xlsx.zahl') });

const preview = previewFile(input);
assert.strictEqual(preview.sheets.length, 1);
assert.strictEqual(preview.sheets[0].totalRows, 4);
assert.deepStrictEqual(preview.sheets[0].rows[0], ['Item', 'Amount', 'Paid']);
assert.strictEqual(preview.sheets[0].rows[2][1], '342.5');
assert.deepStrictEqual(preview.sheets[0].numeric[1], [false, true, false]);

const first = convertFile(input);
assert.strictEqual(first.outputPath, path.join(dir, 'Budget.xlsx'));
const out = XLSX.readFile(first.outputPath);
const back = XLSX.utils.sheet_to_json(out.Sheets[out.SheetNames[0]], { header: 1 });
assert.deepStrictEqual(back, rows);

const second = convertFile(input);
assert.strictEqual(second.outputPath, path.join(dir, 'Budget (1).xlsx'), 'must not overwrite existing output');

assert.throws(() => convertFile(path.join(dir, 'Budget.xlsx')), /Not a \.numbers file/);
const junk = path.join(dir, 'Broken.numbers');
fs.writeFileSync(junk, 'not a spreadsheet');
assert.throws(() => convertFile(junk), /Could not read/);

fs.rmSync(dir, { recursive: true, force: true });
console.log('All conversion tests passed.');
