// Exercises the Node <-> Python helper bridge with real Numbers files.
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { convertFile, previewFile, stopHelper } = require('../src/convert');

const fixture = (name) => path.join(__dirname, 'fixtures', `${name}.numbers`);

(async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'numbers-converter-'));
  const input = path.join(dir, 'Styles & Formats.numbers');
  fs.copyFileSync(fixture('test-styles'), input);

  const preview = await previewFile(input);
  assert.ok(preview.ok, preview.error);
  assert.deepStrictEqual(preview.sheets.map((s) => s.name), ['Styles', 'Headers', 'Borders', 'Large Borders']);
  const [text, styleIdx] = preview.sheets[0].rows[0][0];
  assert.match(text, /^name="Title"/);
  assert.strictEqual(preview.sheets[0].styles[styleIdx].b, 1);

  const first = await convertFile(input);
  assert.ok(first.ok, first.error);
  assert.strictEqual(first.outputPath, path.join(dir, 'Styles & Formats.xlsx'));
  assert.ok(fs.statSync(first.outputPath).size > 0);

  const second = await convertFile(input);
  assert.strictEqual(second.outputPath, path.join(dir, 'Styles & Formats (1).xlsx'), 'must not overwrite existing output');

  const outDir = fs.mkdtempSync(path.join(os.tmpdir(), 'numbers-out-'));
  const elsewhere = await convertFile(input, outDir);
  assert.strictEqual(path.dirname(elsewhere.outputPath), outDir);

  // Charts become Excel charts and show in the preview; Excel has no bubble charts, so that one is reported.
  const charts = await previewFile(fixture('issue-9'));
  assert.strictEqual(charts.unsupported.charts, 1);
  assert.strictEqual(charts.sheets.reduce((n, s) => n + s.charts.length, 0), 10);
  const withCharts = await convertFile(fixture('issue-9'), outDir);
  assert.strictEqual(withCharts.charts, 10);

  const locked = await previewFile(fixture('encrypted'));
  assert.ok(!locked.ok && locked.needsPassword && /password/.test(locked.error));
  const wrong = await previewFile(fixture('encrypted'), 'nope');
  assert.ok(!wrong.ok && wrong.needsPassword && wrong.hint === 's3cr3t');
  const unlocked = await previewFile(fixture('encrypted'), 's3cr3t');
  assert.ok(unlocked.ok, unlocked.error);
  const lockedOut = await convertFile(fixture('encrypted'), outDir, 's3cr3t');
  assert.ok(lockedOut.ok && fs.existsSync(lockedOut.outputPath));

  const pkg = await convertFile(fixture('folder-package'), outDir);
  assert.ok(pkg.ok, pkg.error);
  assert.strictEqual(path.basename(pkg.outputPath), 'folder-package.xlsx');

  const junk = path.join(dir, 'Broken.numbers');
  fs.writeFileSync(junk, 'not a spreadsheet');
  const broken = await previewFile(junk);
  assert.ok(!broken.ok && /Could not read/.test(broken.error));

  const notNumbers = await previewFile(first.outputPath);
  assert.ok(!notNumbers.ok && /Not a \.numbers file/.test(notNumbers.error));

  stopHelper();
  fs.rmSync(dir, { recursive: true, force: true });
  fs.rmSync(outDir, { recursive: true, force: true });
  console.log('All bridge tests passed.');
})().catch((err) => {
  stopHelper();
  console.error(err);
  process.exit(1);
});
