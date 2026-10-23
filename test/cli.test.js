// Runs the numbers-converter command as people do, from a terminal: its help and version, a folder
// converted into another one with a JSON summary, a dry run, a locked file, and its exit codes.
//
//   node test/cli.test.js                  the command in py/, run by the project's .venv
//   node test/cli.test.js --built          numbers-converter.cmd in dist/win-unpacked, as the app ships it
//   node test/cli.test.js --app <folder>   the same in another unpacked build
const assert = require('assert');
const { spawnSync } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const FIXTURES = path.join(ROOT, 'test', 'fixtures');
const VERSION = require('../package.json').version;
const cli = process.argv.slice(2);
const BUILT = cli.includes('--built') || cli.includes('--app');
const APP = cli.includes('--app') ? path.resolve(cli[cli.indexOf('--app') + 1] || '') : path.join(ROOT, 'dist', 'win-unpacked');
const COMMAND = path.join(APP, 'numbers-converter.cmd');

// A scratch folder named as real ones are, with spaces, an ampersand and an accent.
const SCRATCH = fs.realpathSync.native(fs.mkdtempSync(path.join(os.tmpdir(), 'nc-cli-')));
const WORK = path.join(SCRATCH, 'Café & Co');
const REPORTS = path.join(WORK, 'Reports');

// Runs the command in WORK: { code, out, err }. The built one goes through cmd.exe, as a .cmd does from
// PowerShell, so its launcher's quoting and exit code are tested too. Nothing can be typed in reply.
function run(...args) {
  const options = { cwd: WORK, encoding: 'utf8', timeout: 180000, stdio: ['ignore', 'pipe', 'pipe'] };
  const res = BUILT
    ? spawnSync('cmd.exe', ['/d', '/s', '/c', `"${[COMMAND, ...args].map((a) => `"${a}"`).join(' ')}"`],
      { ...options, windowsVerbatimArguments: true })
    : spawnSync(path.join(ROOT, '.venv', 'Scripts', 'python.exe'), [path.join(ROOT, 'py', 'numbers_tool.py'), ...args], options);
  if (res.error) throw res.error;
  const text = (s) => s.replace(/\r\n/g, '\n');
  return { code: res.status, out: text(res.stdout), err: text(res.stderr) };
}

function assertXlsx(file) {
  assert.ok(fs.existsSync(file), `${file} wasn’t saved`);
  assert.strictEqual(fs.readFileSync(file).subarray(0, 2).toString(), 'PK', `${file} isn’t an Excel file`);
}

try {
  if (BUILT) assert.ok(fs.existsSync(COMMAND), `${COMMAND} doesn’t exist. Run \`npm run dist\` first.`);
  fs.mkdirSync(path.join(REPORTS, 'Q1'), { recursive: true });
  fs.copyFileSync(path.join(FIXTURES, 'issue-9.numbers'), path.join(REPORTS, 'issue-9.numbers'));
  fs.copyFileSync(path.join(FIXTURES, 'encrypted.numbers'), path.join(REPORTS, 'Q1', 'encrypted.numbers'));
  fs.copyFileSync(path.join(FIXTURES, 'test-styles.numbers'), path.join(REPORTS, 'Q1', 'Styles.numbers'));
  fs.writeFileSync(path.join(REPORTS, 'notes.txt'), 'not a spreadsheet');

  // Its help, and the version of the app it came with.
  const help = run('--help');
  assert.strictEqual(help.code, 0, help.err);
  assert.match(help.out, /^usage: numbers-converter /);
  assert.match(help.out, /exit codes:/);
  assert.strictEqual(run('--version').out, `numbers-converter ${VERSION}\n`);

  // A folder and its subfolders, into another folder, with a JSON summary. The locked file needs a
  // password, so the exit code is 1.
  const excel = path.join(WORK, 'Excel');
  const first = run('Reports', '--out', 'Excel', '--json');
  assert.strictEqual(first.code, 1, first.out + first.err);
  const summary = JSON.parse(first.out);
  assert.deepStrictEqual([summary.version, summary.saved, summary.locked, summary.failed], [VERSION, 2, 1, 0]);
  assert.deepStrictEqual(summary.files.map((f) => [f.status, f.output || f.input]), [
    ['saved', path.join(excel, 'issue-9.xlsx')],
    ['locked', path.join(REPORTS, 'Q1', 'encrypted.numbers')],
    ['saved', path.join(excel, 'Q1', 'Styles.xlsx')]
  ]);
  assert.strictEqual(summary.files[1].hint, 's3cr3t');
  for (const f of summary.files.filter((x) => x.output)) assertXlsx(f.output);

  // A dry run says what it would save where, and what changes in Excel, and saves nothing.
  const dry = run(path.join('Reports', 'issue-9.numbers'), '--dry-run');
  assert.strictEqual(dry.code, 0, dry.out + dry.err);
  assert.strictEqual(dry.out, `Would save  ${path.join('Reports', 'issue-9.xlsx')} — 1 change to check: 1 chart can’t be converted\n`);
  assert.ok(!fs.existsSync(path.join(REPORTS, 'issue-9.xlsx')), 'the dry run saved a file');

  // With the password from a file, the locked file is saved too, and nothing is written over.
  fs.writeFileSync(path.join(WORK, 'password.txt'), 's3cr3t\n');
  const second = run('Reports', '--out', 'Excel', '--password-file', 'password.txt');
  assert.strictEqual(second.code, 0, second.out + second.err);
  assert.strictEqual(second.out, [
    'Converting 3 Numbers files…',
    `Saved   ${path.join('Excel', 'issue-9 (1).xlsx')} — 1 change to check: 1 chart can’t be converted`,
    `Saved   ${path.join('Excel', 'Q1', 'encrypted.xlsx')}`,
    `Saved   ${path.join('Excel', 'Q1', 'Styles (1).xlsx')}`,
    '3 saved.',
    ''
  ].join('\n'));
  assertXlsx(path.join(excel, 'Q1', 'encrypted.xlsx'));

  // Exit code 1 for a file that isn't there, 2 when there's nothing to convert or the command has a mistake.
  const gone = run('gone.numbers');
  assert.deepStrictEqual([gone.code, gone.out], [1, 'Failed  gone.numbers — This file doesn’t exist.\n']);
  fs.mkdirSync(path.join(WORK, 'Empty'));
  const empty = run('Empty');
  assert.deepStrictEqual([empty.code, empty.err], [2, 'No Numbers files in Empty.\n']);
  assert.strictEqual(run('Reports', '--layout', 'sideways').code, 2);
  assert.strictEqual(run().code, 2);

  fs.rmSync(SCRATCH, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
  console.log(`All command-line tests passed (${BUILT ? path.relative(ROOT, COMMAND) || COMMAND : 'py/numbers_tool.py'}).`);
} catch (err) {
  console.error(err);
  console.error(`\nScratch folder: ${SCRATCH}`);
  process.exit(1);
}
