// Smoke test of the app's window. It starts the packaged app (dist/win-unpacked, built by `npm run dist`)
// and uses it as someone would, through Playwright: files opened the way "Open with" opens them and from
// the window, folders dropped or chosen, their previews, "What changes in Excel", the layout switch, a
// locked file, a damaged one, saving and Settings. Any uncaught error or console error fails it. Every
// .numbers file in test/fixtures is opened, previewed sheet by sheet and saved; expectations.js lists
// what particular files must show.
//
//   node test/ui/smoke.test.js              the app in dist/win-unpacked
//   node test/ui/smoke.test.js --source     the app in src/, with the Python helper in .venv (no build)
//   node test/ui/smoke.test.js --app <exe>  another unpacked build
//
// Each run has a scratch profile (Electron's --user-data-dir, which the single-instance lock goes by too),
// so it never reads or changes the real settings, nor meets or hands files to a copy of the app that's
// open. File dialogs and `shell` are stubbed in the main process, so nothing opens in Excel, Explorer or
// the browser, and only copies that don't update themselves are run, so nothing goes online. When a step
// fails, test-results/ui gets a screenshot, the app's log and what the app printed.
const assert = require('assert');
const { execFileSync, spawn } = require('child_process');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { _electron: electron } = require('playwright-core');
const EXPECTED = require('./expectations');

const ROOT = path.join(__dirname, '..', '..');
const FIXTURES = path.join(ROOT, 'test', 'fixtures');
const RESULTS = path.join(ROOT, 'test-results', 'ui');
const VERSION = require('../../package.json').version;
const STEP_TIMEOUT = 5 * 60 * 1000;

const cli = process.argv.slice(2);
const SOURCE = cli.includes('--source');
const EXE = cli.includes('--app') ? path.resolve(cli[cli.indexOf('--app') + 1] || '')
  : path.join(ROOT, 'dist', 'win-unpacked', 'Numbers Converter.exe');

// Everything a run writes goes in one scratch folder: the app's profile, and copies of the fixtures in a
// folder named as real ones are, with spaces, an ampersand and accents. Excel copies are saved beside them.
const SCRATCH = fs.realpathSync.native(fs.mkdtempSync(path.join(os.tmpdir(), 'nc-ui-')));
const PROFILE = path.join(SCRATCH, 'profile');
const WORK = path.join(SCRATCH, 'Café & Co – Q3');
const DAMAGED = path.join(WORK, 'Damaged.numbers');
// Folders of Numbers files to open: one with a subfolder and things to leave out, and one with none at all.
const ARCHIVE = path.join(SCRATCH, 'Archive 2024');
const ARCHIVE_FILES = [path.join(ARCHIVE, 'Budget.numbers'), path.join(ARCHIVE, 'Q1', 'Sales.numbers')];
const MORE = path.join(SCRATCH, 'More');
const NOTHING = path.join(SCRATCH, 'Nothing here');
const LOG = path.join(PROFILE, 'logs', 'numbers-converter.log');

// Every .numbers file (or Mac package folder) in test/fixtures and the folders under it, by its path there.
function findFixtures(dir = FIXTURES, prefix = '') {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const key = prefix + entry.name;
    if (/\.numbers$/i.test(entry.name)) return [{ key, path: path.join(WORK, ...key.split('/')) }];
    return entry.isDirectory() ? findFixtures(path.join(dir, entry.name), `${key}/`) : [];
  });
}
const FILES = findFixtures();
const fixture = (key) => FILES.find((f) => f.key === key) || assert.fail(`test/fixtures/${key} is missing.`);

let electronApp = null; // the running app (Playwright's ElectronApplication)
let win = null; // its window (a Playwright Page)
let mainPid = null;
let appClosed = false;
const problems = []; // uncaught errors, console errors, network requests: each fails the step it shows up in
const transcript = []; // everything the app printed, saved when a step fails
const checked = new Set(); // fixtures whose preview has been checked sheet by sheet
let logLinesSeen = 0;

/* ---------- Setting up, starting and stopping the app ---------- */

// Copies the installer set up update themselves (so they'd check GitHub) and may be the user's own:
// the test only runs builds.
function checkBuild() {
  for (const key of Object.keys(EXPECTED)) fixture(key);
  if (SOURCE) return;
  assert.ok(fs.existsSync(EXE), `${EXE} doesn’t exist. Run \`npm run dist\` first, or test src/ with --source.`);
  const programs = path.join(process.env.LOCALAPPDATA || 'C:\\nowhere', 'Programs').toLowerCase() + path.sep;
  const installed = fs.existsSync(path.join(path.dirname(EXE), 'Uninstall Numbers Converter.exe'))
    || EXE.toLowerCase().startsWith(programs);
  assert.ok(!installed, `${EXE} is an installed copy. The test only runs builds, such as dist/win-unpacked.`);
}

function setUp() {
  for (const f of FILES) fs.cpSync(path.join(FIXTURES, ...f.key.split('/')), f.path, { recursive: true });
  // A Mac package whose tables can't be read, though the picture Numbers saves beside them can.
  fs.cpSync(path.join(FIXTURES, 'folder-package.numbers'), DAMAGED, { recursive: true });
  fs.writeFileSync(path.join(DAMAGED, 'Index.zip'), 'not a zip archive');
  fs.writeFileSync(path.join(WORK, 'notes.txt'), 'not a spreadsheet');
  fs.mkdirSync(path.join(ARCHIVE, 'Q1'), { recursive: true });
  fs.copyFileSync(path.join(FIXTURES, 'test-styles.numbers'), ARCHIVE_FILES[0]);
  fs.copyFileSync(path.join(FIXTURES, 'images.numbers'), ARCHIVE_FILES[1]);
  fs.writeFileSync(path.join(ARCHIVE, '._Budget.numbers'), 'what a Mac leaves on a drive');
  fs.writeFileSync(path.join(ARCHIVE, 'notes.txt'), 'not a spreadsheet');
  fs.mkdirSync(MORE);
  fs.copyFileSync(path.join(FIXTURES, 'issue-9.numbers'), path.join(MORE, 'Charts.numbers'));
  fs.mkdirSync(NOTHING);
  // Automatic update checks are off too, in case a build ever looks installed.
  fs.mkdirSync(PROFILE);
  fs.writeFileSync(path.join(PROFILE, 'settings.json'), JSON.stringify({ autoUpdate: false }));
}

function appEnv() {
  const env = { ...process.env };
  delete env.ELECTRON_RUN_AS_NODE; // would start the app as plain Node
  return env;
}

async function launch(files) {
  const profile = `--user-data-dir=${PROFILE}`;
  electronApp = await electron.launch({
    ...(SOURCE ? { args: [ROOT, profile, ...files] } : { executablePath: EXE, args: [profile, ...files] }),
    cwd: os.tmpdir(), // not the scratch folder, which a process still exiting would keep from being deleted
    env: appEnv(),
    timeout: 90000
  });
  electronApp.on('close', () => { appClosed = true; });
  const started = await electronApp.evaluate(({ app }) => ({ userData: app.getPath('userData'), pid: process.pid }));
  mainPid = started.pid;
  if (path.resolve(started.userData).toLowerCase() !== PROFILE.toLowerCase()) {
    kill();
    throw new Error(`The app isn’t using the scratch profile (it uses ${started.userData}), so it was stopped.`);
  }

  // File dialogs answer with the files the test queues (pick()), and `shell` and the clipboard only record
  // what they're given: nothing opens in Excel, Explorer or the browser, and the clipboard is left alone.
  await electronApp.evaluate(({ dialog, shell, clipboard }) => {
    const calls = { picks: [], dialogs: [], opened: [], revealed: [], links: [], copied: [], rejections: [] };
    globalThis.uiTest = calls;
    const stub = (fn) => Object.assign(fn, { uiTestStub: true });
    dialog.showOpenDialog = stub(async (_window, options) => {
      calls.dialogs.push({ title: options.title, defaultPath: options.defaultPath });
      const filePaths = calls.picks.shift() || [];
      return { canceled: !filePaths.length, filePaths };
    });
    shell.openPath = stub(async (p) => { calls.opened.push(p); return ''; });
    shell.showItemInFolder = stub((p) => { calls.revealed.push(p); });
    shell.openExternal = stub(async (url) => { calls.links.push(url); });
    clipboard.writeText = stub((text) => { calls.copied.push(text); });
    process.on('unhandledRejection', (reason) => calls.rejections.push(String(reason?.stack || reason)));
  });

  win = await electronApp.firstWindow();
  win.setDefaultTimeout(30000);
  const seen = new Set();
  const once = (key) => !seen.has(key) && seen.add(key);
  const onConsole = (msg) => {
    const { url, lineNumber, columnNumber } = msg.location();
    if (!once(`${msg.type()} ${msg.text()} ${url}:${lineNumber}:${columnNumber}`)) return;
    transcript.push(`[window ${msg.type()}] ${msg.text()}`);
    if (msg.type() === 'error') problems.push(`Console error in the window: ${msg.text()}`);
  };
  const onError = (err) => {
    if (once(err.stack || err.message)) problems.push(`Uncaught error in the window: ${err.stack || err.message}`);
  };
  win.on('console', onConsole);
  win.on('pageerror', onError);
  // The packaged app starts before Playwright attaches: what the window printed until then comes from here.
  (await win.consoleMessages({ filter: 'all' })).forEach(onConsole);
  (await win.pageErrors({ filter: 'all' })).forEach(onError);
  win.on('crash', () => problems.push('The window crashed.'));
  win.on('request', (req) => {
    if (!/^(file|data|blob):/i.test(req.url())) problems.push(`The window made a network request: ${req.url()}`);
  });
  electronApp.on('console', (msg) => {
    transcript.push(`[main ${msg.type()}] ${msg.text()}`);
    if (msg.type() === 'error') problems.push(`Console error in the main process: ${msg.text()}`);
  });
  await until('the window to start', () => document.querySelector('#version-label')?.textContent);
}

// Opening a file from Explorer while the app is open starts it again; that copy hands the file to the
// open window and quits.
async function launchAgain(files) {
  const [command, args] = SOURCE ? [require('electron'), [ROOT]] : [EXE, []];
  const child = spawn(command, [...args, `--user-data-dir=${PROFILE}`, ...files], { env: appEnv(), stdio: 'ignore' });
  const exited = new Promise((resolve, reject) => {
    child.on('exit', resolve);
    child.on('error', reject);
  });
  try {
    await withTimeout(exited, 30000, 'The second launch didn’t quit within 30 s.');
  } catch (err) {
    child.kill();
    throw err;
  }
}

function kill() {
  if (mainPid && isRunning(mainPid)) {
    try { execFileSync('taskkill', ['/pid', String(mainPid), '/t', '/f'], { stdio: 'ignore', timeout: 60000 }); } catch { /* gone */ }
  }
}

// After a failed step: quits the app as closing its window would, or by force.
async function shutDown() {
  if (!electronApp || appClosed) return;
  await withTimeout(electronApp.close(), 15000).catch(() => {});
  kill();
}

function isRunning(pid) {
  try {
    process.kill(pid, 0);
    return true;
  } catch (err) {
    return err.code === 'EPERM';
  }
}

function childProcesses(pid) {
  const out = execFileSync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command',
    `Get-CimInstance Win32_Process -Filter "ParentProcessId=${pid}" | ForEach-Object { "$($_.ProcessId) $($_.Name)" }`],
  { encoding: 'utf8', timeout: 60000 });
  return out.trim().split(/\r?\n/).filter(Boolean).map((line) => {
    const [id, ...name] = line.trim().split(' ');
    return { pid: Number(id), name: name.join(' ') };
  });
}

/* ---------- Small helpers ---------- */

function withTimeout(promise, ms, message) {
  let timer;
  const late = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(message || `Timed out after ${ms / 1000} s.`)), ms);
  });
  return Promise.race([promise, late]).finally(() => clearTimeout(timer));
}

// Waits until `fn`, run in the window, returns something truthy.
async function until(what, fn, arg, timeout = 30000) {
  try {
    return await (await win.waitForFunction(fn, arg, { timeout, polling: 100 })).jsonValue();
  } catch (err) {
    throw new Error(`Waiting for ${what}: ${err.message.split('\n')[0]}`);
  }
}

// Waits until `fn`, run here, returns (or resolves to) something truthy.
async function eventually(what, fn, timeout = 10000) {
  const end = Date.now() + timeout;
  for (;;) {
    try {
      const value = await fn();
      if (value) return value;
    } catch { /* e.g. a file still being written */ }
    if (Date.now() > end) throw new Error(`Timed out after ${timeout / 1000} s waiting for ${what}.`);
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
}

const seconds = (since) => `${((Date.now() - since) / 1000).toFixed(1)} s`;

// Drops files and folders onto the preview, as dragging them from Explorer does.
async function drop(paths) {
  const cdp = await win.context().newCDPSession(win);
  const box = await win.locator('#pv-grid').boundingBox();
  const at = { x: box.x + box.width / 2, y: box.y + box.height / 2 };
  const data = { items: [], files: paths, dragOperationsMask: 1 };
  for (const type of ['dragEnter', 'dragOver', 'drop']) await cdp.send('Input.dispatchDragEvent', { type, ...at, data });
  await cdp.detach();
}

// The toast's text, closed once read. It hides itself after 7 s, so it's read before anything slower.
async function toastText() {
  await win.locator('#toast').waitFor();
  const text = await win.locator('#toast-text').textContent();
  await win.locator('#toast-close').click();
  return text;
}

const readSettings = () => JSON.parse(fs.readFileSync(path.join(PROFILE, 'settings.json'), 'utf8'));
const pick = (paths) => electronApp.evaluate((_electron, p) => { globalThis.uiTest.picks.push(p); }, paths);
const testCalls = () => electronApp.evaluate(() => globalThis.uiTest);

function assertXlsx(file) {
  assert.ok(fs.existsSync(file), `${path.basename(file)} wasn’t saved`);
  const head = Buffer.alloc(2);
  const fd = fs.openSync(file, 'r');
  fs.readSync(fd, head, 0, 2, 0);
  fs.closeSync(fd);
  assert.ok(head.toString() === 'PK' && fs.statSync(file).size > 2000, `${path.basename(file)} isn’t an Excel file`);
}

/* ---------- The file list ---------- */

const rowOf = (p) => win.locator('#list li').filter({ has: win.locator(`.file-row[title=${JSON.stringify(p)}]`) });

// A file's row: its status (loading, locked, ready, converting, done or error), its status text, and
// whether it's selected.
const rowState = (p) => win.evaluate((p) => {
  const li = [...document.querySelectorAll('#list li')].find((x) => x.querySelector('.file-row').title === p);
  return li && { status: li.classList[0], text: li.querySelector('.status').textContent, selected: li.classList.contains('selected') };
}, p);

// Waits until each file is in the list and done opening (or saving).
const settled = (paths, timeout = 120000) => until(paths.length === 1 ? `${path.basename(paths[0])} to open` : `${paths.length} files to open`, (ps) => {
  const rows = new Map([...document.querySelectorAll('#list li')].map((li) => [li.querySelector('.file-row').title, li.classList[0]]));
  return ps.every((p) => rows.has(p) && !['loading', 'converting'].includes(rows.get(p)));
}, paths, timeout);

const statusIs = (p, status, timeout = 60000) => until(`${path.basename(p)} to be ${status}`, ([p, s]) =>
  [...document.querySelectorAll('#list li')].some((li) => li.querySelector('.file-row').title === p && li.classList[0] === s),
[p, status], timeout);

async function select(p) {
  await rowOf(p).locator('.file-row').click();
  await until(`${path.basename(p)} to be selected`, (p) => document.querySelector('#list li.selected .file-row')?.title === p, p);
}

/* ---------- Checking a preview ---------- */

// Clicks through the selected file's sheets: each must show its cells, charts or drawings (or say it's
// empty), with no blank chart and no picture that doesn't load. Returns what the sheets drew in all.
async function walkSheets(label) {
  const tabs = win.locator('#pv-tabs .sheet-tab');
  const count = await tabs.count();
  assert.ok(count > 0, `${label} has no sheet tabs`);
  const totals = { sheets: count, charts: 0, drawings: 0, pictures: 0 };
  for (let i = 0; i < count; i++) {
    await tabs.nth(i).click();
    await until(`sheet ${i + 1} of ${label} to show`, (i) => document.querySelectorAll('#pv-tabs .sheet-tab')[i]?.getAttribute('aria-selected') === 'true', i);
    const sheet = await win.evaluate(async () => {
      const grid = document.querySelector('#pv-grid');
      const images = [...grid.querySelectorAll('.drawing.image img')];
      await Promise.all(images.map((img) => img.decode().catch(() => {})));
      const charts = [...grid.querySelectorAll('svg.chart')];
      return {
        name: document.querySelector('#pv-tabs [aria-selected="true"]').textContent,
        cells: grid.querySelectorAll('table.grid tbody td').length,
        empty: grid.querySelector('.grid-message')?.textContent === 'This sheet is empty.',
        charts: charts.length,
        blankCharts: charts.filter((chart) => !chart.childElementCount).length,
        drawings: grid.querySelectorAll('.drawing').length,
        pictures: grid.querySelectorAll('.drawing.image').length,
        brokenPictures: images.filter((img) => !img.naturalWidth).length
      };
    });
    const where = `${label}, sheet “${sheet.name}”,`;
    assert.ok(sheet.cells || sheet.charts || sheet.drawings || sheet.empty, `${where} shows nothing`);
    assert.strictEqual(sheet.blankCharts, 0, `${where} has a chart with nothing drawn in it`);
    assert.strictEqual(sheet.brokenPictures, 0, `${where} has a picture that doesn’t load`);
    for (const key of ['charts', 'drawings', 'pictures']) totals[key] += sheet[key];
  }
  return totals;
}

// The button beside Convert & save counts what to check before saving; "What changes in Excel" must list
// that many and say so. Returns the count.
async function checkChanges(label) {
  const button = win.locator('#pv-actions button[aria-controls="changes"]');
  const text = (await button.textContent()).trim();
  const count = text === 'Converts fully' ? 0 : Number(text.match(/^([\d,]+) changes?$/)?.[1].replace(/,/g, ''));
  assert.ok(Number.isInteger(count), `${label}: the button beside Convert & save says “${text}”`);
  await button.click();
  const panel = win.locator('#changes');
  await panel.waitFor();
  assert.strictEqual(await button.getAttribute('aria-expanded'), 'true');
  assert.strictEqual(await panel.locator('h2').textContent(), 'What changes in Excel');
  const listed = await panel.locator('.change-card.warn').count();
  assert.strictEqual(listed, count, `${label}: the button says “${text}”, What changes in Excel lists ${listed}`);
  const summary = await panel.locator('.summary').textContent();
  const says = count === 0 ? /^Everything in .+ carries over\./
    : count === 1 ? /Check this before you share/ : new RegExp(`Check these ${count} things before you share`);
  assert.match(summary, says, `${label}: What changes in Excel says “${summary}”`);
  assert.ok(await panel.locator('.kept li').count() > 0, `${label}: nothing is listed as carrying over`);
  await win.keyboard.press('Escape');
  await panel.waitFor({ state: 'hidden' });
  return count;
}

async function unlock(p, password) {
  const input = win.locator('#pw-input');
  await input.fill(password);
  await input.press('Enter');
  await statusIs(p, 'ready');
}

// Switches a Numbers sheet's tables to One sheet and back: its tables become one sheet tab, and the file
// remembers the choice.
async function checkLayoutSwitch(f, numbersSheet, expectedTabs) {
  await select(f.path);
  const tabs = win.locator('#pv-tabs .sheet-tab');
  const before = await tabs.count();
  const label = win.locator('#pv-layout .layout-label');
  const layoutLabel = () => win.evaluate(() => document.querySelector('#pv-layout .layout-label')?.textContent);
  for (let i = 0; i < before && await layoutLabel() !== `Tables on “${numbersSheet}”`; i++) await tabs.nth(i).click();
  assert.strictEqual(await layoutLabel(), `Tables on “${numbersSheet}”`, `${f.key}: no sheet tab shows the tables on “${numbersSheet}”`);
  const tables = Number((await label.getAttribute('title')).match(/^How the (\d+) tables/)[1]);
  const choice = (name) => win.locator('#pv-layout [role="radio"]').filter({ hasText: name });
  assert.strictEqual(await choice('Separate sheets').getAttribute('aria-checked'), 'true');

  await choice('One sheet').click();
  const after = before - tables + 1;
  await until(`${f.key} to show “${numbersSheet}” on one sheet`, (n) => document.querySelectorAll('#pv-tabs .sheet-tab').length === n
    && document.querySelector('#pv-layout [aria-checked="true"]')?.textContent.trim() === 'One sheet', after);
  assert.strictEqual(after, expectedTabs, `${f.key}: sheet tabs with “${numbersSheet}” on one sheet`);
  assert.strictEqual(await win.locator('#pv-tabs [aria-selected="true"]').textContent(), numbersSheet);
  await eventually(`${f.key}’s layout to be remembered`, () => readSettings().fileLayouts[f.path]?.[numbersSheet] === 'sheet');

  await choice('Separate sheets').click();
  await until(`${f.key} to show “${numbersSheet}” as separate sheets again`, (n) => document.querySelectorAll('#pv-tabs .sheet-tab').length === n, before);
  await eventually(`${f.key}’s layout to be remembered`, () => readSettings().fileLayouts[f.path]?.[numbersSheet] === 'tables');
}

// A file the app can't read says so, and shows the picture of its first sheet that Numbers saves in it.
async function checkUnreadable(p, label) {
  await select(p);
  const row = await rowState(p);
  assert.strictEqual(row.text, 'Couldn’t read this file', `${label}: “${row.text}”`);
  assert.strictEqual(await win.locator('#pv-grid .state-card h2').textContent(), 'The tables in this file couldn’t be read');
  const picture = await win.locator('#pv-grid figure.picture img').evaluate(async (img) => {
    await img.decode().catch(() => {});
    return { kind: img.src.slice(0, 23), width: img.naturalWidth };
  });
  assert.ok(picture.kind === 'data:image/jpeg;base64,' && picture.width > 0, `${label}: the picture saved in it doesn’t show`);
}

// A fixture's whole preview: every sheet, "What changes in Excel", and whatever expectations.js lists.
async function checkFile(f) {
  const expected = EXPECTED[f.key] || {};
  checked.add(f.key);
  if (expected.unreadable) {
    await checkUnreadable(f.path, f.key);
    return;
  }
  await select(f.path);
  let row = await rowState(f.path);
  if (row.status === 'locked' && expected.password) {
    await unlock(f.path, expected.password);
    row = await rowState(f.path);
  }
  assert.ok(['ready', 'done'].includes(row.status), `${f.key} didn’t open: “${row.text}”`);
  const totals = await walkSheets(f.key);
  for (const key of ['sheets', 'charts', 'drawings', 'pictures']) {
    if (key in expected) assert.strictEqual(totals[key], expected[key], `${f.key}: ${key} in the preview`);
  }
  const changes = await checkChanges(f.key);
  if ('changes' in expected) assert.strictEqual(changes, expected.changes, `${f.key}: changes to check`);
}

/* ---------- The steps ---------- */

const STEPS = [
  ['Files given on the command line open, as Open with opens them', async () => {
    const nine = fixture('issue-9.numbers');
    const locked = fixture('encrypted.numbers');
    await launch([nine.path, locked.path]);
    await settled([nine.path, locked.path]);
    assert.deepStrictEqual(await win.locator('#list .name').allTextContents(), ['issue-9.numbers', 'encrypted.numbers']);
    const row = await rowState(nine.path);
    assert.ok(row.status === 'ready' && row.selected && row.text.startsWith('Ready'), `issue-9.numbers: “${row.text}”`);
    assert.strictEqual((await rowState(locked.path)).text, 'Needs a password');
    assert.strictEqual(await win.locator('#pv-name').textContent(), 'issue-9.numbers');
    // This copy doesn't update itself, so it never checks GitHub.
    const update = await win.evaluate(() => window.api.getUpdateState());
    assert.strictEqual(update.status, 'unsupported');
    if (!SOURCE) assert.strictEqual(update.managedBy, 'portable');
  }],

  ['The preview draws sheet tabs, cells, charts and text boxes', async () => {
    const nine = fixture('issue-9.numbers');
    await select(nine.path);
    const totals = await walkSheets(nine.key);
    assert.ok(totals.charts > 0 && totals.drawings > 0, 'issue-9.numbers: no charts or text boxes drawn');
    for (const key of ['sheets', 'charts', 'drawings']) {
      assert.strictEqual(totals[key], EXPECTED[nine.key][key], `issue-9.numbers: ${key} in the preview`);
    }
    // Ctrl+PgDn and Ctrl+PgUp go to the next and previous sheet.
    await win.locator('#pv-tabs .sheet-tab').first().click();
    const shown = (i) => until(`sheet ${i + 1} to show`, (i) => document.querySelectorAll('#pv-tabs .sheet-tab')[i].getAttribute('aria-selected') === 'true', i);
    await win.keyboard.press('Control+PageDown');
    await shown(1);
    await win.keyboard.press('Control+PageUp');
    await shown(0);
  }],

  ['What changes in Excel lists as many changes as its button says', async () => {
    const nine = fixture('issue-9.numbers');
    await select(nine.path);
    assert.strictEqual(await checkChanges(nine.key), EXPECTED[nine.key].changes, 'issue-9.numbers: changes to check');
    checked.add(nine.key);
  }],

  ['A locked file shows its password hint, refuses a wrong password and opens with the right one', async () => {
    const locked = fixture('encrypted.numbers');
    const { password, hint } = EXPECTED[locked.key];
    await select(locked.path);
    await win.locator('#pw-input').waitFor();
    assert.strictEqual(await win.locator('#pv-grid h2').textContent(), 'This file is locked');
    assert.strictEqual(await win.locator('#pv-grid .hint-line').textContent(), `Hint: ${hint}`);
    assert.strictEqual(await win.locator('#pv-meta').textContent(), 'Password-protected');
    await win.locator('#pw-input').fill('not the password');
    await win.locator('#pw-input').press('Enter');
    await win.locator('#pv-grid .error-line').waitFor();
    assert.strictEqual(await win.locator('#pv-grid .error-line').textContent(), 'That password isn’t right. Try again.');
    assert.strictEqual(await win.locator('#pv-grid .hint-line').textContent(), `Hint: ${hint}`);
    await win.locator('#pw-input').fill(password);
    await win.getByRole('button', { name: 'Unlock' }).click();
    await statusIs(locked.path, 'ready');
    await win.locator('#pv-grid table.grid').waitFor();
  }],

  ['Opening a file while the app is open hands it to the open window', async () => {
    const formulas = fixture('create-formulas.numbers');
    await launchAgain([formulas.path]);
    await settled([formulas.path]);
    assert.ok((await rowState(formulas.path)).selected, 'create-formulas.numbers isn’t selected');
    assert.strictEqual(await win.locator('#list li').count(), 3);
  }],

  ['Separate sheets and One sheet switch a sheet’s layout, and the file remembers it', async () => {
    const formulas = fixture('create-formulas.numbers');
    await checkLayoutSwitch(formulas, 'Main Sheet', EXPECTED[formulas.key].oneSheet['Main Sheet']);
  }],

  ['Dropping files opens the Numbers files and skips the rest', async () => {
    const images = fixture('images.numbers');
    await drop([images.path, path.join(WORK, 'notes.txt')]);
    assert.strictEqual(await toastText(), 'Skipped 1 item that isn’t a Numbers file: notes.txt.');
    await settled([images.path]);
  }],

  ['Choose files (+ and Ctrl+O) opens files from the file dialog', async () => {
    const listed = new Set(await win.locator('#list .file-row').evaluateAll((rows) => rows.map((r) => r.title)));
    const rest = FILES.map((f) => f.path).filter((p) => !listed.has(p));
    await pick(rest);
    await win.locator('#add').click();
    await settled(rest, 180000);
    await pick([DAMAGED]);
    await win.keyboard.press('Control+o');
    await settled([DAMAGED]);
    assert.strictEqual(await win.locator('#list li').count(), FILES.length + 1);
    const { dialogs } = await testCalls();
    assert.deepStrictEqual(dialogs.map((d) => d.title), ['Choose Numbers files', 'Choose Numbers files']);
    assert.strictEqual(dialogs[1].defaultPath, path.dirname(rest[0]), 'the file dialog doesn’t open where files came from last');
  }],

  ['Every file opens, and each of its sheets draws without errors', async () => {
    for (const f of FILES.filter((x) => !checked.has(x.key))) {
      await checkFile(f);
      for (const [sheet, tabs] of Object.entries(EXPECTED[f.key]?.oneSheet || {})) await checkLayoutSwitch(f, sheet, tabs);
    }
  }],

  ['A file that can’t be read shows the picture Numbers saved in it', async () => {
    await checkUnreadable(DAMAGED, 'Damaged.numbers');
    const card = win.locator('#pv-grid .error-layout .state-card');
    assert.match(await card.locator('details.technical .mono').textContent(), /FileFormatError/);
    // Try again reads it again, with the same result.
    await card.getByRole('button', { name: 'Try again' }).click();
    await statusIs(DAMAGED, 'error');
    await win.locator('#pv-grid figure.picture img').waitFor();
    // Its problem report has the app's version and the error, never the file's name or folder.
    await win.locator('#pv-grid .state-card').getByRole('button', { name: 'Report a problem' }).click();
    await win.locator('#report-dialog').waitFor();
    const details = await win.locator('#report-details').textContent();
    assert.ok(details.startsWith(`Numbers Converter ${VERSION}, `), `the report starts “${details.split('\n')[0]}”`);
    assert.match(details, /FileFormatError/);
    assert.match(details, /File name and contents: not included/);
    for (const text of ['Damaged', 'Café', SCRATCH]) assert.ok(!details.includes(text), `the problem report includes “${text}”`);
    await win.locator('#report-cancel').click();
    await win.locator('#report-dialog').waitFor({ state: 'hidden' });
  }],

  ['Convert & save writes an .xlsx, and saving the file again adds “(1)”', async () => {
    const nine = fixture('issue-9.numbers');
    const first = path.join(WORK, 'issue-9.xlsx');
    const second = path.join(WORK, 'issue-9 (1).xlsx');
    await select(nine.path);
    assert.match(await win.locator('#pv-meta').textContent(), /· Saves as issue-9\.xlsx next to the original$/);
    await win.locator('#pv-actions').getByRole('button', { name: 'Convert & save' }).click();
    await statusIs(nine.path, 'done');
    assertXlsx(first);
    assert.strictEqual((await rowState(nine.path)).text, 'Saved as issue-9.xlsx');
    assert.strictEqual(await win.locator('#pv-meta').textContent(), `Saved as issue-9.xlsx in ${path.basename(SCRATCH)} › ${path.basename(WORK)}`);

    // Show in folder and Open in Excel reach `shell`, stubbed so nothing opens.
    const stubbed = await electronApp.evaluate(({ shell }) => shell.openPath.uiTestStub && shell.showItemInFolder.uiTestStub);
    assert.ok(stubbed, '`shell` isn’t stubbed, so Show in folder and Open in Excel weren’t clicked');
    await win.locator('#pv-actions').getByRole('button', { name: 'Show in folder' }).click();
    await win.locator('#pv-actions').getByRole('button', { name: 'Open in Excel' }).click();
    const calls = await eventually('Show in folder and Open in Excel to reach `shell`', async () => {
      const recorded = await testCalls();
      return recorded.revealed.length && recorded.opened.length && recorded;
    });
    assert.deepStrictEqual([calls.revealed, calls.opened], [[first], [first]]);

    // Opened afresh, the file is saved beside the first copy, which is left as it was.
    const saved = fs.statSync(first).mtimeMs;
    await win.keyboard.press('Delete');
    await until('issue-9.numbers to leave the list', (p) => ![...document.querySelectorAll('#list .file-row')].some((r) => r.title === p), nine.path);
    await pick([nine.path]);
    await win.locator('#add').click();
    await settled([nine.path]);
    assert.match(await win.locator('#pv-meta').textContent(), /· Saves as issue-9 \(1\)\.xlsx next to the original$/);
    await win.keyboard.press('Control+s');
    await statusIs(nine.path, 'done');
    assertXlsx(second);
    assert.strictEqual((await rowState(nine.path)).text, 'Saved as issue-9 (1).xlsx');
    assert.strictEqual(fs.statSync(first).mtimeMs, saved, 'issue-9.xlsx was written over');
  }],

  ['Convert all saves every file that’s ready', async () => {
    const ready = await win.locator('#list li.ready').count();
    assert.strictEqual(await win.locator('#convert-all').textContent(), `Convert all (${ready})`);
    await win.locator('#convert-all').click();
    await until('Convert all to finish', () => !document.querySelector('#list li.ready, #list li.converting'), null, STEP_TIMEOUT - 30000);
    const readable = FILES.filter((f) => !EXPECTED[f.key]?.unreadable);
    for (const f of readable) {
      const row = await rowState(f.path);
      assert.strictEqual(row.status, 'done', `${f.key}: “${row.text}”`);
      assertXlsx(path.join(path.dirname(f.path), row.text.replace(/^Saved as /, '')));
    }
    assert.strictEqual((await rowState(DAMAGED)).status, 'error');
    assert.strictEqual(await win.locator('#files-summary').textContent(), `${FILES.length + 1} · ${readable.length} saved`);
  }],

  ['Dropping a folder opens the Numbers files in it and its subfolders', async () => {
    const before = await win.locator('#list li').count();
    await drop([ARCHIVE, NOTHING]);
    assert.strictEqual(await toastText(), 'No Numbers files in Nothing here.');
    await settled(ARCHIVE_FILES);
    // In name order, a folder's own files before its subfolders'; what a Mac leaves behind isn't opened.
    const listed = await win.locator('#list .file-row').evaluateAll((rows) => rows.map((r) => r.title));
    assert.deepStrictEqual(listed.slice(before), ARCHIVE_FILES);
    assert.ok((await rowState(ARCHIVE_FILES[0])).selected, 'the first file of the folder isn’t selected');
    for (const p of ARCHIVE_FILES) assert.strictEqual((await rowState(p)).status, 'ready', `${p}: “${(await rowState(p)).text}”`);
  }],

  ['Add a folder opens a folder from the folder dialog', async () => {
    const charts = path.join(MORE, 'Charts.numbers');
    await pick([MORE]);
    await win.locator('#more').click();
    await win.locator('#menu-folder').click();
    await settled([charts]);
    assert.ok((await rowState(charts)).selected, 'Charts.numbers isn’t selected');
    const { dialogs } = await testCalls();
    assert.strictEqual(dialogs.at(-1).title, 'Choose a folder of Numbers files, or a Mac package folder');
  }],

  ['Settings opens, and its switches are saved', async () => {
    await win.locator('#settings-btn').click();
    await win.locator('#view-settings').waitFor();
    assert.strictEqual(await win.locator('#version-label').textContent(), `Version ${VERSION}`);
    const status = await win.locator('#update-status').textContent();
    assert.match(status, SOURCE ? /only available in the installed app/ : /runs without being installed/);
    assert.ok(await win.locator('#check').isDisabled(), 'Check for updates is on for a copy that doesn’t update itself');
    for (const on of [true, false]) {
      await win.locator('#contents').click();
      await until('the Contents sheet switch to flip', (on) => document.querySelector('#contents').getAttribute('aria-checked') === String(on), on);
      assert.strictEqual(await win.locator('#contents-word').textContent(), on ? 'On' : 'Off');
      await eventually('the Contents sheet setting to be saved', () => readSettings().contentsSheet === on);
    }
    // How to run the numbers-converter command that comes with the app, ready to copy.
    const command = win.locator('#cli-command');
    if (SOURCE) {
      assert.ok(await command.isHidden(), 'the app in src/ shows a command it doesn’t have');
    } else {
      const cmd = path.join(path.dirname(EXE), 'numbers-converter.cmd');
      assert.ok(fs.existsSync(cmd), 'numbers-converter.cmd isn’t beside the app');
      assert.strictEqual(await command.textContent(), `& "${cmd}" --help`);
      await win.locator('#cli-copy').click();
      const { copied } = await eventually('the command to be copied', async () => {
        const recorded = await testCalls();
        return recorded.copied.length && recorded;
      });
      assert.deepStrictEqual(copied, [`& "${cmd}" --help`]);
      await toastText();
    }
    await win.keyboard.press('Escape');
    await win.locator('#view-files').waitFor();
    await win.keyboard.press('Control+,');
    await win.locator('#view-settings').waitFor();
    await win.locator('#back').click();
    await win.locator('#view-files').waitFor();
  }],

  ['Clear the list can be undone', async () => {
    const count = await win.locator('#list li').count();
    await win.locator('#more').click();
    await win.locator('#menu-clear').click();
    await win.locator('#view-start').waitFor();
    assert.strictEqual(await win.locator('#toast-text').textContent(), `Cleared ${count} files from the list.`);
    await win.locator('#toast-action').click();
    await win.locator('#view-files').waitFor();
    assert.strictEqual(await win.locator('#list li').count(), count);
  }],

  ['A fresh start draws the start screen without errors', async () => {
    await win.reload();
    await win.locator('#view-start').waitFor();
    await until('the window to start', () => document.querySelector('#version-label').textContent);
    assert.strictEqual(await win.locator('#view-start h1').textContent(), 'Drop Numbers files or folders here');
    assert.strictEqual(await win.locator('#out-hint').textContent(), 'Excel copies are saved next to the originals.');
  }],

  ['Closing the window quits the app and stops the converter', async () => {
    const helper = childProcesses(mainPid).find((p) => /^(numbers-tool|python)\.exe$/i.test(p.name));
    assert.ok(helper, 'the converter isn’t running');
    const closed = electronApp.waitForEvent('close', { timeout: 30000 });
    await electronApp.evaluate(({ BrowserWindow }) => BrowserWindow.getAllWindows()[0].close()).catch(() => {});
    await closed;
    await eventually('the app to quit', () => !isRunning(mainPid), 15000);
    await eventually('the converter to stop', () => !isRunning(helper.pid), 15000);
    assert.ok(readSettings().window?.width >= 760, 'the window’s size wasn’t remembered');
  }]
];

/* ---------- Running the steps ---------- */

// Adds unhandled rejections in the main process, and what the app logged, to the problems. In the log, a
// file that couldn't be read is expected (the damaged one, or one listed as unreadable); anything else (an
// uncaught error, a failed save, the converter's error output, an update check) is a problem.
async function collectProblems() {
  if (electronApp && !appClosed) {
    const rejections = await withTimeout(electronApp.evaluate(() => globalThis.uiTest.rejections.splice(0)), 5000).catch(() => []);
    problems.push(...rejections.map((r) => `Unhandled rejection in the main process: ${r}`));
  }
  if (fs.existsSync(LOG)) {
    const entries = fs.readFileSync(LOG, 'utf8').split(/\r?\n(?=\d{4}-\d\d-\d\dT)/).filter(Boolean);
    const fresh = entries.slice(logLinesSeen);
    logLinesSeen = entries.length;
    const expected = /^\S+ open: Could not read this Numbers file /;
    problems.push(...fresh.filter((entry) => !expected.test(entry)).map((entry) => `The app logged: ${entry.trim()}`));
  }
}

async function saveEvidence(step, err) {
  fs.mkdirSync(RESULTS, { recursive: true });
  const base = path.join(RESULTS, step.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 60));
  let list = [];
  if (win && !appClosed) {
    await withTimeout(win.screenshot({ path: `${base}.png` }), 15000).catch((e) => console.log(`        (no screenshot: ${e.message})`));
    list = await withTimeout(win.evaluate(() => [...document.querySelectorAll('#list li')].map((li) =>
      `${li.querySelector('.file-row').title}: ${li.querySelector('.status').textContent} [${li.className}]`)), 5000).catch(() => []);
  }
  fs.writeFileSync(`${base}.txt`, [`Step: ${step}`, '', String(err.stack || err), '', ...problems, '',
    'Files in the list:', ...list, '', 'What the app printed:', ...transcript, ''].join('\n'));
  for (const [from, to] of [[path.join(PROFILE, 'logs'), 'logs'], [path.join(PROFILE, 'settings.json'), 'settings.json']]) {
    if (fs.existsSync(from)) fs.cpSync(from, path.join(RESULTS, to), { recursive: true });
  }
}

async function step(name, fn) {
  const started = Date.now();
  try {
    await withTimeout(fn(), STEP_TIMEOUT, `The step took over ${STEP_TIMEOUT / 60000} minutes.`);
    await collectProblems();
    if (problems.length) throw new Error(problems.splice(0).join('\n'));
    console.log(`  ok    ${name} (${seconds(started)})`);
  } catch (err) {
    console.log(`  FAIL  ${name} (${seconds(started)})`);
    await collectProblems().catch(() => {});
    for (const problem of problems) console.log(`        ${problem.split('\n')[0]}`); // often why the step failed
    await saveEvidence(name, err).catch((e) => console.log(`        (couldn’t save the evidence: ${e.message})`));
    throw err;
  }
}

(async () => {
  const started = Date.now();
  const tested = SOURCE ? 'the app in src/' : EXE.startsWith(ROOT) ? path.relative(ROOT, EXE) : EXE;
  console.log(`UI smoke test of ${tested}, with ${FILES.length} files from test/fixtures`);
  fs.rmSync(RESULTS, { recursive: true, force: true });
  // A last resort, within CI's 10 minutes for the step: if something hangs past every other timeout,
  // keep what there is to see and stop.
  setTimeout(async () => {
    console.error('\nThe UI smoke test took over 9 minutes, so it was stopped.');
    await saveEvidence('stopped after 9 minutes', new Error('The test took over 9 minutes.')).catch(() => {});
    kill();
    process.exit(1);
  }, 9 * 60 * 1000).unref();
  let failed = null;
  try {
    checkBuild();
    setUp();
    for (const [name, fn] of STEPS) await step(name, fn);
  } catch (err) {
    failed = err;
  } finally {
    await shutDown();
  }
  if (failed) {
    console.error(`\n${failed.stack || failed}\n\nScreenshot and logs: ${RESULTS}\nScratch folder: ${SCRATCH}`);
    process.exitCode = 1;
  } else {
    try {
      fs.rmSync(SCRATCH, { recursive: true, force: true, maxRetries: 10, retryDelay: 300 });
    } catch (err) {
      console.log(`  (couldn’t delete the scratch folder ${SCRATCH}: ${err.message})`);
    }
    console.log(`All UI smoke tests passed (${seconds(started)}).`);
  }
})();
