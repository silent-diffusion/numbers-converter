const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

// The Numbers reader/Excel writer is a Python helper (py/numbers_tool.py). The installed app ships it
// as a PyInstaller build in resources/numbers-tool; in development it runs from the project's .venv.
function helperCommand() {
  if (process.resourcesPath) {
    const exe = path.join(process.resourcesPath, 'numbers-tool', 'numbers-tool.exe');
    if (fs.existsSync(exe)) return [exe, []];
  }
  const root = path.join(__dirname, '..');
  return [path.join(root, '.venv', 'Scripts', 'python.exe'), ['-u', path.join(root, 'py', 'numbers_tool.py')]];
}

let child = null;
let nextId = 1;
const pending = new Map();
let logError = () => {}; // where the helper's error output goes (the app's local log)

function setLogger(fn) {
  logError = fn;
}

function startHelper() {
  const [cmd, args] = helperCommand();
  child = spawn(cmd, args, { stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true });
  let buffer = '';
  child.stdout.setEncoding('utf8');
  child.stdout.on('data', (chunk) => {
    buffer += chunk;
    let nl;
    while ((nl = buffer.indexOf('\n')) >= 0) {
      const line = buffer.slice(0, nl);
      buffer = buffer.slice(nl + 1);
      let msg;
      try { msg = JSON.parse(line); } catch { continue; }
      const req = pending.get(msg.id);
      if (req) { pending.delete(msg.id); req.resolve(msg); }
    }
  });
  child.stderr.setEncoding('utf8');
  child.stderr.on('data', (chunk) => logError(chunk));
  const fail = () => {
    child = null;
    for (const req of pending.values()) {
      req.resolve({ ok: false, reason: 'stopped', error: 'The converter stopped unexpectedly. Please try again.' });
    }
    pending.clear();
  };
  child.on('exit', fail);
  child.on('error', fail);
}

function request(payload) {
  if (!child) startHelper();
  const id = nextId++;
  return new Promise((resolve) => {
    pending.set(id, { resolve });
    child.stdin.write(JSON.stringify({ id, ...payload }) + '\n');
  });
}

function stopHelper() {
  if (child) child.kill();
  child = null;
}

// Returns a path in `dir` named after `base` that doesn't exist yet: "Budget.xlsx", "Budget (1).xlsx", ...
function uniquePath(dir, base, ext) {
  let candidate = path.join(dir, base + ext);
  for (let i = 1; fs.existsSync(candidate); i++) {
    candidate = path.join(dir, `${base} (${i})${ext}`);
  }
  return candidate;
}

// Where saving the file now would write its Excel copy: next to it (or in `outputDir`), never over a file.
function targetPath(inputPath, outputDir) {
  const dir = outputDir || path.dirname(inputPath);
  const base = path.basename(inputPath.replace(/[\\/]+$/, ''), '.numbers');
  return uniquePath(dir, base, '.xlsx');
}

// `layout` says how Numbers sheets with several tables are laid out:
// { default: 'tables' | 'sheet', sheets: { [Numbers sheet name]: 'tables' | 'sheet' } }
// ('tables': a worksheet per table; 'sheet': one worksheet per Numbers sheet, laid out as in Numbers).
function previewFile(inputPath, password, layout) {
  return request({ cmd: 'preview', path: inputPath, password, layout });
}

// `options.contents` adds a first sheet linking to every table. `options.replace` is a copy of this file
// the app saved earlier, written over (rather than adding a "(1)" copy) when it saves the file again.
function convertFile(inputPath, outputDir, password, layout, options = {}) {
  const out = options.replace || targetPath(inputPath, outputDir);
  return request({ cmd: 'convert', path: inputPath, out, password, layout, contents: Boolean(options.contents) });
}

module.exports = { convertFile, previewFile, targetPath, stopHelper, setLogger };
