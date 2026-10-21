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

function startHelper() {
  const [cmd, args] = helperCommand();
  child = spawn(cmd, args, { stdio: ['pipe', 'pipe', 'ignore'], windowsHide: true });
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
  const fail = () => {
    child = null;
    for (const req of pending.values()) req.resolve({ ok: false, error: 'The converter stopped unexpectedly. Please try again.' });
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

function previewFile(inputPath) {
  return request({ cmd: 'preview', path: inputPath });
}

function convertFile(inputPath, outputDir) {
  const dir = outputDir || path.dirname(inputPath);
  const base = path.basename(inputPath, path.extname(inputPath));
  return request({ cmd: 'convert', path: inputPath, out: uniquePath(dir, base, '.xlsx') });
}

module.exports = { convertFile, previewFile, stopHelper };
