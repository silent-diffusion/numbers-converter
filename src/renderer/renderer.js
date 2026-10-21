const $ = (sel) => document.querySelector(sel);

const ICONS = {
  ok: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>',
  error: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="M12 8v5M12 16.5v.5"/></svg>',
  busy: '<div class="spinner"></div>'
};

let settings = {};
let counts = { ok: 0, error: 0 };

const baseName = (p) => p.split(/[\\/]/).pop();
const dirName = (p) => p.slice(0, Math.max(p.lastIndexOf('\\'), p.lastIndexOf('/')));

/* ---------- Navigation ---------- */

function showView(name) {
  document.querySelectorAll('.tab').forEach((t) => t.classList.toggle('active', t.dataset.view === name));
  document.querySelectorAll('.view').forEach((v) => (v.hidden = v.id !== `view-${name}`));
}

function showPage(name) {
  document.querySelectorAll('.subtab').forEach((t) => t.classList.toggle('active', t.dataset.page === name));
  document.querySelectorAll('.page').forEach((p) => (p.hidden = p.id !== `page-${name}`));
}

document.querySelectorAll('.tab').forEach((t) => t.addEventListener('click', () => showView(t.dataset.view)));
document.querySelectorAll('.subtab').forEach((t) => t.addEventListener('click', () => showPage(t.dataset.page)));

/* ---------- Converting ---------- */

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text != null) e.textContent = text;
  return e;
}

function updateSummary() {
  const parts = [];
  if (counts.ok) parts.push(`${counts.ok} converted`);
  if (counts.error) parts.push(`${counts.error} failed`);
  $('#results-summary').textContent = parts.join(' · ') || 'Converting…';
}

async function convertAll(paths) {
  const files = paths.filter((p) => p.toLowerCase().endsWith('.numbers'));
  if (!files.length) return;
  showView('convert');
  $('#results').hidden = false;
  $('#view-convert').classList.add('has-results');

  for (const input of files) {
    const li = el('li');
    const icon = el('div', 'status-icon');
    icon.innerHTML = ICONS.busy;
    const info = el('div', 'file');
    info.append(el('div', 'file-name', baseName(input)), el('div', 'file-meta', 'Converting…'));
    const actions = el('div', 'file-actions');
    li.append(icon, info, actions);
    $('#list').prepend(li);

    const res = await window.api.convert(input);
    const meta = info.querySelector('.file-meta');
    if (res.ok) {
      counts.ok++;
      li.className = 'ok';
      icon.innerHTML = ICONS.ok;
      meta.textContent = `${baseName(res.outputPath)} · ${res.sheets} sheet${res.sheets === 1 ? '' : 's'}`;
      meta.title = res.outputPath;
      const open = el('button', 'btn link', 'Open');
      open.onclick = () => window.api.open(res.outputPath);
      const show = el('button', 'btn link', 'Show in folder');
      show.onclick = () => window.api.reveal(res.outputPath);
      actions.append(open, show);
    } else {
      counts.error++;
      li.className = 'error';
      icon.innerHTML = ICONS.error;
      meta.textContent = res.error;
    }
    updateSummary();
  }
}

const drop = $('#drop');
['dragenter', 'dragover'].forEach((ev) =>
  drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach((ev) =>
  drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', (e) => {
  convertAll([...e.dataTransfer.files].map((f) => window.api.pathForFile(f)));
});
// Ignore files dropped outside the drop zone instead of navigating to them.
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', (e) => e.preventDefault());

$('#pick').onclick = async () => convertAll(await window.api.pickFiles());
$('#clear').onclick = () => {
  $('#list').replaceChildren();
  counts = { ok: 0, error: 0 };
  $('#results').hidden = true;
  $('#view-convert').classList.remove('has-results');
};
window.api.onFilesOpened(convertAll);

/* ---------- General settings ---------- */

function renderOutput() {
  const custom = settings.outputDir;
  $('#out-desc').textContent = custom || 'In the same folder as the original file';
  $('#out-hint').textContent = custom ? `Excel files will be saved to ${custom}` : 'Excel files are saved next to the originals';
  $('#reset-out').hidden = !custom;
}

$('#choose-out').onclick = async () => {
  const dir = await window.api.pickFolder();
  if (dir) { settings = await window.api.setSettings({ outputDir: dir }); renderOutput(); }
};
$('#reset-out').onclick = async () => {
  settings = await window.api.setSettings({ outputDir: null });
  renderOutput();
};

/* ---------- Updates ---------- */

function timeAgo(ts) {
  const mins = Math.round((Date.now() - ts) / 60000);
  if (mins < 1) return 'just now';
  if (mins < 60) return `${mins} min ago`;
  const hrs = Math.round(mins / 60);
  return hrs < 24 ? `${hrs} hr ago` : new Date(ts).toLocaleDateString();
}

function renderUpdate(s) {
  const status = $('#update-status');
  const check = $('#check');
  const checked = s.checkedAt ? ` Last checked ${timeAgo(s.checkedAt)}.` : '';
  $('#install').hidden = s.status !== 'downloaded';
  $('#progress').hidden = s.status !== 'downloading';
  $('#update-dot').hidden = s.status !== 'downloaded';
  check.disabled = ['checking', 'downloading', 'unsupported'].includes(s.status);
  check.hidden = s.status === 'downloaded';

  switch (s.status) {
    case 'checking':    status.textContent = 'Checking for updates…'; break;
    case 'downloading': status.textContent = `Downloading version ${s.version}…`;
                        $('#progress-bar').style.width = `${s.percent || 0}%`; break;
    case 'downloaded':  status.textContent = `Version ${s.version} is ready. Restart to finish updating, or it will install when you close the app.`; break;
    case 'up-to-date':  status.textContent = `You're up to date.${checked}`; break;
    case 'error':       status.textContent = `${s.message}${checked}`; break;
    case 'unsupported': status.textContent = s.message; break;
    default:            status.textContent = settings.autoUpdate ? 'Updates will be checked automatically.' : 'Automatic updates are off.';
  }
}

$('#auto-update').onchange = async (e) => {
  settings = await window.api.setSettings({ autoUpdate: e.target.checked });
  renderUpdate(await window.api.getUpdateState());
};
$('#check').onclick = () => window.api.checkForUpdates();
$('#install').onclick = () => window.api.installUpdate();
window.api.onUpdateState(renderUpdate);

/* ---------- Init ---------- */

(async () => {
  settings = await window.api.getSettings();
  const { version } = await window.api.info();
  $('#version').textContent = version;
  $('#auto-update').checked = settings.autoUpdate;
  renderOutput();
  renderUpdate(await window.api.getUpdateState());
  setInterval(async () => renderUpdate(await window.api.getUpdateState()), 60000); // keep "last checked" fresh
})();
