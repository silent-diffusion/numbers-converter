const $ = (sel) => document.querySelector(sel);

const ICONS = {
  ready: '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/></svg>',
  done: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>',
  error: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><circle cx="12" cy="12" r="9"/><path d="M12 8v5M12 16.5v.5"/></svg>',
  locked: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/></svg>',
  busy: '<div class="spinner"></div>'
};

let settings = {};
// Opened files: { id, path, name, status: loading|locked|ready|converting|done|error, preview, sheet,
//                 output, error, password (kept in memory only), hint,
//                 layouts: { [Numbers sheet name]: 'tables' | 'sheet' } chosen in the preview }
let files = [];
let selectedId = null;
let nextId = 1;

const baseName = (p) => p.split(/[\\/]/).pop();

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text != null) e.textContent = text;
  return e;
}

function button(label, cls, onClick) {
  const b = el('button', cls, label);
  b.onclick = onClick;
  return b;
}

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

/* ---------- Opening, previewing, converting ---------- */

async function openFiles(paths) {
  const fresh = paths.filter((p) => p.toLowerCase().endsWith('.numbers'));
  if (!fresh.length) return;
  showView('convert');

  const added = fresh.map((p) => {
    const existing = files.find((f) => f.path === p);
    if (existing) return existing;
    const f = { id: nextId++, path: p, name: baseName(p), status: 'loading', sheet: 0, layouts: {} };
    files.push(f);
    return f;
  });
  selectedId = added[0].id;
  render();

  for (const f of added) {
    if (f.status === 'loading') await load(f);
  }
}

// How the file's Numbers sheets with several tables are laid out: the setting, unless switched in the preview.
const layoutOptions = (f) => ({ default: settings.sheetLayout || 'tables', sheets: f.layouts });

async function load(f) {
  const res = await window.api.preview(f.path, f.password, layoutOptions(f));
  if (res.ok) {
    f.preview = res;
    f.status = 'ready';
  } else if (res.needsPassword) {
    f.status = 'locked';
    f.error = f.password ? res.error : null;
    f.hint = res.hint || f.hint;
  } else {
    f.status = 'error';
    f.error = res.error;
  }
  render();
  if (f.status === 'ready' && settings.autoConvert) await convert(f);
}

async function unlock(f, password) {
  f.password = password;
  f.status = 'loading';
  render();
  await load(f);
}

async function convert(f) {
  if (f.status !== 'ready') return;
  f.status = 'converting';
  render();
  const res = await window.api.convert(f.path, f.password, layoutOptions(f));
  if (res.ok) {
    f.status = 'done';
    f.output = res.outputPath;
    f.formulasAsValues = res.formulasAsValues;
  } else {
    f.status = 'error';
    f.error = res.error;
  }
  render();
}

// Lays a Numbers sheet's tables out the other way, and shows the result.
async function switchLayout(f, numbersSheet, mode) {
  f.layouts = { ...f.layouts, [numbersSheet]: mode };
  const res = await window.api.preview(f.path, f.password, layoutOptions(f));
  if (!res.ok) return;
  f.preview = res;
  f.sheet = Math.max(0, res.sheets.findIndex((s) => s.numbersSheet === numbersSheet));
  if (f.status === 'done') f.status = 'ready'; // the saved copy has the old layout; offer to save again
  render();
}

// When the setting changes, files not saved yet are shown with it.
async function relayout() {
  for (const f of files.filter((x) => x.status === 'ready')) {
    const current = f.preview.sheets[f.sheet]?.numbersSheet;
    const res = await window.api.preview(f.path, f.password, layoutOptions(f));
    if (!res.ok) continue;
    f.preview = res;
    f.sheet = Math.max(0, res.sheets.findIndex((s) => s.numbersSheet === current));
  }
  render();
}

async function convertAll() {
  for (const f of files.filter((x) => x.status === 'ready')) await convert(f);
}

function metaText(f) {
  switch (f.status) {
    case 'loading': return 'Opening…';
    case 'converting': return 'Saving…';
    case 'done': return `Saved as ${baseName(f.output)}`;
    case 'locked': return 'Password needed';
    case 'error': return f.error;
    default: {
      const n = f.preview.sheets.length;
      return `${n} sheet${n === 1 ? '' : 's'} · Not converted yet`;
    }
  }
}

/* ---------- Rendering ---------- */

function render() {
  const hasFiles = files.length > 0;
  $('#drop').hidden = hasFiles;
  $('#workspace').hidden = !hasFiles;
  if (!hasFiles) return;

  const done = files.filter((f) => f.status === 'done').length;
  const ready = files.filter((f) => f.status === 'ready').length;
  $('#files-summary').textContent = `${files.length} file${files.length === 1 ? '' : 's'}${done ? ` · ${done} saved` : ''}`;
  $('#convert-all').hidden = ready < 2;

  $('#list').replaceChildren(...files.map((f) => {
    const li = el('li', `${f.status}${f.id === selectedId ? ' selected' : ''}`);
    const icon = el('div', 'status-icon');
    icon.innerHTML = ICONS[f.status] || ICONS.busy;
    const info = el('div', 'file');
    info.append(el('div', 'file-name', f.name), el('div', 'file-meta', metaText(f)));
    li.title = f.path;
    li.append(icon, info);
    li.onclick = () => { selectedId = f.id; render(); };
    return li;
  }));

  renderPreview(files.find((f) => f.id === selectedId) || files[0]);
}

function renderPreview(f) {
  $('#pv-name').textContent = f.name;
  const meta = $('#pv-meta');
  meta.textContent = f.status === 'done' ? `Saved to ${f.output}` : f.path;
  meta.title = meta.textContent;
  meta.className = `file-meta${f.status === 'done' ? ' done' : ''}`;

  const actions = $('#pv-actions');
  actions.replaceChildren();
  if (f.status === 'ready') {
    actions.append(button('Convert & save', 'btn primary', () => convert(f)));
  } else if (f.status === 'converting') {
    const b = button('Saving…', 'btn primary');
    b.disabled = true;
    actions.append(b);
  } else if (f.status === 'done') {
    actions.append(
      button('Show in folder', 'btn', () => window.api.reveal(f.output)),
      button('Open in Excel', 'btn primary', () => window.api.open(f.output))
    );
  }

  const tabs = $('#pv-tabs');
  const grid = $('#pv-grid');
  const note = $('#pv-note');
  tabs.replaceChildren();
  note.hidden = true;

  if (f.status === 'loading') {
    grid.replaceChildren(el('div', 'grid-message', 'Opening…'));
    return;
  }
  if (f.status === 'locked') {
    grid.replaceChildren(passwordForm(f));
    grid.querySelector('input').focus();
    return;
  }
  if (!f.preview) {
    grid.replaceChildren(el('div', 'grid-message error', f.error));
    return;
  }

  const sheets = f.preview.sheets;
  f.sheet = Math.min(f.sheet, sheets.length - 1);
  const sheet = sheets[f.sheet];
  if (sheets.length > 1) {
    const list = el('div', 'tab-list'); // scrolls on its own, so the layout switch stays in view
    sheets.forEach((s, i) => {
      list.append(button(s.name, `sheet-tab${i === f.sheet ? ' active' : ''}`, () => { f.sheet = i; renderPreview(f); }));
    });
    tabs.append(list);
    list.querySelector('.active')?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }
  if (sheet.tableCount > 1) tabs.append(layoutSwitch(f, sheet));

  if (!sheet.rows.length && !sheet.charts?.length) {
    grid.replaceChildren(el('div', 'grid-message', 'This sheet is empty.'));
  } else {
    const area = buildSheet(sheet);
    grid.replaceChildren(area);
    placeCharts(area);
    grid.scrollTo(0, 0);
  }

  const notes = [];
  const shownRows = sheet.rows.length;
  const shownCols = sheet.rows[0]?.length || 0;
  if (shownRows < sheet.totalRows || shownCols < sheet.totalCols) {
    notes.push(`Preview shows the first ${shownRows.toLocaleString()} of ${sheet.totalRows.toLocaleString()} rows and ${shownCols} of ${sheet.totalCols} columns. The saved file includes everything.`);
  }
  const missing = describeUnsupported(f.preview.unsupported);
  if (missing) {
    const pivotNote = f.preview.unsupported?.pivots ? ' Those pivot tables keep their labels, but their totals come out blank.' : '';
    notes.push(`This file has ${missing}, which can't be converted. Tables, formatting and formulas are kept.${pivotNote}`);
  }
  if (f.formulasAsValues) {
    notes.push(`${f.formulasAsValues} formula${f.formulasAsValues === 1 ? '' : 's'} used Numbers-only features and ${f.formulasAsValues === 1 ? 'was' : 'were'} saved as ${f.formulasAsValues === 1 ? 'its value' : 'values'}.`);
  }
  note.hidden = !notes.length;
  note.replaceChildren(...notes.map((n) => el('div', null, n)));
}

// A Numbers sheet with several tables can become an Excel sheet per table, or one sheet laid out as in Numbers.
function layoutSwitch(f, sheet) {
  const box = el('div', 'layout-switch');
  box.title = `How the ${sheet.tableCount} tables on the Numbers sheet "${sheet.numbersSheet}" are saved: `
    + 'each on its own Excel sheet, or together on one sheet, laid out as in Numbers.';
  box.append(el('span', 'layout-label', `${sheet.numbersSheet}:`));
  const group = el('div', 'segmented small');
  group.setAttribute('role', 'radiogroup');
  for (const [mode, label] of [['tables', 'Separate sheets'], ['sheet', 'One sheet']]) {
    const b = button(label, sheet.layout === mode ? 'active' : '', () => {
      if (sheet.layout !== mode) switchLayout(f, sheet.numbersSheet, mode);
    });
    b.setAttribute('role', 'radio');
    b.setAttribute('aria-checked', String(sheet.layout === mode));
    b.disabled = f.status === 'converting';
    group.append(b);
  }
  box.append(group);
  return box;
}

// The password is only kept in memory for this file while the app is open; it is never saved.
function passwordForm(f) {
  const form = el('form', 'unlock');
  form.append(ICON(ICONS.locked, 'unlock-icon'), el('div', 'unlock-title', 'This file is password-protected'));
  const input = el('input');
  input.type = 'password';
  input.placeholder = 'Password';
  input.autocomplete = 'off';
  const submit = el('button', 'btn primary', 'Open');
  submit.type = 'submit';
  const row = el('div', 'unlock-row');
  row.append(input, submit);
  form.append(row);
  if (f.error) form.append(el('div', 'unlock-error', f.error));
  if (f.hint) form.append(el('div', 'unlock-hint', `Hint: ${f.hint}`));
  form.onsubmit = (e) => {
    e.preventDefault();
    if (input.value) unlock(f, input.value);
  };
  return form;
}

function ICON(svg, cls) {
  const span = el('span', cls);
  span.innerHTML = svg;
  return span;
}

function describeUnsupported(u = {}) {
  const parts = [];
  const add = (n, one, many) => { if (n) parts.push(`${n} ${n === 1 ? one : many}`); };
  add(u.charts, 'chart', 'charts');
  add(u.images, 'image', 'images');
  add(u.shapes, 'shape', 'shapes');
  add(u.pivots, 'pivot table', 'pivot tables');
  if (!parts.length) return '';
  return parts.length === 1 ? parts[0] : `${parts.slice(0, -1).join(', ')} and ${parts.at(-1)}`;
}

function columnName(i) {
  let s = '';
  for (i++; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
  return s;
}

const PT_TO_PX = 4 / 3;
const BORDER_STYLES = { solid: 'solid', dashes: 'dashed', dots: 'dotted' };

// Apply a cell style from the helper (see cell_style in py/numbers_tool.py) as inline CSS.
function applyStyle(td, s) {
  if (!s) return;
  const st = td.style;
  if (s.b) st.fontWeight = '600';
  if (s.i) st.fontStyle = 'italic';
  const deco = [s.u && 'underline', s.st && 'line-through'].filter(Boolean).join(' ');
  if (deco) st.textDecoration = deco;
  if (s.color) st.color = s.color;
  if (s.bg) st.background = s.bg;
  if (s.size) st.fontSize = `${s.size * PT_TO_PX}px`;
  if (s.font) st.fontFamily = `"${s.font}", Arial, sans-serif`;
  if (s.ha) st.textAlign = s.ha;
  if (s.va) st.verticalAlign = s.va;
  if (s.wrap) { st.whiteSpace = 'pre-wrap'; st.wordBreak = 'break-word'; }
  for (const [key, side] of [['bt', 'Top'], ['br', 'Right'], ['bb', 'Bottom'], ['bl', 'Left']]) {
    if (!s[key]) continue;
    const [width, color, kind] = s[key];
    st[`border${side}`] = `${Math.max(1, Math.round(width * PT_TO_PX))}px ${BORDER_STYLES[kind] || 'solid'} ${color}`;
  }
}

// A sheet as it will look in Excel: the table, with its charts drawn over the sheet where they'll be.
function buildSheet(sheet) {
  const area = el('div', 'sheet-area');
  if (sheet.rows.length) area.append(buildTable(sheet));
  if (sheet.charts?.length) {
    const layer = el('div', 'chart-layer');
    for (const chart of sheet.charts) {
      let drawing;
      try {
        drawing = drawChart(chart);
      } catch (err) {
        console.error('Could not draw a chart preview', err); // the saved file still has it
        continue;
      }
      drawing.style.left = `${chart.x}px`;
      drawing.style.top = `${chart.y}px`;
      layer.append(drawing);
    }
    area.append(layer);
  }
  return area;
}

// Chart positions are measured from the sheet's first cell (A1), which sits below and right of
// the row and column labels; once the sheet is on the page, line the charts up with it.
function placeCharts(area) {
  const layer = area.querySelector('.chart-layer');
  if (!layer) return;
  const first = area.querySelector('tbody td');
  if (first) {
    const origin = area.getBoundingClientRect();
    const cell = first.getBoundingClientRect();
    layer.style.left = `${cell.left - origin.left}px`;
    layer.style.top = `${cell.top - origin.top}px`;
  }
}

function buildTable(sheet) {
  // Hidden rows and columns are left out, as in Numbers; row and column labels keep their numbering.
  const hiddenRows = new Set(sheet.hiddenRows || []);
  const hiddenCols = new Set(sheet.hiddenCols || []);
  const visibleCols = sheet.colWidths.map((_, c) => c).filter((c) => !hiddenCols.has(c));

  const table = el('table', 'grid');
  const cols = el('colgroup');
  cols.append(el('col', 'rowhead-col'));
  visibleCols.forEach((c) => {
    const col = el('col');
    col.style.width = `${sheet.colWidths[c] * PT_TO_PX}px`; // sizes are in points, like the fonts
    cols.append(col);
  });
  table.append(cols);

  const headRow = el('tr');
  headRow.append(el('th', 'corner'));
  visibleCols.forEach((c) => headRow.append(el('th', null, columnName(c))));
  const head = el('thead');
  head.append(headRow);
  table.append(head);

  // Merged cells span the rows and columns that are showing.
  const visible = (from, to, hidden) => Array.from({ length: to - from + 1 }, (_, i) => from + i).filter((i) => !hidden.has(i)).length;
  const spans = new Map(sheet.merges.map(([r1, c1, r2, c2]) => [`${r1},${c1}`, [visible(r1, r2, hiddenRows), visible(c1, c2, hiddenCols)]]));
  const body = el('tbody');
  sheet.rows.forEach((row, r) => {
    if (hiddenRows.has(r)) return;
    const tr = el('tr');
    tr.style.height = `${sheet.rowHeights[r] * PT_TO_PX}px`;
    tr.append(el('th', null, String(r + 1)));
    row.forEach((cell, c) => {
      if (cell === null || hiddenCols.has(c)) return; // covered by a merged cell, or hidden
      const [text, styleIdx, extra = {}] = typeof cell === 'string' ? [cell, 0] : cell;
      const td = el('td', null, extra.runs ? null : text);
      applyStyle(td, sheet.styles[styleIdx]);
      if (extra.runs) {
        for (const [piece, runStyle] of extra.runs) {
          const span = el('span', null, piece);
          applyStyle(span, sheet.styles[runStyle]);
          td.append(span);
        }
      }
      if (extra.link) {
        td.classList.add('has-link');
        td.title = extra.link;
      }
      if (extra.note) {
        td.classList.add('has-note');
        td.title = extra.note;
      }
      if (extra.control === 'popup') td.classList.add('has-popup');
      const span = spans.get(`${r},${c}`);
      if (span) { td.rowSpan = Math.max(1, span[0]); td.colSpan = Math.max(1, span[1]); }
      if (!td.title && text.length > 40) td.title = text;
      tr.append(td);
    });
    body.append(tr);
  });
  table.append(body);
  return table;
}

/* ---------- Drag & drop, buttons ---------- */

let dragDepth = 0;
window.addEventListener('dragenter', (e) => { e.preventDefault(); if (dragDepth++ === 0) document.body.classList.add('dragging'); });
window.addEventListener('dragleave', () => { if (--dragDepth === 0) document.body.classList.remove('dragging'); });
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', (e) => {
  e.preventDefault();
  dragDepth = 0;
  document.body.classList.remove('dragging');
  openFiles([...e.dataTransfer.files].map((f) => window.api.pathForFile(f)));
});

const pick = async () => openFiles(await window.api.pickFiles());
$('#pick').onclick = pick;
$('#add').onclick = pick;
$('#convert-all').onclick = convertAll;
$('#clear').onclick = () => { files = []; selectedId = null; render(); };
window.api.onFilesOpened(openFiles);

/* ---------- General settings ---------- */

function renderOutput() {
  const custom = settings.outputDir;
  $('#out-desc').textContent = custom || 'In the same folder as the original file';
  $('#out-hint').textContent = custom ? `Excel files will be saved to ${custom}` : 'Excel files are saved next to the originals';
  $('#reset-out').hidden = !custom;
}

$('#auto-convert').onchange = async (e) => {
  settings = await window.api.setSettings({ autoConvert: e.target.checked });
};

function renderSheetLayout() {
  document.querySelectorAll('#sheet-layout button').forEach((b) => {
    const on = b.dataset.layout === (settings.sheetLayout || 'tables');
    b.classList.toggle('active', on);
    b.setAttribute('aria-checked', String(on));
  });
}
document.querySelectorAll('#sheet-layout button').forEach((b) => {
  b.onclick = async () => {
    if (b.dataset.layout === settings.sheetLayout) return;
    settings = await window.api.setSettings({ sheetLayout: b.dataset.layout });
    renderSheetLayout();
    relayout();
  };
});
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
  $('#auto-convert').checked = settings.autoConvert;
  renderSheetLayout();
  renderOutput();
  renderUpdate(await window.api.getUpdateState());
  setInterval(async () => renderUpdate(await window.api.getUpdateState()), 60000); // keep "last checked" fresh
})();
