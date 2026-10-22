const $ = (sel) => document.querySelector(sel);

// Icons (16 × 16, drawn with the stroke styles in styles.css).
const ICONS = {
  file: '<svg width="16" height="16" viewBox="0 0 16 16"><path d="M9 1.75H4.5c-.69 0-1.25.56-1.25 1.25v10c0 .69.56 1.25 1.25 1.25h7c.69 0 1.25-.56 1.25-1.25V5.5z"/><path d="M9 1.75V5.5h3.75"/></svg>',
  saved: '<svg width="16" height="16" viewBox="0 0 16 16"><circle cx="8" cy="8" r="6.25"/><path d="m5.4 8.2 1.8 1.8 3.4-3.7"/></svg>',
  locked: '<svg width="16" height="16" viewBox="0 0 16 16"><rect x="3.25" y="7" width="9.5" height="7" rx="1.5"/><path d="M5.25 7V5a2.75 2.75 0 0 1 5.5 0v2"/></svg>',
  error: '<svg width="16" height="16" viewBox="0 0 16 16"><circle cx="8" cy="8" r="6.25"/><path d="M8 4.75V8.5"/><path d="M8 11.25h.01"/></svg>',
  warn: '<svg width="16" height="16" viewBox="0 0 16 16"><path d="M7.13 2.5 1.7 12a1 1 0 0 0 .87 1.5h10.86a1 1 0 0 0 .87-1.5L8.87 2.5a1 1 0 0 0-1.74 0z"/><path d="M8 6v3.25"/><path d="M8 11.5h.01"/></svg>',
  info: '<svg width="16" height="16" viewBox="0 0 16 16"><circle cx="8" cy="8" r="6.25"/><path d="M8 7.25V11"/><path d="M8 5h.01"/></svg>',
  check: '<svg width="16" height="16" viewBox="0 0 16 16"><path d="m3.5 8.5 3 3 6-7"/></svg>',
  x: '<svg width="16" height="16" viewBox="0 0 16 16"><path d="m4.25 4.25 7.5 7.5M11.75 4.25l-7.5 7.5"/></svg>',
  folder: '<svg width="16" height="16" viewBox="0 0 16 16"><path d="M1.75 4.25c0-.69.56-1.25 1.25-1.25h3.1l1.4 1.5H13c.69 0 1.25.56 1.25 1.25v6.25c0 .69-.56 1.25-1.25 1.25H3c-.69 0-1.25-.56-1.25-1.25z"/></svg>',
  open: '<svg width="16" height="16" viewBox="0 0 16 16"><path d="M9.5 2.25h4.25V6.5"/><path d="M13.75 2.25 7.75 8.25"/><path d="M12 9.5v3.25c0 .69-.56 1.25-1.25 1.25h-7.5C2.56 14 2 13.44 2 12.75v-7.5C2 4.56 2.56 4 3.25 4H6.5"/></svg>',
  flag: '<svg width="16" height="16" viewBox="0 0 16 16"><path d="M3.5 14.25V2.25"/><path d="M3.5 2.75h8.25L10 5.75l1.75 3H3.5"/></svg>',
  refresh: '<svg width="16" height="16" viewBox="0 0 16 16"><path d="M13.25 8a5.25 5.25 0 1 1-1.54-3.71"/><path d="M13.25 2.5v3.25H10"/></svg>',
  sep: '<svg width="16" height="16" viewBox="0 0 16 16"><rect x="1.75" y="3.25" width="5.25" height="9.5" rx="1"/><rect x="9" y="3.25" width="5.25" height="9.5" rx="1"/></svg>',
  one: '<svg width="16" height="16" viewBox="0 0 16 16"><rect x="1.75" y="2.25" width="12.5" height="11.5" rx="1.25"/><path d="M4.25 5.25h4M4.25 7.25h4M8.5 10h3.25M8.5 11.75h3.25"/></svg>'
};

// What the app says when a file can't be opened, by the helper's reason code.
const ERROR_TEXT = {
  unreadable: ['The tables in this file couldn’t be read',
    'The file may be damaged, or it was saved by a newer version of Numbers than this app can read yet. Nothing was saved, and the original file is untouched.'],
  'no-tables': ['This file has no tables',
    'Numbers Converter converts tables and the charts that go with them. This file doesn’t have any, so there’s nothing to save as an Excel copy.'],
  'not-numbers': ['This isn’t a Numbers file', 'Only files saved by Apple Numbers (.numbers) can be converted.'],
  stopped: ['The converter stopped unexpectedly', 'Try again. If it keeps happening, please report the problem.'],
  failed: ['This file couldn’t be converted', 'Something in it isn’t supported yet. Try again, or report the problem so it can be fixed.']
};

let settings = {};
let appInfo = { version: '', os: '' };
// Opened files: { id, path, name, status: loading|locked|ready|converting|done|error, stage (of an error:
//   open|save), preview (+ version, bumped on every new preview), sheet, output, autoSaved, error, details,
//   reason, picture, password (in memory only), hint, wrongPassword, layouts: { [Numbers sheet]: mode },
//   switching, target (where saving would write), scroll }
let files = [];
let selectedId = null;
let nextId = 1;
let view = 'main'; // or 'settings'
let panelOpen = false; // "What changes in Excel"
let shownBody = null; // what the preview area shows, so updates to other files don't rebuild (and scroll) it
let shownFile = null;
let cleared = null; // files removed by "Clear the list", kept for Undo
let reportText = '';

/* ---------- Small helpers ---------- */

const baseName = (p) => String(p).replace(/[\\/]+$/, '').split(/[\\/]/).pop();
const dirName = (p) => String(p).replace(/[\\/]+$/, '').replace(/[\\/][^\\/]*$/, '');
const stem = (name) => name.replace(/\.numbers$/i, '');
const isNumbersPath = (p) => /\.numbers[\\/]*$/i.test(p);
const selected = () => files.find((f) => f.id === selectedId);
const changeCount = (f) => f.preview?.report?.changes.length || 0;
const plural = (n, one, many = `${one}s`) => `${n.toLocaleString()} ${n === 1 ? one : many}`;

// "Documents › Reports": the last two folders of a path.
function folderLabel(dir) {
  const parts = String(dir).replace(/[\\/]+$/, '').split(/[\\/]/).filter(Boolean);
  return parts.slice(-2).join(' › ') || dir;
}

function listNames(names) {
  if (names.length <= 3) return names.length > 1 ? `${names.slice(0, -1).join(', ')} and ${names.at(-1)}` : names[0];
  return `${names.slice(0, 2).join(', ')} and ${names.length - 2} more`;
}

function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text != null) e.textContent = text;
  return e;
}

function icon(name, size = 16) {
  const span = el('span', 'ic');
  span.setAttribute('aria-hidden', 'true');
  span.innerHTML = ICONS[name].replace('width="16" height="16"', `width="${size}" height="${size}"`);
  return span;
}

function btn(text, cls, onClick, title, iconName) {
  const b = el('button', cls);
  b.type = 'button';
  if (iconName) b.append(icon(iconName));
  b.append(document.createTextNode(text));
  if (title) b.title = title;
  if (onClick) b.onclick = onClick;
  return b;
}

function iconButton(name, label, cls = 'icon-btn') {
  const b = el('button', cls);
  b.type = 'button';
  b.setAttribute('aria-label', label);
  b.append(icon(name, 14));
  return b;
}

/* ---------- Opening, previewing, converting ---------- */

async function openFiles(paths) {
  const accepted = paths.filter(isNumbersPath);
  const skipped = paths.filter((p) => !isNumbersPath(p));
  if (skipped.length) {
    const what = skipped.length === 1 ? '1 item that isn’t a Numbers file' : `${skipped.length} items that aren’t Numbers files`;
    toast(`Skipped ${what}: ${listNames(skipped.map(baseName))}.`);
  }
  if (!accepted.length) return;
  view = 'main';
  const added = accepted.map((p) => {
    const existing = files.find((f) => f.path === p);
    if (existing) {
      if (existing.status === 'error' && existing.stage === 'open') existing.status = 'loading'; // try it again
      return existing;
    }
    const f = { id: nextId++, path: p, name: baseName(p), status: 'loading', sheet: 0, version: 0,
                layouts: { ...(settings.fileLayouts?.[p] || {}) } };
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
const layoutOptions = (f, layouts = f.layouts) => ({ default: settings.sheetLayout || 'tables', sheets: layouts });

async function load(f) {
  f.status = 'loading';
  render();
  const res = await window.api.preview(f.path, f.password, layoutOptions(f));
  Object.assign(f, { error: null, details: null, reason: null, picture: null, stage: null });
  if (res.ok) {
    Object.assign(f, { preview: res, status: 'ready', wrongPassword: false });
    f.version++;
    f.target = await window.api.target(f.path);
  } else if (res.needsPassword) {
    Object.assign(f, { status: 'locked', wrongPassword: Boolean(f.password), hint: res.hint || f.hint });
    f.lockVersion = (f.lockVersion || 0) + 1;
  } else {
    Object.assign(f, { status: 'error', stage: 'open', error: res.error, details: res.details, reason: res.reason, picture: res.picture });
  }
  render();
  if (f.status === 'ready' && settings.autoConvert && files.includes(f)) await convert(f, { auto: true });
}

function unlock(f, password) {
  f.password = password;
  load(f);
}

function retry(f) {
  if (f.stage === 'save') convert(f);
  else load(f);
}

// `again`: the file's layout changed after it was saved automatically, so that copy is written over.
async function convert(f, { auto = false, again = false } = {}) {
  const canSave = f.status === 'ready' || (f.status === 'error' && f.stage === 'save');
  if (!canSave || !files.includes(f)) return;
  f.status = 'converting';
  render();
  const res = await window.api.convert(f.path, f.password, layoutOptions(f), { replace: again ? f.output : undefined });
  if (res.ok) {
    Object.assign(f, { status: 'done', stage: null, error: null, output: res.outputPath, autoSaved: auto });
  } else {
    Object.assign(f, { status: 'error', stage: 'save', error: res.error, details: res.details, reason: res.reason });
  }
  f.target = await window.api.target(f.path);
  render();
}

async function convertAll() {
  for (const f of files.filter((x) => x.status === 'ready')) await convert(f);
}

// Lays a Numbers sheet's tables out the other way, and shows the result.
async function switchLayout(f, numbersSheet, mode) {
  if (f.switching) return;
  f.switching = true;
  render();
  const layouts = { ...f.layouts, [numbersSheet]: mode };
  const res = await window.api.preview(f.path, f.password, layoutOptions(f, layouts));
  f.switching = false;
  if (!res.ok) {
    toast(`Couldn’t switch the layout. ${res.error}`);
    render();
    return;
  }
  f.layouts = layouts;
  rememberLayouts(f);
  Object.assign(f, { preview: res, sheet: Math.max(0, res.sheets.findIndex((s) => s.numbersSheet === numbersSheet)) });
  f.version++;
  const wasSaved = f.status === 'done';
  if (wasSaved) f.status = 'ready'; // the saved copy has the old layout
  render();
  if (wasSaved && f.autoSaved && settings.autoConvert) await convert(f, { auto: true, again: true });
}

// Per-file layout choices are remembered for the next time the file is opened (the last 200 files).
async function rememberLayouts(f) {
  const all = { ...(settings.fileLayouts || {}) };
  delete all[f.path];
  if (Object.keys(f.layouts).length) all[f.path] = f.layouts;
  const paths = Object.keys(all);
  for (const p of paths.slice(0, Math.max(0, paths.length - 200))) delete all[p];
  settings = await window.api.setSettings({ fileLayouts: all });
}

// When the default layout changes, files not saved yet are shown with it.
async function relayout() {
  for (const f of files.filter((x) => x.status === 'ready')) {
    const current = f.preview.sheets[f.sheet]?.numbersSheet;
    const res = await window.api.preview(f.path, f.password, layoutOptions(f));
    if (!res.ok) continue;
    Object.assign(f, { preview: res, sheet: Math.max(0, res.sheets.findIndex((s) => s.numbersSheet === current)) });
    f.version++;
  }
  render();
}

function select(f) {
  selectedId = f.id;
  render();
}

function removeFile(f) {
  if (!f || f.status === 'converting') return;
  const index = files.indexOf(f);
  files.splice(index, 1);
  if (selectedId === f.id) selectedId = files[Math.min(index, files.length - 1)]?.id ?? null;
  render();
  $('#list .selected .file-row')?.focus();
}

function clearAll() {
  hideMenu();
  if (!files.length || files.some((f) => f.status === 'converting')) return;
  cleared = { files, selectedId };
  const count = files.length;
  files = [];
  selectedId = null;
  render();
  toast(`Cleared ${plural(count, 'file')} from the list.`, {
    action: 'Undo',
    onAction: () => {
      files = cleared.files;
      selectedId = cleared.selectedId;
      cleared = null;
      render();
    }
  });
}

function stepFile(delta) {
  const index = files.findIndex((f) => f.id === selectedId);
  const next = files[Math.max(0, Math.min(files.length - 1, index + delta))];
  if (next && next.id !== selectedId) {
    select(next);
    $('#list .selected .file-row')?.focus();
  }
}

function stepSheet(delta) {
  const f = selected();
  if (!f?.preview || view !== 'main') return;
  const next = Math.max(0, Math.min(f.preview.sheets.length - 1, f.sheet + delta));
  if (next !== f.sheet) {
    f.sheet = next;
    render();
  }
}

/* ---------- Status text ---------- */

function statusText(f) {
  switch (f.status) {
    case 'loading': return 'Opening…';
    case 'converting': return 'Saving…';
    case 'done': return `Saved as ${baseName(f.output)}`;
    case 'locked': return 'Needs a password';
    case 'error':
      if (f.stage === 'save') return 'Couldn’t save';
      return f.reason === 'no-tables' ? 'No tables in this file' : 'Couldn’t read this file';
    default: {
      const n = changeCount(f);
      return n ? `Ready · ${plural(n, 'change')} to check` : 'Ready · converts fully';
    }
  }
}

function statusIcon(f) {
  if (f.status === 'loading' || f.status === 'converting') return '<span class="spinner"></span>';
  return ICONS[{ ready: 'file', done: 'saved', locked: 'locked', error: 'error' }[f.status]];
}

/* ---------- Rendering ---------- */

function render() {
  const hasFiles = files.length > 0;
  document.body.classList.toggle('has-files', hasFiles);
  $('#view-start').hidden = view !== 'main' || hasFiles;
  $('#view-files').hidden = view !== 'main' || !hasFiles;
  $('#view-settings').hidden = view !== 'settings';
  $('#settings-btn').classList.toggle('active', view === 'settings');
  if (!hasFiles) {
    shownBody = shownFile = null;
    return;
  }
  if (!selected()) selectedId = files[0].id;
  renderList();
  renderPreview(selected());
}

function renderList() {
  const done = files.filter((f) => f.status === 'done').length;
  const ready = files.filter((f) => f.status === 'ready').length;
  $('#files-summary').textContent = `${files.length}${done ? ` · ${done} saved` : ''}`;
  const all = $('#convert-all');
  all.hidden = ready < 2;
  all.textContent = `Convert all (${ready})`;

  const focusedId = document.activeElement?.closest?.('#list li')?.dataset.id;
  $('#list').replaceChildren(...files.map((f) => {
    const li = el('li', [f.status, f.id === selectedId ? 'selected' : ''].filter(Boolean).join(' '));
    li.dataset.id = f.id;
    const row = el('button', 'file-row');
    row.type = 'button';
    row.title = f.path;
    if (f.id === selectedId) row.setAttribute('aria-current', 'true');
    const statusBox = el('span', 'status-icon');
    statusBox.setAttribute('aria-hidden', 'true');
    statusBox.innerHTML = statusIcon(f);
    const text = el('span', 'text');
    text.append(el('span', 'name', f.name), el('span', 'status', statusText(f)));
    if (f.status === 'converting') text.append(el('span', 'busy-bar'));
    row.append(statusBox, text);
    row.onclick = () => select(f);
    li.append(row);
    if (f.id === selectedId && f.status !== 'converting') {
      const remove = iconButton('x', `Remove ${f.name} from the list`, 'icon-btn small remove');
      remove.title = 'Remove from list (Delete)';
      remove.onclick = () => removeFile(f);
      li.append(remove);
    }
    return li;
  }));
  if (focusedId) $(`#list li[data-id="${focusedId}"] .file-row`)?.focus();
}

function renderPreview(f) {
  $('#pv-name').textContent = f.name;
  renderMeta(f);
  renderActions(f);
  renderSaveError(f);
  renderBody(f);
  renderBottom(f);
  renderNote(f);
  renderChanges(f);
}

function renderMeta(f) {
  const meta = $('#pv-meta');
  meta.className = 'file-meta';
  meta.title = '';
  if (f.status === 'done') {
    meta.classList.add('done');
    meta.title = f.output;
    meta.replaceChildren(icon('check', 14), el('span', null, `Saved as ${baseName(f.output)} in ${folderLabel(dirName(f.output))}`));
    return;
  }
  let text;
  if (f.status === 'loading') text = 'Opening…';
  else if (f.status === 'locked') text = 'Password-protected';
  else if (!f.preview) {
    meta.classList.add('error');
    text = statusText(f);
  } else {
    const sheets = new Set(f.preview.sheets.map((s) => s.numbersSheet)).size;
    const where = settings.outputDir ? `in ${folderLabel(settings.outputDir)}` : 'next to the original';
    text = `${plural(sheets, 'Numbers sheet')} · ${f.status === 'converting' ? 'Saving' : 'Saves'} as ${baseName(f.target || 'an Excel copy')} ${where}`;
  }
  meta.replaceChildren(el('span', null, text));
}

function renderActions(f) {
  const items = [];
  const hasSheet = f.preview && f.status !== 'loading' && f.status !== 'locked';
  if (hasSheet) {
    const n = changeCount(f);
    const changes = n
      ? btn(plural(n, 'change'), 'btn warn', toggleChanges, 'See what will be different in Excel', 'warn')
      : btn('Converts fully', 'btn calm', toggleChanges, 'See what carries over to Excel', 'check');
    changes.classList.toggle('open', panelOpen);
    changes.setAttribute('aria-expanded', String(panelOpen));
    changes.setAttribute('aria-controls', 'changes');
    items.push(changes);
  }
  if (f.status === 'ready') {
    items.push(btn('Convert & save', 'btn primary', () => convert(f), 'Convert & save (Ctrl+S)'));
  } else if (f.status === 'converting') {
    const saving = btn('Saving…', 'btn primary');
    saving.disabled = true;
    items.push(saving);
  } else if (f.status === 'done') {
    items.push(btn('Show in folder', 'btn', () => window.api.reveal(f.output), 'Show in folder', 'folder'),
      btn('Open in Excel', 'btn primary', () => window.api.open(f.output), 'Open in Excel', 'open'));
  }
  $('#pv-actions').replaceChildren(...items);
}

function renderSaveError(f) {
  const bar = $('#pv-save-error');
  bar.hidden = !(f.status === 'error' && f.stage === 'save');
  if (bar.hidden) return;
  bar.replaceChildren(icon('error'), el('span', 'message', f.error),
    btn('Report a problem', 'btn', () => openReport(f)), btn('Try again', 'btn primary', () => convert(f), null, 'refresh'));
}

// What fills the preview area, as a key: when it hasn't changed, the area is left as it is.
function bodyKey(f) {
  if (f.status === 'loading') return `${f.id}|loading`;
  if (f.status === 'locked') return `${f.id}|locked|${f.lockVersion}`;
  if (!f.preview) return `${f.id}|error|${f.reason}|${f.error}`;
  return `${f.id}|sheet|${f.version}|${f.sheet}`;
}

function renderBody(f) {
  const grid = $('#pv-grid');
  const key = bodyKey(f);
  if (key === shownBody) return;
  if (shownFile && !grid.classList.contains('is-state')) {
    shownFile.scroll = { key: shownBody, left: grid.scrollLeft, top: grid.scrollTop };
  }
  shownBody = key;
  shownFile = f;
  grid.className = 'grid-wrap';
  if (f.status === 'loading') {
    const box = el('div', 'grid-message');
    box.append(el('div', 'spinner large'), document.createTextNode('Opening…'));
    stateView(grid, box);
  } else if (f.status === 'locked') {
    stateView(grid, passwordView(f));
    grid.querySelector('input').focus();
  } else if (!f.preview) {
    stateView(grid, errorView(f), true);
  } else {
    showSheet(f, grid, key);
  }
}

function stateView(grid, node, top = false) {
  grid.className = `grid-wrap is-state${top ? ' top' : ''}`;
  grid.replaceChildren(node);
}

function showSheet(f, grid, key) {
  const sheets = f.preview.sheets;
  f.sheet = Math.min(f.sheet, sheets.length - 1);
  const sheet = sheets[f.sheet];
  if (!sheet.rows.length && !sheet.charts?.length) {
    stateView(grid, el('div', 'grid-message', 'This sheet is empty.'));
    return;
  }
  const area = buildSheet(sheet);
  grid.replaceChildren(area);
  placeCharts(area);
  const back = f.scroll?.key === key ? f.scroll : null; // coming back to a sheet: where it was scrolled to
  grid.scrollTo(back?.left || 0, back?.top || 0);
}

// The password is only kept in memory for this file while the app is open; it is never saved.
function passwordView(f) {
  const form = el('form', 'state-card');
  const badge = el('div', 'state-icon');
  badge.append(icon('locked', 20));
  form.append(badge, el('h2', null, 'This file is locked'),
    el('p', 'muted', 'Enter its password to preview and convert it. The password is kept in memory only, and forgotten when you close the app.'));
  const field = el('div', 'field');
  const label = el('label', null, 'Password');
  label.htmlFor = 'pw-input';
  const input = el('input');
  Object.assign(input, { type: 'password', id: 'pw-input', autocomplete: 'off' });
  field.append(label, input);
  if (f.wrongPassword) field.append(el('div', 'error-line', 'That password isn’t right. Try again.'));
  if (f.hint) {
    const hint = el('div', 'hint-line');
    hint.append(icon('info', 14), el('span', null, `Hint: ${f.hint}`));
    field.append(hint);
  }
  const unlockButton = btn('Unlock', 'btn primary');
  unlockButton.type = 'submit';
  const actions = el('div', 'state-actions');
  actions.append(unlockButton, btn('Remove from list', 'btn', () => removeFile(f)));
  form.append(field, actions);
  form.onsubmit = (e) => {
    e.preventDefault();
    if (input.value) unlock(f, input.value);
  };
  return form;
}

function errorView(f) {
  const [title, text] = ERROR_TEXT[f.reason] || [f.error || 'This file couldn’t be opened', 'Try again, or report the problem so it can be fixed.'];
  const layout = el('div', 'error-layout');
  const card = el('div', 'state-card');
  const badge = el('div', 'state-icon bad');
  badge.append(icon('error', 20));
  const actions = el('div', 'state-actions');
  actions.append(btn('Try again', 'btn primary', () => retry(f), null, 'refresh'),
    btn('Report a problem', 'btn', () => openReport(f), null, 'flag'),
    btn('Remove from list', 'btn', () => removeFile(f)));
  card.append(badge, el('h2', null, title), el('p', 'muted', text), actions);
  if (f.details) {
    const details = el('details', 'technical');
    details.append(el('summary', null, 'Technical details'), el('div', 'mono', f.details));
    card.append(details);
  }
  layout.append(card);
  if (f.picture) {
    const figure = el('figure', 'picture');
    const img = el('img');
    img.src = f.picture;
    img.alt = `Picture of ${stem(f.name)} that Numbers saved inside the file`;
    figure.append(el('div', 'caption-title', 'Picture saved inside the file'), img,
      el('figcaption', null, 'Numbers keeps a picture of the first sheet in every file. You can look at it here, but it can’t be edited or converted.'));
    layout.append(figure);
  }
  return layout;
}

function renderBottom(f) {
  const bar = $('#pv-bottom');
  bar.hidden = !(f.preview && f.status !== 'loading' && f.status !== 'locked');
  if (bar.hidden) return;
  const sheets = f.preview.sheets;
  const tabs = $('#pv-tabs');
  const tabFocused = tabs.contains(document.activeElement);
  tabs.replaceChildren(...sheets.map((s, i) => {
    const tab = btn(s.name, 'sheet-tab', () => { f.sheet = i; render(); });
    tab.setAttribute('role', 'tab');
    tab.setAttribute('aria-selected', String(i === f.sheet));
    return tab;
  }));
  const active = tabs.querySelector('[aria-selected="true"]');
  active?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  if (tabFocused) active?.focus();
  const sheet = sheets[Math.min(f.sheet, sheets.length - 1)];
  $('#pv-layout').replaceChildren(...(sheet.tableCount > 1 ? [layoutSwitch(f, sheet)] : []));
}

// A Numbers sheet with several tables can become an Excel sheet per table, or one sheet laid out as in Numbers.
function layoutSwitch(f, sheet) {
  const box = el('div', 'layout-switch');
  if (f.switching) box.append(el('span', 'spinner'));
  const label = el('span', 'layout-label', `Tables on “${sheet.numbersSheet}”`);
  label.title = `How the ${sheet.tableCount} tables on the Numbers sheet “${sheet.numbersSheet}” are saved`;
  label.id = 'layout-label';
  const group = el('div', 'segmented small');
  group.setAttribute('role', 'radiogroup');
  group.setAttribute('aria-labelledby', 'layout-label');
  const choices = [['tables', 'Separate sheets', 'sep', 'Each table on its own Excel sheet'],
    ['sheet', 'One sheet', 'one', 'All the tables together, placed as in Numbers']];
  for (const [mode, text, iconName, tip] of choices) {
    const b = btn(text, '', () => { if (sheet.layout !== mode) switchLayout(f, sheet.numbersSheet, mode); }, tip, iconName);
    b.setAttribute('role', 'radio');
    b.setAttribute('aria-checked', String(sheet.layout === mode));
    b.disabled = f.status === 'converting' || Boolean(f.switching);
    group.append(b);
  }
  box.append(label, group);
  return box;
}

function renderNote(f) {
  const note = $('#pv-note');
  const sheet = f.preview && f.status !== 'loading' && f.status !== 'locked' ? f.preview.sheets[f.sheet] : null;
  const rows = sheet?.rows.length || 0;
  const cols = sheet?.rows[0]?.length || 0;
  note.hidden = !sheet || (rows >= sheet.totalRows && cols >= sheet.totalCols);
  if (!note.hidden) {
    note.textContent = `The preview shows the first ${rows.toLocaleString()} of ${sheet.totalRows.toLocaleString()} rows `
      + `and ${cols} of ${sheet.totalCols} columns. The saved file has everything.`;
  }
}

/* ---------- What changes in Excel ---------- */

function toggleChanges() {
  panelOpen = !panelOpen;
  render();
  if (panelOpen) $('#changes .changes-head button')?.focus();
}

function renderChanges(f) {
  const panel = $('#changes');
  const report = f.preview?.report;
  panel.hidden = !(panelOpen && report && f.status !== 'loading' && f.status !== 'locked');
  if (panel.hidden) {
    delete panel.dataset.key;
    return;
  }
  const key = `${f.id}|${f.version}`;
  if (panel.dataset.key === key) return; // unchanged: keep its scroll position
  panel.dataset.key = key;

  const head = el('div', 'changes-head');
  const close = iconButton('x', 'Close');
  close.title = 'Close (Esc)';
  close.onclick = toggleChanges;
  head.append(el('h2', null, 'What changes in Excel'), close);

  const body = el('div', 'changes-body');
  const n = report.changes.length;
  const name = stem(f.name);
  let summary = `Everything in ${name} carries over.`;
  if (n) summary = `Almost everything in ${name} carries over. Check ${n === 1 ? 'this' : `these ${n} things`} before you share the Excel file.`;
  else if (report.notes.length) summary += ' A few things look a little different:';
  body.append(el('p', 'summary', summary));
  for (const item of report.changes) body.append(changeCard(f, item, true));
  for (const item of report.notes) body.append(changeCard(f, item, false));
  const kept = el('div', 'kept');
  const keptTitle = el('div', 'kept-title');
  keptTitle.append(icon('saved'), el('span', null, 'Carries over as it is'));
  const list = el('ul');
  report.kept.forEach((line) => list.append(el('li', null, line)));
  kept.append(keptTitle, list);
  body.append(kept);
  panel.replaceChildren(head, body, el('div', 'changes-foot', 'Checked on this PC before anything is saved.'));
}

function changeCard(f, item, isChange) {
  const card = el('div', `change-card${isChange ? ' warn' : ''}`);
  const title = el('div', 'change-title');
  title.append(icon(isChange ? 'warn' : 'info'), el('span', null, item.title));
  card.append(title, el('div', 'change-detail', item.detail));
  if (item.reasons?.length) {
    const reasons = el('ul');
    item.reasons.forEach((r) => reasons.append(el('li', null, r)));
    card.append(reasons);
  }
  if (item.cells?.length) {
    const chips = el('div', 'cell-chips');
    for (const cell of item.cells) {
      const chip = el('button', 'cell-chip');
      chip.type = 'button';
      chip.title = 'Show this cell in the preview';
      chip.append(el('span', null, `${cell.sheet} › ${cell.ref}`), el('span', 'show', 'Show'));
      chip.onclick = () => showCell(f, cell);
      chips.append(chip);
    }
    if (item.more > 0) chips.append(el('span', 'more', `and ${item.more.toLocaleString()} more`));
    card.append(chips);
  }
  return card;
}

function showCell(f, cell) {
  const index = f.preview.sheets.findIndex((s) => s.name === cell.sheet);
  if (index < 0) return;
  if (f.sheet !== index) {
    f.sheet = index;
    render();
  }
  const grid = $('#pv-grid');
  const td = grid.querySelector(`td[data-r="${cell.row}"][data-c="${cell.col}"]`);
  if (!td) {
    toast('That cell is hidden, or past the part of the sheet the preview shows. It’s in the saved file.');
    return;
  }
  const visibleWidth = grid.clientWidth - (panelOpen ? $('#changes').offsetWidth : 0);
  grid.scrollTo(Math.max(0, td.offsetLeft + td.offsetWidth / 2 - visibleWidth / 2),
    Math.max(0, td.offsetTop + td.offsetHeight / 2 - grid.clientHeight / 2));
  td.classList.remove('flash');
  void td.offsetWidth; // restart the animation
  td.classList.add('flash');
}

/* ---------- The sheet preview ---------- */

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
      td.dataset.r = r;
      td.dataset.c = c;
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

/* ---------- Toasts, menu, problem reports ---------- */

let toastTimer = null;
function toast(text, { action, onAction, good = false, timeout = 7000 } = {}) {
  const box = $('#toast');
  box.classList.toggle('good', good);
  $('#toast-text').textContent = text;
  const button = $('#toast-action');
  button.hidden = !action;
  if (action) {
    button.textContent = action;
    button.onclick = () => { hideToast(); onAction(); };
  }
  box.hidden = false;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(hideToast, timeout);
}
function hideToast() {
  $('#toast').hidden = true;
}

function hideMenu() {
  $('#more-menu').hidden = true;
  $('#more').setAttribute('aria-expanded', 'false');
}

function openReport(f) {
  let problem = 'Not about a particular file (describe it on GitHub)';
  let error = 'none';
  if (f) {
    problem = f.stage === 'save' ? 'Saving an Excel copy failed' : (ERROR_TEXT[f.reason] || [f.error])[0];
    error = f.details || f.error || 'none';
  }
  reportText = [`Numbers Converter ${appInfo.version}, ${appInfo.os}`, `Problem: ${problem}`, `Error: ${error}`,
    'File name and contents: not included'].join('\n');
  $('#report-details').textContent = reportText;
  $('#report-dialog').hidden = false;
  $('#report-open').focus();
}
function closeReport() {
  $('#report-dialog').hidden = true;
}

$('#report-cancel').onclick = closeReport;
$('#report-copy').onclick = async () => {
  await window.api.copy(reportText);
  toast('Copied the details.', { good: true });
};
$('#report-open').onclick = () => {
  window.api.reportProblem(reportText);
  closeReport();
  toast('Opened a new GitHub issue in your browser.', { good: true });
};
$('#toast-close').onclick = hideToast;

/* ---------- Buttons, drag & drop, keyboard ---------- */

const pick = async () => openFiles(await window.api.pickFiles());
const pickPackages = async () => {
  hideMenu();
  openFiles(await window.api.pickPackages());
};

function showSettings(on) {
  view = on ? 'settings' : 'main';
  hideMenu();
  render();
  (on ? $('#back') : $('#settings-btn')).focus();
}

$('#pick').onclick = pick;
$('#add').onclick = pick;
$('#pick-packages').onclick = pickPackages;
$('#menu-packages').onclick = pickPackages;
$('#menu-clear').onclick = clearAll;
$('#convert-all').onclick = convertAll;
$('#settings-btn').onclick = () => showSettings(view !== 'settings');
$('#back').onclick = () => showSettings(false);
$('#change-out').onclick = () => showSettings(true);
$('#more').onclick = (e) => {
  e.stopPropagation();
  const menu = $('#more-menu');
  menu.hidden = !menu.hidden;
  $('#more').setAttribute('aria-expanded', String(!menu.hidden));
  if (!menu.hidden) menu.querySelector('button').focus();
};
document.addEventListener('click', (e) => {
  if (!e.target.closest('.menu-wrap')) hideMenu();
});
document.querySelectorAll('[data-link]').forEach((b) => { b.onclick = () => window.api.openLink(b.dataset.link); });

let dragDepth = 0;
window.addEventListener('dragenter', (e) => { e.preventDefault(); if (dragDepth++ === 0) document.body.classList.add('dragging'); });
window.addEventListener('dragleave', () => { if (--dragDepth === 0) document.body.classList.remove('dragging'); });
window.addEventListener('dragover', (e) => e.preventDefault());
window.addEventListener('drop', (e) => {
  e.preventDefault();
  dragDepth = 0;
  document.body.classList.remove('dragging');
  openFiles([...e.dataTransfer.files].map((f) => window.api.pathForFile(f)).filter(Boolean));
});
window.api.onFilesOpened(openFiles);

document.addEventListener('keydown', (e) => {
  const ctrl = e.ctrlKey || e.metaKey;
  const key = e.key.toLowerCase();
  const typing = Boolean(e.target.closest?.('input, textarea'));
  const dialogOpen = !$('#report-dialog').hidden;
  if (e.key === 'Escape') {
    if (dialogOpen) closeReport();
    else if (!$('#more-menu').hidden) { hideMenu(); $('#more').focus(); }
    else if (panelOpen && view === 'main') toggleChanges();
    else if (view === 'settings') showSettings(false);
    else return;
    e.preventDefault();
    return;
  }
  if (dialogOpen) return;
  if (ctrl && !e.shiftKey && key === 'o') pick();
  else if (ctrl && key === ',') showSettings(view !== 'settings');
  else if (ctrl && e.shiftKey && key === 's') convertAll();
  else if (ctrl && key === 's') { const f = selected(); if (f && view === 'main') convert(f); }
  else if (ctrl && (e.key === 'PageDown' || e.key === 'PageUp')) stepSheet(e.key === 'PageDown' ? 1 : -1);
  else if (!ctrl && !typing && view === 'main' && (e.key === 'ArrowDown' || e.key === 'ArrowUp')
           && (e.target === document.body || e.target.closest?.('#list'))) stepFile(e.key === 'ArrowDown' ? 1 : -1);
  else if (!ctrl && !typing && view === 'main' && e.key === 'Delete' && !e.target.closest?.('.menu')) removeFile(selected());
  else return;
  e.preventDefault();
});

/* ---------- Settings ---------- */

function setSwitch(id, on) {
  $(`#${id}`).setAttribute('aria-checked', String(Boolean(on)));
  $(`#${id}-word`).textContent = on ? 'On' : 'Off';
}

function renderSettings() {
  setSwitch('auto-convert', settings.autoConvert);
  setSwitch('contents', settings.contentsSheet);
  setSwitch('auto-update', settings.autoUpdate);
  document.querySelectorAll('#sheet-layout button').forEach((b) => {
    b.setAttribute('aria-checked', String(b.dataset.layout === (settings.sheetLayout || 'tables')));
  });
  const custom = settings.outputDir;
  const rule = 'Nothing is ever replaced: if a name is taken, the copy gets “(1)” added.';
  $('#out-desc').textContent = custom ? `In ${custom}. ${rule}` : `Next to the original file. ${rule}`;
  $('#reset-out').hidden = !custom;
  $('#out-hint').textContent = custom ? `Excel copies are saved in ${folderLabel(custom)}.` : 'Excel copies are saved next to the originals.';
  $('#version-label').textContent = `Version ${appInfo.version}`;
  const source = el('button', 'link-btn small', 'Source code on GitHub');
  source.onclick = () => window.api.openLink('https://github.com/silent-diffusion/numbers-converter');
  $('#about').replaceChildren(document.createTextNode(`Numbers Converter ${appInfo.version} · Free and open source (MIT licence) · `), source);
}

async function saveSetting(patch) {
  settings = await window.api.setSettings(patch);
  renderSettings();
}

// Saving somewhere else changes where (and under what name) each file would be saved.
async function refreshTargets() {
  for (const f of files.filter((x) => x.preview)) f.target = await window.api.target(f.path);
  render();
}

$('#auto-convert').onclick = () => saveSetting({ autoConvert: !settings.autoConvert });
$('#contents').onclick = () => saveSetting({ contentsSheet: !settings.contentsSheet });
$('#auto-update').onclick = async () => {
  await saveSetting({ autoUpdate: !settings.autoUpdate });
  renderUpdate(await window.api.getUpdateState());
};
document.querySelectorAll('#sheet-layout button').forEach((b) => {
  b.onclick = async () => {
    if (b.dataset.layout === (settings.sheetLayout || 'tables')) return;
    await saveSetting({ sheetLayout: b.dataset.layout });
    relayout();
  };
});
$('#choose-out').onclick = async () => {
  const dir = await window.api.pickFolder();
  if (!dir) return;
  await saveSetting({ outputDir: dir });
  refreshTargets();
};
$('#reset-out').onclick = async () => {
  await saveSetting({ outputDir: null });
  refreshTargets();
};
$('#report-general').onclick = () => openReport(null);
$('#open-log').onclick = () => window.api.openLogFolder();

/* ---------- Updates ---------- */

function timeAgo(ts) {
  const mins = Math.round((Date.now() - ts) / 60000);
  if (mins < 1) return 'just now';
  if (mins < 60) return `${plural(mins, 'minute')} ago`;
  const hours = Math.round(mins / 60);
  return hours < 24 ? `${plural(hours, 'hour')} ago` : `on ${new Date(ts).toLocaleDateString()}`;
}

function renderUpdate(s) {
  const status = $('#update-status');
  const check = $('#check');
  const checked = s.checkedAt ? ` Last checked ${timeAgo(s.checkedAt)}.` : '';
  const ready = s.status === 'downloaded';
  $('#install').hidden = !ready;
  $('#update-pill').hidden = !ready;
  if (ready) $('#update-pill').title = `Version ${s.version} is ready. Restart to finish updating.`;
  $('#progress').hidden = s.status !== 'downloading';
  check.disabled = ['checking', 'downloading', 'unsupported'].includes(s.status);
  check.hidden = ready;
  $('#auto-update').disabled = Boolean(s.managedBy); // Scoop, or a portable copy: updated another way

  switch (s.status) {
    case 'checking': status.textContent = 'Checking for updates…'; break;
    case 'downloading':
      status.textContent = `Downloading version ${s.version}…`;
      $('#progress-bar').style.width = `${s.percent || 0}%`;
      break;
    case 'downloaded': status.textContent = `Version ${s.version} is ready to install. Restart to finish, or it installs when you close the app.`; break;
    case 'up-to-date': status.textContent = `You’re up to date.${checked}`; break;
    case 'error': status.textContent = `${s.message}${checked}`; break;
    case 'unsupported': status.textContent = s.message; break;
    default: status.textContent = settings.autoUpdate ? 'Updates are checked automatically.' : 'Automatic updates are off.';
  }
}

$('#check').onclick = () => window.api.checkForUpdates();
$('#install').onclick = () => window.api.installUpdate();
$('#update-pill').onclick = () => window.api.installUpdate();
window.api.onUpdateState(renderUpdate);

/* ---------- Start ---------- */

(async () => {
  [settings, appInfo] = await Promise.all([window.api.getSettings(), window.api.info()]);
  renderSettings();
  renderUpdate(await window.api.getUpdateState());
  render();
  setInterval(async () => renderUpdate(await window.api.getUpdateState()), 60000); // keep "last checked" fresh
})();
