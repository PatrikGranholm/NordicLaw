// Improved: parse a Dating string to a year (start of range if present, ignore text)
function parseDatingYear(dating) {
  if (!dating || typeof dating !== 'string') return null;
  // Match a year range (e.g. 1350-1375)
  let m = dating.match(/(\d{3,4})\s*[-–]\s*(\d{3,4})/);
  if (m) return parseInt(m[1], 10);
  // Match a single 4-digit year
  m = dating.match(/(\d{4})/);
  if (m) return parseInt(m[1], 10);
  // Match a 3-digit century (e.g. 1200s)
  m = dating.match(/(\d{3})00s/);
  if (m) return parseInt(m[1] + '50', 10); // use mid-century
  // Match '14th century', '13th cent.'
  m = dating.match(/(\d{1,2})(?:th|st|nd|rd)\s*cent/i);
  if (m) return (parseInt(m[1], 10) - 1) * 100 + 50; // mid-century
  return null;
}

function escapeHtml(text) {
  if (text === null || text === undefined) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function toDomIdToken(text) {
  if (text === null || text === undefined) return '';
  return String(text)
    .trim()
    .replace(/[^A-Za-z0-9_-]+/g, '_');
}

function normalizeLeavesPages(value) {
  if (value === null || value === undefined) return '';
  const s = String(value);
  // Ensure exactly one space after f./p. when immediately followed by a non-space character.
  // Examples: "f.1r" -> "f. 1r", "p.12" -> "p. 12", "f. 1r" stays unchanged.
  return s.replace(/\b([fFpP])\.\s*/g, '$1. ');
}

function normalizeGatherings(value) {
  if (value === null || value === undefined) return '';
  let s = String(value).replace(/\r/g, '').trim();
  if (!s) return '';
  // Collapse whitespace/newlines first.
  s = s.replace(/\s+/g, ' ').trim();
  // If Excel export used spaces between gathering blocks, convert ") II:" -> "); II:".
  s = s.replace(/\)\s+(?=[IVXLCDM]+:)/g, '); ');
  return s;
}

function normalizeScribe(value) {
  if (value === null || value === undefined) return '';
  let s = String(value).replace(/\r/g, '').trim();
  if (!s) return '';
  s = s.replace(/\s+/g, ' ').trim();

  // Normalize to '; '-separated segments for repeated hand markers.
  const markerRe = /(Hand\s+[A-Za-z]\b|Marginal:)/g;
  const matches = Array.from(s.matchAll(markerRe));
  if (matches.length <= 1) return s;

  const starts = matches
    .map(m => (typeof m.index === 'number' ? m.index : -1))
    .filter(i => i >= 0)
    .sort((a, b) => a - b);
  if (starts.length <= 1) return s;

  const parts = [];
  const prefix = s.slice(0, starts[0]).trim();
  if (prefix) parts.push(prefix.replace(/;+\s*$/g, '').trim());

  for (let i = 0; i < starts.length; i++) {
    const start = starts[i];
    const end = (i + 1 < starts.length) ? starts[i + 1] : s.length;
    let part = s.slice(start, end).trim();
    part = part.replace(/^;+\s*/g, '');
    part = part.replace(/;+\s*$/g, '').trim();
    if (part) parts.push(part);
  }

  return parts.join('; ');
}

// Canonical column order for UI display (matches the TSV export order, with Language injected).
const COLUMN_ORDER = [
  "Depository",
  "Shelf mark",
  "Language",
  "Name",
  "Object",
  "Size",
  "Production Unit",
  "Leaves/Pages",
  "Main text",
  "Minor text",
  "Dating",
  "Gatherings",
  "Full size",
  "Leaf size",
  "Catch Words and Gatherings",
  "Pricking",
  "Material",
  "Ruling",
  "Columns",
  "Lines",
  "Script",
  "Rubric",
  "Scribe",
  "Production",
  "Style",
  "Colours",
  "Form of Initials",
  "Size of Initials",
  "Iconography",
  "Place",
  "Related Shelfmarks",
  "Literature",
  "Links to Database",
];

// Dataset globals (populated by loadDataFromParsedRows).
let DATA_HEADERS = null;
let DISPLAY_COLUMNS = null;
let allRows = [];
let TABLE_SOURCE_ROWS = [];

function buildDisplayColumns(headers) {
  const hs = Array.isArray(headers) ? headers.filter(Boolean) : [];
  const headerSet = new Set(hs);

  // Hide internal/derived helper fields from UI column lists.
  const hidden = new Set(["Century", "DatingYear", "Main text group", "Depository_abbr"]);

  const out = [];
  const base = Array.isArray(COLUMN_ORDER) && COLUMN_ORDER.length ? COLUMN_ORDER : hs;

  for (const col of base) {
    if (hidden.has(col)) continue;
    if (headerSet.has(col)) out.push(col);
  }

  // Append any remaining headers not covered by COLUMN_ORDER.
  for (const col of hs) {
    if (hidden.has(col)) continue;
    if (!out.includes(col)) out.push(col);
  }

  return out;
}

// Column title overrides (display-only)
const COLUMN_TITLE_OVERRIDES = {
  "Shelf mark": "Shelfmark",
};

// Shared column presets for both views
const COLUMN_PRESETS = {
  minimal: [
    "Depository",
    "Shelf mark",
    "Name",
    "Language",
    "Production Unit",
    "Dating",
    "Material",
    "Links to Database",
  ],
  textual: [
    "Depository",
    "Shelf mark",
    "Language",
    "Production Unit",
    "Leaves/Pages",
    "Main text",
    "Minor text",
    "Dating",
  ],
  codicology: [
    "Depository",
    "Shelf mark",
    "Language",
    "Object",
    "Material",
    "Size",
    "Leaves/Pages",
    "Gatherings",
    "Full size",
    "Leaf size",
    "Catch Words and Gatherings",
    "Pricking",
    "Ruling",
    "Production Unit",
    "Production",
  ],
  layout: [
    "Depository",
    "Shelf mark",
    "Language",
    "Columns",
    "Lines",
    "Ruling",
    "Rubric",
    "Style",
    "Colours",
    "Form of Initials",
    "Size of Initials",
    "Iconography",
    "Place",
  ],
};

function getPresetColumns(presetName) {
  const key = String(presetName || '').toLowerCase();
  if (key === 'textual' || key === 'textualcontent' || key === 'textual_content') return COLUMN_PRESETS.textual.slice();
  if (key === 'minimal') return COLUMN_PRESETS.minimal.slice();
  if (key === 'codicology') return COLUMN_PRESETS.codicology.slice();
  if (key === 'layout' || key === 'layout&decoration') return COLUMN_PRESETS.layout.slice();
  return [];
}

function applyDefaultMergedColumnVisibility() {
  MERGED_VISIBLE_COLUMNS = new Set(getPresetColumns('textual').filter(c => DISPLAY_COLUMNS && DISPLAY_COLUMNS.includes(c)));
  sanitizeMergedColumnVisibility();
}

function applyDefaultTableColumnVisibility() {
  TABLE_VISIBLE_COLUMNS = new Set(getPresetColumns('textual').filter(c => DISPLAY_COLUMNS && DISPLAY_COLUMNS.includes(c)));
  sanitizeTableColumnVisibility();
}

function getColumnTitle(col) {
  return COLUMN_TITLE_OVERRIDES[col] || col;
}

function saveMergedColumnVisibility() {
  try {
    if (!(MERGED_VISIBLE_COLUMNS instanceof Set)) {
      localStorage.removeItem(MERGED_COLUMNS_STORAGE_KEY);
      return;
    }
    localStorage.setItem(MERGED_COLUMNS_STORAGE_KEY, JSON.stringify(Array.from(MERGED_VISIBLE_COLUMNS)));
  } catch (e) {
    // ignore
  }
}

function sanitizeMergedColumnVisibility() {
  if (!DISPLAY_COLUMNS || !Array.isArray(DISPLAY_COLUMNS)) return;
  if (!(MERGED_VISIBLE_COLUMNS instanceof Set)) return;

  const allowed = new Set(DISPLAY_COLUMNS);
  for (const c of Array.from(MERGED_VISIBLE_COLUMNS)) {
    if (!allowed.has(c)) MERGED_VISIBLE_COLUMNS.delete(c);
  }
  // Avoid rendering a table with zero columns.
  if (MERGED_VISIBLE_COLUMNS.size === 0) {
    MERGED_VISIBLE_COLUMNS = null;
  }
}

function getMergedVisibleColumnsSet() {
  return (MERGED_VISIBLE_COLUMNS instanceof Set) ? MERGED_VISIBLE_COLUMNS : null;
}

function getMergedVisibleColumnsArray(columnsFull) {
  const cols = Array.isArray(columnsFull) ? columnsFull : [];
  return cols.filter(isMergedColumnVisible);
}

// Persist Text View column visibility across reloads.
const TABLE_COLUMNS_STORAGE_KEY = "nordiclaw.tableColumns";
let TABLE_VISIBLE_COLUMNS = null; // null => all columns visible

function isTableColumnVisible(col) {
  if (!(TABLE_VISIBLE_COLUMNS instanceof Set)) return true;
  return TABLE_VISIBLE_COLUMNS.has(col);
}

function loadTableColumnVisibility() {
  try {
    const raw = localStorage.getItem(TABLE_COLUMNS_STORAGE_KEY);
    if (!raw) {
      applyDefaultTableColumnVisibility();
      return;
    }
    const arr = JSON.parse(raw);
    if (Array.isArray(arr) && arr.length > 0) {
      TABLE_VISIBLE_COLUMNS = new Set(arr.map(String));
    } else {
      applyDefaultTableColumnVisibility();
    }
  } catch (e) {
    applyDefaultTableColumnVisibility();
  }
}

function saveTableColumnVisibility() {
  try {
    if (!(TABLE_VISIBLE_COLUMNS instanceof Set)) {
      localStorage.removeItem(TABLE_COLUMNS_STORAGE_KEY);
      return;
    }
    localStorage.setItem(TABLE_COLUMNS_STORAGE_KEY, JSON.stringify(Array.from(TABLE_VISIBLE_COLUMNS)));
  } catch (e) {
    // ignore
  }
}

function sanitizeTableColumnVisibility() {
  if (!DISPLAY_COLUMNS || !Array.isArray(DISPLAY_COLUMNS)) return;
  if (!(TABLE_VISIBLE_COLUMNS instanceof Set)) return;
  const allowed = new Set(DISPLAY_COLUMNS);
  for (const c of Array.from(TABLE_VISIBLE_COLUMNS)) {
    if (!allowed.has(c)) TABLE_VISIBLE_COLUMNS.delete(c);
  }
  if (TABLE_VISIBLE_COLUMNS.size === 0) {
    TABLE_VISIBLE_COLUMNS = null;
  }
}

function getTableVisibleColumnsSet() {
  return (TABLE_VISIBLE_COLUMNS instanceof Set) ? TABLE_VISIBLE_COLUMNS : null;
}

function getTableVisibleColumnsArray(columnsFull) {
  const cols = Array.isArray(columnsFull) ? columnsFull : [];
  return cols.filter(isTableColumnVisible);
}

let TEXT_COLUMN_DEFS = null; // field -> colDef
let TEXT_ALWAYS_COLUMNS = []; // extra hidden columns that should stay attached

function updateTextViewColumnsFromVisibility() {
  if (!TEXT_COLUMN_DEFS || typeof TEXT_COLUMN_DEFS !== 'object') return;
  const full = (DISPLAY_COLUMNS && Array.isArray(DISPLAY_COLUMNS)) ? DISPLAY_COLUMNS : [];
  const visible = getTableVisibleColumnsArray(full);
  const cols = [];
  for (const h of visible) {
    if (TEXT_COLUMN_DEFS[h]) cols.push(TEXT_COLUMN_DEFS[h]);
  }
  for (const extra of (Array.isArray(TEXT_ALWAYS_COLUMNS) ? TEXT_ALWAYS_COLUMNS : [])) {
    if (extra && typeof extra === 'object') cols.push(extra);
  }
  if (table && typeof table.setColumns === 'function') {
    table.setColumns(cols);
    if (currentView === 'table' && typeof table.redraw === 'function') {
      try { table.redraw(true); } catch (e) {}
    }
  }
}

function renderTableColumnsMenu() {
  const menu = document.getElementById("table-columns-menu");
  if (!menu) return;
  if (!DISPLAY_COLUMNS || !Array.isArray(DISPLAY_COLUMNS) || DISPLAY_COLUMNS.length === 0) {
    menu.innerHTML = '<div class="text-secondary small">Columns not ready yet.</div>';
    return;
  }

  const minimal = getPresetColumns('minimal');
  const textualContent = getPresetColumns('textual');
  const codicology = getPresetColumns('codicology');
  const layoutAndDecoration = getPresetColumns('layout');

  const selected = getTableVisibleColumnsSet();
  const isChecked = (c) => !selected || selected.has(c);

  let html = '';
  html += '<div class="d-flex gap-2 mb-2">';
  html += '<button class="btn btn-sm btn-outline-secondary" type="button" data-cols-action="all">All</button>';
  html += '<button class="btn btn-sm btn-outline-secondary" type="button" data-cols-action="minimal">Minimal</button>';
  html += '<button class="btn btn-sm btn-outline-secondary" type="button" data-cols-action="textual">Textual content</button>';
  html += '<button class="btn btn-sm btn-outline-secondary" type="button" data-cols-action="codicology">Codicology</button>';
  html += '<button class="btn btn-sm btn-outline-secondary" type="button" data-cols-action="layout">Layout & decoration</button>';
  html += '<button class="btn btn-sm btn-outline-secondary" type="button" data-cols-action="reset">Reset</button>';
  html += '</div>';
  html += '<div class="mb-2">';
  html += '<input class="form-control form-control-sm" type="text" placeholder="Filter columns…" data-cols-filter>';
  html += '</div>';
  html += '<div data-cols-list>';
  for (const col of DISPLAY_COLUMNS) {
    const idToken = toDomIdToken(col);
    html += `
      <div class="form-check">
        <input class="form-check-input" type="checkbox" data-cols-col="${escapeHtml(col)}" id="tcolvis-${idToken}" ${isChecked(col) ? 'checked' : ''}>
        <label class="form-check-label" for="tcolvis-${idToken}">${escapeHtml(getColumnTitle(col))}</label>
      </div>`;
  }
  html += '</div>';
  menu.innerHTML = html;

  function applyFilter() {
    const q = (menu.querySelector('[data-cols-filter]')?.value || '').toLowerCase();
    const items = menu.querySelectorAll('[data-cols-list] .form-check');
    items.forEach(el => {
      const label = el.textContent || '';
      el.style.display = (!q || label.toLowerCase().includes(q)) ? '' : 'none';
    });
  }

  menu.querySelector('[data-cols-filter]')?.addEventListener('input', applyFilter);

  menu.querySelectorAll('button[data-cols-action]').forEach(btn => {
    btn.addEventListener('click', () => {
      const action = btn.getAttribute('data-cols-action');
      if (action === 'all') {
        TABLE_VISIBLE_COLUMNS = null;
      } else if (action === 'reset') {
        applyDefaultTableColumnVisibility();
      } else if (action === 'minimal') {
        TABLE_VISIBLE_COLUMNS = new Set(minimal.filter(c => DISPLAY_COLUMNS.includes(c)));
      } else if (action === 'textual') {
        TABLE_VISIBLE_COLUMNS = new Set(textualContent.filter(c => DISPLAY_COLUMNS.includes(c)));
      } else if (action === 'codicology') {
        TABLE_VISIBLE_COLUMNS = new Set(codicology.filter(c => DISPLAY_COLUMNS.includes(c)));
      } else if (action === 'layout') {
        TABLE_VISIBLE_COLUMNS = new Set(layoutAndDecoration.filter(c => DISPLAY_COLUMNS.includes(c)));
      }
      saveTableColumnVisibility();
      renderTableColumnsMenu();
      if (currentView === 'table') updateTextViewColumnsFromVisibility();
    });
  });

  menu.querySelectorAll('input[type=checkbox][data-cols-col]').forEach(cb => {
    cb.addEventListener('change', () => {
      const col = cb.getAttribute('data-cols-col');
      if (!col) return;
      if (!(TABLE_VISIBLE_COLUMNS instanceof Set)) {
        TABLE_VISIBLE_COLUMNS = new Set(DISPLAY_COLUMNS);
      }
      if (cb.checked) TABLE_VISIBLE_COLUMNS.add(col);
      else TABLE_VISIBLE_COLUMNS.delete(col);
      sanitizeTableColumnVisibility();
      saveTableColumnVisibility();
      if (currentView === 'table') updateTextViewColumnsFromVisibility();
    });
  });
}

function renderMergedColumnsMenu() {
  const menu = document.getElementById("merged-columns-menu");
  if (!menu) return;
  if (!DISPLAY_COLUMNS || !Array.isArray(DISPLAY_COLUMNS) || DISPLAY_COLUMNS.length === 0) {
    menu.innerHTML = '<div class="text-secondary small">Columns not ready yet.</div>';
    return;
  }

  const minimal = getPresetColumns('minimal');
  const textualContent = getPresetColumns('textual');
  const codicology = getPresetColumns('codicology');
  const layoutAndDecoration = getPresetColumns('layout');

  const selected = getMergedVisibleColumnsSet();
  const isChecked = (c) => !selected || selected.has(c);

  let html = '';
  html += '<div class="d-flex gap-2 mb-2">';
  html += '<button class="btn btn-sm btn-outline-secondary" type="button" data-cols-action="all">All</button>';
  html += '<button class="btn btn-sm btn-outline-secondary" type="button" data-cols-action="minimal">Minimal</button>';
  html += '<button class="btn btn-sm btn-outline-secondary" type="button" data-cols-action="textual">Textual content</button>';
  html += '<button class="btn btn-sm btn-outline-secondary" type="button" data-cols-action="codicology">Codicology</button>';
  html += '<button class="btn btn-sm btn-outline-secondary" type="button" data-cols-action="layout">Layout & decoration</button>';
  html += '<button class="btn btn-sm btn-outline-secondary" type="button" data-cols-action="reset">Reset</button>';
  html += '</div>';
  html += '<div class="mb-2">';
  html += '<input class="form-control form-control-sm" type="text" placeholder="Filter columns…" data-cols-filter>';
  html += '</div>';
  html += '<div data-cols-list>';
  for (const col of DISPLAY_COLUMNS) {
    const idToken = toDomIdToken(col);
    html += `
      <div class="form-check">
        <input class="form-check-input" type="checkbox" data-cols-col="${escapeHtml(col)}" id="colvis-${idToken}" ${isChecked(col) ? 'checked' : ''}>
        <label class="form-check-label" for="colvis-${idToken}">${escapeHtml(getColumnTitle(col))}</label>
      </div>`;
  }
  html += '</div>';
  menu.innerHTML = html;

  function applyFilter() {
    const q = (menu.querySelector('[data-cols-filter]')?.value || '').toLowerCase();
    const items = menu.querySelectorAll('[data-cols-list] .form-check');
    items.forEach(el => {
      const label = el.textContent || '';
      el.style.display = (!q || label.toLowerCase().includes(q)) ? '' : 'none';
    });
  }

  menu.querySelector('[data-cols-filter]')?.addEventListener('input', applyFilter);

  menu.querySelectorAll('button[data-cols-action]').forEach(btn => {
    btn.addEventListener('click', () => {
      const action = btn.getAttribute('data-cols-action');
      if (action === 'all') {
        MERGED_VISIBLE_COLUMNS = null;
      } else if (action === 'reset') {
        applyDefaultMergedColumnVisibility();
      } else if (action === 'minimal') {
        MERGED_VISIBLE_COLUMNS = new Set(minimal.filter(c => DISPLAY_COLUMNS.includes(c)));
      } else if (action === 'textual') {
        MERGED_VISIBLE_COLUMNS = new Set(textualContent.filter(c => DISPLAY_COLUMNS.includes(c)));
      } else if (action === 'codicology') {
        MERGED_VISIBLE_COLUMNS = new Set(codicology.filter(c => DISPLAY_COLUMNS.includes(c)));
      } else if (action === 'layout') {
        MERGED_VISIBLE_COLUMNS = new Set(layoutAndDecoration.filter(c => DISPLAY_COLUMNS.includes(c)));
      }
      saveMergedColumnVisibility();
      renderMergedColumnsMenu();
      if (currentView === 'merged') applyFacetFilters();
    });
  });

  menu.querySelectorAll('input[type=checkbox][data-cols-col]').forEach(cb => {
    cb.addEventListener('change', () => {
      const col = cb.getAttribute('data-cols-col');
      if (!col) return;

      if (!(MERGED_VISIBLE_COLUMNS instanceof Set)) {
        MERGED_VISIBLE_COLUMNS = new Set(DISPLAY_COLUMNS);
      }

      if (cb.checked) MERGED_VISIBLE_COLUMNS.add(col);
      else MERGED_VISIBLE_COLUMNS.delete(col);

      sanitizeMergedColumnVisibility();
      saveMergedColumnVisibility();
      if (currentView === 'merged') applyFacetFilters();
    });
  });
}

// View mode: Text View (SimpleTable) vs Manuscript View (merged cells)
let currentView = "merged";

// Text View table adapter instance (SimpleTable).
let table = null;

class SimpleTable {
  constructor(containerSelector, options = {}) {
    this._root = (typeof containerSelector === 'string')
      ? document.querySelector(containerSelector)
      : containerSelector;
    this._columns = Array.isArray(options.columns) ? options.columns.slice() : [];
    this._baseData = Array.isArray(options.data) ? options.data.slice() : [];
    this._filterFn = null;
    this._filtered = this._baseData.slice();
    if (options && options.pageSize === true) {
      this._pageSize = Number.POSITIVE_INFINITY;
    } else if (options && options.pageSize !== undefined && options.pageSize !== null) {
      const n = Number(options.pageSize);
      this._pageSize = (Number.isFinite(n) && n > 0) ? n : 20;
    } else {
      this._pageSize = 20;
    }
    this._page = 1;
    this._handlers = new Map();

    this._tableEl = document.createElement('table');
    // Match Manuscript View styling as closely as possible.
    this._tableEl.className = 'table table-bordered table-striped merged-table';
    this._thead = document.createElement('thead');
    this._tbody = document.createElement('tbody');
    this._tableEl.appendChild(this._thead);
    this._tableEl.appendChild(this._tbody);

    this._scrollEl = document.createElement('div');
    this._scrollEl.className = 'simple-table-scroll';
    this._scrollEl.appendChild(this._tableEl);

    this._pagerEl = document.createElement('div');
    this._pagerEl.className = 'd-flex align-items-center gap-2';

    this._btnPrev = document.createElement('button');
    this._btnPrev.type = 'button';
    this._btnPrev.className = 'btn btn-outline-secondary';
    this._btnPrev.textContent = 'Prev';
    this._btnPrev.addEventListener('click', () => {
      if (this._page > 1) {
        this._page -= 1;
        this._render();
      }
    });

    this._btnNext = document.createElement('button');
    this._btnNext.type = 'button';
    this._btnNext.className = 'btn btn-outline-secondary';
    this._btnNext.textContent = 'Next';
    this._btnNext.addEventListener('click', () => {
      const totalPages = this._getTotalPages();
      if (this._page < totalPages) {
        this._page += 1;
        this._render();
      }
    });

    this._pageLabel = document.createElement('div');
    this._pageLabel.className = 'text-secondary small';

    this._btnGroupEl = document.createElement('div');
    this._btnGroupEl.className = 'btn-group btn-group-sm';
    this._btnGroupEl.appendChild(this._btnPrev);
    this._btnGroupEl.appendChild(this._btnNext);

    this._pagerEl.appendChild(this._btnGroupEl);
    this._pagerEl.appendChild(this._pageLabel);

    const pagerHost = options && options.pagerContainer
      ? options.pagerContainer
      : document.getElementById('table-pager-host');

    if (this._root) {
      try { this._root.classList.add('simple-table-root'); } catch (e) {}
      this._root.innerHTML = '';
      this._root.appendChild(this._scrollEl);

      if (pagerHost) {
        try { pagerHost.innerHTML = ''; } catch (e) {}
        pagerHost.appendChild(this._pagerEl);
      } else {
        this._pagerEl.classList.add('my-2');
        this._root.appendChild(this._pagerEl);
      }
    }

    this._applyFilter();
    this._render();
  }

  _getRenderColumns() {
    const cols = Array.isArray(this._columns) ? this._columns : [];
    return cols.filter(c => !(c && Object.prototype.hasOwnProperty.call(c, 'visible') && c.visible === false));
  }

  on(eventName, handler) {
    if (!eventName || typeof handler !== 'function') return;
    if (!this._handlers.has(eventName)) this._handlers.set(eventName, []);
    this._handlers.get(eventName).push(handler);
  }

  _emit(eventName, ...args) {
    const hs = this._handlers.get(eventName) || [];
    for (const h of hs) {
      try { h(...args); } catch (e) { /* ignore */ }
    }
  }

  getData() {
    return this._filtered.slice();
  }

  getDataCount() {
    return this._filtered.length;
  }

  setColumns(columns) {
    this._columns = Array.isArray(columns) ? columns.slice() : [];
    this._render();
  }

  replaceData(rows) {
    this._baseData = Array.isArray(rows) ? rows.slice() : [];
    this._page = 1;
    this._applyFilter();
    this._render();
    return Promise.resolve();
  }

  clearFilter(_force) {
    this._filterFn = null;
    this._page = 1;
    this._applyFilter();
    this._render();
  }

  setFilter(fn) {
    this._filterFn = (typeof fn === 'function') ? fn : null;
    this._page = 1;
    this._applyFilter();
    this._render();
  }

  setPageSize(size) {
    if (size === true) {
      this._pageSize = Number.POSITIVE_INFINITY;
    } else {
      const n = Number(size);
      this._pageSize = (Number.isFinite(n) && n > 0) ? n : 50;
    }
    this._page = 1;
    this._render();
  }

  redraw(_force) {
    this._render();
  }

  _applyFilter() {
    if (!this._filterFn) {
      this._filtered = this._baseData.slice();
    } else {
      this._filtered = this._baseData.filter(r => {
        try { return !!this._filterFn(r); } catch (e) { return false; }
      });
    }
    this._emit('dataFiltered', [], this._filtered.slice());
  }

  _getTotalPages() {
    if (!Number.isFinite(this._pageSize)) return 1;
    return Math.max(1, Math.ceil(this._filtered.length / this._pageSize));
  }

  _render() {
    if (!this._root) return;

    const renderCols = this._getRenderColumns();

    // Header
    this._thead.innerHTML = '';
    const trh = document.createElement('tr');
    for (const c of renderCols) {
      const th = document.createElement('th');
      th.textContent = c && c.title ? String(c.title) : '';
      trh.appendChild(th);
    }
    this._thead.appendChild(trh);

    // Body
    this._tbody.innerHTML = '';
    const start = (!Number.isFinite(this._pageSize)) ? 0 : (this._page - 1) * this._pageSize;
    const end = (!Number.isFinite(this._pageSize)) ? this._filtered.length : Math.min(this._filtered.length, start + this._pageSize);
    const pageRows = this._filtered.slice(start, end);

    for (const rowData of pageRows) {
      const tr = document.createElement('tr');
      tr.style.cursor = 'pointer';
      tr.addEventListener('click', (ev) => {
        const rowObj = { getData: () => rowData };
        this._emit('rowClick', ev, rowObj);
      });

      for (const colDef of renderCols) {
        const field = colDef && colDef.field ? colDef.field : '';
        const raw = rowData ? rowData[field] : '';
        const td = document.createElement('td');

        if (colDef && typeof colDef.formatter === 'function') {
          const rowObj = { getData: () => rowData };
          const cellObj = {
            getValue: () => raw,
            getRow: () => rowObj,
          };
          const html = colDef.formatter(cellObj);
          td.innerHTML = (html === null || html === undefined) ? '' : String(html);
        } else {
          td.textContent = (raw === null || raw === undefined) ? '' : String(raw);
        }

        tr.appendChild(td);
      }
      this._tbody.appendChild(tr);
    }

    // Pager
    const totalPages = this._getTotalPages();
    if (this._page > totalPages) this._page = totalPages;
    this._btnPrev.disabled = (this._page <= 1);
    this._btnNext.disabled = (this._page >= totalPages);
    const shown = pageRows.length;
    this._pageLabel.textContent = (!Number.isFinite(this._pageSize))
      ? `Showing ${shown} of ${this._filtered.length}`
      : `Page ${this._page} / ${totalPages} — showing ${shown} of ${this._filtered.length}`;
  }
}

// Persist the selected view across reloads (helps with Live Server reload behavior)
const VIEW_STORAGE_KEY = "nordiclaw.view";

// Persist Facet sidebar visibility across reloads.
const FACETS_HIDDEN_STORAGE_KEY = "nordiclaw.facetsHidden";

function setFacetsHidden(hidden) {
  const isHidden = !!hidden;
  try {
    document.body.classList.toggle('facets-hidden', isHidden);
  } catch (e) {
    // ignore
  }

  const btn = document.getElementById('toggle-facets');
  if (btn) btn.textContent = isHidden ? 'Show facets' : 'Hide facets';

  try {
    localStorage.setItem(FACETS_HIDDEN_STORAGE_KEY, isHidden ? '1' : '0');
  } catch (e) {
    // ignore
  }
}

// Persist Manuscript/Text sort mode across reloads.
const MERGED_SORT_STORAGE_KEY = "nordiclaw.mergedSort";
let MERGED_SORT_MODE = "shelfmark";

// Manuscript View pagination (paginate by manuscript to preserve merged rowspans)
let MERGED_PAGE = 1;

function getMergedPageSize() {
  const sel = document.getElementById('pagination-size');
  const raw = sel ? String(sel.value || '').toLowerCase() : '20';
  if (raw === 'all') return Number.POSITIVE_INFINITY;
  const n = parseInt(raw, 10);
  return (Number.isFinite(n) && n > 0) ? n : 20;
}

function updateMergedPagerUI(page, totalPages) {
  const pager = document.getElementById('merged-pager');
  const btnPrev = document.getElementById('merged-prev-page');
  const btnNext = document.getElementById('merged-next-page');
  const indicator = document.getElementById('merged-page-indicator');

  const show = (currentView === 'merged') && (typeof totalPages === 'number') && totalPages > 1;
  if (pager) pager.style.display = show ? '' : 'none';
  if (btnPrev) btnPrev.disabled = (!show) || page <= 1;
  if (btnNext) btnNext.disabled = (!show) || page >= totalPages;
  if (indicator) indicator.textContent = show ? `Page ${page} / ${totalPages}` : '';
}

function getMergedSortMode() {
  const m = normalizeForCompare(MERGED_SORT_MODE);
  return (m === "dating") ? "dating" : "shelfmark";
}

function compareText(a, b) {
  const ax = normalizeForCompare(a);
  const bx = normalizeForCompare(b);
  return ax.localeCompare(bx, undefined, { numeric: true, sensitivity: 'base' });
}

function compareManuscripts(msA, msB, sortMode) {
  const aRows = (msA && Array.isArray(msA.rows)) ? msA.rows : [];
  const bRows = (msB && Array.isArray(msB.rows)) ? msB.rows : [];
  const a0 = aRows[0] || {};
  const b0 = bRows[0] || {};

  const aDep = normalizeForCompare(a0["Depository_abbr"] || a0["Depository"] || "");
  const bDep = normalizeForCompare(b0["Depository_abbr"] || b0["Depository"] || "");
  const aShelf = normalizeForCompare(a0["Shelf mark"] || "");
  const bShelf = normalizeForCompare(b0["Shelf mark"] || "");

  if (sortMode === "dating") {
    const yearsA = aRows.map(r => r["DatingYear"]).filter(y => typeof y === 'number' && !Number.isNaN(y));
    const yearsB = bRows.map(r => r["DatingYear"]).filter(y => typeof y === 'number' && !Number.isNaN(y));
    const minA = yearsA.length ? Math.min(...yearsA) : Number.POSITIVE_INFINITY;
    const minB = yearsB.length ? Math.min(...yearsB) : Number.POSITIVE_INFINITY;
    if (minA !== minB) return minA - minB;
  }

  // Default: stable manuscript ordering by Depository then Shelf mark.
  const depCmp = compareText(aDep, bDep);
  if (depCmp) return depCmp;
  const shelfCmp = compareText(aShelf, bShelf);
  if (shelfCmp) return shelfCmp;
  return 0;
}

function sortRowsForTextView(rows, sortMode) {
  const arr = Array.isArray(rows) ? rows.slice() : [];
  const decorated = arr.map((r, idx) => ({ r, idx }));
  decorated.sort((a, b) => {
    const ra = a.r || {};
    const rb = b.r || {};

    if (sortMode === "dating") {
      const ay = (typeof ra["DatingYear"] === 'number' && !Number.isNaN(ra["DatingYear"])) ? ra["DatingYear"] : Number.POSITIVE_INFINITY;
      const by = (typeof rb["DatingYear"] === 'number' && !Number.isNaN(rb["DatingYear"])) ? rb["DatingYear"] : Number.POSITIVE_INFINITY;
      if (ay !== by) return ay - by;
    }

    const depCmp = compareText(ra["Depository_abbr"] || ra["Depository"] || "", rb["Depository_abbr"] || rb["Depository"] || "");
    if (depCmp) return depCmp;
    const shelfCmp = compareText(ra["Shelf mark"] || "", rb["Shelf mark"] || "");
    if (shelfCmp) return shelfCmp;
    return a.idx - b.idx;
  });
  return decorated.map(x => x.r);
}

function applyTableSort(mode) {
  MERGED_SORT_MODE = (mode === "dating") ? "dating" : "shelfmark";
  try { localStorage.setItem(MERGED_SORT_STORAGE_KEY, MERGED_SORT_MODE); } catch (e) {}

  // Update Text View ordering (we sort data arrays).
  if (typeof table !== 'undefined' && table && typeof table.replaceData === 'function') {
    const base = (typeof TABLE_SOURCE_ROWS !== 'undefined' && Array.isArray(TABLE_SOURCE_ROWS)) ? TABLE_SOURCE_ROWS : (Array.isArray(allRows) ? allRows : []);
    const sorted = sortRowsForTextView(base, getMergedSortMode());
    try {
      const p = table.replaceData(sorted);
      if (p && typeof p.then === 'function') p.then(() => { try { applyFacetFilters(); } catch (e) {} });
    } catch (e) {
      // ignore
    }
  }
}

// Persist merged-column visibility across reloads.
const MERGED_COLUMNS_STORAGE_KEY = "nordiclaw.mergedColumns";
let MERGED_VISIBLE_COLUMNS = null; // null => all columns visible

function isMergedColumnVisible(col) {
  if (!(MERGED_VISIBLE_COLUMNS instanceof Set)) return true;
  return MERGED_VISIBLE_COLUMNS.has(col);
}

function loadMergedColumnVisibility() {
  try {
    const raw = localStorage.getItem(MERGED_COLUMNS_STORAGE_KEY);
    if (!raw) {
      applyDefaultMergedColumnVisibility();
      return;
    }
    const arr = JSON.parse(raw);
    if (Array.isArray(arr) && arr.length > 0) {
      MERGED_VISIBLE_COLUMNS = new Set(arr.map(String));
    } else {
      applyDefaultMergedColumnVisibility();
    }
  } catch (e) {
    applyDefaultMergedColumnVisibility();
  }
}

// Optional: raw Excel-like dataset (blank cells preserved) + merge coordinates exported by convert_excel_to_tsv.py
// These files are expected to be generated via:
//   python data/scripts/convert_excel_to_tsv.py data/<file>.xlsx --excel-identical --export-merges
// producing:
//   data/<file>_raw.tsv and data/<file>_raw_merges.json
let RAW_EXCEL_LOADING = null;
let RAW_EXCEL_LOADED = false;
let RAW_EXCEL_FAILED = false;
let RAW_BY_MANUSCRIPT_KEY = new Map(); // key -> { rows: RawRow[], sourceId: string }
let RAW_MERGES_BY_SOURCE = new Map();  // sourceId -> { columns: string[], merges: MergeRange[] }
let RAW_ROWS_BY_SOURCE = new Map();    // sourceId -> RawRow[] (in source-row order)
let RAW_EXCEL_HEADERS = null;          // string[] including inserted Language

const RAW_EXCEL_SOURCES = [
  { id: "dan", lang: "da", tsv: "data/1.0_Metadata_Dan_raw.tsv", merges: "data/1.0_Metadata_Dan_raw_merges.json" },
  { id: "isl", lang: "is", tsv: "data/1.1_Metadata_Isl_raw.tsv", merges: "data/1.1_Metadata_Isl_raw_merges.json" },
  { id: "norw", lang: "no", tsv: "data/1.2_Metadata_Norw_raw.tsv", merges: "data/1.2_Metadata_Norw_raw_merges.json" },
  { id: "swe", lang: "sv", tsv: "data/1.1_Metadata_Swe_raw.tsv", merges: "data/1.1_Metadata_Swe_raw_merges.json" },
];

async function ensureRawExcelLoaded() {
  if (RAW_EXCEL_LOADED || RAW_EXCEL_FAILED) return Promise.resolve();
  if (RAW_EXCEL_LOADING) return RAW_EXCEL_LOADING;

  RAW_EXCEL_LOADING = (async () => {
    try {
      const depositoryMap = await loadDepositoryMap();
      const results = await Promise.all(RAW_EXCEL_SOURCES.map(async (src) => {
        try {
          const tsvResp = await fetch(src.tsv);
          if (!tsvResp.ok) throw new Error(`Failed to fetch ${src.tsv}: ${tsvResp.status}`);
          const tsvText = await tsvResp.text();

          // Merges are optional; if missing, the merged view still benefits from the raw blank cells
          try {
            const mergesResp = await fetch(src.merges);
            if (mergesResp.ok) {
              const mergesJson = await mergesResp.json();
              RAW_MERGES_BY_SOURCE.set(src.id, mergesJson);
            } else {
              console.warn(`Missing merge JSON for ${src.id}: ${src.merges} (${mergesResp.status})`);
            }
          } catch (e) {
            console.warn(`Failed to load merge JSON for ${src.id}: ${src.merges}`, e);
          }

          // Parse TSV (keep empty lines to preserve Excel spacing; they will still be associated to a manuscript via hidden key fill)
          const lines = tsvText.split(/\r?\n/);
          if (!lines.length) return false;
          const rawHeaders = (lines[0] || "").split("\t");

          // Insert Language after Shelf mark (to match the combined dataset layout)
          const shelfIdx = rawHeaders.indexOf("Shelf mark");
          const headers = rawHeaders.slice();
          if (shelfIdx !== -1 && !headers.includes("Language")) {
            headers.splice(shelfIdx + 1, 0, "Language");
          }

          if (!RAW_EXCEL_HEADERS) RAW_EXCEL_HEADERS = headers.slice();
          if (!RAW_ROWS_BY_SOURCE.has(src.id)) RAW_ROWS_BY_SOURCE.set(src.id, []);

          let currentShelf = "";
          let currentDepAbbr = "";

          for (let i = 1; i < lines.length; i++) {
            const line = lines[i];
            // Preserve row count even if line is empty
            const rawValues = (line !== undefined) ? line.split("\t") : [];
            const values = rawValues.slice();
            if (shelfIdx !== -1 && !rawHeaders.includes("Language")) {
              values.splice(shelfIdx + 1, 0, src.lang);
            }

            const obj = {};
            headers.forEach((h, colIndex) => {
              let val = values[colIndex] !== undefined ? values[colIndex] : "";
              if (h === "Main text") val = String(val || "").trim();
              if (h === "Depository") {
                obj["Depository_abbr"] = val;
                if (val in depositoryMap) {
                  val = depositoryMap[val];
                }
              }
              obj[h] = val;
            });

            // Hidden key fill (does NOT change visible blank cells)
            const shelfVal = normalizeForCompare(obj["Shelf mark"]);
            const depAbbrVal = normalizeForCompare(obj["Depository_abbr"]);
            if (shelfVal) currentShelf = shelfVal;
            if (depAbbrVal) currentDepAbbr = depAbbrVal;
            obj["Depository_abbr"] = obj["Depository_abbr"] || currentDepAbbr;
            obj["__msKey"] = (currentDepAbbr || "") + "||" + (currentShelf || "");
            obj["__sourceId"] = src.id;
            obj["__sourceRowIndex"] = i - 1; // 0-based index relative to data rows (matches merges JSON)

            // Derived fields (for consistency if used by renderers)
            if (obj["Main text"]) {
              const m = String(obj["Main text"]).match(/^([^\(]+)\s*\(/);
              obj["Main text group"] = m ? m[1].trim() : obj["Main text"];
            } else {
              obj["Main text group"] = "";
            }
            obj["DatingYear"] = parseDatingYear(obj["Dating"]);
            normalizeRowLanguage(obj);

            // Store into manuscript map (skip rows before we have a key)
            const msKey = obj["__msKey"];
            if (!msKey || msKey === "||") continue;
            if (!RAW_BY_MANUSCRIPT_KEY.has(msKey)) {
              RAW_BY_MANUSCRIPT_KEY.set(msKey, { rows: [], sourceId: src.id });
            }
            RAW_BY_MANUSCRIPT_KEY.get(msKey).rows.push(obj);
            RAW_ROWS_BY_SOURCE.get(src.id).push(obj);
          }

          return true;
        } catch (e) {
          console.warn(`Raw TSV not available for ${src.id}: ${src.tsv}`, e);
          return false;
        }
      }));

      const loadedCount = results.filter(Boolean).length;
      if (loadedCount > 0) {
        RAW_EXCEL_LOADED = true;
      } else {
        RAW_EXCEL_FAILED = true;
      }
    } catch (e) {
      console.warn("Raw Excel dataset not available; merged view will fall back to heuristic merging.", e);
      RAW_EXCEL_FAILED = true;
    } finally {
      RAW_EXCEL_LOADING = null;
    }
  })();

  return RAW_EXCEL_LOADING;
}

function applyMergeAwareFillDown(rows, sourceId) {
  const mergesJson = RAW_MERGES_BY_SOURCE.get(sourceId);
  if (!mergesJson || !Array.isArray(mergesJson.merges) || mergesJson.merges.length === 0) return;

  const mergeColumns = Array.isArray(mergesJson.columns) ? mergesJson.columns : null;

  const byIndex = new Map();
  for (const r of rows) {
    if (r && typeof r.__sourceRowIndex === 'number') byIndex.set(r.__sourceRowIndex, r);
  }

  for (const m of mergesJson.merges) {
    if (!m) continue;
    if (typeof m.minRow !== 'number' || typeof m.maxRow !== 'number') continue;

    let colNames = [];
    if (mergeColumns && typeof m.minColIndex === 'number' && typeof m.maxColIndex === 'number') {
      colNames = mergeColumns.slice(m.minColIndex, m.maxColIndex + 1).map(String);
    } else if (m.minCol) {
      colNames = [String(m.minCol)];
    }

    for (const col of colNames) {
      if (!col) continue;
      const topRow = byIndex.get(m.minRow);
      if (!topRow) continue;
      const topVal = topRow[col];
      if (!normalizeForCompare(topVal)) continue;

      for (let i = m.minRow; i <= m.maxRow; i++) {
        const row = byIndex.get(i);
        if (!row) continue;
        if (!normalizeForCompare(row[col])) row[col] = topVal;
      }
    }
  }
}

const _MD_LINK_RE = /^\[([^\]]+)\]\(([^)]+)\)$/;

function splitLinksToDatabase(value) {
  if (value === null || value === undefined) return [];
  let s = String(value).replace(/\r/g, "").trim();
  if (!s) return [];
  if (s === "Unavailable") return [];

  // Split on semicolons and newlines; keep items trimmed.
  const rawParts = [];
  for (const chunk of s.split(";")) {
    for (const line of String(chunk).split("\n")) {
      const t = line.trim();
      if (t) rawParts.push(t);
    }
  }

  // If markdown links exist, drop matching bare labels (Excel exports can include both).
  const labels = new Set();
  for (const p of rawParts) {
    const m = p.match(_MD_LINK_RE);
    if (m && m[1]) labels.add(String(m[1]).trim());
  }

  const seen = new Set();
  const out = [];
  for (const p of rawParts) {
    if (!p) continue;
    if (p === "Unavailable") continue;
    if (labels.size && labels.has(p)) continue;
    const key = normalizeForCompare(p);
    if (!key) continue;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(p);
  }

  return out;
}

function normalizeLinksToDatabase(value) {
  return splitLinksToDatabase(value).join("; ");
}

function renderLinksToDatabaseHtml(value) {
  const parts = splitLinksToDatabase(value);
  if (parts.length === 0) return "";

  const isHttpUrl = (u) => /^https?:\/\//i.test(String(u || "").trim());
  const toLink = (url, label) => {
    const href = String(url || "").trim();
    const text = (label === undefined || label === null) ? href : String(label);
    if (!isHttpUrl(href)) return escapeHtml(text);
    return `<a href="${escapeHtml(href)}" target="_blank" rel="noopener noreferrer">${escapeHtml(text)}</a>`;
  };

  return parts.map(p => {
    const m = String(p).match(_MD_LINK_RE);
    if (m) return toLink(m[2], m[1]);
    if (isHttpUrl(p)) return toLink(p, p);
    return escapeHtml(p);
  }).join("; ");
}

function mergeLinksPerManuscript(rows) {
  const byKey = new Map();
  const labelsWithUrls = new Map();

  for (const r of rows) {
    const dep = normalizeForCompare(r["Depository_abbr"] || "");
    const shelf = normalizeForCompare(r["Shelf mark"] || "");
    const key = `${dep}||${shelf}`;
    const v = normalizeForCompare(r["Links to Database"]);
    if (!v) continue;

    const parts = v.split(";").map(p => String(p).trim()).filter(Boolean);
    if (!byKey.has(key)) byKey.set(key, []);
    byKey.get(key).push(...parts);

    for (const p of parts) {
      const m = p.match(_MD_LINK_RE);
      if (m) {
        if (!labelsWithUrls.has(key)) labelsWithUrls.set(key, new Set());
        labelsWithUrls.get(key).add(m[1]);
      }
    }
  }

  const mergedByKey = new Map();
  for (const [key, parts] of byKey.entries()) {
    const dropLabels = labelsWithUrls.get(key) || new Set();
    const seen = new Set();
    const out = [];
    for (const p of parts) {
      if (!p) continue;
      if (dropLabels.has(p)) continue;
      if (seen.has(p)) continue;
      seen.add(p);
      out.push(p);
    }
    mergedByKey.set(key, out.join("; "));
  }

  for (const r of rows) {
    const dep = normalizeForCompare(r["Depository_abbr"] || "");
    const shelf = normalizeForCompare(r["Shelf mark"] || "");
    const key = `${dep}||${shelf}`;
    if (mergedByKey.has(key)) r["Links to Database"] = mergedByKey.get(key);
  }
}

function dropCompletelyEmptyRows(rows, headers) {
  const ignore = new Set(["Century", "DatingYear", "Main text group", "Depository_abbr", "Language"]);
  return rows.filter(r => {
    for (const h of headers) {
      if (ignore.has(h)) continue;
      if (normalizeForCompare(r[h])) return true;
    }
    return false;
  });
}

function deduplicateRows(rows, headers) {
  const ignore = new Set(["Century", "DatingYear", "Main text group", "Depository_abbr"]);
  const seen = new Set();
  const out = [];
  for (const r of rows) {
    const key = headers
      .filter(h => !ignore.has(h))
      .map(h => normalizeForCompare(r[h] ?? ""))
      .join("\u0001");
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(r);
  }
  return out;
}

async function loadDataFromRawExcelSources() {
  try {
    await loadAbbreviationsMap();
    await ensureRawExcelLoaded();
    if (!RAW_EXCEL_LOADED) throw new Error('Raw Excel sources not loaded');

    const headers = (Array.isArray(RAW_EXCEL_HEADERS) && RAW_EXCEL_HEADERS.length > 0)
      ? RAW_EXCEL_HEADERS.slice()
      : COLUMN_ORDER.slice();

    const combinedRows = [];
    for (const src of RAW_EXCEL_SOURCES) {
      const rawRows = RAW_ROWS_BY_SOURCE.get(src.id) || [];

      // Clone rows so we don't mutate the raw dataset used by Manuscript View.
      const rows = rawRows.map(r => {
        const o = {};
        for (const k of Object.keys(r)) {
          if (k.startsWith('__')) continue;
          o[k] = r[k];
        }
        // Keep source row index for merge-aware fill-down.
        o.__sourceRowIndex = r.__sourceRowIndex;
        return o;
      });

      applyMergeAwareFillDown(rows, src.id);

      for (const r of rows) {
        delete r.__sourceRowIndex;

        // Match filled TSV behavior: replace literal "Unavailable" with empty.
        for (const k of Object.keys(r)) {
          if (r[k] === "Unavailable") r[k] = "";
        }

        // Match filled TSV formatting for some structured fields.
        if ("Gatherings" in r) r["Gatherings"] = normalizeGatherings(r["Gatherings"]);
        if ("Scribe" in r) r["Scribe"] = normalizeScribe(r["Scribe"]);

        // Recompute derived fields after fill-down.
        if (r["Main text"]) {
          const m = String(r["Main text"]).match(/^([^\(]+)\s*\(/);
          r["Main text group"] = m ? m[1].trim() : r["Main text"];
        } else {
          r["Main text group"] = "";
        }
        r["Century"] = parseCentury(r["Dating"]);
        r["DatingYear"] = parseDatingYear(r["Dating"]);
        normalizeRowLanguage(r);

        // Ensure all expected headers exist as keys.
        for (const h of headers) {
          if (!(h in r)) r[h] = "";
        }

        combinedRows.push(r);
      }
    }

    // Mirror the filled TSV pipeline steps: drop empty, merge links per manuscript, deduplicate.
    let filteredRows = dropCompletelyEmptyRows(combinedRows, headers);
    mergeLinksPerManuscript(filteredRows);
    filteredRows = deduplicateRows(filteredRows, headers);

    // Apply UI normalization after dedup so row identity matches the pipeline.
    for (const r of filteredRows) {
      if ("Leaves/Pages" in r) r["Leaves/Pages"] = normalizeLeavesPages(r["Leaves/Pages"]);
    }

    return loadDataFromParsedRows(headers, filteredRows);
  } catch (e) {
    console.warn('Falling back to combined TSV (data/NordicLaw_data.tsv):', e);
    return loadDataTSV('data/NordicLaw_data.tsv');
  }
}

function buildLocalMergeLookup(rawRows, sourceId, columnsFull, columnsVisible) {
  const source = RAW_MERGES_BY_SOURCE.get(sourceId);
  if (!source || !Array.isArray(source.merges)) return null;

  // Map source row index -> local row index
  const rowIndexToLocal = new Map();
  for (let i = 0; i < rawRows.length; i++) {
    rowIndexToLocal.set(rawRows[i].__sourceRowIndex, i);
  }

  const full = (Array.isArray(columnsFull) && columnsFull.length > 0)
    ? columnsFull
    : ((DISPLAY_COLUMNS && Array.isArray(DISPLAY_COLUMNS)) ? DISPLAY_COLUMNS : ((DATA_HEADERS && Array.isArray(DATA_HEADERS)) ? DATA_HEADERS : COLUMN_ORDER));

  const visible = (Array.isArray(columnsVisible) && columnsVisible.length > 0)
    ? columnsVisible
    : full;

  const fullIndex = new Map(full.map((c, idx) => [c, idx]));
  const visibleIndex = new Map(visible.map((c, idx) => [c, idx]));

  const topLeft = new Map(); // "r,c" -> {rowSpan, colSpan}
  const covered = new Set(); // "r,c"

  for (const m of source.merges) {
    const minSrcRow = m.minRow;
    const maxSrcRow = m.maxRow;
    if (!rowIndexToLocal.has(minSrcRow) || !rowIndexToLocal.has(maxSrcRow)) continue;

    // Only apply merges fully contained within this manuscript's row set
    let fullyContained = true;
    for (let r = minSrcRow; r <= maxSrcRow; r++) {
      if (!rowIndexToLocal.has(r)) { fullyContained = false; break; }
    }
    if (!fullyContained) continue;

    const minColName = m.minCol;
    const maxColName = m.maxCol;
    if (!fullIndex.has(minColName) || !fullIndex.has(maxColName)) continue;
    const minFullCol = fullIndex.get(minColName);
    const maxFullCol = fullIndex.get(maxColName);
    const a = Math.min(minFullCol, maxFullCol);
    const b = Math.max(minFullCol, maxFullCol);

    // Map the merge range to the visible column grid.
    const visibleColsInRange = [];
    for (let i = a; i <= b; i++) {
      const colName = full[i];
      if (visibleIndex.has(colName)) visibleColsInRange.push(visibleIndex.get(colName));
    }
    if (visibleColsInRange.length === 0) continue;
    const minCol = Math.min(...visibleColsInRange);
    const maxCol = Math.max(...visibleColsInRange);

    const minRow = rowIndexToLocal.get(minSrcRow);
    const maxRow = rowIndexToLocal.get(maxSrcRow);

    const rowSpan = (maxRow - minRow) + 1;
    const colSpan = (maxCol - minCol) + 1;
    if (rowSpan <= 1 && colSpan <= 1) continue;

    const tlKey = `${minRow},${minCol}`;
    topLeft.set(tlKey, { rowSpan, colSpan });

    for (let r = minRow; r <= maxRow; r++) {
      for (let c = minCol; c <= maxCol; c++) {
        const k = `${r},${c}`;
        if (k === tlKey) continue;
        covered.add(k);
      }
    }
  }

  return { topLeft, covered, columns: visible };
}

function normalizeForCompare(val) {
  if (val === null || val === undefined) return "";
  const v = String(val).trim();
  return v === "." ? "" : v;
}

function aggregateFieldValues(rows, field, separator = "; ") {
  const parts = [];
  for (const r of (rows || [])) {
    if (!r) continue;
    const v = normalizeForCompare(r[field]);
    if (!v) continue;
    parts.push(v);
  }
  return parts.join(separator);
}

function allFieldValuesEmpty(rows, field) {
  for (const r of (rows || [])) {
    if (!r) continue;
    if (normalizeForCompare(r[field])) return false;
  }
  return true;
}

function getManuscriptKey(row) {
  const dep = normalizeForCompare(row["Depository_abbr"] || row["Depository"]);
  const shelf = normalizeForCompare(row["Shelf mark"]);
  return dep + "||" + shelf;
}

function groupByPreserveOrder(rows, keyFn) {
  const map = new Map();
  for (const r of rows) {
    const k = keyFn(r);
    if (!map.has(k)) map.set(k, []);
    map.get(k).push(r);
  }
  return Array.from(map.entries()).map(([key, groupRows]) => ({ key, rows: groupRows }));
}

function getConstantFields(rows, columns, excluded = []) {
  const excludedSet = new Set(excluded);
  const constant = new Set();
  for (const col of columns) {
    if (excludedSet.has(col)) continue;
    const values = new Set(rows.map(r => normalizeForCompare(r[col])));
    if (values.size <= 1) constant.add(col);
  }
  return constant;
}

function splitSemicolonList(value) {
  return String(value || "")
    .split(/\s*;\s*/)
    .map(s => s.trim())
    .filter(Boolean);
}

// General abbreviation mapping (loaded from data/abbreviations.tsv)
let ABBREVIATIONS_MAP = null;

const EN_DASH = " \u2013 ";
function parseTrailingParenSuffix(raw) {
  const s = String(raw || "").trim();
  const m = s.match(/^(.+?)\s*\(([^)]+)\)\s*$/);
  if (!m) return { base: s, suffix: "" };
  return { base: m[1].trim(), suffix: ` (${m[2].trim()})` };
}

function getAbbreviationsMapSync() {
  return (typeof window !== 'undefined' && window.ABBREVIATIONS_MAP)
    ? window.ABBREVIATIONS_MAP
    : (typeof ABBREVIATIONS_MAP !== 'undefined' ? ABBREVIATIONS_MAP : null);
}

function expandParenSuffixWithAbbreviations(suffix) {
  const s = String(suffix || '');
  if (!s) return '';

  const map = getAbbreviationsMapSync();
  if (!map || typeof map !== 'object') return s;

  const trimmed = s.trim();
  if (!(trimmed.startsWith('(') && trimmed.endsWith(')'))) return s;

  const inner = trimmed.slice(1, -1);
  if (!inner) return s;

  // Replace any known tokens with "TOKEN – Expansion".
  // Uses Unicode property escapes to support non-ASCII letters.
  const tokenRe = /[\p{L}0-9]+(?:-[\p{L}0-9]+)*/gu;
  const expandedInner = inner.replace(tokenRe, (tok) => {
    const full = map[tok];
    return full ? `${tok}${EN_DASH}${full}` : tok;
  });

  return ` (${expandedInner})`;
}

function formatAbbrExpansion(map, raw) {
  const rawTrim = String(raw || "").trim();
  if (!rawTrim) return "";

  // Expand via abbreviations.tsv first (exact match).
  const abbrMap = getAbbreviationsMapSync();
  const exactFromAbbr = (abbrMap && typeof abbrMap === 'object') ? (abbrMap[rawTrim] || null) : null;
  if (exactFromAbbr) return `${rawTrim}${EN_DASH}${exactFromAbbr}`;

  // Then expand via texts.tsv map (exact match) to support keys like "K (E)".
  const exactFromTexts = (map && typeof map === 'object') ? (map[rawTrim] || null) : null;
  if (exactFromTexts) return `${rawTrim}${EN_DASH}${exactFromTexts}`;

  const { base, suffix } = parseTrailingParenSuffix(rawTrim);
  if (!base) return "";

  // Expand suffix codes (e.g. "(BGO)") via abbreviations.tsv.
  const expandedSuffix = suffix ? expandParenSuffixWithAbbreviations(suffix) : "";

  // Expand via abbreviations.tsv first (base match).
  const fullFromAbbr = (abbrMap && typeof abbrMap === 'object') ? (abbrMap[base] || null) : null;
  if (fullFromAbbr) return `${base}${EN_DASH}${fullFromAbbr}${expandedSuffix}`;

  // Then expand via texts.tsv map (base match).
  const full = map && typeof map === 'object' ? (map[base] || null) : null;
  if (full) return `${base}${EN_DASH}${full}${expandedSuffix}`;

  // If only the suffix could be expanded, still return the expanded suffix.
  if (expandedSuffix && expandedSuffix !== suffix) return `${base}${expandedSuffix}`;

  return rawTrim;
}

function getEffectiveValue(msRows, rowIndex, field) {
  if (!msRows || !Array.isArray(msRows)) return "";
  for (let i = rowIndex; i >= 0; i--) {
    const v = normalizeForCompare(msRows[i] && msRows[i][field]);
    if (v) return v;
  }
  return "";
}

const MINOR_TEXT_KEY_SEP = "\u0000";
function minorTextKey(mainAbbr, sectionAbbr) {
  return `${String(mainAbbr || "").trim()}${MINOR_TEXT_KEY_SEP}${String(sectionAbbr || "").trim()}`;
}

function renderMergedCell(field, value, ctx = null) {
  const v = normalizeForCompare(value);
  if (!v) return "&nbsp;";
  if (field === "Links to Database") {
    const html = renderLinksToDatabaseHtml(v);
    return html ? html : "&nbsp;";
  }
  if (field === "Leaves/Pages") return escapeHtml(normalizeLeavesPages(v));
  if (field === "Main text") {
    const map = (typeof window !== 'undefined' && window.MAIN_TEXT_MAP)
      ? window.MAIN_TEXT_MAP
      : (typeof MAIN_TEXT_MAP !== 'undefined' ? MAIN_TEXT_MAP : null);

    if (map && typeof map === 'object') {
      const parts = splitSemicolonList(v);
      const expanded = parts.map(key => formatAbbrExpansion(map, key)).join('; ');
      return escapeHtml(expanded);
    }
  }

  if (field === "Minor text") {
    const minorMap = (typeof window !== 'undefined' && window.MINOR_TEXT_MAP)
      ? window.MINOR_TEXT_MAP
      : (typeof MINOR_TEXT_MAP !== 'undefined' ? MINOR_TEXT_MAP : null);

    if (minorMap && typeof minorMap === 'object') {
      const parts = splitSemicolonList(v);

      // Determine main-text context for resolving minor sections.
      let mainCandidates = [];
      if (ctx && ctx.msRows && typeof ctx.rowIndex === 'number') {
        mainCandidates = splitSemicolonList(getEffectiveValue(ctx.msRows, ctx.rowIndex, "Main text"))
          .map(m => parseTrailingParenSuffix(m).base)
          .filter(Boolean);
      } else if (ctx && ctx.row && ctx.row["Main text"]) {
        mainCandidates = splitSemicolonList(ctx.row["Main text"])
          .map(m => parseTrailingParenSuffix(m).base)
          .filter(Boolean);
      }

      const expanded = parts.map(sectionAbbr => {
        const { base: sectionBase, suffix: sectionSuffix } = parseTrailingParenSuffix(sectionAbbr);
        let full = null;

        if (mainCandidates.length === 1) {
          full = minorMap[minorTextKey(mainCandidates[0], sectionBase)] || null;
        } else if (mainCandidates.length > 1) {
          const matches = [];
          for (const m of mainCandidates) {
            const hit = minorMap[minorTextKey(m, sectionBase)];
            if (hit) matches.push(hit);
          }
          if (matches.length === 1) full = matches[0];
        }

        if (full) {
          const expandedSuffix = sectionSuffix ? expandParenSuffixWithAbbreviations(sectionSuffix) : "";
          return `${sectionBase}${EN_DASH}${full}${expandedSuffix}`;
        }
        // Fall back to abbreviations.tsv (and suffix expansion) if applicable.
        return formatAbbrExpansion(null, sectionAbbr);
      }).join('; ');

      return escapeHtml(expanded);
    }
  }

  return escapeHtml(v);
}

function renderMergedView(manuscripts, metaInfo = null) {
  const mergedRoot = document.getElementById("merged-table");
  const meta = document.getElementById("merged-view-meta");
  if (!mergedRoot) return;

  const columnsFull = (DISPLAY_COLUMNS && Array.isArray(DISPLAY_COLUMNS))
    ? DISPLAY_COLUMNS.slice()
    : ((DATA_HEADERS && Array.isArray(DATA_HEADERS)) ? DATA_HEADERS.slice() : COLUMN_ORDER.slice());
  const columns = getMergedVisibleColumnsArray(columnsFull);

  let pageRows = 0;
  manuscripts.forEach(m => { pageRows += m.rows.length; });

  const totalManuscripts = (metaInfo && typeof metaInfo.totalManuscripts === 'number')
    ? metaInfo.totalManuscripts
    : manuscripts.length;
  const totalRows = (metaInfo && typeof metaInfo.totalRows === 'number')
    ? metaInfo.totalRows
    : pageRows;
  const page = (metaInfo && typeof metaInfo.page === 'number') ? metaInfo.page : 1;
  const totalPages = (metaInfo && typeof metaInfo.totalPages === 'number') ? metaInfo.totalPages : 1;

  if (meta) {
    let note = "";
    if (!RAW_EXCEL_LOADED && !RAW_EXCEL_FAILED) note = " (loading raw Excel merges…)";
    if (RAW_EXCEL_FAILED) note = " (raw Excel files not found; showing fallback view)";
    meta.textContent = note;
  }
  const totalEl = document.getElementById('total-records');
  if (totalEl) totalEl.textContent = `${totalManuscripts} manuscripts (${totalRows} texts)`;

  updateMergedPagerUI(page, totalPages);

  let html = '<table class="table table-bordered merged-table"><thead><tr>';
  for (const col of columns) {
    html += `<th>${escapeHtml(getColumnTitle(col))}</th>`;
  }
  html += '</tr></thead><tbody>';

  manuscripts.forEach((ms, msIndex) => {
    const msClass = (msIndex % 2 === 0) ? 'ms-a' : 'ms-b';
    // Prefer raw Excel-like rows (blank cells preserved) if available
    const rawBlock = RAW_BY_MANUSCRIPT_KEY.get(ms.key);
    const msRows = (rawBlock && rawBlock.rows && rawBlock.rows.length > 0) ? rawBlock.rows : ms.rows;
    const msRowCount = msRows.length;

    // Manuscript-level merged fields
    const mergedLinksToDatabase = aggregateFieldValues(msRows, "Links to Database", "; ");
    const literatureAllEmpty = allFieldValuesEmpty(msRows, "Literature");

    const mergeLookup = (rawBlock && RAW_MERGES_BY_SOURCE.has(rawBlock.sourceId))
      ? buildLocalMergeLookup(msRows, rawBlock.sourceId, columnsFull, columns)
      : null;
    const msConst = mergeLookup ? null : getConstantFields(msRows, columns, ["Production Unit"]);

    // Build contiguous Production Unit runs so we can merge/stripe without re-ordering rows.
    const runs = [];
    const rowToRunIndex = new Array(msRowCount);
    let lastPU = "";
    let runStart = 0;
    let runIndex = 0;
    for (let i = 0; i < msRowCount; i++) {
      const v = normalizeForCompare(msRows[i]["Production Unit"]);
      const pu = v || lastPU;
      if (i === 0) {
        lastPU = pu;
        runStart = 0;
        runIndex = 0;
      } else if (pu !== lastPU) {
        runs.push({ start: runStart, end: i - 1, index: runIndex, pu: lastPU });
        runIndex++;
        runStart = i;
        lastPU = pu;
      }
      rowToRunIndex[i] = runIndex;
    }
    runs.push({ start: runStart, end: msRowCount - 1, index: runIndex, pu: lastPU });

    // In heuristic mode we can still merge constant fields within each run.
    const runConstByStart = new Map();
    if (!mergeLookup) {
      for (const r of runs) {
        const slice = msRows.slice(r.start, r.end + 1);
        runConstByStart.set(r.start, getConstantFields(slice, columns, []));
      }
    }

    for (let rIndex = 0; rIndex < msRowCount; rIndex++) {
      const row = msRows[rIndex];
      const isFirstMsRow = (rIndex === 0);
      const trClasses = [msClass];
      if (isFirstMsRow) trClasses.push('ms-sep');
      html += `<tr class="${trClasses.join(' ')}">`;

      const currentRunIdx = rowToRunIndex[rIndex] || 0;
      const puClass = (currentRunIdx % 2 === 0) ? 'pu-a' : 'pu-b';

      // Find the start/end for this run (used for rowspan)
      let runStartIdx = rIndex;
      let runEndIdx = rIndex;
      // Cheap lookup: scan runs array (runs count per manuscript is small)
      for (const rr of runs) {
        if (rr.start <= rIndex && rIndex <= rr.end) {
          runStartIdx = rr.start;
          runEndIdx = rr.end;
          break;
        }
      }
      const isRunStart = (rIndex === runStartIdx);
      const runRowSpan = (runEndIdx - runStartIdx) + 1;
      const runConst = (!mergeLookup && isRunStart) ? runConstByStart.get(runStartIdx) : null;

      for (let c = 0; c < columns.length; c++) {
        const col = columns[c];

        // Always merge Links to Database at the manuscript level (combine URLs across rows).
        if (col === "Links to Database") {
          if (!isFirstMsRow) continue;
          html += `<td class="${msClass}" rowspan="${msRowCount}">${renderMergedCell(col, mergedLinksToDatabase)}</td>`;
          continue;
        }

        // If Literature is empty for all rows in the manuscript, merge it into a single blank cell.
        if (col === "Literature" && literatureAllEmpty) {
          if (!isFirstMsRow) continue;
          html += `<td class="${msClass}" rowspan="${msRowCount}">${renderMergedCell(col, "")}</td>`;
          continue;
        }

        if (mergeLookup) {
          // The raw Excel files don't have a Language column; we inject it.
          // Merge it per manuscript so duplicates don't repeat visually.
          if (col === "Language") {
            if (!isFirstMsRow) continue;
            html += `<td class="${msClass}" rowspan="${msRowCount}">${renderMergedCell(col, row[col], { row, msRows, rowIndex: rIndex })}</td>`;
            continue;
          }

          const key = `${rIndex},${c}`;
          if (mergeLookup.covered.has(key)) continue;
          const span = mergeLookup.topLeft.get(key);
          const attrs = span ? ` rowspan="${span.rowSpan}" colspan="${span.colSpan}"` : "";
          html += `<td class="${(col === "Production Unit") ? puClass : msClass}"${attrs}>${renderMergedCell(col, row[col], { row, msRows, rowIndex: rIndex })}</td>`;
          continue;
        }

        if (msConst && msConst.has(col)) {
          if (!isFirstMsRow) continue;
          html += `<td class="${msClass}" rowspan="${msRowCount}">${renderMergedCell(col, row[col], { row, msRows, rowIndex: rIndex })}</td>`;
          continue;
        }

        if (col === "Production Unit") {
          if (!isRunStart) continue;
          html += `<td class="${puClass}" rowspan="${runRowSpan}">${renderMergedCell(col, row[col], { row, msRows, rowIndex: rIndex })}</td>`;
          continue;
        }

        if (runConst && runConst.has(col)) {
          html += `<td class="${puClass}" rowspan="${runRowSpan}">${renderMergedCell(col, row[col], { row, msRows, rowIndex: rIndex })}</td>`;
          continue;
        }

        html += `<td>${renderMergedCell(col, row[col], { row, msRows, rowIndex: rIndex })}</td>`;
      }

      html += '</tr>';
    }
  });

  html += '</tbody></table>';
  mergedRoot.innerHTML = html;
}

// Apply view-dependent UI visibility without changing filters.
// (Important: the table may be created after initial setView(),
// so we sometimes need to re-apply visibility once it's created.)
function mountRecordCount() {
  const recordCountEl = document.getElementById('record-count');
  if (!recordCountEl) return;

  const hostId = (currentView === 'table') ? 'table-view-count-host' : 'merged-view-count-host';
  const hostEl = document.getElementById(hostId);
  if (!hostEl) return;

  if (recordCountEl.parentElement !== hostEl) {
    hostEl.appendChild(recordCountEl);
  }
}

function applyViewUI() {
  try {
    document.body.classList.toggle('view-merged', currentView === 'merged');
    document.body.classList.toggle('view-table', currentView === 'table');
  } catch (e) {
    // ignore
  }

  const tableView = document.getElementById("table-view");
  const mergedView = document.getElementById("merged-view");
  if (tableView) tableView.style.display = (currentView === "table") ? "block" : "none";
  if (mergedView) mergedView.style.display = (currentView === "merged") ? "block" : "none";

  // Column selector lives in the top control bar; only show it for Manuscript View.
  const mergedColsControl = document.getElementById("merged-columns-control");
  if (mergedColsControl) mergedColsControl.style.display = (currentView === "merged") ? "" : "none";

  // Column selector for Text View.
  const tableColsControl = document.getElementById("table-columns-control");
  if (tableColsControl) tableColsControl.style.display = (currentView === "table") ? "" : "none";

  // Sort selector lives in the top control bar; only show it for Manuscript View.
  const mergedSortControl = document.getElementById("merged-sort-control");
  // Also useful for Text View sorting (shelfmark/dating)
  if (mergedSortControl) mergedSortControl.style.display = "";

  if (currentView === "table" && table && typeof table.redraw === 'function') {
    try { table.redraw(true); } catch (e) {}
  }

  const paginationSizeSelect = document.getElementById("pagination-size");
  if (paginationSizeSelect) paginationSizeSelect.disabled = false;

  const paginationLabel = document.getElementById('pagination-size-label');
  if (paginationLabel) paginationLabel.textContent = 'Results per page:';

  // Show the page-size selector for both views.
  const paginationLabelControl = document.getElementById("pagination-label-control");
  const paginationSelectControl = document.getElementById("pagination-select-control");
  if (paginationLabelControl) paginationLabelControl.style.display = "";
  if (paginationSelectControl) paginationSelectControl.style.display = "";

  // Keep merged column menu in sync
  if (currentView === "merged") {
    renderMergedColumnsMenu();
  }

  // Keep Text View column menu in sync
  if (currentView === "table") {
    renderTableColumnsMenu();
    updateTextViewColumnsFromVisibility();
  }

  // Hide Manuscript pager UI when not in Manuscript View.
  if (currentView !== 'merged') {
    const pager = document.getElementById('merged-pager');
    if (pager) pager.style.display = 'none';
    const indicator = document.getElementById('merged-page-indicator');
    if (indicator) indicator.textContent = '';
    const mergedPageControl = document.getElementById('merged-page-control');
    if (mergedPageControl) mergedPageControl.style.display = 'none';
  } else {
    const mergedPageControl = document.getElementById('merged-page-control');
    if (mergedPageControl) mergedPageControl.style.display = '';
  }

  // Text View pager control lives in the top control bar.
  const tablePageControl = document.getElementById('table-page-control');
  if (tablePageControl) tablePageControl.style.display = (currentView === 'table') ? '' : 'none';

  // Move the shared record-count element into the active view header.
  mountRecordCount();
}

function setView(view) {
  currentView = (view === "merged") ? "merged" : "table";
  const viewSelect = document.getElementById("view-select");
  if (viewSelect) viewSelect.value = currentView;

  try { localStorage.setItem(VIEW_STORAGE_KEY, currentView); } catch (e) {}
  applyViewUI();

  if (currentView === "merged") {
    // Kick off raw Excel + merge JSON loading (if available) and re-render when ready.
    const meta = document.getElementById("merged-view-meta");
    if (meta && !RAW_EXCEL_LOADED && !RAW_EXCEL_FAILED) {
      meta.textContent = "Loading raw Excel layout (if available)…";
    }
    ensureRawExcelLoaded().then(() => {
      if (currentView === "merged") applyFacetFilters();
    });
  }

  applyFacetFilters();
}

// Map 'language' column to 'Language' for facets, expanding ISO codes
const LANGUAGE_MAP = {
  'da': 'Danish',
  'is': 'Icelandic',
  'no': 'Norwegian',
  'sv': 'Swedish',
  // Add more as needed
};
function normalizeRowLanguage(row) {
  if (row["Language"]) {
    const code = row["Language"].toLowerCase();
    row["Language"] = LANGUAGE_MAP[code] || row["Language"];
  }
  return row;
}

async function loadAbbreviationsMap() {
  if (ABBREVIATIONS_MAP) {
    window.ABBREVIATIONS_MAP = ABBREVIATIONS_MAP;
    return ABBREVIATIONS_MAP;
  }
  try {
    const resp = await fetch('data/abbreviations.tsv');
    const text = await resp.text();
    const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
    const map = {};
    lines.forEach(line => {
      const cleaned = String(line || '').replace(/^\uFEFF/, '');
      const cols = cleaned.split('\t');
      const abbr = (cols[0] || '').trim();
      if (!abbr) return;

      // Files often use: abbr \t \t expansion
      let full = '';
      for (let i = cols.length - 1; i >= 1; i--) {
        const v = (cols[i] || '').trim();
        if (v) { full = v; break; }
      }
      if (full) map[abbr] = full;
    });
    ABBREVIATIONS_MAP = map;
    window.ABBREVIATIONS_MAP = map;
    return map;
  } catch (e) {
    console.error('Failed to load abbreviations.tsv:', e);
    ABBREVIATIONS_MAP = {};
    window.ABBREVIATIONS_MAP = {};
    return {};
  }
}

// Main text abbreviation mapping, loaded from texts.tsv at runtime
let MAIN_TEXT_MAP = null;
let MINOR_TEXT_MAP = null;
async function loadMainTextMap() {
  if (MAIN_TEXT_MAP && MINOR_TEXT_MAP) {
    // Always assign to window for modal rendering
    window.MAIN_TEXT_MAP = MAIN_TEXT_MAP;
    window.MINOR_TEXT_MAP = MINOR_TEXT_MAP;
    return MAIN_TEXT_MAP;
  }
  try {
    const resp = await fetch('data/texts.tsv');
    const text = await resp.text();
    const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
    const mainMap = {};
    const minorMap = {};

    lines.forEach(line => {
      const cleaned = String(line || '').replace(/^\uFEFF/, '');
      const cols = cleaned.split('\t');

      // Preferred (new) format: main_abbr \t section_abbr \t expansion
      if (cols.length >= 3) {
        const mainAbbr = (cols[0] || '').trim();
        const sectionAbbr = (cols[1] || '').trim();
        const full = (cols[2] || '').trim();
        if (!mainAbbr || !full) return;

        if (!sectionAbbr) {
          mainMap[mainAbbr] = full;
        } else {
          minorMap[minorTextKey(mainAbbr, sectionAbbr)] = full;
        }
        return;
      }

      // Legacy fallback: main_abbr \t expansion
      if (cols.length === 2) {
        const mainAbbr = (cols[0] || '').trim();
        const full = (cols[1] || '').trim();
        if (mainAbbr && full) mainMap[mainAbbr] = full;
      }
    });

    MAIN_TEXT_MAP = mainMap;
    MINOR_TEXT_MAP = minorMap;
    window.MAIN_TEXT_MAP = mainMap;
    window.MINOR_TEXT_MAP = minorMap;
    return mainMap;
  } catch (e) {
    console.error('Failed to load texts.tsv:', e);
    MAIN_TEXT_MAP = {};
    MINOR_TEXT_MAP = {};
    window.MAIN_TEXT_MAP = {};
    window.MINOR_TEXT_MAP = {};
    return {};
  }
}

// Helper: parse a Dating string to a century (returns string like '13th', '14th', etc. or '')
function parseCentury(dating) {
  if (!dating || typeof dating !== 'string') return '';
  // Try to match '14th century', '13th cent.', 'c. 1200', 'late 1200s', 'early 1300s', 'c. 1350', etc.
  let m = dating.match(/(\d{3,4})/);
  if (m) {
    let year = parseInt(m[1], 10);
    if (year >= 1000 && year < 2000) {
      let cent = Math.floor((year - 1) / 100) + 1;
      return cent + 'th';
    }
  }
  // Try to match '14th century', '13th cent.'
  m = dating.match(/(\d{1,2})(?:th|st|nd|rd)\s*cent/i);
  if (m) {
    return m[1] + 'th';
  }
  // Try to match 'late 1200s', 'early 1300s'
  m = dating.match(/(\d{3})00s/);
  if (m) {
    let year = parseInt(m[1] + '00', 10);
    let cent = Math.floor((year - 1) / 100) + 1;
    return cent + 'th';
  }
  return '';
}

// Facet value normalization (UI-only)
// Keeps the underlying table values intact, but ensures empty values
// can be filtered as a concrete facet option.
//
// Important: We distinguish between a truly empty cell (render as "Empty")
// and a cell containing the literal text "Unknown" (render as "Unknown").
const FACET_FIELDS = [
  "Language",
  "Depository",
  "Object",
  "Material",
  "Size",
  "Main text group",
  "Dating",
  "Script",
  "Pricking",
  "Ruling",
  "Columns",
  "Lines",
  "Rubric",
  "Style",
];

const FACET_EMPTY_LABEL_FIELDS = new Set([
  "Depository",
  "Object",
  "Material",
  "Size",
  "Script",
  "Pricking",
  "Ruling",
  "Columns",
  "Lines",
  "Rubric",
  "Style",
]);

function getFacetValue(row, field) {
  if (!row) return '';
  if (FACET_EMPTY_LABEL_FIELDS.has(field)) {
    const raw = row[field];
    const s = (raw === null || raw === undefined) ? "" : String(raw).trim();
    return s === "" ? "Empty" : s;
  }
  return row[field];
}

function getUniqueValues(rows, field) {
  const set = new Set();
  rows.forEach(row => {
    let val = getFacetValue(row, field);
    if (val === null || val === undefined) return;
    // Most facets should not show empty values; Object/Material are normalized to "Unknown" above.
    if (typeof val === 'string' && val.trim() === "") return;
    if (field === "Dating") {
      // Only keep year or year range
      let m = val.match(/(\d{3,4})\s*[-–]\s*(\d{3,4})/);
      if (m) val = m[1] + '-' + m[2];
      else {
        m = val.match(/(\d{4})/);
        if (m) val = m[1];
        else {
          m = val.match(/(\d{3})00s/);
          if (m) val = m[1] + '00s';
          else {
            m = val.match(/(\d{1,2})(?:th|st|nd|rd)\s*cent/i);
            if (m) val = m[1] + 'th cent.';
            else return;
          }
        }
      }
    }
    set.add(val);
  });
  return Array.from(set).sort((a, b) => a.localeCompare(b, undefined, {numeric:true, sensitivity:"base"}));
}

async function renderFacetSidebar(rows) {
  await loadAbbreviationsMap();
  const mainTextMap = await loadMainTextMap();
  FACET_FIELDS.forEach(field => {
    const facetDiv = document.getElementById(`facet-${field}`);
    if (!facetDiv) return;
    if (field === "Dating") {
      // Render a year-range selector based on parsed DatingYear
      const years = rows
        .map(r => r["DatingYear"])
        .filter(y => typeof y === 'number' && !Number.isNaN(y));
      const minYear = years.length ? Math.min(...years) : '';
      const maxYear = years.length ? Math.max(...years) : '';

      let html = `<div class="fw-bold mb-2">Dating</div>`;
      html += `<div class="small text-secondary mb-2">Filter by year range (uses parsed year from Dating)</div>`;
      html += `<div class="d-flex gap-2 align-items-end mb-2">
        <div class="flex-fill">
          <label class="form-label mb-1" style="font-size:0.75rem; text-transform:uppercase;">From</label>
          <input class="form-control form-control-sm" type="number" inputmode="numeric" data-dating-range="min" id="facet-Dating-min" placeholder="${escapeHtml(minYear)}">
        </div>
        <div class="flex-fill">
          <label class="form-label mb-1" style="font-size:0.75rem; text-transform:uppercase;">To</label>
          <input class="form-control form-control-sm" type="number" inputmode="numeric" data-dating-range="max" id="facet-Dating-max" placeholder="${escapeHtml(maxYear)}">
        </div>
        <button class="btn btn-sm btn-outline-secondary" type="button" data-dating-range-clear>Clear</button>
      </div>`;
      facetDiv.innerHTML = html;
      return;
    }
    if (field === "Main text group") {
      // Build a map: group -> Set(variants)
      const groupMap = {};
      rows.forEach(row => {
        const group = row["Main text group"] || "";
        let variant = "";
        if (row["Main text"]) {
          let m = row["Main text"].match(/^([^\(]+)\s*\(([^\)]+)\)/);
          if (m) variant = m[2].trim();
        }
        if (!groupMap[group]) groupMap[group] = new Set();
        if (variant) groupMap[group].add(variant);
      });
      let html = `<div class="fw-bold mb-2">Main text group</div>`;
      html += `<div class="form-check mb-1"><input class="form-check-input" type="checkbox" value="__ALL__" checked data-facet="Main text group" id="facet-mtg-all"><label class="form-check-label" for="facet-mtg-all">All</label></div>`;
      Object.keys(groupMap).sort().forEach((group, gi) => {
        // Skip empty or dot-only group values
        if (group === '' || group.trim() === '.') return;
        const groupId = `facet-mtg-group-${gi}`;
        const groupLabel = mainTextMap[group]
          ? escapeHtml(formatAbbrExpansion(mainTextMap, group))
          : escapeHtml(group);
        html += `<div style='margin-left:0.5em;'><div class="form-check mb-1"><input class="form-check-input" type="checkbox" value="${escapeHtml(group)}" data-facet="Main text group" id="${groupId}"><label class="form-check-label" for="${groupId}">${groupLabel}</label></div>`;
        const variants = Array.from(groupMap[group]);
        if (variants.length > 0) {
          html += `<div style='margin-left:1.5em;'>`;
          variants.sort().forEach((variant, vi) => {
            // Skip empty or dot-only variant values
            if (variant === '' || variant.trim() === '.') return;
            const variantId = `facet-mtg-variant-${gi}-${vi}`;
            const variantLabel = mainTextMap[variant]
              ? escapeHtml(formatAbbrExpansion(mainTextMap, variant))
              : escapeHtml(variant);
            html += `<div class="form-check mb-1"><input class="form-check-input" type="checkbox" value="${escapeHtml(group + '|' + variant)}" data-facet="Main text group-variant" id="${variantId}"><label class="form-check-label" for="${variantId}">${variantLabel}</label></div>`;
          });
          html += `</div>`;
        }
        html += `</div>`;
      });
      facetDiv.innerHTML = html;
      return;
    }
    // Default facet rendering
    const values = getUniqueValues(rows, field);
    let html = `<div class="fw-bold mb-2">${field}</div>`;
    html += `<div class="form-check mb-1"><input class="form-check-input" type="checkbox" value="__ALL__" checked data-facet="${field}" id="facet-${field}-all"><label class="form-check-label" for="facet-${field}-all">All</label></div>`;
    values.forEach((val, i) => {
      // Skip empty or dot-only values for Main text facet
      if (field === "Main text" && (val === '' || val.trim() === '.' )) return;
      let label = val;
      if (field === "Main text" && mainTextMap[val]) {
        label = formatAbbrExpansion(mainTextMap, val);
      }
      const id = `facet-${field}-${i}`;
      html += `<div class="form-check mb-1"><input class="form-check-input" type="checkbox" value="${escapeHtml(val)}" data-facet="${escapeHtml(field)}" id="${id}"><label class="form-check-label" for="${id}">${escapeHtml(label)}</label></div>`;
    });
    facetDiv.innerHTML = html;
  });
}

function getFacetSelections() {
  const selections = {};
    FACET_FIELDS.forEach(field => {
      const facetDiv = document.getElementById(`facet-${field}`);
      if (!facetDiv) return;
      if (field === "Dating") {
        const minEl = facetDiv.querySelector('input[data-dating-range="min"]');
        const maxEl = facetDiv.querySelector('input[data-dating-range="max"]');
        const minVal = minEl ? minEl.value.trim() : '';
        const maxVal = maxEl ? maxEl.value.trim() : '';
        if (minVal !== '' || maxVal !== '') {
          selections["DatingRange"] = {
            min: minVal === '' ? null : Number(minVal),
            max: maxVal === '' ? null : Number(maxVal),
          };
        }
        return;
      }
      if (field === "Main text group") {
        const groupBoxes = facetDiv.querySelectorAll('input[type=checkbox][data-facet="Main text group"]');
        const checkedGroups = Array.from(groupBoxes).filter(cb => cb.checked && cb.value !== "__ALL__").map(cb => cb.value);
        if (checkedGroups.length > 0) selections[field] = checkedGroups;
        const variantBoxes = facetDiv.querySelectorAll('input[type=checkbox][data-facet="Main text group-variant"]');
        const checkedVariants = Array.from(variantBoxes).filter(cb => cb.checked).map(cb => cb.value);
        if (checkedVariants.length > 0) selections["Main text group-variant"] = checkedVariants;
        return;
      }
      const checkboxes = facetDiv.querySelectorAll('input[type=checkbox][data-facet]');
      const checked = Array.from(checkboxes).filter(cb => cb.checked && cb.value !== "__ALL__").map(cb => cb.value);
      if (checked.length > 0) selections[field] = checked;
    });
    return selections;
}

function updateFacetAllCheckbox(field) {
  const facetDiv = document.getElementById(`facet-${field}`);
  if (!facetDiv) return;
  const allBox = facetDiv.querySelector('input[type=checkbox][value="__ALL__"]');
  const otherBoxes = (field === "Main text group")
    ? facetDiv.querySelectorAll('input[type=checkbox][data-facet="Main text group"]:not([value="__ALL__"])')
    : facetDiv.querySelectorAll('input[type=checkbox][data-facet]:not([value="__ALL__"])');
  if (!allBox) return;
  const allChecked = Array.from(otherBoxes).every(cb => cb.checked);
  allBox.checked = allChecked || Array.from(otherBoxes).every(cb => !cb.checked);
}

function setupFacetEvents() {
  FACET_FIELDS.forEach(field => {
    const facetDiv = document.getElementById(`facet-${field}`);
    if (!facetDiv) return;

    if (field === "Dating") {
      // Range inputs
      facetDiv.addEventListener('input', (e) => {
        if (e.target && e.target.matches('input[data-dating-range]')) {
          applyFacetFilters();
        }
      });
      facetDiv.addEventListener('click', (e) => {
        const btn = e.target && e.target.closest && e.target.closest('button[data-dating-range-clear]');
        if (!btn) return;
        const minEl = facetDiv.querySelector('input[data-dating-range="min"]');
        const maxEl = facetDiv.querySelector('input[data-dating-range="max"]');
        if (minEl) minEl.value = '';
        if (maxEl) maxEl.value = '';
        applyFacetFilters();
      });
      return;
    }

    facetDiv.addEventListener('change', (e) => {
      if (!e.target.matches('input[type=checkbox][data-facet]')) return;
      if (e.target.value === "__ALL__") {
        // All box toggled: check/uncheck all
        const allChecked = e.target.checked;
        const selector = (field === "Main text group")
          ? 'input[type=checkbox][data-facet="Main text group"]:not([value="__ALL__"])'
          : 'input[type=checkbox][data-facet]:not([value="__ALL__"])';
        facetDiv.querySelectorAll(selector).forEach(cb => {
          cb.checked = allChecked;
        });
      } else {
        // If all boxes checked, check All; if none checked, check All
        // For Main text group, don't let variant toggles affect the group "All" state.
        if (!(field === "Main text group" && e.target.dataset.facet === "Main text group-variant")) {
          updateFacetAllCheckbox(field);
        }
      }
      applyFacetFilters();
    });
  });
}

function applyFacetFilters() {
  facetSelections = getFacetSelections();
  const searchInput = document.getElementById('search');
  const query = searchInput ? searchInput.value.toLowerCase() : "";

  if (currentView === "merged") {
    const sortMode = getMergedSortMode();
    const manuscripts = groupByPreserveOrder(allRows || [], getManuscriptKey)
      .map(g => ({ key: g.key, rows: g.rows }))
      .sort((a, b) => compareManuscripts(a, b, sortMode));

    const range = facetSelections["DatingRange"];
    let min = range ? range.min : null;
    let max = range ? range.max : null;
    if (typeof min === 'number' && Number.isNaN(min)) min = null;
    if (typeof max === 'number' && Number.isNaN(max)) max = null;
    if (min !== null && max !== null && min > max) {
      const tmp = min; min = max; max = tmp;
    }

    function manuscriptMatches(ms) {
      for (const field of FACET_FIELDS) {
        if (field === "Dating") {
          if (min === null && max === null) continue;
          const anyInRange = ms.rows.some(r => {
            const y = r["DatingYear"];
            if (typeof y !== 'number' || Number.isNaN(y)) return false;
            if (min !== null && y < min) return false;
            if (max !== null && y > max) return false;
            return true;
          });
          if (!anyInRange) return false;
          continue;
        }

        if (field === "Main text group") {
          const variantSelections = facetSelections["Main text group-variant"];
          if (variantSelections && variantSelections.length > 0) {
            const anyVariant = ms.rows.some(r => {
              const group = r["Main text group"] || "";
              let variant = "";
              if (r["Main text"]) {
                const m = String(r["Main text"]).match(/^([^\(]+)\s*\(([^\)]+)\)/);
                if (m) variant = m[2].trim();
              }
              const key = group + "|" + variant;
              return variantSelections.includes(key);
            });
            if (!anyVariant) return false;
            continue;
          }
          const groups = facetSelections[field];
          if (groups && groups.length > 0) {
            const anyGroup = ms.rows.some(r => groups.includes(getFacetValue(r, field)));
            if (!anyGroup) return false;
          }
          continue;
        }

        const selected = facetSelections[field];
        if (selected && selected.length > 0) {
          const anyMatch = ms.rows.some(r => selected.includes(getFacetValue(r, field)));
          if (!anyMatch) return false;
        }
      }

      if (query) {
        const anySearch = ms.rows.some(r => Object.values(r).some(val => String(val).toLowerCase().includes(query)));
        if (!anySearch) return false;
      }

      return true;
    }

    const filtered = manuscripts.filter(manuscriptMatches);

    let totalRows = 0;
    filtered.forEach(m => { totalRows += m.rows.length; });

    const pageSize = getMergedPageSize();
    const totalPages = (!Number.isFinite(pageSize)) ? 1 : Math.max(1, Math.ceil(filtered.length / pageSize));

    if (MERGED_PAGE < 1) MERGED_PAGE = 1;
    if (MERGED_PAGE > totalPages) MERGED_PAGE = totalPages;

    let pageItems = filtered;
    if (Number.isFinite(pageSize)) {
      const start = (MERGED_PAGE - 1) * pageSize;
      const end = start + pageSize;
      pageItems = filtered.slice(start, end);
    }

    renderMergedView(pageItems, {
      totalManuscripts: filtered.length,
      totalRows,
      page: MERGED_PAGE,
      totalPages,
    });

    // Keep the Text View table internally in sync (even though hidden)
    if (table) {
      const allowed = new Set(filtered.map(m => m.key));
      table.clearFilter(true);
      table.setFilter(function(row) {
        return allowed.has(getManuscriptKey(row));
      });
    }

    return;
  }

  if (!table) return;
  table.clearFilter(true);
  // Compose filter function
  table.setFilter(function(row) {
    // 1. Check Facets
    for (const field of FACET_FIELDS) {
      if (field === "Dating") {
        const range = facetSelections["DatingRange"];
        if (range && (range.min !== null || range.max !== null)) {
          let min = range.min;
          let max = range.max;
          if (typeof min === 'number' && Number.isNaN(min)) min = null;
          if (typeof max === 'number' && Number.isNaN(max)) max = null;
          if (min !== null && max !== null && min > max) {
            const tmp = min; min = max; max = tmp;
          }
          const y = row["DatingYear"];
          if (typeof y !== 'number' || Number.isNaN(y)) return false;
          if (min !== null && y < min) return false;
          if (max !== null && y > max) return false;
        }
        continue;
      }
      if (field === "Main text group") {
        if (facetSelections["Main text group-variant"] && facetSelections["Main text group-variant"].length > 0) {
          let matched = false;
          let group = row["Main text group"] || "";
          let variant = "";
          if (row["Main text"]) {
            let m = row["Main text"].match(/^([^\(]+)\s*\(([^\)]+)\)/);
            if (m) variant = m[2].trim();
          }
          let key = group + "|" + variant;
          if (facetSelections["Main text group-variant"].includes(key)) matched = true;
          if (!matched) return false;
        } else if (facetSelections[field] && facetSelections[field].length > 0) {
          if (!facetSelections[field].includes(row[field])) return false;
        }
      } else {
        if (facetSelections[field] && facetSelections[field].length > 0) {
          if (!facetSelections[field].includes(getFacetValue(row, field))) return false;
        }
      }
    }

    // 2. Check Global Search
    if (query) {
      const matchesSearch = Object.values(row).some(val =>
        String(val).toLowerCase().includes(query)
      );
      if (!matchesSearch) return false;
    }

    return true;
  });

  // Sorting is disabled; order is only changed when sort mode changes.
}



// Depository abbreviation mapping, loaded from depositories.tsv at runtime
let DEPOSITORY_MAP = null;

async function loadDepositoryMap() {
  if (DEPOSITORY_MAP) return DEPOSITORY_MAP;
  try {
    const resp = await fetch('data/depositories.tsv');
    const text = await resp.text();
    const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
    const map = {};
    lines.forEach(line => {
      const [abbr, name] = line.split(/\t/);
      if (abbr && name) map[abbr.trim()] = name.trim();
    });
    DEPOSITORY_MAP = map;
    return map;
  } catch (e) {
    console.error('Failed to load depositories.tsv:', e);
    DEPOSITORY_MAP = {};
    return {};
  }
}

async function loadDataTSV(fileName) {
    try {
      await loadAbbreviationsMap();
      const depositoryMap = await loadDepositoryMap();
      const response = await fetch(fileName);
      const tsvText = await response.text();
      const lines = tsvText.split(/\r?\n/).filter(line => line.trim() !== "");
      if (lines.length === 0) throw new Error("TSV file is empty");

      const headers = lines[0].split("\t");
      const rows = lines.slice(1).map(line => {
        const values = line.split("\t");
        const obj = {};
        headers.forEach((h, i) => {
          let val = values[i] !== undefined ? values[i] : "";
          if (h === "Main text") val = String(val || "").trim();
          if (h === "Leaves/Pages") val = normalizeLeavesPages(val);
          if (h === "Depository") {
            obj["Depository_abbr"] = val;
            if (val in depositoryMap) val = depositoryMap[val];
          }
          obj[h] = val;
        });

        if (obj["Main text"]) {
          const m = String(obj["Main text"]).match(/^([^\(]+)\s*\(/);
          obj["Main text group"] = m ? m[1].trim() : obj["Main text"];
        } else {
          obj["Main text group"] = "";
        }

        obj["Century"] = parseCentury(obj["Dating"]);
        obj["DatingYear"] = parseDatingYear(obj["Dating"]);
        normalizeRowLanguage(obj);
        return obj;
      });

      return loadDataFromParsedRows(headers, rows);
    } catch (err) {
      console.error("Error loading TSV file:", err);
      alert(`Could not load ${fileName}`);
    }
}

async function loadDataFromParsedRows(headers, rows) {
    if (!Array.isArray(headers) || headers.length === 0) throw new Error('Missing headers');
    const safeRows = Array.isArray(rows) ? rows : [];

    DATA_HEADERS = headers.slice();
    DISPLAY_COLUMNS = buildDisplayColumns(headers);

    loadMergedColumnVisibility();
    sanitizeMergedColumnVisibility();
    renderMergedColumnsMenu();

    loadTableColumnVisibility();
    sanitizeTableColumnVisibility();
    renderTableColumnsMenu();

    allRows = safeRows;
    TABLE_SOURCE_ROWS = safeRows;

    let colDefs = {};
    headers.forEach(h => {
      let colDef = {
        title: getColumnTitle(h),
        field: h,
        visible: true,
        hozAlign: 'left',
        width: ["Depository", "Shelf mark", "Production Unit"].includes(h) ? 200 : 100,
        headerSort: false,
        headerFilter: "input",
        headerFilterPlaceholder: "Filter...",
        headerFilterLiveFilter: true
      };

      if (h === "Links to Database") {
        colDef.formatter = function(cell) {
          const v = cell.getValue();
          return renderLinksToDatabaseHtml(v);
        };
      }

      if (h === "Main text") {
        colDef.formatter = function(cell) {
          const v = normalizeForCompare(cell.getValue());
          if (!v) return "";
          const map = (typeof window !== 'undefined' && window.MAIN_TEXT_MAP)
            ? window.MAIN_TEXT_MAP
            : (typeof MAIN_TEXT_MAP !== 'undefined' ? MAIN_TEXT_MAP : null);
          if (!map || typeof map !== 'object') return escapeHtml(v);

          const expanded = splitSemicolonList(v)
            .map(part => formatAbbrExpansion(map, part))
            .join('; ');
          return escapeHtml(expanded);
        };
      }

      if (h === "Minor text") {
        colDef.formatter = function(cell) {
          const v = normalizeForCompare(cell.getValue());
          if (!v) return "";

          const minorMap = (typeof window !== 'undefined' && window.MINOR_TEXT_MAP)
            ? window.MINOR_TEXT_MAP
            : (typeof MINOR_TEXT_MAP !== 'undefined' ? MINOR_TEXT_MAP : null);
          if (!minorMap || typeof minorMap !== 'object') return escapeHtml(v);

          const rowData = cell.getRow && cell.getRow() ? cell.getRow().getData() : null;
          const mainCandidates = splitSemicolonList(rowData && rowData["Main text"] ? rowData["Main text"] : "")
            .map(m => parseTrailingParenSuffix(m).base)
            .filter(Boolean);

          const expanded = splitSemicolonList(v).map(sectionAbbr => {
            const { base: sectionBase, suffix: sectionSuffix } = parseTrailingParenSuffix(sectionAbbr);
            let full = null;

            if (mainCandidates.length === 1) {
              full = minorMap[minorTextKey(mainCandidates[0], sectionBase)] || null;
            } else if (mainCandidates.length > 1) {
              const matches = [];
              for (const m of mainCandidates) {
                const hit = minorMap[minorTextKey(m, sectionBase)];
                if (hit) matches.push(hit);
              }
              if (matches.length === 1) full = matches[0];
            }

            if (full) {
              const expandedSuffix = sectionSuffix ? expandParenSuffixWithAbbreviations(sectionSuffix) : "";
              return `${sectionBase}${EN_DASH}${full}${expandedSuffix}`;
            }
            return formatAbbrExpansion(null, sectionAbbr);
          }).join('; ');

          return escapeHtml(expanded);
        };
      }

      colDefs[h] = colDef;
    });

    TEXT_COLUMN_DEFS = colDefs;

    const columnsForText = getTableVisibleColumnsArray(DISPLAY_COLUMNS || []);
    let columns = [];
    columnsForText.forEach(h => { if (colDefs[h]) columns.push(colDefs[h]); });

    const centuryCol = {
      title: "Century",
      field: "Century",
      visible: false,
      hozAlign: 'left',
      width: 90,
      headerSort: false,
      headerFilter: "input",
      headerFilterPlaceholder: "Filter...",
      headerFilterLiveFilter: true
    };
    columns.push(centuryCol);

    const datingYearCol = {
      title: "DatingYear",
      field: "DatingYear",
      visible: false,
      headerSort: false,
      sorter: function(a, b) {
        const ax = (typeof a === 'number' && !Number.isNaN(a)) ? a : Number.POSITIVE_INFINITY;
        const bx = (typeof b === 'number' && !Number.isNaN(b)) ? b : Number.POSITIVE_INFINITY;
        return ax - bx;
      }
    };
    columns.push(datingYearCol);

    TEXT_ALWAYS_COLUMNS = [centuryCol, datingYearCol];

    await renderFacetSidebar(safeRows);
    setupFacetEvents();

    const sortedRows = sortRowsForTextView(safeRows, getMergedSortMode());

    // Use the shared page-size selector as the default for Text View.
    const paginationSizeSelect = document.getElementById('pagination-size');
    const paginationRaw = paginationSizeSelect ? String(paginationSizeSelect.value || '').toLowerCase() : '20';
    const initialPageSize = (paginationRaw === 'all') ? true : parseInt(paginationRaw, 10);

    if (table) {
      table.setColumns(columns);
      await table.replaceData(sortedRows);
    } else {
      table = new SimpleTable("#table-view", { data: sortedRows, columns: columns, pageSize: initialPageSize });

      window.table = table;

      table.on("dataFiltered", function(filters, filteredRows){
        if (currentView === "table") {
          const totalEl = document.getElementById('total-records');
          if (totalEl) totalEl.textContent = filteredRows.length;
        }
      });

      // Attach rowClick event handler (modal)
      table.on("rowClick", function(e, row){
        const data = row.getData();
        const contentDiv = document.getElementById('row-details-content');
        if (!contentDiv) return;

        const entries = [];
        const dataKeys = Object.keys(data).filter(k => k !== "DatingYear" && k !== "_id" && k !== "Century" && k !== "Main text group" && k !== "Depository_abbr");

        const modalOrder = (Array.isArray(DISPLAY_COLUMNS) && DISPLAY_COLUMNS.length > 0)
          ? DISPLAY_COLUMNS
          : (Array.isArray(COLUMN_ORDER) ? COLUMN_ORDER : []);

        dataKeys.sort((a, b) => {
          const idxA = modalOrder.indexOf(a);
          const idxB = modalOrder.indexOf(b);
          if (idxA !== -1 && idxB !== -1) return idxA - idxB;
          if (idxA !== -1) return -1;
          if (idxB !== -1) return 1;
          return 0;
        });

        dataKeys.forEach(key => {
          let value = data[key] !== undefined && data[key] !== null ? data[key] : "";

          if (key === "Main text") {
            const v = normalizeForCompare(value);
            if (v) {
              const map = (typeof window !== 'undefined' && window.MAIN_TEXT_MAP)
                ? window.MAIN_TEXT_MAP
                : (typeof MAIN_TEXT_MAP !== 'undefined' ? MAIN_TEXT_MAP : null);
              if (map && typeof map === 'object') {
                value = splitSemicolonList(v).map(part => formatAbbrExpansion(map, part)).join('; ');
              }
            }
          }

          if (key === "Minor text") {
            const v = normalizeForCompare(value);
            if (v) {
              const minorMap = (typeof window !== 'undefined' && window.MINOR_TEXT_MAP)
                ? window.MINOR_TEXT_MAP
                : (typeof MINOR_TEXT_MAP !== 'undefined' ? MINOR_TEXT_MAP : null);
              if (minorMap && typeof minorMap === 'object') {
                const mainCandidates = splitSemicolonList(data && data["Main text"] ? data["Main text"] : "")
                  .map(m => parseTrailingParenSuffix(m).base)
                  .filter(Boolean);

                const expanded = splitSemicolonList(v).map(sectionAbbr => {
                  const { base: sectionBase, suffix: sectionSuffix } = parseTrailingParenSuffix(sectionAbbr);
                  let full = null;

                  if (mainCandidates.length === 1) {
                    full = minorMap[minorTextKey(mainCandidates[0], sectionBase)] || null;
                  } else if (mainCandidates.length > 1) {
                    const matches = [];
                    for (const m of mainCandidates) {
                      const hit = minorMap[minorTextKey(m, sectionBase)];
                      if (hit) matches.push(hit);
                    }
                    if (matches.length === 1) full = matches[0];
                  }

                  if (full) {
                    const expandedSuffix = sectionSuffix ? expandParenSuffixWithAbbreviations(sectionSuffix) : "";
                    return `${sectionBase}${EN_DASH}${full}${expandedSuffix}`;
                  }
                  return formatAbbrExpansion(null, sectionAbbr);
                }).join('; ');

                value = expanded;
              }
            }
          }

          if (key === "Links to Database") {
            value = renderLinksToDatabaseHtml(value);
          } else {
            value = escapeHtml(String(value));
          }

          entries.push({ key, value });
        });

        const half = Math.ceil(entries.length / 2);
        const left = entries.slice(0, half);
        const right = entries.slice(half);

        function renderColumn(items) {
          return items.map(e => `
            <div class="mb-2 border-bottom pb-1">
              <div class="fw-bold small">${escapeHtml(getColumnTitle(e.key))}</div>
              <div class="text-break">${e.value || "&nbsp;"}</div>
            </div>
          `).join('');
        }

        contentDiv.innerHTML = `
          <div class="row">
            <div class="col-md-6">${renderColumn(left)}</div>
            <div class="col-md-6">${renderColumn(right)}</div>
          </div>
        `;

        const pdfBtn = document.getElementById('download-pdf-btn');
        if (pdfBtn) {
          const depAbbr = data["Depository_abbr"] || data["Depository"] || "";
          const shelf = data["Shelf mark"] || "";
          const safe = (s) => String(s || "")
            .replace(/[\\/:*?\"<>|]+/g, "-")
            .replace(/\s+/g, " ")
            .trim();
          const filename = `${safe(depAbbr)}_${safe(shelf)}.pdf`;
          pdfBtn.dataset.filename = filename || "manuscript-details.pdf";
        }

        const modalEl = document.getElementById('rowDetailsModal');
        if (modalEl) {
          const bs = window.bootstrap || (typeof bootstrap !== 'undefined' ? bootstrap : null);
          if (bs && bs.Modal) {
            const modal = bs.Modal.getOrCreateInstance(modalEl);
            modal.show();
          }
        }
      });
    }

    // Ensure Text View page size matches the selector immediately (especially for SimpleTable).
    if (table && typeof table.setPageSize === 'function') {
      if (paginationRaw === 'all') table.setPageSize(true);
      else table.setPageSize(parseInt(paginationRaw, 10));
    }

    const totalEl = document.getElementById('total-records');
    if (totalEl) totalEl.textContent = safeRows.length;

    // Reset facet selections (All) and apply filters/render.
    FACET_FIELDS.forEach(field => updateFacetAllCheckbox(field));
    applyFacetFilters();

    applyViewUI();
  }


function setupControls() {
  // Clear search/facets button logic
  const clearBtn = document.getElementById("clear-filters");
  if (clearBtn) {
    clearBtn.addEventListener("click", function() {
      // Clear search box
      const searchInput = document.getElementById("search");
      if (searchInput) {
        searchInput.value = "";
      }
      // Reset all facet checkboxes to 'All'
      FACET_FIELDS.forEach(field => {
        const facetDiv = document.getElementById(`facet-${field}`);
        if (!facetDiv) return;

        if (field === "Dating") {
          const minEl = facetDiv.querySelector('input[data-dating-range="min"]');
          const maxEl = facetDiv.querySelector('input[data-dating-range="max"]');
          if (minEl) minEl.value = "";
          if (maxEl) maxEl.value = "";
          return;
        }

        const allBox = facetDiv.querySelector('input[type=checkbox][value="__ALL__"]');
        if (allBox) {
          allBox.checked = true;
          // Uncheck all others
          facetDiv.querySelectorAll('input[type=checkbox][data-facet]:not([value="__ALL__"])').forEach(cb => {
            cb.checked = false;
          });
        }
      });
      // Re-apply filters
      applyFacetFilters();
    });
  }

  // Search button logic
  const searchBtn = document.getElementById("search-btn");
  if (searchBtn) {
    searchBtn.addEventListener("click", function() {
      applyFacetFilters();
    });
  }

  const searchInput = document.getElementById("search");
  const paginationSizeSelect = document.getElementById("pagination-size");
  const viewSelect = document.getElementById("view-select");
  const mergedSortSelect = document.getElementById("merged-sort");
  const toggleFacetsBtn = document.getElementById("toggle-facets");

  // Restore facet sidebar visibility and wire toggle.
  try {
    setFacetsHidden(localStorage.getItem(FACETS_HIDDEN_STORAGE_KEY) === '1');
  } catch (e) {
    setFacetsHidden(false);
  }
  if (toggleFacetsBtn) {
    toggleFacetsBtn.addEventListener('click', function () {
      const nextHidden = !document.body.classList.contains('facets-hidden');
      setFacetsHidden(nextHidden);
      try { applyViewUI(); } catch (e) {}
      try { if (table && typeof table.redraw === 'function') table.redraw(true); } catch (e) {}
    });
  }

  const mergedPrevBtn = document.getElementById('merged-prev-page');
  const mergedNextBtn = document.getElementById('merged-next-page');
  if (mergedPrevBtn) {
    mergedPrevBtn.addEventListener('click', function () {
      if (currentView !== 'merged') return;
      MERGED_PAGE = Math.max(1, MERGED_PAGE - 1);
      applyFacetFilters();
    });
  }
  if (mergedNextBtn) {
    mergedNextBtn.addEventListener('click', function () {
      if (currentView !== 'merged') return;
      MERGED_PAGE = MERGED_PAGE + 1;
      applyFacetFilters();
    });
  }

  // Restore merged sort mode (if control exists)
  try {
    const raw = localStorage.getItem(MERGED_SORT_STORAGE_KEY);
    if (raw) MERGED_SORT_MODE = raw;
  } catch (e) {
    // ignore
  }
  if (mergedSortSelect) {
    const m = (MERGED_SORT_MODE === "dating") ? "dating" : "shelfmark";
    mergedSortSelect.value = m;
    mergedSortSelect.addEventListener("change", function () {
      MERGED_SORT_MODE = (this.value === "dating") ? "dating" : "shelfmark";
      try { localStorage.setItem(MERGED_SORT_STORAGE_KEY, MERGED_SORT_MODE); } catch (e) {}
      applyTableSort(MERGED_SORT_MODE);
      if (currentView === "merged") applyFacetFilters();
    });
  }

  if (viewSelect) {
    viewSelect.addEventListener("change", function () {
      setView(this.value);
    });
  }

  // Restore view (if available), then initialize.
  if (viewSelect) {
    try {
      const saved = localStorage.getItem(VIEW_STORAGE_KEY);
      if (saved === "merged" || saved === "table") {
        viewSelect.value = saved;
      }
    } catch (e) {
      // ignore
    }
    setView(viewSelect.value);
  } else {
    setView(currentView);
  }

  // Load selected TSV by default
  // Always load the combined file
  loadDataFromRawExcelSources();

  // Global search across all fields
  searchInput.addEventListener("keyup", function () {
    applyFacetFilters();
  });

  // Handle pagination size change
  paginationSizeSelect.addEventListener("change", function () {
    const value = this.value;
    if (currentView === 'merged') {
      MERGED_PAGE = 1;
      applyFacetFilters();
      return;
    }
    if (!table) return;
    if (value === "all") {
      table.setPageSize(true); // Show all rows
    } else {
      table.setPageSize(parseInt(value, 10));
    }
  });

  // PDF Download logic (Custom DOM Parsing for clickable links)
  const pdfBtn = document.getElementById("download-pdf-btn");
  if (pdfBtn) {
    pdfBtn.addEventListener("click", function() {
      if (!window.jspdf) {
        alert("jsPDF library not loaded.");
        return;
      }
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF('p', 'pt', 'a4');
      const content = document.getElementById("row-details-content");
      if (!content) return;

      // Parse modal content: find all label/value pairs in both columns
      // Each column: .col-md-6 > div.mb-2.border-bottom.pb-1
      const columns = content.querySelectorAll('.col-md-6');
      let leftItems = [], rightItems = [];
      if (columns.length === 2) {
        leftItems = Array.from(columns[0].querySelectorAll('.mb-2.border-bottom.pb-1'));
        rightItems = Array.from(columns[1].querySelectorAll('.mb-2.border-bottom.pb-1'));
      }

      // Helper to extract label and value (HTML)
      function extractItem(div) {
        const labelDiv = div.querySelector('.fw-bold');
        const valueDiv = div.querySelector('.text-break');
        const label = labelDiv ? labelDiv.textContent.trim() : '';
        let value = valueDiv ? valueDiv.innerHTML.trim() : '';
        return { label, value };
      }

      // PDF layout
      const marginX = 40, marginY = 60, colGap = 30;
      const colWidth = (515 - colGap) / 2; // 515 = A4 width - 2*marginX
      let y = marginY + 20;
      const lineHeight = 18;
      const maxY = doc.internal.pageSize.getHeight() - 60;

      // Set font to Times for all text
      doc.setFont("helvetica");
      // Header
      doc.setFontSize(16);
      doc.setTextColor(40);
      doc.setFont("helvetica", "bold");
      doc.text("NordicLaw Manuscripts", marginX, marginY);

      // Render two columns with dynamic height calculation
      function renderColumn(items, xStart) {
        let yPos = y;
        for (const item of items) {
          const { label, value } = extractItem(item);
          if (!label && !value) continue;
          // Label: Times Bold, size 8, gray (#6c757d), ALL CAPS
          doc.setFont("times", "bold");
          doc.setFontSize(8);
          doc.setTextColor(108, 117, 125); // Bootstrap 5 text-secondary
          const labelCaps = label ? label.toUpperCase() : "";
          const labelDims = doc.getTextDimensions(labelCaps, { maxWidth: colWidth });
          doc.text(labelCaps, xStart, yPos, { maxWidth: colWidth });
          yPos += (labelDims.h || 12) + 4;
          // Value: Times Regular, size 10, black
          doc.setFont("times", "normal");
          doc.setFontSize(10);
          doc.setTextColor(33, 37, 41); // Bootstrap 5 text-dark
          const tempDiv = document.createElement('div');
          tempDiv.innerHTML = value;
          const anchors = tempDiv.querySelectorAll('a');
          if (anchors.length > 0) {
            tempDiv.childNodes.forEach(node => {
              if (node.nodeType === 3) { // text
                const txt = node.textContent;
                if (txt && txt.trim()) {
                  const dims = doc.getTextDimensions(txt, { maxWidth: colWidth });
                  doc.text(txt, xStart, yPos, { maxWidth: colWidth });
                  yPos += dims.h || lineHeight;
                }
              } else if (node.nodeType === 1 && node.tagName === 'A') {
                const linkText = node.textContent;
                const href = node.getAttribute('href');
                doc.setFont("helvetica", "normal");
                doc.setFontSize(11);
                doc.setTextColor(13, 110, 253); // Bootstrap 5 link color
                const dims = doc.getTextDimensions(linkText, { maxWidth: colWidth });
                doc.textWithLink(linkText, xStart, yPos, { url: href });
                doc.setTextColor(33, 37, 41); // Reset to text-dark
                yPos += dims.h || lineHeight;
              }
            });
          } else {
            // Plain text (strip HTML)
            const plain = tempDiv.textContent || '';
            const dims = doc.getTextDimensions(plain, { maxWidth: colWidth });
            doc.text(plain, xStart, yPos, { maxWidth: colWidth });
            yPos += dims.h || lineHeight;
          }
          yPos += 8;
          if (yPos > maxY) {
            doc.addPage();
            yPos = marginY + 20;
            // Header for new page
            doc.setFont("helvetica", "bold");
            doc.setFontSize(16);
            doc.setTextColor(40);
            doc.text("NordicLaw Manuscripts", marginX, marginY);
          }
        }
      }

      renderColumn(leftItems, marginX);
      renderColumn(rightItems, marginX + colWidth + colGap);

      // Footer (on all pages)
      const pageCount = doc.getNumberOfPages();
      const dateStr = new Date().toLocaleString();
      const pageHeight = doc.internal.pageSize.getHeight();
      for (let i = 1; i <= pageCount; i++) {
        doc.setPage(i);
        doc.setFont("helvetica");
        doc.setFontSize(10);
        doc.setTextColor(100);
        doc.text(`Accessed: ${dateStr}`, marginX, pageHeight - 20);
      }

      // Save
      const filename = this.dataset.filename || "manuscript-details.pdf";
      doc.save(filename);
    });
  }
}

// Check if Bootstrap is loaded
if (typeof bootstrap === 'undefined' && !window.bootstrap) {
  console.error("Bootstrap 5 is not loaded! Modal will not work.");
}

setupControls();
