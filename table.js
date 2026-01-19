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

function isSafeHttpUrl(url) {
  return typeof url === 'string' && /^https?:\/\//i.test(url.trim());
}

function normalizeLinksToDatabase(value) {
  if (!value || typeof value !== 'string') return '';
  let v = value.trim();
  if (!v) return '';

  // 1) Excel-style hyperlink formulas
  // =HYPERLÄNK("url";"label") or =HYPERLINK("url","label")
  if (v.startsWith('=HYPERL')) {
    let m = v.match(/=HYPERL[ÄA]NK\(["']([^"']+)["'];?["']([^"']+)["']\)/i);
    if (!m) m = v.match(/=HYPERLINK\(["']([^"']+)["'],["']([^"']+)["']\)/i);
    if (m && isSafeHttpUrl(m[1])) {
      const href = m[1].trim();
      const label = m[2] ? m[2].trim() : href;
      return `<a href="${escapeHtml(href)}" target="_blank" rel="noopener noreferrer">${escapeHtml(label)}</a>`;
    }
  }

  // 2) Existing HTML: re-build anchors safely (avoid trusting arbitrary HTML from TSV)
  if (v.includes('<a')) {
    try {
      const tmp = document.createElement('div');
      tmp.innerHTML = v;
      const anchors = Array.from(tmp.querySelectorAll('a'));
      if (anchors.length > 0) {
        const rendered = anchors
          .map(a => {
            const href = (a.getAttribute('href') || '').trim();
            const label = (a.textContent || href).trim();
            if (!isSafeHttpUrl(href)) return '';
            return `<a href="${escapeHtml(href)}" target="_blank" rel="noopener noreferrer">${escapeHtml(label)}</a>`;
          })
          .filter(Boolean)
          .join('; ');
        if (rendered) return rendered;
      }
    } catch (e) {
      // Fall through to other parsing strategies
    }
  }

  // 3) Markdown links: [label](url) possibly multiple
  const mdLinks = [];
  const mdRe = /\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)/g;
  let mdMatch;
  while ((mdMatch = mdRe.exec(v)) !== null) {
    const label = (mdMatch[1] || '').trim();
    const href = (mdMatch[2] || '').trim();
    if (!isSafeHttpUrl(href)) continue;
    mdLinks.push({ label: label || href, href });
  }
  if (mdLinks.length > 0) {
    return mdLinks
      .map(l => `<a href="${escapeHtml(l.href)}" target="_blank" rel="noopener noreferrer">${escapeHtml(l.label)}</a>`)
      .join('; ');
  }

  // 4) Plain URLs (single or multiple separated by semicolons/spaces)
  const urlRe = /(https?:\/\/[^\s;]+)/g;
  const urls = (v.match(urlRe) || []).map(s => s.trim()).filter(isSafeHttpUrl);
  if (urls.length > 0) {
    return urls
      .map(href => `<a href="${escapeHtml(href)}" target="_blank" rel="noopener noreferrer">${escapeHtml(href)}</a>`)
      .join('; ');
  }

  // 5) Not a link: show as escaped text
  return escapeHtml(v);
}
let table; // Keep reference to reuse same table instance

// Expand/close groups based on active filters/search; uses current `table` variable
function expandGroupsGlobal() {
  if (!table) return;
  const searchInput = document.getElementById('search');
  const hasGlobalSearch = searchInput && searchInput.value.trim() !== '';
  const activeFilters = (typeof table.getFilters === 'function') ? table.getFilters() : [];
  const hasColumnFilters = activeFilters && activeFilters.length > 0;
  const hasActiveFilter = hasGlobalSearch || hasColumnFilters;

  function groupHasRows(group) {
    try {
      if (group.getRows && group.getRows().length > 0) return true;
      if (group.getSubGroups) {
        const subs = group.getSubGroups();
        for (let i = 0; i < subs.length; i++) {
          if (groupHasRows(subs[i])) return true;
        }
      }
    } catch (e) {
      return false;
    }
    return false;
  }

  let topGroups = [];
  try { topGroups = table.getGroups() || []; } catch (e) { topGroups = []; }
  if (!topGroups || topGroups.length === 0) return;

  if (hasActiveFilter) {
    // Open all groups when a filter/search is active so visible rows are revealed
    (function walk(groups){
      groups.forEach(g => {
        try { g.open(); } catch(e){}
        if (g.getSubGroups) walk(g.getSubGroups());
      });
    })(topGroups);
  } else {
    // restore defaults: open top-level, close deeper
    topGroups.forEach(g => {
      try { g.open(); } catch(e){}
      if (g.getSubGroups) {
        g.getSubGroups().forEach(sg => { try { sg.close(); } catch(e){} });
      }
    });
  }
}

// Natural sorting for Shelf mark
Tabulator.extendModule("sort", "sorters", {
  natural: (a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }),
});


// --- Facet filter logic ---

const FACET_FIELDS = [
  "Language",
  "Depository",
  "Object",
  "Script",
  "Material",
  "Size",
  "Pricking",
  "Ruling",
  "Columns",
  "Lines",
  "Rubric",
  "Style",
  "Main text group",
  "Dating",
];
let facetSelections = {};
let allRows = [];

// Header order from the filled combined TSV (used to match the Excel/TSV column layout in merged view)
let DATA_HEADERS = null;

// Final column order actually used by both the Tabulator table and the merged (Excel-like) view.
// Edit COLUMN_ORDER to change the UI column order.
let DISPLAY_COLUMNS = null;

// Manuscript (merged) view: visible columns (persisted in localStorage)
const MERGED_COLUMNS_STORAGE_KEY = "nordiclaw.mergedVisibleColumns";
let MERGED_VISIBLE_COLUMNS = null; // Set<string> | null (null => all visible)

// User-specified column order (also used by the merged view)
// NOTE: This is the single place to change column ordering in the UI.
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

function buildDisplayColumns(headers) {
  const hs = Array.isArray(headers) ? headers : [];
  const headerSet = new Set(hs);

  // Start with the user order, but only keep columns that actually exist in the data.
  const ordered = COLUMN_ORDER.filter(h => headerSet.has(h));

  // Append any remaining columns (stable, in-file order), excluding internal fields.
  for (const h of hs) {
    if (h === "Century") continue;
    if (!ordered.includes(h)) ordered.push(h);
  }

  return ordered;
}

function getMergedVisibleColumnsSet() {
  if (MERGED_VISIBLE_COLUMNS instanceof Set) return MERGED_VISIBLE_COLUMNS;
  return null;
}

function isMergedColumnVisible(col) {
  const set = getMergedVisibleColumnsSet();
  return !set || set.has(col);
}

function loadMergedColumnVisibility() {
  try {
    const raw = localStorage.getItem(MERGED_COLUMNS_STORAGE_KEY);
    if (!raw) {
      MERGED_VISIBLE_COLUMNS = null;
      return;
    }
    const arr = JSON.parse(raw);
    if (!Array.isArray(arr)) {
      MERGED_VISIBLE_COLUMNS = null;
      return;
    }
    MERGED_VISIBLE_COLUMNS = new Set(arr.map(String));
  } catch (e) {
    MERGED_VISIBLE_COLUMNS = null;
  }
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

function getMergedVisibleColumnsArray(columnsFull) {
  const cols = Array.isArray(columnsFull) ? columnsFull : [];
  return cols.filter(isMergedColumnVisible);
}

function renderMergedColumnsMenu() {
  const menu = document.getElementById("merged-columns-menu");
  if (!menu) return;
  if (!DISPLAY_COLUMNS || !Array.isArray(DISPLAY_COLUMNS) || DISPLAY_COLUMNS.length === 0) {
    menu.innerHTML = '<div class="text-secondary small">Columns not ready yet.</div>';
    return;
  }

  const minimal = [
    "Depository",
    "Shelf mark",
    "Name",
    "Language",
    "Production Unit",
    "Dating",
    "Material",
    "Links to Database",
  ];

  const textualContent = [
    "Depository",
    "Shelf mark",
    "Language",
    "Production Unit",
    "Leaves/Pages",
    "Main text",
    "Minor text",
    "Dating"
  ];

  const codicology = [
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
  ];

  const layoutAndDecoration = [
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
  ];

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
    html += `
      <div class="form-check">
        <input class="form-check-input" type="checkbox" data-cols-col="${escapeHtml(col)}" id="colvis-${escapeHtml(col)}" ${isChecked(col) ? 'checked' : ''}>
        <label class="form-check-label" for="colvis-${escapeHtml(col)}">${escapeHtml(col)}</label>
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
        MERGED_VISIBLE_COLUMNS = null;
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

// View mode: default Tabulator table, optional merged-cells (Excel-like) view
let currentView = "table";

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

function renderMergedCell(field, value) {
  const v = normalizeForCompare(value);
  if (!v) return "&nbsp;";
  if (field === "Links to Database") return normalizeLinksToDatabase(v);
  return escapeHtml(v);
}

function renderMergedView(manuscripts) {
  const mergedRoot = document.getElementById("merged-table");
  const meta = document.getElementById("merged-view-meta");
  if (!mergedRoot) return;

  const columnsFull = (DISPLAY_COLUMNS && Array.isArray(DISPLAY_COLUMNS))
    ? DISPLAY_COLUMNS.slice()
    : ((DATA_HEADERS && Array.isArray(DATA_HEADERS)) ? DATA_HEADERS.slice() : COLUMN_ORDER.slice());
  const columns = getMergedVisibleColumnsArray(columnsFull);
  let totalRows = 0;
  manuscripts.forEach(m => { totalRows += m.rows.length; });

  if (meta) {
    let note = "";
    if (!RAW_EXCEL_LOADED && !RAW_EXCEL_FAILED) note = " (loading raw Excel merges…)";
    if (RAW_EXCEL_FAILED) note = " (raw Excel files not found; showing fallback view)";
    meta.textContent = `${manuscripts.length} manuscripts, ${totalRows} rows${note}`;
  }
  const totalEl = document.getElementById('total-records');
  if (totalEl) totalEl.textContent = `${manuscripts.length} manuscripts (${totalRows} texts)`;

  let html = '<table class="table table-bordered merged-table"><thead><tr>';
  for (const col of columns) {
    html += `<th>${escapeHtml(col)}</th>`;
  }
  html += '</tr></thead><tbody>';

  manuscripts.forEach((ms, msIndex) => {
    const msClass = (msIndex % 2 === 0) ? 'ms-a' : 'ms-b';
    // Prefer raw Excel-like rows (blank cells preserved) if available
    const rawBlock = RAW_BY_MANUSCRIPT_KEY.get(ms.key);
    const msRows = (rawBlock && rawBlock.rows && rawBlock.rows.length > 0) ? rawBlock.rows : ms.rows;
    const msRowCount = msRows.length;
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

        if (mergeLookup) {
          // The raw Excel files don't have a Language column; we inject it.
          // Merge it per manuscript so duplicates don't repeat visually.
          if (col === "Language") {
            if (!isFirstMsRow) continue;
            html += `<td class="${msClass}" rowspan="${msRowCount}">${renderMergedCell(col, row[col])}</td>`;
            continue;
          }

          const key = `${rIndex},${c}`;
          if (mergeLookup.covered.has(key)) continue;
          const span = mergeLookup.topLeft.get(key);
          const attrs = span ? ` rowspan="${span.rowSpan}" colspan="${span.colSpan}"` : "";
          html += `<td class="${(col === "Production Unit") ? puClass : msClass}"${attrs}>${renderMergedCell(col, row[col])}</td>`;
          continue;
        }

        if (msConst && msConst.has(col)) {
          if (!isFirstMsRow) continue;
          html += `<td class="${msClass}" rowspan="${msRowCount}">${renderMergedCell(col, row[col])}</td>`;
          continue;
        }

        if (col === "Production Unit") {
          if (!isRunStart) continue;
          html += `<td class="${puClass}" rowspan="${runRowSpan}">${renderMergedCell(col, row[col])}</td>`;
          continue;
        }

        if (runConst && runConst.has(col)) {
          html += `<td class="${puClass}" rowspan="${runRowSpan}">${renderMergedCell(col, row[col])}</td>`;
          continue;
        }

        html += `<td>${renderMergedCell(col, row[col])}</td>`;
      }

      html += '</tr>';
    }
  });

  html += '</tbody></table>';
  mergedRoot.innerHTML = html;
}

function setView(view) {
  currentView = (view === "merged") ? "merged" : "table";
  const tableView = document.getElementById("table-view");
  const mergedView = document.getElementById("merged-view");
  if (tableView) tableView.style.display = (currentView === "table") ? "block" : "none";
  if (mergedView) mergedView.style.display = (currentView === "merged") ? "block" : "none";

  // Column selector lives in the top control bar; only show it for Manuscript View.
  const mergedColsControl = document.getElementById("merged-columns-control");
  if (mergedColsControl) mergedColsControl.style.display = (currentView === "merged") ? "" : "none";

  if (currentView === "table" && table && typeof table.redraw === 'function') {
    try { table.redraw(true); } catch (e) {}
  }

  const paginationSizeSelect = document.getElementById("pagination-size");
  if (paginationSizeSelect) paginationSizeSelect.disabled = (currentView === "merged");

  // Keep merged column menu in sync
  if (currentView === "merged") {
    renderMergedColumnsMenu();
  }

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

// Main text abbreviation mapping, loaded from texts.tsv at runtime
let MAIN_TEXT_MAP = null;
async function loadMainTextMap() {
  if (MAIN_TEXT_MAP) {
    // Always assign to window for modal rendering
    window.MAIN_TEXT_MAP = MAIN_TEXT_MAP;
    return MAIN_TEXT_MAP;
  }
  try {
    const resp = await fetch('data/texts.tsv');
    const text = await resp.text();
    const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
    const map = {};
    lines.forEach(line => {
      const [abbr, full] = line.split(/\t/);
      if (abbr && full) map[abbr.trim()] = full.trim();
    });
    MAIN_TEXT_MAP = map;
    window.MAIN_TEXT_MAP = map;
    return map;
  } catch (e) {
    console.error('Failed to load texts.tsv:', e);
    MAIN_TEXT_MAP = {};
    window.MAIN_TEXT_MAP = {};
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
        const groupLabel = mainTextMap[group] ? `${escapeHtml(group)} — ${escapeHtml(mainTextMap[group])}` : escapeHtml(group);
        html += `<div style='margin-left:0.5em;'><div class="form-check mb-1"><input class="form-check-input" type="checkbox" value="${escapeHtml(group)}" data-facet="Main text group" id="${groupId}"><label class="form-check-label" for="${groupId}">${groupLabel}</label></div>`;
        const variants = Array.from(groupMap[group]);
        if (variants.length > 0) {
          html += `<div style='margin-left:1.5em;'>`;
          variants.sort().forEach((variant, vi) => {
            // Skip empty or dot-only variant values
            if (variant === '' || variant.trim() === '.') return;
            const variantId = `facet-mtg-variant-${gi}-${vi}`;
            const variantLabel = mainTextMap[variant] ? `${escapeHtml(variant)} — ${escapeHtml(mainTextMap[variant])}` : escapeHtml(variant);
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
        label = `${val} — ${mainTextMap[val]}`;
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
    const manuscripts = groupByPreserveOrder(allRows || [], getManuscriptKey)
      .map(g => ({ key: g.key, rows: g.rows }))
      .sort((a, b) => {
        const a0 = a.rows[0] || {};
        const b0 = b.rows[0] || {};
        const depCmp = String(a0["Depository"] || "").localeCompare(String(b0["Depository"] || ""), undefined, { sensitivity: "base" });
        if (depCmp !== 0) return depCmp;
        return String(a0["Shelf mark"] || "").localeCompare(String(b0["Shelf mark"] || ""), undefined, { numeric: true, sensitivity: "base" });
      });

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
    renderMergedView(filtered);

    // Keep Tabulator internally in sync (even though hidden)
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
    const depositoryMap = await loadDepositoryMap();
    const response = await fetch(fileName);
    const tsvText = await response.text();
    const lines = tsvText.split(/\r?\n/).filter(line => line.trim() !== "");
    if (lines.length === 0) throw new Error("TSV file is empty");
    const headers = lines[0].split("\t");

    // Preserve the actual column order from the combined TSV for the merged view.
    DATA_HEADERS = headers.slice();
    DISPLAY_COLUMNS = buildDisplayColumns(headers);

    // Load/validate column visibility selections now that we know which columns exist.
    loadMergedColumnVisibility();
    sanitizeMergedColumnVisibility();
    renderMergedColumnsMenu();

    let rows = lines.slice(1).map(line => {
      const values = line.split("\t");
      const obj = {};
      headers.forEach((h, i) => {
        let val = values[i] !== undefined ? values[i] : "";
        if (h === "Main text") val = val.trim();
        // Expand Depository abbreviation
        if (h === "Depository") {
          obj["Depository_abbr"] = val; // Store original abbreviation for filename generation
          if (val in depositoryMap) {
            val = depositoryMap[val];
          }
        }
        obj[h] = val;
      });
      // Add Main text group (abbreviation before parentheses)
      if (obj["Main text"]) {
        let m = obj["Main text"].match(/^([^\(]+)\s*\(/);
        obj["Main text group"] = m ? m[1].trim() : obj["Main text"];
      } else {
        obj["Main text group"] = "";
      }
      // Add parsed Century
      obj["Century"] = parseCentury(obj["Dating"]);
      // Add parsed DatingYear (improved)
      obj["DatingYear"] = parseDatingYear(obj["Dating"]);
      // Normalize language column for facet
      normalizeRowLanguage(obj);
      return obj;
    });
    allRows = rows;

  // Build columns from headers, ordered by DISPLAY_COLUMNS/COLUMN_ORDER
    let colDefs = {};
    headers.forEach(h => {
      let colDef = {
        title: h,
        field: h,
        visible: true,
        hozAlign: 'left',
        width: ["Depository", "Shelf mark", "Production Unit"].includes(h) ? 200 : 100,
        headerFilter: "input",
        headerFilterPlaceholder: "Filter...",
        headerFilterLiveFilter: true
      };
      if (h === "Links to Database") {
        colDef.formatter = function(cell) {
          const v = cell.getValue();
          return normalizeLinksToDatabase(v);
        };
        colDef.formatterParams = { allowHtml: true };
      }
      colDefs[h] = colDef;
    });

    // Compose columns array in the UI display order (COLUMN_ORDER + any extras)
    let columns = [];
    (DISPLAY_COLUMNS || []).forEach(h => { if (colDefs[h]) columns.push(colDefs[h]); });
    // Add Century column at the end if present
    columns.push({
      title: "Century",
      field: "Century",
      visible: false, // Hide the Century column
      hozAlign: 'left',
      width: 90,
      headerFilter: "input",
      headerFilterPlaceholder: "Filter...",
      headerFilterLiveFilter: true
    });

    // Render facet sidebar and set up events

  // renderFacetSidebar is now async
  await renderFacetSidebar(rows);
  setupFacetEvents();

    // If table exists, update columns and data; otherwise create it
    if (table) {
      table.setColumns(columns);
      await table.replaceData(rows);
      document.getElementById('total-records').textContent = rows.length;
    } else {
      table = new Tabulator("#table-view", {
        data: rows,
        layout: "fitColumns",
        pagination: true,
        paginationSize: 50,
        placeholder: "No data available",
        headerVisible: true,
        headerFilterPlaceholder: "Filter...",
        groupBy: false,
        movableRows: false,
        responsiveLayout: false,
        columns: columns,
        movableColumns: true,
        autoResize: false,
        height: "100%",
        initialSort: [
          { column: "Depository", dir: "asc" },
          { column: "Shelf mark", dir: "asc" }
        ],
        theme: "bootstrap5",
        selectable: 1, // Allow row selection for visual feedback
      });

      // Expose for index.html sidebar-resizer redraw hook
      window.table = table;

      // Attach rowClick event handler
      table.on("rowClick", function(e, row){
        console.log("Row clicked", row.getData());
        const data = row.getData();
        const contentDiv = document.getElementById('row-details-content');
        if (!contentDiv) {
          console.error("row-details-content element not found!");
          return;
        }
        
        // Collect and sort entries based on the same column order used in the UI
        const entries = [];
        // Exclude internal fields and fields to hide in modal
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
          entries.push({ key: key, value: data[key] });
        });

        // Split into two columns
        const mid = Math.ceil(entries.length / 2);
        const left = entries.slice(0, mid);
        const right = entries.slice(mid);

        function renderItem(item) {
            let displayValue = item.value;

            // Expand Main text using MAIN_TEXT_MAP if available
            if (item.key === "Main text" && typeof displayValue === 'string' && window.MAIN_TEXT_MAP) {
              const expanded = window.MAIN_TEXT_MAP[displayValue];
              if (expanded) {
                displayValue = `${displayValue} — <span class='text-secondary'>${expanded}</span>`;
              }
            }

            // Normalize known link field(s)
            if (item.key === "Links to Database") {
              displayValue = normalizeLinksToDatabase(displayValue);
            } else if (displayValue && typeof displayValue === 'string' && /^https?:\/\//.test(displayValue)) {
              // Auto-convert plain URLs for other fields
              displayValue = `<a href="${escapeHtml(displayValue)}" target="_blank" rel="noopener noreferrer">${escapeHtml(displayValue)}</a>`;
            }

            if (!displayValue) displayValue = '&nbsp;';

            // Compact layout: Label on top (small), value below
            return `<div class="mb-2 border-bottom pb-1">
                  <div class="fw-bold text-secondary" style="font-size: 0.75rem; text-transform: uppercase;">${item.key}</div>
                  <div class="text-break" style="font-size: 0.9rem;">${displayValue}</div>
                </div>`;
        }

        let html = '<div class="container-fluid"><div class="row">';
        
        html += '<div class="col-md-6">';
        left.forEach(item => { html += renderItem(item); });
        html += '</div>';

        html += '<div class="col-md-6">';
        right.forEach(item => { html += renderItem(item); });
        html += '</div>';

        html += '</div></div>';
        contentDiv.innerHTML = html;
        
        // Generate filename for PDF download
        // Format: Depository (abbr) + Shelf mark + Leaves/Pages
        // Replace non-char/non-digit with _
        const depAbbr = data["Depository_abbr"] || data["Depository"] || "Unknown";
        const shelfMark = data["Shelf mark"] || "";
        const leaves = data["Leaves/Pages"] || "";
        
        let rawFilename = `${depAbbr}_${shelfMark}_${leaves}`;
        // Sanitize: replace non-alphanumeric characters (except underscores) with _
        // Actually user asked: "Replace any non character or non digit contents with _"
        let filename = rawFilename.replace(/[^a-zA-Z0-9]/g, "_");
        // Remove duplicate underscores and trim
        filename = filename.replace(/_+/g, "_").replace(/^_|_$/g, "");
        if (!filename) filename = "manuscript_details";
        
        const pdfBtn = document.getElementById("download-pdf-btn");
        if (pdfBtn) {
            pdfBtn.dataset.filename = filename + ".pdf";
        }

        const modalEl = document.getElementById('rowDetailsModal');
        if (!modalEl) {
          console.error("rowDetailsModal element not found!");
          return;
        }

        // Try to find bootstrap
        const bs = window.bootstrap || (typeof bootstrap !== 'undefined' ? bootstrap : null);
        
        if (bs && bs.Modal) {
          try {
            const modal = bs.Modal.getOrCreateInstance(modalEl);
            modal.show();
          } catch (err) {
            console.error("Error showing modal:", err);
            alert("Error showing modal. See console for details.");
          }
        } else {
          console.error("Bootstrap Modal not available.");
          alert("Bootstrap not loaded. Row details:\n" + JSON.stringify(data, null, 2));
        }
      });

      document.getElementById('total-records').textContent = rows.length;
      if (table) {
        table.on("dataFiltered", function(filters, filteredRows){
          if (currentView === "table") {
            document.getElementById('total-records').textContent = filteredRows.length;
          }
        });
      }
    }
    // Reset facet selections (all checked)
    FACET_FIELDS.forEach(field => updateFacetAllCheckbox(field));
    applyFacetFilters();
  } catch (err) {
    console.error("Error loading TSV file:", err);
    alert(`Could not load ${fileName}`);
  }
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

  if (viewSelect) {
    viewSelect.addEventListener("change", function () {
      setView(this.value);
    });
  }

  // Load selected TSV by default
  // Always load the combined file
  loadDataTSV('data/NordicLaw_data.tsv');

  // Global search across all fields
  searchInput.addEventListener("keyup", function () {
    applyFacetFilters();
  });

  // Handle pagination size change
  paginationSizeSelect.addEventListener("change", function () {
    const value = this.value;
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
