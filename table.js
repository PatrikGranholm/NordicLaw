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

const FACET_FIELDS = ["Language", "Depository", "Object", "Script", "Material", "Main text group", "Dating", "Century"];
let facetSelections = {};
let allRows = [];

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

function getUniqueValues(rows, field) {
  const set = new Set();
  rows.forEach(row => {
    let val = row[field];
    if (!val || val.trim() === "") return;
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
        const groupId = `facet-mtg-group-${gi}`;
        const groupLabel = mainTextMap[group] ? `${group} — ${mainTextMap[group]}` : group;
        html += `<div style='margin-left:0.5em;'><div class="form-check mb-1"><input class="form-check-input" type="checkbox" value="${group.replace(/"/g, '&quot;')}" data-facet="Main text group" id="${groupId}"><label class="form-check-label" for="${groupId}">${groupLabel}</label></div>`;
        const variants = Array.from(groupMap[group]);
        if (variants.length > 0) {
          html += `<div style='margin-left:1.5em;'>`;
          variants.sort().forEach((variant, vi) => {
            const variantId = `facet-mtg-variant-${gi}-${vi}`;
            const variantLabel = mainTextMap[variant] ? `${variant} — ${mainTextMap[variant]}` : variant;
            html += `<div class="form-check mb-1"><input class="form-check-input" type="checkbox" value="${group}|${variant}" data-facet="Main text group-variant" id="${variantId}"><label class="form-check-label" for="${variantId}">${variantLabel}</label></div>`;
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
      let label = val;
      if (field === "Main text" && mainTextMap[val]) {
        label = `${val} — ${mainTextMap[val]}`;
      }
      const id = `facet-${field}-${i}`;
      html += `<div class="form-check mb-1"><input class="form-check-input" type="checkbox" value="${val.replace(/"/g, '&quot;')}" data-facet="${field}" id="${id}"><label class="form-check-label" for="${id}">${label}</label></div>`;
    });
    facetDiv.innerHTML = html;
  });
}

function getFacetSelections() {
  const selections = {};
    FACET_FIELDS.forEach(field => {
      const facetDiv = document.getElementById(`facet-${field}`);
      if (!facetDiv) return;
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
  const otherBoxes = facetDiv.querySelectorAll('input[type=checkbox][data-facet]:not([value="__ALL__"])');
  if (!allBox) return;
  const allChecked = Array.from(otherBoxes).every(cb => cb.checked);
  allBox.checked = allChecked || Array.from(otherBoxes).every(cb => !cb.checked);
}

function setupFacetEvents() {
  FACET_FIELDS.forEach(field => {
    const facetDiv = document.getElementById(`facet-${field}`);
    if (!facetDiv) return;
    facetDiv.addEventListener('change', (e) => {
      if (!e.target.matches('input[type=checkbox][data-facet]')) return;
      if (e.target.value === "__ALL__") {
        // All box toggled: check/uncheck all
        const allChecked = e.target.checked;
        facetDiv.querySelectorAll('input[type=checkbox][data-facet]:not([value="__ALL__"])').forEach(cb => {
          cb.checked = allChecked;
        });
      } else {
        // If all boxes checked, check All; if none checked, check All
        updateFacetAllCheckbox(field);
      }
      applyFacetFilters();
    });
  });
}

function applyFacetFilters() {
  facetSelections = getFacetSelections();
  const searchInput = document.getElementById('search');
  const query = searchInput ? searchInput.value.toLowerCase() : "";

  if (!table) return;
  table.clearFilter(true);
  // Compose filter function
  table.setFilter(function(row) {
    // 1. Check Facets
    for (const field of FACET_FIELDS) {
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
          if (!facetSelections[field].includes(row[field])) return false;
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

    let rows = lines.slice(1).map(line => {
      const values = line.split("\t");
      const obj = {};
      headers.forEach((h, i) => {
        let val = values[i] !== undefined ? values[i] : "";
        if (h === "Main text") val = val.trim();
        // Normalize link formats (Markdown/URL/HTML/Excel) to safe HTML anchors
        if (h === "Links to Database") {
          val = normalizeLinksToDatabase(val);
        }
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



    // User-specified column order
    const userColumnOrder = ["Depository","Shelf mark","Language","Name","Object","Size","Production Unit","Leaves/Pages","Main text","Minor text","Dating","Gatherings","Full size","Leaf size","Catch Words and Gatherings","Pricking","Material","Ruling","Columns","Lines","Script","Rubric","Scribe","Production","Style","Colours","Form of Initials","Size of Initials","Iconography","Place","Related Shelfmarks","Literature","Links to Database"];
    
    // Build columns from headers, but order by userColumnOrder, then any extra columns (e.g. Century)
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

    // Compose columns array in user order, then add any extra columns (except Century)
    let columns = [];
    userColumnOrder.forEach(h => { if (colDefs[h]) columns.push(colDefs[h]); });
    // Add any columns from headers not in userColumnOrder (except Century)
    headers.forEach(h => { if (!userColumnOrder.includes(h) && h !== "Century" && colDefs[h]) columns.push(colDefs[h]); });
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
      table = new Tabulator("#table", {
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

      // Attach rowClick event handler
      table.on("rowClick", function(e, row){
        console.log("Row clicked", row.getData());
        const data = row.getData();
        const contentDiv = document.getElementById('row-details-content');
        if (!contentDiv) {
          console.error("row-details-content element not found!");
          return;
        }
        
        // Collect and sort entries based on userColumnOrder
        const entries = [];
        // Exclude internal fields and fields to hide in modal
        const dataKeys = Object.keys(data).filter(k => k !== "DatingYear" && k !== "_id" && k !== "Century" && k !== "Main text group" && k !== "Depository_abbr");

        dataKeys.sort((a, b) => {
          const idxA = userColumnOrder.indexOf(a);
          const idxB = userColumnOrder.indexOf(b);
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
          document.getElementById('total-records').textContent = filteredRows.length;
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
