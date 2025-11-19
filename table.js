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
  if (MAIN_TEXT_MAP) return MAIN_TEXT_MAP;
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
    return map;
  } catch (e) {
    console.error('Failed to load texts.tsv:', e);
    MAIN_TEXT_MAP = {};
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
  if (!table) return;
  table.clearFilter(true);
  // Compose filter function
  table.setFilter(function(row) {
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
        // Parse Excel-style HYPERLINK formulas for Links to Database
        if (h === "Links to Database" && val.startsWith("=HYPERL")) {
          // =HYPERLÄNK("url";"label") or =HYPERLINK("url","label")
          let m = val.match(/=HYPERL[ÄA]NK\(["']([^"']+)["'];?["']([^"']+)["']\)/i);
          if (!m) m = val.match(/=HYPERLINK\(["']([^"']+)["'],["']([^"']+)["']\)/i);
          if (m) val = `<a href="${m[1]}" target="_blank">${m[2]}</a>`;
        }
        // Expand Depository abbreviation
        if (h === "Depository" && val in depositoryMap) {
          val = depositoryMap[val];
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
    const userColumnOrder = [
      "Depository","Shelf mark", "Language", "Name","Object","Size","Dating","Leaves/Pages","Main text","Minor text","Gatherings","Physical Size (mm)","Production Unit","Pricking","Material","Ruling","Columns","Lines","Script","Rubric","Scribe","Production","Style","Colours","Form of Initials","Size of Initials","Iconography","Place","Related Shelfmarks","Literature","Links to Database"
    ];

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
          if (!v) return "";
          if (v.startsWith('<a ')) return v;
          if (/^https?:\/\//.test(v)) return `<a href="${v}" target="_blank">${v}</a>`;
          return v;
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
      visible: true,
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
  const datasetSelect = document.getElementById("dataset");
  const searchInput = document.getElementById("search");
  const paginationSizeSelect = document.getElementById("pagination-size");

  // Load selected TSV by default
  // Always load the combined file
  loadDataTSV('data/NordicLaw_data.tsv');

  // Handle dropdown change
  // Remove dataset switching, always use combined file
  datasetSelect.addEventListener("change", () => {
    loadDataTSV('data/NordicLaw_data.tsv');
    searchInput.value = ""; // clear search when changing dataset
  });

  // Global search across all fields
  searchInput.addEventListener("keyup", function () {
    const query = this.value.toLowerCase();
    if (!table) return;
    table.setFilter(function (data) {
      // Apply global search on top of facet filters
      // First, check facet filters
      for (const field of FACET_FIELDS) {
        if (facetSelections[field] && facetSelections[field].length > 0) {
          if (!facetSelections[field].includes(data[field])) return false;
        }
      }
      // Then, global search
      return Object.values(data).some(val =>
        String(val).toLowerCase().includes(query)
      );
    });
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
}

setupControls();
