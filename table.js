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
        const dataKeys = Object.keys(data).filter(k => k !== "DatingYear" && k !== "_id" && k !== "Century" && k !== "Main text group");

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
            
            // Auto-convert URLs to links if they aren't already HTML anchors
            if (displayValue && typeof displayValue === 'string') {
                if (item.key === "Links to Database" || /^https?:\/\//.test(displayValue)) {
                    if (!displayValue.trim().startsWith('<a ') && /^https?:\/\//.test(displayValue)) {
                        displayValue = `<a href="${displayValue}" target="_blank">${displayValue}</a>`;
                    }
                }
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
}

// Check if Bootstrap is loaded
if (typeof bootstrap === 'undefined' && !window.bootstrap) {
  console.error("Bootstrap 5 is not loaded! Modal will not work.");
}

setupControls();
