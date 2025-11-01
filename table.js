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

const FACET_FIELDS = ["Depository", "Object", "Script", "Material", "Main text group", "Dating", "Century"];
let facetSelections = {};
let allRows = [];

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

function renderFacetSidebar(rows) {
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
      let html = `<strong>Main text group</strong><br>`;
      html += `<label><input type="checkbox" value="__ALL__" checked data-facet="Main text group">All</label>`;
      Object.keys(groupMap).sort().forEach(group => {
        html += `<div style='margin-left:0.5em;'><label><input type="checkbox" value="${group.replace(/"/g, '&quot;')}" data-facet="Main text group">${group}</label>`;
        const variants = Array.from(groupMap[group]);
        if (variants.length > 0) {
          html += `<div style='margin-left:1.5em;'>`;
          variants.sort().forEach(variant => {
            html += `<label><input type="checkbox" value="${group}|${variant}" data-facet="Main text group-variant">${variant}</label>`;
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
    let html = `<strong>${field}</strong><br>`;
    html += `<label><input type="checkbox" value="__ALL__" checked data-facet="${field}">All</label>`;
    values.forEach(val => {
      html += `<label><input type="checkbox" value="${val.replace(/"/g, '&quot;')}" data-facet="${field}">${val}</label>`;
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

async function loadDataTSV(fileName) {
  try {
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
      return obj;
    });
    allRows = rows;


    // Build columns dynamically from headers, add Century column at the end
    let columns = headers.map(h => {
      let colDef = {
        title: h,
        field: h,
        visible: true,
        hozAlign: 'left',
        width: ["Depository", "Shelf mark", "Production Unit"].includes(h) ? 200 : 100
      };
      if (h === "Columns" || h === "Depository") {
        colDef.headerFilter = "select";
        colDef.headerFilterParams = { values: true, clearable: true, multiselect: false, sort: "asc" };
        colDef.headerFilterPlaceholder = "All";
        colDef.headerFilterLiveFilter = true;
      } else {
        colDef.headerFilter = false;
      }
      return colDef;
    });
    // Add Century column
    columns.push({
      title: "Century",
      field: "Century",
      visible: true,
      hozAlign: 'left',
      width: 90,
      headerFilter: "select",
      headerFilterParams: { values: true, clearable: true, multiselect: false, sort: "asc" },
      headerFilterPlaceholder: "All",
      headerFilterLiveFilter: true
    });

    // Render facet sidebar and set up events
  renderFacetSidebar(rows);
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
  loadDataTSV(datasetSelect.value);

  // Handle dropdown change
  datasetSelect.addEventListener("change", () => {
    loadDataTSV(datasetSelect.value);
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
