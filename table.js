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

async function loadData(fileName) {
  try {
    const response = await fetch(fileName);
    const manuscripts = await response.json();

    // Flatten nested structure
    const rows = [];
    manuscripts.forEach(m => {
      const shelf = (m["Shelf mark"] || "").trim();
      // Copy all top-level shelf fields (except Production Units) into base so each row carries shelf metadata
      const base = {};
      Object.keys(m).forEach(k => {
        if (k === 'Production Units') return;
        const v = m[k];
        base[k] = (typeof v === 'string') ? v.trim() : (v === undefined || v === null ? '' : v);
      });
      // Ensure Shelf mark is trimmed and present
      base["Shelf mark"] = shelf;

      (m["Production Units"] || []).forEach(pu => {
        const puId = pu["Production Unit"] || "";
        const material = pu["Material"] || "";

        (pu["Contents"] || []).forEach(content => {
          rows.push({
            ...base,
            "Production Unit": puId,
            "Material": material,
            "Leaves/Pages": content["Leaves/Pages"] || "",
            "Main text": content["Main text"] || "",
            "Minor text": content["Minor text"] || "",
            "Dating": content["Dating"] || "",
            "Scribe": content["Scribe"] || "",
            "Script": content["Script"] || "",
          });
        });
      });
    });

    // Build columns dynamically so all fields show as visible columns
    const fieldSet = new Set();
    rows.forEach(r => Object.keys(r).forEach(k => fieldSet.add(k)));

    // Preferred column order for familiarity
    const preferred = ["Depository", "Shelf mark", "Production Unit", "Leaves/Pages", "Main text", "Minor text", "Dating", "Scribe", "Script", "Material", "Object", "Size", "Literature", "Links to Database"];

    const columns = [];
    // Add preferred fields first (if present)
    preferred.forEach(f => {
      if (fieldSet.has(f)) {
        columns.push({ title: f, field: f, sorter: f === 'Shelf mark' ? 'natural' : 'string', headerFilter: 'input', visible: true, hozAlign: 'left' });
        fieldSet.delete(f);
      }
    });
    // Add remaining fields alphabetically
    Array.from(fieldSet).sort().forEach(f => {
      columns.push({ title: f, field: f, headerFilter: 'input', visible: true, hozAlign: 'left' });
    });

      // use the top-level expandGroupsGlobal helper instead of duplicating logic here

    // If table exists, update columns and data; otherwise create it
    if (table) {
      table.setColumns(columns);
        await table.replaceData(rows);
  // ensure groups reflect current filters/search after render
  setTimeout(() => expandGroupsGlobal(), 150);
    } else {
      table = new Tabulator("#table", {
        data: rows,
        layout: "fitColumns",
        pagination: true,
        paginationSize: 20,
        placeholder: "No data available",
        groupBy: ["Shelf mark", "Production Unit", "Leaves/Pages"],
        groupStartOpen: [true, false, false],
        groupHeader: [
          function(value, count, data) {
            const depository = data[0]?.Depository || "Unknown Depository";
            return `<strong>${depository}, ${value}</strong>`;
          },
          value => `Unit ${value}`,
          value => `${value}`,
        ],
        responsiveLayout: false, // prevent Tabulator from hiding columns responsively
        columns: columns,
      });

        // Ensure groups expand/restore after rendering and filtering
        if (table) {
          table.on("dataFiltered", function(filters, rows){
            // give Tabulator time to update groups then expand
            setTimeout(() => expandGroupsGlobal(), 150);
          });
          // also call after render completes to ensure groups exist
          table.on("renderComplete", function(){
            setTimeout(() => expandGroupsGlobal(), 150);
          });
          // initial expand call after creation
          setTimeout(() => expandGroupsGlobal(), 150);
        }
    }

  } catch (err) {
    console.error("Error loading file:", err);
    alert(`Could not load ${fileName}`);
  }
}

function setupControls() {
  const datasetSelect = document.getElementById("dataset");
  const searchInput = document.getElementById("search");

  // Load default selection
  loadData(datasetSelect.value);

  // Handle dropdown change
  datasetSelect.addEventListener("change", () => {
    loadData(datasetSelect.value);
    searchInput.value = ""; // clear search when changing dataset
  });

  // Global search across all fields
  searchInput.addEventListener("keyup", function () {
    const query = this.value.toLowerCase();
    if (!table) return;
    table.setFilter(function (data) {
      return Object.values(data).some(val =>
        String(val).toLowerCase().includes(query)
      );
    });
  });
}

setupControls();
