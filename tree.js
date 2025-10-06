async function loadTree() {
  const response = await fetch('metadata.json');
  const manuscripts = await response.json();

  // Preprocess and trim Shelf mark
  const cleanedData = manuscripts.map(m => ({
    ...m,
    "Shelf mark": (m["Shelf mark"] || "").trim()
  }));

  // Build jsTree data structure
  const treeData = cleanedData.map((m, mIndex) => ({
    id: `m-${mIndex}`,
    text: `${m["Shelf mark"]} (${m["Object"] || "Unknown"})`,
    state: { opened: false },
    data: m,
    children: (m["Production Units"] || []).map((pu, puIndex) => ({
      id: `m-${mIndex}-pu-${puIndex}`,
      text: `Production Unit ${pu["Production Unit"] || ""} (${pu["Material"] || "Unknown"})`,
      data: pu,
      children: (pu["Contents"] || []).map((c, cIndex) => {
        const leaves = c["Leaves/Pages"] || "";
        const mainText = c["Main text"] || c["Minor text"] || "Untitled";
        const dating = c["Dating"] || "";
        const labelParts = [leaves, mainText, dating].filter(Boolean);
        const label = labelParts.join(" — ");

        return {
          id: `m-${mIndex}-pu-${puIndex}-c-${cIndex}`,
          text: label,
          icon: "jstree-file",
          data: c
        };
      })
    }))
  }));

  // Initialize jsTree
  $('#tree').jstree({
    core: {
      data: treeData,
      themes: { stripes: true }
    },
    plugins: ["search", "sort"],
    sort: function (a, b) {
      const nodeA = this.get_node(a);
      const nodeB = this.get_node(b);
      const textA = (nodeA.text || "").trim();
      const textB = (nodeB.text || "").trim();
      return naturalCompare(textA, textB);
    }
  });

  // Natural sort function (case-insensitive + numeric)
  function naturalCompare(a, b) {
    a = a.toString().toLowerCase();
    b = b.toString().toLowerCase();

    const ax = [];
    const bx = [];

    a.replace(/(\d+)|(\D+)/g, (_, $1, $2) => ax.push([$1 || Infinity, $2 || ""]));
    b.replace(/(\d+)|(\D+)/g, (_, $1, $2) => bx.push([$1 || Infinity, $2 || ""]));

    while (ax.length && bx.length) {
      const an = ax.shift();
      const bn = bx.shift();
      const [a1, a2] = an;
      const [b1, b2] = bn;
      if (a2 !== b2) return a2 > b2 ? 1 : -1;
      if (a1 !== Infinity && b1 !== Infinity) {
        const diff = parseInt(a1) - parseInt(b1);
        if (diff) return diff;
      }
    }
    return ax.length - bx.length;
  }

  // Live search
  let to = false;
  $('#search').on('keyup', function () {
    if (to) clearTimeout(to);
    to = setTimeout(() => {
      $('#tree').jstree(true).search(this.value);
    }, 250);
  });

  // Show details in right panel on click
  $('#tree').on('select_node.jstree', function (e, data) {
    const nodeData = data.node.data;
    const details = document.getElementById('details');

    if (!nodeData || Object.keys(nodeData).length === 0) {
      details.innerHTML = '<p>No details available.</p>';
      return;
    }

    let html = '<table><tbody>';
    for (const [key, value] of Object.entries(nodeData)) {
      if (Array.isArray(value)) continue; // skip nested arrays
      html += `<tr><th>${key}</th><td>${value || ""}</td></tr>`;
    }
    html += '</tbody></table>';
    details.innerHTML = html;
  });
}

loadTree();
