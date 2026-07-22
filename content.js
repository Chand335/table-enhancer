document.querySelectorAll("table").forEach((table, index) => {
  // Skip if already converted
  if (table.dataset.gridConverted) return;

  table.dataset.gridConverted = true;

  // Headers
  const columns = [...table.querySelectorAll("thead th")].map((th) => ({
    name: th.innerText.trim(),
  }));
  const rows = [...table.querySelectorAll("tbody tr")];
  // Data
  const data = [...table.querySelectorAll("tbody tr")].map((tr) =>
    [...tr.children].map((td) => gridjs.html(td.innerHTML)),
  );
  // Container
  const wrapper = document.createElement("div");
  wrapper.id = "gridjs_" + index;

  table.parentNode.insertBefore(wrapper, table);

  table.style.display = "none";

  new gridjs.Grid({
    columns,
    data,

    search: true,

    sort: true,

    pagination:
      rows.length > 10
        ? {
            enabled: true,
            limit: 10,
          }
        : false,

    resizable: true,

    fixedHeader: true,

    height: rows.length > 10 ? "600px" : "auto",
  }).render(wrapper);

  const toolbar = document.createElement("div");
  toolbar.className = "grid-toolbar";
  // wrapper.before(toolbar);
  // Add buttons to the toolbar in .gridjs-head
  document.querySelector(".gridjs-head").appendChild(toolbar);

  // <button type="button" class="grid-export-csv">CSV</button>
  // <button type="button" class="grid-copy">Copy</button>
  toolbar.innerHTML = `
    <button type="button" class="grid-export-excel">Excel</button>
    <button type="button" class="grid-print">Print</button>
`;
  toolbar.querySelector(".grid-export-csv").addEventListener("click", () => {
    exportCSV(columns, data);
  });
  toolbar.querySelector(".grid-print").onclick = () => {
    // print only .gridjs-wrapper
    const wrapper = document.querySelector(".gridjs-wrapper");
    if (wrapper) {
      printWrapperWithStyles(wrapper);
    }
  };
  toolbar.querySelector(".grid-export-excel").addEventListener("click", () => {
    exportExcel(table);
  });
});
function exportExcel(table) {
  // Clone so we don't modify the live DOM
  const clone = table.cloneNode(true);

  // Replace links/buttons/icons with their visible text
  clone.querySelectorAll("a, button").forEach((el) => {
    el.replaceWith(document.createTextNode(el.innerText.trim()));
  });

  clone.querySelectorAll("input").forEach((el) => {
    el.replaceWith(document.createTextNode(el.value));
  });

  clone.querySelectorAll("select").forEach((el) => {
    const txt = el.options[el.selectedIndex]?.text || "";
    el.replaceWith(document.createTextNode(txt));
  });

  // Convert HTML table to worksheet
  const worksheet = XLSX.utils.table_to_sheet(clone);

  // Create workbook
  const workbook = XLSX.utils.book_new();

  const range = XLSX.utils.decode_range(worksheet["!ref"]);

  const cols = [];

  for (let c = range.s.c; c <= range.e.c; c++) {
    let max = 10;

    for (let r = range.s.r; r <= range.e.r; r++) {
      const cell = worksheet[XLSX.utils.encode_cell({ r, c })];

      if (cell && cell.v) {
        max = Math.max(max, cell.v.toString().length);
      }
    }

    cols.push({
      wch: Math.min(max + 3, 50),
    });
  }

  worksheet["!cols"] = cols;

  XLSX.utils.book_append_sheet(workbook, worksheet, "Data");

  // Download
  XLSX.writeFile(workbook, getFileName());
}

function getFileName() {
  const title = document.title.replace(/[<>:"/\\|?*]+/g, "").trim();

  const date = new Date().toISOString().slice(0, 10);

  return `${title}_${date}.xlsx`;
}

function createStylesheetLoadPromise(linkElement) {
  return new Promise((resolve) => {
    const finish = () => resolve();
    linkElement.onload = finish;
    linkElement.onerror = finish;
  });
}

function printWrapperWithStyles(wrapper) {
  const printWindow = window.open("", "_blank");
  if (!printWindow) {
    return;
  }

  const printDoc = printWindow.document;
  const wrapperClone = wrapper.cloneNode(true);
  wrapperClone.querySelectorAll("button").forEach((button) => button.remove());
  const styleNodes = [
    ...document.querySelectorAll("style, link[rel='stylesheet']"),
  ];
  const stylesheetLoads = [];

  printDoc.title = "Print";
  printDoc.head.innerHTML = "";
  printDoc.body.innerHTML = "";

  styleNodes.forEach((node) => {
    if (node.tagName === "STYLE") {
      const style = printDoc.createElement("style");
      style.textContent = node.textContent;
      printDoc.head.appendChild(style);
      return;
    }

    const href = node.getAttribute("href");
    if (!href) {
      return;
    }

    const link = printDoc.createElement("link");
    link.rel = "stylesheet";
    link.href = href;
    printDoc.head.appendChild(link);
    stylesheetLoads.push(createStylesheetLoadPromise(link));
  });

  printDoc.body.appendChild(wrapperClone);

  Promise.all(stylesheetLoads).finally(() => {
    printWindow.focus();
    printWindow.print();
  });
}

function copyTable(columns, data) {
  let text = "";

  text += columns.map((c) => c.name).join("\t") + "\n";

  data.forEach((row) => {
    text += row
      .map((cell) => {
        const div = document.createElement("div");
        div.innerHTML = cell.toString();

        return div.textContent;
      })
      .join("\t");

    text += "\n";
  });

  navigator.clipboard.writeText(text);
}

function exportCSV(columns, data, filename = "table.csv") {
  const csv = [];

  csv.push(columns.map((c) => `"${c.name}"`).join(","));

  data.forEach((row) => {
    csv.push(
      row
        .map((cell) => {
          const div = document.createElement("div");
          div.innerHTML = cell.toString();
          return `"${div.textContent.replaceAll('"', '""')}"`;
        })
        .join(","),
    );
  });

  const blob = new Blob([csv.join("\n")], {
    type: "text/csv;charset=utf-8;",
  });

  const link = document.createElement("a");

  link.href = URL.createObjectURL(blob);
  link.download = filename;
  link.click();
}
