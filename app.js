let globalDemand = [], globalOverstock = [];

function readFile(file) {
  return new Promise(resolve => {
    const reader = new FileReader();
    reader.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      resolve(XLSX.utils.sheet_to_json(sheet, { defval: 0 }));
    };
    reader.readAsArrayBuffer(file);
  });
}

function processFiles() {
  if (!salesFile.files[0] || !stockFile.files[0]) {
    alert("Upload both files");
    return;
  }

  Promise.all([
    readFile(salesFile.files[0]),
    readFile(stockFile.files[0])
  ]).then(([sales, stock]) => calculate(sales, stock));
}

function calculate(sales, stock) {
  const sd = Number(salesDays.value);
  const target = Number(targetSC.value);

  const map = {};
  const bands = { b0:0, b30:0, b60:0, b120:0 };

  sales.forEach(r => {
    const [style, size] = r.SKU.split("-");
    map[style] ??= {};
    map[style][size] ??= { sales: 0, stock: 0 };
    map[style][size].sales += Number(r.Quantity);
  });

  stock.forEach(r => {
    const [style, size] = r.SKU.split("-");
    map[style] ??= {};
    map[style][size] ??= { sales: 0, stock: 0 };
    map[style][size].stock += Number(r["Available Stock"]);
  });

  demandTable.tBodies[0].innerHTML = "";
  overstockTable.tBodies[0].innerHTML = "";
  globalDemand = [];
  globalOverstock = [];

  Object.entries(map).forEach(([style, sizes]) => {
    let totalSales = 0, totalStock = 0;
    Object.values(sizes).forEach(v => {
      totalSales += v.sales;
      totalStock += v.stock;
    });

    const drr = totalSales > 0 ? totalSales / sd : 0;
    const sc = drr > 0 ? totalStock / drr : Infinity;

    if (sc <= 30) bands.b0++;
    else if (sc <= 60) bands.b30++;
    else if (sc <= 120) bands.b60++;
    else bands.b120++;

    const demand = (drr > 0 && sc < target)
      ? Math.ceil((target - sc) * drr)
      : 0;

    if (demand > 0) {
      globalDemand.push({ style, totalSales, totalStock, drr, sc, demand, sizes });
      addRow(demandTable, style, sizes, totalSales, totalStock, drr, sc, demand, true);
    }

    if (sc > 120) {
      globalOverstock.push({ style, totalSales, totalStock, drr, sc, sizes });
      addRow(overstockTable, style, sizes, totalSales, totalStock, drr, sc, 0, false);
    }
  });

  Object.keys(bands).forEach(k => document.getElementById(k).innerText = bands[k]);
}

function addRow(table, style, sizes, sales, stock, drr, sc, demand, showDemand) {
  const id = Math.random().toString(36).slice(2,7);
  const tb = table.tBodies[0];

  tb.insertAdjacentHTML("beforeend", `
    <tr class="main-row">
      <td class="expand-btn" onclick="toggle('${id}')">+</td>
      <td>${style}</td>
      <td>${sales}</td>
      <td>${stock}</td>
      <td>${drr.toFixed(2)}</td>
      <td>${isFinite(sc) ? sc.toFixed(1) : "No Sales"}</td>
      ${showDemand ? `<td>${demand}</td>` : ""}
    </tr>
  `);

  Object.entries(sizes).forEach(([size, v]) => {
    const drrS = v.sales > 0 ? v.sales / Number(salesDays.value) : 0;
    const scS = drrS > 0 ? v.stock / drrS : Infinity;

    let demandS = 0;
    if (drrS > 0 && scS < Number(targetSC.value)) {
      demandS = Math.ceil((Number(targetSC.value) - scS) * drrS);
    }

    tb.insertAdjacentHTML("beforeend", `
      <tr class="sub-row hidden ${id}">
        <td></td>
        <td>${style}-${size}</td>
        <td>${v.sales}</td>
        <td>${v.stock}</td>
        <td>${drrS.toFixed(2)}</td>
        <td>${isFinite(scS) ? scS.toFixed(1) : "No Sales"}</td>
        ${showDemand ? `<td>${demandS}</td>` : ""}
      </tr>
    `);
  });
}

function toggle(id) {
  document.querySelectorAll("." + id).forEach(r => r.classList.toggle("hidden"));
}

function showTab(tab) {
  document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
  document.querySelectorAll(".tab-content").forEach(c => c.classList.remove("active"));

  if (tab === "demand") {
    document.querySelector(".tab-btn:nth-child(1)").classList.add("active");
    demandTab.classList.add("active");
  } else {
    document.querySelector(".tab-btn:nth-child(2)").classList.add("active");
    overstockTab.classList.add("active");
  }
}

function filterTables() {
  const q = search.value.toLowerCase();
  document.querySelectorAll(".main-row").forEach(r => {
    r.style.display = r.children[1].innerText.toLowerCase().includes(q) ? "" : "none";
  });
}

function exportExcel() {
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(globalDemand.map(r => ({
      Style: r.style,
      Sales: r.totalSales,
      Stock: r.totalStock,
      DRR: r.drr,
      SC: r.sc,
      Demand: r.demand
    }))),
    "Demand Report"
  );

  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.json_to_sheet(globalOverstock.map(r => ({
      Style: r.style,
      Sales: r.totalSales,
      Stock: r.totalStock,
      DRR: r.drr,
      SC: r.sc
    }))),
    "Overstock"
  );

  XLSX.writeFile(wb, "Demand_Planner_Report.xlsx");
}

/* DOM */
const salesFile = document.getElementById("salesFile");
const stockFile = document.getElementById("stockFile");
const salesDays = document.getElementById("salesDays");
const targetSC = document.getElementById("targetSC");
const demandTable = document.getElementById("demandTable");
const overstockTable = document.getElementById("overstockTable");
const demandTab = document.getElementById("demandTab");
const overstockTab = document.getElementById("overstockTab");
const search = document.getElementById("search");
