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
  Promise.all([
    readFile(salesFile.files[0]),
    readFile(stockFile.files[0])
  ]).then(([sales, stock]) => calculate(sales, stock));
}

function calculate(sales, stock) {
  const salesDays = +salesDaysInput.value;
  const targetSC = +targetSCInput.value;

  const map = {};

  sales.forEach(r => {
    const [style, size] = r.SKU.split("-");
    map[style] ??= {};
    map[style][size] ??= { sales: 0, stock: 0 };
    map[style][size].sales += +r.Quantity;
  });

  stock.forEach(r => {
    const [style, size] = r.SKU.split("-");
    map[style] ??= {};
    map[style][size] ??= { sales: 0, stock: 0 };
    map[style][size].stock += +r["Available Stock"];
  });

  render(map, salesDays, targetSC);
}

function render(data, salesDays, targetSC) {
  demandTable.tBodies[0].innerHTML = "";
  overstockTable.tBodies[0].innerHTML = "";

  Object.entries(data).forEach(([style, sizes]) => {
    let sales = 0, stock = 0;
    Object.values(sizes).forEach(v => {
      sales += v.sales;
      stock += v.stock;
    });

    const drr = sales ? sales / salesDays : 0;
    const sc = drr ? stock / drr : 9999;
    const demand = sc < targetSC ? Math.ceil((targetSC - sc) * drr) : 0;

    if (demand > 0) addRow(demandTable, style, sizes, sales, stock, drr, sc, demand, true);
    if (sc > 120) addRow(overstockTable, style, sizes, sales, stock, drr, sc, 0, false);
  });
}

function addRow(table, style, sizes, sales, stock, drr, sc, demand, showDemand) {
  const tbody = table.tBodies[0];
  const id = Math.random().toString(36).substr(2, 5);

  const tr = tbody.insertRow();
  tr.innerHTML = `
    <td class="expand-btn" onclick="toggle('${id}')">+</td>
    <td>${style}</td>
    <td>${sales}</td>
    <td>${stock}</td>
    <td>${drr.toFixed(2)}</td>
    <td>${sc.toFixed(1)}</td>
    ${showDemand ? `<td>${demand}</td>` : ""}
  `;

  Object.entries(sizes).forEach(([size, v]) => {
    const sr = tbody.insertRow();
    sr.className = `sub-row hidden ${id}`;
    const drrS = v.sales ? v.sales / salesDays : 0;
    const scS = drrS ? v.stock / drrS : 9999;

    sr.innerHTML = `
      <td></td>
      <td>${style}-${size}</td>
      <td>${v.sales}</td>
      <td>${v.stock}</td>
      <td>${drrS.toFixed(2)}</td>
      <td>${scS.toFixed(1)}</td>
      ${showDemand ? `<td></td>` : ""}
    `;
  });
}

function toggle(id) {
  document.querySelectorAll(`.${id}`).forEach(r => {
    r.classList.toggle("hidden");
  });
}

/* DOM shortcuts */
const salesFile = document.getElementById("salesFile");
const stockFile = document.getElementById("stockFile");
const salesDaysInput = document.getElementById("salesDays");
const targetSCInput = document.getElementById("targetSC");
const demandTable = document.getElementById("demandTable");
const overstockTable = document.getElementById("overstockTable");
