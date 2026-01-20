const SIZE_ORDER = ["S","M","L","XL","XXL","3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"];
const NORMAL_SIZES = new Set(["S","M","L","XL","XXL"]);
const PLUS_SIZES = new Set(["3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"]);

let demandRows = [], overstockRows = [];

function readFile(file) {
  return new Promise(resolve => {
    const r = new FileReader();
    r.onload = e => {
      const wb = XLSX.read(e.target.result, { type: 'array' });
      resolve(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: 0 }));
    };
    r.readAsArrayBuffer(file);
  });
}

function normalizeSKU(sku) {
  if (!sku) return { style: "", size: null };
  const parts = sku.split("-");
  if (parts.length < 2 || !parts[1] || parts[1] === "undefined") {
    return { style: parts[0], size: null };
  }
  return { style: parts[0], size: parts[1] };
}

function processFiles() {
  Promise.all([
    readFile(salesFile.files[0]),
    readFile(stockFile.files[0])
  ]).then(([sales, stock]) => calculate(sales, stock));
}

function calculate(sales, stock) {
  const sd = Number(salesDays.value);
  const target = Number(targetSC.value);
  const map = {};
  let normalSales = 0, plusSales = 0;

  sales.forEach(r => {
    const { style, size } = normalizeSKU(r.SKU);
    map[style] ??= {};
    map[style][size] ??= { sales:0, stock:0 };
    map[style][size].sales += Number(r.Quantity);

    if (NORMAL_SIZES.has(size)) normalSales += Number(r.Quantity);
    if (PLUS_SIZES.has(size)) plusSales += Number(r.Quantity);
  });

  stock.forEach(r => {
    const { style, size } = normalizeSKU(r.SKU);
    map[style] ??= {};
    map[style][size] ??= { sales:0, stock:0 };
    map[style][size].stock += Number(r["Available Stock"]);
  });

  const totalSales = normalSales + plusSales;
  normalSalesEl.innerText = normalSales;
  plusSalesEl.innerText = plusSales;
  normalPctEl.innerText = totalSales ? ((normalSales/totalSales)*100).toFixed(1)+"%" : "0%";
  plusPctEl.innerText = totalSales ? ((plusSales/totalSales)*100).toFixed(1)+"%" : "0%";

  demandRows = [];
  overstockRows = [];

  Object.entries(map).forEach(([style, sizes]) => {
    let s=0, st=0;
    Object.values(sizes).forEach(v => { s+=v.sales; st+=v.stock; });
    const drr = s ? s/sd : 0;
    const sc = drr ? st/drr : Infinity;
    const demand = (drr && sc < target) ? Math.ceil((target-sc)*drr) : 0;

    const sortedSizes = Object.entries(sizes)
      .sort((a,b)=>SIZE_ORDER.indexOf(a[0]) - SIZE_ORDER.indexOf(b[0]));

    if (demand > 0) demandRows.push({ style, s, st, drr, sc, demand, sizes: sortedSizes });
    if (sc > 120) overstockRows.push({ style, s, st, drr, sc, sizes: sortedSizes });
  });

  demandRows.sort((a,b)=>b.demand-a.demand);
  overstockRows.sort((a,b)=>b.sc-a.sc);

  renderTable(demandTable, demandRows, true);
  renderTable(overstockTable, overstockRows, false);
}

function renderTable(table, rows, showDemand) {
  const tb = table.tBodies[0];
  tb.innerHTML = "";

  rows.forEach((r,i)=>{
    const id = "row_"+i;
    tb.insertAdjacentHTML("beforeend", `
      <tr class="main-row" data-style="${r.style}">
        <td class="expand-btn" onclick="toggle('${id}',this)">+</td>
        <td>${r.style}</td><td>${r.s}</td><td>${r.st}</td>
        <td>${r.drr.toFixed(2)}</td>
        <td>${isFinite(r.sc)?r.sc.toFixed(1):"No Sales"}</td>
        ${showDemand?`<td>${r.demand}</td>`:""}
      </tr>
    `);

    r.sizes.forEach(([size,v])=>{
      if (!size) return;
      const drrS = v.sales ? v.sales/Number(salesDays.value) : 0;
      const scS = drrS ? v.stock/drrS : Infinity;
      const dS = (drrS && scS < Number(targetSC.value)) ? Math.ceil((Number(targetSC.value)-scS)*drrS) : 0;

      tb.insertAdjacentHTML("beforeend", `
        <tr class="sub-row hidden ${id}">
          <td></td>
          <td>${r.style}-${size}</td>
          <td>${v.sales}</td>
          <td>${v.stock}</td>
          <td>${drrS.toFixed(2)}</td>
          <td>${isFinite(scS)?scS.toFixed(1):"No Sales"}</td>
          ${showDemand?`<td>${dS}</td>`:""}
        </tr>
      `);
    });
  });
}

function toggle(id, el) {
  document.querySelectorAll("."+id).forEach(r=>r.classList.toggle("hidden"));
  el.innerText = el.innerText === "+" ? "−" : "+";
}

function filterTables() {
  const q = search.value.toLowerCase();
  document.querySelectorAll(".main-row").forEach(r=>{
    r.style.display = r.dataset.style.toLowerCase().includes(q) ? "" : "none";
  });
}

function showTab(tab) {
  document.querySelectorAll(".tab-btn").forEach(b=>b.classList.remove("active"));
  document.querySelectorAll(".tab-content").forEach(c=>c.classList.remove("active"));
  if (tab==="demand") { btns[0].classList.add("active"); demandTab.classList.add("active"); }
  if (tab==="overstock") { btns[1].classList.add("active"); overstockTab.classList.add("active"); }
  if (tab==="size") { btns[2].classList.add("active"); sizeTab.classList.add("active"); }
}

/* DOM */
const salesFile = document.getElementById("salesFile");
const stockFile = document.getElementById("stockFile");
const salesDays = document.getElementById("salesDays");
const targetSC = document.getElementById("targetSC");
const demandTable = document.getElementById("demandTable");
const overstockTable = document.getElementById("overstockTable");
const search = document.getElementById("search");
const btns = document.querySelectorAll(".tab-btn");
const demandTab = document.getElementById("demandTab");
const overstockTab = document.getElementById("overstockTab");
const sizeTab = document.getElementById("sizeTab");

const normalSalesEl = document.getElementById("normalSales");
const plusSalesEl = document.getElementById("plusSales");
const normalPctEl = document.getElementById("normalPct");
const plusPctEl = document.getElementById("plusPct");
