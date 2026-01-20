let globalDemand = [], globalOverstock = [];

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

function processFiles() {
  Promise.all([
    readFile(salesFile.files[0]),
    readFile(stockFile.files[0])
  ]).then(([s, st]) => calculate(s, st));
}

function calculate(sales, stock) {
  const sd = +salesDays.value, target = +targetSC.value;
  const map = {}, bands = { b0:0, b30:0, b60:0, b120:0 };

  sales.forEach(r => {
    const [style, size] = r.SKU.split("-");
    map[style] ??= {};
    map[style][size] ??= { sales:0, stock:0 };
    map[style][size].sales += +r.Quantity;
  });

  stock.forEach(r => {
    const [style, size] = r.SKU.split("-");
    map[style] ??= {};
    map[style][size] ??= { sales:0, stock:0 };
    map[style][size].stock += +r["Available Stock"];
  });

  demandTable.tBodies[0].innerHTML = "";
  overstockTable.tBodies[0].innerHTML = "";
  globalDemand = []; globalOverstock = [];

  Object.entries(map).forEach(([style, sizes]) => {
    let sales=0, stock=0;
    Object.values(sizes).forEach(v => { sales+=v.sales; stock+=v.stock; });

    const drr = sales ? sales/sd : 0;
    const sc = drr ? stock/drr : 9999;

    if (sc <= 30) bands.b0++;
    else if (sc <= 60) bands.b30++;
    else if (sc <= 120) bands.b60++;
    else bands.b120++;

    const demand = sc < target ? Math.ceil((target-sc)*drr) : 0;

    if (demand > 0) {
      globalDemand.push({style,sales,stock,drr,sc,demand,sizes});
      addRow(demandTable, style, sizes, sales, stock, drr, sc, demand, true);
    }
    if (sc > 120) {
      globalOverstock.push({style,sales,stock,drr,sc,sizes});
      addRow(overstockTable, style, sizes, sales, stock, drr, sc, 0, false);
    }
  });

  Object.keys(bands).forEach(k => document.getElementById(k).innerText = bands[k]);
}

function addRow(table, style, sizes, sales, stock, drr, sc, demand, showDemand) {
  const id = Math.random().toString(36).slice(2,7);
  const tb = table.tBodies[0];
  tb.insertAdjacentHTML("beforeend",
    `<tr class="main-row">
      <td onclick="toggle('${id}')" class="expand-btn">+</td>
      <td>${style}</td><td>${sales}</td><td>${stock}</td>
      <td>${drr.toFixed(2)}</td><td>${sc.toFixed(1)}</td>
      ${showDemand?`<td>${demand}</td>`:""}
    </tr>`
  );

  Object.entries(sizes).forEach(([size,v])=>{
    const drrS = v.sales ? v.sales/salesDays.value : 0;
    const scS = drrS ? v.stock/drrS : 9999;
    tb.insertAdjacentHTML("beforeend",
      `<tr class="sub-row hidden ${id}">
        <td></td><td>${style}-${size}</td>
        <td>${v.sales}</td><td>${v.stock}</td>
        <td>${drrS.toFixed(2)}</td><td>${scS.toFixed(1)}</td>
        ${showDemand?"<td></td>":""}
      </tr>`
    );
  });
}

function toggle(id){
  document.querySelectorAll(`.${id}`).forEach(r=>r.classList.toggle("hidden"));
}

function filterTables(){
  const q = search.value.toLowerCase();
  document.querySelectorAll("table tbody tr.main-row").forEach(r=>{
    r.style.display = r.children[1].innerText.toLowerCase().includes(q) ? "" : "none";
  });
}

function exportExcel(){
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb,
    XLSX.utils.json_to_sheet(globalDemand.map(r=>({
      Style:r.style, Sales:r.sales, Stock:r.stock, DRR:r.drr, SC:r.sc, Demand:r.demand
    }))), "Demand Report");

  XLSX.utils.book_append_sheet(wb,
    XLSX.utils.json_to_sheet(globalOverstock.map(r=>({
      Style:r.style, Sales:r.sales, Stock:r.stock, DRR:r.drr, SC:r.sc
    }))), "Overstock");

  XLSX.writeFile(wb,"Demand_Planning_Report.xlsx");
}

/* DOM */
const salesFile=document.getElementById("salesFile");
const stockFile=document.getElementById("stockFile");
const salesDays=document.getElementById("salesDays");
const targetSC=document.getElementById("targetSC");
const demandTable=document.getElementById("demandTable");
const overstockTable=document.getElementById("overstockTable");
const search=document.getElementById("search");
