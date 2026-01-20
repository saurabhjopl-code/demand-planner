const SIZE_ORDER = ["S","M","L","XL","XXL","3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"];
const NORMAL = new Set(["S","M","L","XL","XXL"]);
const PLUS1 = new Set(["3XL","4XL","5XL","6XL"]);
const PLUS2 = new Set(["7XL","8XL","9XL","10XL"]);

function readFile(file) {
  return new Promise(resolve => {
    const r = new FileReader();
    r.onload = e => {
      const wb = XLSX.read(e.target.result, { type: "array" });
      resolve(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: 0 }));
    };
    r.readAsArrayBuffer(file);
  });
}

function normalizeSKU(sku) {
  if (!sku) return { style:"", size:null };
  const p = sku.split("-");
  if (p.length < 2 || !p[1] || p[1] === "undefined") return { style:p[0], size:null };
  return { style:p[0], size:p[1] };
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
  let n=0,p1=0,p2=0;

  sales.forEach(r => {
    const {style,size} = normalizeSKU(r.SKU);
    map[style] ??= {};
    map[style][size] ??= { sales:0, stock:0 };
    map[style][size].sales += Number(r.Quantity);

    if (NORMAL.has(size)) n += Number(r.Quantity);
    if (PLUS1.has(size)) p1 += Number(r.Quantity);
    if (PLUS2.has(size)) p2 += Number(r.Quantity);
  });

  stock.forEach(r => {
    const {style,size} = normalizeSKU(r.SKU);
    map[style] ??= {};
    map[style][size] ??= { sales:0, stock:0 };
    map[style][size].stock += Number(r["Available Stock"]);
  });

  renderDemandAndOverstock(map, sd, target);
  renderSizeMix(map);
  renderSizeAnalysis(n, p1, p2);
}

function renderDemandAndOverstock(map, sd, target) {
  demandTable.tBodies[0].innerHTML = "";
  overstockTable.tBodies[0].innerHTML = "";

  let demandRows=[], overstockRows=[];

  Object.entries(map).forEach(([style,sizes])=>{
    let s=0, st=0;
    Object.values(sizes).forEach(v=>{ s+=v.sales; st+=v.stock; });
    const drr = s ? s/sd : 0;
    const sc = drr ? st/drr : Infinity;
    const demand = (drr && sc < target) ? Math.ceil((target-sc)*drr) : 0;

    const sortedSizes = Object.entries(sizes)
      .filter(([sz])=>sz)
      .sort((a,b)=>SIZE_ORDER.indexOf(a[0])-SIZE_ORDER.indexOf(b[0]));

    if (demand>0) demandRows.push({style,s,st,drr,sc,demand,sortedSizes});
    if (sc>120) overstockRows.push({style,s,st,drr,sc,sortedSizes});
  });

  demandRows.sort((a,b)=>b.demand-a.demand);
  overstockRows.sort((a,b)=>b.sc-a.sc);

  renderTable(demandTable,demandRows,true);
  renderTable(overstockTable,overstockRows,false);
}

function renderTable(table, rows, showDemand) {
  const tb = table.tBodies[0];
  rows.forEach((r,i)=>{
    const id="r"+i;
    tb.insertAdjacentHTML("beforeend",`
      <tr class="main-row" data-style="${r.style}">
        <td class="expand-btn" onclick="toggle('${id}',this)">+</td>
        <td>${r.style}</td><td>${r.s}</td><td>${r.st}</td>
        <td>${r.drr.toFixed(2)}</td>
        <td>${isFinite(r.sc)?r.sc.toFixed(1):"No Sales"}</td>
        ${showDemand?`<td>${r.demand}</td>`:""}
      </tr>`);

    r.sortedSizes.forEach(([size,v])=>{
      const drrS = v.sales ? v.sales/Number(salesDays.value) : 0;
      const scS = drrS ? v.stock/drrS : Infinity;
      const dS = (drrS && scS<Number(targetSC.value)) ? Math.ceil((Number(targetSC.value)-scS)*drrS) : 0;

      tb.insertAdjacentHTML("beforeend",`
        <tr class="sub-row hidden ${id}">
          <td></td><td>${r.style}-${size}</td>
          <td>${v.sales}</td><td>${v.stock}</td>
          <td>${drrS.toFixed(2)}</td>
          <td>${isFinite(scS)?scS.toFixed(1):"No Sales"}</td>
          ${showDemand?`<td>${dS}</td>`:""}
        </tr>`);
    });
  });
}

function renderSizeMix(map) {
  const tb = sizeMixTable.tBodies[0];
  tb.innerHTML = "";

  Object.entries(map).forEach(([style,sizes])=>{
    let t=0,n=0,p1=0,p2=0;
    Object.entries(sizes).forEach(([sz,v])=>{
      if (!sz) return;
      t += v.sales;
      if (NORMAL.has(sz)) n+=v.sales;
      if (PLUS1.has(sz)) p1+=v.sales;
      if (PLUS2.has(sz)) p2+=v.sales;
    });
    if (t<20) return;

    tb.insertAdjacentHTML("beforeend",`
      <tr>
        <td>${style}</td>
        <td>${t}</td>
        <td>${((n/t)*100).toFixed(1)}%</td>
        <td>${((p1/t)*100).toFixed(1)}%</td>
        <td>${((p2/t)*100).toFixed(1)}%</td>
      </tr>`);
  });
}

function renderSizeAnalysis(n,p1,p2){
  const total=n+p1+p2;
  normalSales.innerText=n;
  plus1Sales.innerText=p1;
  plus2Sales.innerText=p2;
  normalPct.innerText=total?((n/total)*100).toFixed(1)+"%":"0%";
  plus1Pct.innerText=total?((p1/total)*100).toFixed(1)+"%":"0%";
  plus2Pct.innerText=total?((p2/total)*100).toFixed(1)+"%":"0%";
}

function toggle(id,el){
  document.querySelectorAll("."+id).forEach(r=>r.classList.toggle("hidden"));
  el.innerText = el.innerText==="+"?"−":"+";
}

function filterTables(){
  const q = search.value.toLowerCase();
  document.querySelectorAll(".main-row").forEach(r=>{
    r.style.display = r.dataset.style.toLowerCase().includes(q) ? "" : "none";
  });
}

function showTab(tab){
  document.querySelectorAll(".tab-btn").forEach(b=>b.classList.remove("active"));
  document.querySelectorAll(".tab-content").forEach(c=>c.classList.remove("active"));
  if(tab==="demand"){btns[0].classList.add("active");demandTab.classList.add("active");}
  if(tab==="overstock"){btns[1].classList.add("active");overstockTab.classList.add("active");}
  if(tab==="sizemix"){btns[2].classList.add("active");sizemixTab.classList.add("active");}
  if(tab==="sizeanalysis"){btns[3].classList.add("active");sizeanalysisTab.classList.add("active");}
}

/* DOM */
const salesFile=document.getElementById("salesFile");
const stockFile=document.getElementById("stockFile");
const salesDays=document.getElementById("salesDays");
const targetSC=document.getElementById("targetSC");
const demandTable=document.getElementById("demandTable");
const overstockTable=document.getElementById("overstockTable");
const sizeMixTable=document.getElementById("sizeMixTable");
const search=document.getElementById("search");
const btns=document.querySelectorAll(".tab-btn");
const demandTab=document.getElementById("demandTab");
const overstockTab=document.getElementById("overstockTab");
const sizemixTab=document.getElementById("sizemixTab");
const sizeanalysisTab=document.getElementById("sizeanalysisTab");

const normalSales=document.getElementById("normalSales");
const plus1Sales=document.getElementById("plus1Sales");
const plus2Sales=document.getElementById("plus2Sales");
const normalPct=document.getElementById("normalPct");
const plus1Pct=document.getElementById("plus1Pct");
const plus2Pct=document.getElementById("plus2Pct");
