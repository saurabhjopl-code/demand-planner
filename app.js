const SIZE_ORDER=["S","M","L","XL","XXL","3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"];
const NORMAL=new Set(["S","M","L","XL","XXL"]);
const PLUS1=new Set(["3XL","4XL","5XL","6XL"]);
const PLUS2=new Set(["7XL","8XL","9XL","10XL"]);

function readFile(f){
  return new Promise(r=>{
    const fr=new FileReader();
    fr.onload=e=>{
      const wb=XLSX.read(e.target.result,{type:"array"});
      r(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:0}));
    };
    fr.readAsArrayBuffer(f);
  });
}

function normalizeSKU(s){
  if(!s) return {style:"",size:null};
  const p=s.split("-");
  if(p.length<2||!p[1]||p[1]==="undefined") return {style:p[0],size:null};
  return {style:p[0],size:p[1]};
}

function processFiles(){
  Promise.all([
    readFile(salesFile.files[0]),
    readFile(stockFile.files[0])
  ]).then(([sales,stock])=>calculate(sales,stock));
}

function calculate(sales,stock){
  const sd=+salesDays.value, target=+targetSC.value;
  const map={}, bands={b0:0,b30:0,b60:0,b120:0};
  let n=0,p1=0,p2=0;

  sales.forEach(r=>{
    const {style,size}=normalizeSKU(r.SKU);
    map[style]??={};
    map[style][size]??={sales:0,stock:0};
    map[style][size].sales+=+r.Quantity;
    if(NORMAL.has(size))n+=+r.Quantity;
    if(PLUS1.has(size))p1+=+r.Quantity;
    if(PLUS2.has(size))p2+=+r.Quantity;
  });

  stock.forEach(r=>{
    const {style,size}=normalizeSKU(r.SKU);
    map[style]??={};
    map[style][size]??={sales:0,stock:0};
    map[style][size].stock+=+r["Available Stock"];
  });

  renderDemandOverstock(map,sd,target,bands);
  renderSizeMix(map);
  renderSizeAnalysis(n,p1,p2);
  renderSizeCurve(map,sd,target);

  Object.keys(bands).forEach(k=>document.getElementById(k).innerText=bands[k]);
}

/* ---------------- DEMAND + OVERSTOCK ---------------- */
function renderDemandOverstock(map,sd,target,bands){
  demandTable.tBodies[0].innerHTML="";
  overstockTable.tBodies[0].innerHTML="";
  let dRows=[],oRows=[];

  Object.entries(map).forEach(([style,sizes])=>{
    let s=0,st=0;
    Object.values(sizes).forEach(v=>{s+=v.sales;st+=v.stock});
    const drr=s?s/sd:0;
    const sc=drr?st/drr:Infinity;

    if(sc<=30)bands.b0++; else if(sc<=60)bands.b30++;
    else if(sc<=120)bands.b60++; else bands.b120++;

    const demand=(drr&&sc<target)?Math.ceil((target-sc)*drr):0;

    const sz=Object.entries(sizes)
      .filter(([z])=>z)
      .sort((a,b)=>SIZE_ORDER.indexOf(a[0])-SIZE_ORDER.indexOf(b[0]));

    if(demand>0)dRows.push({style,s,st,drr,sc,demand,sz});
    if(sc>120)oRows.push({style,s,st,drr,sc,sz});
  });

  dRows.sort((a,b)=>b.demand-a.demand);
  oRows.sort((a,b)=>b.sc-a.sc);

  renderTable(demandTable,dRows,true);
  renderTable(overstockTable,oRows,false);
}

function renderTable(table,rows,showDemand){
  const tb=table.tBodies[0];
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
    r.sz.forEach(([z,v])=>{
      const drrS=v.sales?v.sales/+salesDays.value:0;
      const scS=drrS?v.stock/drrS:Infinity;
      const dS=(drrS&&scS<+targetSC.value)?Math.ceil((+targetSC.value-scS)*drrS):0;
      tb.insertAdjacentHTML("beforeend",`
        <tr class="sub-row hidden ${id}">
          <td></td><td>${r.style}-${z}</td>
          <td>${v.sales}</td><td>${v.stock}</td>
          <td>${drrS.toFixed(2)}</td>
          <td>${isFinite(scS)?scS.toFixed(1):"No Sales"}</td>
          ${showDemand?`<td>${dS}</td>`:""}
        </tr>`);
    });
  });
}

/* ---------------- SIZE MIX ---------------- */
function renderSizeMix(map){
  const tb=sizeMixTable.tBodies[0]; tb.innerHTML="";
  Object.entries(map).forEach(([style,sizes])=>{
    let t=0,n=0,p1=0,p2=0;
    Object.entries(sizes).forEach(([z,v])=>{
      if(!z)return;
      t+=v.sales;
      if(NORMAL.has(z))n+=v.sales;
      if(PLUS1.has(z))p1+=v.sales;
      if(PLUS2.has(z))p2+=v.sales;
    });
    if(t<20)return;
    tb.insertAdjacentHTML("beforeend",`
      <tr>
        <td>${style}</td><td>${t}</td>
        <td>${((n/t)*100).toFixed(1)}%</td>
        <td>${((p1/t)*100).toFixed(1)}%</td>
        <td>${((p2/t)*100).toFixed(1)}%</td>
      </tr>`);
  });
}

/* ---------------- SIZE ANALYSIS ---------------- */
function renderSizeAnalysis(n,p1,p2){
  const t=n+p1+p2;
  normalSales.innerText=n;
  plus1Sales.innerText=p1;
  plus2Sales.innerText=p2;
  normalPct.innerText=t?((n/t)*100).toFixed(1)+"%":"0%";
  plus1Pct.innerText=t?((p1/t)*100).toFixed(1)+"%":"0%";
  plus2Pct.innerText=t?((p2/t)*100).toFixed(1)+"%":"0%";
}

/* ---------------- DEMAND-AWARE SIZE CURVE ---------------- */
function renderSizeCurve(map,sd,target){
  const tb=sizeCurveTable.tBodies[0];
  tb.innerHTML="";

  Object.entries(map).forEach(([style,sizes])=>{
    let totalSales=0,totalStock=0;
    Object.values(sizes).forEach(v=>{totalSales+=v.sales; totalStock+=v.stock});
    if(totalSales<20) return;

    const drr=totalSales/sd;
    const sc=drr?totalStock/drr:Infinity;
    const styleDemand=(drr&&sc<target)?Math.ceil((target-sc)*drr):0;
    if(styleDemand<=0) return;

    // size shares
    let sizeShares=[];
    Object.entries(sizes).forEach(([z,v])=>{
      if(!z||v.sales===0) return;
      const drrS=v.sales/sd;
      const scS=drrS?v.stock/drrS:Infinity;

      let weight=v.sales/totalSales;

      // SC adjustment
      if(scS>target*1.5) weight*=0.3;
      else if(scS>target) weight*=0.6;
      else if(scS<target*0.5) weight*=1.3;

      sizeShares.push({z,weight});
    });

    const totalWeight=sizeShares.reduce((a,b)=>a+b.weight,0);
    let allocated=0;
    let curve=[];

    sizeShares.forEach((s,i)=>{
      let qty = Math.round((s.weight/totalWeight)*styleDemand);
      if(i===sizeShares.length-1) qty=styleDemand-allocated;
      allocated+=qty;
      if(qty>0) curve.push(`${s.z}:${qty}`);
    });

    tb.insertAdjacentHTML("beforeend",`
      <tr>
        <td>${style}</td>
        <td>${styleDemand}</td>
        <td>${curve.join(", ")}</td>
      </tr>`);
  });
}

/* ---------------- UI HELPERS ---------------- */
function toggle(id,el){
  document.querySelectorAll("."+id).forEach(r=>r.classList.toggle("hidden"));
  el.innerText=el.innerText==="+"?"−":"+";
}
function filterTables(){
  const q=search.value.toLowerCase();
  document.querySelectorAll(".main-row").forEach(r=>{
    r.style.display=r.dataset.style.toLowerCase().includes(q)?"":"none";
  });
}
function showTab(t){
  document.querySelectorAll(".tab-btn").forEach(b=>b.classList.remove("active"));
  document.querySelectorAll(".tab-content").forEach(c=>c.classList.remove("active"));
  if(t==="demand"){btns[0].classList.add("active");demandTab.classList.add("active")}
  if(t==="overstock"){btns[1].classList.add("active");overstockTab.classList.add("active")}
  if(t==="sizemix"){btns[2].classList.add("active");sizemixTab.classList.add("active")}
  if(t==="sizeanalysis"){btns[3].classList.add("active");sizeanalysisTab.classList.add("active")}
  if(t==="sizecurve"){btns[4].classList.add("active");sizecurveTab.classList.add("active")}
}

/* DOM */
const salesFile=document.getElementById("salesFile");
const stockFile=document.getElementById("stockFile");
const salesDays=document.getElementById("salesDays");
const targetSC=document.getElementById("targetSC");
const demandTable=document.getElementById("demandTable");
const overstockTable=document.getElementById("overstockTable");
const sizeMixTable=document.getElementById("sizeMixTable");
const sizeCurveTable=document.getElementById("sizeCurveTable");
const search=document.getElementById("search");
const btns=document.querySelectorAll(".tab-btn");
const demandTab=document.getElementById("demandTab");
const overstockTab=document.getElementById("overstockTab");
const sizemixTab=document.getElementById("sizemixTab");
const sizeanalysisTab=document.getElementById("sizeanalysisTab");
const sizecurveTab=document.getElementById("sizecurveTab");

const normalSales=document.getElementById("normalSales");
const plus1Sales=document.getElementById("plus1Sales");
const plus2Sales=document.getElementById("plus2Sales");
const normalPct=document.getElementById("normalPct");
const plus1Pct=document.getElementById("plus1Pct");
const plus2Pct=document.getElementById("plus2Pct");
