const NORMAL = new Set(["S","M","L","XL","XXL"]);
const PLUS1 = new Set(["3XL","4XL","5XL","6XL"]);
const PLUS2 = new Set(["7XL","8XL","9XL","10XL"]);

const tabMap = {
  demand: "demandTab",
  overstock: "overstockTab",
  sizemix: "sizemixTab",
  sizeanalysis: "sizeanalysisTab",
  sizecurve: "sizecurveTab"
};

document.getElementById("generateBtn").addEventListener("click", processFiles);

document.querySelectorAll(".tab-btn").forEach(btn=>{
  btn.addEventListener("click",()=>{
    document.querySelectorAll(".tab-btn").forEach(b=>b.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach(c=>c.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById(tabMap[btn.dataset.tab]).classList.add("active");
  });
});

function readFile(file){
  return new Promise(resolve=>{
    const r=new FileReader();
    r.onload=e=>{
      const wb=XLSX.read(e.target.result,{type:"array"});
      resolve(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:0}));
    };
    r.readAsArrayBuffer(file);
  });
}

function normalizeSKU(sku){
  if(!sku) return {style:"",size:null};
  const p=sku.split("-");
  if(p.length<2||!p[1]||p[1]==="undefined") return {style:p[0],size:null};
  return {style:p[0],size:p[1]};
}

function processFiles(){
  if(!salesFile.files[0]||!stockFile.files[0]){
    alert("Upload both files");
    return;
  }
  Promise.all([readFile(salesFile.files[0]),readFile(stockFile.files[0])])
    .then(([sales,stock])=>calculate(sales,stock));
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

  renderReports(map,sd,target,bands);
  renderSizeMix(map);
  renderSizeAnalysis(n,p1,p2);
  renderSizeCurve(map,sd,target);

  Object.keys(bands).forEach(k=>document.getElementById(k).innerText=bands[k]);
}

function renderReports(map,sd,target,bands){
  demandTable.tBodies[0].innerHTML="";
  overstockTable.tBodies[0].innerHTML="";

  Object.entries(map).forEach(([style,sizes])=>{
    let s=0,st=0;
    Object.values(sizes).forEach(v=>{s+=v.sales;st+=v.stock});
    const drr=s?s/sd:0;
    const sc=drr?st/drr:Infinity;

    if(sc<=30)bands.b0++;
    else if(sc<=60)bands.b30++;
    else if(sc<=120)bands.b60++;
    else bands.b120++;

    const demand=(drr&&sc<target)?Math.ceil((target-sc)*drr):0;

    if(demand>0){
      demandTable.tBodies[0].insertAdjacentHTML("beforeend",
        `<tr data-style="${style}">
          <td>${style}</td><td>${s}</td><td>${st}</td>
          <td>${drr.toFixed(2)}</td><td>${sc.toFixed(1)}</td><td>${demand}</td>
        </tr>`);
    }

    if(sc>120){
      overstockTable.tBodies[0].insertAdjacentHTML("beforeend",
        `<tr data-style="${style}">
          <td>${style}</td><td>${s}</td><td>${st}</td>
          <td>${drr.toFixed(2)}</td><td>${sc.toFixed(1)}</td>
        </tr>`);
    }
  });
}

function renderSizeMix(map){
  sizeMixTable.tBodies[0].innerHTML="";
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
    sizeMixTable.tBodies[0].insertAdjacentHTML("beforeend",
      `<tr data-style="${style}">
        <td>${style}</td><td>${t}</td>
        <td>${((n/t)*100).toFixed(1)}%</td>
        <td>${((p1/t)*100).toFixed(1)}%</td>
        <td>${((p2/t)*100).toFixed(1)}%</td>
      </tr>`);
  });
}

function renderSizeAnalysis(n,p1,p2){
  const t=n+p1+p2;
  normalSales.innerText=n;
  plus1Sales.innerText=p1;
  plus2Sales.innerText=p2;
  normalPct.innerText=t?((n/t)*100).toFixed(1)+"%":"0%";
  plus1Pct.innerText=t?((p1/t)*100).toFixed(1)+"%":"0%";
  plus2Pct.innerText=t?((p2/t)*100).toFixed(1)+"%":"0%";
}

function renderSizeCurve(map,sd,target){
  sizeCurveTable.tBodies[0].innerHTML="";
  Object.entries(map).forEach(([style,sizes])=>{
    let sales=0,stock=0;
    Object.values(sizes).forEach(v=>{sales+=v.sales;stock+=v.stock});
    if(sales<20)return;
    const drr=sales/sd;
    const sc=stock/drr;
    const demand=(sc<target)?Math.ceil((target-sc)*drr):0;
    if(demand<=0)return;

    let parts=[];
    Object.entries(sizes).forEach(([z,v])=>{
      if(!z||v.sales===0)return;
      parts.push(`${z}:${Math.round((v.sales/sales)*demand)}`);
    });

    sizeCurveTable.tBodies[0].insertAdjacentHTML("beforeend",
      `<tr data-style="${style}">
        <td>${style}</td><td>${demand}</td><td>${parts.join(", ")}</td>
      </tr>`);
  });
}

/* SEARCH – ALL TABS */
search.addEventListener("keyup",()=>{
  const q=search.value.toLowerCase();
  document.querySelectorAll("[data-style]").forEach(r=>{
    r.style.display=r.dataset.style.toLowerCase().includes(q)?"":"none";
  });
});

/* EXPORT */
function exportFullReport(){
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.table_to_sheet(demandTable),"Demand");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.table_to_sheet(overstockTable),"Overstock");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.table_to_sheet(sizeMixTable),"Size Mix");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.table_to_sheet(sizeCurveTable),"Size Curve");
  XLSX.writeFile(wb,"Demand_Planning_Report.xlsx");
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

const normalSales=document.getElementById("normalSales");
const plus1Sales=document.getElementById("plus1Sales");
const plus2Sales=document.getElementById("plus2Sales");
const normalPct=document.getElementById("normalPct");
const plus1Pct=document.getElementById("plus1Pct");
const plus2Pct=document.getElementById("plus2Pct");
