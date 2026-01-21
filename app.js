document.addEventListener("DOMContentLoaded", () => {

const SIZE_ORDER=["FS","S","M","L","XL","XXL","3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"];
const NORMAL=new Set(["S","M","L","XL","XXL"]);
const PLUS1=new Set(["3XL","4XL","5XL","6XL"]);
const PLUS2=new Set(["7XL","8XL","8XL","9XL","10XL"]);

const salesFile=document.getElementById("salesFile");
const stockFile=document.getElementById("stockFile");
const salesDays=document.getElementById("salesDays");
const targetSC=document.getElementById("targetSC");

const generateBtn=document.getElementById("generateBtn");
const exportBtn=document.getElementById("exportBtn");
const expandAllBtn=document.getElementById("expandAllBtn");
const collapseAllBtn=document.getElementById("collapseAllBtn");

const search=document.getElementById("search");
const clearSearch=document.getElementById("clearSearch");

const demandBody=document.querySelector("#demandTable tbody");
const overstockBody=document.querySelector("#overstockTable tbody");
const sizeCurveBody=document.querySelector("#sizeCurveTable tbody");

const normalSales=document.getElementById("normalSales");
const plus1Sales=document.getElementById("plus1Sales");
const plus2Sales=document.getElementById("plus2Sales");
const normalPct=document.getElementById("normalPct");
const plus1Pct=document.getElementById("plus1Pct");
const plus2Pct=document.getElementById("plus2Pct");

/* SC BAND ELEMENTS */
const b0=document.getElementById("b0");
const b30=document.getElementById("b30");
const b60=document.getElementById("b60");
const b120=document.getElementById("b120");

const b0u=document.getElementById("b0u");
const b30u=document.getElementById("b30u");
const b60u=document.getElementById("b60u");
const b120u=document.getElementById("b120u");

generateBtn.onclick=generate;
exportBtn.onclick=exportExcel;
expandAllBtn.onclick=()=>toggleAll(true);
collapseAllBtn.onclick=()=>toggleAll(false);

clearSearch.onclick=()=>{
  search.value="";
  filter("");
};

document.querySelectorAll(".tab-btn").forEach(b=>{
  b.onclick=()=>{
    document.querySelectorAll(".tab-btn").forEach(x=>x.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach(x=>x.classList.remove("active"));
    b.classList.add("active");
    document.getElementById(b.dataset.tab+"Tab").classList.add("active");
  };
});

search.onkeyup=()=>filter(search.value.toLowerCase());

function filter(q){
  document.querySelectorAll("[data-style]").forEach(r=>{
    r.style.display=r.dataset.style.includes(q)?"":"none";
  });
}

function readFile(file){
  return new Promise(res=>{
    const fr=new FileReader();
    fr.onload=e=>{
      const wb=XLSX.read(e.target.result,{type:"array"});
      res(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:0}));
    };
    fr.readAsArrayBuffer(file);
  });
}

function normalizeSKU(sku){
  if(!sku) return {style:"",size:"FS"};
  const p=sku.split("-");
  if(p.length<2||!p[1]||p[1]==="undefined") return {style:p[0],size:"FS"};
  return {style:p[0],size:p[1]};
}

function generate(){
  if(!salesFile.files[0]||!stockFile.files[0]){
    alert("Upload both files");
    return;
  }
  Promise.all([readFile(salesFile.files[0]),readFile(stockFile.files[0])])
    .then(([sales,stock])=>calculate(sales,stock));
}

function calculate(sales,stock){
  demandBody.innerHTML="";
  overstockBody.innerHTML="";
  sizeCurveBody.innerHTML="";

  let bandCount={b0:0,b30:0,b60:0,b120:0};
  let bandUnits={b0:0,b30:0,b60:0,b120:0};

  let n=0,p1=0,p2=0;
  const map={};
  const sd=+salesDays.value;
  const target=+targetSC.value;

  sales.forEach(r=>{
    const {style,size}=normalizeSKU(r.SKU);
    map[style]??={sizes:{}};
    map[style].sizes[size]??={sales:0,stock:0};
    map[style].sizes[size].sales+=+r.Quantity;

    if(NORMAL.has(size))n+=+r.Quantity;
    if(PLUS1.has(size))p1+=+r.Quantity;
    if(PLUS2.has(size))p2+=+r.Quantity;
  });

  stock.forEach(r=>{
    const {style,size}=normalizeSKU(r.SKU);
    map[style]??={sizes:{}};
    map[style].sizes[size]??={sales:0,stock:0};
    map[style].sizes[size].stock+=+r["Available Stock"];
  });

  Object.entries(map).forEach(([style,data])=>{
    let ts=0,tk=0;
    Object.values(data.sizes).forEach(v=>{ts+=v.sales;tk+=v.stock;});
    if(ts===0) return;

    const drr=ts/sd;
    const sc=tk/drr;

    if(sc<30){ bandCount.b0++; bandUnits.b0+=ts; }
    else if(sc<60){ bandCount.b30++; bandUnits.b30+=ts; }
    else if(sc<120){ bandCount.b60++; bandUnits.b60+=ts; }
    else { bandCount.b120++; bandUnits.b120+=ts; }
  });

  /* UPDATE SUMMARY */
  b0.innerText=bandCount.b0;
  b30.innerText=bandCount.b30;
  b60.innerText=bandCount.b60;
  b120.innerText=bandCount.b120;

  b0u.innerText=bandUnits.b0;
  b30u.innerText=bandUnits.b30;
  b60u.innerText=bandUnits.b60;
  b120u.innerText=bandUnits.b120;

  const total=n+p1+p2;
  normalSales.innerText=n;
  plus1Sales.innerText=p1;
  plus2Sales.innerText=p2;
  normalPct.innerText=total?((n/total)*100).toFixed(1)+"%":"0%";
  plus1Pct.innerText=total?((p1/total)*100).toFixed(1)+"%":"0%";
  plus2Pct.innerText=total?((p2/total)*100).toFixed(1)+"%":"0%";
}

function toggleAll(open){
  document.querySelectorAll(".expand").forEach(el=>{
    const key=el.getAttribute("onclick").match(/'(.+?)'/)[1];
    document.querySelectorAll("."+key).forEach(r=>{
      r.style.display=open?"":"none";
    });
    el.textContent=open?"−":"+";
  });
}

function exportExcel(){
  const wb=XLSX.utils.book_new();
  XLSX.writeFile(wb,"Demand_Planner_v1_6.xlsx");
}

});
