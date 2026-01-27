document.addEventListener("DOMContentLoaded", () => {

const SIZE_ORDER=["FS","S","M","L","XL","XXL","3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"];
const NORMAL=new Set(["S","M","L","XL","XXL"]);
const PLUS1=new Set(["3XL","4XL","5XL","6XL"]);
const PLUS2=new Set(["7XL","8XL","9XL","10XL"]);

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
const sizeSummaryBody=document.querySelector("#sizeSummaryTable tbody");
const brokenBody=document.querySelector("#brokenTable tbody");

generateBtn.onclick=generate;
exportBtn.onclick=exportExcel;
expandAllBtn.onclick=()=>toggleAll(true);
collapseAllBtn.onclick=()=>toggleAll(false);
clearSearch.onclick=()=>{search.value="";filter("");};

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
  Promise.all([
    readFile(salesFile.files[0]),
    readFile(stockFile.files[0]),
    fetch("data/sizes.xlsx").then(r=>r.ok?r.arrayBuffer():Promise.reject()).then(b=>XLSX.read(b,{type:"array"})).then(w=>XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]])).catch(()=>[])
  ]).then(([sales,stock,sizeMaster])=>{
    calculate(sales,stock);
    calculateBrokenSize(sales,stock,sizeMaster);
  });
}

/* ==== EXISTING CALCULATE (UNCHANGED) ==== */
/* ... YOUR calculate(), renderExpandable(), toggle(), toggleAll(), exportExcel()
   ARE IDENTICAL TO WHAT YOU PASTED ABOVE ...
   (intentionally not rewritten to avoid accidental breakage) */

/* ==== BROKEN SIZE (APPENDED) ==== */
function calculateBrokenSize(sales,stock,sizeMaster){
  brokenBody.innerHTML="";
  const ref={};
  sizeMaster.forEach(r=>ref[r["Style ID"]]=+r["Total Sizes"]);
  const map={};

  sales.forEach(r=>{
    const {style,size}=normalizeSKU(r.SKU);
    map[style]??={sold:0,stock:0,sizes:{}};
    map[style].sold+=+r.Quantity;
    map[style].sizes[size]??={stock:0};
  });

  stock.forEach(r=>{
    const {style,size}=normalizeSKU(r.SKU);
    map[style]??={sold:0,stock:0,sizes:{}};
    map[style].stock+=+r["Available Stock"];
    map[style].sizes[size]??={stock:0};
    map[style].sizes[size].stock+=+r["Available Stock"];
  });

  Object.entries(map)
    .filter(([s,d])=>ref[s] && d.sold>=30)
    .map(([s,d])=>{
      const brokenNames=Object.entries(d.sizes)
        .filter(([_,v])=>v.stock<10)
        .map(([size])=>size);
      return {s,total:ref[s],broken:brokenNames.length,names:brokenNames.join(", "),sold:d.sold,stock:d.stock};
    })
    .filter(r=>r.broken>1)
    .sort((a,b)=>b.sold-a.sold||b.broken-a.broken)
    .forEach(r=>{
      brokenBody.insertAdjacentHTML("beforeend",`
        <tr>
          <td>${r.s}</td>
          <td>${r.total}</td>
          <td>${r.broken}</td>
          <td>${r.names}</td>
          <td>${r.sold}</td>
          <td>${r.stock}</td>
        </tr>
      `);
    });
}

});
