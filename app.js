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
const brokenBody=document.querySelector("#brokenTable tbody");

generateBtn.onclick=generate;
exportBtn.onclick=exportExcel;
expandAllBtn.onclick=()=>toggleAll(true);
collapseAllBtn.onclick=()=>toggleAll(false);
clearSearch.onclick=()=>{search.value="";filter("");};
search.onkeyup=()=>filter(search.value.toLowerCase());

document.querySelectorAll(".tab-btn").forEach(b=>{
  b.onclick=()=>{
    document.querySelectorAll(".tab-btn").forEach(x=>x.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach(x=>x.classList.remove("active"));
    b.classList.add("active");
    document.getElementById(b.dataset.tab+"Tab").classList.add("active");
  };
});

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

/* DEMAND / OVERSTOCK LOGIC — RESTORED FROM v1.6 */
function calculate(sales,stock){
  demandBody.innerHTML="";
  overstockBody.innerHTML="";
  sizeCurveBody.innerHTML="";

  const map={};
  const sd=+salesDays.value;
  const target=+targetSC.value;

  sales.forEach(r=>{
    const {style,size}=normalizeSKU(r.SKU);
    map[style]??={sizes:{}};
    map[style].sizes[size]??={sales:0,stock:0};
    map[style].sizes[size].sales+=+r.Quantity;
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
    const demand=sc<target?Math.ceil((target-sc)*drr):0;

    let cls="", remark="";
    if(ts>=50){
      if(drr>20 && sc<20){cls="green";remark="High Demand";}
      else if(drr>10 && drr<20 && sc<30 && sc>10){cls="amber";remark="Mid Demand";}
      else if(drr<10 && sc<10){cls="red";remark="Low Demand";}
    } else {
      remark="Low Sale Units";
    }

    if(demand>0){
      demandBody.insertAdjacentHTML("beforeend",`
        <tr class="${cls}" data-style="${style.toLowerCase()}">
          <td></td><td>${style}</td><td>${ts}</td><td>${tk}</td>
          <td>${drr.toFixed(2)}</td><td>${sc.toFixed(1)}</td>
          <td>${demand}</td><td>${remark}</td>
        </tr>`);
    }

    if(sc>120){
      overstockBody.insertAdjacentHTML("beforeend",`
        <tr data-style="${style.toLowerCase()}">
          <td></td><td>${style}</td><td>${ts}</td><td>${tk}</td>
          <td>${drr.toFixed(2)}</td><td>${sc.toFixed(1)}</td>
        </tr>`);
    }

    if(demand>0){
      let row={}; SIZE_ORDER.forEach(s=>row[s]=0);
      Object.entries(data.sizes).forEach(([z,v])=>{
        row[z]=Math.round((v.sales/ts)*demand);
      });
      sizeCurveBody.insertAdjacentHTML("beforeend",`
        <tr data-style="${style.toLowerCase()}">
          <td>${style}</td><td>${demand}</td>
          ${SIZE_ORDER.map(s=>`<td>${row[s]||""}</td>`).join("")}
        </tr>`);
    }
  });
}

/* BROKEN SIZE WITH NAMES */
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
      const brokenSizes=Object.entries(d.sizes).filter(([_,v])=>v.stock<10).map(([k])=>k);
      return {s,total:ref[s],broken:brokenSizes.length,names:brokenSizes.join(", "),sold:d.sold,stock:d.stock};
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
        </tr>`);
    });
}

function toggleAll(open){
  document.querySelectorAll(".expand").forEach(el=>{
    const key=el.getAttribute("onclick")?.match(/'(.+?)'/)?.[1];
    if(!key) return;
    const rows=document.querySelectorAll("."+key);
    rows.forEach(r=>r.style.display=open?"":"none");
    el.textContent=open?"−":"+";
  });
}

function exportExcel(){
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.table_to_sheet(document.getElementById("demandTable")),"Demand");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.table_to_sheet(document.getElementById("overstockTable")),"Overstock");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.table_to_sheet(document.getElementById("sizeCurveTable")),"Size_Curve");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.table_to_sheet(document.getElementById("brokenTable")),"Broken_Size");
  XLSX.writeFile(wb,"Demand_Planner_v1_8_1.xlsx");
}

});
