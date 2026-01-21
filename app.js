document.addEventListener("DOMContentLoaded", () => {

const SIZE_ORDER = ["S","M","L","XL","XXL","3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"];
const NORMAL = new Set(["S","M","L","XL","XXL"]);
const PLUS1 = new Set(["3XL","4XL","5XL","6XL"]);
const PLUS2 = new Set(["7XL","8XL","9XL","10XL"]);

const salesFile = document.getElementById("salesFile");
const stockFile = document.getElementById("stockFile");
const salesDaysEl = document.getElementById("salesDays");
const targetSCEl = document.getElementById("targetSC");
const generateBtn = document.getElementById("generateBtn");
const exportBtn = document.getElementById("exportBtn");
const search = document.getElementById("search");

const demandBody = document.querySelector("#demandTable tbody");
const overstockBody = document.querySelector("#overstockTable tbody");
const sizeMixBody = document.querySelector("#sizeMixTable tbody");
const sizeCurveBody = document.querySelector("#sizeCurveTable tbody");

generateBtn.onclick = generate;
exportBtn.onclick = exportExcel;

document.querySelectorAll(".tab-btn").forEach(btn=>{
  btn.onclick=()=>{
    document.querySelectorAll(".tab-btn").forEach(b=>b.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach(c=>c.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById(btn.dataset.tab+"Tab").classList.add("active");
  };
});

search.onkeyup = ()=>{
  const q = search.value.toLowerCase();
  document.querySelectorAll("[data-style]").forEach(r=>{
    r.style.display = r.dataset.style.includes(q) ? "" : "none";
  });
};

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
  if(!sku) return {style:"",size:null};
  const p=sku.split("-");
  if(p.length<2||!p[1]||p[1]==="undefined") return {style:p[0],size:null};
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
  sizeMixBody.innerHTML="";
  sizeCurveBody.innerHTML="";

  let b0=b30=b60=b120=0;

  const map={};
  const sd=+salesDaysEl.value;
  const target=+targetSCEl.value;

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

  Object.entries(map).forEach(([style,data],i)=>{
    let ts=0, tk=0, n=0,p1=0,p2=0;
    Object.entries(data.sizes).forEach(([z,v])=>{
      ts+=v.sales; tk+=v.stock;
      if(NORMAL.has(z))n+=v.sales;
      if(PLUS1.has(z))p1+=v.sales;
      if(PLUS2.has(z))p2+=v.sales;
    });

    if(ts===0) return;

    const drr=ts/sd;
    const sc=tk/drr;
    const demand=sc<target?Math.ceil((target-sc)*drr):0;

    if(sc<=30)b0++; else if(sc<=60)b30++; else if(sc<=120)b60++; else b120++;

    if(demand>0){
      demandBody.insertAdjacentHTML("beforeend",`
        <tr data-style="${style.toLowerCase()}">
          <td class="expand" onclick="this.parentElement.nextSibling.classList.toggle('sub')">+</td>
          <td>${style}</td><td>${ts}</td><td>${tk}</td>
          <td>${drr.toFixed(2)}</td><td>${sc.toFixed(1)}</td><td>${demand}</td>
        </tr>`);
    }

    if(sc>120){
      overstockBody.insertAdjacentHTML("beforeend",`
        <tr data-style="${style.toLowerCase()}">
          <td></td><td>${style}</td><td>${ts}</td><td>${tk}</td>
          <td>${drr.toFixed(2)}</td><td>${sc.toFixed(1)}</td>
        </tr>`);
    }

    if(ts>=20){
      sizeMixBody.insertAdjacentHTML("beforeend",`
        <tr data-style="${style.toLowerCase()}">
          <td>${style}</td><td>${ts}</td>
          <td>${((n/ts)*100).toFixed(1)}%</td>
          <td>${((p1/ts)*100).toFixed(1)}%</td>
          <td>${((p2/ts)*100).toFixed(1)}%</td>
        </tr>`);
    }

    if(demand>0){
      let curve=[];
      Object.entries(data.sizes)
        .sort((a,b)=>SIZE_ORDER.indexOf(a[0])-SIZE_ORDER.indexOf(b[0]))
        .forEach(([z,v])=>{
          if(v.sales>0){
            curve.push(`${z}:${Math.round((v.sales/ts)*demand)}`);
          }
        });
      sizeCurveBody.insertAdjacentHTML("beforeend",`
        <tr data-style="${style.toLowerCase()}">
          <td>${style}</td><td>${demand}</td><td>${curve.join(", ")}</td>
        </tr>`);
    }
  });

  document.getElementById("b0").innerText=b0;
  document.getElementById("b30").innerText=b30;
  document.getElementById("b60").innerText=b60;
  document.getElementById("b120").innerText=b120;
}

function exportExcel(){
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.table_to_sheet(document.getElementById("demandTable")),"Demand");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.table_to_sheet(document.getElementById("overstockTable")),"Overstock");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.table_to_sheet(document.getElementById("sizeMixTable")),"Size Mix");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.table_to_sheet(document.getElementById("sizeCurveTable")),"Size Curve");
  XLSX.writeFile(wb,"Demand_Planning_Report.xlsx");
}

});
