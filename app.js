document.addEventListener("DOMContentLoaded", () => {

const SIZE_ORDER=["S","M","L","XL","XXL","3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"];
const NORMAL=new Set(["S","M","L","XL","XXL"]);
const PLUS1=new Set(["3XL","4XL","5XL","6XL"]);
const PLUS2=new Set(["7XL","8XL","9XL","10XL"]);

const salesFile=salesFileEl=document.getElementById("salesFile");
const stockFile=document.getElementById("stockFile");
const salesDays=document.getElementById("salesDays");
const targetSC=document.getElementById("targetSC");
const generateBtn=document.getElementById("generateBtn");
const exportBtn=document.getElementById("exportBtn");
const search=document.getElementById("search");

const demandBody=document.querySelector("#demandTable tbody");
const overstockBody=document.querySelector("#overstockTable tbody");
const sizeCurveBody=document.querySelector("#sizeCurveTable tbody");

const normalSales=document.getElementById("normalSales");
const plus1Sales=document.getElementById("plus1Sales");
const plus2Sales=document.getElementById("plus2Sales");
const normalPct=document.getElementById("normalPct");
const plus1Pct=document.getElementById("plus1Pct");
const plus2Pct=document.getElementById("plus2Pct");

generateBtn.onclick=generate;
exportBtn.onclick=exportExcel;

document.querySelectorAll(".tab-btn").forEach(b=>{
  b.onclick=()=>{
    document.querySelectorAll(".tab-btn").forEach(x=>x.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach(x=>x.classList.remove("active"));
    b.classList.add("active");
    document.getElementById(b.dataset.tab+"Tab").classList.add("active");
  };
});

search.onkeyup=()=>{
  const q=search.value.toLowerCase();
  document.querySelectorAll("[data-style]").forEach(r=>{
    r.style.display=r.dataset.style.includes(q)?"":"none";
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
  sizeCurveBody.innerHTML="";

  let b0=0,b30=0,b60=0,b120=0;
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

  Object.entries(map).forEach(([style,data],i)=>{
    let ts=0,tk=0;
    Object.values(data.sizes).forEach(v=>{ts+=v.sales;tk+=v.stock;});
    if(ts===0) return;

    const drr=ts/sd;
    const sc=tk/drr;
    const demand=sc<target?Math.ceil((target-sc)*drr):0;

    if(sc<=30)b0++; else if(sc<=60)b30++; else if(sc<=120)b60++; else b120++;

    const styleRow=`
      <tr data-style="${style.toLowerCase()}">
        <td class="expand" onclick="toggle(${i})">+</td>
        <td>${style}</td><td>${ts}</td><td>${tk}</td>
        <td>${drr.toFixed(2)}</td><td>${sc.toFixed(1)}</td>
        ${demand>0?`<td>${demand}</td>`:`<td></td>`}
      </tr>`;
    if(demand>0) demandBody.insertAdjacentHTML("beforeend",styleRow);
    if(sc>120) overstockBody.insertAdjacentHTML("beforeend",styleRow.replace(`<td>${demand}</td>`,""));

    let curve=[];
    Object.entries(data.sizes)
      .sort((a,b)=>SIZE_ORDER.indexOf(a[0])-SIZE_ORDER.indexOf(b[0]))
      .forEach(([z,v])=>{
        if(v.sales>0) curve.push(`${z}:${Math.round((v.sales/ts)*demand)}`);
      });
    if(demand>0){
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

  const total=n+p1+p2;
  normalSales.innerText=n; plus1Sales.innerText=p1; plus2Sales.innerText=p2;
  normalPct.innerText=total?((n/total)*100).toFixed(1)+"%":"0%";
  plus1Pct.innerText=total?((p1/total)*100).toFixed(1)+"%":"0%";
  plus2Pct.innerText=total?((p2/total)*100).toFixed(1)+"%":"0%";
}

function exportExcel(){
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.table_to_sheet(document.getElementById("demandTable")),"Demand_Style");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.table_to_sheet(document.getElementById("overstockTable")),"Overstock_Style");
  XLSX.utils.book_append_sheet(wb,XLSX.utils.table_to_sheet(document.getElementById("sizeCurveTable")),"Size_Curve");
  XLSX.writeFile(wb,"Demand_Planner_V1_1.xlsx");
}

});
