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

generateBtn.onclick=generate;
exportBtn.onclick=exportExcel;

expandAllBtn.onclick=()=>toggleAll(true);
collapseAllBtn.onclick=()=>toggleAll(false);

clearSearch.onclick=()=>{ search.value=""; filter(""); };

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
  sizeSummaryBody.innerHTML="";

  let b0=0,b30=0,b60=0,b120=0;
  let b0u=0,b30u=0,b60u=0,b120u=0;

  const map={};
  const sizeAgg={};
  const sd=+salesDays.value;
  const target=+targetSC.value;

  sales.forEach(r=>{
    const {style,size}=normalizeSKU(r.SKU);
    map[style]??={sizes:{}};
    map[style].sizes[size]??={sales:0,stock:0};
    map[style].sizes[size].sales+=+r.Quantity;

    sizeAgg[size]??={sold:0,stock:0};
    sizeAgg[size].sold+=+r.Quantity;
  });

  stock.forEach(r=>{
    const {style,size}=normalizeSKU(r.SKU);
    map[style]??={sizes:{}};
    map[style].sizes[size]??={sales:0,stock:0};
    map[style].sizes[size].stock+=+r["Available Stock"];

    sizeAgg[size]??={sold:0,stock:0};
    sizeAgg[size].stock+=+r["Available Stock"];
  });

  let demandRows=[], overstockRows=[];

  Object.entries(map).forEach(([style,data])=>{
    let ts=0,tk=0;
    Object.values(data.sizes).forEach(v=>{ts+=v.sales;tk+=v.stock;});
    if(ts===0) return;

    const drr=ts/sd;
    const sc=tk/drr;
    const demand=sc<target?Math.ceil((target-sc)*drr):0;

    if(sc<30){b0++;b0u+=ts;}
    else if(sc<60){b30++;b30u+=ts;}
    else if(sc<120){b60++;b60u+=ts;}
    else {b120++;b120u+=ts;}

    if(demand>0) demandRows.push({style,data,ts,tk,drr,sc,demand});
    if(sc>120) overstockRows.push({style,data,ts,tk,drr,sc});
  });

  demandRows.sort((a,b)=>b.demand-a.demand);

  renderExpandable(demandRows,demandBody,true);
  renderExpandable(overstockRows,overstockBody,false);

  document.getElementById("b0").innerText=b0;
  document.getElementById("b30").innerText=b30;
  document.getElementById("b60").innerText=b60;
  document.getElementById("b120").innerText=b120;
  document.getElementById("b0u").innerText=b0u;
  document.getElementById("b30u").innerText=b30u;
  document.getElementById("b60u").innerText=b60u;
  document.getElementById("b120u").innerText=b120u;

  /* SIZE-WISE ANALYSIS SUMMARY */
  const totalSold=Object.values(sizeAgg).reduce((a,b)=>a+b.sold,0);
  const catTotals={Normal:0,"Plus 1":0,"Plus 2":0,"Free Size":0};

  SIZE_ORDER.forEach(s=>{
    if(s==="FS")catTotals["Free Size"]+=sizeAgg[s]?.sold||0;
    else if(NORMAL.has(s))catTotals["Normal"]+=sizeAgg[s]?.sold||0;
    else if(PLUS1.has(s))catTotals["Plus 1"]+=sizeAgg[s]?.sold||0;
    else if(PLUS2.has(s))catTotals["Plus 2"]+=sizeAgg[s]?.sold||0;
  });

  const printed={};
  SIZE_ORDER.forEach(s=>{
    const d=sizeAgg[s]||{sold:0,stock:0};
    let cat="Free Size";
    if(NORMAL.has(s))cat="Normal";
    else if(PLUS1.has(s))cat="Plus 1";
    else if(PLUS2.has(s))cat="Plus 2";

    const catShare=printed[cat]?"":(totalSold?((catTotals[cat]/totalSold)*100).toFixed(1)+"%":"0%");
    printed[cat]=true;

    sizeSummaryBody.insertAdjacentHTML("beforeend",`
      <tr>
        <td>${s}</td>
        <td>${cat}</td>
        <td>${d.sold}</td>
        <td>${totalSold?((d.sold/totalSold)*100).toFixed(1):"0"}%</td>
        <td>${catShare}</td>
        <td>${d.stock}</td>
      </tr>
    `);
  });

  /* SIZE CURVE */
  demandRows.forEach(r=>{
    let row={};
    SIZE_ORDER.forEach(s=>row[s]=0);
    Object.entries(r.data.sizes).forEach(([z,v])=>{
      row[z]=Math.round((v.sales/r.ts)*r.demand);
    });
    sizeCurveBody.insertAdjacentHTML("beforeend",`
      <tr data-style="${r.style.toLowerCase()}">
        <td>${r.style}</td>
        <td>${r.demand}</td>
        ${SIZE_ORDER.map(s=>`<td>${row[s]||""}</td>`).join("")}
      </tr>`);
  });
}

function renderExpandable(rows,tbody,isDemand){
  rows.forEach(r=>{
    const key="k"+Math.random().toString(36).slice(2);
    tbody.insertAdjacentHTML("beforeend",`
      <tr data-style="${r.style.toLowerCase()}">
        <td class="expand" onclick="toggle('${key}',this)">+</td>
        <td>${r.style}</td>
        <td>${r.ts}</td>
        <td>${r.tk}</td>
        <td>${r.drr.toFixed(2)}</td>
        <td>${r.sc.toFixed(1)}</td>
        ${isDemand?`<td>${r.demand}</td><td></td>`:""}
      </tr>`);

    Object.entries(r.data.sizes)
      .sort((a,b)=>SIZE_ORDER.indexOf(a[0])-SIZE_ORDER.indexOf(b[0]))
      .forEach(([z,v])=>{
        const drrS=v.sales/(+salesDays.value);
        const scS=drrS?v.stock/drrS:0;
        const dS=(isDemand&&scS<+targetSC.value)?Math.ceil((+targetSC.value-scS)*drrS):"";
        tbody.insertAdjacentHTML("beforeend",`
          <tr class="sub-row ${key}" style="display:none" data-style="${r.style.toLowerCase()}">
            <td></td>
            <td>${r.style}-${z}</td>
            <td>${v.sales}</td>
            <td>${v.stock}</td>
            <td>${drrS.toFixed(2)}</td>
            <td>${scS.toFixed(1)}</td>
            ${isDemand?`<td>${dS}</td><td></td>`:""}
          </tr>`);
      });
  });
}

window.toggle=(key,el)=>{
  const rows=document.querySelectorAll("."+key);
  const open=rows[0].style.display==="none";
  rows.forEach(r=>r.style.display=open?"":"none");
  el.textContent=open?"−":"+";
};

function toggleAll(open){
  document.querySelectorAll(".expand").forEach(el=>{
    const key=el.getAttribute("onclick").match(/'(.+?)'/)[1];
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
  XLSX.writeFile(wb,"Demand_Planner_v1_7.xlsx");
}

});
