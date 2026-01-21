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

const normalSales=document.getElementById("normalSales");
const plus1Sales=document.getElementById("plus1Sales");
const plus2Sales=document.getElementById("plus2Sales");
const normalPct=document.getElementById("normalPct");
const plus1Pct=document.getElementById("plus1Pct");
const plus2Pct=document.getElementById("plus2Pct");

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

  let demandRows=[];
  let overstockRows=[];

  Object.entries(map).forEach(([style,data])=>{
    let ts=0,tk=0;
    Object.values(data.sizes).forEach(v=>{ts+=v.sales;tk+=v.stock;});
    if(ts===0) return;

    const drr=ts/sd;
    const sc=tk/drr;
    const demand=sc<target?Math.ceil((target-sc)*drr):0;

    if(sc<30)b0++; else if(sc<60)b30++; else if(sc<120)b60++; else b120++;

    /* -------- DEMAND COLOR LOGIC -------- */
    let dClass="", dRemark="";

    if(ts < 50){
      dRemark="Low Sale Units";
    } else {
      if(drr>20 && sc<20){
        dClass="green"; dRemark="High Demand";
      } else if(drr>10 && drr<20 && sc<30 && sc>10){
        dClass="amber"; dRemark="Mid Demand";
      } else if(drr<10 && sc<10){
        dClass="red"; dRemark="Low Demand";
      }
    }

    /* -------- OVERSTOCK COLOR LOGIC -------- */
    let oClass="", oRemark="";

    if(ts < 50 && sc > 90){
      oClass="red"; oRemark="High Risk";
    } else {
      if(drr>30){
        oClass="green"; oRemark="Low Risk";
      } else if(drr<30 && drr>10){
        oClass="amber"; oRemark="Mid Risk";
      } else if(drr<10){
        oClass="red"; oRemark="High Risk";
      }
    }

    if(demand>0) demandRows.push({style,data,ts,tk,drr,sc,demand,dClass,dRemark});
    if(sc>120) overstockRows.push({style,data,ts,tk,drr,sc,oClass,oRemark});
  });

  demandRows.sort((a,b)=>b.demand-a.demand);

  renderExpandable(demandRows,demandBody,true);
  renderExpandable(overstockRows,overstockBody,false);

  document.getElementById("b0").innerText=b0;
  document.getElementById("b30").innerText=b30;
  document.getElementById("b60").innerText=b60;
  document.getElementById("b120").innerText=b120;

  const total=n+p1+p2;
  normalSales.innerText=n;
  plus1Sales.innerText=p1;
  plus2Sales.innerText=p2;
  normalPct.innerText=total?((n/total)*100).toFixed(1)+"%":"0%";
  plus1Pct.innerText=total?((p1/total)*100).toFixed(1)+"%":"0%";
  plus2Pct.innerText=total?((p2/total)*100).toFixed(1)+"%":"0%";

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
    const cls=isDemand?r.dClass:r.oClass;
    const remark=isDemand?r.dRemark:r.oRemark;

    tbody.insertAdjacentHTML("beforeend",`
      <tr class="${cls}" data-style="${r.style.toLowerCase()}">
        <td class="expand" onclick="toggle('${key}',this)">+</td>
        <td>${r.style}</td>
        <td>${r.ts}</td>
        <td>${r.tk}</td>
        <td>${r.drr.toFixed(2)}</td>
        <td>${r.sc.toFixed(1)}</td>
        ${isDemand?`<td>${r.demand}</td><td>${remark||""}</td>`:""}
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
  XLSX.writeFile(wb,"Demand_Planner_v1_6.xlsx");
}

});
