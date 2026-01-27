document.addEventListener("DOMContentLoaded", () => {

/* ================= CONSTANTS ================= */
const SIZE_ORDER = ["FS","S","M","L","XL","XXL","3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"];

const CATEGORY = s =>
  s === "FS" ? "Free Size" :
  ["S","M","L","XL","XXL"].includes(s) ? "Normal" :
  ["3XL","4XL","5XL","6XL"].includes(s) ? "Plus 1" : "Plus 2";

/* ================= DOM ================= */
const salesFile = document.getElementById("salesFile");
const stockFile = document.getElementById("stockFile");
const salesDays = document.getElementById("salesDays");
const targetSC = document.getElementById("targetSC");

const generateBtn = document.getElementById("generateBtn");
const expandAllBtn = document.getElementById("expandAllBtn");
const collapseAllBtn = document.getElementById("collapseAllBtn");

const search = document.getElementById("search");
const clearSearch = document.getElementById("clearSearch");

const demandBody = document.querySelector("#demandTable tbody");
const overstockBody = document.querySelector("#overstockTable tbody");
const sizeCurveBody = document.querySelector("#sizeCurveTable tbody");
const brokenBody = document.querySelector("#brokenTable tbody");
const sizeSummaryBody = document.querySelector("#sizeSummaryTable tbody");

/* SC Band */
const b0=document.getElementById("b0"), b0u=document.getElementById("b0u");
const b30=document.getElementById("b30"), b30u=document.getElementById("b30u");
const b60=document.getElementById("b60"), b60u=document.getElementById("b60u");
const b120=document.getElementById("b120"), b120u=document.getElementById("b120u");

/* ================= TAB SWITCH ================= */
document.querySelectorAll(".tab-btn").forEach(btn=>{
  btn.onclick=()=>{
    document.querySelectorAll(".tab-btn").forEach(b=>b.classList.remove("active"));
    document.querySelectorAll(".tab-content").forEach(c=>c.classList.remove("active"));
    btn.classList.add("active");
    document.getElementById(btn.dataset.tab+"Tab").classList.add("active");
  };
});

/* ================= SEARCH ================= */
search.onkeyup=()=>{
  const q=search.value.toLowerCase();
  document.querySelectorAll("[data-style]").forEach(r=>{
    r.style.display=r.dataset.style.includes(q)?"":"none";
  });
};
clearSearch.onclick=()=>{search.value="";search.onkeyup();};

/* ================= FILE READ ================= */
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
const splitSKU=sku=>{
  if(!sku) return ["","FS"];
  const p=sku.split("-");
  return [p[0],p[1]||"FS"];
};

/* ================= GENERATE ================= */
generateBtn.onclick=()=>{
  if(!salesFile.files[0]||!stockFile.files[0]){
    alert("Upload Sales & Stock files");
    return;
  }
  Promise.all([
    readFile(salesFile.files[0]),
    readFile(stockFile.files[0]),
    fetch("data/sizes.xlsx")
      .then(r=>r.arrayBuffer())
      .then(b=>XLSX.read(b,{type:"array"}))
      .then(w=>XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]]))
  ]).then(([sales,stock,sizeMaster])=>{
    calculate(sales,stock,sizeMaster);
  });
};

/* ================= CALCULATE ================= */
function calculate(sales,stock,sizeMaster){

  /* RESET */
  demandBody.innerHTML="";
  overstockBody.innerHTML="";
  sizeCurveBody.innerHTML="";
  brokenBody.innerHTML="";
  sizeSummaryBody.innerHTML="";

  let bandCount={b0:0,b30:0,b60:0,b120:0};
  let bandUnits={b0:0,b30:0,b60:0,b120:0};

  const styleMap={}, sizeMap={}, sizeRef={};

  sizeMaster.forEach(r=>sizeRef[r["Style ID"]]=r["Total Sizes"]);

  /* SALES */
  sales.forEach(r=>{
    const [style,size]=splitSKU(r.SKU);
    sizeMap[size]??={sold:0,stock:0};
    sizeMap[size].sold+=+r.Quantity;

    styleMap[style]??={sold:0,stock:0,sizes:{}};
    styleMap[style].sold+=+r.Quantity;
    styleMap[style].sizes[size]=(styleMap[style].sizes[size]||0)+r.Quantity;
  });

  /* STOCK */
  stock.forEach(r=>{
    const [style,size]=splitSKU(r.SKU);
    sizeMap[size]??={sold:0,stock:0};
    sizeMap[size].stock+=+r["Available Stock"];

    styleMap[style]??={sold:0,stock:0,sizes:{}};
    styleMap[style].stock+=+r["Available Stock"];
  });

  /* SIZE SUMMARY */
  const totalSold=Object.values(sizeMap).reduce((a,b)=>a+b.sold,0);
  SIZE_ORDER.forEach(s=>{
    const d=sizeMap[s]||{sold:0,stock:0};
    sizeSummaryBody.insertAdjacentHTML("beforeend",`
      <tr>
        <td>${s}</td>
        <td>${CATEGORY(s)}</td>
        <td>${d.sold}</td>
        <td>${totalSold?((d.sold/totalSold)*100).toFixed(1):"0"}%</td>
        <td>${d.stock}</td>
      </tr>`);
  });

  /* DEMAND + OVERSTOCK + SIZE CURVE */
  Object.entries(styleMap).forEach(([style,d])=>{
    if(d.sold===0) return;

    const drr=d.sold/+salesDays.value;
    const sc=d.stock/drr;
    const demand=sc<+targetSC.value?Math.ceil((+targetSC.value-sc)*drr):0;

    /* SC BAND */
    if(sc<30){bandCount.b0++;bandUnits.b0+=d.sold;}
    else if(sc<60){bandCount.b30++;bandUnits.b30+=d.sold;}
    else if(sc<120){bandCount.b60++;bandUnits.b60+=d.sold;}
    else {bandCount.b120++;bandUnits.b120+=d.sold;}

    /* DEMAND */
    if(demand>0){
      demandBody.insertAdjacentHTML("beforeend",`
        <tr data-style="${style.toLowerCase()}">
          <td></td><td>${style}</td><td>${d.sold}</td><td>${d.stock}</td>
          <td>${drr.toFixed(2)}</td><td>${sc.toFixed(1)}</td><td>${demand}</td><td></td>
        </tr>`);
    }

    /* OVERSTOCK */
    if(sc>120){
      overstockBody.insertAdjacentHTML("beforeend",`
        <tr data-style="${style.toLowerCase()}">
          <td></td><td>${style}</td><td>${d.sold}</td><td>${d.stock}</td>
          <td>${drr.toFixed(2)}</td><td>${sc.toFixed(1)}</td>
        </tr>`);
    }

    /* SIZE CURVE */
    if(demand>0){
      const row={};
      SIZE_ORDER.forEach(s=>row[s]=0);
      Object.entries(d.sizes).forEach(([z,v])=>{
        row[z]=Math.round((v/d.sold)*demand);
      });
      sizeCurveBody.insertAdjacentHTML("beforeend",`
        <tr>
          <td>${style}</td><td>${demand}</td>
          ${SIZE_ORDER.map(s=>`<td>${row[s]||""}</td>`).join("")}
        </tr>`);
    }
  });

  /* UPDATE BAND TABLE */
  b0.innerText=bandCount.b0; b0u.innerText=bandUnits.b0;
  b30.innerText=bandCount.b30; b30u.innerText=bandUnits.b30;
  b60.innerText=bandCount.b60; b60u.innerText=bandUnits.b60;
  b120.innerText=bandCount.b120; b120u.innerText=bandUnits.b120;

  /* BROKEN SIZE */
  Object.entries(styleMap)
    .filter(([s,d])=>sizeRef[s] && d.sold>=30)
    .map(([s,d])=>{
      const broken=Object.entries(d.sizes)
        .filter(([z])=>(sizeMap[z]?.stock||0)<10)
        .map(([z])=>z);
      return {s,d,broken};
    })
    .filter(x=>x.broken.length>1)
    .sort((a,b)=>b.d.sold-a.d.sold||b.broken.length-a.broken.length)
    .forEach(x=>{
      brokenBody.insertAdjacentHTML("beforeend",`
        <tr>
          <td>${x.s}</td>
          <td>${sizeRef[x.s]}</td>
          <td>${x.broken.length}</td>
          <td>${x.broken.join(", ")}</td>
          <td>${x.d.sold}</td>
          <td>${x.d.stock}</td>
        </tr>`);
    });
}

});
