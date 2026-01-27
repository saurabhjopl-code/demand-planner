document.addEventListener("DOMContentLoaded", () => {

const SIZE_ORDER=["FS","S","M","L","XL","XXL","3XL","4XL","5XL","6XL","7XL","8XL","9XL","10XL"];
const CATEGORY = s =>
  s==="FS" ? "Free Size" :
  ["S","M","L","XL","XXL"].includes(s) ? "Normal" :
  ["3XL","4XL","5XL","6XL"].includes(s) ? "Plus 1" : "Plus 2";

const salesFile=document.getElementById("salesFile");
const stockFile=document.getElementById("stockFile");
const salesDays=document.getElementById("salesDays");

const generateBtn=document.getElementById("generateBtn");
const expandAllBtn=document.getElementById("expandAllBtn");
const collapseAllBtn=document.getElementById("collapseAllBtn");

const sizeSummaryBody=document.querySelector("#sizeSummaryTable tbody");
const brokenBody=document.querySelector("#brokenTable tbody");

generateBtn.onclick=generate;

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

function generate(){
  Promise.all([
    readFile(salesFile.files[0]),
    readFile(stockFile.files[0]),
    fetch("data/sizes.xlsx").then(r=>r.arrayBuffer())
      .then(b=>XLSX.read(b,{type:"array"}))
      .then(w=>XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]]))
  ]).then(([sales,stock,sizes])=>calculate(sales,stock,sizes));
}

function calculate(sales,stock,sizeMaster){
  const sizeMap={}, styleMap={}, sizeRef={};

  sizeMaster.forEach(r=>sizeRef[r["Style ID"]]=r["Total Sizes"]);

  sales.forEach(r=>{
    const [style,size="FS"]=r.SKU.split("-");
    sizeMap[size]??={sold:0,stock:0};
    sizeMap[size].sold+=+r.Quantity;
    styleMap[style]??={sold:0,stock:0,sizes:{}};
    styleMap[style].sold+=+r.Quantity;
    styleMap[style].sizes[size]??=0;
    styleMap[style].sizes[size]+=+r.Quantity;
  });

  stock.forEach(r=>{
    const [style,size="FS"]=r.SKU.split("-");
    sizeMap[size]??={sold:0,stock:0};
    sizeMap[size].stock+=+r["Available Stock"];
    styleMap[style]??={sold:0,stock:0,sizes:{}};
    styleMap[style].stock+=+r["Available Stock"];
  });

  // SIZE SUMMARY
  const totalSold=Object.values(sizeMap).reduce((a,b)=>a+b.sold,0);
  sizeSummaryBody.innerHTML="";
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

  // BROKEN SIZE REPORT
  brokenBody.innerHTML="";
  Object.entries(styleMap)
    .filter(([s,d])=>sizeRef[s] && d.sold>=30)
    .map(([s,d])=>{
      const broken=Object.entries(d.sizes)
        .filter(([z])=> (sizeMap[z]?.stock||0) <10)
        .map(([z])=>z);
      return {s,d,broken};
    })
    .filter(x=>x.broken.length>1)
    .sort((a,b)=>b.d.sold-a.d.sold || b.broken.length-a.broken.length)
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
