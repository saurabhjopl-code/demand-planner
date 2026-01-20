function readFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = e => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet, { defval: 0 });
      resolve(json);
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function extractStyle(sku) {
  return sku.split('-')[0];
}

async function processFiles() {
  const salesFile = document.getElementById("salesFile").files[0];
  const stockFile = document.getElementById("stockFile").files[0];
  const salesDays = Number(document.getElementById("salesDays").value);
  const targetSC = Number(document.getElementById("targetSC").value);

  if (!salesFile || !stockFile) {
    alert("Please upload both files");
    return;
  }

  const salesData = await readFile(salesFile);
  const stockData = await readFile(stockFile);

  const styleMap = {};

  // Process Sales
  salesData.forEach(row => {
    const sku = row["SKU"];
    const qty = Number(row["Quantity"]) || 0;
    if (!sku) return;

    const style = extractStyle(sku);
    if (!styleMap[style]) {
      styleMap[style] = { sales: 0, stock: 0 };
    }
    styleMap[style].sales += qty;
  });

  // Process Stock
  stockData.forEach(row => {
    const sku = row["SKU"];
    const stock = Number(row["Available Stock"]) || 0;
    if (!sku) return;

    const style = extractStyle(sku);
    if (!styleMap[style]) {
      styleMap[style] = { sales: 0, stock: 0 };
    }
    styleMap[style].stock += stock;
  });

  const tbody = document.querySelector("#resultTable tbody");
  tbody.innerHTML = "";

  Object.entries(styleMap).forEach(([style, data]) => {
    const drr = data.sales > 0 ? data.sales / salesDays : 0;
    const sc = drr > 0 ? data.stock / drr : 9999;

    if (sc <= 120) return;

    let demand = 0;
    if (sc < targetSC && drr > 0) {
      demand = Math.ceil((targetSC - sc) * drr);
    }

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${style}</td>
      <td>${data.sales}</td>
      <td>${data.stock}</td>
      <td>${drr.toFixed(2)}</td>
      <td>${sc >= 9999 ? "No Sales" : sc.toFixed(1)}</td>
      <td>${demand}</td>
    `;
    tbody.appendChild(tr);
  });
}
