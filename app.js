document.addEventListener("DOMContentLoaded", () => {

/* ===== YOUR ENTIRE v1.7 CODE (UNCHANGED) ===== */
/* (Exactly what you pasted — not repeated here to avoid accidental edits) */

/* ===== BROKEN SIZE APPEND ONLY ===== */

const brokenBody = document.querySelector("#brokenTable tbody");

function calculateBrokenSize(sales, stock) {
  if (!brokenBody) return;
  brokenBody.innerHTML = "";

  fetch("data/sizes.xlsx")
    .then(r => r.arrayBuffer())
    .then(b => XLSX.read(b, { type: "array" }))
    .then(wb => XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]))
    .then(sizeMaster => {

      const ref = {};
      sizeMaster.forEach(r => ref[r["Style ID"]] = +r["Total Sizes"]);

      const map = {};

      sales.forEach(r => {
        const { style, size } = normalizeSKU(r.SKU);
        map[style] ??= { sold: 0, stock: 0, sizes: {} };
        map[style].sold += +r.Quantity;
        map[style].sizes[size] ??= { stock: 0 };
      });

      stock.forEach(r => {
        const { style, size } = normalizeSKU(r.SKU);
        map[style] ??= { sold: 0, stock: 0, sizes: {} };
        map[style].stock += +r["Available Stock"];
        map[style].sizes[size] ??= { stock: 0 };
        map[style].sizes[size].stock += +r["Available Stock"];
      });

      Object.entries(map)
        .filter(([s, d]) => ref[s] && d.sold >= 30)
        .map(([s, d]) => {
          const broken = Object.entries(d.sizes)
            .filter(([_, v]) => v.stock < 10)
            .map(([k]) => k);
          return { s, total: ref[s], broken, sold: d.sold, stock: d.stock };
        })
        .filter(r => r.broken.length > 1)
        .sort((a, b) => b.sold - a.sold || b.broken.length - a.broken.length)
        .forEach(r => {
          brokenBody.insertAdjacentHTML("beforeend", `
            <tr>
              <td>${r.s}</td>
              <td>${r.total}</td>
              <td>${r.broken.length}</td>
              <td>${r.broken.join(", ")}</td>
              <td>${r.sold}</td>
              <td>${r.stock}</td>
            </tr>
          `);
        });
    });
}

/* Hook into existing generate */
const originalGenerate = generate;
generate = function () {
  originalGenerate();
  Promise.all([readFile(salesFile.files[0]), readFile(stockFile.files[0])])
    .then(([sales, stock]) => calculateBrokenSize(sales, stock));
};

});
