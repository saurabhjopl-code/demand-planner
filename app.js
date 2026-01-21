document.addEventListener("DOMContentLoaded", () => {

  console.log("Demand Planner loaded");

  const salesFile = document.getElementById("salesFile");
  const stockFile = document.getElementById("stockFile");
  const generateBtn = document.getElementById("generateBtn");
  const exportBtn = document.getElementById("exportBtn");
  const search = document.getElementById("search");

  generateBtn.addEventListener("click", () => {
    console.log("Generate button clicked");
    processFiles();
  });

  exportBtn.addEventListener("click", exportFullReport);

  document.querySelectorAll(".tab-btn").forEach(btn => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".tab-btn").forEach(b => b.classList.remove("active"));
      document.querySelectorAll(".tab-content").forEach(c => c.classList.remove("active"));
      btn.classList.add("active");
      document.getElementById(btn.dataset.tab + "Tab").classList.add("active");
    });
  });

  search.addEventListener("keyup", () => {
    const q = search.value.toLowerCase();
    document.querySelectorAll("[data-style]").forEach(r => {
      r.style.display = r.dataset.style.toLowerCase().includes(q) ? "" : "none";
    });
  });

  function processFiles() {
    if (!salesFile.files.length || !stockFile.files.length) {
      alert("Please upload both Sales and Stock files");
      return;
    }
    alert("Files loaded. Processing started.");
  }

  function exportFullReport() {
    alert("Export will be added after confirmation");
  }

});
