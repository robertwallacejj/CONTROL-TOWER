(function () {
  "use strict";

  let globalChartInstance = null;
  let statusChartInstance = null;
  let basesChartInstance = null;

  function destroyLocal(instance) {
    if (instance && typeof instance.destroy === "function") instance.destroy();
  }

  function renderGlobalChart(canvas, summary) {
    if (!canvas || typeof Chart === "undefined") return;
    destroyLocal(globalChartInstance);
    globalChartInstance = new Chart(canvas, {
      type: "doughnut",
      data: {
        labels: ["Entregue", "Insucesso", "Pendente"],
        datasets: [{ data: [summary.entregue, summary.insucesso, summary.pendente], backgroundColor: ["#16a34a", "#ef4444", "#f59e0b"] }]
      },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: "bottom" } } }
    });
  }

  function renderStatusChart(canvas, totals) {
    if (!canvas || typeof Chart === "undefined") return;
    destroyLocal(statusChartInstance);
    statusChartInstance = new Chart(canvas, {
      type: "doughnut",
      data: {
        labels: ["Entregue", "Não Entregue", "Pendente", "Problemático"],
        datasets: [{ data: [totals.entregue, totals.naoEntregue, totals.pendente, totals.problematico], backgroundColor: ["#16a34a", "#ef4444", "#f59e0b", "#d81f26"] }]
      },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: "bottom" } } }
    });
  }

  function renderBasesChart(canvas, metrics, sortMode) {
    if (!canvas || typeof Chart === "undefined") return;
    destroyLocal(basesChartInstance);
    const sorted = [...metrics].sort(function (a, b) {
      return sortMode === "asc" ? a.taxa - b.taxa : b.taxa - a.taxa;
    });
    basesChartInstance = new Chart(canvas, {
      type: "bar",
      data: {
        labels: sorted.map((item) => item.base),
        datasets: [{ label: "SLA", data: sorted.map((item) => Number(item.taxa.toFixed(2))), backgroundColor: "#d81f26" }]
      },
      options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { y: { beginAtZero: true, max: 100 } } }
    });
  }

  function destroyBasesChart() {
    destroyLocal(basesChartInstance);
    basesChartInstance = null;
  }

  window.CTCharts = { renderGlobalChart, renderStatusChart, renderBasesChart, destroyBasesChart };
})();
