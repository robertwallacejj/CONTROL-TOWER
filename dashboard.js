(function () {
  "use strict";

  const U = window.CTUtils;
  const Metrics = window.CTMetrics;
  const Charts = window.CTCharts;

  const state = { currentRows: [], panelState: null, filters: { regional: "all", base: "all", status: "all", search: "", target: 90 } };

  function getComparisonSummary(currentGlobal) {
    const previous = U.storageGet("ct_previous_snapshot_summary", null);
    if (!previous || !previous.total) {
      U.setText("compareBadge", "Sem comparação anterior");
      return;
    }
    const delta = currentGlobal.entregue - previous.entregue;
    const label = delta === 0 ? "Sem variação vs. anterior" : `${delta > 0 ? "+" : ""}${U.formatNumber(delta)} entregues vs. anterior`;
    U.setText("compareBadge", label);
  }

  function renderGlobalSummary(filteredMetrics, globalSummary) {
    const avgTax = globalSummary.total > 0 ? (globalSummary.entregue / globalSummary.total) * 100 : 0;
    U.setHtml("globalSummary", [
      `<strong>Total:</strong> ${U.formatNumber(globalSummary.total)}`,
      `<strong>Entregues:</strong> ${U.formatNumber(globalSummary.entregue)}`,
      `<strong>Insucesso:</strong> ${U.formatNumber(globalSummary.insucesso)}`,
      `<strong>Taxa:</strong> ${U.formatPercent(avgTax, 2)}`,
      `<strong>Bases visíveis:</strong> ${U.formatNumber(filteredMetrics.length)}`
    ].join(" &nbsp;•&nbsp; "));

    const globalCanvas = U.byId("globalChart");
    if (globalCanvas) Charts.renderGlobalChart(globalCanvas, globalSummary);
    getComparisonSummary(globalSummary);
  }

  function getFiltersFromUI() {
    state.filters.regional = U.byId("regionalFilter") ? U.byId("regionalFilter").value : "all";
    state.filters.base = U.byId("baseFilter") ? U.byId("baseFilter").value : "all";
    state.filters.status = U.byId("statusFilter") ? U.byId("statusFilter").value : "all";
    state.filters.search = U.byId("searchInput") ? U.byId("searchInput").value.trim() : "";
    state.filters.target = U.byId("metaSlaInput") ? Number(U.byId("metaSlaInput").value) || 90 : 90;
    return state.filters;
  }

  function getFilteredMetrics() {
    const metrics = Metrics.aggregateBaseMetrics(state.currentRows);
    return Metrics.filterMetrics(metrics, getFiltersFromUI());
  }

  function buildDriverRankingHtml(drivers) {
    if (!drivers.length) return `<tr><td colspan="4" class="text-soft">Sem dados detalhados de entregador.</td></tr>`;
    return drivers.sort(function (a, b) { return b.taxa - a.taxa; }).slice(0, 10).map(function (item) {
      return `<tr><td>${U.escapeHtml(item.driver)}</td><td>${U.escapeHtml(item.base)}</td><td class="t-right">${U.formatNumber(item.total)}</td><td class="t-right">${U.formatPercent(item.taxa, 2)}</td></tr>`;
    }).join("");
  }

  function buildCriticalBasesHtml(metrics, target) {
    const critical = metrics.filter(function (item) { return item.taxa < target; }).sort(function (a, b) {
      if (a.taxa !== b.taxa) return a.taxa - b.taxa;
      return b.insucesso - a.insucesso;
    }).slice(0, 5);

    if (!critical.length) return `<tr><td colspan="4" class="text-soft">Nenhuma base crítica no filtro atual.</td></tr>`;
    return critical.map(function (item) {
      return `<tr><td>${U.escapeHtml(item.base)}</td><td class="t-right text-danger">${U.formatPercent(item.taxa, 2)}</td><td class="t-right">${U.formatNumber(item.pendente)}</td><td class="t-right">${U.formatNumber(item.insucesso)}</td></tr>`;
    }).join("");
  }

  function buildBaseDashboardHtml(item, target, panelState) {
    const colorClass = item.taxa >= target ? "status-ok" : item.taxa >= target - 10 ? "status-warn" : "status-bad";
    const rowClass = item.taxa < target ? "capture-container critical" : "capture-container";
    return `
      <article class="${rowClass}" data-base="${U.escapeHtml(item.base)}">
        <header class="capture-header">
          <div>
            <div class="base-name-label">Base</div>
            <h3 class="base-name-value">${U.escapeHtml(item.base)}</h3>
            <div class="text-soft">${U.escapeHtml(item.regional)}</div>
          </div>
          <div class="base-meta">
            <div class="report-date">Atualização: ${U.escapeHtml(U.formatDateTimeBR(panelState && panelState.last_update))}</div>
            <div class="eficacia-pill ${colorClass}">SLA ${U.formatPercent(item.taxa, 2)}</div>
          </div>
        </header>

        <div class="kpi-grid">
          <div class="kpi-card"><small>Total Expedido</small><div class="kpi-value">${U.formatNumber(item.total)}</div></div>
          <div class="kpi-card"><small>Entregues</small><div class="kpi-value">${U.formatNumber(item.entregue)}</div></div>
          <div class="kpi-card"><small>Não Entregue</small><div class="kpi-value">${U.formatNumber(item.naoEntregue)}</div></div>
          <div class="kpi-card kpi-highlight"><small>Problemático</small><div class="kpi-value">${U.formatNumber(item.problematico)}</div></div>
        </div>

        <div class="kpi-grid">
          <div class="kpi-card"><small>Pendente</small><div class="kpi-value">${U.formatNumber(item.pendente)}</div></div>
          <div class="kpi-card"><small>Insucesso</small><div class="kpi-value">${U.formatNumber(item.insucesso)}</div></div>
          <div class="kpi-card"><small>Status da meta</small><div class="kpi-value">${item.taxa >= target ? "OK" : "Ação"}</div></div>
          <div class="kpi-card"><small>Regional</small><div class="kpi-value">${U.escapeHtml(item.regional)}</div></div>
        </div>

        <div class="action-area">
          <button class="btn-secondary" data-action="download" data-base="${U.escapeHtml(item.base)}">Baixar PNG</button>
        </div>
      </article>`;
  }

  function renderMonitorByBase(filteredMetrics) {
    const content = U.byId("dashboardsContent");
    const empty = U.byId("dashboardsEmpty");
    if (!content || !empty) return;
    if (!filteredMetrics.length) {
      content.innerHTML = "";
      empty.hidden = false;
      return;
    }
    empty.hidden = true;
    content.innerHTML = filteredMetrics.map(function (item) {
      return buildBaseDashboardHtml(item, state.filters.target, state.panelState);
    }).join("");

    U.qsa("[data-action='download']", content).forEach(function (button) {
      button.addEventListener("click", function () {
        downloadReportPNG(button.getAttribute("data-base"));
      });
    });
  }

  async function downloadReportPNG(baseName) {
    const selector = `[data-base="${CSS.escape(baseName)}"]`;
    const node = document.querySelector(selector);
    if (!node || typeof html2canvas === "undefined") return;
    const canvas = await html2canvas(node, { scale: 1.5, backgroundColor: "#ffffff" });
    const url = canvas.toDataURL("image/png");
    U.downloadBlobURL(url, `${U.sanitizeFilename(baseName)}.png`);
  }

  function populateDownloadList(metrics) {
    const wrap = U.byId("downloadItems");
    if (!wrap) return;
    wrap.innerHTML = metrics.map(function (item) {
      return `<label><input type="checkbox" value="${U.escapeHtml(item.base)}" checked> ${U.escapeHtml(item.base)}</label>`;
    }).join("");
  }

  async function downloadSelectedImages() {
    const checked = U.qsa("#downloadItems input[type='checkbox']:checked").map(function (input) {
      return input.value;
    });
    for (let i = 0; i < checked.length; i += 1) {
      await downloadReportPNG(checked[i]);
    }
  }

  function populateBaseFilter(metrics) {
    const baseFilter = U.byId("baseFilter");
    if (!baseFilter) return;
    const currentValue = baseFilter.value || "all";
    const bases = metrics.map(function (item) { return item.base; });
    baseFilter.innerHTML = `<option value="all">Todas as Bases</option>` + bases.map(function (base) {
      return `<option value="${U.escapeHtml(base)}">${U.escapeHtml(base)}</option>`;
    }).join("");
    if (bases.includes(currentValue)) baseFilter.value = currentValue;
  }

  function populateRegionalFilter(metrics) {
    const regionalFilter = U.byId("regionalFilter");
    if (!regionalFilter) return;
    const currentValue = regionalFilter.value || "all";
    const regionais = Array.from(new Set(metrics.map(function (item) { return item.regional || "Não definida"; }))).sort(function (a, b) {
      return a.localeCompare(b, "pt-BR");
    });
    regionalFilter.innerHTML = `<option value="all">Todas</option>` + regionais.map(function (regional) {
      return `<option value="${U.escapeHtml(regional)}">${U.escapeHtml(regional)}</option>`;
    }).join("");
    if (regionais.includes(currentValue)) regionalFilter.value = currentValue;
  }

  function renderTowerRanking(metrics, sortMode) {
    const top = [...metrics].sort(function (a, b) { return b.taxa - a.taxa; }).slice(0, 10);
    const worst = [...metrics].sort(function (a, b) { return a.taxa - b.taxa; }).slice(0, 10);
    U.setHtml("topBases", top.map(function (item) {
      return `<tr><td>${U.escapeHtml(item.base)}</td><td class="t-right">${U.formatPercent(item.taxa, 2)}</td></tr>`;
    }).join("") || `<tr><td colspan="2" class="text-soft">Sem dados</td></tr>`);
    U.setHtml("worstBases", worst.map(function (item) {
      return `<tr><td>${U.escapeHtml(item.base)}</td><td class="t-right">${U.formatPercent(item.taxa, 2)}</td></tr>`;
    }).join("") || `<tr><td colspan="2" class="text-soft">Sem dados</td></tr>`);
    const basesCanvas = U.byId("basesChart");
    if (metrics.length) Charts.renderBasesChart(basesCanvas, metrics, sortMode);
    else Charts.destroyBasesChart();
  }

  function renderRegionalTable(baseList, targetId, summaryId, metrics) {
    const allowed = baseList.map(U.normalizar);
    const rows = metrics.filter(function (item) { return allowed.includes(U.normalizar(item.base)); }).sort(function (a, b) { return b.taxa - a.taxa; });
    U.setHtml(targetId, rows.map(function (item) {
      const colorClass = item.taxa >= state.filters.target ? "text-success" : item.taxa >= state.filters.target - 10 ? "text-warning" : "text-danger";
      return `<tr><td>${U.escapeHtml(item.base)}</td><td class="t-right">${U.formatNumber(item.total)}</td><td class="t-right">${U.formatNumber(item.entregue)}</td><td class="t-right">${U.formatNumber(item.pendente)}</td><td class="t-right">${U.formatNumber(item.problematico)}</td><td class="t-right ${colorClass}">${U.formatPercent(item.taxa, 2)}</td></tr>`;
    }).join("") || `<tr><td colspan="6" class="text-soft">Sem bases no filtro atual.</td></tr>`);

    const summary = rows.reduce(function (acc, item) {
      acc.total += item.total;
      acc.entregue += item.entregue;
      return acc;
    }, { total: 0, entregue: 0 });
    const rate = summary.total ? (summary.entregue / summary.total) * 100 : 0;
    U.setText(summaryId, summary.total ? `Média regional: ${U.formatPercent(rate, 2)}` : "Sem dados");
  }

  function renderRegionais(metrics) {
    renderRegionalTable(Metrics.REGIONAIS.claudio, "regionalClaudio", "regionalSummaryClaudio", metrics);
    renderRegionalTable(Metrics.REGIONAIS.rodrigo, "regionalRodrigo", "regionalSummaryRodrigo", metrics);
    renderRegionalTable(Metrics.REGIONAIS.neto, "regionalNeto", "regionalSummaryNeto", metrics);
    renderRegionalTable(Metrics.REGIONAIS.luana, "regionalLuana", "regionalSummaryLuana", metrics);
  }

  function updateControlTowerView(filteredMetrics) {
    const selectedBase = U.byId("baseFilter") ? U.byId("baseFilter").value : "all";
    const selectedMetrics = selectedBase === "all" ? filteredMetrics : filteredMetrics.filter(function (item) { return item.base === selectedBase; });
    const totals = selectedMetrics.reduce(function (acc, item) {
      acc.total += item.total;
      acc.entregue += item.entregue;
      acc.problematico += item.problematico;
      acc.naoEntregue += item.naoEntregue;
      acc.pendente += item.pendente;
      acc.insucesso += item.insucesso;
      return acc;
    }, { total: 0, entregue: 0, problematico: 0, naoEntregue: 0, pendente: 0, insucesso: 0 });

    const taxa = totals.total ? (totals.entregue / totals.total) * 100 : 0;
    const target = state.filters.target;
    const taxaCard = U.byId("taxaCard");
    if (taxaCard) taxaCard.classList.toggle("alert", taxa < target);

    U.setText("totalExpedido", U.formatNumber(totals.total));
    U.setText("assinados", U.formatNumber(totals.entregue));
    U.setText("naoExpedido", U.formatNumber(totals.pendente));
    U.setText("problematico", U.formatNumber(totals.problematico));
    U.setText("naoEntregue", U.formatNumber(totals.naoEntregue));
    U.setText("taxa", U.formatPercent(taxa, 2));

    const belowTarget = filteredMetrics.filter(function (item) { return item.taxa < target; }).length;
    const aboveTarget = filteredMetrics.filter(function (item) { return item.taxa >= target; }).length;
    const failureRate = totals.total ? (totals.insucesso / totals.total) * 100 : 0;
    U.setText("belowTargetCount", U.formatNumber(belowTarget));
    U.setText("aboveTargetCount", U.formatNumber(aboveTarget));
    U.setText("pendingCurrentCount", U.formatNumber(totals.pendente));
    U.setText("failureRateText", U.formatPercent(failureRate, 2));

    U.setHtml("criticalBases", buildCriticalBasesHtml(filteredMetrics, target));
    const drivers = Metrics.aggregateDrivers(state.currentRows);
    const driverFiltered = drivers.filter(function (item) {
      const regional = state.filters.regional;
      const search = U.normalizar(state.filters.search || "");
      const matchesRegional = regional === "all" || Metrics.getRegionalFromBase(item.base) === regional;
      const matchesSearch = !search || U.normalizar(item.base).includes(search) || U.normalizar(item.driver).includes(search);
      const matchesBase = selectedBase === "all" || item.base === selectedBase;
      return matchesRegional && matchesSearch && matchesBase;
    });
    U.setHtml("driverRanking", buildDriverRankingHtml(driverFiltered));

    const statusCanvas = U.byId("statusChart");
    Charts.renderStatusChart(statusCanvas, { entregue: totals.entregue, naoEntregue: totals.naoEntregue, pendente: totals.pendente, problematico: totals.problematico });
    renderTowerRanking(filteredMetrics, U.byId("towerSort") ? U.byId("towerSort").value : "asc");
    renderRegionais(filteredMetrics);
  }

  async function renderAll() {
    state.panelState = await window.CTSupabase.loadPanelState();
    state.currentRows = U.safeArray(state.panelState.raw_rows);
    U.setText("lastUpdateText", `Última atualização: ${U.formatDateTimeBR(state.panelState.last_update)}`);
    const allMetrics = Metrics.aggregateBaseMetrics(state.currentRows);
    populateRegionalFilter(allMetrics);
    populateBaseFilter(allMetrics);
    const filteredMetrics = getFilteredMetrics();
    const globalSummary = filteredMetrics.reduce(function (acc, item) {
      acc.total += item.total; acc.entregue += item.entregue; acc.problematico += item.problematico; acc.naoEntregue += item.naoEntregue; acc.pendente += item.pendente; acc.insucesso += item.insucesso; return acc;
    }, { total: 0, entregue: 0, problematico: 0, naoEntregue: 0, pendente: 0, insucesso: 0 });
    renderMonitorByBase(filteredMetrics);
    renderGlobalSummary(filteredMetrics, globalSummary);
    populateDownloadList(filteredMetrics);
    updateControlTowerView(filteredMetrics);
  }

  function initTabs() {
    U.qsa(".tab-btn").forEach(function (button) {
      button.addEventListener("click", function () {
        U.qsa(".tab-btn").forEach(function (item) { item.classList.remove("active"); });
        U.qsa(".tab-panel").forEach(function (panel) { panel.classList.remove("active"); });
        button.classList.add("active");
        const panel = U.byId(button.dataset.tab);
        if (panel) panel.classList.add("active");
      });
    });
  }

  function initGridSearch() {
    const gridSelector = U.byId("gridSelector");
    const dashboards = U.byId("dashboardsContent");
    const refresh = U.debounce(function () { renderAll().catch(console.error); }, 200);

    if (gridSelector && dashboards) {
      gridSelector.addEventListener("change", function () {
        dashboards.style.gridTemplateColumns = `repeat(${gridSelector.value}, 1fr)`;
      });
      dashboards.style.gridTemplateColumns = `repeat(${gridSelector.value}, 1fr)`;
    }

    ["searchInput", "statusFilter", "regionalFilter", "baseFilter", "metaSlaInput", "towerSort"].forEach(function (id) {
      const el = U.byId(id);
      if (!el) return;
      el.addEventListener("input", refresh);
      el.addEventListener("change", refresh);
    });

    const searchBtn = U.byId("searchBtn");
    if (searchBtn) searchBtn.addEventListener("click", refresh);

    const toggle = U.byId("downloadToggle");
    const list = U.byId("downloadList");
    if (toggle && list) toggle.addEventListener("click", function () { list.hidden = !list.hidden; });

    const selectAll = U.byId("selectAllDownloads");
    if (selectAll) selectAll.addEventListener("change", function () {
      U.qsa("#downloadItems input[type='checkbox']").forEach(function (input) { input.checked = selectAll.checked; });
    });

    const downloadSelected = U.byId("downloadSelected");
    if (downloadSelected) downloadSelected.addEventListener("click", function () {
      downloadSelectedImages().catch(console.error);
    });
  }

  async function initIndexPage() {
    if (!window.CTSupabase) return;
    const auth = await window.CTSupabase.requireAuth();
    if (!auth) return;

    const current = await window.CTSupabase.getCurrentUserWithProfile();
    if (!current || !current.user) {
      window.location.href = "login.html";
      return;
    }

    const profile = current.profile || { role: "viewer", full_name: current.user.email, email: current.user.email };
    U.setText("loggedUser", profile.full_name || profile.email || current.user.email);
    U.setText("loggedRole", `Perfil ${profile.role || "viewer"}`);

    const goAdminBtn = U.byId("goAdminBtn");
    if (goAdminBtn) {
      goAdminBtn.hidden = profile.role !== "admin";
      goAdminBtn.addEventListener("click", function () { window.location.href = "admin.html"; });
    }

    const logoutBtn = U.byId("logoutBtn");
    if (logoutBtn) logoutBtn.addEventListener("click", async function () {
      await window.CTSupabase.signOut();
      window.location.href = "login.html";
    });

    initTabs();
    initGridSearch();
    await renderAll();
    setInterval(function () { renderAll().catch(function (error) { console.error("Falha na atualização visual automática.", error); }); }, 1200000);
  }

  window.CTDashboard = { initIndexPage, renderAll };
})();
