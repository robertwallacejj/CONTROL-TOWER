(function () {
  "use strict";

  let globalChartInstance = null;
  let statusChartInstance = null;
  let basesChartInstance = null;
  let currentRows = [];
  let currentPanelState = null;

  const REGIONAIS = {
    claudio: [
      "S-CRDR-SP", "GRU-SP", "S-CSVD-SP", "S-BRFD-SP", "S-FREG-SP", "F GRU-SP",
      "S-BRAS-SP", "F S-JRG-SP", "F S-VLMR-SP", "GRU 03-SP", "S-VLGUI-SP",
      "F S-BRSL-SP", "F S-BLV-SP"
    ],
    rodrigo: [
      "S-SAPOP-SP", "S-PENHA-SP", "S-MGUE-SP", "MGC-SP", "ARJ-SP", "SDR-SP",
      "S-SRAF-SP", "F ITQ-SP", "F S-PENHA-SP", "F S-PENHA 02-SP", "F S-MGUE-SP"
    ],
    neto: [
      "CARAP-SP", "CHM-SP", "COT-SP", "JDR-SP", "OSC-SP", "S-VLANA-SP",
      "S-VLLEO-SP", "S-VLSN-SP", "TBA-SP", "VRG-SP"
    ],
    luana: [
      "AME-SP", "FRCLR-SP", "F VCP-SP", "MGG-SP", "PIR-SP", "RCLR-SP", "SMR-SP",
      "VCP 03-SP", "VCP 05-SP", "VIN-SP", "FJND-SP", "ITUP-SP", "JND-SP",
      "BRG-SP", "CAIE-SP", "ATB-SP", "F SOD 02-SP", "IBUN-SP", "ITPT-SP",
      "ITPV-SP", "ITU-SP", "SOD02-SP", "SOD-SP", "SRQ-SP", "INDTR SD"
    ]
  };

  function formatNumber(value) {
    return Number(value || 0).toLocaleString("pt-BR");
  }

  function escapeHtml(str) {
    return String(str || "").replace(/[&<>"']/g, (s) => ({
      "&": "&amp;",
      "<": "&lt;",
      ">": "&gt;",
      "\"": "&quot;",
      "'": "&#39;"
    }[s]));
  }

  function escapeJs(str) {
    return String(str || "")
      .replace(/\\/g, "\\\\")
      .replace(/'/g, "\\'")
      .replace(/\r?\n/g, " ");
  }

  function sanitizeFilename(name) {
    return String(name || "base").replace(/[^a-z0-9_\-]/gi, "_");
  }

  function normalizar(txt) {
    return String(txt || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, "")
      .replace(/-/g, "")
      .toUpperCase();
  }

  function toNumber(value) {
    if (value === null || value === undefined || value === "") return 0;
    if (typeof value === "number") return Number.isFinite(value) ? value : 0;

    const normalized = String(value)
      .trim()
      .replace(/\./g, "")
      .replace(",", ".")
      .replace(/[^\d.-]/g, "");

    const num = Number(normalized);
    return Number.isFinite(num) ? num : 0;
  }

  function getField(row, keys) {
    for (const key of keys) {
      if (Object.prototype.hasOwnProperty.call(row, key)) {
        return row[key];
      }
    }

    const normalizedTargets = keys.map((k) => normalizar(k));
    for (const rowKey of Object.keys(row)) {
      if (normalizedTargets.includes(normalizar(rowKey))) {
        return row[rowKey];
      }
    }
    return null;
  }

  function getBaseName(row) {
    return String(getField(row, ["Base de entrega", "Base"]) || "BASE INDEFINIDA").trim();
  }

  function getDriverName(row) {
    return String(getField(row, ["Entregador", "Motorista", "Courier"]) || "NÃO ATRIBUÍDO").trim();
  }

  function inferStatus(row) {
    const deliveredTime = getField(row, ["Horário da entrega", "Horario da entrega"]);
    const failureReason = getField(row, [
      "Motivos dos pacotes problemáticos",
      "Motivos dos pacotes problematicos",
      "Pacote problemático",
      "Pacote problematico"
    ]);

    if (deliveredTime) return "delivered";
    if (failureReason) return "failure";
    return "pending";
  }

  function getRegionalFromBase(baseName) {
    const normalizedBase = normalizar(baseName);

    if (REGIONAIS.claudio.some((b) => normalizar(b) === normalizedBase)) return "Claudio";
    if (REGIONAIS.rodrigo.some((b) => normalizar(b) === normalizedBase)) return "Rodrigo";
    if (REGIONAIS.neto.some((b) => normalizar(b) === normalizedBase)) return "Neto";
    if (REGIONAIS.luana.some((b) => normalizar(b) === normalizedBase)) return "Luana";

    return "Não definida";
  }

  function isSummaryRow(row) {
    const explicitTotal = getField(row, [
      "Número total de expedido",
      "Numero total de expedido",
      "número total de expedido",
      "numero total de expedido",
      "Total Expedido"
    ]);

    return toNumber(explicitTotal) > 0;
  }

  function aggregateBaseMetrics(rows) {
    const grouped = {};

    rows.forEach((row) => {
      const base = getBaseName(row);

      if (!grouped[base]) {
        grouped[base] = {
          base,
          regional: String(getField(row, ["Regional"]) || getRegionalFromBase(base)),
          total: 0,
          entregue: 0,
          insucesso: 0,
          pendente: 0
        };
      }

      if (isSummaryRow(row)) {
        grouped[base].total += toNumber(getField(row, [
          "Número total de expedido",
          "Numero total de expedido",
          "número total de expedido",
          "numero total de expedido",
          "Total Expedido"
        ]));

        grouped[base].entregue += toNumber(getField(row, [
          "Número de pacotes assinados",
          "Numero de pacotes assinados",
          "Pacotes assinados",
          "Entregues"
        ]));

        grouped[base].pendente += toNumber(getField(row, [
          "Pacote não expedido",
          "Pacote nao expedido",
          "Não expedido",
          "Nao expedido",
          "Pendente"
        ]));

        grouped[base].insucesso += toNumber(getField(row, [
          "Pacote problemático",
          "Pacote problematico",
          "Não entregue",
          "Nao entregue",
          "Problemático"
        ]));
      } else {
        grouped[base].total += 1;
        const status = inferStatus(row);

        if (status === "delivered") grouped[base].entregue += 1;
        else if (status === "failure") grouped[base].insucesso += 1;
        else grouped[base].pendente += 1;
      }
    });

    return Object.values(grouped).map((item) => ({
      ...item,
      taxa: item.total > 0 ? (item.entregue / item.total) * 100 : 0
    }));
  }

  function aggregateGlobal(rows) {
    const bases = aggregateBaseMetrics(rows);
    return bases.reduce((acc, item) => {
      acc.total += item.total;
      acc.entregue += item.entregue;
      acc.insucesso += item.insucesso;
      acc.pendente += item.pendente;
      return acc;
    }, { total: 0, entregue: 0, insucesso: 0, pendente: 0 });
  }

  function createSampleRows() {
    const sample = [];
    const definitions = [
      { base: "VRG-SP", regional: "Neto", total: 40, delivered: 33, failure: 4, pending: 3 },
      { base: "AME-SP", regional: "Luana", total: 55, delivered: 49, failure: 3, pending: 3 },
      { base: "SMR-SP", regional: "Luana", total: 60, delivered: 51, failure: 5, pending: 4 },
      { base: "GRU-SP", regional: "Claudio", total: 35, delivered: 31, failure: 2, pending: 2 },
      { base: "SDR-SP", regional: "Rodrigo", total: 38, delivered: 26, failure: 8, pending: 4 },
      { base: "CHM-SP", regional: "Neto", total: 30, delivered: 24, failure: 3, pending: 3 }
    ];

    definitions.forEach((item) => {
      for (let i = 1; i <= item.total; i++) {
        const isDelivered = i <= item.delivered;
        const isFailure = !isDelivered && i <= item.delivered + item.failure;

        sample.push({
          "Base de entrega": item.base,
          "Entregador": `Motorista ${((i - 1) % 5) + 1}`,
          "Horário da entrega": isDelivered ? "10:30" : "",
          "Motivos dos pacotes problemáticos": isFailure ? "Cliente ausente" : "",
          "Regional": item.regional
        });
      }
    });

    return sample;
  }

  async function initLoginPage() {
    const loginForm = document.getElementById("loginForm");
    if (!loginForm) return;

    const loginMessage = document.getElementById("loginMessage");

    loginForm.addEventListener("submit", async (e) => {
      e.preventDefault();

      const email = document.getElementById("username").value.trim();
      const password = document.getElementById("password").value.trim();

      try {
        await CTSupabase.signIn(email, password);
        const profile = await CTSupabase.getProfile();

        if (profile && profile.role === "admin") {
          window.location.href = "admin.html";
        } else {
          window.location.href = "index.html";
        }
      } catch (error) {
        loginMessage.className = "message-box message-error";
        loginMessage.textContent = error.message || "Erro ao entrar.";
      }
    });
  }

  async function initAdminPage() {
    const auth = await CTSupabase.requireAdmin();
    if (!auth) return;

    const currentUser = await CTSupabase.getCurrentUserWithProfile();
    const adminUser = document.getElementById("adminUser");
    const goPanelBtn = document.getElementById("goPanelBtn");
    const logoutAdminBtn = document.getElementById("logoutAdminBtn");
    const importExcelBtn = document.getElementById("importExcelBtn");
    const loadSampleBtn = document.getElementById("loadSampleBtn");
    const clearDataBtn = document.getElementById("clearDataBtn");
    const messageBox = document.getElementById("adminMessage");

    if (!adminUser) return;

    adminUser.textContent =
      currentUser?.profile?.full_name ||
      currentUser?.profile?.email ||
      currentUser?.user?.email ||
      "Administrador";

    goPanelBtn.addEventListener("click", () => {
      window.location.href = "index.html";
    });

    logoutAdminBtn.addEventListener("click", async () => {
      await CTSupabase.signOut();
      window.location.href = "login.html";
    });

    importExcelBtn.addEventListener("click", async () => {
      const input = document.getElementById("excelFiles");
      const files = Array.from(input.files || []);

      if (!files.length) {
        messageBox.className = "message-box message-error";
        messageBox.textContent = "Selecione ao menos um arquivo Excel.";
        return;
      }

      if (typeof XLSX === "undefined") {
        messageBox.className = "message-box message-error";
        messageBox.textContent = "Biblioteca XLSX não carregada.";
        return;
      }

      try {
        let imported = [];

        for (const file of files) {
          const buffer = await file.arrayBuffer();
          const workbook = XLSX.read(buffer, { type: "array" });
          const sheet = workbook.Sheets[workbook.SheetNames[0]];
          const json = XLSX.utils.sheet_to_json(sheet, { defval: null });
          if (Array.isArray(json)) imported = imported.concat(json);
        }

        const userInfo = await CTSupabase.getCurrentUserWithProfile();
        await CTSupabase.savePanelState(imported, userInfo?.user?.email || null);
        await updateAdminStatus();

        messageBox.className = "message-box message-success";
        messageBox.textContent = `Importação concluída com sucesso. ${formatNumber(imported.length)} linhas carregadas.`;
      } catch (error) {
        messageBox.className = "message-box message-error";
        messageBox.textContent = `Erro ao importar Excel: ${error.message}`;
      }
    });

    loadSampleBtn.addEventListener("click", async () => {
      const rows = createSampleRows();
      const userInfo = await CTSupabase.getCurrentUserWithProfile();
      await CTSupabase.savePanelState(rows, userInfo?.user?.email || null);
      await updateAdminStatus();

      messageBox.className = "message-box message-success";
      messageBox.textContent = "Exemplo carregado com sucesso.";
    });

    clearDataBtn.addEventListener("click", async () => {
      const userInfo = await CTSupabase.getCurrentUserWithProfile();
      await CTSupabase.savePanelState([], userInfo?.user?.email || null);
      await updateAdminStatus();

      messageBox.className = "message-box message-success";
      messageBox.textContent = "Dados limpos com sucesso.";
    });

    await updateAdminStatus();
  }

  async function updateAdminStatus() {
    const state = await CTSupabase.loadPanelState();
    const rows = state.raw_rows || [];
    const baseMetrics = aggregateBaseMetrics(rows);
    const global = aggregateGlobal(rows);

    document.getElementById("adminLastUpdate").textContent =
      `Última atualização: ${CTSupabase.formatDateTimeBR(state.last_update)}`;

    document.getElementById("statusLastUpdate").textContent =
      `${CTSupabase.formatDateTimeBR(state.last_update)} (${state.updated_by_email || "--"})`;

    document.getElementById("statusRows").textContent = formatNumber(rows.length);
    document.getElementById("statusBases").textContent = formatNumber(baseMetrics.length);
    document.getElementById("statusTotal").textContent = formatNumber(global.total);
    document.getElementById("statusDelivered").textContent = formatNumber(global.entregue);
  }

  async function initIndexPage() {
    const session = await CTSupabase.requireAuth();
    if (!session) return;

    const currentUser = await CTSupabase.getCurrentUserWithProfile();
    const loggedUser = document.getElementById("loggedUser");
    const loggedRole = document.getElementById("loggedRole");

    if (!loggedUser || !loggedRole) return;

    loggedUser.textContent =
      currentUser?.profile?.full_name ||
      currentUser?.profile?.email ||
      currentUser?.user?.email ||
      "Usuário";

    loggedRole.textContent =
      currentUser?.profile?.role === "admin" ? "Administrador" : "Visualização";

    const goAdminBtn = document.getElementById("goAdminBtn");
    if (currentUser?.profile?.role !== "admin") {
      goAdminBtn.style.display = "none";
    } else {
      goAdminBtn.addEventListener("click", () => {
        window.location.href = "admin.html";
      });
    }

    document.getElementById("logoutBtn").addEventListener("click", async () => {
      await CTSupabase.signOut();
      window.location.href = "login.html";
    });

    initTabs();
    initGridSearch();
    await renderAll();

    setInterval(async () => {
      await renderAll();
    }, 1200000);
  }

  function initTabs() {
    const buttons = document.querySelectorAll(".tab-btn");
    const panels = document.querySelectorAll(".tab-panel");

    buttons.forEach((btn) => {
      btn.addEventListener("click", () => {
        const tabId = btn.getAttribute("data-tab");
        buttons.forEach((b) => b.classList.remove("active"));
        panels.forEach((p) => p.classList.remove("active"));
        btn.classList.add("active");
        document.getElementById(tabId).classList.add("active");
      });
    });
  }

  function initGridSearch() {
    const gridSelector = document.getElementById("gridSelector");
    const dashboardsContent = document.getElementById("dashboardsContent");
    const searchInput = document.getElementById("searchInput");
    const searchBtn = document.getElementById("searchBtn");
    const baseFilter = document.getElementById("baseFilter");

    if (gridSelector && dashboardsContent) {
      gridSelector.addEventListener("change", function () {
        const value = parseInt(this.value, 10) || 2;
        dashboardsContent.style.gridTemplateColumns = `repeat(${value}, 1fr)`;
        syncExpandButtons();
      });
    }

    if (searchInput) {
      searchInput.addEventListener("input", filterDashboardBlocks);
    }

    if (searchBtn) {
      searchBtn.addEventListener("click", filterDashboardBlocks);
    }

    if (baseFilter) {
      baseFilter.addEventListener("change", updateControlTowerView);
    }

    const downloadToggle = document.getElementById("downloadToggle");
    const downloadList = document.getElementById("downloadList");
    const selectAllDownloads = document.getElementById("selectAllDownloads");
    const downloadSelected = document.getElementById("downloadSelected");

    if (downloadToggle && downloadList) {
      downloadToggle.addEventListener("click", () => {
        downloadList.style.display = downloadList.style.display === "none" ? "block" : "none";
      });
    }

    if (selectAllDownloads) {
      selectAllDownloads.addEventListener("change", function () {
        document.querySelectorAll(".download-item").forEach((cb) => {
          cb.checked = this.checked;
        });
      });
    }

    if (downloadSelected) {
      downloadSelected.addEventListener("click", downloadSelectedImages);
    }
  }

  async function renderAll() {
    currentPanelState = await CTSupabase.loadPanelState();
    currentRows = currentPanelState.raw_rows || [];

    document.getElementById("lastUpdateText").textContent =
      `Última atualização: ${CTSupabase.formatDateTimeBR(currentPanelState.last_update)} - ${currentPanelState.updated_by_email || "--"}`;

    renderMonitorByBase();
    renderGlobalSummary();
    renderGlobalChart();
    populateDownloadList();
    populateBaseFilter();
    updateControlTowerView();
  }

  function renderMonitorByBase() {
    const dashboardsContent = document.getElementById("dashboardsContent");
    if (!dashboardsContent) return;

    const grouped = currentRows.reduce((acc, row) => {
      const base = getBaseName(row);
      if (!acc[base]) acc[base] = [];
      acc[base].push(row);
      return acc;
    }, {});

    dashboardsContent.innerHTML = "";

    Object.keys(grouped)
      .sort((a, b) => a.localeCompare(b, "pt-BR"))
      .forEach((baseName) => {
        dashboardsContent.insertAdjacentHTML("beforeend", buildBaseDashboardHtml(baseName, grouped[baseName]));
      });

    attachExpandHandlers();
    filterDashboardBlocks();
  }

  function buildBaseDashboardHtml(baseName, rows) {
    const totals = { total: 0, entregue: 0, insucesso: 0, pendente: 0 };
    const drivers = {};

    rows.forEach((row) => {
      if (isSummaryRow(row)) {
        const total = toNumber(getField(row, [
          "Número total de expedido",
          "Numero total de expedido",
          "número total de expedido",
          "numero total de expedido",
          "Total Expedido"
        ]));
        const delivered = toNumber(getField(row, [
          "Número de pacotes assinados",
          "Numero de pacotes assinados",
          "Pacotes assinados",
          "Entregues"
        ]));
        const pending = toNumber(getField(row, [
          "Pacote não expedido",
          "Pacote nao expedido",
          "Não expedido",
          "Nao expedido",
          "Pendente"
        ]));
        const failure = toNumber(getField(row, [
          "Pacote problemático",
          "Pacote problematico",
          "Não entregue",
          "Nao entregue",
          "Problemático"
        ]));

        const driver = getDriverName(row);
        if (!drivers[driver]) {
          drivers[driver] = { total: 0, entregue: 0, insucesso: 0, pendente: 0 };
        }

        totals.total += total;
        totals.entregue += delivered;
        totals.pendente += pending;
        totals.insucesso += failure;

        drivers[driver].total += total;
        drivers[driver].entregue += delivered;
        drivers[driver].pendente += pending;
        drivers[driver].insucesso += failure;
      } else {
        totals.total += 1;
        const driver = getDriverName(row);

        if (!drivers[driver]) {
          drivers[driver] = { total: 0, entregue: 0, insucesso: 0, pendente: 0 };
        }

        drivers[driver].total += 1;
        const status = inferStatus(row);

        if (status === "delivered") {
          totals.entregue += 1;
          drivers[driver].entregue += 1;
        } else if (status === "failure") {
          totals.insucesso += 1;
          drivers[driver].insucesso += 1;
        } else {
          totals.pendente += 1;
          drivers[driver].pendente += 1;
        }
      }
    });

    const efficiency = totals.total ? ((totals.entregue / totals.total) * 100).toFixed(1) : "0.0";
    const efficiencyColor =
      Number(efficiency) >= 90 ? "#16a34a" :
      Number(efficiency) >= 70 ? "#f59e0b" :
      "#d81f26";

    const safeId = `panel-${baseName.replace(/[^a-z0-9]/gi, "-").toLowerCase()}`;

    const driverRows = Object.entries(drivers)
      .sort((a, b) => b[1].total - a[1].total)
      .map(([name, d]) => {
        const pct = d.total ? ((d.entregue / d.total) * 100).toFixed(1) : "0.0";
        const statusClass = pct < 50 ? "status-bad" : pct < 90 ? "status-warn" : "status-ok";

        return `
          <tr>
            <td style="font-weight:700">${escapeHtml(name)}</td>
            <td class="t-center">${formatNumber(d.total)}</td>
            <td class="t-center">${formatNumber(d.entregue)}</td>
            <td class="t-center">${formatNumber(d.insucesso)}</td>
            <td class="t-right"><span class="${statusClass}">${pct}%</span></td>
          </tr>
        `;
      })
      .join("");

    return `
      <div class="card dashboard-block">
        <div id="${safeId}" class="capture-container">
          <div class="capture-header">
            <div>
              <div class="base-name-label">UNIDADE OPERACIONAL</div>
              <h2 class="base-name-value">${escapeHtml(baseName)}</h2>
            </div>
            <div class="base-meta">
              <div class="report-date">${new Date().toLocaleString("pt-BR")}</div>
              <div class="eficacia-pill">
                EFICÁCIA:
                <strong style="color:${efficiencyColor};margin-left:8px">${efficiency}%</strong>
              </div>
            </div>
          </div>

          <div class="kpi-grid">
            <div class="kpi-card kpi-highlight">
              <small>Pendente</small>
              <div class="kpi-value">${formatNumber(totals.pendente)}</div>
            </div>
            <div class="kpi-card">
              <small>Entregue</small>
              <div class="kpi-value">${formatNumber(totals.entregue)}</div>
            </div>
            <div class="kpi-card">
              <small>Insucesso</small>
              <div class="kpi-value">${formatNumber(totals.insucesso)}</div>
            </div>
            <div class="kpi-card">
              <small>Total</small>
              <div class="kpi-value">${formatNumber(totals.total)}</div>
            </div>
          </div>

          <div class="data-table-wrapper">
            <table>
              <thead>
                <tr>
                  <th>Motorista</th>
                  <th class="t-center">Carga</th>
                  <th class="t-center">Entregue</th>
                  <th class="t-center">Insucesso</th>
                  <th class="t-right">Eficácia</th>
                </tr>
              </thead>
              <tbody>${driverRows}</tbody>
            </table>
          </div>
        </div>

        <div class="action-area">
          <button class="btn-secondary" onclick="window.downloadReportPNG('${safeId}','${escapeJs(baseName)}')">📥 BAIXAR PNG</button>
          <button class="expand-btn" data-target="${safeId}">Expandir</button>
        </div>
      </div>
    `;
  }

  function filterDashboardBlocks() {
    const term = String(document.getElementById("searchInput")?.value || "").toLowerCase();
    document.querySelectorAll(".dashboard-block").forEach((block) => {
      const nameNode = block.querySelector(".base-name-value");
      const base = nameNode ? nameNode.textContent.toLowerCase() : "";
      block.style.display = base.includes(term) ? "" : "none";
    });
  }

  function renderGlobalSummary() {
    const global = aggregateGlobal(currentRows);
    const el = document.getElementById("globalSummary");
    const eficiencia = global.total ? ((global.entregue / global.total) * 100).toFixed(1) : "0.0";

    el.innerHTML = `
      TOTAL GERAL: ${formatNumber(global.total)}
      | ENTREGUE: ${formatNumber(global.entregue)}
      | INSUCESSO: ${formatNumber(global.insucesso)}
      | EFICÁCIA: ${eficiencia}%
    `;
  }

  function renderGlobalChart() {
    const canvas = document.getElementById("globalChart");
    if (!canvas || typeof Chart === "undefined") return;

    const global = aggregateGlobal(currentRows);

    const data = {
      labels: ["Entregue", "Insucesso", "Pendente"],
      datasets: [{
        data: [global.entregue, global.insucesso, global.pendente],
        backgroundColor: ["#16a34a", "#d81f26", "#f59e0b"]
      }]
    };

    if (globalChartInstance) {
      globalChartInstance.data = data;
      globalChartInstance.update();
      return;
    }

    globalChartInstance = new Chart(canvas.getContext("2d"), {
      type: "doughnut",
      data,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } }
      }
    });
  }

  function attachExpandHandlers() {
    const buttons = document.querySelectorAll(".expand-btn");

    buttons.forEach((btn) => {
      const targetId = btn.getAttribute("data-target");
      const container = document.getElementById(targetId);
      if (!container) return;

      const tableWrapper = container.querySelector(".data-table-wrapper");
      if (!tableWrapper) return;

      const gridSelector = document.getElementById("gridSelector");
      const gridCols = gridSelector ? parseInt(gridSelector.value, 10) || 2 : 2;

      if (gridCols >= 3) {
        tableWrapper.classList.add("collapsed");
        tableWrapper.classList.remove("expanded");
        btn.textContent = "Expandir";
      } else {
        tableWrapper.classList.remove("collapsed");
        tableWrapper.classList.add("expanded");
        btn.textContent = "Recolher";
      }

      btn.onclick = () => {
        const isCollapsed = tableWrapper.classList.toggle("collapsed");
        if (isCollapsed) {
          tableWrapper.classList.remove("expanded");
          btn.textContent = "Expandir";
        } else {
          tableWrapper.classList.add("expanded");
          btn.textContent = "Recolher";
        }
      };
    });
  }

  function syncExpandButtons() {
    const gridSelector = document.getElementById("gridSelector");
    const cols = gridSelector ? parseInt(gridSelector.value, 10) || 2 : 2;

    document.querySelectorAll(".dashboard-block .data-table-wrapper").forEach((wrapper) => {
      const button = wrapper.parentElement.nextElementSibling?.querySelector(".expand-btn");

      if (cols >= 3) {
        wrapper.classList.add("collapsed");
        wrapper.classList.remove("expanded");
        if (button) button.textContent = "Expandir";
      } else {
        wrapper.classList.remove("collapsed");
        wrapper.classList.add("expanded");
        if (button) button.textContent = "Recolher";
      }
    });
  }

  async function downloadReportPNG(elementId, baseName) {
    const el = document.getElementById(elementId);
    if (!el || typeof html2canvas === "undefined") return;

    try {
      const canvas = await html2canvas(el, { scale: 2, backgroundColor: "#ffffff" });
      const link = document.createElement("a");
      link.download = `OperationalReport_${sanitizeFilename(baseName)}.png`;
      link.href = canvas.toDataURL("image/png");
      link.click();
    } catch (error) {
      console.error("Erro exportando PNG", error);
    }
  }

  function populateDownloadList() {
    const container = document.getElementById("downloadItems");
    if (!container) return;

    container.innerHTML = "";
    const panels = document.querySelectorAll(".capture-container");

    panels.forEach((panel) => {
      const id = panel.id;
      const base = panel.querySelector(".base-name-value")?.textContent.trim() || "";
      const label = document.createElement("label");
      label.innerHTML = `<input type="checkbox" class="download-item" value="${id}"> ${escapeHtml(base)}`;
      container.appendChild(label);
    });

    const selectAll = document.getElementById("selectAllDownloads");
    if (selectAll) selectAll.checked = false;
  }

  async function downloadSelectedImages() {
    const checked = Array.from(document.querySelectorAll(".download-item:checked")).map((i) => i.value);

    if (!checked.length) {
      alert("Selecione ao menos uma base.");
      return;
    }

    for (const id of checked) {
      const el = document.getElementById(id);
      if (!el || typeof html2canvas === "undefined") continue;

      try {
        const canvas = await html2canvas(el, { scale: 2, backgroundColor: "#ffffff" });
        const link = document.createElement("a");
        const baseName = el.querySelector(".base-name-value")?.textContent.trim() || id;
        link.download = `OperationalReport_${sanitizeFilename(baseName)}.png`;
        link.href = canvas.toDataURL("image/png");
        link.click();
        await new Promise((resolve) => setTimeout(resolve, 200));
      } catch (error) {
        console.error("Erro ao gerar imagem", error);
      }
    }
  }

  function populateBaseFilter() {
    const baseFilter = document.getElementById("baseFilter");
    if (!baseFilter) return;

    const baseMetrics = aggregateBaseMetrics(currentRows)
      .map((item) => item.base)
      .sort((a, b) => a.localeCompare(b, "pt-BR"));

    const currentValue = baseFilter.value || "all";
    baseFilter.innerHTML = `<option value="all">Todas as Bases</option>`;

    baseMetrics.forEach((base) => {
      const option = document.createElement("option");
      option.value = base;
      option.textContent = base;
      baseFilter.appendChild(option);
    });

    if ([...baseFilter.options].some((o) => o.value === currentValue)) {
      baseFilter.value = currentValue;
    }
  }

  function updateControlTowerView() {
    const baseFilter = document.getElementById("baseFilter");
    if (!baseFilter) return;

    const selectedBase = baseFilter.value;
    const allMetrics = aggregateBaseMetrics(currentRows);
    const filteredMetrics = selectedBase === "all"
      ? allMetrics
      : allMetrics.filter((item) => item.base === selectedBase);

    const totalExpedido = filteredMetrics.reduce((sum, item) => sum + item.total, 0);
    const entregues = filteredMetrics.reduce((sum, item) => sum + item.entregue, 0);
    const pendente = filteredMetrics.reduce((sum, item) => sum + item.pendente, 0);
    const problematico = filteredMetrics.reduce((sum, item) => sum + item.insucesso, 0);
    const naoEntregue = problematico;
    const taxa = totalExpedido ? (entregues / totalExpedido) * 100 : 0;

    setText("totalExpedido", formatNumber(totalExpedido));
    setText("assinados", formatNumber(entregues));
    setText("naoExpedido", formatNumber(pendente));
    setText("problematico", formatNumber(problematico));
    setText("naoEntregue", formatNumber(naoEntregue));
    setText("taxa", `${taxa.toFixed(2)}%`);

    const taxaEl = document.getElementById("taxa");
    const taxaCard = document.getElementById("taxaCard");
    if (taxaCard) taxaCard.classList.remove("alert");

    if (taxaEl) {
      if (taxa >= 90) taxaEl.style.color = "#16a34a";
      else if (taxa >= 80) taxaEl.style.color = "#f59e0b";
      else {
        taxaEl.style.color = "#ef4444";
        if (taxaCard) taxaCard.classList.add("alert");
      }
    }

    renderStatusChart(entregues, naoEntregue, pendente, problematico);

    if (selectedBase === "all") {
      renderBasesChart(allMetrics);
      renderTowerRanking(allMetrics);
      renderRegionais(allMetrics);
    } else {
      clearTbody("topBases");
      clearTbody("worstBases");
      clearTbody("regionalClaudio");
      clearTbody("regionalRodrigo");
      clearTbody("regionalNeto");
      clearTbody("regionalLuana");
      destroyBasesChart();
    }
  }

  function setText(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
  }

  function clearTbody(id) {
    const el = document.getElementById(id);
    if (el) el.innerHTML = "";
  }

  function renderStatusChart(entregues, naoEntregue, pendente, problematico) {
    const canvas = document.getElementById("statusChart");
    if (!canvas || typeof Chart === "undefined") return;

    if (statusChartInstance) statusChartInstance.destroy();

    statusChartInstance = new Chart(canvas, {
      type: "doughnut",
      data: {
        labels: ["Entregue", "Não Entregue", "Pendente", "Problemático"],
        datasets: [{
          data: [entregues, naoEntregue, pendente, problematico],
          backgroundColor: ["#16a34a", "#ef4444", "#f59e0b", "#d81f26"]
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false
      }
    });
  }

  function renderBasesChart(metrics) {
    const canvas = document.getElementById("basesChart");
    if (!canvas || typeof Chart === "undefined") return;

    const sorted = [...metrics].sort((a, b) => b.taxa - a.taxa);
    const labels = sorted.map((item) => item.base);
    const values = sorted.map((item) => Number(item.taxa.toFixed(2)));

    if (basesChartInstance) basesChartInstance.destroy();

    basesChartInstance = new Chart(canvas, {
      type: "bar",
      data: {
        labels,
        datasets: [{
          data: values,
          backgroundColor: "#d81f26"
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false
      }
    });
  }

  function destroyBasesChart() {
    if (basesChartInstance) {
      basesChartInstance.destroy();
      basesChartInstance = null;
    }
  }

  function renderTowerRanking(metrics) {
    const top = [...metrics].sort((a, b) => b.taxa - a.taxa).slice(0, 10);
    const worst = [...metrics].sort((a, b) => a.taxa - b.taxa).slice(0, 10);

    document.getElementById("topBases").innerHTML = top.map((item) => `
      <tr>
        <td>${escapeHtml(item.base)}</td>
        <td class="t-right">${item.taxa.toFixed(2)}%</td>
      </tr>
    `).join("");

    document.getElementById("worstBases").innerHTML = worst.map((item) => `
      <tr>
        <td>${escapeHtml(item.base)}</td>
        <td class="t-right">${item.taxa.toFixed(2)}%</td>
      </tr>
    `).join("");
  }

  function renderRegionais(metrics) {
    renderRegionalTable(REGIONAIS.claudio, "regionalClaudio", metrics);
    renderRegionalTable(REGIONAIS.rodrigo, "regionalRodrigo", metrics);
    renderRegionalTable(REGIONAIS.neto, "regionalNeto", metrics);
    renderRegionalTable(REGIONAIS.luana, "regionalLuana", metrics);
  }

  function renderRegionalTable(baseList, targetId, metrics) {
    const normalizedList = baseList.map((b) => normalizar(b));
    const rows = metrics
      .filter((item) => normalizedList.includes(normalizar(item.base)))
      .sort((a, b) => b.taxa - a.taxa);

    const tbody = document.getElementById(targetId);
    if (!tbody) return;

    tbody.innerHTML = rows.map((item) => {
      const color = item.taxa >= 90 ? "#16a34a" : item.taxa >= 80 ? "#f59e0b" : "#ef4444";
      return `
        <tr>
          <td>${escapeHtml(item.base)}</td>
          <td class="t-right" style="color:${color}">${formatNumber(item.total)}</td>
          <td class="t-right" style="color:${color}">${formatNumber(item.entregue)}</td>
          <td class="t-right">${formatNumber(item.pendente)}</td>
          <td class="t-right">${formatNumber(item.insucesso)}</td>
          <td class="t-right" style="color:${color}">${item.taxa.toFixed(2)}%</td>
        </tr>
      `;
    }).join("");
  }

  function boot() {
    if (!window.CTSupabase) {
      console.error("CTSupabase não foi carregado.");
      return;
    }

    if (document.getElementById("loginForm")) initLoginPage();
    if (document.getElementById("importExcelBtn")) initAdminPage();
    if (document.getElementById("dashboardsContent")) initIndexPage();

    window.downloadReportPNG = downloadReportPNG;
  }

  document.addEventListener("DOMContentLoaded", boot);
})();