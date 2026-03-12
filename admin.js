(function () {
  "use strict";

  const U = window.CTUtils;
  const Excel = window.CTExcel;
  const Metrics = window.CTMetrics;
  const STORAGE_LAST_PANEL = "ct_last_panel_state_backup";
  let fileAnalyses = [];
  let selectedSheetsMap = {};

  function setPreviewCounters(report) {
    U.setText("previewFileCount", U.formatNumber(report.fileCount));
    U.setText("previewValidRows", U.formatNumber(report.validRows));
    U.setText("previewInvalidRows", U.formatNumber(report.invalidRows));
    U.setText("previewIgnoredSheets", U.formatNumber(report.ignoredSheets));
  }

  function buildFilePreviewCard(fileReport, analysis) {
    const missingBase = fileReport.validation.missing.base.join(", ");
    const missingDetailed = fileReport.validation.missing.detailed.join(", ");
    const missingSummary = fileReport.validation.missing.summary.join(", ");
    const statusBadge = fileReport.validation.isUsable ? `<span class="badge badge-success">Estrutura válida</span>` : `<span class="badge badge-danger">Estrutura incompleta</span>`;
    const sheetOptions = analysis.sheets.map(function (sheet) {
      const selected = sheet.sheetName === fileReport.selectedSheetName ? "selected" : "";
      const ignored = sheet.ignored ? " (ignorada automaticamente)" : "";
      return `<option value="${U.escapeHtml(sheet.sheetName)}" ${selected}>${U.escapeHtml(sheet.sheetName)}${ignored}</option>`;
    }).join("");

    return `<article class="preview-card" data-file="${U.escapeHtml(fileReport.fileName)}">
      <div class="preview-card-head">
        <div>
          <h4>${U.escapeHtml(fileReport.fileName)}</h4>
          <p class="text-soft">Linhas válidas: ${U.formatNumber(fileReport.normalizedRows.length)} • Linhas inválidas: ${U.formatNumber(fileReport.invalidRows.length)}</p>
        </div>
        ${statusBadge}
      </div>
      <div class="grid-2">
        <div class="field">
          <label>Aba selecionada</label>
          <select class="sheet-selector" data-file="${U.escapeHtml(fileReport.fileName)}">${sheetOptions}</select>
        </div>
        <div class="field">
          <label>Colunas detectadas</label>
          <div class="preview-chips">
            ${fileReport.validation.headers.length ? fileReport.validation.headers.map(function (header) { return `<span class="chip">${U.escapeHtml(header)}</span>`; }).join("") : `<span class="chip chip-danger">Nenhuma coluna detectada</span>`}
          </div>
        </div>
      </div>
      <div class="preview-missing">
        <div><strong>Base:</strong> ${missingBase || "OK"}</div>
        <div><strong>Detalhado:</strong> ${missingDetailed || "OK"}</div>
        <div><strong>Resumo:</strong> ${missingSummary || "OK"}</div>
      </div>
      <div class="table-scroll">
        <table>
          <thead><tr><th>Base</th><th>Entregador</th><th>Status</th><th class="t-right">Total</th></tr></thead>
          <tbody>
            ${fileReport.normalizedRows.slice(0, 8).map(function (row) {
              return `<tr><td>${U.escapeHtml(row.base)}</td><td>${U.escapeHtml(row.driver)}</td><td>${U.escapeHtml(row.isSummary ? "Resumo" : row.status)}</td><td class="t-right">${U.formatNumber(row.isSummary ? row.total : 1)}</td></tr>`;
            }).join("") || `<tr><td colspan="4" class="text-soft">Sem linhas utilizáveis nesta aba.</td></tr>`}
          </tbody>
        </table>
      </div>
    </article>`;
  }

  function bindSheetSelectors() {
    U.qsa(".sheet-selector").forEach(function (select) {
      select.addEventListener("change", function () {
        selectedSheetsMap[select.dataset.file] = select.value;
        renderImportPreview();
      });
    });
  }

  function renderImportPreview() {
    const previewWrap = U.byId("importPreviewWrap");
    const previewEmpty = U.byId("importPreviewEmpty");
    const badge = U.byId("importPreviewBadge");
    if (!previewWrap || !previewEmpty || !badge) return;

    if (!fileAnalyses.length) {
      previewWrap.hidden = true;
      previewEmpty.hidden = false;
      badge.textContent = "Nenhum arquivo selecionado";
      setPreviewCounters({ fileCount: 0, validRows: 0, invalidRows: 0, ignoredSheets: 0 });
      return;
    }

    const preview = Excel.buildPreviewReport(fileAnalyses, selectedSheetsMap);
    setPreviewCounters(preview.report);
    previewWrap.hidden = false;
    previewEmpty.hidden = true;
    badge.textContent = `${preview.report.fileCount} arquivo(s) prontos para importação`;
    U.setText("previewTimestamp", `Prévia gerada em ${U.formatDateTimeBR(new Date().toISOString())}`);
    previewWrap.innerHTML = preview.files.map(function (fileReport) {
      const analysis = fileAnalyses.find(function (item) { return item.fileName === fileReport.fileName; });
      return buildFilePreviewCard(fileReport, analysis);
    }).join("");
    bindSheetSelectors();
  }

  async function refreshPreviewFromInput() {
    const input = U.byId("excelFiles");
    const files = input && input.files ? input.files : [];
    if (!files.length) {
      fileAnalyses = [];
      selectedSheetsMap = {};
      renderImportPreview();
      return;
    }

    U.showMessage("adminMessage", "Lendo arquivos e validando abas...", "info");
    try {
      fileAnalyses = await Excel.inspectFiles(files);
      selectedSheetsMap = {};
      fileAnalyses.forEach(function (item) { selectedSheetsMap[item.fileName] = item.selectedSheetName; });
      renderImportPreview();
      U.showMessage("adminMessage", "Pré-visualização pronta. Revise as abas e as colunas antes de importar.", "success");
    } catch (error) {
      U.showMessage("adminMessage", error.message || "Falha ao ler os arquivos Excel.", "error");
    }
  }

  async function updateAdminStatus() {
    const panelState = await window.CTSupabase.loadPanelState();
    const rows = U.safeArray(panelState.raw_rows);
    const metrics = Metrics.aggregateBaseMetrics(rows);
    const global = Metrics.aggregateGlobal(rows);
    U.setText("adminLastUpdate", `Última atualização: ${U.formatDateTimeBR(panelState.last_update)}`);
    U.setText("statusLastUpdate", U.formatDateTimeBR(panelState.last_update));
    U.setText("statusRows", U.formatNumber(rows.length));
    U.setText("statusBases", U.formatNumber(metrics.length));
    U.setText("statusTotal", U.formatNumber(global.total));
    U.setText("statusDelivered", U.formatNumber(global.entregue));
    U.setText("statusUpdatedBy", panelState.updated_by_email || "--");
  }

  async function importExcel() {
    const preview = Excel.buildPreviewReport(fileAnalyses, selectedSheetsMap);
    const mode = U.byId("importMode") ? U.byId("importMode").value : "replace";
    const auth = await window.CTSupabase.getCurrentUserWithProfile();
    const userEmail = auth && auth.user ? auth.user.email : null;
    const invalidFiles = preview.files.filter(function (item) { return !item.validation.isUsable; });
    if (!preview.files.length) {
      U.showMessage("adminMessage", "Selecione pelo menos um arquivo para importar.", "warning");
      return;
    }
    if (invalidFiles.length) {
      U.showMessage("adminMessage", "Há arquivos com estrutura incompleta. Corrija as colunas destacadas antes de continuar.", "warning");
      return;
    }

    const importedRows = preview.files.flatMap(function (item) { return item.normalizedRows; });
    const previousState = await window.CTSupabase.loadPanelState();
    U.storageSet(STORAGE_LAST_PANEL, previousState);
    const mergedRows = Excel.mergeImportedRows(previousState.raw_rows, importedRows, mode);
    await window.CTSupabase.savePanelState(mergedRows, userEmail);
    U.storageSet("ct_previous_snapshot_summary", Metrics.aggregateGlobal(previousState.raw_rows));

    const metrics = Metrics.aggregateBaseMetrics(mergedRows);
    await window.CTSupabase.tryInsertImportLog([{ imported_at: new Date().toISOString(), imported_by_email: userEmail, import_mode: mode, file_count: preview.report.fileCount, valid_rows: preview.report.validRows, invalid_rows: preview.report.invalidRows }]);
    await window.CTSupabase.tryReplaceNormalizedRows(importedRows);
    await window.CTSupabase.tryReplaceMetrics(metrics);

    U.showMessage("adminMessage", `${mode === "append" ? "Dados adicionados" : "Dados substituídos"} com sucesso. ${preview.report.validRows} linhas válidas publicadas.`, "success");
    await updateAdminStatus();
  }

  async function loadSample() {
    const auth = await window.CTSupabase.getCurrentUserWithProfile();
    const userEmail = auth && auth.user ? auth.user.email : null;
    const previousState = await window.CTSupabase.loadPanelState();
    U.storageSet(STORAGE_LAST_PANEL, previousState);
    U.storageSet("ct_previous_snapshot_summary", Metrics.aggregateGlobal(previousState.raw_rows));
    const sample = Excel.createSampleRows();
    await window.CTSupabase.savePanelState(sample, userEmail);
    await updateAdminStatus();
    U.showMessage("adminMessage", "Dados de exemplo carregados com sucesso.", "success");
  }

  async function clearData() {
    const auth = await window.CTSupabase.getCurrentUserWithProfile();
    const userEmail = auth && auth.user ? auth.user.email : null;
    const previousState = await window.CTSupabase.loadPanelState();
    U.storageSet(STORAGE_LAST_PANEL, previousState);
    U.storageSet("ct_previous_snapshot_summary", Metrics.aggregateGlobal(previousState.raw_rows));
    await window.CTSupabase.savePanelState([], userEmail);
    await updateAdminStatus();
    U.showMessage("adminMessage", "Dados do painel removidos.", "success");
  }

  async function undoLastImport() {
    const backup = U.storageGet(STORAGE_LAST_PANEL, null);
    const auth = await window.CTSupabase.getCurrentUserWithProfile();
    const userEmail = auth && auth.user ? auth.user.email : null;
    if (!backup) {
      U.showMessage("adminMessage", "Nenhum backup local disponível para desfazer.", "warning");
      return;
    }
    await window.CTSupabase.savePanelState(U.safeArray(backup.raw_rows), userEmail);
    await updateAdminStatus();
    U.showMessage("adminMessage", "Última importação desfeita com sucesso.", "success");
  }

  async function initAdminPage() {
    if (!window.CTSupabase || !U.byId("importExcelBtn")) return;
    const auth = await window.CTSupabase.requireAdmin();
    if (!auth) return;

    U.setText("adminUser", auth.profile.full_name || auth.profile.email || "Administrador");
    U.setText("adminImportModeBadge", `Modo: ${U.byId("importMode").selectedOptions[0].textContent}`);

    U.byId("goPanelBtn").addEventListener("click", function () { window.location.href = "index.html"; });
    U.byId("logoutAdminBtn").addEventListener("click", async function () { await window.CTSupabase.signOut(); window.location.href = "login.html"; });
    U.byId("excelFiles").addEventListener("change", function () { refreshPreviewFromInput().catch(console.error); });
    U.byId("importMode").addEventListener("change", function () { U.setText("adminImportModeBadge", `Modo: ${U.byId("importMode").selectedOptions[0].textContent}`); });
    U.byId("importExcelBtn").addEventListener("click", function () { importExcel().catch(function (error) { console.error(error); U.showMessage("adminMessage", error.message || "Falha ao salvar a importação.", "error"); }); });
    U.byId("loadSampleBtn").addEventListener("click", function () { loadSample().catch(function (error) { console.error(error); U.showMessage("adminMessage", error.message || "Falha ao carregar exemplo.", "error"); }); });
    U.byId("clearDataBtn").addEventListener("click", function () { clearData().catch(function (error) { console.error(error); U.showMessage("adminMessage", error.message || "Falha ao limpar os dados.", "error"); }); });
    U.byId("undoImportBtn").addEventListener("click", function () { undoLastImport().catch(function (error) { console.error(error); U.showMessage("adminMessage", error.message || "Falha ao desfazer a última importação.", "error"); }); });

    await updateAdminStatus();
    renderImportPreview();
  }

  window.CTAdmin = { initAdminPage };
})();
