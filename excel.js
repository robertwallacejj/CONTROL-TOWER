(function () {
  "use strict";

  const U = window.CTUtils;

  const COLUMN_GROUPS = {
    base: ["Base de entrega", "Base"],
    driver: ["Entregador", "Motorista", "Courier"],
    deliveredTime: ["Horário da entrega", "Horario da entrega"],
    problemReason: [
      "Motivos dos pacotes problemáticos",
      "Motivos dos pacotes problematicos",
      "Pacote problemático",
      "Pacote problematico"
    ],
    total: [
      "Número total de expedido",
      "Numero total de expedido",
      "Total Expedido"
    ],
    signed: [
      "Número de pacotes assinados",
      "Numero de pacotes assinados",
      "Pacotes assinados",
      "Entregues"
    ],
    undelivered: ["Não entregue", "Nao entregue"],
    problematic: ["Pacote problemático", "Pacote problematico", "Problemático", "Problematico"],
    pending: ["Pacote não expedido", "Pacote nao expedido", "Não expedido", "Nao expedido", "Pendente"],
    regional: ["Regional"]
  };

  function getField(row, keys) {
    const keyList = Array.isArray(keys) ? keys : [keys];

    for (let i = 0; i < keyList.length; i += 1) {
      if (Object.prototype.hasOwnProperty.call(row, keyList[i])) {
        return row[keyList[i]];
      }
    }

    const normalizedTargets = keyList.map(U.normalizar);

    const rowKeys = Object.keys(row);
    for (let i = 0; i < rowKeys.length; i += 1) {
      if (normalizedTargets.includes(U.normalizar(rowKeys[i]))) {
        return row[rowKeys[i]];
      }
    }

    return null;
  }

  function hasAnyColumn(headers, aliases) {
    const normalizedHeaders = headers.map(U.normalizar);
    return aliases.some(function (alias) {
      return normalizedHeaders.includes(U.normalizar(alias));
    });
  }

  function getMissingColumns(headers) {
    return {
      base: hasAnyColumn(headers, COLUMN_GROUPS.base) ? [] : ["Base de entrega"],
      detailed: [
        hasAnyColumn(headers, COLUMN_GROUPS.driver) ? null : "Entregador",
        hasAnyColumn(headers, COLUMN_GROUPS.deliveredTime) ? null : "Horário da entrega",
        hasAnyColumn(headers, COLUMN_GROUPS.problemReason) ? null : "Motivos dos pacotes problemáticos"
      ].filter(Boolean),
      summary: [
        hasAnyColumn(headers, COLUMN_GROUPS.total) ? null : "Número total de expedido",
        hasAnyColumn(headers, COLUMN_GROUPS.signed) ? null : "Número de pacotes assinados",
        hasAnyColumn(headers, COLUMN_GROUPS.undelivered) ? null : "Não entregue",
        hasAnyColumn(headers, COLUMN_GROUPS.problematic) ? null : "Pacote problemático",
        hasAnyColumn(headers, COLUMN_GROUPS.pending) ? null : "Pacote não expedido"
      ].filter(Boolean)
    };
  }

  function scoreSheet(headers) {
    let score = 0;
    Object.keys(COLUMN_GROUPS).forEach(function (key) {
      if (hasAnyColumn(headers, COLUMN_GROUPS[key])) score += 2;
    });

    const missing = getMissingColumns(headers);
    if (missing.base.length === 0) score += 5;
    if (missing.detailed.length <= 1) score += 3;
    if (missing.summary.length <= 2) score += 3;
    return score;
  }

  function analyzeWorkbook(file, arrayBuffer) {
    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    const sheets = workbook.SheetNames.map(function (sheetName) {
      const sheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      const headers = rows.length ? Object.keys(rows[0]) : [];
      const sheetScore = scoreSheet(headers);

      return {
        fileName: file.name,
        workbook,
        sheetName,
        headers,
        rawRows: rows,
        score: sheetScore,
        ignored: sheetScore === 0 || rows.length === 0
      };
    });

    const relevantSheets = sheets.filter(function (item) {
      return !item.ignored;
    });

    const selected = relevantSheets.sort(function (a, b) {
      return b.score - a.score;
    })[0] || sheets[0] || null;

    return {
      fileName: file.name,
      workbook,
      sheets,
      selectedSheetName: selected ? selected.sheetName : "",
      ignoredSheets: sheets.filter(function (item) {
        return item.ignored;
      }).length
    };
  }

  function classifyRow(rawRow) {
    const base = String(getField(rawRow, COLUMN_GROUPS.base) || "BASE INDEFINIDA").trim();
    const driver = String(getField(rawRow, COLUMN_GROUPS.driver) || "NÃO ATRIBUÍDO").trim();
    const regional = String(getField(rawRow, COLUMN_GROUPS.regional) || "").trim();

    const total = U.toNumber(getField(rawRow, COLUMN_GROUPS.total));
    const delivered = U.toNumber(getField(rawRow, COLUMN_GROUPS.signed));
    const undelivered = U.toNumber(getField(rawRow, COLUMN_GROUPS.undelivered));
    const problematic = U.toNumber(getField(rawRow, COLUMN_GROUPS.problematic));
    const pending = U.toNumber(getField(rawRow, COLUMN_GROUPS.pending));

    const deliveredTime = getField(rawRow, COLUMN_GROUPS.deliveredTime);
    const problemReason = getField(rawRow, COLUMN_GROUPS.problemReason);

    const isSummary = total > 0 || delivered > 0 || undelivered > 0 || problematic > 0 || pending > 0;

    let status = "pendente";
    if (isSummary) {
      status = "resumo";
    } else if (deliveredTime) {
      status = "entregue";
    } else if (problemReason && undelivered > 0) {
      status = "nao_entregue";
    } else if (problemReason) {
      status = "problematico";
    } else {
      status = "pendente";
    }

    const isValid = Boolean(base && base !== "BASE INDEFINIDA") && (isSummary || driver !== "NÃO ATRIBUÍDO" || deliveredTime || problemReason);

    return {
      base,
      regional,
      driver,
      deliveredTime: deliveredTime || "",
      problemReason: problemReason || "",
      total,
      delivered,
      undelivered,
      problematic,
      pending,
      isSummary,
      status,
      isValid,
      raw: rawRow
    };
  }

  function normalizeRows(rows) {
    const normalizedRows = [];
    const invalidRows = [];

    rows.forEach(function (row, index) {
      const normalized = classifyRow(row);
      normalized.rowIndex = index + 2;

      if (normalized.isValid || normalized.isSummary) normalizedRows.push(normalized);
      else invalidRows.push(normalized);
    });

    return {
      normalizedRows,
      invalidRows
    };
  }

  function validateSheet(rows) {
    const headers = rows.length ? Object.keys(rows[0]) : [];
    const missing = getMissingColumns(headers);
    const hasBase = missing.base.length === 0;
    const canUseDetailed = missing.detailed.length < 3;
    const canUseSummary = missing.summary.length < 5;

    return {
      headers,
      missing,
      isUsable: hasBase && (canUseDetailed || canUseSummary)
    };
  }

  function summarizeSelection(fileAnalysis, selectedSheetName) {
    const selectedSheet = fileAnalysis.sheets.find(function (item) {
      return item.sheetName === selectedSheetName;
    });

    if (!selectedSheet) {
      return {
        fileName: fileAnalysis.fileName,
        selectedSheetName: "",
        validation: {
          headers: [],
          missing: { base: ["Base de entrega"], detailed: [], summary: [] },
          isUsable: false
        },
        normalizedRows: [],
        invalidRows: [],
        ignoredSheets: fileAnalysis.ignoredSheets
      };
    }

    const validation = validateSheet(selectedSheet.rawRows);
    const normalized = normalizeRows(selectedSheet.rawRows);

    return {
      fileName: fileAnalysis.fileName,
      selectedSheetName,
      validation,
      normalizedRows: normalized.normalizedRows,
      invalidRows: normalized.invalidRows,
      ignoredSheets: fileAnalysis.ignoredSheets
    };
  }

  async function inspectFiles(fileList) {
    const files = Array.from(fileList || []);
    const analyses = [];

    for (let i = 0; i < files.length; i += 1) {
      const file = files[i];
      const buffer = await file.arrayBuffer();
      analyses.push(analyzeWorkbook(file, buffer));
    }

    return analyses;
  }

  function buildPreviewReport(analyses, selectedSheetsMap) {
    const files = analyses.map(function (analysis) {
      return summarizeSelection(analysis, selectedSheetsMap[analysis.fileName] || analysis.selectedSheetName);
    });

    const report = files.reduce(function (acc, item) {
      acc.fileCount += 1;
      acc.validRows += item.normalizedRows.length;
      acc.invalidRows += item.invalidRows.length;
      acc.ignoredSheets += item.ignoredSheets;
      return acc;
    }, {
      fileCount: 0,
      validRows: 0,
      invalidRows: 0,
      ignoredSheets: 0
    });

    return {
      files,
      report
    };
  }

  function mergeImportedRows(currentRows, importedRows, mode) {
    if (mode === "append") return U.safeArray(currentRows).concat(U.safeArray(importedRows));
    return U.safeArray(importedRows);
  }

  function createSampleRows() {
    const sample = [];
    const definitions = [
      { base: "VRG-SP", regional: "Neto", total: 40, delivered: 33, problematic: 2, undelivered: 2, pending: 3 },
      { base: "AME-SP", regional: "Luana", total: 55, delivered: 49, problematic: 2, undelivered: 1, pending: 3 },
      { base: "SMR-SP", regional: "Luana", total: 60, delivered: 51, problematic: 3, undelivered: 2, pending: 4 },
      { base: "GRU-SP", regional: "Claudio", total: 35, delivered: 31, problematic: 1, undelivered: 1, pending: 2 },
      { base: "SDR-SP", regional: "Rodrigo", total: 38, delivered: 26, problematic: 4, undelivered: 4, pending: 4 },
      { base: "CHM-SP", regional: "Neto", total: 30, delivered: 24, problematic: 1, undelivered: 2, pending: 3 }
    ];

    definitions.forEach(function (item) {
      for (let i = 1; i <= item.total; i += 1) {
        const deliveredLimit = item.delivered;
        const problematicLimit = deliveredLimit + item.problematic;
        const undeliveredLimit = problematicLimit + item.undelivered;

        let rowStatus = "pendente";
        if (i <= deliveredLimit) rowStatus = "entregue";
        else if (i <= problematicLimit) rowStatus = "problematico";
        else if (i <= undeliveredLimit) rowStatus = "nao_entregue";

        sample.push({
          base: item.base,
          regional: item.regional,
          driver: `Entregador ${String(((i - 1) % 6) + 1).padStart(2, "0")}`,
          deliveredTime: rowStatus === "entregue" ? "10:30" : "",
          problemReason: rowStatus === "problematico"
            ? "Pacote avariado"
            : rowStatus === "nao_entregue"
              ? "Cliente ausente"
              : "",
          total: 0,
          delivered: 0,
          undelivered: 0,
          problematic: 0,
          pending: 0,
          isSummary: false,
          status: rowStatus,
          isValid: true,
          raw: {}
        });
      }
    });

    return sample;
  }

  window.CTExcel = {
    COLUMN_GROUPS,
    getField,
    inspectFiles,
    buildPreviewReport,
    mergeImportedRows,
    createSampleRows
  };
})();