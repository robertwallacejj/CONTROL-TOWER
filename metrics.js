(function () {
  "use strict";

  const U = window.CTUtils;
  const Excel = window.CTExcel || {};

  const REGIONAIS = {
    claudio: ["S-CRDR-SP", "GRU-SP", "S-CSVD-SP", "S-BRFD-SP", "S-FREG-SP", "F GRU-SP", "S-BRAS-SP", "F S-JRG-SP", "F S-VLMR-SP", "GRU 03-SP", "S-VLGUI-SP", "F S-BRSL-SP", "F S-BLV-SP"],
    rodrigo: ["S-SAPOP-SP", "S-PENHA-SP", "S-MGUE-SP", "MGC-SP", "ARJ-SP", "SDR-SP", "S-SRAF-SP", "F ITQ-SP", "F S-PENHA-SP", "F S-PENHA 02-SP", "F S-MGUE-SP"],
    neto: ["CARAP-SP", "CHM-SP", "COT-SP", "JDR-SP", "OSC-SP", "S-VLANA-SP", "S-VLLEO-SP", "S-VLSN-SP", "TBA-SP", "VRG-SP"],
    luana: ["AME-SP", "FRCLR-SP", "F VCP-SP", "MGG-SP", "PIR-SP", "RCLR-SP", "SMR-SP", "VCP 03-SP", "VCP 05-SP", "VIN-SP", "FJND-SP", "ITUP-SP", "JND-SP", "BRG-SP", "CAIE-SP", "ATB-SP", "F SOD 02-SP", "IBUN-SP", "ITPT-SP", "ITPV-SP", "ITU-SP", "SOD02-SP", "SOD-SP", "SRQ-SP", "INDTR SD"]
  };

  let cacheKey = "";
  let cacheValue = null;

  function getRegionalFromBase(baseName) {
    const normalizedBase = U.normalizar(baseName);
    if (REGIONAIS.claudio.some((b) => U.normalizar(b) === normalizedBase)) return "Claudio";
    if (REGIONAIS.rodrigo.some((b) => U.normalizar(b) === normalizedBase)) return "Rodrigo";
    if (REGIONAIS.neto.some((b) => U.normalizar(b) === normalizedBase)) return "Neto";
    if (REGIONAIS.luana.some((b) => U.normalizar(b) === normalizedBase)) return "Luana";
    return "Não definida";
  }

  function normalizeLegacyRow(row) {
    if (row && typeof row === "object" && "isSummary" in row && "base" in row) {
      if (!row.regional) row.regional = getRegionalFromBase(row.base);
      return row;
    }

    const base = String((Excel.getField && Excel.getField(row, ["Base de entrega", "Base"])) || "BASE INDEFINIDA").trim();
    const driver = String((Excel.getField && Excel.getField(row, ["Entregador", "Motorista", "Courier"])) || "NÃO ATRIBUÍDO").trim();
    const total = U.toNumber(Excel.getField && Excel.getField(row, ["Número total de expedido", "Numero total de expedido", "Total Expedido"]));
    const delivered = U.toNumber(Excel.getField && Excel.getField(row, ["Número de pacotes assinados", "Numero de pacotes assinados", "Pacotes assinados", "Entregues"]));
    const undelivered = U.toNumber(Excel.getField && Excel.getField(row, ["Não entregue", "Nao entregue"]));
    const problematic = U.toNumber(Excel.getField && Excel.getField(row, ["Pacote problemático", "Pacote problematico", "Problemático", "Problematico"]));
    const pending = U.toNumber(Excel.getField && Excel.getField(row, ["Pacote não expedido", "Pacote nao expedido", "Não expedido", "Nao expedido", "Pendente"]));
    const deliveredTime = Excel.getField && Excel.getField(row, ["Horário da entrega", "Horario da entrega"]);
    const problemReason = Excel.getField && Excel.getField(row, ["Motivos dos pacotes problemáticos", "Motivos dos pacotes problematicos", "Pacote problemático", "Pacote problematico"]);
    const isSummary = total > 0 || delivered > 0 || undelivered > 0 || problematic > 0 || pending > 0;
    let status = "pendente";
    if (isSummary) status = "resumo";
    else if (deliveredTime) status = "entregue";
    else if (problemReason && undelivered > 0) status = "nao_entregue";
    else if (problemReason) status = "problematico";

    return {
      base,
      driver,
      regional: String((Excel.getField && Excel.getField(row, ["Regional"])) || getRegionalFromBase(base)).trim(),
      deliveredTime: deliveredTime || "",
      problemReason: problemReason || "",
      total,
      delivered,
      undelivered,
      problematic,
      pending,
      isSummary,
      status,
      isValid: true,
      raw: row
    };
  }

  function getRowsSignature(rows) {
    const safeRows = U.safeArray(rows);
    const first = safeRows[0] || {};
    const last = safeRows[safeRows.length - 1] || {};
    return `${safeRows.length}::${Object.keys(first).join("|")}::${Object.keys(last).join("|")}`;
  }

  function aggregateBaseMetrics(rows) {
    const signature = getRowsSignature(rows);
    if (cacheKey === signature && cacheValue) return cacheValue;

    const grouped = {};
    U.safeArray(rows).forEach(function (sourceRow) {
      const row = normalizeLegacyRow(sourceRow);
      const base = row.base || "BASE INDEFINIDA";
      if (!grouped[base]) {
        grouped[base] = { base, regional: row.regional || getRegionalFromBase(base), total: 0, entregue: 0, problematico: 0, naoEntregue: 0, pendente: 0, insucesso: 0, taxa: 0 };
      }

      if (row.isSummary) {
        grouped[base].total += row.total;
        grouped[base].entregue += row.delivered;
        grouped[base].problematico += row.problematic;
        grouped[base].naoEntregue += row.undelivered;
        grouped[base].pendente += row.pending;
      } else {
        grouped[base].total += 1;
        if (row.status === "entregue") grouped[base].entregue += 1;
        else if (row.status === "problematico") grouped[base].problematico += 1;
        else if (row.status === "nao_entregue") grouped[base].naoEntregue += 1;
        else grouped[base].pendente += 1;
      }

      grouped[base].insucesso = grouped[base].problematico + grouped[base].naoEntregue;
      grouped[base].taxa = grouped[base].total > 0 ? (grouped[base].entregue / grouped[base].total) * 100 : 0;
    });

    cacheKey = signature;
    cacheValue = Object.values(grouped).sort(function (a, b) {
      return a.base.localeCompare(b.base, "pt-BR");
    });
    return cacheValue;
  }

  function aggregateGlobal(rows) {
    return aggregateBaseMetrics(rows).reduce(function (acc, item) {
      acc.total += item.total;
      acc.entregue += item.entregue;
      acc.problematico += item.problematico;
      acc.naoEntregue += item.naoEntregue;
      acc.pendente += item.pendente;
      acc.insucesso += item.insucesso;
      return acc;
    }, { total: 0, entregue: 0, problematico: 0, naoEntregue: 0, pendente: 0, insucesso: 0 });
  }

  function aggregateDrivers(rows) {
    const grouped = {};
    U.safeArray(rows).forEach(function (sourceRow) {
      const row = normalizeLegacyRow(sourceRow);
      if (row.isSummary) return;
      const key = `${row.base}__${row.driver}`;
      if (!grouped[key]) grouped[key] = { base: row.base, driver: row.driver, total: 0, entregue: 0, problematico: 0, naoEntregue: 0, pendente: 0, taxa: 0 };
      grouped[key].total += 1;
      if (row.status === "entregue") grouped[key].entregue += 1;
      else if (row.status === "problematico") grouped[key].problematico += 1;
      else if (row.status === "nao_entregue") grouped[key].naoEntregue += 1;
      else grouped[key].pendente += 1;
      grouped[key].taxa = grouped[key].total > 0 ? (grouped[key].entregue / grouped[key].total) * 100 : 0;
    });
    return Object.values(grouped);
  }

  function filterMetrics(metrics, filters) {
    const regional = filters && filters.regional ? filters.regional : "all";
    const base = filters && filters.base ? filters.base : "all";
    const status = filters && filters.status ? filters.status : "all";
    const target = Number(filters && filters.target) || 90;
    const search = U.normalizar(filters && filters.search ? filters.search : "");

    return U.safeArray(metrics).filter(function (item) {
      const matchesRegional = regional === "all" || item.regional === regional;
      const matchesBase = base === "all" || item.base === base;
      const matchesSearch = !search || U.normalizar(item.base).includes(search);
      let matchesStatus = true;
      if (status === "critical") matchesStatus = item.taxa < target;
      if (status === "healthy") matchesStatus = item.taxa >= target;
      return matchesRegional && matchesBase && matchesStatus && matchesSearch;
    });
  }

  window.CTMetrics = { REGIONAIS, normalizeLegacyRow, getRegionalFromBase, aggregateBaseMetrics, aggregateGlobal, aggregateDrivers, filterMetrics };
})();
