(function () {
  "use strict";

  function qs(selector, root) {
    return (root || document).querySelector(selector);
  }

  function qsa(selector, root) {
    return Array.from((root || document).querySelectorAll(selector));
  }

  function byId(id) {
    return document.getElementById(id);
  }

  function setText(idOrElement, value) {
    const el = typeof idOrElement === "string" ? byId(idOrElement) : idOrElement;
    if (el) el.textContent = value;
  }

  function setHtml(idOrElement, value) {
    const el = typeof idOrElement === "string" ? byId(idOrElement) : idOrElement;
    if (el) el.innerHTML = value;
  }

  function toggleHidden(idOrElement, hidden) {
    const el = typeof idOrElement === "string" ? byId(idOrElement) : idOrElement;
    if (el) el.hidden = Boolean(hidden);
  }

  function formatNumber(value) {
    return Number(value || 0).toLocaleString("pt-BR");
  }

  function formatPercent(value, digits) {
    const num = Number(value || 0);
    return `${num.toFixed(digits ?? 2)}%`;
  }

  function formatDateTimeBR(value) {
    if (!value) return "Última atualização indisponível";
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return "Última atualização indisponível";
    return date.toLocaleString("pt-BR");
  }

  function escapeHtml(str) {
    return String(str || "").replace(/[&<>"']/g, function (s) {
      return {
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;"
      }[s];
    });
  }

  function escapeJs(str) {
    return String(str || "")
      .replace(/\\/g, "\\\\")
      .replace(/'/g, "\\'")
      .replace(/\r?\n/g, " ");
  }

  function sanitizeFilename(name) {
    return String(name || "arquivo").replace(/[^a-z0-9_\-]/gi, "_");
  }

  function normalizar(txt) {
    return String(txt || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, "")
      .replace(/-/g, "")
      .replace(/[^\w]/g, "")
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

  function safeArray(value) {
    return Array.isArray(value) ? value : [];
  }

  function clone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function debounce(fn, wait) {
    let timer = null;
    return function () {
      const ctx = this;
      const args = arguments;
      clearTimeout(timer);
      timer = setTimeout(function () {
        fn.apply(ctx, args);
      }, wait);
    };
  }

  function showMessage(targetId, message, type) {
    const box = byId(targetId);
    if (!box) return;
    box.className = `message-box is-${type || "info"}`;
    box.innerHTML = escapeHtml(message);
  }

  function clearMessage(targetId) {
    const box = byId(targetId);
    if (!box) return;
    box.className = "message-box";
    box.innerHTML = "";
  }

  function storageSet(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch (error) {
      console.warn("Falha ao salvar no localStorage.", error);
    }
  }

  function storageGet(key, fallbackValue) {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallbackValue;
    } catch (error) {
      console.warn("Falha ao ler do localStorage.", error);
      return fallbackValue;
    }
  }

  function downloadBlobURL(url, filename) {
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    link.click();
  }

  window.CTUtils = {
    qs,
    qsa,
    byId,
    setText,
    setHtml,
    toggleHidden,
    formatNumber,
    formatPercent,
    formatDateTimeBR,
    escapeHtml,
    escapeJs,
    sanitizeFilename,
    normalizar,
    toNumber,
    safeArray,
    clone,
    debounce,
    showMessage,
    clearMessage,
    storageSet,
    storageGet,
    downloadBlobURL
  };
})();