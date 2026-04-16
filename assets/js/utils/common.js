(function initForm15Utils(global) {
  function nowMs() {
    return Date.now();
  }

  function formatTime(ts) {
    try {
      return new Date(ts).toLocaleString("vi-VN");
    } catch (_) {
      return "";
    }
  }

  function normalizeText(input) {
    return String(input || "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  }

  function normalizeCompact(input) {
    return normalizeText(input).replace(/\s+/g, "");
  }

  function isSameLoose(a, b) {
    return normalizeCompact(a) === normalizeCompact(b);
  }

  function htmlEscape(str) {
    return String(str ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  global.Form15Utils = {
    nowMs,
    formatTime,
    normalizeText,
    normalizeCompact,
    isSameLoose,
    htmlEscape,
  };
})(window);

