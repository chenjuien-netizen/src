/**
 * Helpers Microstore (spécifiques module)
 */

function msGetCell_(row, col1Based) {
  var v = row[col1Based - 1];
  return (v === null || typeof v === "undefined") ? "" : String(v);
}

function msNormalizeRef_(s) {
  var t = (s === null || typeof s === "undefined") ? "" : String(s);
  if (typeof cleanRef_ === "function") {
    return cleanRef_(t).toUpperCase();
  }
  return t.trim().toUpperCase();
}

function msNormalizeCouleur_(s) {
  var t = (s === null || typeof s === "undefined") ? "" : String(s);
  t = t.trim().toUpperCase();
  if (t === "MIXTE" || t === "MIX") return "MIX";
  return t;
}

function msMergeRemarqueNom_(remarque, nom) {
  var r = (remarque || "").trim();
  var n = (nom || "").trim();
  if (!n) return r;
  if (!r) return n;
  if (r.indexOf(n) !== -1) return r;
  return r + " | " + n;
}

function msStripDoublonTag_(s) {
  var t = (s || "");
  t = t.replace(/\s*\[MS_DOUBLON:\d+\]\s*/g, " ").trim();
  return t.replace(/\s{2,}/g, " ").trim();
}

function msApplyDoublonTag_(rem, count) {
  var base = msStripDoublonTag_(rem);
  if (count && count > 1) {
    return base ? (base + " [MS_DOUBLON:" + count + "]") : ("[MS_DOUBLON:" + count + "]");
  }
  return base;
}

function msFormatNowText_() {
  var tz = Session.getScriptTimeZone();
  return Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
}

function msGetEffectiveDimensions_(values) {
  if (!values || values.length === 0) return { rows: 0, cols: 0 };

  var lastRow = -1;
  var lastCol = -1;

  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    var rowHasData = false;

    for (var c = 0; c < row.length; c++) {
      var v = row[c];
      if (v !== "" && v !== null && typeof v !== "undefined") {
        rowHasData = true;
        if (c > lastCol) lastCol = c;
      }
    }
    if (rowHasData) lastRow = r;
  }

  return {
    rows: (lastRow >= 0) ? (lastRow + 1) : 0,
    cols: (lastCol >= 0) ? (lastCol + 1) : 0
  };
}

function msOpenSpreadsheetWithRetry_(fileId, attempts, sleepMs) {
  attempts = attempts || 6;
  sleepMs = sleepMs || 1000;

  var lastErr = null;
  for (var i = 0; i < attempts; i++) {
    try {
      return SpreadsheetApp.openById(fileId);
    } catch (e) {
      lastErr = e;
      Utilities.sleep(sleepMs);
    }
  }
  throw lastErr || new Error("Impossible d'ouvrir le fichier converti après plusieurs tentatives.");
}
