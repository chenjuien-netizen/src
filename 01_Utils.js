// 01_Utils.gs
let __uiGuard = false;
function isUiGuardOn_() { return __uiGuard === true; }
function withUiGuard_(fn) {
  __uiGuard = true;
  try { return fn(); }
  finally { __uiGuard = false; }
}

function notify_(title, msg) {
  try {
    SpreadsheetApp.getUi().alert(String(title || "Info"), String(msg || ""), SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    try { SpreadsheetApp.getActive().toast(String(msg || ""), String(title || "Info"), 6); } catch (e2) {}
  }
}

function cleanRef_(s) {
  return String(s || "")
    .trim()
    .replace(/\s+/g, "")
    .replace(/[＃#]/g, "");
}

function toIntSafe_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return Math.trunc(v);
  const s = String(v).trim();
  if (!s) return 0;
  const m = s.match(/-?\d+(\.\d+)?/);
  if (!m) return 0;
  const n = Number(m[0]);
  return Number.isFinite(n) ? Math.trunc(n) : 0;
}

function parseMixPartnerFromNoteSystem_(noteS) {
  const s = String(noteS || "").trim();
  // On ignore MIX_START et tout le reste : on ne prend que le partner après "MIX avec:"
  // Accepte: "MIX avec: JY72-1", "MIX avec: JY72-1 MIX_START", "MIX_START MIX avec: JY72-1", etc.
  const m = s.match(/MIX\s+avec\s*:\s*([A-Za-z0-9\-_.#*]+)/i);
  if (!m) return "";
  const raw = m[1];
  return (typeof cleanRef_ === "function" ? cleanRef_(raw) : String(raw || "")).toUpperCase();
}

function normalizeSheetHeaderKey_(header) {
  if (header === null || typeof header === "undefined") return "";
  return String(header)
    .trim()
    .replace(/[’`´]/g, "'")
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function headerMap_(headersRow) {
  var map = {};
  for (var c = 0; c < headersRow.length; c++) {
    var h = headersRow[c];
    if (h === null || typeof h === "undefined") continue;
    var key = normalizeSheetHeaderKey_(h);
    if (!key) continue;
    map[key] = c + 1; // 1-based
  }
  return map;
}

function ensureHeadersExist_(map, requiredHeaders, where) {
  var missing = [];
  for (var i = 0; i < requiredHeaders.length; i++) {
    var k = requiredHeaders[i];
    if (!map.hasOwnProperty(normalizeSheetHeaderKey_(k))) missing.push(k);
  }
  if (missing.length) {
    throw new Error("Headers manquants dans " + where + ": " + missing.join(", "));
  }
}

function normalizeUpper_(s) {
  var t = (s === null || typeof s === "undefined") ? "" : String(s);
  return t.trim().toUpperCase();
}
