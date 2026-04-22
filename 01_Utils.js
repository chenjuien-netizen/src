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

function normalizeBasicSheetHeaderKey_(header) {
  if (header === null || typeof header === "undefined") return "";
  return String(header)
    .trim()
    .replace(/[’`´]/g, "'")
    .replace(/\s+/g, " ")
    .toLowerCase();
}

var SHEET_HEADER_ALIASES_ = {};
SHEET_HEADER_ALIASES_[normalizeBasicSheetHeaderKey_("进货")] = normalizeBasicSheetHeaderKey_("修改日期");
SHEET_HEADER_ALIASES_[normalizeBasicSheetHeaderKey_("当前尾箱件数")] = normalizeBasicSheetHeaderKey_("尾箱");
SHEET_HEADER_ALIASES_[normalizeBasicSheetHeaderKey_("箱件")] = normalizeBasicSheetHeaderKey_("件/箱");
SHEET_HEADER_ALIASES_[normalizeBasicSheetHeaderKey_("当前箱数")] = normalizeBasicSheetHeaderKey_("箱数");
SHEET_HEADER_ALIASES_[normalizeBasicSheetHeaderKey_("出-Sortie/箱")] = normalizeBasicSheetHeaderKey_("开箱/包");
SHEET_HEADER_ALIASES_[normalizeBasicSheetHeaderKey_("出库记录")] = normalizeBasicSheetHeaderKey_("清库记录");

function normalizeSheetHeaderKey_(header) {
  var basicKey = normalizeBasicSheetHeaderKey_(header);
  return SHEET_HEADER_ALIASES_[basicKey] || basicKey;
}

function headerMap_(headersRow) {
  var map = {};
  for (var c = 0; c < headersRow.length; c++) {
    var h = headersRow[c];
    if (h === null || typeof h === "undefined") continue;
    var rawKey = normalizeBasicSheetHeaderKey_(h);
    var key = normalizeSheetHeaderKey_(h);
    if (!rawKey && !key) continue;
    if (rawKey) map[rawKey] = c + 1; // compat alias
    if (key) map[key] = c + 1; // canonical
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

function findColumnByHeaderNoteKey_(sheet, key) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return 0;

  var notes = sheet.getRange(1, 1, 1, lastCol).getNotes()[0];
  for (var i = 0; i < notes.length; i++) {
    if (notes[i] && String(notes[i]).indexOf(key) !== -1) return i + 1;
  }
  return 0;
}

function rebuildSheetFilterToLastColumn_(sheet) {
  var lastRow = Math.max(1, sheet.getLastRow());
  var lastCol = Math.max(1, sheet.getLastColumn());
  var existing = sheet.getFilter();
  if (existing) existing.remove();
  sheet.getRange(1, 1, lastRow, lastCol).createFilter();
}

var STOCK_TEMPLATE_PROPAGATION_HEADERS = [
  "Poids (en gramme)",
  "Pays d'origine",
  "SortKey",
  "Colisage",
  "Couleur",
  "Couleurs",
  "Promo",
  "Prix@",
  "Promo@",
  "剩下 / RESTE",
  "包/箱",
  "Carton ouvert (reste)"
];

var STOCK_TEMPLATE_PROPAGATION_NOTE_KEYS = [
  "KEY:TOTAL_BOX",
  "KEY:TOTAL_PQS",
  "KEY:TOTAL_PCS"
];

function applyTemplateCellToRange_(shTpl, shStock, tplCol, stockCol, addCount, startRow) {
  if (!tplCol || !stockCol || !addCount || !startRow) return;

  var tplCell = shTpl.getRange(2, tplCol);
  var formula = tplCell.getFormulaR1C1();

  if (formula) {
    var formulas = [];
    for (var i = 0; i < addCount; i++) formulas.push([formula]);
    shStock.getRange(startRow, stockCol, addCount, 1).setFormulasR1C1(formulas);
    return;
  }

  var value = tplCell.getValue();
  var values = [];
  for (var j = 0; j < addCount; j++) values.push([value]);
  shStock.getRange(startRow, stockCol, addCount, 1).setValues(values);
}

function applyStockTemplatePropagation_(shTpl, shStock, addCount, startRow, headers, noteKeys) {
  if (!addCount || addCount <= 0) return;
  if (!startRow || startRow <= 0) throw new Error("applyStockTemplatePropagation_: startRow invalide");

  var tplHeaders = shTpl.getRange(1, 1, 1, Math.max(1, shTpl.getLastColumn())).getValues()[0];
  var stockHeaders = shStock.getRange(1, 1, 1, Math.max(1, shStock.getLastColumn())).getValues()[0];
  var tplHeaderMap = headerMap_(tplHeaders);
  var stockHeaderMap = headerMap_(stockHeaders);
  var effectiveHeaders = Array.isArray(headers) && headers.length ? headers : STOCK_TEMPLATE_PROPAGATION_HEADERS;
  var effectiveNoteKeys = Array.isArray(noteKeys) && noteKeys.length ? noteKeys : STOCK_TEMPLATE_PROPAGATION_NOTE_KEYS;

  for (var i = 0; i < effectiveHeaders.length; i++) {
    var header = String(effectiveHeaders[i] || "");
    var key = normalizeSheetHeaderKey_(header);
    var tplCol = tplHeaderMap[key];
    var stockCol = stockHeaderMap[key];
    if (!tplCol || !stockCol) continue;
    applyTemplateCellToRange_(shTpl, shStock, tplCol, stockCol, addCount, startRow);
  }

  for (var j = 0; j < effectiveNoteKeys.length; j++) {
    var noteKey = String(effectiveNoteKeys[j] || "");
    if (!noteKey) continue;
    var tplNoteCol = findColumnByHeaderNoteKey_(shTpl, noteKey);
    var stockNoteCol = findColumnByHeaderNoteKey_(shStock, noteKey);
    if (!tplNoteCol || !stockNoteCol) continue;
    applyTemplateCellToRange_(shTpl, shStock, tplNoteCol, stockNoteCol, addCount, startRow);
  }
}

function rebuildAndSortStockSheet_(sheet) {
  rebuildSheetFilterToLastColumn_(sheet);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var lastCol = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = headerMap_(headers);
  var sortCol = map["sortkey"];
  if (!sortCol) return;

  sheet.getRange(2, 1, lastRow - 1, lastCol).sort({ column: sortCol, ascending: true });
}

function readSheetRowsByNumbers_(sheet, rowNumbers, lastCol) {
  var rows = Array.isArray(rowNumbers) ? rowNumbers.slice() : [];
  if (!sheet || !rows.length || !lastCol) return [];

  rows.sort(function(a, b) { return a - b; });

  var out = [];
  var blockStart = rows[0];
  var blockEnd = rows[0];

  function pushBlock_(startRow, endRow) {
    var numRows = endRow - startRow + 1;
    var values = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
    for (var i = 0; i < values.length; i++) {
      var rowNumber = startRow + i;
      out.push({
        row: rowNumber,
        rowIndex: rowNumber - 2,
        values: values[i]
      });
    }
  }

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (row === blockEnd + 1) {
      blockEnd = row;
      continue;
    }
    pushBlock_(blockStart, blockEnd);
    blockStart = row;
    blockEnd = row;
  }

  pushBlock_(blockStart, blockEnd);
  return out;
}

function buildSelectedStockExportContext_(requiredHeaders) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stock = ss.getSheetByName(SHEET_STOCK);

  if (!stock) {
    throw new Error("Feuille introuvable: " + SHEET_STOCK);
  }

  var lastRow = stock.getLastRow();
  var lastCol = stock.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return {
      ss: ss,
      stock: stock,
      stockHeaders: [],
      stockMap: {},
      lastRow: lastRow,
      lastCol: lastCol,
      selectionCol: 0,
      selectedRows: [],
      selectedRecords: []
    };
  }

  var stockHeaders = stock.getRange(1, 1, 1, lastCol).getValues()[0];
  var stockMap = headerMap_(stockHeaders);
  if (Array.isArray(requiredHeaders) && requiredHeaders.length) {
    ensureHeadersExist_(stockMap, requiredHeaders, "STOCK");
  }

  var selectionCol = stockMap["选择"];
  if (!selectionCol) {
    throw new Error("Colonne '选择' introuvable dans STOCK.");
  }

  var selectedRows = [];
  if (lastRow >= 2) {
    var values = stock.getRange(2, selectionCol, lastRow - 1, 1).getValues().flat();
    for (var i = 0; i < values.length; i++) {
      if (values[i] === true) selectedRows.push(i + 2);
    }
  }

  return {
    ss: ss,
    stock: stock,
    stockHeaders: stockHeaders,
    stockMap: stockMap,
    lastRow: lastRow,
    lastCol: lastCol,
    selectionCol: selectionCol,
    selectedRows: selectedRows,
    selectedRecords: readSheetRowsByNumbers_(stock, selectedRows, lastCol)
  };
}

function logPerfStep_(scope, label, startMs, extra) {
  var duration = Date.now() - Number(startMs || 0);
  var suffix = extra ? " | " + String(extra) : "";
  Logger.log(String(scope || "perf") + " | " + String(label || "step") + "=" + duration + " ms" + suffix);
}

function normalizeColorCatalogKey_(value) {
  var s = (value === null || typeof value === "undefined") ? "" : String(value);
  return s
    .trim()
    .replace(/[’`´']/g, "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Za-z0-9]+/g, "")
    .toUpperCase();
}

var STOCK_COLOR_CANONICAL_LIST = [
  "Écru",
  "Ivoire",
  "Vanille",
  "Nude",
  "Crême",
  "Beige",
  "Blanc",
  "Transparent",
  "Bleu Ciel",
  "Bleu Clair",
  "Bleu",
  "Cyan",
  "Turquoise",
  "Bleu Roi",
  "Jeans",
  "Denim",
  "Bleu Canard",
  "Bleu Pétrole",
  "Bleu Foncé",
  "Bleu Irisé",
  "Marine",
  "Ciel nocturne",
  "Gris Clair",
  "Gris Perle",
  "Argent",
  "Gris",
  "Gris Souris",
  "Acier",
  "Gris Foncé",
  "Carbone",
  "Gris Ardoise",
  "Anthracite",
  "Jaune Clair",
  "Jaune Citron",
  "Jaune",
  "Jaune Foncé",
  "Jaune Soleil",
  "Jaune Fluo",
  "Or",
  "Moutarde",
  "Doré",
  "Ocre",
  "Caramel",
  "Bronze",
  "Camel",
  "Champagne",
  "Taupe",
  "Cognac",
  "Brun",
  "Terracotta",
  "Brun foncé",
  "Marron Clair",
  "Chocolat",
  "Marron Foncé",
  "Multicolore",
  "Bicolore",
  "Noir Irisé",
  "Noir",
  "Saumon",
  "Corail",
  "Abricot",
  "Orange Fluo",
  "Orange",
  "Rouge Orangé",
  "Cuivre",
  "Brique",
  "Rouille",
  "Blush",
  "Rose",
  "Rose Fluo",
  "Fuchsia",
  "Magenta",
  "Framboise",
  "Vieux Rose",
  "Rouge Clair",
  "Rouge",
  "Carmin",
  "Rouge Foncé",
  "Bordeaux",
  "Vert Clair",
  "Vert d'Eau",
  "Céladon",
  "Vert Fluo",
  "Vert Pomme",
  "Vert",
  "Vert Foncé",
  "Vert Bouteille",
  "Vert Sapin",
  "Vert Canard",
  "Olive",
  "Kaki",
  "Lilas",
  "Lavande",
  "Mauve",
  "Violet",
  "Indigo",
  "Prune"
];

function buildCanonicalColorMap_() {
  var map = {};
  for (var i = 0; i < STOCK_COLOR_CANONICAL_LIST.length; i++) {
    var canonical = STOCK_COLOR_CANONICAL_LIST[i];
    map[normalizeColorCatalogKey_(canonical)] = canonical;
  }

  map["CREME"] = "Crême";
  map["CRÈME"] = "Crême";
  map["CRÊME"] = "Crême";
  map["VERTDEAU"] = "Vert d'Eau";

  return map;
}

var STOCK_COLOR_CANONICAL_MAP = buildCanonicalColorMap_();

function normalizeStockCatalogColorName_(raw) {
  var key = normalizeColorCatalogKey_(raw);
  if (!key) return "";
  return STOCK_COLOR_CANONICAL_MAP[key] || "";
}
