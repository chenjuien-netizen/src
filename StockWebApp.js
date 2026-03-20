function doGet() {
  return HtmlService.createHtmlOutputFromFile("StockMobile")
    .setTitle("SZFashion Stock")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function StockWebApp_getList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_STOCK);

  if (!sh) {
    throw new Error("Feuille introuvable: " + SHEET_STOCK);
  }

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) {
    throw new Error("La feuille STOCK ne contient aucun en-tête.");
  }

  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const cols = StockWebApp_resolveColumns_(headers);
  const rowCount = Math.max(0, lastRow - 1);
  const values = rowCount ? sh.getRange(2, 1, rowCount, lastCol).getDisplayValues() : [];
  const items = [];

  for (let i = 0; i < values.length; i++) {
    const item = StockWebApp_buildItem_(values[i], cols, i + 2);
    if (item) items.push(item);
  }

  Logger.log(
    "StockWebApp_getList_ | sheet=%s | columns=%s | rowsRead=%s | itemsReturned=%s",
    sh.getName(),
    JSON.stringify({
      reference: headers[cols.reference - 1] || "",
      warehouse: headers[cols.warehouse - 1] || "",
      stockDisplay: headers[cols.stockDisplay - 1] || ""
    }),
    rowCount,
    items.length
  );

  return {
    items: items,
    meta: {
      sheetName: sh.getName(),
      rowsRead: rowCount,
      itemsReturned: items.length,
      columns: {
        reference: headers[cols.reference - 1] || "",
        warehouse: headers[cols.warehouse - 1] || "",
        stockDisplay: headers[cols.stockDisplay - 1] || ""
      }
    }
  };
}

function StockWebApp_resolveColumns_(headers) {
  const normalizedHeaders = Array.isArray(headers) ? headers : [];
  const exactMap = (typeof headerMap_ === "function") ? headerMap_(normalizedHeaders) : {};
  const lowerHeaders = normalizedHeaders.map(function(header) {
    return StockWebApp_normalizeHeaderCandidate_(header);
  });

  const findCol = function(candidates) {
    for (let i = 0; i < candidates.length; i++) {
      const candidate = String(candidates[i] || "");
      const exact = exactMap[candidate];
      if (exact) return exact;
    }

    for (let i = 0; i < candidates.length; i++) {
      const wanted = StockWebApp_normalizeHeaderCandidate_(candidates[i]);
      if (!wanted) continue;
      const idx = lowerHeaders.indexOf(wanted);
      if (idx !== -1) return idx + 1;
    }

    return 0;
  };

  const cols = {
    reference: findCol(["货号", "ref", "réf", "référence", "reference"]),
    warehouse: findCol(["仓库", "entrepot", "entrepôt", "warehouse"]),
    stockDisplay: findCol(["剩下 / RESTE", "reste", "剩下", "restant", "stock restant"])
  };

  const missing = [];
  if (!cols.reference) missing.push("reference (`货号`)");
  if (!cols.warehouse) missing.push("warehouse (`仓库`)");
  if (!cols.stockDisplay) missing.push("stockDisplay (`剩下 / RESTE`)");

  if (missing.length) {
    throw new Error("Colonnes introuvables dans STOCK: " + missing.join(", "));
  }

  return cols;
}

function StockWebApp_buildItem_(row, cols, rowIndex) {
  const reference = StockWebApp_normalizeReference_(row[cols.reference - 1]);
  if (!reference) return null;

  const warehouse = StockWebApp_toDisplayText_(row[cols.warehouse - 1]);
  const stockDisplay = StockWebApp_toDisplayText_(row[cols.stockDisplay - 1]);

  return {
    id: [reference, warehouse || "NO_WAREHOUSE", String(rowIndex || 0)].join("__"),
    reference: reference,
    warehouse: warehouse,
    stockDisplay: stockDisplay
  };
}

function StockWebApp_normalizeReference_(value) {
  if (typeof cleanRef_ === "function") {
    return cleanRef_(value).toUpperCase();
  }
  return String(value || "").trim().toUpperCase();
}

function StockWebApp_toDisplayText_(value) {
  return String(value === null || typeof value === "undefined" ? "" : value).trim();
}

function StockWebApp_normalizeHeaderCandidate_(value) {
  return String(value === null || typeof value === "undefined" ? "" : value)
    .trim()
    .replace(/[’`´]/g, "'")
    .replace(/\s+/g, " ")
    .toLowerCase();
}
