function doGet() {
  return HtmlService.createHtmlOutputFromFile("StockMobile")
    .setTitle("SZFashion Stock")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, viewport-fit=cover")
    .addMetaTag("mobile-web-app-capable", "yes")
    .addMetaTag("apple-mobile-web-app-capable", "yes")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function StockWebApp_getList() {
  return StockWebApp_getList_();
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

  if (!items.length) {
    Logger.log(
      "StockWebApp_getList_ | no items returned | sheet=%s | headers=%s | columns=%s | missingOptional=%s | rowsRead=%s | sampleReferenceValues=%s",
      sh.getName(),
      JSON.stringify(headers),
      JSON.stringify(StockWebApp_logColumns_(headers, cols)),
      JSON.stringify(StockWebApp_missingOptionalColumns_(cols)),
      rowCount,
      JSON.stringify(values.slice(0, 10).map(function(row) {
        return row[cols.reference - 1];
      }))
    );
  }

  Logger.log(
    "StockWebApp_getList_ | sheet=%s | headers=%s | columns=%s | missingOptional=%s | rowsRead=%s | itemsReturned=%s",
    sh.getName(),
    JSON.stringify(headers),
    JSON.stringify(StockWebApp_logColumns_(headers, cols)),
    JSON.stringify(StockWebApp_missingOptionalColumns_(cols)),
    rowCount,
    items.length
  );

  return { items: items };
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
    reference: findCol(["货号", "Reference", "Référence", "ref"]),
    stockDisplay: findCol(["剩下 / RESTE", "剩下", "RESTE"]),
    arrivageRef: findCol(["到货单", "arrivage", "arrivage id"]),
    warehouse: findCol(["仓库", "entrepot", "entrepôt"])
  };

  if (!cols.reference) {
    throw new Error("Colonne reference introuvable dans STOCK. Attendu: `货号`, `Reference`, `Référence` ou `ref`.");
  }

  return cols;
}

function StockWebApp_buildItem_(row, cols, rowIndex) {
  const reference = StockWebApp_normalizeReference_(row[cols.reference - 1]);
  if (!reference) return null;

  return {
    id: "row_" + String(rowIndex || 0),
    reference: reference,
    stockDisplay: StockWebApp_optionalText_(row, cols.stockDisplay),
    arrivageRef: StockWebApp_optionalText_(row, cols.arrivageRef),
    warehouse: StockWebApp_optionalText_(row, cols.warehouse)
  };
}

function StockWebApp_normalizeReference_(value) {
  if (typeof cleanRef_ === "function") {
    return cleanRef_(value).toUpperCase();
  }
  return String(value || "").trim().toUpperCase();
}

function StockWebApp_optionalText_(row, col) {
  if (!col) return "";
  return String(row[col - 1] === null || typeof row[col - 1] === "undefined" ? "" : row[col - 1]).trim();
}

function StockWebApp_logColumns_(headers, cols) {
  return {
    reference: cols.reference ? (headers[cols.reference - 1] || "") : "",
    stockDisplay: cols.stockDisplay ? (headers[cols.stockDisplay - 1] || "") : "",
    arrivageRef: cols.arrivageRef ? (headers[cols.arrivageRef - 1] || "") : "",
    warehouse: cols.warehouse ? (headers[cols.warehouse - 1] || "") : ""
  };
}

function StockWebApp_missingOptionalColumns_(cols) {
  const missing = [];
  if (!cols.stockDisplay) missing.push("剩下 / RESTE");
  if (!cols.arrivageRef) missing.push("到货单");
  if (!cols.warehouse) missing.push("仓库");
  return missing;
}

function StockWebApp_normalizeHeaderCandidate_(value) {
  return String(value === null || typeof value === "undefined" ? "" : value)
    .trim()
    .replace(/[’`´]/g, "'")
    .replace(/\s+/g, " ")
    .toLowerCase();
}
