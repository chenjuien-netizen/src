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
  const referenceCol = StockWebApp_resolveReferenceColumn_(headers);
  const rowCount = Math.max(0, lastRow - 1);
  const values = rowCount ? sh.getRange(2, referenceCol, rowCount, 1).getDisplayValues().flat() : [];
  const items = [];

  for (let i = 0; i < values.length; i++) {
    const item = StockWebApp_buildReferenceItem_(values[i], i + 2);
    if (item) items.push(item);
  }

  if (!items.length) {
    Logger.log(
      "StockWebApp_getList_ | no items returned | sheet=%s | headers=%s | referenceColumn=%s | rowsRead=%s | sampleValues=%s",
      sh.getName(),
      JSON.stringify(headers),
      headers[referenceCol - 1] || "",
      rowCount,
      JSON.stringify(values.slice(0, 10))
    );
  }

  Logger.log(
    "StockWebApp_getList_ | sheet=%s | headers=%s | referenceColumn=%s | rowsRead=%s | itemsReturned=%s",
    sh.getName(),
    JSON.stringify(headers),
    headers[referenceCol - 1] || "",
    rowCount,
    items.length
  );

  return { items: items };
}

function StockWebApp_resolveReferenceColumn_(headers) {
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

  const referenceCol = findCol(["货号", "Reference", "Référence", "ref"]);
  if (!referenceCol) {
    throw new Error("Colonne reference introuvable dans STOCK. Attendu: `货号`, `Reference`, `Référence` ou `ref`.");
  }

  return referenceCol;
}

function StockWebApp_buildReferenceItem_(value, rowIndex) {
  const reference = StockWebApp_normalizeReference_(value);
  if (!reference) return null;

  return {
    id: "row_" + String(rowIndex || 0),
    reference: reference
  };
}

function StockWebApp_normalizeReference_(value) {
  if (typeof cleanRef_ === "function") {
    return cleanRef_(value).toUpperCase();
  }
  return String(value || "").trim().toUpperCase();
}

function StockWebApp_normalizeHeaderCandidate_(value) {
  return String(value === null || typeof value === "undefined" ? "" : value)
    .trim()
    .replace(/[’`´]/g, "'")
    .replace(/\s+/g, " ")
    .toLowerCase();
}
