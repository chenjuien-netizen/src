const INITIAL_SERVER_RENDER_COUNT = 100;

function doGet() {
  const initialResult = StockWebApp_collectPayload_(INITIAL_SERVER_RENDER_COUNT);
  Logger.log(
    "StockWebApp_doGet | initialItems=%s | rowsFetched=%s | readRange=%s:%s | timingsMs=%s",
    initialResult.payload.items.length,
    initialResult.meta.rowsFetched,
    initialResult.meta.minCol,
    initialResult.meta.maxCol,
    JSON.stringify(initialResult.meta.timings)
  );

  const template = HtmlService.createTemplateFromFile("StockMobile");
  template.initialBootJson = StockWebApp_safeJsonForTemplate_({
    payload: initialResult.payload,
    source: "server",
    generatedAt: new Date().toISOString()
  });
  template.initialListMarkup = StockWebApp_renderItemsHtml_(initialResult.payload.items);
  template.initialCountText = String(initialResult.payload.items.length || 0);
  template.initialStatusText = initialResult.payload.items.length
    ? initialResult.payload.items.length + " référence" + (initialResult.payload.items.length > 1 ? "s" : "") + " affichée" + (initialResult.payload.items.length > 1 ? "s" : "") + "."
    : "";

  return template.evaluate()
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
  const result = StockWebApp_collectPayload_(null);
  const headers = result.meta.headers;
  const cols = result.meta.cols;

  if (!result.payload.items.length) {
    Logger.log(
      "StockWebApp_getList_ | no items returned | sheet=%s | headers=%s | columns=%s | missingOptional=%s | missingRaw=%s | rowsRead=%s | readRange=%s:%s | timingsMs=%s | sampleReferenceValues=%s",
      result.meta.sheetName,
      JSON.stringify(headers),
      JSON.stringify(StockWebApp_logColumns_(headers, cols)),
      JSON.stringify(StockWebApp_missingOptionalColumns_(cols)),
      JSON.stringify(StockWebApp_missingRawColumns_(cols)),
      result.meta.rowsFetched,
      result.meta.minCol,
      result.meta.maxCol,
      JSON.stringify(result.meta.timings),
      JSON.stringify(result.meta.sampleReferenceValues)
    );
  }

  Logger.log(
    "StockWebApp_getList_ | sheet=%s | headers=%s | columns=%s | missingOptional=%s | missingRaw=%s | rowsRead=%s | readRange=%s:%s | itemsReturned=%s | timingsMs=%s",
    result.meta.sheetName,
    JSON.stringify(headers),
    JSON.stringify(StockWebApp_logColumns_(headers, cols)),
    JSON.stringify(StockWebApp_missingOptionalColumns_(cols)),
    JSON.stringify(StockWebApp_missingRawColumns_(cols)),
    result.meta.rowsFetched,
    result.meta.minCol,
    result.meta.maxCol,
    result.payload.items.length,
    JSON.stringify(result.meta.timings)
  );

  return result.payload;
}

function StockWebApp_collectPayload_(limit) {
  const totalStart = Date.now();
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

  const headersStart = Date.now();
  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const headersMs = Date.now() - headersStart;

  const resolveStart = Date.now();
  const cols = StockWebApp_resolveColumns_(headers);
  const resolveMs = Date.now() - resolveStart;

  const rowCount = Math.max(0, lastRow - 1);
  const usedCols = StockWebApp_collectUsedColumns_(cols);
  const minCol = usedCols.length ? Math.min.apply(null, usedCols) : 1;
  const maxCol = usedCols.length ? Math.max.apply(null, usedCols) : 1;
  const items = [];
  const sampleReferenceValues = [];
  let rowsFetched = 0;
  let fetchMs = 0;
  let buildMs = 0;

  if (rowCount > 0) {
    const width = maxCol - minCol + 1;
    if (limit && limit > 0) {
      const batchSize = Math.max(limit, 250);
      let startRow = 2;

      while (startRow <= lastRow && items.length < limit) {
        const rowsToRead = Math.min(batchSize, lastRow - startRow + 1);

        const fetchStart = Date.now();
        const values = sh.getRange(startRow, minCol, rowsToRead, width).getDisplayValues();
        fetchMs += Date.now() - fetchStart;
        rowsFetched += rowsToRead;

        const buildStart = Date.now();
        for (let i = 0; i < values.length && items.length < limit; i++) {
          if (sampleReferenceValues.length < 10) {
            sampleReferenceValues.push(StockWebApp_getCellByAbsCol_(values[i], cols.reference, minCol));
          }
          const item = StockWebApp_buildItem_(values[i], cols, startRow + i, minCol);
          if (item) items.push(item);
        }
        buildMs += Date.now() - buildStart;

        startRow += rowsToRead;
      }
    } else {
      const fetchStart = Date.now();
      const values = sh.getRange(2, minCol, rowCount, width).getDisplayValues();
      fetchMs += Date.now() - fetchStart;
      rowsFetched = rowCount;

      const buildStart = Date.now();
      for (let i = 0; i < values.length; i++) {
        if (sampleReferenceValues.length < 10) {
          sampleReferenceValues.push(StockWebApp_getCellByAbsCol_(values[i], cols.reference, minCol));
        }
        const item = StockWebApp_buildItem_(values[i], cols, i + 2, minCol);
        if (item) items.push(item);
      }
      buildMs += Date.now() - buildStart;
    }
  }

  return {
    payload: { items: items },
    meta: {
      sheetName: sh.getName(),
      headers: headers,
      cols: cols,
      rowCount: rowCount,
      rowsFetched: rowsFetched,
      minCol: minCol,
      maxCol: maxCol,
      sampleReferenceValues: sampleReferenceValues,
      timings: {
        headers: headersMs,
        resolve: resolveMs,
        fetch: fetchMs,
        build: buildMs,
        total: Date.now() - totalStart
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
    reference: findCol(["货号", "Reference", "Référence", "ref"]),
    stockDisplay: findCol(["剩下 / RESTE", "剩下", "RESTE"]),
    arrivageRef: findCol(["到货单", "arrivage", "arrivage id"]),
    warehouse: findCol(["仓库", "entrepot", "entrepôt"]),
    tailRaw: findCol(["尾箱"]),
    unitsPerBoxRaw: findCol(["件/箱", "每箱件数2"]),
    boxesRaw: findCol(["箱数"]),
    fractionRaw: findCol(["当前箱数分数"]),
    stockRaw: findCol(["stock"])
  };

  if (!cols.reference) {
    throw new Error("Colonne reference introuvable dans STOCK. Attendu: `货号`, `Reference`, `Référence` ou `ref`.");
  }

  return cols;
}

function StockWebApp_buildItem_(row, cols, rowIndex, baseCol) {
  const reference = StockWebApp_normalizeReference_(StockWebApp_getCellByAbsCol_(row, cols.reference, baseCol));
  if (!reference) return null;

  return {
    id: "row_" + String(rowIndex || 0),
    reference: reference,
    stockDisplay: StockWebApp_optionalText_(row, cols.stockDisplay, baseCol),
    arrivageRef: StockWebApp_optionalText_(row, cols.arrivageRef, baseCol),
    warehouse: StockWebApp_optionalText_(row, cols.warehouse, baseCol),
    stockState: StockWebApp_computeStockState_(row, cols, baseCol).state
  };
}

function StockWebApp_normalizeReference_(value) {
  if (typeof cleanRef_ === "function") {
    return cleanRef_(value).toUpperCase();
  }
  return String(value || "").trim().toUpperCase();
}

function StockWebApp_getCellByAbsCol_(row, absCol, baseCol) {
  if (!absCol) return "";
  const index = absCol - (baseCol || 1);
  if (index < 0 || index >= row.length) return "";
  const value = row[index];
  return value === null || typeof value === "undefined" ? "" : value;
}

function StockWebApp_optionalText_(row, col, baseCol) {
  if (!col) return "";
  return String(StockWebApp_getCellByAbsCol_(row, col, baseCol)).trim();
}

function StockWebApp_collectUsedColumns_(cols) {
  const out = [];
  Object.keys(cols || {}).forEach(function(key) {
    const col = cols[key];
    if (col && out.indexOf(col) === -1) out.push(col);
  });
  return out;
}

function StockWebApp_logColumns_(headers, cols) {
  return {
    reference: cols.reference ? (headers[cols.reference - 1] || "") : "",
    stockDisplay: cols.stockDisplay ? (headers[cols.stockDisplay - 1] || "") : "",
    arrivageRef: cols.arrivageRef ? (headers[cols.arrivageRef - 1] || "") : "",
    warehouse: cols.warehouse ? (headers[cols.warehouse - 1] || "") : "",
    tailRaw: cols.tailRaw ? (headers[cols.tailRaw - 1] || "") : "",
    unitsPerBoxRaw: cols.unitsPerBoxRaw ? (headers[cols.unitsPerBoxRaw - 1] || "") : "",
    boxesRaw: cols.boxesRaw ? (headers[cols.boxesRaw - 1] || "") : "",
    fractionRaw: cols.fractionRaw ? (headers[cols.fractionRaw - 1] || "") : "",
    stockRaw: cols.stockRaw ? (headers[cols.stockRaw - 1] || "") : ""
  };
}

function StockWebApp_missingOptionalColumns_(cols) {
  const missing = [];
  if (!cols.stockDisplay) missing.push("剩下 / RESTE");
  if (!cols.arrivageRef) missing.push("到货单");
  if (!cols.warehouse) missing.push("仓库");
  return missing;
}

function StockWebApp_missingRawColumns_(cols) {
  const missing = [];
  if (!cols.tailRaw) missing.push("尾箱");
  if (!cols.unitsPerBoxRaw) missing.push("件/箱|每箱件数2");
  if (!cols.boxesRaw) missing.push("箱数");
  if (!cols.fractionRaw) missing.push("当前箱数分数");
  if (!cols.stockRaw) missing.push("stock");
  return missing;
}

function StockWebApp_computeStockState_(row, cols, baseCol) {
  const tail = StockWebApp_parsePositiveNumber_(row, cols.tailRaw, baseCol);
  const unitsPerBox = StockWebApp_parsePositiveNumber_(row, cols.unitsPerBoxRaw, baseCol);
  const boxes = StockWebApp_parsePositiveNumber_(row, cols.boxesRaw, baseCol);
  const fractionPositive = StockWebApp_isPositiveFractionLike_(row, cols.fractionRaw, baseCol);

  if (tail > 0) return { state: "positive", source: "raw-columns" };
  if (unitsPerBox > 0 && boxes > 0) return { state: "positive", source: "raw-columns" };
  if (unitsPerBox > 0 && fractionPositive) return { state: "positive", source: "raw-columns" };

  if (cols.tailRaw || cols.unitsPerBoxRaw || cols.boxesRaw || cols.fractionRaw) {
    const stockNumeric = StockWebApp_parseFiniteNumber_(row, cols.stockRaw, baseCol);
    if (stockNumeric !== null) {
      return { state: stockNumeric > 0 ? "positive" : "zero", source: "stock" };
    }

    const displayNumber = StockWebApp_extractFirstNumber_(StockWebApp_optionalText_(row, cols.stockDisplay, baseCol));
    if (displayNumber !== null) {
      return { state: displayNumber > 0 ? "positive" : "zero", source: "rest-display" };
    }

    return { state: "zero", source: "default-zero" };
  }

  const stockNumeric = StockWebApp_parseFiniteNumber_(row, cols.stockRaw, baseCol);
  if (stockNumeric !== null) {
    return { state: stockNumeric > 0 ? "positive" : "zero", source: "stock" };
  }

  const displayNumber = StockWebApp_extractFirstNumber_(StockWebApp_optionalText_(row, cols.stockDisplay, baseCol));
  if (displayNumber !== null) {
    return { state: displayNumber > 0 ? "positive" : "zero", source: "rest-display" };
  }

  return { state: "zero", source: "default-zero" };
}

function StockWebApp_parseFiniteNumber_(row, col, baseCol) {
  if (!col) return null;
  const raw = StockWebApp_getCellByAbsCol_(row, col, baseCol);
  if (raw === null || typeof raw === "undefined" || String(raw).trim() === "") return null;
  const normalized = String(raw).trim().replace(",", ".");
  if (!/^[-+]?\d+(?:\.\d+)?$/.test(normalized)) return null;
  const numberValue = Number(normalized);
  return Number.isFinite(numberValue) ? numberValue : null;
}

function StockWebApp_parsePositiveNumber_(row, col, baseCol) {
  const numberValue = StockWebApp_parseFiniteNumber_(row, col, baseCol);
  return numberValue !== null && numberValue > 0 ? numberValue : 0;
}

function StockWebApp_isPositiveFractionLike_(row, col, baseCol) {
  const numberValue = StockWebApp_parseFiniteNumber_(row, col, baseCol);
  if (numberValue !== null) return numberValue > 0;
  if (!col) return false;

  const raw = StockWebApp_getCellByAbsCol_(row, col, baseCol);
  const normalized = String(raw === null || typeof raw === "undefined" ? "" : raw)
    .trim()
    .replace(/\s*\/\s*/g, "/");
  if (!normalized) return false;

  const match = normalized.match(/^(\d+)\/(\d+)$/);
  if (!match) return false;

  const numerator = Number(match[1]);
  const denominator = Number(match[2]);
  return Number.isFinite(numerator) &&
    Number.isFinite(denominator) &&
    numerator > 0 &&
    denominator > 0;
}

function StockWebApp_extractFirstNumber_(value) {
  const text = String(value || "").trim().replace(",", ".");
  if (!text) return null;
  const match = text.match(/[-+]?\d+(?:\.\d+)?/);
  if (!match) return null;
  const numberValue = Number(match[0]);
  return Number.isFinite(numberValue) ? numberValue : null;
}

function StockWebApp_normalizeHeaderCandidate_(value) {
  return String(value === null || typeof value === "undefined" ? "" : value)
    .trim()
    .replace(/[’`´]/g, "'")
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function StockWebApp_renderItemsHtml_(items) {
  return (Array.isArray(items) ? items : []).map(function(item) {
    return StockWebApp_renderItemHtml_(item);
  }).join("");
}

function StockWebApp_renderItemHtml_(item) {
  const itemClass = item && item.stockState === "positive"
    ? "item inventory-row stock-positive"
    : "item inventory-row stock-zero";
  const reference = item && item.reference ? item.reference : "-";
  const stockDisplay = item && item.stockDisplay ? item.stockDisplay : "";

  return '' +
    '<tr class="' + itemClass + '" data-item-id="' + StockWebApp_escapeHtml_(item && item.id ? item.id : "") + '" data-reference="' + StockWebApp_escapeHtml_(item && item.reference ? item.reference : "") + '" data-selection-mode="false" tabindex="0" role="button" aria-label="Copier ' + StockWebApp_escapeHtml_(reference) + '">' +
      '<td class="px-3 py-2 text-center">' +
        '<input class="item-checkbox h-3.5 w-3.5 rounded border-outline-variant text-primary focus:ring-primary/30" type="checkbox" aria-label="Sélectionner ' + StockWebApp_escapeHtml_(reference) + '">' +
      '</td>' +
      '<td class="px-3 py-2 truncate font-sans font-bold text-on-surface">' + StockWebApp_escapeHtml_(reference) + '</td>' +
      '<td class="px-3 py-2 text-right tabular-nums font-bold ' + (item && item.stockState === "zero" ? 'text-on-surface-variant' : 'text-primary') + '">' + StockWebApp_escapeHtml_(stockDisplay || "-") + '</td>' +
    '</tr>';
}

function StockWebApp_escapeHtml_(value) {
  return String(value === null || typeof value === "undefined" ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function StockWebApp_safeJsonForTemplate_(value) {
  return JSON.stringify(value)
    .replace(/</g, "\\u003c")
    .replace(/>/g, "\\u003e")
    .replace(/&/g, "\\u0026")
    .replace(/\u2028/g, "\\u2028")
    .replace(/\u2029/g, "\\u2029");
}
