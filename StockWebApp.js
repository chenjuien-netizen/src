const INITIAL_SERVER_RENDER_COUNT = 120;

function doGet() {
  const initialResult = StockWebApp_collectPayload_(INITIAL_SERVER_RENDER_COUNT);
  const template = HtmlService.createTemplateFromFile("StockMobile");

  template.initialBootJson = StockWebApp_safeJsonForTemplate_({
    payload: initialResult.payload,
    source: "server",
    generatedAt: new Date().toISOString()
  });
  template.initialListMarkup = StockWebApp_renderItemsHtml_(initialResult.payload.items);

  Logger.log(
    "StockWebApp_doGet | initialItems=%s | totalRows=%s | partial=%s | readRange=%s:%s | timingsMs=%s",
    initialResult.payload.items.length,
    initialResult.payload.summary.totalRows,
    initialResult.payload.summary.isPartial,
    initialResult.meta.minCol,
    initialResult.meta.maxCol,
    JSON.stringify(initialResult.meta.timings)
  );

  return template.evaluate()
    .setTitle("SZFashion | Inventaire")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, viewport-fit=cover")
    .addMetaTag("mobile-web-app-capable", "yes")
    .addMetaTag("apple-mobile-web-app-capable", "yes")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function StockWebApp_getList() {
  const result = StockWebApp_collectPayload_(null);

  Logger.log(
    "StockWebApp_getList | items=%s | totalRows=%s | readRange=%s:%s | timingsMs=%s",
    result.payload.items.length,
    result.payload.summary.totalRows,
    result.meta.minCol,
    result.meta.maxCol,
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
    throw new Error("La feuille STOCK ne contient aucun en-tete.");
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
  const width = maxCol - minCol + 1;
  const items = [];
  let rowsFetched = 0;
  let fetchMs = 0;
  let buildMs = 0;

  if (rowCount > 0 && width > 0) {
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
        const item = StockWebApp_buildItem_(values[i], cols, i + 2, minCol);
        if (item) items.push(item);
      }
      buildMs += Date.now() - buildStart;
    }
  }

  StockWebApp_sortItemsByReference_(items);

  const availableFilters = StockWebApp_collectAvailableFilters_(items);
  const generatedAt = new Date().toISOString();

  return {
    payload: {
      items: items,
      summary: StockWebApp_buildSummary_(items, rowCount, !!(limit && rowCount > items.length), generatedAt),
      filters: availableFilters,
      ui: {
        defaultSort: "ref_asc",
        availableSorts: [
          { value: "ref_asc", label: "REF ASC" },
          { value: "ref_desc", label: "REF DESC" },
          { value: "stock_first", label: "STOCK" }
        ]
      },
      generatedAt: generatedAt
    },
    meta: {
      headers: headers,
      cols: cols,
      rowsFetched: rowsFetched,
      minCol: minCol,
      maxCol: maxCol,
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
    tailRaw: findCol(["尾箱"]),
    unitsPerBoxRaw: findCol(["件/箱", "每箱件数2"]),
    boxesRaw: findCol(["箱数"]),
    signRaw: findCol(["当前signe"]),
    fractionRaw: findCol(["当前箱数分数"]),
    colisage: findCol(["Colisage"]),
    warehouse: findCol(["仓库", "entrepot", "entrepôt"]),
    createdAt: findCol(["date de création", "修改日期", "进货"])
  };

  if (!cols.reference) {
    throw new Error("Colonne reference introuvable dans STOCK. Attendu: `货号`, `Reference`, `Référence` ou `ref`.");
  }

  return cols;
}

function StockWebApp_buildItem_(row, cols, rowIndex, baseCol) {
  const reference = StockWebApp_normalizeReference_(StockWebApp_getCellByAbsCol_(row, cols.reference, baseCol));
  if (!reference) return null;

  const stateModel = StockWebApp_extractStateModel_(row, cols, baseCol);
  const stockDisplay = StockWebApp_buildStockDisplay_(stateModel);

  return {
    id: "row_" + String(rowIndex || 0),
    reference: reference,
    stockDisplay: stockDisplay,
    stockState: StockWebApp_computeStockStateFromModel_(stateModel),
    tail: stateModel.tail,
    unitsPerBox: stateModel.unitsPerBox,
    itemBoxes: stateModel.itemBoxes,
    fractionText: stateModel.fractionText,
    fractionValue: stateModel.fractionValue,
    colisage: stateModel.colisage,
    warehouse: StockWebApp_optionalText_(row, cols.warehouse, baseCol),
    createdAt: StockWebApp_optionalText_(row, cols.createdAt, baseCol)
  };
}

function StockWebApp_extractStateModel_(row, cols, baseCol) {
  const fractionRawValue = StockWebApp_getCellByAbsCol_(row, cols.fractionRaw, baseCol);
  const fractionText = StockWebApp_normalizeFractionText_(fractionRawValue);
  return {
    tail: StockWebApp_tailDisplayToTotal_(StockWebApp_getCellByAbsCol_(row, cols.tailRaw, baseCol)),
    unitsPerBox: StockWebApp_toInt_(StockWebApp_getCellByAbsCol_(row, cols.unitsPerBoxRaw, baseCol)),
    itemBoxes: StockWebApp_toInt_(StockWebApp_getCellByAbsCol_(row, cols.boxesRaw, baseCol)),
    sign: StockWebApp_normalizeSign_(StockWebApp_getCellByAbsCol_(row, cols.signRaw, baseCol)),
    fractionText: fractionText,
    fractionValue: StockWebApp_parseFractionValue_(fractionRawValue),
    colisage: cols.colisage ? StockWebApp_parsePositiveNumber_(StockWebApp_getCellByAbsCol_(row, cols.colisage, baseCol)) : 0
  };
}

function StockWebApp_buildStockDisplay_(stateInput) {
  const state = StockWebApp_normalizeStateModel_(stateInput || {});
  const tail = state.tail;
  const unitsPerBox = state.unitsPerBox;
  const itemBoxes = state.itemBoxes;
  const sign = state.sign;
  const fractionText = state.fractionText || StockWebApp_fractionToText_(state.fractionValue);

  let core = "";
  if (unitsPerBox > 0) {
    core = String(unitsPerBox) + "p";
    if (sign === "×" && fractionText) {
      if (itemBoxes > 1) {
        core += "×" + String(itemBoxes) + "×" + fractionText;
      } else {
        core += "×" + fractionText;
      }
    } else if (sign === "+" && fractionText) {
      core += "×" + String(itemBoxes) + "+" + fractionText;
    } else if (itemBoxes > 0) {
      core += "×" + String(itemBoxes);
    }
  }

  if (tail > 0) {
    return "(" + String(tail) + "p)" + (core ? "+" + core : "");
  }

  return core || "-";
}

function StockWebApp_normalizeStateModel_(stateInput) {
  const state = {
    tail: Math.max(0, StockWebApp_toInt_(stateInput.tail)),
    unitsPerBox: Math.max(0, StockWebApp_toInt_(stateInput.unitsPerBox)),
    itemBoxes: Math.max(0, StockWebApp_toInt_(stateInput.itemBoxes)),
    sign: StockWebApp_normalizeSign_(stateInput.sign),
    fractionText: StockWebApp_normalizeFractionText_(stateInput.fractionText),
    fractionValue: StockWebApp_parseFractionValue_(stateInput.fractionValue),
    colisage: StockWebApp_parsePositiveNumber_(stateInput.colisage)
  };

  if (!(state.fractionValue > 0)) {
    state.fractionValue = 0;
    state.fractionText = "";
    state.sign = "";
  } else if (!state.sign) {
    state.sign = state.itemBoxes > 0 ? "+" : "×";
  }

  if (!state.fractionText && state.fractionValue > 0) {
    state.fractionText = StockWebApp_fractionToText_(state.fractionValue);
  }

  if (state.sign === "×" && state.itemBoxes <= 0) {
    state.itemBoxes = 1;
  }

  if (state.sign === "+" && state.itemBoxes <= 0) {
    state.sign = "×";
    state.itemBoxes = 1;
  }

  return state;
}

function StockWebApp_computeStockStateFromModel_(stateInput) {
  const state = StockWebApp_normalizeStateModel_(stateInput || {});
  if (state.tail > 0) return "positive";
  if (state.unitsPerBox > 0 && state.itemBoxes > 0) return "positive";
  if (state.unitsPerBox > 0 && state.fractionValue > 0) return "positive";
  return "zero";
}

function StockWebApp_buildSummary_(items, totalRows, isPartial, generatedAt) {
  const positiveCount = (items || []).filter(function(item) {
    return item && item.stockState === "positive";
  }).length;

  return {
    visibleCount: (items || []).length,
    positiveCount: positiveCount,
    zeroCount: Math.max(0, (items || []).length - positiveCount),
    totalRows: totalRows || 0,
    isPartial: !!isPartial,
    generatedAt: generatedAt
  };
}

function StockWebApp_collectAvailableFilters_(items) {
  const warehouses = [];
  (items || []).forEach(function(item) {
    const value = String(item && item.warehouse ? item.warehouse : "").trim();
    if (value && warehouses.indexOf(value) === -1) warehouses.push(value);
  });

  warehouses.sort();

  return {
    warehouses: warehouses,
    stockStates: [
      { value: "all", label: "Tous" },
      { value: "positive", label: "En stock" },
      { value: "zero", label: "Zero" }
    ]
  };
}

function StockWebApp_sortItemsByReference_(items) {
  (items || []).sort(function(a, b) {
    const left = String(a && a.reference ? a.reference : "");
    const right = String(b && b.reference ? b.reference : "");
    return left.localeCompare(right);
  });
}

function StockWebApp_normalizeReference_(value) {
  const normalized = typeof cleanRef_ === "function"
    ? cleanRef_(value).toUpperCase()
    : String(value || "").trim().toUpperCase();

  return normalized
    .replace(/([A-Z0-9]+-\d+)(?=[A-Z0-9]+-\d+)/g, "$1 ")
    .replace(/\s+/g, " ")
    .trim();
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

function StockWebApp_tailDisplayToTotal_(value) {
  if (value === null || typeof value === "undefined" || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? Math.max(0, Math.trunc(value)) : 0;
  const matches = String(value).match(/-?\d+/g);
  if (!matches || !matches.length) return 0;
  return Math.max(0, matches.reduce(function(sum, part) {
    return sum + (Number(part) || 0);
  }, 0));
}

function StockWebApp_toInt_(value) {
  if (value === null || typeof value === "undefined" || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) ? Math.trunc(value) : 0;
  const match = String(value).trim().match(/-?\d+/);
  return match ? Number(match[0]) : 0;
}

function StockWebApp_parseFractionValue_(value) {
  if (value === null || typeof value === "undefined" || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) && value > 0 ? value : 0;

  const text = String(value).trim().replace(",", ".");
  if (!text) return 0;

  const fractionMatch = text.match(/^(\d+)\s*\/\s*(\d+)$/);
  if (fractionMatch) {
    const numerator = Number(fractionMatch[1]);
    const denominator = Number(fractionMatch[2]);
    if (!denominator) return 0;
    return numerator > 0 ? numerator / denominator : 0;
  }

  const numeric = Number(text);
  return Number.isFinite(numeric) && numeric > 0 ? numeric : 0;
}

function StockWebApp_parsePositiveNumber_(value) {
  if (value === null || typeof value === "undefined" || value === "") return 0;
  if (typeof value === "number") return Number.isFinite(value) && value > 0 ? value : 0;
  const normalized = String(value).trim().replace(",", ".");
  if (!normalized) return 0;
  const numberValue = Number(normalized);
  return Number.isFinite(numberValue) && numberValue > 0 ? numberValue : 0;
}

function StockWebApp_normalizeFractionText_(value) {
  const text = String(value === null || typeof value === "undefined" ? "" : value).trim().replace(/\s+/g, "");
  if (!text) return "";
  const match = text.match(/^(\d+)\/(\d+)$/);
  if (match) return match[1] + "/" + match[2];
  const numeric = Number(text.replace(",", "."));
  return Number.isFinite(numeric) && numeric > 0 ? StockWebApp_fractionToText_(numeric) : "";
}

function StockWebApp_normalizeSign_(value) {
  const sign = String(value || "").trim();
  return (sign === "+" || sign === "×" || sign === "x" || sign === "X") ? (sign === "+" ? "+" : "×") : "";
}

function StockWebApp_fractionToText_(value) {
  const numeric = Number(value || 0);
  if (!Number.isFinite(numeric) || numeric <= 0) return "";

  const candidates = [2, 3, 4, 6, 8, 12];
  let best = null;
  let bestErr = Infinity;

  for (let i = 0; i < candidates.length; i++) {
    const den = candidates[i];
    const num = Math.round(numeric * den);
    const approx = num / den;
    const err = Math.abs(numeric - approx);
    if (err < bestErr) {
      bestErr = err;
      best = { num: num, den: den };
    }
  }

  if (!best || !best.den || best.num <= 0) return "";
  const reduced = StockWebApp_reduceFraction_(best.num, best.den);
  return reduced.num + "/" + reduced.den;
}

function StockWebApp_reduceFraction_(num, den) {
  if (!den) return { num: 0, den: 1 };
  const divisor = StockWebApp_gcd_(num, den);
  return {
    num: num / divisor,
    den: den / divisor
  };
}

function StockWebApp_gcd_(a, b) {
  let left = Math.abs(Math.trunc(a || 0));
  let right = Math.abs(Math.trunc(b || 0));
  while (right) {
    const next = right;
    right = left % right;
    left = next;
  }
  return left || 1;
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
  const reference = item && item.reference ? item.reference : "-";
  const stockDisplay = item && item.stockDisplay ? item.stockDisplay : "-";
  const accentClass = item && item.stockState === "positive"
    ? "border-emerald-400/50"
    : "border-rose-400/50";
  const stockClass = item && item.stockState === "positive"
    ? "text-primary"
    : "text-on-surface-variant";

  return '' +
    '<article class="inventory-card bg-surface-container-lowest relative border-l-4 ' + accentClass + ' flex min-h-[4.25rem] flex-col justify-between px-2.5 py-2 transition-colors duration-150 hover:bg-surface-container" data-reference="' + StockWebApp_escapeHtml_(reference) + '" data-stock-display="' + StockWebApp_escapeHtml_(stockDisplay) + '" data-stock-state="' + StockWebApp_escapeHtml_(item && item.stockState ? item.stockState : "zero") + '">' +
      '<div class="flex items-start justify-between gap-2">' +
        '<span class="truncate pr-2 text-[12px] font-bold tracking-tight text-on-surface">' + StockWebApp_escapeHtml_(reference) + '</span>' +
        '<span class="material-symbols-outlined shrink-0 text-outline-variant !text-[14px]">inventory_2</span>' +
      '</div>' +
      '<div class="mt-2">' +
        '<span class="block truncate text-[13px] font-medium ' + stockClass + '">' + StockWebApp_escapeHtml_(stockDisplay) + '</span>' +
      '</div>' +
    '</article>';
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
