const INITIAL_SERVER_RENDER_COUNT = 80;
const STOCK_WEBAPP_HISTORY_SHEET = "STOCK_HISTORY";
const STOCK_WEBAPP_STOCK_EXTRA_COLUMNS = ["Notation paquets"];

function StockWebApp_legacyDoGet() {
  const initialResult = StockWebApp_collectPayload_(INITIAL_SERVER_RENDER_COUNT);
  const template = HtmlService.createTemplateFromFile("StockMobile");
  const bootPayload = StockWebApp_buildInitialBootPayload_(initialResult.payload);

  template.initialBootJson = StockWebApp_safeJsonForTemplate_(bootPayload);
  template.initialListMarkup = StockWebApp_renderColumnLayoutHtml_(initialResult.payload.items, 2);

  Logger.log(
    "StockWebApp_doGet | bootMode=hybrid | initialItems=%s | totalRows=%s | partial=%s | readRange=%s:%s | timingsMs=%s",
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

function StockWebAppV2_getInventory() {
  const result = StockWebApp_collectPayload_(null);

  Logger.log(
    "StockWebAppV2_getInventory | items=%s | totalRows=%s | readRange=%s:%s | timingsMs=%s",
    result.payload.items.length,
    result.payload.summary.totalRows,
    result.meta.minCol,
    result.meta.maxCol,
    JSON.stringify(result.meta.timings)
  );

  return result.payload;
}

function StockWebApp_getList() {
  return StockWebAppV2_getInventory();
}

function StockWebApp_buildInitialBootPayload_(initialPayload) {
  const payload = initialPayload || {};
  const items = Array.isArray(payload.items) ? payload.items : [];
  return {
    payload: payload,
    loadingInventory: false,
    needsInventoryHydration: !!(payload.summary && payload.summary.isPartial),
    initialCount: items.length,
    source: "server-hybrid",
    generatedAt: new Date().toISOString()
  };
}

function StockWebApp_getHistory(input) {
  const payload = input || {};
  const result = StockWebApp_collectHistoryPayload_(payload);
  return result;
}

function StockWebApp_getItemDetail(reference, input) {
  const normalizedReference = StockWebApp_normalizeReference_(reference || "");
  if (!normalizedReference) {
    return {
      item: null,
      history: [],
      nextHistoryOffset: 0,
      hasMoreHistory: false,
      generatedAt: new Date().toISOString(),
      lastMovementAt: "",
      notFoundInStock: true
    };
  }

  const item = StockWebApp_findItemByReference_(normalizedReference);
  const historyPayload = StockWebApp_collectHistoryPayload_({
    reference: normalizedReference,
    offset: input && input.historyOffset,
    limit: input && input.historyLimit
  });

  return {
    item: item,
    history: historyPayload.items,
    nextHistoryOffset: historyPayload.nextOffset,
    hasMoreHistory: historyPayload.hasMore,
    generatedAt: new Date().toISOString(),
    lastMovementAt: historyPayload.items.length ? historyPayload.items[0].timestampRaw : "",
    notFoundInStock: !item
  };
}

function StockWebApp_saveQuickEdit(input) {
  const payload = input || {};
  const mode = String(payload.mode || "").trim();
  const itemId = String(payload.id || "").trim();
  const inputReference = StockWebApp_normalizeReference_(payload.reference || "");
  const rowIndex = StockWebApp_parseRowId_(itemId);

  if (!rowIndex || rowIndex < 2) {
    throw new Error("Ligne invalide. Merci de rafraichir la liste.");
  }
  if (!inputReference) {
    throw new Error("Reference invalide. Merci de rafraichir la liste.");
  }
  if (mode !== "edit" && mode !== "quick-exit") {
    throw new Error("Mode de sauvegarde invalide.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_STOCK);
  if (!sh) {
    throw new Error("Feuille introuvable: " + SHEET_STOCK);
  }

  StockWebApp_ensureStockColumns_(sh);

  const lastCol = sh.getLastColumn();
  if (lastCol < 1) {
    throw new Error("La feuille STOCK ne contient aucun en-tete.");
  }

  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const cols = StockWebApp_resolveColumns_(headers);
  StockWebApp_assertQuickEditColumns_(cols);

  const usedCols = StockWebApp_collectUsedColumns_(cols);
  const minCol = usedCols.length ? Math.min.apply(null, usedCols) : 1;
  const maxCol = usedCols.length ? Math.max.apply(null, usedCols) : 1;
  const width = maxCol - minCol + 1;

  const currentRow = sh.getRange(rowIndex, minCol, 1, width).getDisplayValues()[0];
  const currentReference = StockWebApp_normalizeReference_(StockWebApp_getCellByAbsCol_(currentRow, cols.reference, minCol));
  if (!currentReference) {
    throw new Error("La ligne cible est introuvable. Merci de rafraichir la liste.");
  }
  if (currentReference !== inputReference) {
    throw new Error("La ligne a change depuis l'ouverture. Merci de rafraichir la liste.");
  }

  const currentItem = StockWebApp_buildItem_(currentRow, cols, rowIndex, minCol);
  if (!currentItem) {
    throw new Error("Impossible de relire la reference cible.");
  }

  const nextState = mode === "edit"
    ? StockWebApp_buildEditStateFromInput_(payload)
    : StockWebApp_buildQuickExitStateFromInput_(payload, currentItem);

  StockWebApp_writeQuickEditState_(sh, rowIndex, cols, nextState);

  const refreshedRow = sh.getRange(rowIndex, minCol, 1, width).getDisplayValues()[0];
  const refreshedItem = StockWebApp_buildItem_(refreshedRow, cols, rowIndex, minCol);
  if (!refreshedItem) {
    throw new Error("Impossible de relire la reference apres sauvegarde.");
  }

  StockWebApp_appendHistoryEntry_(ss, {
    actionType: mode === "quick-exit" ? "sortie_rapide" : "modifier",
    rowId: itemId,
    reference: refreshedItem.reference,
    beforeItem: currentItem,
    afterItem: refreshedItem,
    remark: nextState.remark || "",
    source: "stock_mobile_quick_edit"
  });

  return {
    item: refreshedItem,
    savedAt: new Date().toISOString(),
    message: mode === "quick-exit" ? "Sortie appliquee" : "Modification enregistree"
  };
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

  StockWebApp_sortItems_(items, "ref_asc");

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
    sortKey: findCol(["SortKey"]),
    tailRaw: findCol(["尾箱"]),
    unitsPerBoxRaw: findCol(["件/箱", "每箱件数2"]),
    boxesRaw: findCol(["箱数"]),
    signRaw: findCol(["当前signe"]),
    fractionRaw: findCol(["当前箱数分数"]),
    colisage: findCol(["Colisage"]),
    packNotation: findCol(["Notation paquets"]),
    remark: findCol(["放位/提醒"]),
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
  const packMeta = StockWebApp_buildPackMeta_(stateModel);
  const stockDisplay = StockWebApp_buildStockDisplay_(stateModel);

  return {
    id: "row_" + String(rowIndex || 0),
    reference: reference,
    sortKey: StockWebApp_optionalText_(row, cols.sortKey, baseCol),
    stockDisplay: stockDisplay,
    stockState: StockWebApp_computeStockStateFromModel_(stateModel),
    tail: stateModel.tail,
    unitsPerBox: stateModel.unitsPerBox,
    itemBoxes: stateModel.itemBoxes,
    sign: stateModel.sign,
    fractionText: stateModel.fractionText,
    fractionValue: stateModel.fractionValue,
    colisage: stateModel.colisage,
    packNotation: stateModel.packNotation,
    remark: stateModel.remark,
    packsPerBox: packMeta.packsPerBox,
    packCounterText: packMeta.packCounterText,
    dynamicFractions: packMeta.dynamicFractions,
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
    colisage: cols.colisage ? StockWebApp_parsePositiveNumber_(StockWebApp_getCellByAbsCol_(row, cols.colisage, baseCol)) : 0,
    packNotation: cols.packNotation ? StockWebApp_normalizePackNotation_(StockWebApp_getCellByAbsCol_(row, cols.packNotation, baseCol), false) : "",
    remark: cols.remark ? String(StockWebApp_getCellByAbsCol_(row, cols.remark, baseCol)).trim() : ""
  };
}

function StockWebApp_buildStockDisplay_(stateInput) {
  const state = StockWebApp_normalizeStateModel_(stateInput || {});
  const rawDisplay = StockWebApp_buildRawStockDisplay_(state);
  const mainPackNotation = StockWebApp_getMainPackNotationFromState_(state);
  if (!mainPackNotation) return rawDisplay;
  return rawDisplay === "-" ? mainPackNotation : (rawDisplay + mainPackNotation);
}

function StockWebApp_buildRawStockDisplay_(stateInput) {
  const state = StockWebApp_normalizeStateModel_(stateInput || {});
  const tailDisplay = StockWebApp_buildTailDisplay_(state);
  const unitsPerBox = state.unitsPerBox;
  const itemBoxes = state.itemBoxes;
  const sign = state.sign;
  const fractionText = state.fractionText || StockWebApp_fractionToText_(state.fractionValue);
  const hasMainContent = itemBoxes > 0 || !!fractionText;

  let core = "";
  if (unitsPerBox > 0 && hasMainContent) {
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

  if (tailDisplay) {
    return tailDisplay + (core ? "+" + core : "");
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
    colisage: StockWebApp_parsePositiveNumber_(stateInput.colisage),
    packNotation: StockWebApp_normalizePackNotation_(stateInput.packNotation, false),
    remark: String(stateInput.remark || "").trim()
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

function StockWebApp_buildEditStateFromInput_(payload) {
  const tail = Math.max(0, StockWebApp_toInt_(payload.tail));
  const unitsPerBox = Math.max(0, StockWebApp_toInt_(payload.unitsPerBox));
  const itemBoxes = Math.max(0, StockWebApp_toInt_(payload.itemBoxes));
  const sign = StockWebApp_normalizeSign_(payload.sign);
  const rawFractionText = String(payload.fractionText || "").trim().replace(/\s+/g, "");
  const fractionText = rawFractionText ? StockWebApp_normalizeFractionText_(rawFractionText) : "";

  if (rawFractionText && !/^\d+\/\d+$/.test(rawFractionText)) {
    throw new Error("Fraction invalide. Utilise un format du type 1/2.");
  }
  if (fractionText) {
    const parts = fractionText.split("/");
    const numerator = Number(parts[0]);
    const denominator = Number(parts[1]);
    if (!(numerator > 0) || !(denominator > 0)) {
      throw new Error("Fraction invalide. Utilise un format du type 1/2.");
    }
  }

  return StockWebApp_normalizeStateModel_({
    tail: tail,
    unitsPerBox: unitsPerBox,
    itemBoxes: itemBoxes,
    sign: sign,
    fractionText: fractionText,
    fractionValue: StockWebApp_parseFractionValue_(fractionText),
    packNotation: StockWebApp_buildPackNotationFromParts_(payload.packNotationSign, payload.packNotationCount, payload.packNotation),
    remark: String(payload.remark || "").trim()
  });
}

function StockWebApp_buildQuickExitStateFromInput_(payload, currentItem) {
  const currentState = StockWebApp_normalizeStateModel_({
    tail: currentItem.tail,
    unitsPerBox: currentItem.unitsPerBox,
    itemBoxes: currentItem.itemBoxes,
    sign: currentItem.sign,
    fractionText: currentItem.fractionText,
    fractionValue: currentItem.fractionValue,
    colisage: currentItem.colisage,
    packNotation: currentItem.packNotation,
    remark: currentItem.remark
  });

  const unitsPerBox = currentState.unitsPerBox;
  const totalPieces = StockWebApp_stateModelToPieces_(currentState);
  const exitMode = String(payload.exitMode || "").trim();
  const exitValue = Math.max(0, StockWebApp_toInt_(payload.exitValue));
  const exitFractionText = StockWebApp_normalizeFractionText_(payload.exitFractionText);
  let exitPieces = 0;

  if (unitsPerBox <= 0) {
    throw new Error("Sortie rapide impossible sans 件/箱.");
  }

  if (exitMode === "boxes") {
    if (!(exitValue > 0)) throw new Error("Indique un nombre de cartons positif.");
    exitPieces = unitsPerBox * exitValue;
  } else if (exitMode === "fraction") {
    const fractionValue = StockWebApp_parseFractionValue_(exitFractionText);
    if (!(fractionValue > 0) || !exitFractionText) {
      throw new Error("Selectionne une fraction valide.");
    }
    exitPieces = unitsPerBox * fractionValue;
  } else if (exitMode === "packs") {
    const colisage = StockWebApp_parsePositiveNumber_(currentItem.colisage);
    if (!(colisage > 0)) {
      throw new Error("Mode paquets indisponible pour cette reference.");
    }
    if (!(exitValue > 0)) throw new Error("Indique un nombre de paquets positif.");
    exitPieces = colisage * exitValue;
  } else {
    throw new Error("Mode de sortie invalide.");
  }

  const newTotal = totalPieces - exitPieces;
  if (newTotal < 0) {
    throw new Error("La sortie depasse le stock disponible.");
  }

  return StockWebApp_buildStateFromPieces_(newTotal, {
    unitsPerBox: unitsPerBox,
    colisage: currentItem.colisage,
    remark: String(payload.remark || "").trim(),
    reconstructionMode: exitMode
  });
}

function StockWebApp_computeStockStateFromModel_(stateInput) {
  return StockWebApp_stateModelToPieces_(stateInput || {}) > 0 ? "positive" : "zero";
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

function StockWebApp_assertQuickEditColumns_(cols) {
  if (!cols.tailRaw || !cols.unitsPerBoxRaw || !cols.boxesRaw || !cols.signRaw || !cols.fractionRaw) {
    throw new Error("Colonnes quick edit introuvables dans STOCK.");
  }
}

function StockWebApp_parseRowId_(value) {
  const match = String(value || "").match(/^row_(\d+)$/);
  return match ? Number(match[1]) : 0;
}

function StockWebApp_writeQuickEditState_(sheet, rowIndex, cols, state) {
  const normalized = StockWebApp_normalizeStateModel_(state || {});
  sheet.getRange(rowIndex, cols.tailRaw).setValue(normalized.tail);
  sheet.getRange(rowIndex, cols.unitsPerBoxRaw).setValue(normalized.unitsPerBox);
  sheet.getRange(rowIndex, cols.boxesRaw).setValue(normalized.itemBoxes);
  sheet.getRange(rowIndex, cols.signRaw).setValue(normalized.sign || "");
  if (cols.packNotation) {
    sheet.getRange(rowIndex, cols.packNotation).setValue(normalized.packNotation || "");
  }
  if (cols.remark) {
    sheet.getRange(rowIndex, cols.remark).setValue(normalized.remark || "");
  }

  const fractionCell = sheet.getRange(rowIndex, cols.fractionRaw);
  if (normalized.fractionText) {
    fractionCell.setNumberFormat("@");
    fractionCell.setValue(normalized.fractionText);
  } else {
    fractionCell.clearContent();
    fractionCell.setNumberFormat("@");
  }
}

function StockWebApp_sortItems_(items, sortMode) {
  (items || []).sort(function(a, b) {
    if (sortMode === "ref_desc") {
      return -StockWebApp_compareBySortKey_(a, b);
    }

    if (sortMode === "stock_first") {
      const leftState = a && a.stockState === "positive" ? 0 : 1;
      const rightState = b && b.stockState === "positive" ? 0 : 1;
      if (leftState !== rightState) return leftState - rightState;
    }

    return StockWebApp_compareBySortKey_(a, b);
  });
}

function StockWebApp_compareBySortKey_(a, b) {
  const leftSortKey = String(a && a.sortKey ? a.sortKey : "").trim();
  const rightSortKey = String(b && b.sortKey ? b.sortKey : "").trim();
  const leftReference = String(a && a.reference ? a.reference : "");
  const rightReference = String(b && b.reference ? b.reference : "");

  if (leftSortKey && rightSortKey && leftSortKey !== rightSortKey) {
    return leftSortKey.localeCompare(rightSortKey);
  }
  if (leftSortKey && !rightSortKey) return -1;
  if (!leftSortKey && rightSortKey) return 1;
  return leftReference.localeCompare(rightReference);
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

function StockWebApp_normalizePackNotation_(value, strict) {
  const text = String(value || "").trim();
  if (!text) return "";
  const match = text.match(/^([+-])\s*(\d+)\s*包$/i);
  if (!match) {
    if (strict) throw new Error("Notation paquets invalide. Utilise +N包 ou -N包.");
    return text;
  }
  const count = Math.max(0, Math.trunc(Number(match[2]) || 0));
  if (!(count > 0)) return "";
  return match[1] + count + "包";
}

function StockWebApp_normalizeTailRelativeNotation_(value) {
  const text = String(value || "").trim().replace(/^TAIL:/i, "");
  if (!text) return "";
  const packNotation = StockWebApp_normalizePackNotation_(text, false);
  if (/^-\d+包$/.test(packNotation)) return packNotation;
  const fractionText = StockWebApp_normalizeFractionText_(text);
  return fractionText && StockWebApp_parseFractionValue_(fractionText) > 0 ? fractionText : "";
}

function StockWebApp_splitCompositePackNotation_(value) {
  const text = String(value || "").trim();
  if (!text) return { tailNotation: "", mainNotation: "" };
  const parts = text.split("|");
  let tailNotation = "";
  let mainNotation = "";
  parts.forEach(function(part) {
    const token = String(part || "").trim();
    if (!token) return;
    if (/^TAIL:/i.test(token)) {
      if (!tailNotation) tailNotation = StockWebApp_normalizeTailRelativeNotation_(token);
      return;
    }
    const packNotation = StockWebApp_normalizePackNotation_(token, false);
    if (!mainNotation && /^([+-])\d+包$/.test(packNotation)) {
      mainNotation = packNotation;
    }
  });
  return {
    tailNotation: tailNotation,
    mainNotation: mainNotation
  };
}

function StockWebApp_buildCompositePackNotation_(tailNotation, mainNotation) {
  const normalizedTail = StockWebApp_normalizeTailRelativeNotation_(tailNotation);
  const normalizedMain = StockWebApp_normalizePackNotation_(mainNotation, false);
  if (normalizedTail && normalizedMain) return "TAIL:" + normalizedTail + "|" + normalizedMain;
  if (normalizedTail) return "TAIL:" + normalizedTail;
  return /^([+-])\d+包$/.test(normalizedMain) ? normalizedMain : "";
}

function StockWebApp_getTailNotationFromState_(stateInput) {
  return StockWebApp_splitCompositePackNotation_(stateInput && stateInput.packNotation).tailNotation;
}

function StockWebApp_getMainPackNotationFromState_(stateInput) {
  return StockWebApp_splitCompositePackNotation_(stateInput && stateInput.packNotation).mainNotation;
}

function StockWebApp_getTailActualPieces_(stateInput) {
  const tailBase = Math.max(0, StockWebApp_toInt_(stateInput && stateInput.tail));
  const tailNotation = StockWebApp_getTailNotationFromState_(stateInput);
  if (!(tailBase > 0) || !tailNotation) return tailBase;
  if (/^-\d+包$/.test(tailNotation)) {
    const packMeta = StockWebApp_parsePackNotation_(tailNotation);
    const colisage = Math.max(0, StockWebApp_toInt_(stateInput && stateInput.colisage));
    return Math.max(0, tailBase - (packMeta.count * colisage));
  }
  const fractionValue = StockWebApp_parseFractionValue_(tailNotation);
  if (!(fractionValue > 0)) return tailBase;
  return Math.max(0, tailBase * fractionValue);
}

function StockWebApp_buildTailDisplay_(stateInput) {
  const tailBase = Math.max(0, StockWebApp_toInt_(stateInput && stateInput.tail));
  if (!(tailBase > 0)) return "";
  const tailNotation = StockWebApp_getTailNotationFromState_(stateInput);
  return "(" + String(tailBase) + "p)" + (tailNotation || "");
}

function StockWebApp_buildPackNotationFromParts_(signValue, countValue, fallbackValue) {
  const sign = String(signValue || "").trim();
  const count = Math.max(0, StockWebApp_toInt_(countValue));
  if ((sign === "+" || sign === "-") && count > 0) {
    return sign + count + "包";
  }
  const composite = StockWebApp_splitCompositePackNotation_(fallbackValue);
  if (composite.tailNotation || composite.mainNotation) {
    return StockWebApp_buildCompositePackNotation_(composite.tailNotation, composite.mainNotation);
  }
  return StockWebApp_normalizePackNotation_(fallbackValue, true);
}

function StockWebApp_parsePackNotation_(value) {
  const normalized = StockWebApp_normalizePackNotation_(value, false);
  const match = normalized.match(/^([+-])(\d+)包$/);
  if (!match) {
    return { notation: normalized, sign: "", count: 0, valid: !normalized };
  }
  return {
    notation: normalized,
    sign: match[1],
    count: Number(match[2]) || 0,
    valid: true
  };
}

function StockWebApp_computePacksPerBox_(unitsPerBox, colisage) {
  const units = Math.max(0, StockWebApp_toInt_(unitsPerBox));
  const packSize = Math.max(0, StockWebApp_toInt_(colisage));
  if (!(units > 0) || !(packSize > 0)) return 0;
  if (units % packSize !== 0) return 0;
  const packsPerBox = units / packSize;
  return packsPerBox > 0 ? packsPerBox : 0;
}

function StockWebApp_stateModelToPieces_(stateInput) {
  const state = StockWebApp_normalizeStateModel_(stateInput || {});
  const unitsPerBox = state.unitsPerBox;
  const fractionValue = state.fractionValue;
  let totalPieces = StockWebApp_getTailActualPieces_(state);

  if (unitsPerBox > 0) {
    if (state.sign === "+") {
      totalPieces += (unitsPerBox * state.itemBoxes) + (unitsPerBox * fractionValue);
    } else if (state.sign === "×") {
      if (state.itemBoxes > 1) {
        totalPieces += unitsPerBox * state.itemBoxes * fractionValue;
      } else if (state.itemBoxes > 0 || fractionValue > 0) {
        totalPieces += unitsPerBox * fractionValue;
      }
    } else {
      totalPieces += unitsPerBox * state.itemBoxes;
    }
  }

  const packMeta = StockWebApp_parsePackNotation_(StockWebApp_getMainPackNotationFromState_(state));
  if (packMeta.count > 0 && state.colisage > 0) {
    totalPieces += (packMeta.sign === "-" ? -1 : 1) * (packMeta.count * state.colisage);
  }

  return totalPieces;
}

function StockWebApp_fractionTextFromPieces_(pieces, unitsPerBox) {
  const safePieces = Math.max(0, StockWebApp_toInt_(pieces));
  const safeUnits = Math.max(0, StockWebApp_toInt_(unitsPerBox));
  if (!(safePieces > 0) || !(safeUnits > 0)) return "";
  const reduced = StockWebApp_reduceFraction_(safePieces, safeUnits);
  return reduced.num + "/" + reduced.den;
}

function StockWebApp_buildSimpleStateFromPieces_(totalPiecesInput, options) {
  const totalPieces = Math.max(0, Number(totalPiecesInput || 0));
  const unitsPerBox = Math.max(0, StockWebApp_toInt_(options && options.unitsPerBox));
  const colisage = Math.max(0, StockWebApp_toInt_(options && options.colisage));
  const remark = String(options && options.remark || "").trim();

  if (!(unitsPerBox > 0)) {
    throw new Error("Sortie rapide impossible sans 件/箱.");
  }

  const wholeBoxes = Math.floor(totalPieces / unitsPerBox);
  const remainderPieces = Math.max(0, totalPieces - (wholeBoxes * unitsPerBox));

  if (!(remainderPieces > 0)) {
    return StockWebApp_normalizeStateModel_({
      tail: 0,
      unitsPerBox: unitsPerBox,
      itemBoxes: wholeBoxes,
      sign: "",
      fractionText: "",
      fractionValue: 0,
      colisage: colisage,
      packNotation: "",
      remark: remark
    });
  }

  return StockWebApp_normalizeStateModel_({
    tail: 0,
    unitsPerBox: unitsPerBox,
    itemBoxes: wholeBoxes > 0 ? wholeBoxes : 1,
    sign: wholeBoxes > 0 ? "+" : "×",
    fractionText: StockWebApp_fractionTextFromPieces_(remainderPieces, unitsPerBox),
    fractionValue: remainderPieces / unitsPerBox,
    colisage: colisage,
    packNotation: "",
    remark: remark
  });
}

function StockWebApp_buildPackFriendlyStateFromPieces_(totalPiecesInput, options) {
  const totalPieces = Math.max(0, Number(totalPiecesInput || 0));
  const unitsPerBox = Math.max(0, StockWebApp_toInt_(options && options.unitsPerBox));
  const colisage = Math.max(0, StockWebApp_toInt_(options && options.colisage));
  const remark = String(options && options.remark || "").trim();

  if (!(unitsPerBox > 0)) {
    throw new Error("Sortie rapide impossible sans 件/箱.");
  }

  const packsPerBox = StockWebApp_computePacksPerBox_(unitsPerBox, colisage);
  const wholeBoxes = Math.floor(totalPieces / unitsPerBox);
  const remainderPieces = Math.max(0, totalPieces - (wholeBoxes * unitsPerBox));

  if (!(remainderPieces > 0)) {
    return StockWebApp_buildSimpleStateFromPieces_(totalPieces, options);
  }

  if (packsPerBox > 0 && colisage > 0 && remainderPieces % colisage === 0) {
    const packCount = remainderPieces / colisage;
    return StockWebApp_normalizeStateModel_({
      tail: 0,
      unitsPerBox: unitsPerBox,
      itemBoxes: wholeBoxes > 0 ? wholeBoxes : 1,
      sign: wholeBoxes > 0 ? "+" : "×",
      fractionText: packCount + "/" + packsPerBox,
      fractionValue: packCount / packsPerBox,
      colisage: colisage,
      packNotation: "",
      remark: remark
    });
  }

  if (packsPerBox > 0 && colisage > 0) {
    const packCount = Math.floor(remainderPieces / colisage);
    const loosePieces = remainderPieces - (packCount * colisage);
    return StockWebApp_normalizeStateModel_({
      tail: 0,
      unitsPerBox: unitsPerBox,
      itemBoxes: wholeBoxes > 0 ? wholeBoxes : 1,
      sign: wholeBoxes > 0 ? "+" : "×",
      fractionText: StockWebApp_fractionTextFromPieces_(loosePieces || remainderPieces, unitsPerBox),
      fractionValue: (loosePieces || remainderPieces) / unitsPerBox,
      colisage: colisage,
      packNotation: loosePieces > 0 && packCount > 0 ? ("+" + packCount + "包") : "",
      remark: remark
    });
  }

  return StockWebApp_buildSimpleStateFromPieces_(totalPieces, options);
}

function StockWebApp_buildStateFromPieces_(totalPiecesInput, options) {
  const reconstructionMode = String(options && options.reconstructionMode || "").trim();
  return reconstructionMode === "packs"
    ? StockWebApp_buildPackFriendlyStateFromPieces_(totalPiecesInput, options)
    : StockWebApp_buildSimpleStateFromPieces_(totalPiecesInput, options);
}

function StockWebApp_buildPackMeta_(stateInput) {
  const state = StockWebApp_normalizeStateModel_(stateInput || {});
  const packsPerBox = StockWebApp_computePacksPerBox_(state.unitsPerBox, state.colisage);
  const dynamicFractions = [];
  let packCounterText = "";
  const baseFractions = { "1/2": true, "1/3": true, "1/4": true, "2/3": true };

  if (packsPerBox > 1) {
    for (let i = 1; i < packsPerBox; i++) {
      const value = i + "/" + packsPerBox;
      if (!baseFractions[value]) dynamicFractions.push(value);
    }
  }

  if (packsPerBox > 0) {
    const packMeta = StockWebApp_parsePackNotation_(StockWebApp_getMainPackNotationFromState_(state));
    const basePackCount = state.fractionValue > 0 ? (state.fractionValue * packsPerBox) : 0;
    const deltaPackCount = packMeta.count > 0 ? (packMeta.sign === "-" ? -packMeta.count : packMeta.count) : 0;
    const hasRelevantCounter = state.fractionValue > 0 || packMeta.count > 0;
    if (hasRelevantCounter) {
      const counter = Math.max(0, Math.floor(basePackCount + deltaPackCount));
      packCounterText = counter + "/" + packsPerBox;
    }
  }

  return {
    packsPerBox: packsPerBox,
    packCounterText: packCounterText,
    dynamicFractions: dynamicFractions
  };
}

function StockWebApp_ensureStockColumns_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0] : [];
  const existing = {};
  headers.forEach(function(header) {
    existing[String(header || "").trim()] = true;
  });

  const missing = STOCK_WEBAPP_STOCK_EXTRA_COLUMNS.filter(function(name) {
    return !existing[name];
  });

  if (!missing.length) {
    StockWebApp_applyStockTextFormats_(sheet, headers);
    return;
  }

  const insertAt = sheet.getLastColumn();
  sheet.insertColumnsAfter(insertAt || 1, missing.length);
  sheet.getRange(1, insertAt + 1, 1, missing.length).setValues([missing]);

  const refreshedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  StockWebApp_applyStockTextFormats_(sheet, refreshedHeaders);
}

function StockWebApp_collectHistoryPayload_(input) {
  const payload = input || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = StockWebApp_getOrCreateHistorySheet_(ss);
  const loadAll = !!payload.loadAll;
  const limit = loadAll
    ? Number.MAX_SAFE_INTEGER
    : Math.max(1, Math.min(50, StockWebApp_toInt_(payload.limit) || 20));
  const offset = Math.max(0, StockWebApp_toInt_(payload.offset));
  const query = StockWebApp_normalizeHeaderCandidate_(payload.query || "");
  const reference = StockWebApp_normalizeReference_(payload.reference || "");
  const actionType = String(payload.actionType || "").trim();
  const rows = StockWebApp_readHistoryRows_(sheet);
  const filtered = rows.filter(function(entry) {
    if (reference && entry.reference !== reference) return false;
    if (actionType && entry.actionType !== actionType) return false;
    if (!query) return true;
    const haystack = [
      entry.reference,
      entry.actionType,
      entry.remark,
      entry.source,
      entry.beforeDisplay,
      entry.afterDisplay
    ].map(StockWebApp_normalizeHeaderCandidate_).join(" ");
    return haystack.indexOf(query) !== -1;
  });
  const slice = filtered.slice(offset, offset + limit);
  const nextOffset = offset + slice.length;
  return {
    items: slice,
    nextOffset: nextOffset,
    hasMore: nextOffset < filtered.length,
    totalMatched: filtered.length,
    generatedAt: new Date().toISOString()
  };
}

function StockWebApp_findItemByReference_(reference) {
  const normalizedReference = StockWebApp_normalizeReference_(reference || "");
  if (!normalizedReference) return null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_STOCK);
  if (!sh) return null;
  StockWebApp_ensureStockColumns_(sh);

  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  if (lastCol < 1 || lastRow < 2) return null;

  const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const cols = StockWebApp_resolveColumns_(headers);
  const usedCols = StockWebApp_collectUsedColumns_(cols);
  const minCol = usedCols.length ? Math.min.apply(null, usedCols) : 1;
  const maxCol = usedCols.length ? Math.max.apply(null, usedCols) : 1;
  const width = maxCol - minCol + 1;
  const values = sh.getRange(2, minCol, lastRow - 1, width).getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const item = StockWebApp_buildItem_(values[i], cols, i + 2, minCol);
    if (item && item.reference === normalizedReference) return item;
  }
  return null;
}

function StockWebApp_readHistoryRows_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  const headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  const cols = StockWebApp_resolveHistoryColumns_(headers);
  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const displayValues = range.getDisplayValues();
  const rawValues = range.getValues();
  const rows = [];

  for (let i = rawValues.length - 1; i >= 0; i--) {
    const entry = StockWebApp_buildHistoryEntry_(displayValues[i], rawValues[i], cols);
    if (entry) rows.push(entry);
  }
  return rows;
}

function StockWebApp_resolveHistoryColumns_(headers) {
  const indexByKey = {};
  (Array.isArray(headers) ? headers : []).forEach(function(header, idx) {
    indexByKey[String(header || "").trim()] = idx;
  });
  return {
    timestamp: indexByKey.timestamp,
    actionType: indexByKey.action_type,
    reference: indexByKey.reference,
    rowId: indexByKey.row_id,
    beforeDisplay: indexByKey.before_display,
    afterDisplay: indexByKey.after_display,
    remark: indexByKey.remark,
    source: indexByKey.source,
    beforeTotalPieces: indexByKey.before_total_pieces,
    afterTotalPieces: indexByKey.after_total_pieces
  };
}

function StockWebApp_buildHistoryEntry_(displayRow, rawRow, cols) {
  if (!displayRow || !rawRow || !cols) return null;
  const reference = StockWebApp_normalizeReference_(displayRow[cols.reference]);
  const actionType = String(displayRow[cols.actionType] || "").trim();
  const timestampValue = rawRow[cols.timestamp];
  const timestampRaw = StockWebApp_historyTimestampRaw_(timestampValue);
  if (!reference && !actionType && !timestampRaw) return null;
  return {
    timestampRaw: timestampRaw,
    timestampLabel: StockWebApp_formatHistoryTimestamp_(timestampValue),
    actionType: actionType,
    reference: reference,
    rowId: String(displayRow[cols.rowId] || "").trim(),
    beforeDisplay: String(displayRow[cols.beforeDisplay] || "").trim(),
    afterDisplay: String(displayRow[cols.afterDisplay] || "").trim(),
    remark: String(displayRow[cols.remark] || "").trim(),
    source: String(displayRow[cols.source] || "").trim(),
    beforeTotalPieces: Number(rawRow[cols.beforeTotalPieces] || 0),
    afterTotalPieces: Number(rawRow[cols.afterTotalPieces] || 0)
  };
}

function StockWebApp_historyTimestampRaw_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.toISOString();
  }
  const text = String(value || "").trim();
  if (!text) return "";
  const date = new Date(text);
  return isNaN(date.getTime()) ? text : date.toISOString();
}

function StockWebApp_formatHistoryTimestamp_(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (!(date instanceof Date) || isNaN(date.getTime())) return String(value || "").trim();
  const timeZone = Session.getScriptTimeZone ? Session.getScriptTimeZone() : "Europe/Paris";
  return Utilities.formatDate(date, timeZone || "Europe/Paris", "dd/MM HH:mm");
}

function StockWebApp_getOrCreateHistorySheet_(spreadsheet) {
  const expectedHeaders = [[
    "timestamp",
    "action_type",
    "reference",
    "row_id",
    "before_display",
    "after_display",
    "before_tail",
    "before_units_per_box",
    "before_boxes",
    "before_sign",
    "before_fraction",
    "before_pack_notation",
    "after_tail",
    "after_units_per_box",
    "after_boxes",
    "after_sign",
    "after_fraction",
    "after_pack_notation",
    "remark",
    "source",
    "before_total_pieces",
    "after_total_pieces"
  ]];
  let sheet = spreadsheet.getSheetByName(STOCK_WEBAPP_HISTORY_SHEET);
  if (sheet) {
    const lastCol = Math.max(1, sheet.getLastColumn());
    const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    if (existingHeaders.length < expectedHeaders[0].length) {
      sheet.insertColumnsAfter(lastCol, expectedHeaders[0].length - existingHeaders.length);
    }
    sheet.getRange(1, 1, 1, expectedHeaders[0].length).setValues(expectedHeaders);
    sheet.setFrozenRows(1);
    StockWebApp_applyHistoryColumnFormats_(sheet, expectedHeaders[0]);
    StockWebApp_sanitizeHistorySignColumns_(sheet, expectedHeaders[0]);
    return sheet;
  }

  sheet = spreadsheet.insertSheet(STOCK_WEBAPP_HISTORY_SHEET);
  sheet.getRange(1, 1, 1, expectedHeaders[0].length).setValues(expectedHeaders);
  sheet.setFrozenRows(1);
  StockWebApp_applyHistoryColumnFormats_(sheet, expectedHeaders[0]);
  return sheet;
}

function StockWebApp_appendHistoryEntry_(spreadsheet, payload) {
  const sheet = StockWebApp_getOrCreateHistorySheet_(spreadsheet);
  const beforeItem = payload.beforeItem || {};
  const afterItem = payload.afterItem || {};
  const beforeTotalPieces = StockWebApp_stateModelToPieces_(beforeItem);
  const afterTotalPieces = StockWebApp_stateModelToPieces_(afterItem);
  const beforeSign = StockWebApp_normalizeHistorySign_(beforeItem.sign);
  const afterSign = StockWebApp_normalizeHistorySign_(afterItem.sign);
  sheet.appendRow([
    new Date(),
    String(payload.actionType || ""),
    String(payload.reference || ""),
    String(payload.rowId || ""),
    String(beforeItem.stockDisplay || ""),
    String(afterItem.stockDisplay || ""),
    Number(beforeItem.tail || 0),
    Number(beforeItem.unitsPerBox || 0),
    Number(beforeItem.itemBoxes || 0),
    beforeSign,
    String(beforeItem.fractionText || ""),
    String(beforeItem.packNotation || ""),
    Number(afterItem.tail || 0),
    Number(afterItem.unitsPerBox || 0),
    Number(afterItem.itemBoxes || 0),
    afterSign,
    String(afterItem.fractionText || ""),
    String(afterItem.packNotation || ""),
    String(payload.remark || ""),
    String(payload.source || ""),
    Number(beforeTotalPieces || 0),
    Number(afterTotalPieces || 0)
  ]);
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
    if (!(numerator > 0) || !(denominator > 0)) return 0;
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

function StockWebApp_normalizeHistorySign_(value) {
  return StockWebApp_normalizeSign_(value);
}

function StockWebApp_applyStockTextFormats_(sheet, headers) {
  if (!sheet) return;
  const headerRow = Array.isArray(headers) ? headers : [];
  if (!headerRow.length) return;
  const cols = StockWebApp_resolveColumns_(headerRow);
  const textCols = [
    cols.reference,
    cols.signRaw,
    cols.fractionRaw,
    cols.packNotation,
    cols.remark
  ].filter(function(col) {
    return !!col;
  });
  StockWebApp_applyTextFormatToColumns_(sheet, textCols);
}

function StockWebApp_applyHistoryColumnFormats_(sheet, headers) {
  if (!sheet) return;
  const headerRow = Array.isArray(headers) ? headers : [];
  if (!headerRow.length) return;
  const cols = StockWebApp_resolveHistoryColumns_(headerRow);
  const textCols = [
    cols.actionType,
    cols.reference,
    cols.rowId,
    cols.beforeDisplay,
    cols.afterDisplay,
    headerRow.indexOf("before_sign") + 1,
    headerRow.indexOf("before_fraction") + 1,
    headerRow.indexOf("before_pack_notation") + 1,
    headerRow.indexOf("after_sign") + 1,
    headerRow.indexOf("after_fraction") + 1,
    headerRow.indexOf("after_pack_notation") + 1,
    cols.remark,
    cols.source
  ].filter(function(col) {
    return !!col;
  });
  StockWebApp_applyTextFormatToColumns_(sheet, textCols);
}

function StockWebApp_applyTextFormatToColumns_(sheet, columns) {
  if (!sheet) return;
  const rowCount = Math.max(0, sheet.getMaxRows() - 1);
  if (!(rowCount > 0)) return;
  const uniqueCols = {};
  (Array.isArray(columns) ? columns : []).forEach(function(col) {
    const numericCol = Math.max(0, StockWebApp_toInt_(col));
    if (numericCol > 0) uniqueCols[numericCol] = true;
  });
  Object.keys(uniqueCols).forEach(function(key) {
    sheet.getRange(2, Number(key), rowCount, 1).setNumberFormat("@");
  });
}

function StockWebApp_sanitizeHistorySignColumns_(sheet, headers) {
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const headerRow = Array.isArray(headers) ? headers : [];
  const beforeCol = headerRow.indexOf("before_sign") + 1;
  const afterCol = headerRow.indexOf("after_sign") + 1;
  [beforeCol, afterCol].forEach(function(col) {
    if (!(col > 0)) return;
    const range = sheet.getRange(2, col, lastRow - 1, 1);
    const displayValues = range.getDisplayValues();
    const formulas = range.getFormulas();
    let needsRewrite = false;
    const nextValues = displayValues.map(function(row, idx) {
      const currentDisplay = row && row.length ? row[0] : "";
      const currentFormula = formulas[idx] && formulas[idx].length ? formulas[idx][0] : "";
      const normalizedSign = StockWebApp_normalizeHistorySign_(currentDisplay);
      if (currentFormula || String(currentDisplay || "").trim() !== normalizedSign) {
        needsRewrite = true;
      }
      return [normalizedSign];
    });
    if (needsRewrite) {
      range.setNumberFormat("@");
      range.setValues(nextValues);
    }
  });
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

function StockWebApp_renderColumnLayoutHtml_(items, columnCount) {
  const normalizedCount = Math.max(1, Math.trunc(Number(columnCount) || 1));
  const sourceItems = Array.isArray(items) ? items : [];
  const perColumn = Math.ceil(sourceItems.length / normalizedCount) || 1;
  const columns = [];

  for (let i = 0; i < normalizedCount; i++) {
    const start = i * perColumn;
    const slice = sourceItems.slice(start, start + perColumn);
    if (!slice.length) continue;
    columns.push(
      '<div class="inventory-column flex min-w-0 flex-1 flex-col gap-px bg-outline-variant/20">' +
        StockWebApp_renderItemsHtml_(slice) +
      '</div>'
    );
  }

  return columns.length
    ? '<div class="flex items-start gap-px bg-outline-variant/20">' + columns.join("") + '</div>'
    : "";
}

function StockWebApp_renderItemHtml_(item) {
  const reference = item && item.reference ? item.reference : "-";
  const stockDisplay = item && item.stockDisplay ? item.stockDisplay : "-";
  const itemId = item && item.id ? item.id : reference;
  const accentClass = item && item.stockState === "positive"
    ? "border-emerald-400/50"
    : "border-rose-400/50";
  const stockClass = item && item.stockState === "positive"
    ? "text-primary"
    : "text-on-surface-variant";

  return '' +
    '<article class="inventory-card bg-surface-container-lowest relative border-l-4 ' + accentClass + ' flex min-h-[4.1rem] items-stretch transition-colors duration-150 hover:bg-surface-container select-none" data-item-id="' + StockWebApp_escapeHtml_(itemId) + '" data-reference="' + StockWebApp_escapeHtml_(reference) + '" data-stock-display="' + StockWebApp_escapeHtml_(stockDisplay) + '" data-stock-state="' + StockWebApp_escapeHtml_(item && item.stockState ? item.stockState : "zero") + '">' +
      '<button class="inventory-card-main flex min-w-0 flex-1 flex-col justify-between px-2.5 py-2 text-left" type="button" data-action="open-quick-edit" data-item-id="' + StockWebApp_escapeHtml_(itemId) + '">' +
        '<div class="flex items-start gap-2">' +
          '<span class="truncate pr-2 text-[12px] font-bold tracking-tight text-on-surface">' + StockWebApp_escapeHtml_(reference) + '</span>' +
        '</div>' +
        '<div class="mt-1.5">' +
          '<span class="block truncate text-[13px] font-medium ' + stockClass + '">' + StockWebApp_escapeHtml_(stockDisplay) + '</span>' +
        '</div>' +
      '</button>' +
      '<button class="reference-detail-trigger flex w-10 shrink-0 touch-manipulation select-none items-center justify-center border-l border-outline-variant/20 text-outline-variant transition-colors duration-150 hover:bg-surface-container-highest hover:text-on-surface-variant active:bg-surface-container-high" type="button" aria-label="Ouvrir la fiche de ' + StockWebApp_escapeHtml_(reference) + '" data-action="open-detail" data-item-id="' + StockWebApp_escapeHtml_(itemId) + '" data-reference="' + StockWebApp_escapeHtml_(reference) + '">' +
        '<span class="material-symbols-outlined !text-[16px]">chevron_right</span>' +
      '</button>' +
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
