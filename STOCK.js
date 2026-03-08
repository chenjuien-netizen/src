function StockMoves_validateAll_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_STOCK);
  if (!sh) throw new Error("Feuille STOCK introuvable.");

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return;

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = (typeof headerMap_ === "function") ? headerMap_(headers) : StockMoves_headerMapLocal_(headers);

  const colSel = StockMoves_col_(map, "选择");
  const colRef = StockMoves_col_(map, "货号");
  const colIn = StockMoves_col_(map, "进货");
  const colOut = StockMoves_col_(map, "出-Sortie/箱");

  const colTail = StockMoves_col_(map, "当前尾箱件数");
  const colPpc = StockMoves_col_(map, "每箱件数2");
  const colBoxes = StockMoves_col_(map, "当前箱数");
  const colSign = StockMoves_col_(map, "当前signe");
  const colFrac = StockMoves_col_(map, "当前箱数分数");
  const colMissing = StockMoves_col_(map, "当前缺包");
  const colOpenRest = StockMoves_col_(map, "Carton ouvert (reste)");

  if (!colSel || !colRef || !colIn || !colOut || !colTail || !colPpc || !colBoxes || !colSign || !colFrac || !colMissing) {
    throw new Error("Colonnes requises manquantes dans STOCK (选择, 货号, 进货, 出-Sortie/箱, 当前尾箱件数, 每箱件数2, 当前箱数, 当前signe, 当前箱数分数, 当前缺包).");
  }

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const refToRow = {};
  for (let i = 0; i < data.length; i++) {
    const ref = String(data[i][colRef - 1] || "").trim().toUpperCase();
    if (ref) refToRow[ref] = i + 2;
  }

  const writes = []; // {r,c,v}
  const push = (r, c, v) => { if (r && c) writes.push({ r, c, v }); };
  const getPendingValue = (r, c, fallback) => {
    for (let k = writes.length - 1; k >= 0; k--) {
      const w = writes[k];
      if (w.r === r && w.c === c) return w.v;
    }
    return fallback;
  };

  const nowStr = StockMoves_nowStr_();
  const touchedRows = [];
  const fracRows = [];
  const errors = [];
  let applied = 0;

  for (let i = 0; i < data.length; i++) {
    const row = i + 2;
    const rowValues = data[i];

    if (!StockMoves_isTruthy_(rowValues[colSel - 1])) continue;

    const ref = String(rowValues[colRef - 1] || "").trim().toUpperCase();
    if (!ref) {
      errors.push(`Ligne ${row}: 货号 vide`);
      continue;
    }

    const outRaw = String(rowValues[colOut - 1] || "").trim();
    const oldHistory = String(rowValues[colIn - 1] || "").trim();

    const currentState = StockMoves_stateFromRowValues_(rowValues, map);
    const openRest = colOpenRest ? StockMoves_toInt_(rowValues[colOpenRest - 1]) : 0;
    const currentStateText = StockMoves_buildNormalizedStateText_(currentState);
    const latestHistoryStateText = StockMoves_extractLatestStateFromHistory_(oldHistory);
    const manualChanged = StockMoves_hasManualStateChange_(currentStateText, latestHistoryStateText);

    // Do not block stock exits just because the current row shape differs from 进货.
    // Without a real snapshot, that comparison creates false conflicts on untouched lines.
    // If 出-Sortie/箱 is filled, we let the exit flow continue and use the current row state as source.

    if (!outRaw) {
      if (manualChanged) {
        push(row, colIn, StockMoves_prependHistoryState_(oldHistory, nowStr, currentStateText));
      }
      push(row, colSel, false);
      touchedRows.push(row);
      applied++;
      continue;
    }

    const allowedOut = StockMoves_getAllowedOutValues_(currentState, openRest);
    const canonicalOut = StockMoves_canonicalizeOutValue_(outRaw);
    if (!canonicalOut || allowedOut.indexOf(canonicalOut) === -1) {
      errors.push(`Ligne ${row} (${ref}): sortie non autorisée pour l'état courant "${outRaw}"`);
      continue;
    }

    const parsed = StockMoves_parseOutValue_(outRaw);
    if (!parsed.type) {
      errors.push(`Ligne ${row} (${ref}): sortie invalide "${outRaw}"`);
      continue;
    }

    const packsPerBox = StockMoves_getPacksPerBox_(sh, map, row, rowValues);
    let nextState;
    try {
      nextState = StockMoves_applyOutCommandToState_(currentState, parsed, row, ref, packsPerBox);
    } catch (err) {
      errors.push(String(err && err.message ? err.message : err));
      continue;
    }

    if (parsed.type === "CLEAR" || parsed.type === "TAIL") {
      StockMoves_closeMixIfAny_(sh, map, refToRow, ref, row, rowValues, nowStr, writes, getPendingValue);
    }

    push(row, colTail, nextState.tail);
    push(row, colPpc, nextState.ppc);
    push(row, colBoxes, nextState.boxes);
    push(row, colSign, nextState.sign);
    push(row, colFrac, nextState.fraction);
    push(row, colMissing, nextState.missingPacks);
    fracRows.push(row);

    const nextStateText = StockMoves_buildNormalizedStateText_(nextState);
    push(row, colIn, StockMoves_prependHistoryState_(oldHistory, nowStr, nextStateText));

    push(row, colOut, "");
    push(row, colSel, false);

    touchedRows.push(row);
    applied++;
  }

  for (const w of writes) sh.getRange(w.r, w.c).setValue(w.v);

  for (const r of Array.from(new Set(fracRows))) {
    sh.getRange(r, colFrac).setValue(StockMoves_toNumber_(sh.getRange(r, colFrac).getValue()));
    StockMoves_applyFractionDisplayFormat_(sh.getRange(r, colFrac));
  }

  for (const r of Array.from(new Set(touchedRows))) {
    StockMoves_setOutDropdown_(sh, map, r);
  }

  if (errors.length) {
    const head = `✅ Lignes traitées: ${applied}`;
    throw new Error(head + "\n" + errors.join("\n"));
  }

  try { ss.toast(`✅ Lignes traitées: ${applied}`, "STOCK", 5); } catch (e) {}
}

function stockResetPending_() {
  StockMoves_resetPendingRows_();
}

function StockMoves_resetPendingRows_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_STOCK);
  if (!sh) throw new Error("Feuille STOCK introuvable.");

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return;

  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = (typeof headerMap_ === "function") ? headerMap_(headers) : StockMoves_headerMapLocal_(headers);

  const colSel = StockMoves_col_(map, "选择");
  const colIn = StockMoves_col_(map, "进货");
  const colOut = StockMoves_col_(map, "出-Sortie/箱");
  const colTail = StockMoves_col_(map, "当前尾箱件数");
  const colPpc = StockMoves_col_(map, "每箱件数2");
  const colBoxes = StockMoves_col_(map, "当前箱数");
  const colSign = StockMoves_col_(map, "当前signe");
  const colFrac = StockMoves_col_(map, "当前箱数分数");
  const colMissing = StockMoves_col_(map, "当前缺包");

  if (!colSel || !colIn || !colOut || !colTail || !colPpc || !colBoxes || !colSign || !colFrac || !colMissing) {
    throw new Error("Colonnes requises manquantes pour reset pending.");
  }

  const data = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const writes = [];
  const push = (r, c, v) => { if (r && c) writes.push({ r, c, v }); };

  const touchedRows = [];
  const errors = [];
  let resetCount = 0;

  for (let i = 0; i < data.length; i++) {
    const row = i + 2;
    const rowValues = data[i];
    const isSelected = StockMoves_isTruthy_(rowValues[colSel - 1]);
    const outRaw = String(rowValues[colOut - 1] || "").trim();
    if (!isSelected && !outRaw) continue;

    const latestStateText = StockMoves_extractLatestStateFromHistory_(String(rowValues[colIn - 1] || "").trim());
    const packsPerBox = StockMoves_getPacksPerBox_(sh, map, row, rowValues);

    let restored;
    try {
      // Safe fallback when history is empty: restore zero/default state.
      restored = latestStateText ? StockMoves_parseNormalizedStateText_(latestStateText) : { tail: 0, ppc: 0, boxes: 0, sign: "", fraction: 0, missingPacks: 0 };

      // If history is pure pack form like "9包", keep the current 每箱件数2 value
      if (/^\s*\d+\s*包\s*$/.test(latestStateText)) {
        const currentPpc = StockMoves_toInt_(rowValues[colPpc - 1]);
        if (currentPpc > 0) restored.ppc = currentPpc;
      }

      restored.packsPerBox = packsPerBox;
      restored = StockMoves_normalizeState_(restored);
    } catch (err) {
      errors.push(`Ligne ${row}: impossible de parser l'état historisé "${latestStateText}"`);
      continue;
    }

    push(row, colTail, restored.tail);
    push(row, colPpc, restored.ppc);
    push(row, colBoxes, restored.boxes);
    push(row, colSign, restored.sign);
    push(row, colFrac, restored.fraction);
    push(row, colMissing, restored.missingPacks);
    push(row, colOut, "");
    push(row, colSel, false);

    touchedRows.push(row);
    resetCount++;
  }

  for (const w of writes) sh.getRange(w.r, w.c).setValue(w.v);

  for (const r of Array.from(new Set(touchedRows))) {
    StockMoves_applyFractionDisplayFormat_(sh.getRange(r, colFrac));
    StockMoves_setOutDropdown_(sh, map, r);
  }

  if (errors.length) {
    throw new Error(`✅ Lignes reset: ${resetCount}\n` + errors.join("\n"));
  }

  try { ss.toast(`↩️ Lignes reset: ${resetCount}`, "STOCK", 5); } catch (e) {}
}

// ---------- MIX CLOSE (sur sortie tail ou 清空库存) ----------
function StockMoves_closeMixIfAny_(sh, map, refToRow, ref, row, rowValues, nowStr, writes, getPendingValue) {
  const colMix = StockMoves_col_(map, "混箱数");
  const colIsMix = StockMoves_col_(map, "is_mix");
  const colLoc = StockMoves_col_(map, "放位/提醒");

  const push = (r, c, v) => { if (r && c) writes.push({ r, c, v }); };

  const isMixVal = colIsMix ? StockMoves_toInt_(rowValues[colIsMix - 1]) : 0;
  const mixVal = colMix ? StockMoves_toInt_(rowValues[colMix - 1]) : 0;
  const noteS = colLoc ? String(rowValues[colLoc - 1] || "").trim() : "";

  const partner = (typeof parseMixPartnerFromNoteSystem_ === "function") ? parseMixPartnerFromNoteSystem_(noteS) : "";

  const hasMix = (isMixVal > 0) || (mixVal > 0) || (!!partner);
  if (!hasMix) return;

  if (colMix) push(row, colMix, 0);
  if (colIsMix) push(row, colIsMix, 0);

  const pRef = String(partner || "").trim().toUpperCase();
  const pRow = pRef ? refToRow[pRef] : 0;
  if (pRow) {
    if (colMix) push(pRow, colMix, 0);
    if (colIsMix) push(pRow, colIsMix, 0);
  }
}

// ---------- onEdit routing (STOCK) ----------
function StockMoves_onEditWarehouse_(e) {
  try {
    if (typeof StockMoves_onEdit_ === "function") StockMoves_onEdit_(e);

    const range = e && e.range;
    if (!range) return;

    const sh = range.getSheet();
    if (!sh || sh.getName() !== SHEET_STOCK) return;

    const row = range.getRow();
    if (row < 2) return;

    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const map = (typeof headerMap_ === "function") ? headerMap_(headers) : StockMoves_headerMapLocal_(headers);

    const colWh = StockMoves_col_(map, "仓库");
    const colLog = StockMoves_col_(map, "出库记录");
    if (!colWh || !colLog) return;

    if (range.getColumn() !== colWh) return;

    const oldV = (typeof e.oldValue === "undefined") ? "" : String(e.oldValue || "").trim();
    const newV = String(range.getValue() || "").trim();
    if (!newV || newV === oldV) return;

    const nowStr = StockMoves_nowStr_();
    const oldLog = String(sh.getRange(row, colLog).getValue() || "").trim();
    const line = `${nowStr} | 仓库: ${oldV || "∅"} → ${newV}`;
    sh.getRange(row, colLog).setValue(StockMoves_concatLog_(oldLog, line));
  } catch (err) {
    // silencieux pour éviter de casser onEdit
  }
}

function StockMoves_onEdit_(e) {
  const range = e && e.range;
  if (!range) return;

  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;

  const sh = range.getSheet();
  if (!sh || sh.getName() !== SHEET_STOCK) return;

  const row = range.getRow();
  if (row < 2) return;

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = (typeof headerMap_ === "function") ? headerMap_(headers) : StockMoves_headerMapLocal_(headers);

  const colSel = StockMoves_col_(map, "选择");
  const colIn = StockMoves_col_(map, "进货");
  const colOut = StockMoves_col_(map, "出-Sortie/箱");
  const colFrac = StockMoves_col_(map, "当前箱数分数");
  if (!colSel || !colIn || !colOut || !colFrac) return;

  const col = range.getColumn();
  const manualCols = StockMoves_manualStateColumns_(map);
  const isManualCol = manualCols.indexOf(col) !== -1;

  if (col === colFrac) {
    const rawEditValue = (e && Object.prototype.hasOwnProperty.call(e, "value"))
      ? e.value
      : range.getDisplayValue();

    try {
      const parsedFraction = StockMoves_parseFractionInput_(rawEditValue);
      if (parsedFraction !== null) {
        range.setNumberFormat("@");
        range.setValue(parsedFraction);
        StockMoves_applyFractionDisplayFormat_(range);
      } else {
        StockMoves_applyFractionDisplayFormat_(range);
      }
    } catch (err) {
      range.clearContent();
      StockMoves_applyFractionDisplayFormat_(range);
      try { SpreadsheetApp.getActive().toast(String(err && err.message ? err.message : err), "STOCK", 4); } catch (e2) {}
      return;
    }

    const fracNow = StockMoves_toNumber_(range.getValue());
    const colSign = StockMoves_col_(map, "当前signe");
    if (colSign) {
      if (!(fracNow > 0)) {
        sh.getRange(row, colSign).setValue("");
      } else {
        const signCell = sh.getRange(row, colSign);
        const signNow = String(signCell.getValue() || "").trim();
        if (!signNow) signCell.setValue("+");
      }
    }

    sh.getRange(row, colSel).setValue(true);
    StockMoves_setOutDropdown_(sh, map, row);
    return;
  }

  if (col === colOut) {
    const outRaw = String(range.getValue() || "").trim();
    if (outRaw) {
      sh.getRange(row, colSel).setValue(true);
    } else {
      const latestState = StockMoves_extractLatestStateFromHistory_(String(sh.getRange(row, colIn).getValue() || ""));
      const rowObj = StockMoves_stateFromSheetRow_(sh, map, row);
      const currentState = StockMoves_buildNormalizedStateText_(rowObj);
      if (!StockMoves_hasManualStateChange_(currentState, latestState)) {
        sh.getRange(row, colSel).setValue(false);
      }
    }
    StockMoves_setOutDropdown_(sh, map, row);
    return;
  }

  if (isManualCol) {
    sh.getRange(row, colSel).setValue(true);
    StockMoves_setOutDropdown_(sh, map, row);
  }
}

// ---------- helpers ----------
function StockMoves_parseOutValue_(s) {
  const raw = String(s || "").trim();
  if (!raw) return { type: "" };

  if (/清空库存/i.test(raw) || /^vider$/i.test(raw) || /^all$/i.test(raw)) {
    return { type: "CLEAR" };
  }

  const mTail = raw.match(/^\(?\s*(\d+)\s*[pP]\s*\)?$/);
  if (mTail) return { type: "TAIL", qty: Number(mTail[1]) };

  const mBox = raw.match(/^\s*(\d+)\s*箱\s*$/);
  if (mBox) return { type: "BOXES", qty: Number(mBox[1]) };

  const mPack = raw.match(/^\s*(\d+)\s*包\s*$/);
  if (mPack) return { type: "PACKS", qty: Number(mPack[1]) };

  const mFrac = raw.match(/^\s*(\d+)\s*\/\s*(\d+)\s*$/);
  if (mFrac) {
    const num = Number(mFrac[1]);
    const den = Number(mFrac[2]);
    if (!den) return { type: "" };
    return { type: "FRACTION", qty: num / den, num, den };
  }

  const mNum = raw.match(/^\s*(\d+)\s*$/);
  if (mNum) return { type: "BOXES", qty: Number(mNum[1]) };

  return { type: "" };
}

function StockMoves_fractionStringToNumber_(txt) {
  const s = String(txt || "").trim();
  const m = s.match(/^(\d+)\s*\/\s*(\d+)$/);
  if (!m) throw new Error("Fraction invalide: " + s);
  const num = Number(m[1]);
  const den = Number(m[2]);
  if (!den) throw new Error("Fraction invalide: " + s);
  return num / den;
}

function StockMoves_parseNormalizedStateText_(stateText) {
  let raw = String(stateText || "").trim();
  const out = { tail: 0, ppc: 0, boxes: 0, sign: "", fraction: 0, missingPacks: 0 };
  if (!raw) return out;

  let mPurePack = raw.match(/^(\d+)\s*包$/);
  if (mPurePack) {
    out.missingPacks = Number(mPurePack[1]) || 0;
    return out;
  }

  let m = raw.match(/^\((\d+)\s*[pP]\)\s*\+?\s*(.*)$/);
  if (m) {
    out.tail = Number(m[1]) || 0;
    raw = String(m[2] || "").trim();
  }

  m = raw.match(/([+-]\d+)\s*包$/);
  if (m) {
    out.missingPacks = Number(m[1]) || 0;
    raw = raw.slice(0, m.index).trim();
  }

  if (!raw) return out;

  m = raw.match(/^(\d+)\s*[pP]$/);
  if (m) {
    out.ppc = Number(m[1]);
    out.boxes = 1;
    return out;
  }

  m = raw.match(/^(\d+)\s*[pP]\s*×\s*(\d+)$/);
  if (m) {
    out.ppc = Number(m[1]);
    out.boxes = Number(m[2]);
    return out;
  }

  m = raw.match(/^(\d+)\s*[pP]\s*\+\s*(\d+\s*\/\s*\d+)$/);
  if (m) {
    out.ppc = Number(m[1]);
    out.boxes = 1;
    out.sign = "+";
    out.fraction = StockMoves_fractionStringToNumber_(m[2]);
    return out;
  }

  m = raw.match(/^(\d+)\s*[pP]\s*×\s*(\d+)\s*\+\s*(\d+\s*\/\s*\d+)$/);
  if (m) {
    out.ppc = Number(m[1]);
    out.boxes = Number(m[2]);
    out.sign = "+";
    out.fraction = StockMoves_fractionStringToNumber_(m[3]);
    return out;
  }

  m = raw.match(/^(\d+)\s*[pP]\s*×\s*(\d+\s*\/\s*\d+)$/);
  if (m) {
    out.ppc = Number(m[1]);
    out.boxes = 1;
    out.sign = "×";
    out.fraction = StockMoves_fractionStringToNumber_(m[2]);
    return out;
  }

  m = raw.match(/^(\d+)\s*[pP]\s*×\s*(\d+)\s*×\s*(\d+\s*\/\s*\d+)$/);
  if (m) {
    out.ppc = Number(m[1]);
    out.boxes = Number(m[2]);
    out.sign = "×";
    out.fraction = StockMoves_fractionStringToNumber_(m[3]);
    return out;
  }

  throw new Error("Etat normalisé non reconnu: " + stateText);
}

function StockMoves_applyOutCommandToState_(stateInput, parsed, row, ref, packsPerBox) {
  const state = StockMoves_normalizeState_(Object.assign({}, stateInput, { packsPerBox: StockMoves_toInt_(packsPerBox) }));
  const next = StockMoves_copyState_(state);

  if (parsed.type === "CLEAR") {
    next.tail = 0;
    next.boxes = 0;
    next.sign = "";
    next.fraction = 0;
    next.missingPacks = 0;
    return StockMoves_normalizeState_(next);
  }

  if (parsed.type === "TAIL") {
    const qtyTail = StockMoves_toInt_(parsed.qty);
    if (qtyTail <= 0) throw new Error(`尾箱数量 invalide ligne ${row} (${ref}).`);
    if (qtyTail > next.tail) throw new Error(`尾箱不足 ligne ${row} (${ref}) : demandé ${qtyTail}p > 当前尾箱件数 ${next.tail}p`);
    next.tail = next.tail - qtyTail;
    return StockMoves_normalizeState_(next);
  }

  let totalBoxes = StockMoves_stateToBoxesEquivalent_(next);
  let consume = 0;

  if (parsed.type === "BOXES") {
    consume = Number(parsed.qty || 0);
  } else if (parsed.type === "PACKS") {
    const ppb = StockMoves_toInt_(state.packsPerBox);
    if (ppb <= 0) throw new Error(`包/箱 invalide ligne ${row} (${ref}) pour sortie 包.`);
    consume = Number(parsed.qty || 0) / ppb;
  } else if (parsed.type === "FRACTION") {
    consume = Number(parsed.qty || 0);
  }

  if (!(consume > 0)) throw new Error(`Sortie invalide ligne ${row} (${ref}).`);
  if (parsed.type === "BOXES" && Number(parsed.qty || 0) === 1 && StockMoves_hasOpenState_(state)) {
    // Business rule: on any open state, 1箱 means clear the currently open remainder only.
    // It must not consume another closed carton.
    if (state.sign === "+" && state.boxes > 0) {
      next.boxes = Math.max(0, state.boxes - 1);
    } else {
      next.boxes = 0;
    }
    next.sign = "";
    next.fraction = 0;
    next.missingPacks = 0;
    return StockMoves_normalizeState_(next);
  }
  if (consume > totalBoxes + 1e-9) {
    throw new Error(`Stock insuffisant ligne ${row} (${ref}) : demandé ${consume}箱-equivalent > disponible ${totalBoxes}.`);
  }

  totalBoxes = Math.max(0, totalBoxes - consume);
  let shape;
  if (parsed.type === "PACKS" && StockMoves_toInt_(state.packsPerBox) > 0) {
    // After a pack move, prefer canonical packs form over fraction form.
    const ppb = StockMoves_toInt_(state.packsPerBox);
    if (totalBoxes > 0 && totalBoxes < 1) {
      next.boxes = 0;
      next.sign = "";
      next.fraction = 0;
      next.missingPacks = Math.max(0, Math.floor(totalBoxes * ppb + 1e-9));
      next.packsPerBox = ppb;
      return StockMoves_normalizeState_(next);
    }
    const totalPacks = Math.max(0, Math.floor(totalBoxes * ppb + 1e-9));
    const whole = Math.floor(totalPacks / ppb);
    const remPacks = totalPacks - (whole * ppb);
    shape = { boxes: whole, sign: "", fraction: 0, missingPacks: remPacks, packsPerBox: ppb };
  } else {
    shape = StockMoves_boxesEquivalentToState_(totalBoxes, state.packsPerBox);
  }

  next.boxes = shape.boxes;
  next.sign = shape.sign;
  next.fraction = shape.fraction;
  next.missingPacks = shape.missingPacks;

  return StockMoves_normalizeState_(next);
}

function StockMoves_stateToBoxesEquivalent_(stateInput) {
  const s = StockMoves_normalizeState_(stateInput);
  const frac = Number(s.fraction || 0);
  let total = 0;
  if (s.sign === "+") total = s.boxes + frac;
  else if (s.sign === "×") {
    const mult = s.boxes > 1 ? s.boxes : 1;
    total = mult * frac;
  } else {
    total = s.boxes;
  }

  const miss = Number(s.missingPacks || 0);
  if (miss !== 0) {
    const ppb = StockMoves_toInt_(s.packsPerBox);
    if (ppb <= 0) throw new Error("包/箱 invalide pour intégrer 当前缺包.");
    total += miss / ppb;
  }

  return Math.max(0, total);
}

function StockMoves_boxesEquivalentToState_(totalBoxesInput, packsPerBox) {
  const n = Number(totalBoxesInput || 0);
  const ppb = StockMoves_toInt_(packsPerBox);
  if (!(n > 1e-9)) return { boxes: 0, sign: "", fraction: 0, missingPacks: 0, packsPerBox: ppb };

  // Prefer canonical "boxes + missing packs" when 包/箱 is valid.
  if (ppb > 0) {
    const eps = 1e-6;

    // Special case: for sub-1 carton quantities, keep a pure packs form
    // like 0箱 + 9包 instead of 1箱 - 7包.
    if (n < 1 - eps) {
      const remPacks = Math.round(n * ppb);
      const reconSmall = remPacks / ppb;
      if (remPacks >= 0 && remPacks < ppb && Math.abs(reconSmall - n) <= eps) {
        return { boxes: 0, sign: "", fraction: 0, missingPacks: remPacks, packsPerBox: ppb };
      }
    }

    let baseBoxes = Math.ceil(n - eps);
    let miss = Math.round((n - baseBoxes) * ppb); // usually <= 0

    if (miss <= -ppb || miss >= ppb) {
      const delta = Math.trunc(miss / ppb);
      baseBoxes += delta;
      miss -= delta * ppb;
    }

    if (baseBoxes < 0) {
      baseBoxes = 0;
      miss = 0;
    }

    const recon = baseBoxes + (miss / ppb);
    if (Math.abs(recon - n) <= eps) {
      return { boxes: baseBoxes, sign: "", fraction: 0, missingPacks: miss, packsPerBox: ppb };
    }
  }

  const whole = Math.floor(n + 1e-9);
  const rem = n - whole;

  if (Math.abs(rem) <= 1e-9) {
    return { boxes: whole, sign: "", fraction: 0, missingPacks: 0, packsPerBox: ppb };
  }

  if (whole >= 1) {
    return { boxes: whole, sign: "+", fraction: rem, missingPacks: 0, packsPerBox: ppb };
  }

  return { boxes: 1, sign: "×", fraction: n, missingPacks: 0, packsPerBox: ppb };
}

function StockMoves_parseFractionInput_(value) {
  if (typeof value === "number") return null;
  const s = String(value || "").trim();
  if (!s) return null;
  const m = s.match(/^(\d+)\s*\/\s*(\d+)$/);
  if (!m) return null;

  const num = Number(m[1]);
  const den = Number(m[2]);
  if (!den) throw new Error("分数格式错误: " + s);
  if ([2, 3, 4].indexOf(den) === -1) throw new Error("分数只支持 /2 /3 /4");
  return num / den;
}

function StockMoves_applyFractionDisplayFormat_(range) {
  if (!range) return;
  range.setNumberFormat("# ?/?");
}

function StockMoves_gcd_(a, b) {
  a = Math.abs(Math.trunc(a || 0));
  b = Math.abs(Math.trunc(b || 0));
  while (b) {
    const t = b;
    b = a % b;
    a = t;
  }
  return a || 1;
}

function StockMoves_reduceFraction_(num, den) {
  if (!den) return { num: 0, den: 1 };
  const g = StockMoves_gcd_(num, den);
  return { num: num / g, den: den / g };
}

function StockMoves_fractionToText_(v) {
  const n = Number(v);
  if (!Number.isFinite(n) || n <= 0) return "";

  const candidates = [2, 3, 4, 6, 8, 12];
  let best = null;
  let bestErr = Infinity;

  for (const den of candidates) {
    const num = Math.round(n * den);
    const approx = num / den;
    const err = Math.abs(n - approx);
    if (err < bestErr) {
      bestErr = err;
      best = { num, den };
    }
  }

  if (!best) return "";
  const reduced = StockMoves_reduceFraction_(best.num, best.den);
  return reduced.num + "/" + reduced.den;
}

function StockMoves_normalizeState_(stateInput) {
  const s = StockMoves_copyState_(stateInput || {});

  s.tail = Math.max(0, Math.trunc(Number(s.tail) || 0));
  s.ppc = Math.max(0, Math.trunc(Number(s.ppc) || 0));
  s.boxes = Math.max(0, Math.trunc(Number(s.boxes) || 0));
  s.missingPacks = Math.trunc(Number(s.missingPacks) || 0);
  s.packsPerBox = Math.max(0, Math.trunc(Number(s.packsPerBox) || 0));

  const signRaw = String(s.sign || "").trim();
  s.sign = (signRaw === "+" || signRaw === "×") ? signRaw : "";

  let fracNum = Number(s.fraction || 0);
  if (!Number.isFinite(fracNum) || fracNum <= 0) fracNum = 0;

  if (fracNum > 0) {
    const fracTxt = StockMoves_fractionToText_(fracNum);
    if (fracTxt) {
      const parts = fracTxt.split("/");
      fracNum = Number(parts[0]) / Number(parts[1]);
    }
  }

  if (fracNum <= 0) {
    s.fraction = 0;
    s.sign = "";
  } else {
    s.fraction = fracNum;
    if (!s.sign) s.sign = s.boxes > 0 ? "+" : "×";
  }

  if (s.sign === "×" && s.fraction > 0 && s.boxes <= 0) s.boxes = 1;
  if (s.sign === "+" && s.fraction > 0 && s.boxes <= 0) {
    s.sign = "×";
    s.boxes = 1;
  }

  // Prefer canonical boxes+packs when 包/箱 is available.
  if (s.packsPerBox > 0) {
    let baseBoxesEq = 0;
    if (s.sign === "+") {
      baseBoxesEq = s.boxes + s.fraction;
    } else if (s.sign === "×") {
      const mult = s.boxes > 1 ? s.boxes : 1;
      baseBoxesEq = mult * s.fraction;
    } else {
      baseBoxesEq = s.boxes;
    }
    const totalBoxesEq = Math.max(0, baseBoxesEq + (s.missingPacks / s.packsPerBox));
    const canon = StockMoves_boxesEquivalentToState_(totalBoxesEq, s.packsPerBox);
    s.boxes = canon.boxes;
    s.missingPacks = canon.missingPacks;
    s.sign = canon.sign;
    s.fraction = canon.fraction;
  }

  if (s.packsPerBox > 0 && s.missingPacks !== 0) {
    let delta = 0;
    if (s.missingPacks >= s.packsPerBox) delta = Math.floor(s.missingPacks / s.packsPerBox);
    if (s.missingPacks <= -s.packsPerBox) delta = Math.ceil(s.missingPacks / s.packsPerBox);
    if (delta !== 0) {
      s.boxes = Math.max(0, s.boxes + delta);
      s.missingPacks = s.missingPacks - (delta * s.packsPerBox);
    }
  }

  if (s.missingPacks !== 0 && s.packsPerBox > 0) {
    // Keep canonical packs form; avoid mixed fraction + missing packs.
    s.sign = "";
    s.fraction = 0;
  }

  if (!(s.fraction > 0)) {
    s.sign = "";
  }

  return s;
}

function StockMoves_buildNormalizedStateText_(rowObj) {
  const s = StockMoves_normalizeState_(rowObj);

  const ppcTxt = s.ppc > 0 ? String(s.ppc) + "p" : "";
  const wholeN = Math.max(0, Math.trunc(s.boxes));
  const signTxt = String(s.sign || "").trim();
  const fracTxt = StockMoves_fractionToText_(s.fraction);
  const missN = Number(s.missingPacks) || 0;
  const tailN = Number(s.tail) || 0;

  if (tailN <= 0 && wholeN <= 0 && !signTxt && !(s.fraction > 0) && missN > 0) {
    return String(missN) + "包";
  }

  let core = "";
  if (ppcTxt) {
    if (signTxt === "×") {
      core = ppcTxt + "×" + (wholeN <= 1 ? "" : String(wholeN) + "×") + fracTxt;
    } else if (signTxt === "+") {
      core = ppcTxt + (wholeN <= 1 ? "" : "×" + String(wholeN)) + "+" + fracTxt;
    } else {
      const showWhole = (wholeN > 1) || (wholeN === 1 && missN !== 0);
      core = ppcTxt + (showWhole ? "×" + String(wholeN) : "");
    }

    if (missN !== 0) core += (missN > 0 ? "+" : "") + String(missN) + "包";
  }

  if (tailN > 0) {
    core = "(" + String(Math.trunc(tailN)) + "p)" + (core ? "+" + core : "");
  }

  return core;
}

function StockMoves_extractLatestStateFromHistory_(historyText) {
  const raw = String(historyText || "").trim();
  if (!raw) return "";
  const first = raw.split(/\r?\n/).find(Boolean) || "";
  const idx = first.indexOf("|");
  if (idx < 0) return "";
  return String(first.slice(idx + 1) || "").trim();
}

function StockMoves_prependHistoryState_(oldHistory, nowStr, stateText) {
  const base = String(oldHistory || "").trim();
  const st = String(stateText || "").trim();
  if (!st) return base;

  const latest = StockMoves_extractLatestStateFromHistory_(base);
  if (latest === st) return base;

  const line = `${nowStr} | ${st}`;
  if (!base) return line;
  return line + "\n" + base;
}

function StockMoves_hasManualStateChange_(currentStateText, latestHistoryStateText) {
  const a = String(currentStateText || "").trim();
  const b = String(latestHistoryStateText || "").trim();
  return a !== b;
}

function StockMoves_concatLog_(oldText, line) {
  const a = String(oldText || "").trim();
  const b = String(line || "").trim();
  if (!b) return a;
  if (!a) return b;
  return b + "\n" + a;
}

function StockMoves_toInt_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return Math.trunc(v);
  const s = String(v).trim();
  const m = s.match(/-?\d+/);
  return m ? Number(m[0]) : 0;
}

function StockMoves_toNumber_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const s = String(v).trim().replace(",", ".");
  if (!s) return 0;
  const m = s.match(/-?\d+(\.\d+)?/);
  if (!m) return 0;
  const n = Number(m[0]);
  return Number.isFinite(n) ? n : 0;
}

function StockMoves_nowStr_() {
  const tz = Session.getScriptTimeZone() || "Europe/Paris";
  return Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm");
}

function StockMoves_headerMapLocal_(headersRow) {
  const map = {};
  for (let c = 0; c < headersRow.length; c++) {
    const h = headersRow[c];
    if (h === null || typeof h === "undefined") continue;
    const key = String(h).trim();
    if (!key) continue;
    map[key.toLowerCase()] = c + 1;
  }
  return map;
}

function StockMoves_col_(map, header) {
  const k = String(header || "").toLowerCase();
  return map[k] || map[header] || 0;
}

function StockMoves_isTruthy_(v) {
  if (v === true) return true;
  if (typeof v === "number") return v !== 0;
  const s = String(v || "").trim().toLowerCase();
  return s === "true" || s === "1" || s === "yes" || s === "y" || s === "oui";
}

function StockMoves_manualStateColumns_(map) {
  return [
    StockMoves_col_(map, "当前尾箱件数"),
    StockMoves_col_(map, "每箱件数2"),
    StockMoves_col_(map, "当前箱数"),
    StockMoves_col_(map, "当前signe"),
    StockMoves_col_(map, "当前箱数分数"),
    StockMoves_col_(map, "当前缺包")
  ].filter(Boolean);
}

function StockMoves_stateFromRowValues_(rowValues, map) {
  const colTail = StockMoves_col_(map, "当前尾箱件数");
  const colPpc = StockMoves_col_(map, "每箱件数2");
  const colBoxes = StockMoves_col_(map, "当前箱数");
  const colSign = StockMoves_col_(map, "当前signe");
  const colFrac = StockMoves_col_(map, "当前箱数分数");
  const colMissing = StockMoves_col_(map, "当前缺包");
  const colPackPerBox = StockMoves_col_(map, "包/箱");

  return StockMoves_normalizeState_({
    tail: colTail ? StockMoves_toInt_(rowValues[colTail - 1]) : 0,
    ppc: colPpc ? StockMoves_toInt_(rowValues[colPpc - 1]) : 0,
    boxes: colBoxes ? StockMoves_toInt_(rowValues[colBoxes - 1]) : 0,
    sign: colSign ? String(rowValues[colSign - 1] || "").trim() : "",
    fraction: colFrac ? StockMoves_toNumber_(rowValues[colFrac - 1]) : 0,
    missingPacks: colMissing ? StockMoves_toInt_(rowValues[colMissing - 1]) : 0,
    packsPerBox: colPackPerBox ? StockMoves_toInt_(rowValues[colPackPerBox - 1]) : 0
  });
}

function StockMoves_stateFromSheetRow_(sh, map, row) {
  const lastCol = sh.getLastColumn();
  const values = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  return StockMoves_stateFromRowValues_(values, map);
}

function StockMoves_copyState_(s) {
  return {
    tail: Number(s && s.tail ? s.tail : 0),
    ppc: Number(s && s.ppc ? s.ppc : 0),
    boxes: Number(s && s.boxes ? s.boxes : 0),
    sign: String(s && s.sign ? s.sign : ""),
    fraction: Number(s && s.fraction ? s.fraction : 0),
    missingPacks: Number(s && s.missingPacks ? s.missingPacks : 0),
    packsPerBox: Number(s && s.packsPerBox ? s.packsPerBox : 0)
  };
}

function StockMoves_getPacksPerBox_(sh, map, row, rowValues) {
  const colPackPerBox = StockMoves_col_(map, "包/箱");
  if (!colPackPerBox) return 0;

  const raw = rowValues ? rowValues[colPackPerBox - 1] : sh.getRange(row, colPackPerBox).getValue();
  const n = StockMoves_toNumber_(raw);
  if (!Number.isFinite(n) || n <= 0) return 0;
  return Math.max(0, Math.trunc(n));
}

function StockMoves_hasOpenState_(state) {
  const miss = Number(state && state.missingPacks ? state.missingPacks : 0);
  const frac = Number(state && state.fraction ? state.fraction : 0);
  return miss !== 0 || frac > 0;
}

function StockMoves_getAllowedOutValues_(stateInput, openRestInput) {
  const state = StockMoves_normalizeState_(stateInput || {});
  const curTail = StockMoves_toInt_(state.tail);
  const curBoxes = StockMoves_toInt_(state.boxes);
  const curSign = String(state.sign || "").trim();
  const curFracText = StockMoves_fractionToText_(state.fraction);
  const curMissing = StockMoves_toInt_(state.missingPacks);
  const hasFraction = !!curFracText;
  const isOpen = StockMoves_hasOpenState_(state);
  const purePacks = curMissing > 0 && curBoxes <= 0 && !curSign && !hasFraction;
  const openRest = Math.max(0, StockMoves_toInt_(openRestInput));

  const list = [];
  const seen = {};
  const addOpt = (v) => {
    const s = String(v || "").trim();
    if (!s || seen[s]) return;
    seen[s] = true;
    list.push(s);
  };

  if (curTail > 0) addOpt("(" + Math.trunc(curTail) + "p)");

  if (isOpen) {
    if (curSign === "+") {
      if (curFracText) addOpt(curFracText);
    } else if (curSign === "×") {
      const m = String(curFracText || "").match(/^(\d+)\/(\d+)$/);
      if (m) {
        const num = Number(m[1]);
        const den = Number(m[2]);
        if (num > 1) {
          for (let n = 1; n <= num; n++) addOpt(n + "/" + den);
        } else {
          addOpt(num + "/" + den);
        }
      } else if (curFracText) {
        addOpt(curFracText);
      }
    } else if (curFracText) {
      addOpt(curFracText);
    }

    const packMax = purePacks ? Math.max(0, curMissing) : openRest;
    for (let k = 1; k <= packMax; k++) addOpt(k + "包");

    addOpt("1箱");
  } else {
    addOpt("1/2");
    addOpt("1/3");
    addOpt("1/4");

    const packMax = Math.min(5, openRest);
    for (let k = 1; k <= packMax; k++) addOpt(k + "包");

    const boxMax = Math.min(5, Math.max(0, curBoxes));
    for (let k = 1; k <= boxMax; k++) addOpt(k + "箱");
  }

  addOpt("清空库存");
  return list;
}

function StockMoves_canonicalizeOutValue_(rawInput) {
  const raw = String(rawInput || "").trim();
  if (!raw) return "";
  const parsed = StockMoves_parseOutValue_(raw);
  if (!parsed.type) return "";

  if (parsed.type === "CLEAR") return "清空库存";
  if (parsed.type === "TAIL") return "(" + StockMoves_toInt_(parsed.qty) + "p)";
  if (parsed.type === "BOXES") return StockMoves_toInt_(parsed.qty) + "箱";
  if (parsed.type === "PACKS") return StockMoves_toInt_(parsed.qty) + "包";
  if (parsed.type === "FRACTION") {
    const reduced = StockMoves_reduceFraction_(StockMoves_toInt_(parsed.num), StockMoves_toInt_(parsed.den));
    return reduced.num + "/" + reduced.den;
  }
  return "";
}

function StockMoves_setOutDropdown_(sh, map, row) {
  const colOut = StockMoves_col_(map, "出-Sortie/箱");
  const colWh = StockMoves_col_(map, "仓库");
  if (!colOut) return;

  const state = StockMoves_stateFromSheetRow_(sh, map, row);
  const colOpenRest = StockMoves_col_(map, "Carton ouvert (reste)");
  const openRest = colOpenRest ? StockMoves_toInt_(sh.getRange(row, colOpenRest).getValue()) : 0;

  const list = StockMoves_getAllowedOutValues_(state, openRest);
  if (list.length <= 1 && list[0] === "清空库存") {
    sh.getRange(row, colOut).clearDataValidations();
    if (colWh) {
      const cellWh = sh.getRange(row, colWh);
      cellWh.clearDataValidations();
      cellWh.setValue("");
    }
    return;
  }

  const ruleOut = SpreadsheetApp.newDataValidation()
    .requireValueInList(list, true)
    .setAllowInvalid(true)
    .build();

  sh.getRange(row, colOut).setDataValidation(ruleOut);

  if (colWh) {
    const ruleWh = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Aulnay", "Epinay"], true)
      .setAllowInvalid(false)
      .build();
    sh.getRange(row, colWh).setDataValidation(ruleWh);
  }
}
