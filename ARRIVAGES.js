/***********************
 * Arrivages_Service.gs — Use cases (load, quick insert)
 ***********************/

function ArrivagesService_load_(arrivageId) {
  const ss = SpreadsheetApp.getActive();
  const ui = ss.getSheetByName(SHEET_UI);
  const db = ss.getSheetByName(SHEET_DB);
  if (!ui || !db) throw new Error("Feuilles manquantes: ARRIVAGES / ARRIVAGES_DB.");

  const id = String(arrivageId || "").trim();
  if (!id || id === LABEL_ADD) return;

  const rows = ArrivagesRepo_getRowsByArrivageId_(db, id);
  if (!rows.length) {
    try { ss.toast("⚠️ ID introuvable dans DB", "ARRIVAGES", 4); } catch (e) {}
    return;
  }

  const model = ArrivagesDomain_buildUiModelFromDbRows_(rows);

  const writeUi = () => {
    ui.getRange(UI_UPDATED_CELL).setValue(model.updatedAtMax || "");
    ui.getRange(UI_CREATED_CELL).setValue(model.createdAt || "");
    ui.getRange(UI_ENTREPOT_CELL).setValue(model.entrepot || "");

    ui.getRange(UI_TABLE_RANGE).clearContent();
    const sliced = model.lines.slice(0, UI_TABLE_ROWS);
    if (sliced.length) {
      ui.getRange(UI_TABLE_START_ROW, 1, sliced.length, UI_TABLE_COLS).setValues(sliced);
    }
  };

  if (typeof withUiGuard_ === "function") withUiGuard_(writeUi);
  else writeUi();
}

function ArrivagesService_quickInsert_(input) {
  const ss = SpreadsheetApp.getActive();
  const ui = ss.getSheetByName(SHEET_UI);
  if (!ui) throw new Error("Feuille ARRIVAGES introuvable.");

  const parsed = ArrivagesDomain_parseQuickInput_(input);
  // parsed = { ref, tail, ppc, cartons }

  const grid = ui.getRange(UI_TABLE_RANGE).getValues();
  let idx = -1;
  for (let i = 0; i < grid.length; i++) {
    const empty = (grid[i] || []).every(c => String(c ?? "").trim() === "");
    if (empty) { idx = i; break; }
  }
  if (idx === -1) throw new Error("表格已满：A4:F305 没有空行");

  const row = UI_TABLE_START_ROW + idx;

  const writeRow = () => {
    ui.getRange(row, 1, 1, UI_TABLE_COLS).setValues([[
      parsed.ref,
      parsed.tail === "" ? "" : parsed.tail,
      parsed.ppc,
      parsed.cartons,
      "", // E NoteSystem
      ""  // F NoteUser
    ]]);
  };

  if (typeof withUiGuard_ === "function") withUiGuard_(writeRow);
  else writeRow();
}
function ArrivagesService_quickInsertMany_(lines) {
  const ss = SpreadsheetApp.getActive();
  const ui = ss.getSheetByName(SHEET_UI);
  if (!ui) throw new Error("Feuille ARRIVAGES introuvable.");

  const parsed = [];
  for (const s of lines) {
    parsed.push(ArrivagesDomain_parseQuickInput_(s));
  }

  const tableRange = ui.getRange(UI_TABLE_RANGE);
  const grid = tableRange.getValues();

  // Trouver la première ligne vide
  let idx = -1;
  for (let i = 0; i < grid.length; i++) {
    const empty = (grid[i] || []).every(c => String(c ?? "").trim() === "");
    if (empty) { idx = i; break; }
  }
  if (idx === -1) throw new Error("表格已满：没有空行");

  const remaining = grid.length - idx;
  if (parsed.length > remaining) {
    throw new Error(`空行不足：还剩 ${remaining} 行，但你粘贴了 ${parsed.length} 行`);
  }

  const out = [];
  for (const p of parsed) {
    out.push([
      p.ref,
      p.tail === "" ? "" : p.tail,
      p.ppc,
      p.cartons,
      "",
      ""
    ]);
  }

  const write = () => {
    ui.getRange(UI_TABLE_START_ROW + idx, 1, out.length, UI_TABLE_COLS).setValues(out);
  };
  if (typeof withUiGuard_ === "function") withUiGuard_(write);
  else write();
}

/***********************
 * Arrivages_Service.gs — Save + Sync STOCK
 ***********************/

// Public (menu)
function arrivagesSaveCurrent() {
  ArrivagesService_saveCurrent_();
}

// Public (menu)
function deleteArrivage_() {
  ArrivagesService_deleteCurrent_();
}

function ArrivagesService_saveCurrent_() {
  const ss = SpreadsheetApp.getActive();
  const ui = ss.getSheetByName(SHEET_UI);
  const db = ss.getSheetByName(SHEET_DB);
  const stock = ss.getSheetByName(SHEET_STOCK);
  const tpl = ss.getSheetByName(SHEET_TEMPLATE_STOCK);
  if (!ui || !db || !stock || !tpl) throw new Error("Feuilles manquantes: ARRIVAGES / ARRIVAGES_DB / STOCK / TEMPLATE_STOCK.");

  const now = new Date();

  // --- read UI header cells
  const entrepot = String(ui.getRange(UI_ENTREPOT_CELL).getValue() || "").trim();
  let id = String(ui.getRange(UI_ID_CELL).getValue() || "").trim();

  

  if (!entrepot) throw new Error("E1 (Entrepot) est vide. Mets Epinay ou Aulnay avant d’enregistrer.");

  // Determine if new
const isNew = (!id || id === LABEL_ADD);

// ✅ Edit seulement si l'ID est réellement présent en DB
const existsInDb = (!isNew) && ArrivagesRepo_existsArrivageId_(db, id);
const isEdit = existsInDb;

if (isNew) {
  id = ArrivagesService_newArrivageId_();
}

if (isEdit) {
  const uiPrompt = SpreadsheetApp.getUi();

  const confirm = uiPrompt.alert(
    "Modifier arrivage",
    "Cet arrivage existe déjà (" + id + ").\n\n" +
    "Les anciennes lignes DB vont être supprimées et le stock des refs de cet arrivage sera réinitialisé avant ré-enregistrement.\n\nContinuer ?",
    uiPrompt.ButtonSet.YES_NO
  );

  if (confirm !== uiPrompt.Button.YES) return;

  // 1️⃣ récupérer les refs existantes
  const oldRefs = ArrivagesRepo_getRefsByArrivageId_(db, id);

  // 2️⃣ supprimer les lignes DB
  ArrivagesRepo_deleteByArrivageId_(db, id);

  // 3️⃣ reset stock pour ces refs
  ArrivagesStock_resetRefsAndDeleteSuffix_(stock, oldRefs);
}

  // CreatedAt logic
  let createdAt = ui.getRange(UI_CREATED_CELL).getValue();
  if (!(createdAt instanceof Date)) createdAt = isNew ? now : "";

  // --- read UI lines A4:F305
  const grid = ui.getRange(UI_TABLE_RANGE).getDisplayValues(); // A..F
  const dbRows = [];
  const payload = []; // for STOCK sync (only rows w/ ref)

  for (let i = 0; i < grid.length; i++) {
    const r = grid[i];
    const rawRef = r[0];
    const ref = (typeof cleanRef_ === "function" ? cleanRef_(rawRef) : String(rawRef || "")).toUpperCase();
    if (!ref) continue;

    const tailParsed = ArrivagesDomain_parseTailInput_(r[1]);
    const tail = Number(tailParsed.total || 0);
    const tailDisplay = String(tailParsed.display || "").trim();
    const ppcParsed = ArrivagesDomain_parsePpcInput_(r[2]);
    const ppc = Number(ppcParsed.primary || 0);
    const ppcDisplay = String(ppcParsed.display || (ppc ? String(ppc) : "")).trim();
    const boxPackParsed = ArrivagesDomain_parseBoxesAndPacks_(r[3], ppc);
    const cartons = Number(boxPackParsed.boxesValue || 0);
    const missingPacks = Number(boxPackParsed.missingPacks || 0);
    const noteS = String(r[4] || "").trim();
    let dbNoteSystem = ArrivagesDomain_mergeBoxPackRawIntoNoteSystem_(noteS, String(r[3] ?? "").trim());
    dbNoteSystem = ArrivagesDomain_mergeTailRawIntoNoteSystem_(dbNoteSystem, tailParsed.raw);
    dbNoteSystem = ArrivagesDomain_mergePpcRawIntoNoteSystem_(dbNoteSystem, ppcParsed.values.length > 1 ? ppcParsed.raw : "");
    const noteU = String(r[5] || "").trim();

    // DB row (10 cols)
    dbRows.push([
      now,         // UpdatedAt
      id,          // ArrivageID
      entrepot,    // Entrepot
      ref,         // 货号
      tail || 0,   // 尾箱件数
      ppc || 0,    // 每箱件数
      cartons || 0,// 标准箱数
      createdAt,   // CreatedAt
      dbNoteSystem,// NoteSystem
      noteU        // NoteUser
    ]);

    // STOCK payload line
    const mixPartner = (typeof parseMixPartnerFromNoteSystem_ === "function")
  ? parseMixPartnerFromNoteSystem_(noteS)
  : "";
  if (/MIX/i.test(noteS) && !mixPartner) {
  throw new Error("MIX détecté mais partenaire introuvable dans NoteSystem: " + ref);
}

// la ref participe à un mix
const isMix = !!mixPartner;

// seule la ligne MIX_START porte 混箱数
const isMixStart = /MIX_START/i.test(noteS);

// valeur qui sera mise dans 混箱数
const mixUsed = (isMixStart && tail > 0);
    payload.push({
  ref: ref,
  entrepot: entrepot,
  arrivageId: id,
  tail: tail || 0,
  tailDisplay: tailDisplay,
  ppc: ppc || 0,
  ppcDisplay: ppcDisplay,
  cartons: cartons || 0,
  missingPacks: missingPacks || 0,
  boxPackRaw: String(r[3] ?? "").trim(),
  boxPackKind: boxPackParsed.kind,
  noteSystem: noteS,
  noteUser: noteU,
  mixUsed: mixUsed,
  isMix: isMix
});
  }

  if (!dbRows.length) throw new Error("Aucune ligne à enregistrer (A4:F305 vide).");

  const writeAll = () => {
    // 1) DB append
    ArrivagesRepo_appendDbRows_(db, dbRows);

    // 2) Update UI meta
    ui.getRange(UI_ID_CELL).setValue(id);
    ui.getRange(UI_CREATED_CELL).setValue(createdAt);
    ui.getRange(UI_UPDATED_CELL).setValue(now);

    // 3) STOCK sync from payload
    ArrivagesStock_applyFromArrivagePayload_(stock, tpl, payload);

    try { ss.toast("✅ Arrivage enregistré + STOCK synchronisé", "ARRIVAGES", 4); } catch (e) {}
  };

  if (typeof withUiGuard_ === "function") withUiGuard_(writeAll);
  else writeAll();
}

function ArrivagesService_deleteCurrent_() {
  const ss = SpreadsheetApp.getActive();
  const ui = ss.getSheetByName(SHEET_UI);
  const db = ss.getSheetByName(SHEET_DB);
  const stock = ss.getSheetByName(SHEET_STOCK);
  if (!ui || !db || !stock) throw new Error("Feuilles manquantes: ARRIVAGES / ARRIVAGES_DB / STOCK.");

  const id = String(ui.getRange(UI_ID_CELL).getValue() || "").trim();
  if (!id || id === LABEL_ADD) {
    try { ss.toast("⚠️ Aucun ArrivageID en B1", "ARRIVAGES", 4); } catch (e) {}
    return;
  }

  const uiPrompt = SpreadsheetApp.getUi();

  const c1 = uiPrompt.alert(
    "Supprimer arrivage",
    "Supprimer toutes les lignes ARRIVAGES_DB avec ArrivageID = " + id + " ?",
    uiPrompt.ButtonSet.YES_NO
  );
  if (c1 !== uiPrompt.Button.YES) return;

  const c2 = uiPrompt.alert(
    "⚠️ Reset STOCK (irréversible)",
    "On va aussi RESET les valeurs de STOCK pour les refs de cet arrivage (et supprimer les refs avec *).\n\nConfirmer ?",
    uiPrompt.ButtonSet.YES_NO
  );
  if (c2 !== uiPrompt.Button.YES) return;

  // 1) Liste des refs concernées (AVANT delete DB)
  const refs = ArrivagesRepo_getRefsByArrivageId_(db, id);

  // 2) Delete DB
  const n = ArrivagesRepo_deleteByArrivageId_(db, id);

  // 3) Reset STOCK + delete suffix refs
  ArrivagesStock_resetRefsAndDeleteSuffix_(stock, refs);

  // 4) Reset UI
  if (typeof ArrivagesUI_reset_ === "function") ArrivagesUI_reset_();

  try { ss.toast("🗑️ Supprimé DB: " + n + " lignes | STOCK reset: " + refs.length + " refs", "ARRIVAGES", 6); } catch (e) {}
}

function ArrivagesService_newArrivageId_() {
  const tz = Session.getScriptTimeZone() || "Europe/Paris";
  const s = Utilities.formatDate(new Date(), tz, "yyyyMMdd-HHmmss");
  const rnd = Math.floor(Math.random() * 900 + 100);
  return "A-" + s + "-" + rnd;
}

/**
 * STOCK sync rules
 * - Match by header "货号"
 * - Insert if missing (+ template formulas)
 * - Safe-add / Replace / Suffix (货号*, 货号**...)
 * - Update dropdowns ONLY for refs touched (no disabling elsewhere)
 * - Rebuild filter A:AS + sort by SortKey
 */
function ArrivagesStock_applyFromArrivagePayload_(shStock, shTpl, payload) {
  if (!payload || !payload.length) return;

  const stockLastRow = shStock.getLastRow();
  const stockLastCol = shStock.getLastColumn();
  if (stockLastRow < 1) throw new Error("STOCK vide (pas d'en-têtes).");

  // Header maps
  const stockHeaders = shStock.getRange(1, 1, 1, stockLastCol).getValues()[0];
  const tplHeaders = shTpl.getRange(1, 1, 1, shTpl.getLastColumn()).getValues()[0];

  const stockMap = (typeof headerMap_ === "function") ? headerMap_(stockHeaders) : ArrivagesStock_headerMapLocal_(stockHeaders);
  const tplMap = (typeof headerMap_ === "function") ? headerMap_(tplHeaders) : ArrivagesStock_headerMapLocal_(tplHeaders);

  // Required columns (skip if not found; but 货号 must exist)
  const colRef = stockMap["货号".toLowerCase()];
  if (!colRef) throw new Error("STOCK: colonne '货号' introuvable.");

  const colTailCur = stockMap["当前尾箱件数".toLowerCase()];
  const colBoxesCur = stockMap["当前箱数".toLowerCase()];
  const colSignCur = stockMap["当前signe".toLowerCase()];
  const colFracCur = stockMap["当前箱数分数".toLowerCase()];
  const colMissingCur = stockMap["当前缺包".toLowerCase()];
  const colMix = stockMap["混箱数".toLowerCase()];
  const colSortKey = stockMap["sortkey"];

  const colPpc2 = stockMap["每箱件数2".toLowerCase()];
  const colIsMix = stockMap["is_mix".toLowerCase()];

  const colLoc = stockMap["放位/提醒".toLowerCase()];
  const colNote2 = stockMap["备注2".toLowerCase()];
  const colArrId = stockMap["到货单".toLowerCase()];
  const colWh = stockMap["仓库".toLowerCase()];
  const colIn = stockMap["进货".toLowerCase()];

  const colOut = stockMap["出-sortie/箱".toLowerCase()];
  const colOpenRest = stockMap["carton ouvert (reste)".toLowerCase()];

  // Build ref->rowIndex map (existing)
  const nData = Math.max(0, stockLastRow - 1);
  const refValues = nData ? shStock.getRange(2, colRef, nData, 1).getValues().flat() : [];
  const refToRow = {};
  for (let i = 0; i < refValues.length; i++) {
    const r = String(refValues[i] || "").trim().toUpperCase();
    if (r) refToRow[r] = 2 + i;
  }

  // Read needed existing cols for decisions
  const readCol = (col) => (col && nData) ? shStock.getRange(2, col, nData, 1).getValues().flat() : [];
  const ppcCol = readCol(colPpc2);
  const tailCurCol = readCol(colTailCur);
  const boxesCurCol = readCol(colBoxesCur);
  const signCurCol = readCol(colSignCur);
  const fracCurCol = readCol(colFracCur);
  const mixCol = readCol(colMix);

  const getExistingAtRow = (rowIdx) => {
    const i = rowIdx - 2;
    const fracRaw = fracCurCol[i];
    const fracTxt = String(fracRaw === null || typeof fracRaw === "undefined" ? "" : fracRaw).trim();
    return {
      ppc: colPpc2 ? ArrivagesStock_toNumber_(ppcCol[i]) : 0,
      tailCur: colTailCur ? ArrivagesStock_tailDisplayToTotal_(tailCurCol[i]) : 0,
      boxesCur: colBoxesCur ? ArrivagesStock_toNumber_(boxesCurCol[i]) : 0,
      signCur: colSignCur ? String(signCurCol[i] || "").trim() : "",
      hasFractionCur: colFracCur ? (fracTxt !== "" && fracTxt !== "0") : false,
      mix: colMix ? ArrivagesStock_toNumber_(mixCol[i]) : 0
    };
  };

  // Prepare append buffer
  const rowsToAppend = [];
  const appendMeta = []; // {targetRef, valuesToWrite}
  const updateMeta = []; // existing row updates {row, op, fields...}
  const touchedRowIndices = []; // for dropdowns later

  const ensureSuffixRef_ = (baseRef) => {
    let k = 1;
    while (k < 200) {
      const stars = Array(k + 1).join("*");
      const candidate = baseRef + stars;
      if (!refToRow[candidate]) return candidate;
      k++;
    }
    throw new Error("Suffix overflow for ref: " + baseRef);
  };

  const safeStr = (v) => String(v || "").trim();

  for (const item of payload) {
    const baseRef = String(item.ref || "").trim().toUpperCase();
    if (!baseRef) continue;

    const newPpc = ArrivagesStock_toNumber_(item.ppc);
    const newPpcDisplay = String(item.ppcDisplay || (newPpc ? String(Math.trunc(newPpc)) : "")).trim();
    const newTail = ArrivagesStock_toNumber_(item.tail);
    const newTailDisplay = String(item.tailDisplay || "").trim();
    const newBoxes = ArrivagesStock_toNumber_(item.cartons);
    const newMissingPacks = ArrivagesStock_toNumber_(item.missingPacks);
    const newHasTail = newTail > 0;
   const newMixUsed = ArrivagesStock_toInt_(item.mixUsed) > 0;
    const newBoxPackKind = String(item.boxPackKind || "plain").trim().toLowerCase() || "plain";
    const newBoxParts = ArrivagesStock_boxPartsFromRaw_(item.boxPackRaw, newBoxPackKind, newBoxes);
    const historyEntry = ArrivagesStock_buildIncomingHistoryEntry_(
      new Date(),
      newPpc,
      newPpcDisplay,
      newBoxParts.whole,
      newBoxParts.sign,
      newBoxParts.fraction,
      newMissingPacks,
      newTail,
      newTailDisplay,
      item.boxPackRaw
    );

    // Decide target
    const baseRow = refToRow[baseRef] || 0;

    let mode = "INSERT"; // INSERT | SAFE_ADD | REPLACE | SUFFIX
    let targetRef = baseRef;
    let targetRow = 0;

    if (!baseRow) {
      mode = "INSERT";
    } else {
      const ex = getExistingAtRow(baseRow);
      const oldActive = (ex.tailCur > 0) || (ex.boxesCur > 0) || (ex.mix > 0);
      const ppcSame = (colPpc2 ? (ex.ppc === newPpc) : true);
      const oldMixUsed = ex.mix > 0;

      if (ppcSame) {
        const existingIsPlain = !ex.signCur && !ex.hasFractionCur;
        const safeAdd = (!newHasTail) && (!newMixUsed) && (newBoxPackKind === "plain") && existingIsPlain;
        if (safeAdd) {
          mode = "SAFE_ADD";
        } else if (oldMixUsed || newMixUsed) {
          mode = "SUFFIX";
        } else if ((ex.tailCur > 0) && (newHasTail)) {
          // old has tail + new has tail => suffix
          mode = "SUFFIX";
        } else if (!oldActive) {
          mode = "REPLACE";
        } else {
          mode = "SUFFIX";
        }
      } else {
        // ppc different
        mode = oldActive ? "SUFFIX" : "REPLACE";
      }
    }

    if (mode === "SUFFIX") {
      targetRef = ensureSuffixRef_(baseRef);
    }

    // Resolve row index
    targetRow = refToRow[targetRef] || 0;

    if (!targetRow && (mode === "INSERT" || mode === "SUFFIX")) {
      // Append new row (blank), formulas later
      const rowArr = new Array(stockLastCol).fill("");
      rowArr[colRef - 1] = targetRef;

      rowsToAppend.push(rowArr);
      appendMeta.push({
        targetRef: targetRef,
        // write values after append (we need actual row index)
        data: {
          entrepot: safeStr(item.entrepot),
          arrivageId: safeStr(item.arrivageId),
          noteUser: safeStr(item.noteUser),
          noteSystem: safeStr(item.noteSystem),
          newPpc: newPpc,
          newPpcDisplay: newPpcDisplay,
          newTail: newTail,
          newTailDisplay: newTailDisplay,
          newBoxes: newBoxes,
          newWholeBoxes: newBoxParts.whole,
          newBoxSign: newBoxParts.sign,
          newBoxFraction: newBoxParts.fraction,
          extraNote2: newBoxParts.extraNote2,
          newMissingPacks: newMissingPacks,
          mixUsed: newMixUsed,
          isMix: !!item.isMix,
          historyEntry: historyEntry,
        }
      });

      // Pre-reserve row index (virtual) so suffix collisions don’t happen inside same save
      refToRow[targetRef] = -1;
    } else {
      // Existing row update
      updateMeta.push({
        row: targetRow || baseRow,
        mode: mode,
        entrepot: safeStr(item.entrepot),
        arrivageId: safeStr(item.arrivageId),
        noteUser: safeStr(item.noteUser),
        noteSystem: safeStr(item.noteSystem),
        newPpc: newPpc,
        newPpcDisplay: newPpcDisplay,
        newTail: newTail,
        newTailDisplay: newTailDisplay,
        newBoxes: newBoxes,
        newWholeBoxes: newBoxParts.whole,
        newBoxSign: newBoxParts.sign,
        newBoxFraction: newBoxParts.fraction,
        extraNote2: newBoxParts.extraNote2,
        newMissingPacks: newMissingPacks,
        mixUsed: newMixUsed,
        isMix: !!item.isMix,
        historyEntry: historyEntry,
      });
    }
  }

  // 1) Append rows in one shot
  let appendStartRow = 0;
  if (rowsToAppend.length) {
    appendStartRow = shStock.getLastRow() + 1;
    shStock.getRange(appendStartRow, 1, rowsToAppend.length, stockLastCol).setValues(rowsToAppend);

    // Apply template formulas
    const formulaCols = ["包/箱", "Colisage", "Poids (en gramme)", "Pays d'origine", "Promo", "Prix@", "Promo@", "剩下 / RESTE", "SortKey"];

// Keep only headers present in both TEMPLATE and STOCK
const safeFormulaCols = formulaCols.filter(h => {
  const k = String(h).toLowerCase();
  return !!tplMap[k] && !!stockMap[k];
});

if (safeFormulaCols.length) {
  if (typeof applyTemplateFormulas_ === "function") {
    applyTemplateFormulas_(shTpl, shStock, tplMap, stockMap, safeFormulaCols, rowsToAppend.length, appendStartRow);
  } else {
    ArrivagesStock_applyTemplateFormulasLocal_(shTpl, shStock, tplMap, stockMap, safeFormulaCols, rowsToAppend.length, appendStartRow);
  }
}
// KEY:* columns detected from header notes keep their template formulas.
ArrivagesStock_applyTemplateFormulasByHeaderNoteKey_(
  shTpl,
  shStock,
  ["KEY:TOTAL_BOX", "KEY:TOTAL_PCS"],
  rowsToAppend.length,
  appendStartRow
);

    // Fix refToRow for appended
    for (let i = 0; i < appendMeta.length; i++) {
      const r = appendStartRow + i;
      const ref = appendMeta[i].targetRef;
      refToRow[ref] = r;
      touchedRowIndices.push(r);
    }
  }

  // 2) Write values (updates + appended metas)
  const writes = [];

  // Helper to schedule writes (single cell)
  const pushWrite = (row, col, val) => {
    if (!row || !col) return;
    writes.push({ row: row, col: col, val: val });
  };

  const concatIfNeeded = (oldVal, addVal) => {
    const a = String(oldVal || "").trim();
    const b = String(addVal || "").trim();
    if (!b) return a;
    if (!a) return b;
    // keep simple
    return a + "\n" + b;
  };
  const concatNewestFirst = (oldVal, addVal) => {
    const a = String(oldVal || "").trim();
    const b = String(addVal || "").trim();
    if (!b) return a;
    if (!a) return b;
    return b + "\n" + a;
  };

  // Apply appended meta writes
  for (const meta of appendMeta) {
    const row = refToRow[meta.targetRef];
    const d = meta.data;

    // Replace/Insert/Suffix all behave same for new line: set values fresh
    if (colTailCur) pushWrite(row, colTailCur, d.newTailDisplay || "");
    if (colPpc2) pushWrite(row, colPpc2, d.newPpcDisplay || (d.newPpc || 0));
    if (colBoxesCur) pushWrite(row, colBoxesCur, d.newWholeBoxes || 0);
    if (colSignCur) pushWrite(row, colSignCur, d.newBoxSign || "");
    if (colFracCur) pushWrite(row, colFracCur, d.newBoxFraction || "");
    if (colMissingCur) pushWrite(row, colMissingCur, d.newMissingPacks || 0);

    if (colMix) pushWrite(row, colMix, d.mixUsed ? 1 : 0);
    if (colIsMix) pushWrite(row, colIsMix, d.isMix ? 1 : 0);

    if (colArrId) pushWrite(row, colArrId, d.arrivageId);
    if (colWh) pushWrite(row, colWh, d.entrepot);
    if (colIn) pushWrite(row, colIn, d.historyEntry || "");

    // Notes mapping
    if (colNote2) pushWrite(row, colNote2, [d.noteUser, d.extraNote2].filter(Boolean).join("\n"));
    if (colLoc)   pushWrite(row, colLoc, d.noteSystem || "");
  }

  // Apply existing updates
  for (const up of updateMeta) {
    const row = up.row;
    if (!row) continue;
    touchedRowIndices.push(row);

    if (up.mode === "SAFE_ADD") {
      // Only add current boxes/missing; do NOT touch tails
      if (colBoxesCur) {
        const old = ArrivagesStock_toNumber_(shStock.getRange(row, colBoxesCur).getValue());
        pushWrite(row, colBoxesCur, old + (up.newWholeBoxes || 0));
      }
      if (colMissingCur) {
        const old = ArrivagesStock_toNumber_(shStock.getRange(row, colMissingCur).getValue());
        pushWrite(row, colMissingCur, old + (up.newMissingPacks || 0));
      }
      if (colIn && up.historyEntry) {
        const oldIn = shStock.getRange(row, colIn).getValue();
        pushWrite(row, colIn, concatNewestFirst(oldIn, up.historyEntry));
      }

      // Notes: concat
      if (colNote2 && up.noteUser) {
        const oldN = shStock.getRange(row, colNote2).getValue();
        pushWrite(row, colNote2, concatIfNeeded(oldN, up.noteUser));
      }
      if (colLoc && up.noteSystem) {
        const oldS = shStock.getRange(row, colLoc).getValue();
        pushWrite(row, colLoc, concatIfNeeded(oldS, up.noteSystem));
      }

      // Arrivage ID: concat
      if (colArrId && up.arrivageId) {
        const oldA = shStock.getRange(row, colArrId).getValue();
        pushWrite(row, colArrId, concatIfNeeded(oldA, up.arrivageId));
      }

      // Warehouse: only fill if empty
      if (colWh && up.entrepot) {
        const oldW = String(shStock.getRange(row, colWh).getValue() || "").trim();
        if (!oldW) pushWrite(row, colWh, up.entrepot);
      }

    } else {
      // REPLACE (or any non-safe update on existing line): replace full set
      if (colTailCur) pushWrite(row, colTailCur, up.newTailDisplay || "");
      if (colPpc2) pushWrite(row, colPpc2, up.newPpcDisplay || (up.newPpc || 0));
      if (colBoxesCur) pushWrite(row, colBoxesCur, up.newWholeBoxes || 0);
      if (colSignCur) pushWrite(row, colSignCur, up.newBoxSign || "");
      if (colFracCur) pushWrite(row, colFracCur, up.newBoxFraction || "");
      if (colMissingCur) pushWrite(row, colMissingCur, up.newMissingPacks || 0);

      if (colMix) pushWrite(row, colMix, up.mixUsed ? 1 : 0);
      if (colIsMix) pushWrite(row, colIsMix, (up.isMix ? 1 : 0));

      if (colArrId) pushWrite(row, colArrId, up.arrivageId);
      if (colWh) pushWrite(row, colWh, up.entrepot);
      if (colIn && up.historyEntry) {
        const oldIn = shStock.getRange(row, colIn).getValue();
        pushWrite(row, colIn, concatNewestFirst(oldIn, up.historyEntry));
      }

      if (colNote2) pushWrite(row, colNote2, [up.noteUser, up.extraNote2].filter(Boolean).join("\n"));
      if (colLoc) pushWrite(row, colLoc, up.noteSystem || "");
    }
  }

  // Batch apply writes (group by row for performance)
  if (writes.length) {
    const byRow = {};
    for (const w of writes) {
      if (!byRow[w.row]) byRow[w.row] = [];
      byRow[w.row].push(w);
    }
    for (const rowStr in byRow) {
      const row = Number(rowStr);
      const list = byRow[row].sort((a,b)=>a.col-b.col);
      // Write cell-by-cell (safe & simple)
      for (const w of list) shStock.getRange(row, w.col).setValue(w.val);
    }
  }

  // 3) Rows touched by this sync
  const uniqTouched = Array.from(new Set(touchedRowIndices)).filter(r => r && r > 1);

  // 4) Dropdowns only for touched rows

  // 仓库 dropdown (activate only for touched)
  if (colWh) {
    const ruleWh = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Epinay", "Aulnay"], true)
      .setAllowInvalid(true)
      .build();
    for (const r of uniqTouched) shStock.getRange(r, colWh).setDataValidation(ruleWh);
  }

  // 出-Sortie/箱 dropdown (activate only for touched)
  if (colOut) {
    const toFracText = (v) => {
      const s = String(v === null || typeof v === "undefined" ? "" : v).trim();
      if (!s) return "";
      if (/^\d+\/\d+$/.test(s)) return s;
      if (/^\d+\/\d+(?:\+\d+\/\d+)+$/.test(s)) return s;
      if (/^\d+(\.\d+)?$/.test(s)) return ArrivagesStock_fractionToText_(Number(s));
      return "";
    };
    const parseFrac = (txt) => {
      const m = String(txt || "").trim().match(/^(\d+)\/(\d+)$/);
      if (!m) return null;
      const num = Number(m[1]);
      const den = Number(m[2]);
      if (!den || num <= 0) return null;
      return ArrivagesStock_reduceFraction_(num, den);
    };

    for (const r of uniqTouched) {
      const curBoxes = colBoxesCur ? ArrivagesStock_toNumber_(shStock.getRange(r, colBoxesCur).getValue()) : 0;
      const curTail = colTailCur ? ArrivagesStock_tailDisplayToTotal_(shStock.getRange(r, colTailCur).getValue()) : 0;
      const curSign = colSignCur ? String(shStock.getRange(r, colSignCur).getValue() || "").trim() : "";
      const fracRaw = colFracCur ? shStock.getRange(r, colFracCur).getValue() : "";
      const curFracText = toFracText(fracRaw);
      const curMissing = colMissingCur ? ArrivagesStock_toNumber_(shStock.getRange(r, colMissingCur).getValue()) : 0;
      const openRest = colOpenRest ? ArrivagesStock_toNumber_(shStock.getRange(r, colOpenRest).getValue()) : 0;
      const hasFraction = !!curFracText;
      const isOpen = (curMissing !== 0) || !!curSign || hasFraction;
      const purePacks = curMissing > 0 && curBoxes <= 0 && !curSign && !hasFraction;
      const packMaxOpen = purePacks ? Math.max(0, Math.floor(curMissing)) : Math.max(0, Math.floor(openRest));
      const packMaxPlain = Math.min(5, Math.max(0, Math.floor(openRest)));
      const boxMaxPlain = Math.min(5, Math.max(0, Math.floor(curBoxes)));

      if (curBoxes <= 0 && curTail <= 0 && !isOpen && packMaxPlain <= 0) continue;

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
          const frac = parseFrac(curFracText);
          if (frac) {
            if (frac.num > 1) {
              for (let n = 1; n <= frac.num; n++) addOpt(n + "/" + frac.den);
            } else {
              addOpt(frac.num + "/" + frac.den);
            }
          } else if (curFracText) {
            addOpt(curFracText);
          }
        } else if (curFracText) {
          addOpt(curFracText);
        }

        for (let k = 1; k <= packMaxOpen; k++) addOpt(k + "包");
        addOpt("1箱");
      } else {
        addOpt("1/2");
        addOpt("1/3");
        addOpt("1/4");
        for (let k = 1; k <= packMaxPlain; k++) addOpt(k + "包");
        for (let k = 1; k <= boxMaxPlain; k++) addOpt(k + "箱");
      }

      addOpt("清空库存");

      const ruleOut = SpreadsheetApp.newDataValidation()
        .requireValueInList(list, true)
        .setAllowInvalid(true)
        .build();
      shStock.getRange(r, colOut).setDataValidation(ruleOut);
    }
  }

  // 5) Filter A:AS + sort by SortKey
  if (typeof rebuildLockedFilter_A_to_AS_ === "function") rebuildLockedFilter_A_to_AS_(shStock);
  else ArrivagesStock_rebuildFilterAtoASLocal_(shStock);

  if (colSortKey) {
    const last = shStock.getLastRow();
    if (last >= 2) {
      // sort only within A:AS (45 cols)
      shStock.getRange(2, 1, last - 1, 45).sort({ column: colSortKey, ascending: true });
    }
  }
}

// ---- Local fallbacks (if Microstore helpers not loaded)

function ArrivagesStock_headerMapLocal_(headersRow) {
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

function ArrivagesStock_applyTemplateFormulasLocal_(shTpl, shStock, tplHeaderMap, stockHeaderMap, formulaCols, addCount, startRow) {
  if (!addCount || addCount <= 0) return;
  for (let i = 0; i < formulaCols.length; i++) {
    const h = formulaCols[i];
    const tplCol = tplHeaderMap[String(h).toLowerCase()];
    const stockCol = stockHeaderMap[String(h).toLowerCase()];
    if (!tplCol || !stockCol) continue;

    const f = shTpl.getRange(2, tplCol).getFormulaR1C1();
    if (!f) continue;

    const formulas = [];
    for (let r = 0; r < addCount; r++) formulas.push([f]);
    shStock.getRange(startRow, stockCol, addCount, 1).setFormulasR1C1(formulas);
  }
}

function ArrivagesStock_rebuildFilterAtoASLocal_(sh) {
  const lastRow = Math.max(1, sh.getLastRow());
  const existing = sh.getFilter();
  if (existing) existing.remove();
  sh.getRange(1, 1, lastRow, 45).createFilter(); // A..AS
}

function ArrivagesStock_toInt_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return Math.trunc(v);
  const s = String(v).trim();
  if (!s) return 0;
  const m = s.match(/-?\d+(\.\d+)?/);
  if (!m) return 0;
  const n = Number(m[0]);
  return Number.isFinite(n) ? Math.trunc(n) : 0;
}

function ArrivagesStock_tailDisplayToTotal_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return Number.isFinite(v) ? Math.trunc(v) : 0;
  const matches = String(v).match(/-?\d+/g);
  if (!matches || !matches.length) return 0;
  return matches.reduce((sum, part) => sum + (Number(part) || 0), 0);
}

function ArrivagesStock_toNumber_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const s = String(v).trim().replace(",", ".");
  if (!s) return 0;
  const m = s.match(/-?\d+(\.\d+)?/);
  if (!m) return 0;
  const n = Number(m[0]);
  return Number.isFinite(n) ? n : 0;
}

function ArrivagesStock_gcd_(a, b) {
  a = Math.abs(Math.trunc(a || 0));
  b = Math.abs(Math.trunc(b || 0));
  while (b) {
    const t = b;
    b = a % b;
    a = t;
  }
  return a || 1;
}

function ArrivagesStock_reduceFraction_(num, den) {
  if (!den) return { num: 0, den: 1 };
  const g = ArrivagesStock_gcd_(num, den);
  return { num: num / g, den: den / g };
}

function ArrivagesStock_boxPartsFromRaw_(rawInput, kind, boxesValue) {
  const raw = String(rawInput || "").trim();
  const compact = raw.replace(/\s+/g, " ").trim();
  const noSpace = compact.replace(/\s+/g, "");
  const safeKind = String(kind || "plain").trim().toLowerCase();
  const asFractionText = (num, den) => {
    if (!den) return "";
    const reduced = ArrivagesStock_reduceFraction_(Number(num), Number(den));
    return reduced.num + "/" + reduced.den;
  };

  if (safeKind === "plain") {
    return {
      whole: Math.max(0, Math.trunc(Number(boxesValue) || 0)),
      sign: "",
      fraction: "",
      extraNote2: ""
    };
  }

  if (safeKind === "packs_delta") {
    return {
      whole: Math.max(0, Math.trunc(Number(boxesValue) || 0)),
      sign: "",
      fraction: "",
      extraNote2: ""
    };
  }

  let m = compact.match(/^(\d+)\s+(\d+)\/(\d+)$/);
  if (m) {
    return {
      whole: Number(m[1]),
      sign: "+",
      fraction: asFractionText(m[2], m[3]),
      extraNote2: ""
    };
  }

  m = noSpace.match(/^×(\d+)$/);
  if (m) {
    return {
      whole: Number(m[1]),
      sign: "",
      fraction: "",
      extraNote2: ""
    };
  }

  m = noSpace.match(/^×(\d+)箱$/);
  if (m) {
    return {
      whole: Number(m[1]),
      sign: "",
      fraction: "",
      extraNote2: ""
    };
  }

  m = noSpace.match(/^×(\d+)\+(\d+)\/(\d+)$/);
  if (m) {
    return {
      whole: Number(m[1]),
      sign: "+",
      fraction: asFractionText(m[2], m[3]),
      extraNote2: ""
    };
  }

  m = noSpace.match(/^×(\d+)\+(\d+)\/(\d+)\+(\d+)\/(\d+)$/);
  if (m) {
    return {
      whole: Number(m[1]),
      sign: "+",
      fraction: asFractionText(m[2], m[3]) + "+" + asFractionText(m[4], m[5]),
      extraNote2: ""
    };
  }

  m = noSpace.match(/^(\d+)\+(\d+)\/(\d+)$/);
  if (m) {
    return {
      whole: Number(m[1]),
      sign: "+",
      fraction: asFractionText(m[2], m[3]),
      extraNote2: ""
    };
  }

  m = noSpace.match(/^\+(\d+)\/(\d+)(?:[+-]\d+包)?$/);
  if (m) {
    return {
      whole: 1,
      sign: "+",
      fraction: asFractionText(m[1], m[2]),
      extraNote2: ""
    };
  }

  m = noSpace.match(/^(\d+)\/(\d+)\+(\d+)\/(\d+)$/);
  if (m) {
    return {
      whole: 1,
      sign: "+",
      fraction: asFractionText(m[1], m[2]) + "+" + asFractionText(m[3], m[4]),
      extraNote2: ""
    };
  }

  m = noSpace.match(/^(\d+)\/(\d+)$/);
  if (m) {
    return {
      whole: 1,
      sign: "×",
      fraction: asFractionText(m[1], m[2]),
      extraNote2: ""
    };
  }

  return {
    whole: Math.max(0, Math.trunc(Number(boxesValue) || 0)),
    sign: "",
    fraction: "",
    extraNote2: ""
  };
}

function ArrivagesStock_fractionToText_(v) {
  const raw = String(v || "").trim();
  if (raw && /^\d+\/\d+(?:\+\d+\/\d+)*$/.test(raw)) return raw;

  const n = Number(v);
  if (!Number.isFinite(n) || n <= 0) return "";

  const candidates = [2, 3, 4, 5, 6, 8, 12];
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
  const reduced = ArrivagesStock_reduceFraction_(best.num, best.den);
  return reduced.num + "/" + reduced.den;
}

function ArrivagesStock_buildIncomingHistoryEntry_(dt, ppc, ppcDisplay, whole, sign, fraction, missingPacks, tail, tailDisplay, boxPackRaw) {
  const tz = Session.getScriptTimeZone() || "Europe/Paris";
  const stamp = Utilities.formatDate(dt instanceof Date ? dt : new Date(), tz, "yyyy-MM-dd HH:mm");

  const ppcDisplayTxt = String(ppcDisplay || "").trim();
  const ppcTxt = ppcDisplayTxt
    ? ppcDisplayTxt
        .split("+")
        .map(part => String(part || "").trim())
        .filter(Boolean)
        .map(part => part + "p")
        .join("+")
    : (Number(ppc) > 0 ? String(Math.trunc(Number(ppc))) + "p" : "");
  const wholeN = Math.max(0, Math.trunc(Number(whole) || 0));
  const signTxt = String(sign || "").trim();
  const fracTxt = ArrivagesStock_fractionToText_(fraction);
  const missN = Number(missingPacks) || 0;
  const tailN = Number(tail) || 0;
  const tailDisplayTxt = String(tailDisplay || "").trim();
  const boxPackRawTxt = String(boxPackRaw || "").trim();

  let core = "";
  if (ppcTxt) {
    if (boxPackRawTxt) {
      core = ppcTxt + boxPackRawTxt;
    } else if (signTxt === "×") {
      core = ppcTxt + "×" + (wholeN <= 1 ? "" : String(wholeN) + "×") + fracTxt;
    } else if (signTxt === "+") {
      core = ppcTxt + (wholeN <= 1 ? "" : "×" + String(wholeN)) + "+" + fracTxt;
    } else {
      core = ppcTxt + (wholeN > 1 ? "×" + String(wholeN) : "");
    }

    if (missN !== 0 && !boxPackRawTxt) core += (missN > 0 ? "+" : "") + String(missN) + "包";
  }

  if (tailN > 0) {
    const tailText = tailDisplayTxt
      ? tailDisplayTxt
          .split("+")
          .map(part => String(part || "").trim())
          .filter(Boolean)
          .map(part => "(" + part + "p)")
          .join("+")
      : "(" + String(Math.trunc(tailN)) + "p)";
    core = tailText + (core ? "+" + core : "");
  }

  return stamp + " | " + core;
}

function ArrivagesStock_findColByHeaderNoteKey_(sheet, key) {
  const lastCol = sheet.getLastColumn();
  for (let c = 1; c <= lastCol; c++) {
    const note = sheet.getRange(1, c).getNote();
    if (note && String(note).indexOf(key) !== -1) return c;
  }
  return 0;
}

function ArrivagesStock_applyTemplateFormulasByHeaderNoteKey_(shTpl, shStock, keys, addCount, startRow) {
  if (!addCount || addCount <= 0) return;

  for (const key of keys) {
    const tplCol = ArrivagesStock_findColByHeaderNoteKey_(shTpl, key);
    const stockCol = ArrivagesStock_findColByHeaderNoteKey_(shStock, key);
    if (!tplCol || !stockCol) continue;

    const f = shTpl.getRange(2, tplCol).getFormulaR1C1();
    if (!f) continue;

    const formulas = Array.from({ length: addCount }, () => [f]);
    shStock.getRange(startRow, stockCol, addCount, 1).setFormulasR1C1(formulas);
  }
}

function ArrivagesStock_resetRefsAndDeleteSuffix_(shStock, refs) {
  if (!refs || !refs.length) return;

  const lastRow = shStock.getLastRow();
  const lastCol = shStock.getLastColumn();
  if (lastRow < 1) return;

  // headers
  const headers = shStock.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = (typeof headerMap_ === "function") ? headerMap_(headers) : ArrivagesStock_headerMapLocal_(headers);

  const colRef      = map["货号"];
  const colTailCur  = map["当前尾箱件数"];
  const colPpc2     = map["每箱件数2"];
  const colBoxesCur = map["当前箱数"];
  const colSignCur  = map["当前signe"];
  const colFracCur  = map["当前箱数分数"];
  const colMissingCur = map["当前缺包"];
  const colMix      = map["混箱数"];
  const colIsMix    = map["is_mix"];
  const colLoc      = map["放位/提醒"];
  const colNote2    = map["备注2"];
  const colArrId    = map["到货单"];
  const colWh       = map["仓库"];
  const colOut      = map["出-sortie/箱"];

  if (!colRef) throw new Error("STOCK: colonne '货号' introuvable.");

  const nData = Math.max(0, lastRow - 1);
  if (!nData) return;

  const refCol = shStock.getRange(2, colRef, nData, 1).getValues().flat().map(v => String(v || "").trim().toUpperCase());

  // index existing
  const refToRow = {};
  for (let i = 0; i < refCol.length; i++) {
    if (refCol[i]) refToRow[refCol[i]] = 2 + i;
  }

  // rows to delete (suffix)
  const rowsToDelete = [];

  // rows to reset (base)
  const resets = []; // {row, col, val}

  const pushReset = (row, col, val) => {
    if (!row || !col) return;
    resets.push({ row, col, val });
  };

  // build a Set for quick contains
  const wanted = {};
  for (const r of refs) wanted[String(r || "").trim().toUpperCase()] = true;

  // 1) collect suffix rows to delete
  for (let i = 0; i < refCol.length; i++) {
    const r = refCol[i];
    if (!r) continue;

    // match base + stars only: BASE*, BASE**, ...
    for (const base in wanted) {
      if (!base) continue;
      if (r.length <= base.length) continue;
      if (r.startsWith(base) && /^[*]+$/.test(r.slice(base.length))) {
        rowsToDelete.push(2 + i);
        break;
      }
    }
  }

  // 2) reset base rows values
  for (const base in wanted) {
    const row = refToRow[base];
    if (!row) continue;

    pushReset(row, colTailCur, 0);
    pushReset(row, colPpc2, 0);
    pushReset(row, colBoxesCur, 0);
    pushReset(row, colSignCur, "");
    pushReset(row, colFracCur, "");
    pushReset(row, colMissingCur, 0);

    pushReset(row, colMix, 0);
    pushReset(row, colIsMix, 0);

    // text columns -> ""
    pushReset(row, colLoc, "");
    pushReset(row, colNote2, "");
    pushReset(row, colArrId, "");
    pushReset(row, colWh, "");
    pushReset(row, colOut, "");
    // Keep 进货 history untouched.
  }

  // apply resets (simple & safe)
  if (resets.length) {
    // group by row, but write cell-by-cell to avoid messing with formulas elsewhere
    for (const r of resets) {
      shStock.getRange(r.row, r.col).setValue(r.val);
    }
  }

  // delete suffix rows bottom-up
  rowsToDelete.sort((a,b)=>b-a);
  for (const r of rowsToDelete) shStock.deleteRow(r);

  // rebuild filter A:AS + sort if you want
  try {
    if (typeof rebuildLockedFilter_A_to_AS_ === "function") rebuildLockedFilter_A_to_AS_(shStock);
    else ArrivagesStock_rebuildFilterAtoASLocal_(shStock);
  } catch (e) {}

  // sort by SortKey if exists
  const colSortKey = map["sortkey"];
  if (colSortKey) {
    const lr = shStock.getLastRow();
    if (lr >= 2) shStock.getRange(2, 1, lr - 1, 45).sort({ column: colSortKey, ascending: true });
  }
}

/***********************
 * Arrivages_Repo.gs — DB access only (Sheets I/O)
 ***********************/

function ArrivagesRepo_getRowsByArrivageId_(dbSheet, id) {
  const wanted = String(id || "").trim();
  if (!wanted) return [];

  const last = dbSheet.getLastRow();
  if (last < 2) return [];

  // Col B = arrivageId (range: B2:B)
  const rng = dbSheet.getRange(2, 2, last - 1, 1);
  const hits = rng.createTextFinder(wanted).matchEntireCell(true).findAll();
  if (!hits || !hits.length) return [];

  const rows = hits.map(c => c.getRow()).sort((a,b)=>a-b);

  const out = [];
  let s = rows[0], p = rows[0];
  const flush = (a,b) => out.push(...dbSheet.getRange(a, 1, b-a+1, 10).getValues());

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r === p + 1) p = r;
    else { flush(s,p); s = p = r; }
  }
  flush(s,p);

  return out;
}

/***********************
 * Arrivages_Repo.gs — DB write ops
 ***********************/

function ArrivagesRepo_appendDbRows_(dbSheet, rows10) {
  if (!rows10 || !rows10.length) return;
  dbSheet.getRange(dbSheet.getLastRow() + 1, 1, rows10.length, 10).setValues(rows10);
}

function ArrivagesRepo_deleteByArrivageId_(dbSheet, id) {
  const wanted = String(id || "").trim();
  if (!wanted) return 0;

  const last = dbSheet.getLastRow();
  if (last < 2) return 0;

  const rng = dbSheet.getRange(2, 2, last - 1, 1); // Col B
  const hits = rng.createTextFinder(wanted).matchEntireCell(true).findAll();
  if (!hits || !hits.length) return 0;

  const rows = hits.map(c => c.getRow()).sort((a,b)=>b-a); // delete bottom-up
  for (const r of rows) dbSheet.deleteRow(r);
  return rows.length;
}

function ArrivagesRepo_getRefsByArrivageId_(dbSheet, id) {
  const wanted = String(id || "").trim();
  if (!wanted) return [];

  const last = dbSheet.getLastRow();
  if (last < 2) return [];

  // Trouve les lignes DB pour cet ArrivageID (col B)
  const rngId = dbSheet.getRange(2, 2, last - 1, 1);
  const hits = rngId.createTextFinder(wanted).matchEntireCell(true).findAll();
  if (!hits || !hits.length) return [];

  // Récupère les refs (col D) sur les mêmes lignes
  const set = {};
  for (const cell of hits) {
    const r = cell.getRow();
    const ref = String(dbSheet.getRange(r, 4).getValue() || "").trim().toUpperCase(); // col D
    if (ref) set[ref] = true;
  }
  return Object.keys(set);
}

function ArrivagesRepo_existsArrivageId_(dbSheet, id) {
  const wanted = String(id || "").trim();
  if (!wanted) return false;

  const last = dbSheet.getLastRow();
  if (last < 2) return false;

  const hits = dbSheet.getRange(2, 2, last - 1, 1)  // col B ArrivageID
    .createTextFinder(wanted)
    .matchEntireCell(true)
    .findNext();

  return !!hits;
}

/***********************
 * Arrivages_Domain.gs — Pure domain logic (parse, DB->UI aggregation)
 * Pas de SpreadsheetApp ici.
 ***********************/

function ArrivagesDomain_parseBoxesAndPacks_(input, ppcInput) {
  let raw = String(input || "").trim().replace(/^'+\s*/, "");

  // rule: empty 箱数/包 means:
  // - 1 carton if 每箱件数 exists
  // - 0 carton if 每箱件数 is empty
  if (!raw) {
    return {
      boxesValue: Number(ppcInput) > 0 ? 1 : 0,
      missingPacks: 0,
      kind: "plain"
    };
  }

  // tolerate spaces around operators while preserving mixed fraction with space, e.g. "2 1/2"
  const compact = raw.replace(/\s+/g, " ").trim();
  const noSpace = compact.replace(/\s+/g, "");
  const hasPpc = Number(ppcInput) > 0;

  // integer + fraction (ex: 1+1/4, 1 + 1/4, 2 1/2) -> fractional
  let m = compact.match(/^(\d+)\s+(\d+)\/(\d+)$/);
  if (m) {
    const whole = Number(m[1]);
    const n = Number(m[2]);
    const d = Number(m[3]);
    if (!d) throw new Error("箱数/包 分数格式错误: " + raw);
    return {
      boxesValue: whole + n / d,
      missingPacks: 0,
      kind: "fractional"
    };
  }

  m = noSpace.match(/^(\d+)\+(\d+)\/(\d+)$/);
  if (m) {
    const whole = Number(m[1]);
    const n = Number(m[2]);
    const d = Number(m[3]);
    if (!d) throw new Error("箱数/包 分数格式错误: " + raw);
    return {
      boxesValue: whole + n / d,
      missingPacks: 0,
      kind: "fractional"
    };
  }

  // leading +fraction (ex: +2/3) -> treat as 1 + fraction
  m = noSpace.match(/^\+(\d+)\/(\d+)$/);
  if (m) {
    const n = Number(m[1]);
    const d = Number(m[2]);
    if (!d) throw new Error("箱数/包 分数格式错误: " + raw);
    return {
      boxesValue: 1 + n / d,
      missingPacks: 0,
      kind: "fractional"
    };
  }

  // leading +fraction with packs delta (ex: +3/4-5包, +1/2+3包)
  m = noSpace.match(/^\+(\d+)\/(\d+)([+-])(\d+)包$/);
  if (m) {
    const n = Number(m[1]);
    const d = Number(m[2]);
    const op = m[3];
    const packs = Number(m[4]);
    if (!d) throw new Error("箱数/包 分数格式错误: " + raw);
    return {
      boxesValue: 1 + n / d,
      missingPacks: op === "+" ? packs : -packs,
      kind: "fractional"
    };
  }

  // normalize spaces for other cases
  raw = noSpace;

  // tolerate explicit carton markers: ×4箱, 4箱, ×4
  m = raw.match(/^×?(\d+)箱?$/);
  if (m) {
    return {
      boxesValue: Number(m[1]),
      missingPacks: 0,
      kind: "plain"
    };
  }

  // explicit cartons + fraction (ex: ×10+1/2, ×2+2/3)
  m = raw.match(/^×(\d+)\+(\d+)\/(\d+)$/);
  if (m) {
    const whole = Number(m[1]);
    const n = Number(m[2]);
    const d = Number(m[3]);
    if (!d) throw new Error("箱数/包 分数格式错误: " + raw);
    return {
      boxesValue: whole + n / d,
      missingPacks: 0,
      kind: "fractional"
    };
  }

  // explicit cartons + fraction + fraction (ex: ×6+2/3+1/2)
  m = raw.match(/^×(\d+)\+(\d+)\/(\d+)\+(\d+)\/(\d+)$/);
  if (m) {
    const whole = Number(m[1]);
    const n1 = Number(m[2]);
    const d1 = Number(m[3]);
    const n2 = Number(m[4]);
    const d2 = Number(m[5]);
    if (!d1 || !d2) throw new Error("箱数/包 分数格式错误: " + raw);
    return {
      boxesValue: whole + (n1 / d1) + (n2 / d2),
      missingPacks: 0,
      kind: "fractional"
    };
  }

  // packs only with 包 suffix (ex: +3包, 10包, -2包)
  // if 每箱件数 exists, treat as 1 carton + packs; otherwise 0 carton + packs
  m = raw.match(/^([+-]?)(\d+)包$/);
  if (m) {
    const sign = m[1] === "-" ? -1 : 1;
    const packs = Number(m[2]) * sign;
    return {
      boxesValue: hasPpc ? 1 : 0,
      missingPacks: packs,
      kind: "packs_delta"
    };
  }

  // explicit cartons + packs with 包 suffix (ex: ×2+8包, ×2-5包)
  m = raw.match(/^×(\d+)([+-])(\d+)包$/);
  if (m) {
    return {
      boxesValue: Number(m[1]),
      missingPacks: m[2] === "+" ? Number(m[3]) : -Number(m[3]),
      kind: "packs_delta"
    };
  }

  // mixed fraction with packs delta (ex: 2+2/3-5, 2+2/3+5)
  m = raw.match(/^(\d+)\+(\d+)\/(\d+)([+-])(\d+)$/);
  if (m) {
    const whole = Number(m[1]);
    const n = Number(m[2]);
    const d = Number(m[3]);
    const op = m[4];
    const packs = Number(m[5]);
    if (!d) throw new Error("箱数/包 分数格式错误: " + raw);
    return {
      boxesValue: whole + n / d,
      missingPacks: op === "+" ? packs : -packs,
      kind: "fractional"
    };
  }

  // fraction + fraction (ex: 2/3+2/3) -> fractional
  if (/^\d+\/\d+\+\d+\/\d+$/.test(raw)) {
    const [a, b] = raw.split("+");
    const [n1, d1] = a.split("/").map(Number);
    const [n2, d2] = b.split("/").map(Number);
    if (!d1 || !d2) throw new Error("箱数/包 分数格式错误: " + raw);

    const v1 = n1 / d1;
    const v2 = n2 / d2;

    return {
      boxesValue: v1 + v2,
      missingPacks: 0,
      kind: "fractional"
    };
  }

  // pure fraction (ex: 1/2, 2/3) -> fractional
  if (/^\d+\/\d+$/.test(raw)) {
    const [n, d] = raw.split("/").map(Number);
    if (!d) throw new Error("箱数/包 分数格式错误: " + raw);

    return {
      boxesValue: n / d,
      missingPacks: 0,
      kind: "fractional"
    };
  }

  // integer +/- packs (ex: 3+7, 2-4) -> packs_delta
  if (/^\d+[+-]\d+$/.test(raw)) {
    const sign = raw.includes("+") ? "+" : "-";
    const [boxes, packs] = raw.split(/[+-]/);

    return {
      boxesValue: Number(boxes),
      missingPacks: sign === "+" ? Number(packs) : -Number(packs),
      kind: "packs_delta"
    };
  }

  // leading +packs (ex: +3) -> if 每箱件数 exists, treat as 1 carton + packs; otherwise 0 carton + packs
  if (/^\+\d+$/.test(raw)) {
    return {
      boxesValue: hasPpc ? 1 : 0,
      missingPacks: Number(raw),
      kind: "packs_delta"
    };
  }
  // packs only negative (ex: -3) -> packs_delta
  if (/^-\d+$/.test(raw)) {
    return {
      boxesValue: 1,
      missingPacks: Number(raw),
      kind: "packs_delta"
    };
  }

  // pure integer (ex: 4, 10) -> plain
  if (/^\d+$/.test(raw)) {
    return {
      boxesValue: Number(raw),
      missingPacks: 0,
      kind: "plain"
    };
  }

  throw new Error("箱数/包 格式不支持: " + raw);
}

function ArrivagesDomain_parseTailInput_(input) {
  const raw = String(input || "").trim();
  if (!raw) return { total: 0, display: "", raw: "" };

  const matches = raw.match(/\d+/g);
  if (!matches || !matches.length) throw new Error("尾箱格式不支持: " + raw);

  const nums = matches.map(Number).filter(n => Number.isFinite(n));
  if (!nums.length) throw new Error("尾箱格式不支持: " + raw);

  return {
    total: nums.reduce((sum, n) => sum + n, 0),
    display: nums.join("+"),
    raw: raw
  };
}

function ArrivagesDomain_parsePpcInput_(input) {
  const raw = String(input || "").trim();
  if (!raw) return { primary: 0, values: [], raw: "", display: "" };

  const normalized = raw
    .replace(/[（）]/g, "")
    .replace(/\s+/g, "")
    .replace(/件/gi, "p");

  const parts = normalized.split("+").map(s => String(s || "").trim()).filter(Boolean);
  if (!parts.length) return { primary: 0, values: [], raw: "", display: "" };

  const values = [];
  for (const part of parts) {
    const m = part.match(/^(\d+)(?:[pP])?$/);
    if (!m) throw new Error("每箱件数格式不支持: " + raw);
    const n = Number(m[1]);
    if (!Number.isFinite(n)) throw new Error("每箱件数格式不支持: " + raw);
    values.push(n);
  }

  if (!values.length) return { primary: 0, values: [], raw: "", display: "" };

  return {
    primary: values[0] || 0,
    values: values,
    raw: raw,
    display: values.join("+")
  };
}

function ArrivagesDomain_mergeTailRawIntoNoteSystem_(noteSystem, tailRaw) {
  const note = String(noteSystem || "").trim();
  const raw = String(tailRaw || "").trim();
  if (!raw) return note;

  const cleaned = note
    .replace(/\s*\|\s*TAIL_RAW:[^|]*/gi, "")
    .replace(/^\s*TAIL_RAW:[^|]*\s*\|?\s*/i, "")
    .trim();

  return cleaned ? (cleaned + " | TAIL_RAW:" + raw) : ("TAIL_RAW:" + raw);
}

function ArrivagesDomain_extractTailRawFromNoteSystem_(noteSystem) {
  const s = String(noteSystem || "").trim();
  if (!s) return { cleanNoteSystem: "", tailRaw: "" };

  const m = s.match(/(?:^|\|)\s*TAIL_RAW:([^|]+)/i);
  const tailRaw = m ? String(m[1] || "").trim() : "";
  const cleanNoteSystem = s
    .replace(/\s*\|\s*TAIL_RAW:[^|]*/gi, "")
    .replace(/^\s*TAIL_RAW:[^|]*\s*\|?\s*/i, "")
    .trim()
    .replace(/^\|\s*|\s*\|$/g, "")
    .trim();

  return { cleanNoteSystem, tailRaw };
}

function ArrivagesDomain_mergePpcRawIntoNoteSystem_(noteSystem, ppcRaw) {
  const note = String(noteSystem || "").trim();
  const raw = String(ppcRaw || "").trim();
  if (!raw) return note;

  const cleaned = note
    .replace(/\s*\|\s*PPC_RAW:[^|]*/gi, "")
    .replace(/^\s*PPC_RAW:[^|]*\s*\|?\s*/i, "")
    .trim();

  return cleaned ? (cleaned + " | PPC_RAW:" + raw) : ("PPC_RAW:" + raw);
}

function ArrivagesDomain_extractPpcRawFromNoteSystem_(noteSystem) {
  const s = String(noteSystem || "").trim();
  if (!s) return { cleanNoteSystem: "", ppcRaw: "" };

  const m = s.match(/(?:^|\|)\s*PPC_RAW:([^|]+)/i);
  const ppcRaw = m ? String(m[1] || "").trim() : "";
  const cleanNoteSystem = s
    .replace(/\s*\|\s*PPC_RAW:[^|]*/gi, "")
    .replace(/^\s*PPC_RAW:[^|]*\s*\|?\s*/i, "")
    .trim()
    .replace(/^\|\s*|\s*\|$/g, "")
    .trim();

  return { cleanNoteSystem, ppcRaw };
}

function ArrivagesDomain_mergeBoxPackRawIntoNoteSystem_(noteSystem, boxPackRaw) {
  const note = String(noteSystem || "").trim();
  const raw = String(boxPackRaw || "").trim();
  if (!raw) return note;

  const cleaned = note
    .replace(/\s*\|\s*BOXPACK_RAW:[^|]*/gi, "")
    .replace(/^\s*BOXPACK_RAW:[^|]*\s*\|?\s*/i, "")
    .trim();

  return cleaned ? (cleaned + " | BOXPACK_RAW:" + raw) : ("BOXPACK_RAW:" + raw);
}

function ArrivagesDomain_extractBoxPackRawFromNoteSystem_(noteSystem) {
  const s = String(noteSystem || "").trim();
  if (!s) return { cleanNoteSystem: "", boxPackRaw: "" };

  const m = s.match(/(?:^|\|)\s*BOXPACK_RAW:([^|]+)/i);
  const boxPackRaw = m ? String(m[1] || "").trim() : "";
  const cleanNoteSystem = s
    .replace(/\s*\|\s*BOXPACK_RAW:[^|]*/gi, "")
    .replace(/^\s*BOXPACK_RAW:[^|]*\s*\|?\s*/i, "")
    .trim()
    .replace(/^\|\s*|\s*\|$/g, "")
    .trim();

  return { cleanNoteSystem, boxPackRaw };
}

function ArrivagesDomain_parseQuickInput_(input) {
  const tokens = String(input || "").trim().split(/\s+/).filter(Boolean);
  if (tokens.length !== 3 && tokens.length !== 4) {
    throw new Error("格式错误：`ref ppc cartons` 或 `ref tail ppc cartons`，例如：jy70 120 16 / jy70 56p 120 16");
  }

  const ref = (typeof cleanRef_ === "function" ? cleanRef_(tokens[0]) : String(tokens[0] || "").trim())
    .toUpperCase();
  if (!ref) throw new Error("货号为空");

  const parsePToken = (t) => {
    const s = String(t || "").trim();
    const m = s.match(/^\(?\s*(\d+)\s*[pP箱]?\s*\)?$/);
    if (!m) return NaN;
    const n = Number(m[1]);
    return Number.isFinite(n) ? n : NaN;
  };

  let tail = "";
  let ppc = 0;
  let cartons = 0;

  if (tokens.length === 3) {
    ppc = parsePToken(tokens[1]);
    cartons = parsePToken(tokens[2]);
  } else {
    const t = parsePToken(tokens[1]);
    if (!Number.isFinite(t) || t <= 0) throw new Error("尾箱件数格式错误：例如 56p");
    tail = Math.trunc(t);

    ppc = parsePToken(tokens[2]);
    cartons = parsePToken(tokens[3]);
  }

  if (!Number.isFinite(ppc) || ppc <= 0) throw new Error("每箱件数格式错误：例如 120");
  if (!Number.isFinite(cartons) || cartons <= 0) throw new Error("箱数格式错误：例如 16");

  if (Math.abs(cartons - Math.round(cartons)) > 1e-9) throw new Error("箱数必须是整数（不支持 16.5）");

  return {
    ref,
    tail: tail === "" ? "" : tail,
    ppc: Math.trunc(ppc),
    cartons: Math.round(cartons),
  };
}

/**
 * Transforme des rows DB (A:J) en modèle UI agrégé.
 * DB row format attendu (0-index):
 * 0 A: updatedAt (Date)
 * 2 C: entrepot
 * 3 D: ref
 * 4 E: surplus
 * 5 F: ppc
 * 6 G: cartons
 * 7 H: createdAt (Date)
 * 8 I: noteSystem
 * 9 J: noteUser
 */
function ArrivagesDomain_buildUiModelFromDbRows_(rows) {
  const entrepot  = String(rows[0][2] || "").trim();
  const createdAt = (rows[0][7] instanceof Date) ? rows[0][7] : "";

  let updatedAtMax = null;
  for (const r of rows) {
    const at = (r[0] instanceof Date) ? r[0] : null;
    if (at && (!updatedAtMax || at > updatedAtMax)) updatedAtMax = at;
  }

  const toIntSafe = (v) => (typeof toIntSafe_ === "function" ? toIntSafe_(v) : (Number(v) || 0));
  const toNumberSafe = (v) => {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return Number.isFinite(v) ? v : 0;
    const s = String(v).trim().replace(",", ".");
    if (!s) return 0;
    const n = Number(s);
    return Number.isFinite(n) ? n : 0;
  };
  const cleanRef = (v) => (typeof cleanRef_ === "function" ? cleanRef_(v) : String(v || "").trim());

  const parseMixPartner = (noteSystem) => {
    if (typeof parseMixPartnerFromNoteSystem_ === "function") {
      return parseMixPartnerFromNoteSystem_(noteSystem);
    }
    // fallback minimal: "MIX avec: XXX"
    const s = String(noteSystem || "");
    const m = s.match(/MIX\s+avec:\s*([A-Za-z0-9\-_#]+)/i);
    return m ? m[1] : "";
  };

  const byRef = new Map();
  const seenOrder = [];

  for (const r of rows) {
    const at = (r[0] instanceof Date) ? r[0] : new Date(0);
    const ref = cleanRef(r[3]).toUpperCase();
    if (!ref) continue;

    if (!byRef.has(ref)) {
      byRef.set(ref, {
        ref,
        cartonsSum: 0,
        boxPackRawLatest: { at: new Date(0), v: "" },
        tailRawLatest: { at: new Date(0), v: "" },
        ppcRawLatest: { at: new Date(0), v: "" },
        tailLatest: { at: new Date(0), v: "" },
        ppcLatest:  { at: new Date(0), v: "" },
        noteULatest:{ at: new Date(0), v: "" },
        mixLatest:  { at: new Date(0), partner: "", isCarrier: false },
      });
      seenOrder.push(ref);
    }
    const o = byRef.get(ref);

    const surplus = Math.max(0, toIntSafe(r[4]));
    const ppc     = Math.max(0, toIntSafe(r[5]));
    const cartons = Math.max(0, toNumberSafe(r[6]));
    const noteRaw = String(r[8] || "").trim();
    const tailInfo = ArrivagesDomain_extractTailRawFromNoteSystem_(noteRaw);
    const ppcInfo = ArrivagesDomain_extractPpcRawFromNoteSystem_(tailInfo.cleanNoteSystem);
    const noteInfo = ArrivagesDomain_extractBoxPackRawFromNoteSystem_(ppcInfo.cleanNoteSystem);
    const noteS   = noteInfo.cleanNoteSystem;
    const noteU   = String(r[9] || "").trim();

    // Somme cartons (normal only)
    if (ppc > 0 && cartons > 0) {
      o.cartonsSum += cartons;
      if (at > o.ppcLatest.at) o.ppcLatest = { at, v: ppc };
    }

    // Tail latest (toutes lignes)
    if (surplus > 0 && at > o.tailLatest.at) o.tailLatest = { at, v: surplus };
    if (tailInfo.tailRaw && at > o.tailRawLatest.at) o.tailRawLatest = { at, v: tailInfo.tailRaw };
    if (ppcInfo.ppcRaw && at > o.ppcRawLatest.at) o.ppcRawLatest = { at, v: ppcInfo.ppcRaw };
    if (noteInfo.boxPackRaw && at > o.boxPackRawLatest.at) o.boxPackRawLatest = { at, v: noteInfo.boxPackRaw };

    // Note user latest
    if (noteU && at > o.noteULatest.at) o.noteULatest = { at, v: noteU };

    // MIX detection (strict-ish)
    const partner = parseMixPartner(noteS);
    const isMixRow = !!partner && ppc === 0 && surplus > 0 && (cartons === 1 || cartons === 0);
    if (isMixRow && at > o.mixLatest.at) {
      o.mixLatest = { at, partner, isCarrier: cartons === 1 };
    }
  }

  const out = [];
  for (const ref of seenOrder) {
    const o = byRef.get(ref);
    if (!o) continue;

    const tail = o.tailRawLatest.v || o.tailLatest.v || "";
    const ppc  = o.ppcRawLatest.v || o.ppcLatest.v || "";
    const cartons = o.boxPackRawLatest.v || (o.cartonsSum > 0 ? o.cartonsSum : 1);

    let noteSystem = "";
    if (o.mixLatest.partner) {
      noteSystem = (o.mixLatest.isCarrier ? "MIX_START " : "") + `MIX avec: ${o.mixLatest.partner}`;
      noteSystem = noteSystem.trim();
    }

    const noteUser = o.noteULatest.v || "";

    out.push([o.ref, tail, ppc, cartons, noteSystem, noteUser]);
  }

  out.sort((a, b) => String(a[0]).localeCompare(String(b[0])));

  return {
    entrepot,
    createdAt,
    updatedAtMax: updatedAtMax || "",
    lines: out,
  };
}

/***********************
 * Arrivages_UI.gs — UI controller (onEdit, reset, dispatch)
 ***********************/

function ArrivagesUI_onEdit(e) {
  if (!e || !e.range) return;
  if (typeof isUiGuardOn_ === "function" && isUiGuardOn_()) return;

  const sh = e.range.getSheet();
  if (sh.getName() === SHEET_UI) {
    ArrivagesUI_handleEdit_(e);
  }
}

function ArrivagesUI_handleEdit_(e) {
  if (!e || !e.range) return;
  if (typeof isUiGuardOn_ === "function" && isUiGuardOn_()) return;

  const sh = e.range.getSheet();
  if (sh.getName() !== SHEET_UI) return;

  // Quick input zone G4:G305 (single edit OR paste)
  if (e.range.getColumn() === UI_QUICK_COL &&
      e.range.getRow() >= UI_QUICK_ROW_START &&
      e.range.getLastRow() <= UI_QUICK_ROW_END) {

    const edited = e.range.getValues().flat().map(v => String(v ?? "").trim());
    const lines = edited.filter(s => s);
    if (!lines.length) return;

    try {
      ArrivagesService_quickInsertMany_(lines);
    } catch (err) {
      if (typeof notify_ === "function") notify_("Erreur", String(err && (err.message || err)));
      else throw err;
    } finally {
      const clear = () => e.range.clearContent();
      if (typeof withUiGuard_ === "function") withUiGuard_(clear);
      else clear();
    }
    return;
  }

  // Hors quick zone : on ignore les edits multi-cells
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;

  const a1  = e.range.getA1Notation();
  const val = String(e.value ?? "").trim();

  // B1 selector
  if (a1 === UI_ID_CELL) {
    if (!val) return;
    if (val === LABEL_ADD) { ArrivagesUI_reset_(); return; }

    try {
      ArrivagesService_load_(val);
    } catch (err) {
      if (typeof notify_ === "function") notify_("Erreur", String(err && (err.message || err)));
      else throw err;
    }
  }
}

function ArrivagesUI_reset_() {
  const ss = SpreadsheetApp.getActive();
  const ui = ss.getSheetByName(SHEET_UI);
  if (!ui) throw new Error("Feuille ARRIVAGES introuvable.");

  const doReset = () => {
    ui.getRange(UI_TABLE_RANGE).clearContent();
    ui.getRange(UI_UPDATED_CELL).clearContent();
    ui.getRange(UI_CREATED_CELL).clearContent();
    ui.getRange(UI_ENTREPOT_CELL).clearContent();
    ui.getRange(UI_ID_CELL).setValue("");
  };

  if (typeof withUiGuard_ === "function") withUiGuard_(doReset);
  else doReset();

  try { ss.toast("✅ 新建：表格已重置", "ARRIVAGES", 3); } catch (e) {}
}

/***********************
 * Arrivages.gs — STUB (routing)
 * Garde les anciens noms publics et redirige vers les modules.
 ***********************/

function onEdit(e) {
  // ARRIVAGES UI handler (si tu l’as)
  if (typeof ArrivagesUI_onEdit === "function") ArrivagesUI_onEdit(e);

  // STOCK warehouse logger
  if (typeof StockMoves_onEditWarehouse_ === "function") StockMoves_onEditWarehouse_(e);
}

function handleArrivageEdit_(e) {
  ArrivagesUI_handleEdit_(e);
}

function resetArrivagesUi_() {
  ArrivagesUI_reset_();
}

function loadArrivage_(arrivageId) {
  ArrivagesService_load_(arrivageId);
}

function quickInsertArrivageLine_(input) {
  ArrivagesService_quickInsert_(input);
}

function getRowsByArrivageId_(dbSheet, id) {
  return ArrivagesRepo_getRowsByArrivageId_(dbSheet, id);
} 

/***********************
 * Conversion.gs — SUPPLIER -> UI (H:K -> A:F)
 ***********************/

function menuClearSupplierZone_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_UI);
  if (!sh) {
    try { ss.toast("❌ 找不到 ARRIVAGES", "供应商", 3); } catch (e) {}
    return;
  }
  withUiGuard_(() => sh.getRange(SUP_RANGE).clearContent());
  try { ss.toast("✅ 已清空供应商区", "供应商", 3); } catch (e) {}
}

function menuConvertSupplierToUi_() {
  const ss = SpreadsheetApp.getActive();

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    try { ss.toast("⏳ 系统繁忙，请稍后重试", "供应商", 4); } catch (e) {}
    return;
  }

  try {
    const ui = ss.getSheetByName(SHEET_UI);
    if (!ui) throw new Error("Feuille ARRIVAGES introuvable.");

    const supGrid = ui.getRange(SUP_RANGE).getValues(); // H..K
    const result = parseSupplierToUiRows_(supGrid);

    if (!result.rows.length) {
      try { ss.toast("提示：供应商区没有有效数据", "供应商", 5); } catch (e) {}
      return;
    }

    appendUiRowsSafe_(ui, result.rows);

    if (result.warnings && result.warnings.length) {
      const first = result.warnings[0];
      const more = result.warnings.length - 1;
      try {
        ss.toast(
          "✅ 已追加：" + result.rows.length + " 行\n⚠️ " + first + (more > 0 ? ` (+${more})` : ""),
          "转换完成（有警告）",
          10
        );
      } catch (e) {}
    } else {
      try { ss.toast(`✅ 转换完成：已追加 ${result.rows.length} 行`, "供应商", 6); } catch (e) {}
    }

  } catch (err) {
    const msg = String(err && (err.message || err));
    try { ss.toast("❌ " + msg, "供应商-错误", 10); } catch (e) {}
  } finally {
    lock.releaseLock();
  }
}

function parseSupplierToUiRows_(supGrid) {
  const warnings = [];

  const isBlank_ = (v) => String(v ?? "").trim() === "";
  const normRef_ = (v) => cleanRef_(String(v ?? "")).toUpperCase();
  const asPosInt_ = (v) => {
    const n = toIntSafe_(v);
    return n > 0 ? n : 0;
  };
  const lineNo_ = (i) => SUP_START_ROW + i;

  const agg = new Map();
  const order = [];

  const getAgg_ = (ref) => {
    if (!agg.has(ref)) {
      agg.set(ref, { cartons: 0, tailMax: 0, ppcSet: new Set(), noteSystem: "", userNotes: [] });
      order.push(ref);
    }
    return agg.get(ref);
  };

  const ppcDisplay_ = (set) => {
    const arr = Array.from(set || []).map(Number).filter(n => n > 0).sort((a,b)=>a-b);
    if (!arr.length) return "";
    if (arr.length === 1) return arr[0];
    return arr.join(" - ");
  };

  let lastRef = "";

  for (let i = 0; i < (supGrid || []).length; i++) {
    const r = supGrid[i] || ["", "", "", ""];

    const rawRef = String(r[0] ?? "").trim();
    const hasRef = rawRef !== "";

    const rowHasSomething = !isBlank_(r[0]) || !isBlank_(r[1]) || !isBlank_(r[2]) || !isBlank_(r[3]);
    if (!hasRef && !lastRef && rowHasSomething) {
      warnings.push(`L${lineNo_(i)}: 货号为空且没有上一行货号可继承（该行可能被忽略）`);
      continue;
    }

    const ref = hasRef ? normRef_(rawRef) : lastRef;
    if (!ref) continue;
    if (hasRef) lastRef = ref;

    const cartonsRaw = r[1];
    const qtyRaw     = r[2];
    const noteUser   = String(r[3] || "").trim();

    const cartons = asPosInt_(cartonsRaw);
    const qty     = asPosInt_(qtyRaw);

    if (rowHasSomething && !hasRef && cartons > 0 && qty === 0 && isBlank_(qtyRaw)) {
      warnings.push(`L${lineNo_(i)} (${ref || "?"}): 尾箱可能填错到I列（应填J列）。`);
    }

    // inherited tail: ref empty + cartons>0 + qty>0 => qty is tail
    const isInheritedTail = !hasRef && cartons > 0 && qty > 0;
    if (isInheritedTail) {
      const o = getAgg_(ref);
      o.tailMax = Math.max(o.tailMax, qty);
      if (noteUser) mergeUserNotes_(o.userNotes, noteUser);
      continue;
    }

    const isTailOnly = cartons === 0 && !isBlank_(qtyRaw) && qty > 0;
    const isNormal   = cartons > 0 && qty > 0;

    if (!isNormal && !isTailOnly) continue;

    // MIX detection: [refA cartons=1 qty=X] then [refB tail-only qty=Y]
    if (isNormal && cartons === 1) {
      const next = (i + 1 < supGrid.length) ? (supGrid[i + 1] || ["", "", "", ""]) : null;
      if (next) {
        const rawRefB = String(next[0] ?? "").trim();
        const hasRefB = rawRefB !== "";
        const refB = hasRefB ? normRef_(rawRefB) : "";

        const cartonsB  = asPosInt_(next[1]);
        const qtyB      = asPosInt_(next[2]);
        const noteUserB = String(next[3] || "").trim();

        const isTailOnlyB = cartonsB === 0 && !isBlank_(next[2]) && qtyB > 0;

        if (refB && refB !== ref && isTailOnlyB) {
          const oA = getAgg_(ref);
          const oB = getAgg_(refB);

          oA.tailMax = Math.max(oA.tailMax, qty);
          oB.tailMax = Math.max(oB.tailMax, qtyB);

          oA.noteSystem = mergeSystemNote_(oA.noteSystem, `MIX avec: ${refB} MIX_START`);
          oB.noteSystem = mergeSystemNote_(oB.noteSystem, `MIX avec: ${ref}`);

          if (noteUser)  mergeUserNotes_(oA.userNotes, noteUser);
          if (noteUserB) mergeUserNotes_(oB.userNotes, noteUserB);

          i++; // consume next
          continue;
        }
      }
    }

    const o = getAgg_(ref);
    if (isNormal) {
      o.cartons += cartons;
      o.ppcSet.add(qty);
      if (noteUser) mergeUserNotes_(o.userNotes, noteUser);
    } else if (isTailOnly) {
      o.tailMax = Math.max(o.tailMax, qty);
      if (noteUser) mergeUserNotes_(o.userNotes, noteUser);
    }
  }

  const out = [];
  for (const ref of order) {
    const o = agg.get(ref);
    if (!o) continue;
    out.push([
      ref,
      o.tailMax > 0 ? o.tailMax : "",
      ppcDisplay_(o.ppcSet),
      o.cartons > 0 ? o.cartons : "",
      o.noteSystem || "",
      (o.userNotes && o.userNotes.length) ? o.userNotes.join(" | ") : "",
    ]);
  }

  if (out.length > UI_TABLE_ROWS) {
    warnings.push(`输出超过 ${UI_TABLE_ROWS} 行，将被截断`);
    out.length = UI_TABLE_ROWS;
  }

  return { rows: out, warnings };
}

function _rowHasData_(row) {
  return (row || []).some((c) => String(c ?? "").trim() !== "");
}

function uiHasHoles_(uiSheet) {
  const grid = uiSheet.getRange(UI_TABLE_RANGE).getValues();
  let seenData = false;
  let seenEmptyAfterData = false;

  for (let i = 0; i < grid.length; i++) {
    const has = _rowHasData_(grid[i]);
    if (has) {
      if (seenEmptyAfterData) return true;
      seenData = true;
    } else {
      if (seenData) seenEmptyAfterData = true;
    }
  }
  return false;
}

function uiContiguousFilledCount_(uiSheet) {
  const grid = uiSheet.getRange(UI_TABLE_RANGE).getValues();
  let count = 0;
  for (let i = 0; i < grid.length; i++) {
    if (_rowHasData_(grid[i])) count++;
    else break;
  }
  return count;
}

function appendUiRowsSafe_(uiSheet, uiRows) {
  if (!uiRows || !uiRows.length) throw new Error("Aucune ligne à ajouter.");

  for (const r of uiRows) {
    if (!Array.isArray(r) || r.length !== UI_TABLE_COLS) {
      throw new Error("uiRows invalide: chaque ligne doit avoir 6 colonnes (A..F).");
    }
  }

  if (uiHasHoles_(uiSheet)) {
    throw new Error("A4:F305 contient des trous (lignes vides au milieu). Nettoie/compacte avant d'ajouter.");
  }

  const filled = uiContiguousFilledCount_(uiSheet);
  if (filled >= UI_TABLE_ROWS) throw new Error("A4:F305 est plein. Impossible d'ajouter.");

  const remaining = UI_TABLE_ROWS - filled;
  if (uiRows.length > remaining) {
    throw new Error(`Pas assez de place dans A4:F305: restant=${remaining} lignes, besoin=${uiRows.length}.`);
  }

  const startRow = UI_TABLE_START_ROW + filled;

  withUiGuard_(() => {
    uiSheet.getRange(startRow, 1, uiRows.length, UI_TABLE_COLS).setValues(uiRows);
  });
}

function compactUiTable_(uiSheet) {
  const grid = uiSheet.getRange(UI_TABLE_RANGE).getValues();
  const packed = grid.filter(_rowHasData_);
  while (packed.length < UI_TABLE_ROWS) packed.push(new Array(UI_TABLE_COLS).fill(""));

  withUiGuard_(() => uiSheet.getRange(UI_TABLE_RANGE).setValues(packed));

  try {
    SpreadsheetApp.getActive().toast("A4:F305 compacté (trous supprimés)", "ARRIVAGES", 5);
  } catch (e) {}
}

function mergeUserNotes_(arr, note) {
  if (!note) return;
  const n = String(note).trim();
  if (!n) return;
  if (!Array.isArray(arr)) return;
  if (arr.indexOf(n) === -1) arr.push(n);
}

function mergeSystemNote_(existing, add) {
  const a = normalizeMixStartToken_(String(add || "").trim());
  if (!a) return existing || "";
  const e = normalizeMixStartToken_(String(existing || "").trim());
  if (!e) return a;

  if (e.indexOf(a) !== -1) return e;
  if (a.indexOf(e) !== -1) return a;

  return e + " | " + a;
}

function normalizeMixStartToken_(s) {
  if (!s) return "";
  return s.replace(/mix[\s_-]*start_?/gi, "MIX_START").trim();
}
