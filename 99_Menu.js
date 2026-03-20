// 99_Menu.gs
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Microstore")
    .addItem("Sync Microstore", "menuSyncMicrostore")
    .addItem("Exporter Microstore", "menuExportMicrostore")
    .addToUi();

  ui.createMenu("eFashion/PFS")
    .addItem("Exporter eFashion", "menuExportEFashion")
    .addItem("Exporter PFS", "menuExportPFS")
    .addItem("Exporter eFashion + PFS", "menuExportEFashionAndPFS")
    .addSeparator()
    .addSubMenu(
      ui.createMenu("Édition eFashion")
        .addItem("eFashion • Importer Template", "menuImportEFashionTemplate")
        .addItem("eFashion • Comparer STOCK ↔ E_IMPORT", "menuCompareStockWithEFashionImport")
    )
    .addSubMenu(
      ui.createMenu("Édition PFS")
        .addItem("PFS • Importer Template", "menuImportPFSTemplate")
        .addItem("PFS • Comparer STOCK ↔ PFS_IMPORT", "menuCompareStockWithPfsImport")
    )
    .addToUi();

  ui.createMenu("Arrivages")
    .addItem("Enregistrer", "arrivagesSaveCurrent")
    .addItem("Supprimer", "deleteArrivage_")
    .addSeparator()
    .addItem("Convert supplier → UI (H:K → A:F)", "menuConvertSupplierToUi_")
    .addItem("Clear supplier zone (H4:K305)", "menuClearSupplierZone_")
    .addToUi();

  ui.createMenu("STOCK")
    .addItem("Valider mouvements (开箱/包)", "StockMoves_validateAll_")
    .addItem("Annuler modifs non validées", "stockResetPending_")
    .addToUi();
}

function menuCompactArrivagesUi_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_UI);
  if (!sh) {
    try { ss.toast("❌ 找不到 ARRIVAGES", "ARRIVAGES", 5); } catch (e) {}
    return;
  }
  compactUiTable_(sh);
}

function menuSyncMicrostore() {
  syncMsImportToStock();
}

function menuExportMicrostore() {
  exportStockToMsExport();
}

function menuExportEFashion() {
  exportStockToEFashion();
}

function menuExportPFS() {
  exportStockToPFS();
}

function menuImportPFSTemplate() {
  importPFSTemplateToSheet();
}

function menuImportEFashionTemplate() {
  importEFashionTemplateToSheet();
}

function menuCompareStockWithEFashionImport() {
  compareStockWithEFashionImport();
}

function menuCompareStockWithPfsImport() {
  compareStockWithPfsImport();
}

function menuExportEFashionAndPFS() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const selection = getSelectedStockSelectionState_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!selection.rows.length) {
      ss.toast("Aucune ligne cochée (选择)", "eFashion/PFS", 5);
      return;
    }

    exportStockToEFashion();
    restoreStockSelectionState_(selection);

    try {
      exportStockToPFS();
    } finally {
      clearStockSelectionState_(selection);
    }
  } finally {
    lock.releaseLock();
  }
}

function getSelectedStockSelectionState_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stock = ss.getSheetByName(SHEET_STOCK);

  if (!stock) {
    throw new Error("Feuille introuvable: " + SHEET_STOCK);
  }

  const lastRow = stock.getLastRow();
  const lastCol = stock.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { sheet: stock, selectionCol: 0, rows: [] };
  }

  const headers = stock.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = headerMap_(headers);
  const selectionCol = map["选择"];
  if (!selectionCol) {
    throw new Error("Colonne '选择' introuvable dans STOCK.");
  }

  const values = stock.getRange(2, selectionCol, lastRow - 1, 1).getValues().flat();
  const rows = [];
  for (let i = 0; i < values.length; i++) {
    if (values[i] === true) rows.push(i + 2);
  }

  return {
    sheet: stock,
    selectionCol: selectionCol,
    rows: rows
  };
}

function restoreStockSelectionState_(selection) {
  setStockSelectionState_(selection, true);
}

function clearStockSelectionState_(selection) {
  setStockSelectionState_(selection, false);
}

function setStockSelectionState_(selection, value) {
  if (!selection || !selection.sheet || !selection.selectionCol || !selection.rows.length) {
    return;
  }

  try {
    const a1 = selection.rows.map(function(row) {
      return selection.sheet.getRange(row, selection.selectionCol).getA1Notation();
    });
    selection.sheet.getRangeList(a1).setValue(value);
  } catch (e) {
    for (let i = 0; i < selection.rows.length; i++) {
      selection.sheet.getRange(selection.rows[i], selection.selectionCol).setValue(value);
    }
  }
}
