// 99_Menu.gs
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Microstore")
    .addItem("Importer dernier export (.xlsx)", "importMicrostoreLatestExport")
    .addItem("Sync MS_IMPORT → STOCK", "syncMsImportToStock")
     .addSeparator()
    .addItem("Exporter STOCK → MS_EXPORT", "exportStockToMsExport")
    .addItem(
  "Exporter MS_EXPORT → Drive (overwrite)",
  "exportMsExportSheetToDriveXlsx"
)
    .addToUi();
  

  ui.createMenu("eFashion")
    .addItem("Exporter STOCK → eFashion", "exportStockToEFashion")
    .addItem("Exporter eFashion → Drive", "exportEFashionToDriveXlsx")
    .addToUi();

  ui.createMenu("PFS")
    .addItem("Générer PFS_EXPORT (sélection STOCK)", "exportStockToPFS")
    .addItem("Exporter PFS_EXPORT → Drive (overwrite)", "exportPFSToDriveXlsx")
    .addToUi();

  ui.createMenu("Arrivages")
    .addItem("Enregistrer", "arrivagesSaveCurrent")
    .addItem("Supprimer", "deleteArrivage_")
    .addSeparator()
    .addItem("Convert supplier → UI (H:K → A:F)", "menuConvertSupplierToUi_")
    .addItem("Clear supplier zone (H4:K200)", "menuClearSupplierZone_")
    .addToUi();

  ui.createMenu("STOCK")
    .addItem("Valider mouvements (出-Sortie/箱)", "StockMoves_validateAll_")
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
