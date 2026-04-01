function SheetRepository_getInventoryRecords_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(STOCK_APP_SHEET_NAME);

  if (!sheet) {
    throw new Error("Feuille introuvable: " + STOCK_APP_SHEET_NAME);
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 1 || lastColumn < 1) {
    return [];
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const referenceColumn = headers.indexOf(STOCK_APP_REFERENCE_HEADER) + 1;

  if (referenceColumn < 1) {
    throw new Error("Colonne requise introuvable: " + STOCK_APP_REFERENCE_HEADER);
  }

  if (lastRow === 1) {
    return [];
  }

  const values = sheet.getRange(2, referenceColumn, lastRow - 1, 1).getDisplayValues();
  const seen = {};
  const records = [];

  for (let index = 0; index < values.length; index++) {
    const reference = String(values[index][0] || "").trim();
    if (!reference) continue;

    if (seen[reference]) {
      throw new Error("Reference dupliquee detectee: " + reference);
    }

    seen[reference] = true;
    records.push({ reference: reference });
  }

  records.sort(function(left, right) {
    return String(left.reference).localeCompare(String(right.reference));
  });

  return records;
}
