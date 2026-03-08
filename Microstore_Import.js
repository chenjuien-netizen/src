/**
 * Import du dernier export Microstore (.xlsx)
 * - Actifs -> MS_IMPORT
 * - Désactivés -> MS_IMPORT_DISABLED
 */
function importMicrostoreLatestExport(promptForOtherImport) {
  msImportLatestExportCore_({
    sourceFolderId: MS_IMPORT_FOLDER_ID,
    targetSheetName: SHEET_MS_IMPORT,
    logType: "IMPORT_MS_ACTIVE",
    importLabel: "actifs",
    noFileMessage: "Aucun fichier .xlsx trouvé dans le dossier des exports actifs.",
    promptForOtherImport: promptForOtherImport !== false,
    otherImportPrompt: "Importer aussi les produits désactivés ?",
    runOtherImport: function() { importMicrostoreLatestDisabledExport(false); }
  });
}

function importMicrostoreLatestDisabledExport(promptForOtherImport) {
  msImportLatestExportCore_({
    sourceFolderId: MS_IMPORT_DISABLED_FOLDER_ID,
    targetSheetName: SHEET_MS_IMPORT_DISABLED,
    logType: "IMPORT_MS_DISABLED",
    importLabel: "désactivés",
    noFileMessage: "Aucun fichier .xlsx trouvé dans le dossier des exports désactivés.",
    promptForOtherImport: promptForOtherImport !== false,
    otherImportPrompt: "Importer aussi les produits actifs ?",
    runOtherImport: function() { importMicrostoreLatestExport(false); }
  });
}

function msImportLatestExportCore_(options) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  ss.toast("Import Microstore " + options.importLabel + " : démarrage…", "Microstore", 5);

  try {
    var shImport = ss.getSheetByName(options.targetSheetName);
    var shLog = ss.getSheetByName(SHEET_LOG_IMPORT_EXPORT);

    if (!shImport) throw new Error("Feuille introuvable: " + options.targetSheetName);
    if (!shLog) throw new Error("Feuille introuvable: " + SHEET_LOG_IMPORT_EXPORT);

    ss.toast("Recherche du dernier export .xlsx…", "Microstore", 5);
    var latestXlsx = msFindLatestXlsxByCreatedTime_(options.sourceFolderId);
    if (!latestXlsx) {
      ss.toast(options.noFileMessage, "Microstore", 6);
      ui.alert(options.noFileMessage);
      return;
    }

    ss.toast("Dernier fichier: " + latestXlsx.name, "Microstore", 5);

    ss.toast("Conversion .xlsx → Google Sheet temporaire…", "Microstore", 6);
    var tempFileId = msConvertXlsxToGoogleSheet_(latestXlsx.id);

    ss.toast("Ouverture du fichier converti…", "Microstore", 6);
    var tempSs = msOpenSpreadsheetWithRetry_(tempFileId, 8, 1500);
    var tempSheets = tempSs.getSheets();
    if (!tempSheets || tempSheets.length === 0) {
      throw new Error("Le fichier converti ne contient aucune feuille.");
    }

    var tempFirst = tempSheets[0];

    ss.toast("Lecture des données…", "Microstore", 6);
    var lastRow = tempFirst.getLastRow();
    var lastCol = tempFirst.getLastColumn();
    if (!lastRow || !lastCol) {
      lastRow = 0;
      lastCol = 0;
    }

    var values = [];
    if (lastRow > 0 && lastCol > 0) {
      values = tempFirst.getRange(1, 1, lastRow, lastCol).getValues();
    }

    ss.toast("Import en texte brut dans " + options.targetSheetName + "…", "Microstore", 6);
    shImport.clearContents();

    var dims = msGetEffectiveDimensions_(values);
    var numRows = dims.rows;
    var numCols = dims.cols;

    if (numRows > 0 && numCols > 0) {
      var trimmed = [];
      for (var r = 0; r < numRows; r++) {
        var row = new Array(numCols);
        for (var c = 0; c < numCols; c++) {
          var v = values[r][c];
          row[c] = (v === null || typeof v === "undefined") ? "" : String(v);
        }
        trimmed.push(row);
      }

      var targetRange = shImport.getRange(1, 1, numRows, numCols);
      targetRange.setNumberFormat("@");
      targetRange.setValues(trimmed);
    }

    ss.toast("Écriture du log…", "Microstore", 5);
    msLogImport_(shLog, {
      date: new Date(),
      type: options.logType,
      fileName: latestXlsx.name,
      fileCreated: latestXlsx.createdTime,
      rows: numRows,
      cols: numCols,
      user: msGetActiveUserEmail_()
    });

    ss.toast("Archivage Drive…", "Microstore", 6);
    msArchiveFileToDatedFolder_(options.sourceFolderId, latestXlsx.id, new Date());

    Drive.Files.remove(tempFileId);

    var successMessage = (numRows === 0 || numCols === 0)
      ? ("Import " + options.importLabel + " OK (fichier vide).")
      : ("Import " + options.importLabel + " OK: " + latestXlsx.name + "\n" + numRows + " lignes × " + numCols + " colonnes");

    ss.toast("Import " + options.importLabel + " terminé.", "Microstore", 8);
    ui.alert(successMessage);

    if (options.promptForOtherImport && typeof options.runOtherImport === "function") {
      var ask = ui.alert("Import Microstore", options.otherImportPrompt, ui.ButtonSet.OK_CANCEL);
      if (ask === ui.Button.OK) {
        options.runOtherImport();
      }
    }
  } catch (err) {
    var msg = (err && err.message) ? err.message : String(err);
    ss.toast("Import échoué: " + msg, "Microstore", 10);
    ui.alert("Erreur import Microstore:\n" + msg);
    Logger.log("Import_MS failed: " + msg);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function msFindLatestXlsxByCreatedTime_(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();

  var best = null;
  while (files.hasNext()) {
    var f = files.next();
    var name = f.getName() || "";
    if (name.toLowerCase().slice(-5) !== ".xlsx") continue;

    var meta = Drive.Files.get(f.getId(), { fields: "id,name,createdTime" });
    var created = meta.createdTime ? new Date(meta.createdTime) : new Date(0);

    if (!best || created.getTime() > best.createdTime.getTime()) {
      best = { id: meta.id, name: meta.name, createdTime: created };
    }
  }
  return best;
}

function msConvertXlsxToGoogleSheet_(xlsxFileId) {
  var xlsxMeta = Drive.Files.get(xlsxFileId, { fields: "name" });
  var resource = {
    title: "TMP_MS_CONVERT__" + xlsxMeta.name + "__" + Utilities.getUuid(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  var converted = Drive.Files.copy(resource, xlsxFileId);
  return converted.id;
}

function msLogImport_(shLog, info) {
  msEnsureLogHeader_(shLog);

  var row = [
    info.date,
    info.type,
    info.fileName,
    info.fileCreated,
    info.rows,
    info.cols,
    info.user
  ];

  var last = shLog.getLastRow();
  shLog.getRange(last + 1, 1, 1, row.length).setValues([row]);
  shLog.getRange(last + 1, 1).setNumberFormat("yyyy-MM-dd HH:mm:ss");
  shLog.getRange(last + 1, 4).setNumberFormat("yyyy-MM-dd HH:mm:ss");
}

function msEnsureLogHeader_(shLog) {
  var expected = ["Date", "Type", "Fichier", "Date du fichier", "Lignes", "Colonnes", "Utilisateur"];
  var needWrite = shLog.getLastRow() === 0;

  if (!needWrite) {
    var current = shLog.getRange(1, 1, 1, expected.length).getValues()[0];
    for (var i = 0; i < expected.length; i++) {
      if (String(current[i] || "").trim() !== expected[i]) {
        needWrite = true;
        break;
      }
    }
  }

  if (!needWrite) return;

  shLog.getRange(1, 1, 1, expected.length).setValues([expected]);
  shLog.getRange(1, 1, 1, expected.length).setFontWeight("bold");
}

function msArchiveFileToDatedFolder_(sourceFolderId, fileId, importDate) {
  var sourceFolder = DriveApp.getFolderById(sourceFolderId);
  var dayName = Utilities.formatDate(importDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

  var targetFolder = null;
  var subfolders = sourceFolder.getFoldersByName(dayName);
  if (subfolders.hasNext()) {
    targetFolder = subfolders.next();
  } else {
    targetFolder = sourceFolder.createFolder(dayName);
  }

  var file = DriveApp.getFileById(fileId);
  targetFolder.addFile(file);
  sourceFolder.removeFile(file);
}

function msGetActiveUserEmail_() {
  try {
    var email = Session.getActiveUser().getEmail();
    return email || "(inconnu)";
  } catch (e) {
    return "(inconnu)";
  }
}
