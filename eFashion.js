/****************************************************
 * Export STOCK → e_export (eFashion)
 ****************************************************/

var SHEET_E_EXPORT = "e_export";
var SHEET_E_IMPORT = "E_IMPORT";
var SHEET_E_DIFF = "E_DIFF";
var EFASHION_IMPORT_FOLDER_ID = "1oamwmVGHzns3yh2tPE_WpezCKkNhTqJX";
var EFASHION_IMPORT_PREFIX = "efashion_produits_";

function importEFashionTemplateToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  let tempFileId = "";
  lock.waitLock(30000);

  ss.toast("Import template eFashion : démarrage…", "eFashion Audit", 5);

  try {
    const importFile = efFindLatestImportXlsx_();
    if (!importFile) {
      throw new Error("Aucun export eFashion .xlsx récent trouvé dans le dossier Drive.");
    }

    ss.toast("Conversion XLSX → Google Sheet temporaire…", "eFashion Audit", 5);
    tempFileId = efConvertXlsxToGoogleSheet_(importFile.id);

    ss.toast("Ouverture du fichier converti…", "eFashion Audit", 5);
    const tempSs = efOpenSpreadsheetWithRetry_(tempFileId, 8, 1500);
    const sourceSheet = efSelectBestImportSheet_(tempSs);
    if (!sourceSheet) {
      throw new Error("Le fichier eFashion converti ne contient aucune feuille exploitable.");
    }

    const analysis = efAnalyzeSheetStructure_(sourceSheet);
    const importedSheet = efReplaceSheetFromTemplate_(ss, SHEET_E_IMPORT, sourceSheet);
    SpreadsheetApp.flush();

    const summary = efBuildImportAnalysisSummary_(analysis, {
      fileName: importFile.name,
      targetSheetName: importedSheet.getName()
    });

    ss.toast("Template eFashion importé dans " + SHEET_E_IMPORT, "eFashion Audit", 6);
    notify_("Import eFashion terminé", summary);
    Logger.log(summary);
  } catch (err) {
    const msg = err && err.message ? err.message : String(err);
    ss.toast("Import eFashion échoué: " + msg, "eFashion Audit", 10);
    throw err;
  } finally {
    if (tempFileId) {
      try { Drive.Files.remove(tempFileId); } catch (e) { Logger.log("Cleanup temp eFashion failed: " + e); }
    }
    lock.releaseLock();
  }
}

function compareStockWithEFashionImport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stock = ss.getSheetByName(SHEET_STOCK);
  const imp = ss.getSheetByName(SHEET_E_IMPORT);

  if (!stock) throw new Error("Feuille introuvable: " + SHEET_STOCK);
  if (!imp) throw new Error("Feuille introuvable: " + SHEET_E_IMPORT);

  ss.toast("Comparaison STOCK ↔ E_IMPORT : analyse…", "eFashion Audit", 5);

  const stockLastRow = stock.getLastRow();
  const stockLastCol = stock.getLastColumn();
  if (stockLastRow < 2 || stockLastCol < 1) {
    ss.toast("STOCK vide", "eFashion Audit", 5);
    return;
  }

  const stockHeaders = stock.getRange(1, 1, 1, stockLastCol).getValues()[0];
  const stockMap = headerMap_(stockHeaders);
  ensureHeadersExist_(stockMap, ["货号"], "STOCK");

  const impAnalysis = efAnalyzeSheetStructure_(imp);
  if (!impAnalysis.headerRow || !impAnalysis.usefulWidth) {
    throw new Error("Impossible de détecter les headers de " + SHEET_E_IMPORT + ".");
  }

  const impHeaders = imp.getRange(impAnalysis.headerRow, 1, 1, impAnalysis.usefulWidth).getValues()[0];
  const impDataStartRow = impAnalysis.headerRow + 1;
  const impRowCount = Math.max(0, impAnalysis.usefulRows - impAnalysis.headerRow);
  if (!impRowCount) {
    ss.toast("E_IMPORT ne contient aucune ligne à comparer", "eFashion Audit", 6);
    return;
  }

  const matchCols = efFindColumnsByAliases_(impHeaders, [
    "référence",
    "reference",
    "réf",
    "ref",
    "sku"
  ]);
  const typeCol = efFindColumnByAliases_(impHeaders, ["type", "type produit", "type vente", "vendu par"]);
  const stockRefCol = stockMap["货号"] || 0;

  if (!matchCols.length) {
    throw new Error("Aucune colonne de matching claire trouvée dans " + SHEET_E_IMPORT + " (référence/ref/sku).");
  }

  const stockValues = stock.getRange(2, 1, stockLastRow - 1, stockLastCol).getValues();
  const impValues = imp.getRange(impDataStartRow, 1, impRowCount, impAnalysis.usefulWidth).getValues();

  const stockByRef = {};
  const duplicateStockRefs = [];
  for (let i = 0; i < stockValues.length; i++) {
    const ref = efNormalizeRef_(stockValues[i][stockRefCol - 1]);
    if (!ref) continue;
    if (stockByRef[ref]) {
      duplicateStockRefs.push(ref);
      continue;
    }
    stockByRef[ref] = stockValues[i];
  }

  const stockRefs = Object.keys(stockByRef);
  if (!stockRefs.length) {
    ss.toast("Aucune ref exploitable trouvée dans STOCK", "eFashion Audit", 6);
    return;
  }

  const impRowsByRef = {};
  const impRowsByBaseRef = {};
  const seenImportRefs = {};

  for (let i = 0; i < impValues.length; i++) {
    const row = impValues[i];
    const baseRef = efExtractBaseRefFromRow_(row, matchCols);
    if (baseRef) {
      seenImportRefs[baseRef] = true;
      if (!impRowsByBaseRef[baseRef]) impRowsByBaseRef[baseRef] = [];
      impRowsByBaseRef[baseRef].push(row);
    }

    const ref = efFindMatchingStockRefForImportRow_(row, stockRefs, matchCols);
    if (!ref) continue;
    if (!impRowsByRef[ref]) impRowsByRef[ref] = [];
    impRowsByRef[ref].push(row);
  }

  const compareDefs = [
    {
      key: "prix",
      label: "Prix",
      stockCol: stockMap["prix@"] || 0,
      importCol: efFindColumnByAliases_(impHeaders, ["prix de vente ht", "prix ht", "prix de vente"]),
      type: "number",
      stockValue: function(row) { return row[stockMap["prix@"] - 1]; }
    },
    {
      key: "promo",
      label: "Promo",
      stockCol: stockMap["promo@"] || 0,
      importCol: efFindColumnByAliases_(impHeaders, ["prix promo / prix réduit", "prix promo/prix réduit", "prix promo / prix reduit", "prix promo", "prix réduit", "prix reduit"]),
      type: "number",
      stockValue: function(row) { return row[stockMap["promo@"] - 1]; }
    },
    {
      key: "stock",
      label: "Stock",
      stockCol: stockMap["stock"] || 0,
      importCol: efFindColumnByAliases_(impHeaders, ["stock", "quantité", "quantite"]),
      type: "number",
      stockValue: function(row) { return row[stockMap["stock"] - 1]; }
    },
    {
      key: "poids",
      label: "Poids",
      stockCol: stockMap["poids (en gramme)"] || 0,
      importCol: efFindColumnByAliases_(impHeaders, ["poids", "poids kg", "poids_kg"]),
      type: "number",
      stockValue: function(row) { return formatWeightKgEFashionText_(row[stockMap["poids (en gramme)"] - 1]); }
    },
    {
      key: "active",
      label: "Active",
      stockCol: stockMap["ms_statut"] || 0,
      importCol: efFindColumnByAliases_(impHeaders, ["actif", "active", "statut", "status"]),
      type: "status",
      stockValue: function(row) { return efMapActiveFromMsStatus_(row[stockMap["ms_statut"] - 1]); }
    }
  ];

  const availableDefs = compareDefs.filter(function(def) {
    return def.stockCol && def.importCol;
  });
  const missingDefs = compareDefs
    .filter(function(def) { return !def.stockCol || !def.importCol; })
    .map(function(def) { return def.label; });

  if (!availableDefs.length) {
    throw new Error("Aucune colonne comparable claire trouvée entre STOCK et " + SHEET_E_IMPORT + ".");
  }

  const outRows = [];
  const matchedRefs = [];
  const absentImportRefs = [];
  const absentStockRefs = [];
  const nonComparableWarnings = [];

  for (let i = 0; i < stockRefs.length; i++) {
    const ref = stockRefs[i];
    const stockRow = stockByRef[ref];
    const targetRows = impRowsByRef[ref];

    if (!targetRows || !targetRows.length) {
      absentImportRefs.push(ref);
      outRows.push(efBuildAuditOutputRow_({
        ref: ref,
        typeValue: "",
        sku: "",
        fields: efBuildEmptyAuditFields_(availableDefs, stockRow),
        status: "ABSENT_EFASHION",
        comment: "Ref absente de eFashion"
      }));
      continue;
    }

    matchedRefs.push(ref);

    for (let r = 0; r < targetRows.length; r++) {
      const snapshot = efBuildAuditSnapshot_(ref, stockRow, targetRows[r], availableDefs, {
        matchCols: matchCols,
        typeCol: typeCol
      });

      if (snapshot.nonComparable.length) {
        for (let w = 0; w < snapshot.nonComparable.length; w++) {
          nonComparableWarnings.push(ref + " / " + snapshot.nonComparable[w].label + ": " + snapshot.nonComparable[w].reason);
        }
      }

      outRows.push(efBuildAuditOutputRow_({
        ref: ref,
        typeValue: snapshot.typeValue,
        sku: snapshot.debugSku,
        fields: snapshot.fields,
        status: snapshot.diffLabels.length || snapshot.nonComparable.length ? "DIFF" : "OK",
        comment: efBuildAuditComment_(snapshot)
      }));
    }
  }

  Object.keys(seenImportRefs).forEach(function(ref) {
    if (stockByRef[ref]) return;
    absentStockRefs.push(ref);
    const rows = impRowsByBaseRef[ref] || [];
    if (!rows.length) return;

    const snapshot = efBuildAuditSnapshot_(ref, null, rows[0], availableDefs, {
      matchCols: matchCols,
      typeCol: typeCol
    });
    outRows.push(efBuildAuditOutputRow_({
      ref: ref,
      typeValue: snapshot.typeValue,
      sku: snapshot.debugSku,
      fields: snapshot.fields,
      status: "ABSENT_STOCK",
      comment: "Ref absente de STOCK"
    }));
  });

  const diffSheet = efRecreateSheet_(ss, SHEET_E_DIFF);
  const header = [[
    "Ref",
    "Type",
    "Prix STOCK",
    "Prix eFashion",
    "Promo STOCK",
    "Promo eFashion",
    "Stock STOCK",
    "Stock eFashion",
    "Poids STOCK",
    "Poids eFashion",
    "Active STOCK",
    "Active eFashion",
    "Statut",
    "Commentaire",
    "SKU"
  ]];
  diffSheet.getRange(1, 1, 1, header[0].length).setValues(header);
  diffSheet.getRange(1, 1, 1, header[0].length).setFontWeight("bold");
  diffSheet.setFrozenRows(1);

  if (outRows.length) {
    diffSheet.getRange(2, 1, outRows.length, header[0].length).setValues(outRows);
  }
  for (let c = 1; c <= header[0].length; c++) {
    diffSheet.autoResizeColumn(c);
  }

  const summary = [
    "Comparaison STOCK ↔ " + SHEET_E_IMPORT,
    "",
    "Fichier matché via colonnes: " + matchCols.map(function(col) { return impHeaders[col - 1]; }).join(", "),
    "Refs STOCK analysées : " + stockRefs.length,
    "Refs matchées eFashion : " + matchedRefs.length,
    "Refs absentes de eFashion : " + absentImportRefs.length,
    "Lignes avec écarts : " + outRows.filter(function(row) { return row[12] !== 'OK'; }).length
  ];

  if (missingDefs.length) summary.push("Colonnes non comparées: " + missingDefs.join(", "));
  if (duplicateStockRefs.length) summary.push("Refs dupliquées dans STOCK ignorées: " + efUniqueList_(duplicateStockRefs).slice(0, 10).join(", "));
  if (absentImportRefs.length) summary.push("Refs absentes de eFashion (aperçu): " + absentImportRefs.slice(0, 15).join(", ") + (absentImportRefs.length > 15 ? " …" : ""));
  if (absentStockRefs.length) summary.push("Bonus - refs eFashion absentes de STOCK (aperçu): " + absentStockRefs.slice(0, 15).join(", ") + (absentStockRefs.length > 15 ? " …" : ""));
  if (nonComparableWarnings.length) summary.push("Valeurs non comparables (aperçu): " + efUniqueList_(nonComparableWarnings).slice(0, 8).join(" | ") + (nonComparableWarnings.length > 8 ? " …" : ""));

  const message = summary.join("\n");

  if (absentImportRefs.length) {
    Logger.log("Refs STOCK absentes de eFashion (" + absentImportRefs.length + "):\n" + absentImportRefs.join("\n"));
  }
  if (absentStockRefs.length) {
    Logger.log("Refs eFashion absentes de STOCK (" + absentStockRefs.length + "):\n" + absentStockRefs.join("\n"));
  }
  if (nonComparableWarnings.length) {
    Logger.log("Valeurs eFashion non comparables (" + nonComparableWarnings.length + "):\n" + efUniqueList_(nonComparableWarnings).join("\n"));
  }
  Logger.log(message);

  ss.toast("Comparaison eFashion terminée", "eFashion Audit", 6);
  notify_("Comparaison eFashion terminée", message);
  return {
    analyzedRefs: stockRefs.length,
    matchedRefs: matchedRefs.length,
    absentImportRefs: absentImportRefs,
    absentStockRefs: absentStockRefs,
    diffCount: outRows.filter(function(row) { return row[12] !== 'OK'; }).length
  };
}

function efFindLatestImportXlsx_() {
  const folder = DriveApp.getFolderById(EFASHION_IMPORT_FOLDER_ID);
  const files = folder.getFiles();
  let best = null;

  while (files.hasNext()) {
    const file = files.next();
    const name = String(file.getName() || "");
    const lower = name.toLowerCase();
    if (lower.indexOf(EFASHION_IMPORT_PREFIX) !== 0) continue;
    if (lower.slice(-5) !== ".xlsx") continue;

    const meta = Drive.Files.get(file.getId(), { fields: "id,name,modifiedTime,createdTime" });
    const stamp = meta.modifiedTime || meta.createdTime || "";
    const when = stamp ? new Date(stamp) : new Date(0);
    if (!best || when.getTime() > best.when.getTime()) {
      best = {
        id: meta.id,
        name: meta.name,
        when: when
      };
    }
  }

  return best;
}

function efConvertXlsxToGoogleSheet_(xlsxFileId) {
  const meta = Drive.Files.get(xlsxFileId, { fields: "name" });
  const resource = {
    name: "TMP_EFASHION_IMPORT__" + meta.name + "__" + Utilities.getUuid(),
    title: "TMP_EFASHION_IMPORT__" + meta.name + "__" + Utilities.getUuid(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  const converted = Drive.Files.copy(resource, xlsxFileId);
  return converted.id;
}

function efOpenSpreadsheetWithRetry_(fileId, attempts, sleepMs) {
  if (typeof msOpenSpreadsheetWithRetry_ === "function") {
    return msOpenSpreadsheetWithRetry_(fileId, attempts, sleepMs);
  }

  attempts = attempts || 6;
  sleepMs = sleepMs || 1000;
  let lastErr = null;

  for (let i = 0; i < attempts; i++) {
    try {
      return SpreadsheetApp.openById(fileId);
    } catch (e) {
      lastErr = e;
      Utilities.sleep(sleepMs);
    }
  }
  throw lastErr || new Error("Impossible d'ouvrir le fichier eFashion converti.");
}

function efAnalyzeSheetStructure_(sheet) {
  if (typeof pfsAnalyzeSheetStructure_ === "function") {
    return pfsAnalyzeSheetStructure_(sheet);
  }
  throw new Error("Helper d'analyse de feuille indisponible.");
}

function efSelectBestImportSheet_(spreadsheet) {
  const sheets = spreadsheet.getSheets();
  if (!sheets || !sheets.length) return null;

  const tokens = ["référence", "reference", "prix", "stock", "poids", "actif", "statut", "collection", "vendu"];
  let bestSheet = null;
  let bestScore = -1;

  for (let i = 0; i < sheets.length; i++) {
    const analysis = efAnalyzeSheetStructure_(sheets[i]);
    const headerPreview = (analysis.headerValues || []).map(function(v) {
      return String(v || "").toLowerCase();
    });
    let hits = 0;
    for (let t = 0; t < tokens.length; t++) {
      for (let h = 0; h < headerPreview.length; h++) {
        if (headerPreview[h].indexOf(tokens[t]) !== -1) {
          hits++;
          break;
        }
      }
    }
    const score = hits * 100 + analysis.usefulRows * 2 + analysis.usefulWidth;
    if (score > bestScore) {
      bestScore = score;
      bestSheet = sheets[i];
    }
  }

  return bestSheet || sheets[0];
}

function efBuildImportAnalysisSummary_(analysis, context) {
  const preview = (analysis.headerValues || [])
    .slice(0, Math.min((analysis.headerValues || []).length, 12))
    .map(function(v) { return String(v || "").trim(); })
    .filter(Boolean)
    .join(" | ");

  return [
    "Fichier importé: " + (context && context.fileName ? context.fileName : ""),
    "Feuille source retenue: " + analysis.sheetName,
    "Feuille cible: " + (context && context.targetSheetName ? context.targetSheetName : SHEET_E_IMPORT),
    "Dimensions utiles: " + analysis.usefulRows + " lignes × " + analysis.usefulWidth + " colonnes",
    "Header row détecté: " + analysis.headerRow,
    "Headers (aperçu): " + (preview || "(aucun header lisible)")
  ].join("\n");
}

function efReplaceSheetFromTemplate_(spreadsheet, targetSheetName, sourceSheet) {
  if (typeof pfsReplaceSheetFromTemplate_ === "function") {
    return pfsReplaceSheetFromTemplate_(spreadsheet, targetSheetName, sourceSheet);
  }
  throw new Error("Helper de remplacement de feuille indisponible.");
}

function efFindColumnByAliases_(headers, aliases) {
  const findCol = buildEFashionColumnIndex_(headers || []);
  for (let i = 0; i < aliases.length; i++) {
    const col = findCol(aliases[i]);
    if (col) return col;
  }
  return 0;
}

function efFindColumnsByAliases_(headers, aliases) {
  const out = [];
  const seen = {};
  for (let i = 0; i < aliases.length; i++) {
    const col = efFindColumnByAliases_(headers, [aliases[i]]);
    if (col && !seen[col]) {
      seen[col] = true;
      out.push(col);
    }
  }
  return out;
}

function efNormalizeRef_(value) {
  return (typeof cleanRef_ === "function" ? cleanRef_(value) : String(value || "").trim()).toUpperCase();
}

function efCellMatchesStockRef_(cellValue, ref) {
  const cell = efNormalizeRef_(cellValue);
  const stockRef = efNormalizeRef_(ref);
  if (!cell || !stockRef) return false;
  return cell === stockRef || cell.indexOf(stockRef + "_") === 0;
}

function efExtractBaseRefFromRow_(row, cols) {
  const list = Array.isArray(cols) ? cols : [];
  for (let i = 0; i < list.length; i++) {
    const raw = efNormalizeRef_(row[list[i] - 1]);
    if (!raw) continue;
    const pos = raw.indexOf("_");
    return pos >= 0 ? raw.slice(0, pos) : raw;
  }
  return "";
}

function efFindMatchingStockRefForImportRow_(row, stockRefs, cols) {
  const refs = Array.isArray(stockRefs) ? stockRefs : [];
  const list = Array.isArray(cols) ? cols : [];
  for (let c = 0; c < list.length; c++) {
    const value = row[list[c] - 1];
    for (let i = 0; i < refs.length; i++) {
      if (efCellMatchesStockRef_(value, refs[i])) return refs[i];
    }
  }
  return "";
}

function efResolveComparableValue_(value) {
  if (value && typeof value === "object" && Object.prototype.hasOwnProperty.call(value, "write")) {
    return {
      write: value.write === true,
      value: String(value.value || "").trim(),
      reason: value.reason ? String(value.reason) : ""
    };
  }
  return {
    write: true,
    value: String(value === null || typeof value === "undefined" ? "" : value).trim(),
    reason: ""
  };
}

function efPrepareComparableValue_(value, compareType) {
  const resolved = efResolveComparableValue_(value);
  if (!resolved.write) {
    return {
      comparable: false,
      value: "",
      display: resolved.value,
      reason: resolved.reason || "Valeur non comparable"
    };
  }

  const raw = resolved.value;
  if (compareType === "number") {
    if (raw === "") return { comparable: true, value: "", display: "" };
    const num = efToComparableNumber_(raw);
    if (num === null) {
      return {
        comparable: false,
        value: raw,
        display: raw,
        reason: "Valeur numérique invalide (" + raw + ")"
      };
    }
    return {
      comparable: true,
      value: num,
      display: raw
    };
  }

  if (compareType === "status") {
    if (raw === "") return { comparable: true, value: "", display: "" };
    const mapped = efNormalizeActiveValue_(raw);
    if (!mapped) {
      return {
        comparable: false,
        value: raw,
        display: raw,
        reason: "Valeur actif/statut invalide (" + raw + ")"
      };
    }
    return {
      comparable: true,
      value: mapped,
      display: raw
    };
  }

  return {
    comparable: true,
    value: raw,
    display: raw
  };
}

function efMapActiveFromMsStatus_(status) {
  const raw = String(status || "").trim().toUpperCase();
  if (!raw) return { write: true, value: "" };
  if (raw === "MS") return { write: true, value: "Oui" };
  if (raw === "MS_DISABLED") return { write: true, value: "Non" };
  return {
    write: false,
    value: "",
    reason: "MS_STATUT non géré (" + raw + ")"
  };
}

function efNormalizeActiveValue_(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (!raw) return "";
  if (["oui", "yes", "true", "1", "actif", "active", "enabled"].indexOf(raw) !== -1) return "oui";
  if (["non", "no", "false", "0", "inactif", "inactive", "disabled"].indexOf(raw) !== -1) return "non";
  return "";
}

function efComparableValuesEqual_(left, right, compareType) {
  if (compareType === "number") {
    if (left.value === "" && right.value === "") return true;
    if (typeof left.value !== "number" || typeof right.value !== "number") return false;
    return Math.abs(left.value - right.value) < 1e-9;
  }
  return String(left.value || "") === String(right.value || "");
}

function efToComparableNumber_(value) {
  const raw = String(value || "").trim();
  if (!raw) return null;
  const normalized = raw.replace(/\s+/g, "").replace(",", ".");
  if (!/^-?\d+(?:\.\d+)?$/.test(normalized)) return null;
  const num = Number(normalized);
  return Number.isFinite(num) ? num : null;
}

function efBuildAuditSnapshot_(ref, stockRow, importRow, defs, options) {
  const compareDefs = Array.isArray(defs) ? defs : [];
  const opts = options || {};
  const fields = {};
  const diffLabels = [];
  const nonComparable = [];

  for (let i = 0; i < compareDefs.length; i++) {
    const def = compareDefs[i];
    const stockPrepared = stockRow
      ? efPrepareComparableValue_(def.stockValue(stockRow), def.type)
      : { comparable: true, value: "", display: "" };
    const importPrepared = efPrepareComparableValue_(importRow ? importRow[def.importCol - 1] : "", def.type);

    fields[def.key] = {
      stock: stockPrepared,
      imp: importPrepared
    };

    if (!stockPrepared.comparable || !importPrepared.comparable) {
      nonComparable.push({
        label: def.label,
        reason: stockPrepared.reason || importPrepared.reason || "Valeur non comparable"
      });
      continue;
    }

    if (!efComparableValuesEqual_(stockPrepared, importPrepared, def.type)) {
      diffLabels.push(def.label);
    }
  }

  return {
    ref: ref,
    typeValue: importRow && opts.typeCol ? String(importRow[opts.typeCol - 1] || "").trim() : "",
    debugSku: importRow ? efBuildDebugSkuFromRow_(importRow, opts.matchCols) : "",
    fields: fields,
    diffLabels: diffLabels,
    nonComparable: nonComparable
  };
}

function efBuildDebugSkuFromRow_(row, cols) {
  const list = Array.isArray(cols) ? cols : [];
  for (let i = 0; i < list.length; i++) {
    const value = String(row[list[i] - 1] || "").trim();
    if (value) return value;
  }
  return "";
}

function efBuildAuditComment_(snapshot) {
  const parts = [];
  if (snapshot && snapshot.diffLabels && snapshot.diffLabels.length) {
    parts.push("Écarts STOCK/eFashion: " + snapshot.diffLabels.join(", "));
  }
  if (snapshot && snapshot.nonComparable && snapshot.nonComparable.length) {
    parts.push("Non comparable: " + snapshot.nonComparable.map(function(item) { return item.label; }).join(", "));
  }
  return parts.join(" | ");
}

function efBuildAuditOutputRow_(info) {
  const data = info || {};
  const fields = data.fields || {};
  const getDisplay = function(key, side) {
    const cell = fields[key] && fields[key][side];
    return cell ? String(cell.display || "").trim() : "";
  };

  return [
    data.ref || "",
    data.typeValue || "",
    getDisplay("prix", "stock"),
    getDisplay("prix", "imp"),
    getDisplay("promo", "stock"),
    getDisplay("promo", "imp"),
    getDisplay("stock", "stock"),
    getDisplay("stock", "imp"),
    getDisplay("poids", "stock"),
    getDisplay("poids", "imp"),
    getDisplay("active", "stock"),
    getDisplay("active", "imp"),
    data.status || "",
    data.comment || "",
    data.sku || ""
  ];
}

function efBuildEmptyAuditFields_(defs, stockRow) {
  const compareDefs = Array.isArray(defs) ? defs : [];
  const fields = {};
  for (let i = 0; i < compareDefs.length; i++) {
    const def = compareDefs[i];
    fields[def.key] = {
      stock: stockRow ? efPrepareComparableValue_(def.stockValue(stockRow), def.type) : { comparable: true, value: "", display: "" },
      imp: { comparable: true, value: "", display: "" }
    };
  }
  return fields;
}

function efRecreateSheet_(spreadsheet, sheetName) {
  const existing = spreadsheet.getSheetByName(sheetName);
  let index = spreadsheet.getSheets().length + 1;

  if (existing) {
    index = existing.getIndex();
    spreadsheet.deleteSheet(existing);
  }

  const safeIndex = Math.max(1, Math.min(index, spreadsheet.getSheets().length + 1));
  return spreadsheet.insertSheet(sheetName, safeIndex);
}

function efUniqueList_(values) {
  const out = [];
  const seen = {};
  for (let i = 0; i < values.length; i++) {
    const key = String(values[i] || "");
    if (!key || seen[key]) continue;
    seen[key] = true;
    out.push(key);
  }
  return out;
}

function exportStockToEFashion() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stock = ss.getSheetByName(SHEET_STOCK);
  const exp = ss.getSheetByName(SHEET_E_EXPORT);

  if (!stock) throw new Error("Feuille STOCK introuvable");
  if (!exp) throw new Error("Feuille e_export introuvable");

  // --- Map STOCK headers
  const stockHeaders = stock.getRange(1, 1, 1, stock.getLastColumn()).getValues()[0];
  const stockMap = headerMap_(stockHeaders);

  const need = [
    "选择",
    "货号",
    "Prix@",
    "Date de création",
    "Contenu colis",
    "Poids (en gramme)",
    "Pays d'origine",
    "Composition matérielle",
    "Couleurs",
    "Catégorie",
    "Colisage"
  ];

  ensureHeadersExist_(stockMap, need, "STOCK");

  const lastRow = stock.getLastRow();
  if (lastRow < 2) {
    ss.toast("STOCK vide", "eFashion", 4);
    return;
  }

  const n = lastRow - 1;

  // Read only what we need
  const cSel = stockMap["选择"]; 
  const cRef = stockMap["货号"];
  const cPrix = stockMap["prix@"];
  const cDate = stockMap["date de création"];
  const cContenu = stockMap["contenu colis"];
  const cPoids = stockMap["poids (en gramme)"];
  const cPays = stockMap["pays d'origine"];
  const cCompo = stockMap["composition matérielle"];
  const cCouleursStock = stockMap["couleurs"];
  const cCategorieStock = stockMap["catégorie"];
  const cColisage = stockMap["colisage"];

  const refVals = stock.getRange(2, cRef, n, 1).getValues().flat();
  const prixVals = stock.getRange(2, cPrix, n, 1).getValues().flat();
  const dateVals = stock.getRange(2, cDate, n, 1).getValues().flat();
  const contenuVals = stock.getRange(2, cContenu, n, 1).getValues().flat();
  const poidsVals = stock.getRange(2, cPoids, n, 1).getValues().flat();
  const paysVals = stock.getRange(2, cPays, n, 1).getValues().flat();
  const compoVals = stock.getRange(2, cCompo, n, 1).getValues().flat();
  const selVals = stock.getRange(2, cSel, n, 1).getValues().flat();
  const couleursStockVals = stock.getRange(2, cCouleursStock, n, 1).getValues().flat();
  const categorieStockVals = stock.getRange(2, cCategorieStock, n, 1).getValues().flat();
  const colisageVals = stock.getRange(2, cColisage, n, 1).getValues().flat();

  let rows = [];
  const bad = []; // {row, ref, reason}

  for (let i = 0; i < n; i++) {

    if (selVals[i] !== true) continue;

    const ref = String(refVals[i] || "").trim();
    if (!ref) continue;

    const prix = formatPriceEFashionText_(prixVals[i]);
    const dt = normalizeDateForSortEf_(dateVals[i]);
    const contenuColis = String(contenuVals[i] ?? "").trim();
    const colisage = parseColisageEFashion_(colisageVals[i]);
    const tailles = colisage.ok
      ? extractTaillesFromContenuColisEFashion_(contenuColis, colisage.value)
      : "";
    const poids = String(poidsVals[i] ?? "").trim();
    const pays = String(paysVals[i] ?? "").trim();

    // Composition: prefer shared normalizer from PFS.js if present
    let compo = "";
    if (typeof normalizeCompositionPFS_ === "function") {
      compo = normalizeCompositionPFS_(compoVals[i]);
    } else {
      compo = normalizeCompositionEFashionFallback_(compoVals[i]);
    }

    const couleursStock = String(couleursStockVals[i] ?? "").trim();
    const catStock = String(categorieStockVals[i] ?? "").trim();
    const efCat = mapStockCategoryToEFashion_(catStock);
    const taillesInfo = tailles ? parseEFashionTaillesStructure_(tailles) : { ok: false, reason: "Tailles introuvables ou invalides dans Contenu colis" };

    const sheetRow = 2 + i;
    if (!colisage.ok) {
      bad.push({ row: sheetRow, ref: ref || "(vide)", reason: colisage.reason });
      continue;
    }
    if (!tailles) {
      bad.push({ row: sheetRow, ref: ref || "(vide)", reason: "Tailles introuvables ou invalides dans Contenu colis" });
      continue;
    }
    if (!taillesInfo.ok) {
      bad.push({ row: sheetRow, ref: ref || "(vide)", reason: "Structure tailles invalide: " + taillesInfo.reason });
      continue;
    }
    const colorPack = validateEFashionColorPack_(couleursStock, colisage.value, taillesInfo.sizes);
    if (!colorPack.ok) {
      bad.push({ row: sheetRow, ref: ref || "(vide)", reason: colorPack.reason });
      continue;
    }
    const couleursEf = formatEFashionMixedColorsFromStock_(colorPack);
    if (!couleursEf) {
      bad.push({ row: sheetRow, ref: ref || "(vide)", reason: "Couleurs incompatibles avec la structure tailles" });
      continue;
    }
    rows.push({
      ref: ref,
      prix: prix,
      empty: dt.empty,
      ts: dt.ts,
      tailles: tailles,
      poids: poids,
      pays: pays,
      compo: compo,
      row: sheetRow,
      couleursEf: couleursEf,
      sousCategorie: efCat.sousCategorie,
      sousSousCategorie: efCat.sousSousCategorie,
      colisage: colisage.value,
    });

  }
  if (bad.length) {
    const details = bad
      .slice(0, 15)
      .map(function (x) {
        return [
          "• Ligne " + x.row,
          "Ref : " + x.ref,
          "Motif : " + appendValidationHint_(x.reason)
        ].join("\n");
      })
      .join("\n\n");

    const more = bad.length > 15
      ? "\n\n(" + (bad.length - 15) + " autres lignes avec erreur...)"
      : "";

    const message = [
      "eFashion export",
      "",
      "Erreurs détectées dans STOCK :",
      "",
      details + more
    ].join("\n");

    SpreadsheetApp.getUi().alert(message);
    throw new Error(message);
  }

  if (rows.length === 0) {
    const hinfo0 = readEFashionHeaders_(exp);
    const lr0 = exp.getLastRow();
    if (lr0 > 1) exp.getRange(2, 1, lr0 - 1, hinfo0.width).clearContent();
    ss.toast("Aucune ligne cochée (选择)", "eFashion", 4);
    return;
  }

  // Sort (keep your existing logic): empty dates first, then newest first
  rows.sort(function (a, b) {
    if (a.empty && !b.empty) return -1;
    if (!a.empty && b.empty) return 1;
    return b.ts - a.ts;
  });

  // --- Read eFashion headers robustly
  const hinfo = readEFashionHeaders_(exp);
  const headers = hinfo.headers;
  const col = buildEFashionColumnIndex_(headers);

  // Header-based lookup (preferred)
  let cMarque = col("marque");
  let cReference = col("référence");
  let cCategorie = col("catégorie");
  let cSousCategorie = col("sous catégorie");
  let cSousSousCategorie = col("sous-sous catégorie");
  let cProvenances = col("provenances");
  let cVenduPar = col("vendu par");
  let cQteMin = col("qte min");
  let cCollection = col("collection");
  let cTailles = col("tailles");
  let cCouleurs = col("couleurs");
  let cPrixHT = col("prix ht");
  let cPoidsKG = col("poids kg");
  let cCompositions = col("compositions");

  // Fallback fixed positions A..U (1-based)
  if (!cMarque) cMarque = 1;        // A
  if (!cReference) cReference = 2;  // B
  if (!cCategorie) cCategorie = 3;  // C
  if (!cSousCategorie) cSousCategorie = 4;     // D
  if (!cSousSousCategorie) cSousSousCategorie = 5; // E
  if (!cProvenances) cProvenances = 6; // F
  if (!cVenduPar) cVenduPar = 7;    // G
  if (!cQteMin) cQteMin = 8;              // H
  if (!cCollection) cCollection = 9; // I
  if (!cTailles) cTailles = 10;     // J
  if (!cCouleurs) cCouleurs = 11;   // K
  if (!cPrixHT) cPrixHT = 12;       // L
  if (!cPoidsKG) cPoidsKG = 14;     // N
  if (!cCompositions) cCompositions = 15; // O

  const out = [];

  for (let it of rows) {
    const line = new Array(hinfo.width).fill("");

    // Fixed values
    line[cMarque - 1] = "J&S FASHION";
    line[cCategorie - 1] = "Femme";
    line[cSousCategorie - 1] = it.sousCategorie;
    line[cSousSousCategorie - 1] = it.sousSousCategorie;

    // User wants mixed-color selling with no minimum quantity for now
    line[cVenduPar - 1] = "melangees";
    line[cQteMin - 1] = "";

    // Extra safety: also write fixed-position H (Qte min) in case header matching is off
    if (hinfo.width >= 8) line[7] = "";

    line[cCollection - 1] = "Printemps / Été";

    // From STOCK
    line[cReference - 1] = it.ref;
    line[cProvenances - 1] = it.pays;

    line[cTailles - 1] = it.tailles;
    line[cCouleurs - 1] = it.couleursEf;

    // Extra safety: fixed-position J/K as well
    if (hinfo.width >= 10) line[9] = it.tailles;
    if (hinfo.width >= 11) line[10] = it.couleursEf;

    // Prix: write as TEXT "0.00" (dot)
    line[cPrixHT - 1] = it.prix;

    // Poids: convert to KG text
    line[cPoidsKG - 1] = formatWeightKgEFashionText_(it.poids);

    // Composition: normalized style with import normalizer
    line[cCompositions - 1] = normalizeCompositionEFashionImport_(it.compo) || "Viscose*90,Polyester*10";

    // Extra safety fallback by fixed template positions (A..U) for test import
    if (hinfo.width >= 1)  line[0] = "J&S FASHION";         // A Marque
    if (hinfo.width >= 2)  line[1] = it.ref;                  // B Référence
    if (hinfo.width >= 3)  line[2] = "Femme";               // C Catégorie
    if (hinfo.width >= 4)  line[3] = it.sousCategorie;       // D Sous catégorie
    if (hinfo.width >= 5)  line[4] = it.sousSousCategorie;   // E Sous-sous Catégorie
    if (hinfo.width >= 6)  line[5] = it.pays;                 // F Provenances
    if (hinfo.width >= 7)  line[6] = "melangees";          // G Vendu Par
    if (hinfo.width >= 8)  line[7] = "";                   // H Qte min
    if (hinfo.width >= 9)  line[8] = "Printemps / Été";     // I Collection
    if (hinfo.width >= 10) line[9] = it.tailles;             // J Tailles
    if (hinfo.width >= 11) line[10] = it.couleursEf;         // K Couleurs
    if (hinfo.width >= 12) line[11] = it.prix;                // L Prix HT
    if (hinfo.width >= 14) line[13] = formatWeightKgEFashionText_(it.poids); // N Poids KG
    if (hinfo.width >= 15) line[14] = normalizeCompositionEFashionImport_(it.compo) || "Viscose*90,Polyester*10"; // O Compositions

    out.push(line);
  }

  // Clear + write (force TEXT format to avoid Excel coercion)
  const lr = exp.getLastRow();
  if (lr > 1) {
    exp.getRange(2, 1, lr - 1, hinfo.width).clearContent();
  }

  if (out.length) {
    exp.getRange(2, 1, out.length, hinfo.width).setNumberFormat("@");
    exp.getRange(2, 1, out.length, hinfo.width).setValues(out);
    exp.getRange(2, 1, out.length, hinfo.width).setNumberFormat("@");
    SpreadsheetApp.flush();

    // Auto-uncheck: décoche 选择 pour les lignes exportées
    try {
      const a1 = rows.map(it => stock.getRange(it.row, cSel).getA1Notation());
      stock.getRangeList(a1).setValue(false);
    } catch (e) {
      for (const it of rows) {
        stock.getRange(it.row, cSel).setValue(false);
      }
    }

    // Auto-export Drive after sheet generation
    // Important: flush first so the XLSX export sees the newly written data, not stale previous values.
    try {
      SpreadsheetApp.flush();
      Utilities.sleep(1200);
      exportEFashionToDriveXlsx();
    } catch (e) {
      Logger.log("Auto export Drive eFashion failed: " + e);
    }

    // Also generate photo ZIP helper files in the export Drive folder
    try {
      exportEFashionPhotoZipHelperToDrive_(rows);
    } catch (e) {
      Logger.log("Photo ZIP helper export failed: " + e);
    }
  }

  ss.toast("Export eFashion terminé : " + out.length + " lignes", "eFashion", 5);
}




/****************************************************
 * Normalisation date pour tri
 ****************************************************/
function normalizeDateForSortEf_(v){

  if (!v) return {empty:true,ts:0};

  if (Object.prototype.toString.call(v) === "[object Date]"){
    return {empty:false,ts:v.getTime()};
  }

  const d = new Date(v);

  if (!isNaN(d.getTime())){
    return {empty:false,ts:d.getTime()};
  }

  return {empty:true,ts:0};

}

/****************************************************
 * Export e_export → Google Drive (XLSX)
 * overwrite activé
 ****************************************************/

var SHEET_E_EXPORT = "e_export";
var EFASHION_EXPORT_FOLDER_ID = "1KgBX7CaWqhQZF3lzEmYo4yTStRhCwRkt";
var EFASHION_EXPORT_FILENAME = "EFASHION_EXPORT.xlsx";

function exportEFashionToDriveXlsx() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_E_EXPORT);

  if (!sheet) throw new Error("Feuille introuvable: " + SHEET_E_EXPORT);

  const folder = DriveApp.getFolderById(EFASHION_EXPORT_FOLDER_ID);

  const spreadsheetId = ss.getId();
  const gid = sheet.getSheetId();

  const url =
    "https://docs.google.com/spreadsheets/d/" +
    spreadsheetId +
    "/export?format=xlsx&gid=" +
    gid;

  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + token
    }
  });

  const blob = response.getBlob().setName(EFASHION_EXPORT_FILENAME);

  // overwrite : supprimer anciens fichiers
  const existing = folder.getFilesByName(EFASHION_EXPORT_FILENAME);

  while (existing.hasNext()) {
    existing.next().setTrashed(true);
  }

  // créer nouveau fichier
  const file = folder.createFile(blob);

  SpreadsheetApp.getActive().toast(
    "Export eFashion Drive terminé",
    "eFashion",
    5
  );

  Logger.log("File created: " + file.getUrl());

  return file.getUrl();
}

/****************************************************
 * Robust header reader/matcher for eFashion template
 ****************************************************/
var EF_HEADER_MAX_COLS = 30;
var EF_MIN_EXPORT_COLS = 21; // A..U

function readEFashionHeaders_(sheet) {
  const maxCols = sheet.getMaxColumns();
  const widthScan = Math.min(maxCols, EF_HEADER_MAX_COLS);
  const row = sheet.getRange(1, 1, 1, widthScan).getValues()[0];

  let last = 0;
  for (let i = row.length - 1; i >= 0; i--) {
    if (String(row[i] || "").trim() !== "") { last = i + 1; break; }
  }
  if (last === 0) throw new Error("e_export: ligne 1 sans headers");

  last = Math.max(last, EF_MIN_EXPORT_COLS);
  return { headers: row.slice(0, last), width: last };
}

function normalizeEFHeaderKey_(s) {
  return String(s || "")
    .replace(/^\s*"|"\s*$/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function buildEFashionColumnIndex_(headers) {
  const keys = headers.map(h => normalizeEFHeaderKey_(h));
  return function findCol(needle) {
    const n = normalizeEFHeaderKey_(needle);
    if (!n) return 0;

    // exact
    let idx = keys.indexOf(n);
    if (idx >= 0) return idx + 1;

    // includes
    for (let i = 0; i < keys.length; i++) {
      if (keys[i] && keys[i].indexOf(n) !== -1) return i + 1;
    }
    return 0;
  };
}

/****************************************************
 * Price formatting for eFashion: "4.50" (dot, 2 decimals)
 ****************************************************/
function formatPriceEFashionText_(v) {
  if (v === null || v === "") return "";
  const s = String(v).trim().replace(/\s+/g, "").replace(",", ".");
  const n = Number(s);
  if (!isFinite(n)) return "";
  return n.toFixed(2);
}

/****************************************************
 * Fallback composition normalizer if PFS helper is not loaded
 ****************************************************/
function normalizeCompositionEFashionFallback_(v) {
  const raw = String(v || "").trim();
  if (!raw) return "";
  // Minimal: collapse spaces, fix POLYSTER, and insert " - " between % blocks
  let s = raw.replace(/\u00A0/g, " ").replace(/\s+/g, " ").trim();
  s = s.replace(/POLYSTER/gi, "POLYESTER");
  // Add separator between consecutive percentage blocks if missing
  s = s.replace(/(\d+%\s*[A-Za-zÉÈÊËÀÂÎÏÔÖÛÜÙÇ]+)\s+(?=\d+%)/g, "$1 - ");
  return s;
}

/****************************************************
 * Extract tailles from Contenu colis
 * Examples:
 * "6 x M/L - 6 x XL/XXL" -> "6*M/L,6*XL/XXL"
 * "M/L - XL/XXL" -> "6*M/L,6*XL/XXL"
 ****************************************************/
function extractTaillesFromContenuColisEFashion_(v, colisage) {
  const raw = String(v || "").trim();
  if (!raw) return "";

  const normalized = raw
    .replace(/\u00A0/g, " ")
    .replace(/[／⁄∕]/g, "/")
    .replace(/[–—−]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/\s*-\s*/g, "|")
    .trim();

  const segments = normalized
    .split("|")
    .map(function (part) { return String(part || "").trim(); })
    .filter(Boolean);

  if (!segments.length) return "";

  const total = Number(colisage);
  if (!Number.isFinite(total) || total <= 0) return "";
  if (total % segments.length !== 0) return "";

  const qtyPerSize = total / segments.length;
  const parts = [];

  for (let i = 0; i < segments.length; i++) {
    const parsed = parseContenuColisSegmentEFashion_(segments[i]);
    if (!parsed || !parsed.size) return "";
    parts.push(qtyPerSize + "*" + parsed.size);
  }

  return parts.join(",");
}

/****************************************************
 * eFashion import composition format
 * Expected: "Matière*Pourcentage" separated by commas
 * Example: "Coton*80,Polyester*20"
 ****************************************************/
function normalizeCompositionEFashionImport_(v) {
  const raw0 = String(v || "").trim();
  if (!raw0) return "";

  // Accept inputs like "90% Viscose - 10% Polyester" or "90%VISCOSE 10%POLYESTER"
  // Extract pairs (pct, material)
  const up = raw0
    .replace(/\u00A0/g, " ")
    .replace(/POLYSTER/gi, "POLYESTER")
    .replace(/RAYONNE/gi, "RAYON")
    .replace(/LINEN/gi, "LIN")
    .replace(/ELASTHANNE/gi, "ÉLASTHANNE")
    .replace(/\s+/g, " ")
    .trim();

  const re = /(\d+)\s*%\s*([A-Za-zÉÈÊËÀÂÎÏÔÖÛÜÙÇ]+(?:\s+[A-Za-zÉÈÊËÀÂÎÏÔÖÛÜÙÇ]+)*)/g;
  const parts = [];
  let m;
  while ((m = re.exec(up)) !== null) {
    const pct = m[1];
    const mat = normalizeMaterialEfashion_(m[2]);
    if (mat) parts.push(`${mat}*${pct}`);
  }

  if (!parts.length) return "";
  return parts.join(",");
}

function normalizeMaterialEfashion_(matRaw) {
  const s = String(matRaw || "").trim();
  if (!s) return "";
  const up = s.toUpperCase().replace(/\s+/g, "");

  const MAP = {
    "COTON": "Coton",
    "POLYESTER": "Polyester",
    "VISCOSE": "Viscose",
    "POLYAMIDE": "Polyamide",
    "POLYURÉTHANE": "Polyuréthane",
    "POLYURETHANE": "Polyuréthane",
    "ÉLASTHANNE": "Élasthanne",
    "ELASTHANNE": "Élasthanne",
    "NYLON": "Nylon",
    "LAINE": "Laine",
    "LIN": "Lin",
    "RAYON": "Rayon",
    "SPANDEX": "Spandex"
  };

  if (MAP[up]) return MAP[up];

  // Title-case fallback
  const low = s.toLowerCase();
  return low.charAt(0).toUpperCase() + low.slice(1);
}

/****************************************************
 * Convert STOCK Couleurs to eFashion mixed-color format
 * Input example: "4 ROUGE 2 VERT 2 MARRON 4 BLEU"
 * Output example: "Rouge*2-2,Vert*1-1,Marron*1-1,Bleu*2-2"
 ****************************************************/
function formatEFashionMixedColorsFromStock_(colorPack) {
  if (!colorPack || !colorPack.ok || !Array.isArray(colorPack.entries) || !colorPack.entries.length) return "";

  const parts = [];
  const sizeCount = Number(colorPack.sizeCount) || 0;

  for (let i = 0; i < colorPack.entries.length; i++) {
    const entry = colorPack.entries[i];
    if (!entry || !entry.color) return "";

    if (colorPack.mode === "detailed") {
      if (!Array.isArray(entry.split) || !entry.split.length) return "";
      parts.push(entry.color + "*" + entry.split.join("-"));
      continue;
    }

    if (!sizeCount || !Number.isFinite(entry.qty) || entry.qty <= 0 || entry.qty % sizeCount !== 0) return "";
    const qtyPerSize = entry.qty / sizeCount;
    parts.push(entry.color + "*" + new Array(sizeCount).fill(qtyPerSize).join("-"));
  }

  return parts.join(",");
}

function normalizeEFashionColorName_(raw) {
  const s = String(raw || "").trim();
  if (!s) return "";
  const up = s.toUpperCase();
  const MAP = {
    "ROUGE": "Rouge",
    "VERT": "Vert",
    "MARRON": "Marron",
    "BLEU": "Bleu",
    "TAUPE": "Taupe",
    "NOIR": "Noir",
    "KAKI": "Kaki",
    "BLANC": "Blanc",
    "BEIGE": "Beige",
    "JAUNE": "Jaune",
    "GRIS": "Gris"
  };
  if (MAP[up]) return MAP[up];
  if (typeof normalizePfsColorName_ === "function") {
    return normalizePfsColorName_(s);
  }
  return "";
}

function parseContenuColisSegmentEFashion_(segment) {
  const s = String(segment || "").trim();
  if (!s) return null;

  // Examples:
  // "6 x S/M" -> keep only size S/M
  // "6 x XL/XXL" -> keep only size XL/XXL
  // "S/M" -> keep S/M
  // "XL/XXL" -> keep XL/XXL

  const withQty = s.match(/^(\d+)\s*[xX×*]?\s*(.+)$/);
  if (withQty) {
    const size = normalizeEFashionSizeTokenEFashion_(withQty[2]);
    if (!size) return null;
    return { size: size };
  }

  const sizeOnly = normalizeEFashionSizeTokenEFashion_(s);
  if (sizeOnly) {
    return { size: sizeOnly };
  }

  return null;
}

function normalizeEFashionSizeTokenEFashion_(raw) {
  const s = String(raw || "")
    .replace(/\u00A0/g, " ")
    .replace(/[／⁄∕]/g, "/")
    .replace(/\s+/g, " ")
    .replace(/\s*\/\s*/g, "/")
    .trim()
    .toUpperCase();

  if (!s) return "";

  // Keep slash-based dual sizes intact: S/M, L/XL, M/L, XL/XXL
  if (/^[A-Z0-9]+(?:\/[A-Z0-9]+)+$/.test(s)) return s;

  // Allow simple single tokens if ever needed
  if (/^[A-Z0-9]+$/.test(s)) return s;

  return "";
}

/****************************************************
 * Convert STOCK weight in grams to eFashion KG text
 * Example: 250 -> "0.25"
 ****************************************************/
function formatWeightKgEFashionText_(v) {
  const s = String(v ?? "").trim().replace(/\s+/g, "").replace(",", ".");
  const n = Number(s);
  if (!isFinite(n)) return "";
  const kg = n / 1000;
  return String(kg.toFixed(3)).replace(/\.0+$/, "").replace(/(\.\d*[1-9])0+$/, "$1");
}

/****************************************************
 * Map STOCK category to eFashion Sous catégorie / Sous-sous catégorie
 ****************************************************/
function mapStockCategoryToEFashion_(raw) {
  const s = String(raw || "").trim().toUpperCase();

  const MAP = {
    "CHEMISES / TUNIQUES": { sousCategorie: "Hauts", sousSousCategorie: "Tuniques" },
    "COMBI PANTALON": { sousCategorie: "Robes & Combinaisons", sousSousCategorie: "Combinaisons" },
    "COMBI SHORT": { sousCategorie: "Robes & Combinaisons", sousSousCategorie: "Combinaisons" },
    "CROCHETS": { sousCategorie: "Hauts", sousSousCategorie: "Tops" },
    "ENSEMBLES": { sousCategorie: "Robes & Combinaisons", sousSousCategorie: "Ensembles" },
    "JUPES": { sousCategorie: "Bas", sousSousCategorie: "Jupes" },
    "MANTEAUX / VESTES": { sousCategorie: "Extérieur", sousSousCategorie: "Vestes" },
    "PANTALONS": { sousCategorie: "Bas", sousSousCategorie: "Pantalons" },
    "PULLS / GILETS": { sousCategorie: "Hauts", sousSousCategorie: "Pulls" },
    "ROBES COURTES": { sousCategorie: "Robes & Combinaisons", sousSousCategorie: "Robes courtes" },
    "ROBES LONGUES": { sousCategorie: "Robes & Combinaisons", sousSousCategorie: "Robes longues" },
    "SHORTS": { sousCategorie: "Bas", sousSousCategorie: "Shorts" },
    "TOPS": { sousCategorie: "Hauts", sousSousCategorie: "Tops" },
    "VÊTEMENTS PLAGE": { sousCategorie: "Maillots de bain", sousSousCategorie: "Vêtements de plage" }
  };

  return MAP[s] || { sousCategorie: "Robes & Combinaisons", sousSousCategorie: "Robes longues" };
}

/****************************************************
 * Strict validation for STOCK Couleurs before eFashion export
 * Rules:
 * - syntax must be pairs like: 4 ROUGE 2 VERT 2 BLEU
 * - total quantity must equal 12
 ****************************************************/
function validateEFashionColorPack_(raw, expectedTotal, sizes) {
  if (typeof parseStockColorPackStrict_ === "function") {
    return parseStockColorPackStrict_(raw, {
      normalizeColor: normalizeEFashionColorName_,
      invalidColorReason: "Couleur non reconnue pour eFashion",
      sizes: sizes,
      expectedTotal: expectedTotal
    });
  }
  return { ok: false, reason: "Parseur couleurs partagé introuvable" };
}

function parseEFashionTaillesStructure_(taillesStr) {
  const s = String(taillesStr || "").trim();
  if (!s) return { ok: false, reason: "Tailles vides" };

  const parts = s.split(/\s*,\s*/).filter(Boolean);
  if (!parts.length) return { ok: false, reason: "Tailles vides" };

  const out = [];
  let total = 0;
  let expectedQty = null;

  for (let i = 0; i < parts.length; i++) {
    const m = parts[i].match(/^(\d+)\*\s*(.+)$/);
    if (!m) return { ok: false, reason: "Format taille invalide: " + parts[i] };

    const qty = Number(m[1]);
    const size = normalizeEFashionSizeTokenEFashion_(m[2]);
    if (!Number.isFinite(qty) || qty <= 0 || !size) {
      return { ok: false, reason: "Format taille invalide: " + parts[i] };
    }

    if (expectedQty === null) expectedQty = qty;
    if (qty !== expectedQty) {
      return { ok: false, reason: "Quantités tailles incohérentes" };
    }

    total += qty;
    out.push({ size: size, qty: qty });
  }

  return { ok: true, sizes: out, total: total };
}

function appendValidationHint_(reason) {
  const base = String(reason || "").trim();
  if (!base) return "";

  if (/non divisible par \d+ taille\(s\)/i.test(base)) {
    return base + "\nExemple attendu : 1-2 ORANGE 1-2 VERT 2-1 BLEU 2-1 NOIR";
  }

  return base;
}

function parseColisageEFashion_(v) {
  const s = String(v ?? "").trim().replace(/\s+/g, "").replace(",", ".");
  const n = Number(s);

  if (!Number.isFinite(n) || n <= 0 || !Number.isInteger(n)) {
    return { ok: false, reason: "Colisage invalide: '" + String(v ?? "") + "'" };
  }

  if (n % 2 !== 0) {
    return { ok: false, reason: "Colisage impair non supporté pour un pack 2 tailles: " + n };
  }

  return { ok: true, value: n };
}

/****************************************************
 * Generate photo ZIP helper files in Drive
 * - efashion_photo_colors.json
 * - efashion_photo_zip.py
 ****************************************************/
function exportEFashionPhotoZipHelperToDrive_(rows) {
  const folder = DriveApp.getFolderById(EFASHION_EXPORT_FOLDER_ID);

  const refs = rows.map(function (it) {
    return {
      ref: String(it.ref || "").trim().toUpperCase(),
      color: firstEFashionColor_(it.couleursEf)
    };
  }).filter(function (x) {
    return x.ref;
  });

  const payload = {
    generated_at: new Date().toISOString(),
    refs: refs
  };

  // Always refresh the ref -> color mapping
  upsertDriveTextFile_(folder, "efashion_photo_colors.json", JSON.stringify(payload, null, 2), MimeType.PLAIN_TEXT);

  // Create the Python helper only once (or if missing)
  ensureDriveTextFileExists_(folder, "efashion_photo_zip.py", buildEFashionPhotoZipPythonScript_(), MimeType.PLAIN_TEXT);
}

function firstEFashionColor_(couleursEf) {
  const s = String(couleursEf || "").trim();
  if (!s) return "NOIR";

  const first = s.split(",")[0] || "";
  const color = first.split("*")[0] || "";
  return String(color || "NOIR").trim().toUpperCase();
}

function upsertDriveTextFile_(folder, filename, content, mimeType) {
  const existing = folder.getFilesByName(filename);
  while (existing.hasNext()) {
    existing.next().setTrashed(true);
  }
  folder.createFile(filename, content, mimeType || MimeType.PLAIN_TEXT);
}

function ensureDriveTextFileExists_(folder, filename, content, mimeType) {
  const existing = folder.getFilesByName(filename);
  if (existing.hasNext()) return;
  folder.createFile(filename, content, mimeType || MimeType.PLAIN_TEXT);
}

function buildEFashionPhotoZipPythonScript_() {
  return [
    "#!/usr/bin/env python3",
    "from __future__ import annotations",
    "",
    "import argparse",
    "import json",
    "import re",
    "import zipfile",
    "from collections import defaultdict",
    "from pathlib import Path",
    "",
    "IMAGE_EXTS = {'.jpg', '.jpeg', '.png', '.webp'}",
    "DEFAULT_COLOR = 'NOIR'",
    "MAX_PHOTOS_PER_REF = 20",
    "",
    "def slugify_color(color: str) -> str:",
    "    color = (color or DEFAULT_COLOR).strip().upper()",
    "    color = color.replace(' ', '-')",
    "    color = re.sub(r'[^A-Z0-9\\-À-ÿ]', '-', color)",
    "    color = re.sub(r'-{2,}', '-', color).strip('-')",
    "    return color or DEFAULT_COLOR",
    "",
    "def normalize_ref(ref: str) -> str:",
    "    return ref.strip().upper()",
    "",
    "def parse_input_name(path: Path) -> tuple[str, int] | None:",
    "    stem = path.stem.strip()",
    "    m = re.fullmatch(r'(.+?)_(\\d+)$', stem)",
    "    if m:",
    "        ref = normalize_ref(m.group(1))",
    "        raw_index = int(m.group(2))",
    "        pos = max(0, raw_index - 1)",
    "        return ref, pos",
    "    return normalize_ref(stem), 0",
    "",
    "def load_mapping(mapping_path: Path | None) -> dict[str, str]:",
    "    if mapping_path is None or not mapping_path.exists():",
    "        return {}",
    "    data = json.loads(mapping_path.read_text(encoding='utf-8'))",
    "    out: dict[str, str] = {}",
    "    if isinstance(data, dict) and isinstance(data.get('refs'), list):",
    "        for item in data['refs']:",
    "            if isinstance(item, dict) and item.get('ref'):",
    "                out[normalize_ref(str(item['ref']))] = slugify_color(str(item.get('color', DEFAULT_COLOR)))",
    "    return out",
    "",
    "def build_zip(input_dir: Path, output_zip: Path, mapping: dict[str, str]) -> None:",
    "    files = [p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_EXTS]",
    "    if not files:",
    "        raise SystemExit(f'Aucune image trouvée dans {input_dir}')",
    "",
    "    by_ref: dict[str, list[tuple[int, Path]]] = defaultdict(list)",
    "    for path in files:",
    "        parsed = parse_input_name(path)",
    "        if parsed is None:",
    "            continue",
    "        ref, pos = parsed",
    "        by_ref[ref].append((pos, path))",
    "",
    "    output_zip.parent.mkdir(parents=True, exist_ok=True)",
    "    with zipfile.ZipFile(output_zip, 'w', compression=zipfile.ZIP_DEFLATED) as zf:",
    "        for ref in sorted(by_ref):",
    "            color = mapping.get(ref, DEFAULT_COLOR)",
    "            entries = sorted(by_ref[ref], key=lambda x: (x[0], x[1].name.lower()))",
    "            used_positions: set[int] = set()",
    "            count = 0",
    "            for suggested_pos, path in entries:",
    "                if count >= MAX_PHOTOS_PER_REF:",
    "                    break",
    "                pos = suggested_pos",
    "                while pos in used_positions:",
    "                    pos += 1",
    "                used_positions.add(pos)",
    "                arcname = f'{ref}-{color}-{pos}{path.suffix.lower()}'",
    "                zf.write(path, arcname)",
    "                count += 1",
    "",
    "def main() -> int:",
    "    parser = argparse.ArgumentParser(description='Prépare un ZIP de photos renommées au format {reference-couleur}-{position}.jpg')",
    "    parser.add_argument('--input', default='input', help='Dossier d\'entrée contenant les photos')",
    "    parser.add_argument('--output', default='output/photos_upload.zip', help='Chemin du ZIP de sortie')",
    "    parser.add_argument('--mapping', default='efashion_photo_colors.json', help='Fichier JSON de mapping ref -> couleur')",
    "    args = parser.parse_args()",
    "",
    "    input_dir = Path(args.input)",
    "    output_zip = Path(args.output)",
    "    mapping_path = Path(args.mapping) if args.mapping else None",
    "",
    "    if not input_dir.exists() or not input_dir.is_dir():",
    "        raise SystemExit(f'Dossier input introuvable: {input_dir}')",
    "",
    "    mapping = load_mapping(mapping_path)",
    "    build_zip(input_dir, output_zip, mapping)",
    "    print(f'ZIP créé: {output_zip}')",
    "    return 0",
    "",
    "if __name__ == '__main__':",
    "    raise SystemExit(main())",
    ""
  ].join("\n");
}
