/**
 * Microstore → Sync MS_IMPORT → STOCK (v1.0)
 *
 * Hypothèses:
 * - MS_IMPORT: headers row 2, data row 3+
 * - STOCK: headers row 1, data row 2+
 * - Colonnes déplaçables: repérage par NOM d’en-tête
 * - Écrit sur toutes les colonnes importées de MS_IMPORT :
 *   货号, Nom, Catégorie, Contenu colis, Composition matérielle, Marque, Année, Saison, Colisage,
 *   Couleur, Stock, Nbr de pièces hors unité de colisage, Poids (en gramme), Prix, Pays d'origine,
 *   Remise (%), Remarque, MS_STATUT, MS_LAST_SEEN
 * - MS_LAST_SEEN = DATE (date+heure du lancement du sync)
 * - Statuts:
 *   * si ref vue dans MS_IMPORT => MS
 *   * si ref absente de MS_IMPORT:
 *        - si ancien statut == "MS" => MS_SUPPRIME
 *        - sinon => A_CREER
 * - Refs MS_IMPORT absentes de STOCK => création de ligne + statut MS
 * - Formules à prolonger sur les nouvelles lignes depuis TEMPLATE_STOCK (ligne 2):
 *   Colisage, 包/箱, Stock, Poids (en gramme), Promo, Prix@, Promo@, 进货, 剩下 / RESTE, SortKey
 *
 * Popup confirmation: affiche la date du dernier import depuis LOG_IMPORT (col A, dernière ligne)
 *
 * Prérequis: feuille TEMPLATE_STOCK existe.
 */

var SHEET_MS_IMPORT = "MS_IMPORT"; // source sheet filled by Import_MS

var SHEET_TEMPLATE = "TEMPLATE_STOCK";
var SHEET_LOG_IMPORT = "LOG_IMPORT";



function syncMsImportToStock() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Popup de confirmation avec date dernier import (LOG_IMPORT col A)
  var lastImportDate = getLastImportDateText_();
  var confirmText =
    "Dernier import Microstore (LOG_IMPORT): " + (lastImportDate || "—") +
    "\n\nLancer la synchro MS_IMPORT → STOCK ?";

  var btn = ui.alert("Confirmation", confirmText, ui.ButtonSet.OK_CANCEL);
  if (btn !== ui.Button.OK) return;

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  ss.toast("Sync MS_IMPORT → STOCK : démarrage…", "Microstore", 5);

  try {
    var shMs = ss.getSheetByName(SHEET_MS_IMPORT);
    var shStock = ss.getSheetByName(SHEET_STOCK);
    var shTpl = ss.getSheetByName(SHEET_TEMPLATE);

    if (!shMs) throw new Error("Feuille introuvable: " + SHEET_MS_IMPORT);
    if (!shStock) throw new Error("Feuille introuvable: " + SHEET_STOCK);
    if (!shTpl) throw new Error("Feuille introuvable: " + SHEET_TEMPLATE);

    var nowText = formatNowText_(); // MS_LAST_SEEN = DATE

    // --- 1) Lire MS_IMPORT (headers row2, data row3+) ---
    ss.toast("Lecture MS_IMPORT…", "Microstore", 5);
    var msLastRow = shMs.getLastRow();
    var msLastCol = shMs.getLastColumn();
    if (msLastRow < 3) {
      ss.toast("MS_IMPORT vide (pas de données).", "Microstore", 6);
      return;
    }

    var msHeaders = shMs.getRange(2, 1, 1, msLastCol).getValues()[0];
    var msHeaderMap = headerMap_(msHeaders);

    // Colonnes nécessaires dans MS_IMPORT
    var msNeed = [
      "Référence",
      "Nom",
      "Catégorie",
      "Contenu colis",
      "Composition matérielle",
      "Marque",
      "Année",
      "Saison",
      "Colisage",
      "Couleur",
      "Stock",
      "Nbr de pièces hors unité de colisage",
      "Poids (en gramme)",
      "Prix",
      "Pays d'origine",
      "Remise (%)",
      "Remarque",
      "Date de création"
    ];
    ensureHeadersExist_(msHeaderMap, msNeed, "MS_IMPORT (ligne 2)");

    var msData = shMs.getRange(3, 1, msLastRow - 2, msLastCol).getValues();

    // --- 2) Construire msMap par ref normalisée + doublonCount ---
    ss.toast("Indexation des refs Microstore…", "Microstore", 5);

    var doublonCount = {}; // ref -> count
    var msMap = {};        // ref -> record (dernière occurrence gagne)

    var msOrder = [];        // refs dans l'ordre du fichier (avec répétitions)
    var lastIdxByRef = {};   // ref -> dernier index rencontré
    for (var i = 0; i < msData.length; i++) {
      var row = msData[i];
      var refRaw = getCell_(row, msHeaderMap["référence"]);
      var ref = normalizeRef_(refRaw);
      if (!ref) continue;
      msOrder.push(ref);
      lastIdxByRef[ref] = i;

      doublonCount[ref] = (doublonCount[ref] || 0) + 1;

      // dernière occurrence gagne
      msMap[ref] = {
        ref: ref,
        nom: getCell_(row, msHeaderMap["nom"]),
        categorie: getCell_(row, msHeaderMap["catégorie"]),
        contenuColis: getCell_(row, msHeaderMap["contenu colis"]),
        compo: normalizeUpper_(getCell_(row, msHeaderMap["composition matérielle"])),
        marque: getCell_(row, msHeaderMap["marque"]),
        annee: getCell_(row, msHeaderMap["année"]),
        saison: getCell_(row, msHeaderMap["saison"]),
        colisage: getCell_(row, msHeaderMap["colisage"]),
        couleur: normalizeCouleur_(getCell_(row, msHeaderMap["couleur"])),
        stock: getCell_(row, msHeaderMap["stock"]),
        horsColisage: getCell_(row, msHeaderMap["nbr de pièces hors unité de colisage"]),
        poidsG: getCell_(row, msHeaderMap["poids (en gramme)"]),
        prix: getCell_(row, msHeaderMap["prix"]),
        paysOrigine: getCell_(row, msHeaderMap["pays d'origine"]),
        remise: getCell_(row, msHeaderMap["remise (%)"]),
        remarque: getCell_(row, msHeaderMap["remarque"]),
        dateCreation: getCell_(row, msHeaderMap["date de création"])
      };
    }

    // post-process: Remarque fusion + tag doublon
    var refsMs = [];
var seen = {};
for (var t = 0; t < msOrder.length; t++) {
  var refT = msOrder[t];

  // On ne garde que la DERNIÈRE occurrence (la dernière gagne)
  if (lastIdxByRef[refT] !== t) continue;

  // éviter doublon au cas où
  if (seen[refT]) continue;
  seen[refT] = true;

  refsMs.push(refT);
}
    for (var r = 0; r < refsMs.length; r++) {
      var k = refsMs[r];
      var rec = msMap[k];
      var merged = mergeRemarqueNom_(rec.remarque, rec.nom);
      merged = applyDoublonTag_(merged, doublonCount[k] || 1);
      rec.remarqueFinale = merged;
    }

    // --- 3) Lire STOCK headers + index refs existantes ---
    ss.toast("Lecture STOCK…", "Microstore", 5);
    var stockLastRow = shStock.getLastRow();
    var stockLastCol = shStock.getLastColumn();
    if (stockLastRow < 2) {
      // STOCK vide (juste headers) -> on va créer tout ce qu’il y a dans MS_IMPORT
      stockLastRow = 1;
    }

    var stockHeaders = shStock.getRange(1, 1, 1, stockLastCol).getValues()[0];
    var stockHeaderMap = headerMap_(stockHeaders);

    var stockNeed = [
      "货号",
      "Nom",
      "Catégorie",
      "Contenu colis",
      "Composition matérielle",
      "Marque",
      "Année",
      "Saison",
      "Colisage",
      "Couleur",
      "Stock",
      "Nbr de pièces hors unité de colisage",
      "Poids (en gramme)",
      "Prix",
      "Pays d'origine",
      "Remise (%)",
      "Remarque",
      "Date de création",
      "MS_STATUT",
      "MS_LAST_SEEN"
    ];
    ensureHeadersExist_(stockHeaderMap, stockNeed, "STOCK (ligne 1)");

    // Formules à prolonger sur nouvelles lignes
    var formulaCols = [
      "Colisage",
      "包/箱",
      "Stock",
      "Poids (en gramme)",
      "Promo",
      "Prix@",
      "Promo@",
      "进货",
      "剩下 / RESTE",
      "SortKey"
    ];
    ensureHeadersExist_(stockHeaderMap, formulaCols, "STOCK (ligne 1) - colonnes formules");
    // TEMPLATE doit avoir les mêmes headers pour retrouver les colonnes
    var tplHeaders = shTpl.getRange(1, 1, 1, shTpl.getLastColumn()).getValues()[0];
    var tplHeaderMap = headerMap_(tplHeaders);
    ensureHeadersExist_(tplHeaderMap, formulaCols, "TEMPLATE_STOCK (ligne 1) - colonnes formules");

    var nExisting = Math.max(0, stockLastRow - 1);

    // Lire colonnes nécessaires en STOCK (valeurs existantes)
    var colRef = stockHeaderMap["货号"];
    var colStatus = stockHeaderMap["ms_statut"];
    var colSeen = stockHeaderMap["ms_last_seen"];

    var stockRefs = (nExisting > 0)
      ? shStock.getRange(2, colRef, nExisting, 1).getValues()
      : [];
    var stockStatus = (nExisting > 0)
      ? shStock.getRange(2, colStatus, nExisting, 1).getValues()
      : [];
    var stockSeen = (nExisting > 0)
      ? shStock.getRange(2, colSeen, nExisting, 1).getValues()
      : [];

    // Map ref -> idx (0-based within data block)
    var stockIndex = {};
    for (var s = 0; s < stockRefs.length; s++) {
      var stockRef = normalizeRef_(stockRefs[s][0]);
      if (!stockRef) continue;
      stockIndex[stockRef] = s;
      // On normalise aussi la cellule STOCK "货号" si besoin? (on l’écrit dans colonne 货号 lors update)
    }

    // Préparer tableaux de sortie pour les colonnes cibles (on repart des valeurs existantes)
    var out = initOutColumns_(shStock, stockHeaderMap, stockNeed, nExisting);

    // --- 4) Appliquer updates + préparer ajouts ---
    ss.toast("Calcul updates + ajouts…", "Microstore", 6);

    var toAdd = []; // records à append (dans l’ordre)
    var seenInMs = {}; // ref -> true

    for (var j = 0; j < refsMs.length; j++) {
      var refKey = refsMs[j];
      var rec2 = msMap[refKey];
      seenInMs[refKey] = true;

      if (stockIndex.hasOwnProperty(refKey)) {
        var idx = stockIndex[refKey]; // row offset (0..nExisting-1)
        // Remplir colonnes cibles
        setOutRow_(out, idx, stockHeaderMap, {
          "货号": rec2.ref,
          "Nom": rec2.nom,
          "Catégorie": rec2.categorie,
          "Contenu colis": rec2.contenuColis,
          "Composition matérielle": rec2.compo,
          "Marque": rec2.marque,
          "Année": rec2.annee,
          "Saison": rec2.saison,
          "Colisage": rec2.colisage,
          "Couleur": rec2.couleur,
          "Stock": rec2.stock,
          "Nbr de pièces hors unité de colisage": rec2.horsColisage,
          "Poids (en gramme)": rec2.poidsG,
          "Prix": rec2.prix,
          "Pays d'origine": rec2.paysOrigine,
          "Remise (%)": rec2.remise,
          "Remarque": rec2.remarqueFinale,
          "Date de création": rec2.dateCreation,
          "MS_STATUT": "MS",
          "MS_LAST_SEEN": nowText
        });
      } else {
        // absent de STOCK => créer ligne avec statut MS
        toAdd.push({
          "货号": rec2.ref,
          "Nom": rec2.nom,
          "Catégorie": rec2.categorie,
          "Contenu colis": rec2.contenuColis,
          "Composition matérielle": rec2.compo,
          "Marque": rec2.marque,
          "Année": rec2.annee,
          "Saison": rec2.saison,
          "Colisage": rec2.colisage,
          "Couleur": rec2.couleur,
          "Stock": rec2.stock,
          "Nbr de pièces hors unité de colisage": rec2.horsColisage,
          "Poids (en gramme)": rec2.poidsG,
          "Prix": rec2.prix,
          "Pays d'origine": rec2.paysOrigine,
          "Remise (%)": rec2.remise,
          "Remarque": rec2.remarqueFinale,
          "Date de création": rec2.dateCreation,
          "MS_STATUT": "MS",
          "MS_LAST_SEEN": nowText
        });
      }
    }

    // Refs STOCK absentes de MS_IMPORT => statut MS_SUPPRIME / A_CREER, MS_LAST_SEEN inchangé
    if (nExisting > 0) {
      ss.toast("Mise à jour statuts (absents MS)…", "Microstore", 5);
      for (var rr = 0; rr < stockRefs.length; rr++) {
        var rref = normalizeRef_(stockRefs[rr][0]);
        if (!rref) continue;
        if (seenInMs[rref]) continue;

        var old = (stockStatus[rr] && stockStatus[rr][0]) ? String(stockStatus[rr][0]).trim() : "";
        var newStatus = (old === "MS") ? "MS_SUPPRIME" : "A_CREER";

        // Écrire uniquement MS_STATUT, laisser MS_LAST_SEEN inchangé
        setOutRow_(out, rr, stockHeaderMap, {
          "MS_STATUT": newStatus
        });
      }
    }

    // --- 5) Append nouvelles lignes si besoin ---
var addCount = toAdd.length;
var addedStartRow = 0;

if (addCount > 0) {
  ss.toast("Ajout " + addCount + " nouvelles refs…", "Microstore", 6);

  var lastRowBeforeInsert = shStock.getLastRow();
  if (lastRowBeforeInsert < 1) lastRowBeforeInsert = 1;

  // On ajoute après la dernière ligne "réelle"
  shStock.insertRowsAfter(lastRowBeforeInsert, addCount);

  // Première nouvelle ligne
  addedStartRow = lastRowBeforeInsert + 1;

  // Étendre out columns avec les nouvelles valeurs
  extendOutForAdds_(out, stockHeaderMap, toAdd);
}

    // --- 6) Écrire en batch uniquement les colonnes cibles ---
    ss.toast("Écriture des colonnes cibles…", "Microstore", 6);
    writeOutColumns_(shStock, stockHeaderMap, stockNeed, out);

    // --- 7) Prolonger les formules sur les nouvelles lignes ---
    if (addCount > 0) {
      ss.toast("Prolongation formules (TEMPLATE_STOCK)…", "Microstore", 8);
      applyTemplateFormulas_(shTpl, shStock, tplHeaderMap, stockHeaderMap, formulaCols, addCount, addedStartRow);
    }
    ss.toast("Rebuild filtre (A→AU)…", "Microstore", 5);
    // Prolonger les colonnes dont l'entête contient des notes KEY:* (ex: totaux)
    if (addCount > 0) {
      applyTemplateColumnByHeaderNoteKey_(shTpl, shStock, "KEY:TOTAL_BOX", addCount, addedStartRow);
      applyTemplateColumnByHeaderNoteKey_(shTpl, shStock, "KEY:TOTAL_PCS", addCount, addedStartRow);
    }
    rebuildLockedFilter_A_to_AU_(shStock);
    ss.toast("Sync OK (MS_IMPORT → STOCK).", "Microstore", 8);
  } catch (err) {
    var msg = (err && err.message) ? err.message : String(err);
    SpreadsheetApp.getActiveSpreadsheet().toast("Sync échoué: " + msg, "Microstore", 10);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/** ---------------- Helpers ---------------- **/

function getLastImportDateText_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_LOG_IMPORT);
  if (!sh) return "";
  var lr = sh.getLastRow();
  if (lr < 2) return ""; // header only
  var v = sh.getRange(lr, 1).getDisplayValue(); // Date import
  return v || "";
}

function formatNowText_() {
  var tz = Session.getScriptTimeZone();
  return Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
}

function headerMap_(headersRow) {
  var map = {};
  for (var c = 0; c < headersRow.length; c++) {
    var h = headersRow[c];
    if (h === null || typeof h === "undefined") continue;
    var key = String(h).trim();
    if (!key) continue;
    map[key.toLowerCase()] = c + 1; // 1-based
  }
  return map;
}

function ensureHeadersExist_(map, requiredHeaders, where) {
  var missing = [];
  for (var i = 0; i < requiredHeaders.length; i++) {
    var k = requiredHeaders[i];
    if (!map.hasOwnProperty(k.toLowerCase())) missing.push(k);
  }
  if (missing.length) {
    throw new Error("Headers manquants dans " + where + ": " + missing.join(", "));
  }
}

function getCell_(row, col1Based) {
  var v = row[col1Based - 1];
  return (v === null || typeof v === "undefined") ? "" : String(v);
}

function normalizeRef_(s) {
  var t = (s === null || typeof s === "undefined") ? "" : String(s);
  t = t.trim();
  if (!t) return "";
  return t.toUpperCase();
}

function normalizeCouleur_(s) {
  var t = (s === null || typeof s === "undefined") ? "" : String(s);
  t = t.trim().toUpperCase();
  if (t === "MIXTE") return "MIX";
  if (t === "MIX") return "MIX";
  return t;
}

function mergeRemarqueNom_(remarque, nom) {
  var r = (remarque || "").trim();
  var n = (nom || "").trim();
  if (!n) return r;
  if (!r) return n;

  // Si r contient déjà n (simple contains)
  if (r.indexOf(n) !== -1) return r;
  return r + " | " + n;
}

function stripDoublonTag_(s) {
  var t = (s || "");
  // retire toutes variantes [MS_DOUBLON:x]
  t = t.replace(/\s*\[MS_DOUBLON:\d+\]\s*/g, " ").trim();
  // normalise espaces
  t = t.replace(/\s{2,}/g, " ").trim();
  return t;
}

function applyDoublonTag_(rem, count) {
  var base = stripDoublonTag_(rem);
  if (count && count > 1) {
    if (base) return base + " [MS_DOUBLON:" + count + "]";
    return "[MS_DOUBLON:" + count + "]";
  }
  return base;
}

function initOutColumns_(shStock, stockHeaderMap, stockNeed, nExisting) {
  var out = {};
  for (var i = 0; i < stockNeed.length; i++) {
    var h = stockNeed[i];
    var col = stockHeaderMap[h.toLowerCase()];
    // Lire la colonne existante (texte brut)
    if (nExisting > 0) {
      out[h] = shStock.getRange(2, col, nExisting, 1).getValues(); // 2D
      // force string
      for (var r = 0; r < out[h].length; r++) {
        var v = out[h][r][0];
        out[h][r][0] = (v === null || typeof v === "undefined") ? "" : String(v);
      }
    } else {
      out[h] = [];
    }
  }
  return out;
}

function setOutRow_(out, idx0, stockHeaderMap, patchObj) {
  // patchObj keys are header names in STOCK (exact)
  for (var k in patchObj) {
    if (!patchObj.hasOwnProperty(k)) continue;
    var arr = out[k];
    if (!arr) continue; // sécurité
    while (arr.length <= idx0) arr.push([""]);
    arr[idx0][0] = (patchObj[k] === null || typeof patchObj[k] === "undefined") ? "" : String(patchObj[k]);
  }
}

function extendOutForAdds_(out, stockHeaderMap, toAdd) {
  for (var i = 0; i < toAdd.length; i++) {
    var rec = toAdd[i];
    // étendre chaque colonne de out d’une ligne
    for (var colName in out) {
      if (!out.hasOwnProperty(colName)) continue;
      out[colName].push([""]); // placeholder
    }
    var newIdx = out["货号"].length - 1; // idx de la ligne ajoutée
    setOutRow_(out, newIdx, stockHeaderMap, rec);
  }
}

function writeOutColumns_(shStock, stockHeaderMap, stockNeed, out) {
  var totalRows = 0;
  if (stockNeed.length > 0) totalRows = out[stockNeed[0]].length;

  if (totalRows === 0) return;

  for (var i = 0; i < stockNeed.length; i++) {
    var h = stockNeed[i];
    var col = stockHeaderMap[h.toLowerCase()];
    var data = out[h];
    shStock.getRange(2, col, totalRows, 1).setValues(data);
    // option: forcer format texte
    shStock.getRange(2, col, totalRows, 1).setNumberFormat("@");
  }
}

function applyTemplateFormulas_(shTpl, shStock, tplHeaderMap, stockHeaderMap, formulaCols, addCount, startRow) {
  if (!addCount || addCount <= 0) return;
  if (!startRow || startRow <= 0) throw new Error("applyTemplateFormulas_: startRow invalide");

  for (var i = 0; i < formulaCols.length; i++) {
    var h = String(formulaCols[i] || "");
    var tplCol = tplHeaderMap[h.toLowerCase()];
    var stockCol = stockHeaderMap[h.toLowerCase()];

    if (!tplCol || !stockCol) continue;

    var tplCell = shTpl.getRange(2, tplCol);
    var f = tplCell.getFormulaR1C1();

    if (f) {
      var formulas = [];
      for (var r = 0; r < addCount; r++) formulas.push([f]);
      shStock.getRange(startRow, stockCol, addCount, 1).setFormulasR1C1(formulas);
    } else {
      var v = tplCell.getValue();
      var values = [];
      for (var rr = 0; rr < addCount; rr++) values.push([v]);
      shStock.getRange(startRow, stockCol, addCount, 1).setValues(values);
    }
  }
}
function rebuildLockedFilter_A_to_AU_(sh) {
  // A..AU = 47 colonnes
  var lastRow = sh.getLastRow();
  if (lastRow < 1) lastRow = 1;

  var existing = sh.getFilter();
  if (existing) existing.remove();

  sh.getRange(1, 1, lastRow, 47).createFilter();
}
function applyTemplateColumnByHeaderNoteKey_(shTpl, shStock, key, addCount, startRow) {
  // Prolonge la formule/valeur de TEMPLATE_STOCK (ligne 2) vers STOCK, sur les nouvelles lignes,
  // en identifiant la colonne via la NOTE de l'entête (ligne 1) contenant `key`.
  if (!addCount || addCount <= 0) return;
  if (!startRow || startRow <= 0) throw new Error("applyTemplateColumnByHeaderNoteKey_: startRow invalide");

  var stockLastCol = shStock.getLastColumn();
  var tplLastCol = shTpl.getLastColumn();

  var stockHeaderRange = shStock.getRange(1, 1, 1, stockLastCol);
  var tplHeaderRange = shTpl.getRange(1, 1, 1, tplLastCol);

  var stockNotes = stockHeaderRange.getNotes()[0];
  var tplNotes = tplHeaderRange.getNotes()[0];

  var stockCol = -1;
  var tplCol = -1;

  for (var c = 0; c < stockNotes.length; c++) {
    if (stockNotes[c] && stockNotes[c].indexOf(key) !== -1) {
      stockCol = c + 1;
      break;
    }
  }

  for (var t = 0; t < tplNotes.length; t++) {
    if (tplNotes[t] && tplNotes[t].indexOf(key) !== -1) {
      tplCol = t + 1;
      break;
    }
  }

  if (stockCol <= 0 || tplCol <= 0) return;

  var tplCell = shTpl.getRange(2, tplCol);
  var f = tplCell.getFormulaR1C1();

  if (f) {
    var formulas = [];
    for (var r = 0; r < addCount; r++) formulas.push([f]);
    shStock.getRange(startRow, stockCol, addCount, 1).setFormulasR1C1(formulas);
  } else {
    var v = tplCell.getValue();
    var values = [];
    for (var rr = 0; rr < addCount; rr++) values.push([v]);
    shStock.getRange(startRow, stockCol, addCount, 1).setValues(values);
  }
}
function normalizeUpper_(s) {
  var t = (s === null || typeof s === "undefined") ? "" : String(s);
  return t.trim().toUpperCase();
}

/****************************************************
 * Microstore XLSX Importer (v1.1)
 *
 * - Prend le dernier .xlsx par DATE DE CRÉATION (1A)
 * - Convertit en Google Sheet temporaire
 * - Importe TOUT en texte brut dans MS_IMPORT
 * - Log dans LOG_IMPORT
 * - Archive le .xlsx dans sous-dossier yyyy-MM-dd
 * - Supprime le fichier temporaire converti (2A)
 *
 * Prérequis:
 * 1) Apps Script > Services avancés Google > activer "Drive API"
 * 2) Projet GCP associé > activer aussi "Google Drive API" si demandé
 ****************************************************/

// Dossier Drive source (dépôt des .xlsx Microstore à importer)
var MS_IMPORT_FOLDER_ID = "1nIFyDmc4Qg2Kf8gzHWbmcGgRDRWa_mM8";
var SHEET_IMPORT_NAME = "MS_IMPORT";
var SHEET_LOG_NAME = "LOG_IMPORT";



function importMicrostoreLatestExport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  ss.toast("Import Microstore : démarrage…", "Microstore", 5);

  try {
    var shImport = ss.getSheetByName(SHEET_IMPORT_NAME);
    var shLog = ss.getSheetByName(SHEET_LOG_NAME);

    if (!shImport) throw new Error("Feuille introuvable: " + SHEET_IMPORT_NAME);
    if (!shLog) throw new Error("Feuille introuvable: " + SHEET_LOG_NAME);

    // 1) Trouver le dernier XLSX par DATE DE CRÉATION
    ss.toast("Recherche du dernier export .xlsx…", "Microstore", 5);
    var latestXlsx = findLatestXlsxByCreatedTime_(MS_IMPORT_FOLDER_ID);
    if (!latestXlsx) {
      ss.toast("Aucun .xlsx trouvé dans MS_Export.", "Microstore", 6);
      ui.alert("Aucun fichier .xlsx trouvé dans le dossier MS_Export.");
      return;
    }

    ss.toast("Dernier fichier: " + latestXlsx.name, "Microstore", 5);

    // 2) Convertir le XLSX en Google Sheet temporaire (Drive API)
    ss.toast("Conversion .xlsx → Google Sheet temporaire…", "Microstore", 6);
    var tempFileId = convertXlsxToGoogleSheet_(latestXlsx.id);

    // 3) Ouvrir le fichier converti (la conversion peut prendre du temps côté Google)
    ss.toast("Ouverture du fichier converti…", "Microstore", 6);
    var tempSs = openSpreadsheetWithRetry_(tempFileId, 8, 1500);
    var tempSheets = tempSs.getSheets();
    if (!tempSheets || tempSheets.length === 0) throw new Error("Le fichier converti ne contient aucune feuille.");
    var tempFirst = tempSheets[0];

    // 4) Lecture des données — éviter getDataRange() (peut être énorme si 'range fantôme')
    ss.toast("Lecture des données…", "Microstore", 6);
    var lastRow = tempFirst.getLastRow();
    var lastCol = tempFirst.getLastColumn();
    if (!lastRow || !lastCol) {
      lastRow = 0;
      lastCol = 0;
    }

    var values = [];
    if (lastRow > 0 && lastCol > 0) {
      // Lecture en une fois du rectangle utilisé
      values = tempFirst.getRange(1, 1, lastRow, lastCol).getValues();
    }

    // 5) Import en TEXTE BRUT dans MS_IMPORT
    ss.toast("Import en texte brut dans " + SHEET_IMPORT_NAME + "…", "Microstore", 6);

    shImport.clearContents();

    var dims = getEffectiveDimensions_(values);
    var numRows = dims.rows;
    var numCols = dims.cols;

    if (numRows === 0 || numCols === 0) {
      logImport_(shLog, {
        importDate: new Date(),
        fileName: latestXlsx.name,
        fileCreated: latestXlsx.createdTime,
        rows: 0,
        cols: 0,
        user: getActiveUserEmail_()
      });

      archiveFileToDatedFolder_(MS_IMPORT_FOLDER_ID, latestXlsx.id, new Date());
      Drive.Files.remove(tempFileId);

      ss.toast("Import OK (fichier vide).", "Microstore", 6);
      ui.alert("Import OK (fichier vide).");
      return;
    }

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
    targetRange.setNumberFormat("@"); // TEXTE
    targetRange.setValues(trimmed);

    // 6) Logger
    ss.toast("Écriture du log…", "Microstore", 5);
    logImport_(shLog, {
      importDate: new Date(),
      fileName: latestXlsx.name,
      fileCreated: latestXlsx.createdTime,
      rows: numRows,
      cols: numCols,
      user: getActiveUserEmail_()
    });

    // 7) Archivage
    ss.toast("Archivage Drive…", "Microstore", 6);
    archiveFileToDatedFolder_(MS_IMPORT_FOLDER_ID, latestXlsx.id, new Date());

    // 8) Supprimer le fichier temporaire
    Drive.Files.remove(tempFileId);

    ss.toast("Import OK: " + numRows + " lignes × " + numCols + " colonnes", "Microstore", 8);
    ui.alert("Import OK: " + latestXlsx.name + "\n" + numRows + " lignes × " + numCols + " colonnes");
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

/** -------- Helpers -------- **/

function findLatestXlsxByCreatedTime_(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();

  var best = null; // {id, name, createdTime: Date}
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

function convertXlsxToGoogleSheet_(xlsxFileId) {
  var xlsxMeta = Drive.Files.get(xlsxFileId, { fields: "name" });
  var resource = {
    title: "TMP_MS_CONVERT__" + xlsxMeta.name + "__" + Utilities.getUuid(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  var converted = Drive.Files.copy(resource, xlsxFileId);
  return converted.id;
}

function getEffectiveDimensions_(values) {
  if (!values || values.length === 0) return { rows: 0, cols: 0 };

  var lastRow = -1;
  var lastCol = -1;

  for (var r = 0; r < values.length; r++) {
    var row = values[r];
    var rowHasData = false;
    for (var c = 0; c < row.length; c++) {
      var v = row[c];
      if (v !== "" && v !== null && typeof v !== "undefined") {
        rowHasData = true;
        if (c > lastCol) lastCol = c;
      }
    }
    if (rowHasData) lastRow = r;
  }

  return {
    rows: (lastRow >= 0) ? (lastRow + 1) : 0,
    cols: (lastCol >= 0) ? (lastCol + 1) : 0
  };
}

function logImport_(shLog, info) {
  ensureLogHeader_(shLog);

  var row = [
    info.importDate,
    info.fileName,
    info.fileCreated,
    info.rows,
    info.cols,
    info.user
  ];

  var last = shLog.getLastRow();
  shLog.getRange(last + 1, 1, 1, row.length).setValues([row]);

  shLog.getRange(last + 1, 1).setNumberFormat("yyyy-MM-dd HH:mm:ss");
  shLog.getRange(last + 1, 3).setNumberFormat("yyyy-MM-dd HH:mm:ss");
}

function ensureLogHeader_(shLog) {
  if (shLog.getLastRow() === 0) {
    shLog.getRange(1, 1, 1, 6).setValues([[
      "Date import",
      "Fichier",
      "Date du fichier",
      "Lignes",
      "Colonnes",
      "Utilisateur"
    ]]);
    shLog.getRange(1, 1, 1, 6).setFontWeight("bold");
  }
}

function archiveFileToDatedFolder_(sourceFolderId, fileId, importDate) {
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

function getActiveUserEmail_() {
  try {
    var email = Session.getActiveUser().getEmail();
    return email || "(inconnu)";
  } catch (e) {
    return "(inconnu)";
  }
}

function openSpreadsheetWithRetry_(fileId, attempts, sleepMs) {
  attempts = attempts || 6;
  sleepMs = sleepMs || 1000;

  var lastErr = null;
  for (var i = 0; i < attempts; i++) {
    try {
      return SpreadsheetApp.openById(fileId);
    } catch (e) {
      lastErr = e;
      // Attendre un peu : Drive a parfois créé le fichier mais Sheets n'est pas encore prêt
      Utilities.sleep(sleepMs);
    }
  }
  throw lastErr || new Error("Impossible d'ouvrir le fichier converti après plusieurs tentatives.");
}

/****************************************************
 * Export STOCK → MS_EXPORT (Microstore) (v1.0)
 *
 * - MS_EXPORT: headers row 2 (A2:Q2), data row 3+
 * - STOCK: headers row 1, data row 2+
 * - Export colonnes:
 *   Référence, Nom, Catégorie, Contenu colis (généré),
 *   Composition matérielle (UPPER), Marque, Année, Saison,
 *   Colisage, Couleur, Stock, Nbr de pièces hors unité de colisage,
 *   Poids (en gramme), Prix, Pays d'origine, Remise (%),
 *   Remarque (généré)
 *
 * - Tri STOCK: Date de création
 *   * vides en haut
 *   * puis décroissant (Z→A = plus récent → plus ancien)
 ****************************************************/

var SHEET_MS_EXPORT = "MS_EXPORT"; // adapte si ton onglet s’appelle autrement

function exportStockToMsExport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shStock = ss.getSheetByName(SHEET_STOCK);
  var shExp = ss.getSheetByName(SHEET_MS_EXPORT);

  if (!shStock) throw new Error("Feuille introuvable: " + SHEET_STOCK);
  if (!shExp) throw new Error("Feuille introuvable: " + SHEET_MS_EXPORT);

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  ss.toast("Export STOCK → MS_EXPORT : démarrage…", "Microstore", 5);

  try {
    // --- 1) Lire headers MS_EXPORT (row 2) ---
    var expLastCol = shExp.getLastColumn();
    if (expLastCol < 1) throw new Error("MS_EXPORT semble vide (aucune colonne).");

    var expHeaders = shExp.getRange(2, 1, 1, expLastCol).getValues()[0];
    var expMap = headerMap_(expHeaders);

    var expNeed = [
      "Référence",
      "Nom",
      "Catégorie",
      "Contenu colis",
      "Composition matérielle",
      "Marque",
      "Année",
      "Saison",
      "Colisage",
      "Couleur",
      "Stock",
      "Nbr de pièces hors unité de colisage",
      "Poids (en gramme)",
      "Prix",
      "Pays d'origine",
      "Remise (%)",
      "Remarque"
    ];
    ensureHeadersExist_(expMap, expNeed, "MS_EXPORT (ligne 2)");

    // --- 2) Lire STOCK ---
    var stockLastRow = shStock.getLastRow();
    var stockLastCol = shStock.getLastColumn();
    if (stockLastRow < 2) {
      ss.toast("STOCK vide (rien à exporter).", "Microstore", 6);
      clearMsExportData_(shExp, expLastCol);
      return;
    }

    var stockHeaders = shStock.getRange(1, 1, 1, stockLastCol).getValues()[0];
    var stockMap = headerMap_(stockHeaders);

    var stockNeed = [
      "货号",
      "Nom",
      "Catégorie",
      "Contenu colis",
      "Composition matérielle",
      "Marque",
      "Année",
      "Saison",
      "Colisage",
      "Couleur",
      "Stock",
      "Nbr de pièces hors unité de colisage",
      "Poids (en gramme)",
      "Prix",
      "Pays d'origine",
      "Remise (%)",
      "Date de création",
      "每箱件数",
      "包/箱"
    ];
    ensureHeadersExist_(stockMap, stockNeed, "STOCK (ligne 1)");

    var data = shStock.getRange(2, 1, stockLastRow - 1, stockLastCol).getValues();

    // --- Confirmation simple d'exécution ---
    var ui = SpreadsheetApp.getUi();
    var msg = "⚠️ L'export va remplacer tout le contenu de MS_EXPORT avec les données de STOCK.\n\nContinuer ?";
    var btn = ui.alert("Confirmer export Microstore", msg, ui.ButtonSet.OK_CANCEL);
    if (btn !== ui.Button.OK) {
      ss.toast("Export annulé.", "Microstore", 4);
      return;
    }

    // --- 3) Construire lignes exportables ---
    var cRef = stockMap["货号"] - 1;
    var cNom = stockMap["nom"] - 1;
    var cCat = stockMap["catégorie"] - 1;
    var cContenu = stockMap["contenu colis"] - 1;
    var cComp = stockMap["composition matérielle"] - 1;
    var cMarque = stockMap["marque"] - 1;
    var cAnnee = stockMap["année"] - 1;
    var cSaison = stockMap["saison"] - 1;
    var cStock = stockMap["stock"] - 1;
    var cColisage = stockMap["colisage"] - 1;
    var cCouleur = stockMap["couleur"] - 1;
    var cPrix = stockMap["prix"] - 1;
    var cRemise = stockMap["remise (%)"] - 1;

    var cDate = stockMap["date de création"] - 1;
    var cTotalPcs = stockMap["每箱件数"] - 1;
    var cPacks = stockMap["包/箱"] - 1;

    var cHorsColisage = stockMap["nbr de pièces hors unité de colisage"] - 1;
    var cPoidsG = stockMap["poids (en gramme)"] - 1;
    var cPaysOrigine = stockMap["pays d'origine"] - 1;

    var rows = [];
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      var ref = normalizeRef_(r[cRef]);
      if (!ref) continue;

      var dt = r[cDate];
      var dtInfo = normalizeDateForSort_(dt); // { empty: bool, ts: number }

      rows.push({
        sortEmpty: dtInfo.empty,
        sortTs: dtInfo.ts,
        ref: ref,
        nom: (r[cNom] === null || typeof r[cNom] === "undefined") ? "" : String(r[cNom]).trim(),
        cat: (r[cCat] === null || typeof r[cCat] === "undefined") ? "" : String(r[cCat]).trim(),
        comp: normalizeUpper_(r[cComp]),
        marque: (r[cMarque] === null || typeof r[cMarque] === "undefined") ? "" : String(r[cMarque]).trim(),
        annee: (r[cAnnee] === null || typeof r[cAnnee] === "undefined") ? "" : String(r[cAnnee]).trim(),
        saison: (r[cSaison] === null || typeof r[cSaison] === "undefined") ? "" : String(r[cSaison]).trim(),
        contenuColis: (r[cContenu] === null || typeof r[cContenu] === "undefined") ? "" : String(r[cContenu]).trim(),
        colisage: (r[cColisage] === null || typeof r[cColisage] === "undefined") ? "" : r[cColisage],
        couleur: normalizeCouleur_(r[cCouleur]),
        prix: (r[cPrix] === null || typeof r[cPrix] === "undefined") ? "" : r[cPrix],
        remise: (r[cRemise] === null || typeof r[cRemise] === "undefined") ? "" : r[cRemise],
        totalPcs: r[cTotalPcs],
        stock: r[cStock],
        packs: r[cPacks],
        horsColisage: r[cHorsColisage],
        poidsG: r[cPoidsG],
        paysOrigine: (r[cPaysOrigine] === null || typeof r[cPaysOrigine] === "undefined") ? "" : String(r[cPaysOrigine]).trim()
      });
    }

    // --- 4) Trier: dates vides en haut, sinon décroissant ---
    rows.sort(function(a, b) {
      if (a.sortEmpty && !b.sortEmpty) return -1;
      if (!a.sortEmpty && b.sortEmpty) return 1;
      // les deux vides -> stable-ish
      if (a.sortEmpty && b.sortEmpty) return 0;
      // décroissant ts
      return b.sortTs - a.sortTs;
    });

    // --- 5) Construire matrice d’export ---
    var out = [];
    for (var k = 0; k < rows.length; k++) {
      var it = rows[k];

      var total = (typeof toIntSafe_ === "function") ? toIntSafe_(it.totalPcs) : safeInt_(it.totalPcs);
      var packs = (typeof toIntSafe_ === "function") ? toIntSafe_(it.packs) : safeInt_(it.packs);
      var ppp   = (typeof toIntSafe_ === "function") ? toIntSafe_(it.colisage) : safeInt_(it.colisage);

      var contenu = buildContenuColis_(total, packs, ppp);

      // ligne MS_EXPORT = largeur expLastCol (A..Q), on ne remplit que les colonnes utiles
      var line = new Array(expLastCol).fill("");

      line[expMap["référence"] - 1] = it.ref;
      line[expMap["nom"] - 1] = it.nom;
      line[expMap["catégorie"] - 1] = it.cat;
      line[expMap["contenu colis"] - 1] = it.contenuColis;
      line[expMap["composition matérielle"] - 1] = it.comp;
      line[expMap["marque"] - 1] = it.marque;
      line[expMap["année"] - 1] = it.annee;
      line[expMap["saison"] - 1] = it.saison;
      line[expMap["colisage"] - 1] = it.colisage;
      line[expMap["couleur"] - 1] = it.couleur;
      line[expMap["stock"] - 1] = it.stock;
      line[expMap["nbr de pièces hors unité de colisage"] - 1] = it.horsColisage;
      line[expMap["poids (en gramme)"] - 1] = it.poidsG;
      line[expMap["prix"] - 1] = it.prix;
      line[expMap["pays d'origine"] - 1] = it.paysOrigine;
      line[expMap["remise (%)"] - 1] = it.remise;
      line[expMap["remarque"] - 1] = contenu; // Remarque générée

      out.push(line);
    }

    // --- 6) Écrire MS_EXPORT ---
    clearMsExportData_(shExp, expLastCol);

    if (out.length) {
      shExp.getRange(3, 1, out.length, expLastCol).setValues(out);
    }

    ss.toast("Export terminé: " + out.length + " lignes.", "Microstore", 6);
  } finally {
    lock.releaseLock();
  }
}

// ------------------ Helpers ------------------

function clearMsExportData_(shExp, expLastCol) {
  var lr = shExp.getLastRow();
  if (lr >= 3) {
    shExp.getRange(3, 1, lr - 2, expLastCol).clearContent();
  }
}

/**
 * Retourne { empty:boolean, ts:number }
 * - empty=true si vide / non parseable
 * - ts=timestamp si ok
 */
function normalizeDateForSort_(v) {
  if (v === null || typeof v === "undefined" || v === "") {
    return { empty: true, ts: 0 };
  }
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    return { empty: false, ts: v.getTime() };
  }
  var s = String(v).trim();
  if (!s) return { empty: true, ts: 0 };

  // Tentative parse JS (accepte yyyy-mm-dd etc.)
  var d = new Date(s);
  if (!isNaN(d.getTime())) return { empty: false, ts: d.getTime() };

  return { empty: true, ts: 0 };
}

function buildContenuColis_(totalPieces, packs, pcsPerPack) {
  var t = (totalPieces && totalPieces !== 0) ? String(totalPieces) : "x";
  var p = (packs && packs !== 0) ? String(packs) : "x";
  var u = (pcsPerPack && pcsPerPack !== 0) ? String(pcsPerPack) : "x";
  return "Colis: " + t + " pièces avec " + p + " paquets de " + u + " pièces";
}

// fallback si jamais toIntSafe_ n’existe pas (mais chez toi il existe)
function safeInt_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return Math.trunc(v);
  var m = String(v).match(/-?\d+(\.\d+)?/);
  if (!m) return 0;
  var n = Number(m[0]);
  return isFinite(n) ? Math.trunc(n) : 0;
}

/****************************************************
 * Export MS_EXPORT → Google Drive (XLSX)
 * overwrite activé
 ****************************************************/

var SHEET_MS_EXPORT = "MS_EXPORT";
var MS_EXPORT_DRIVE_FOLDER_ID = "1B2P4SwJbwbPEXW5XO_LSNdIMYA-aYMZa";
var MS_EXPORT_FILENAME = "MS_EXPORT.xlsx";

function exportMsExportSheetToDriveXlsx() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MS_EXPORT);
  if (!sheet) throw new Error("Feuille introuvable: " + SHEET_MS_EXPORT);

  const folder = DriveApp.getFolderById(MS_EXPORT_DRIVE_FOLDER_ID);

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

  const blob = response.getBlob().setName(MS_EXPORT_FILENAME);

  // overwrite : supprimer anciens fichiers
  const existing = folder.getFilesByName(MS_EXPORT_FILENAME);

  while (existing.hasNext()) {
    existing.next().setTrashed(true);
  }

  // créer nouveau fichier
  const file = folder.createFile(blob);

  SpreadsheetApp.getActive().toast(
    "Export Drive terminé (overwrite)",
    "MS Export",
    5
  );

  Logger.log("File created: " + file.getUrl());

  return file.getUrl();
}