/**
 * Microstore → Sync MS_IMPORT → STOCK
 * - Conserve le comportement métier existant
 */
function syncMsImportToStock() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var lastImportDate = msGetLastImportDateText_();
  var confirmText =
    "Dernier import Microstore (LOG_IMPORT_EXPORT): " + (lastImportDate || "—") +
    "\n\nLancer la synchro MS_IMPORT → STOCK ?";

  var btn = ui.alert("Confirmation", confirmText, ui.ButtonSet.OK_CANCEL);
  if (btn !== ui.Button.OK) return;

  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  ss.toast("Sync MS_IMPORT → STOCK : démarrage…", "Microstore", 5);

  try {
    var shMs = ss.getSheetByName(SHEET_MS_IMPORT);
    var shMsDisabled = ss.getSheetByName(SHEET_MS_IMPORT_DISABLED);
    var shStock = ss.getSheetByName(SHEET_STOCK);
    var shTpl = ss.getSheetByName(SHEET_TEMPLATE_STOCK);

    if (!shMs) throw new Error("Feuille introuvable: " + SHEET_MS_IMPORT);
    if (!shStock) throw new Error("Feuille introuvable: " + SHEET_STOCK);
    if (!shTpl) throw new Error("Feuille introuvable: " + SHEET_TEMPLATE_STOCK);
    if (!shMsDisabled) throw new Error("Feuille introuvable: " + SHEET_MS_IMPORT_DISABLED);

    var nowText = msFormatNowText_();

    ss.toast("Lecture MS_IMPORT…", "Microstore", 5);
    var msLastRow = shMs.getLastRow();
    var msLastCol = shMs.getLastColumn();
    if (msLastRow < 3) {
      ss.toast("MS_IMPORT vide (pas de données).", "Microstore", 6);
      return;
    }

    var msHeaders = shMs.getRange(2, 1, 1, msLastCol).getValues()[0];
    var msHeaderMap = headerMap_(msHeaders);

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

    ss.toast("Indexation des refs Microstore…", "Microstore", 5);

    var doublonCount = {};
    var msMap = {};
    var msOrder = [];
    var lastIdxByRef = {};

    for (var i = 0; i < msData.length; i++) {
      var row = msData[i];
      var refRaw = msGetCell_(row, msHeaderMap["référence"]);
      var ref = msNormalizeRef_(refRaw);
      if (!ref) continue;

      msOrder.push(ref);
      lastIdxByRef[ref] = i;
      doublonCount[ref] = (doublonCount[ref] || 0) + 1;

      msMap[ref] = {
        ref: ref,
        nom: msGetCell_(row, msHeaderMap["nom"]),
        categorie: msGetCell_(row, msHeaderMap["catégorie"]),
        contenuColis: msGetCell_(row, msHeaderMap["contenu colis"]),
        compo: normalizeUpper_(msGetCell_(row, msHeaderMap["composition matérielle"])),
        marque: msGetCell_(row, msHeaderMap["marque"]),
        annee: msGetCell_(row, msHeaderMap["année"]),
        saison: msGetCell_(row, msHeaderMap["saison"]),
        colisage: msGetCell_(row, msHeaderMap["colisage"]),
        couleur: msNormalizeCouleur_(msGetCell_(row, msHeaderMap["couleur"])),
        stock: msGetCell_(row, msHeaderMap["stock"]),
        horsColisage: msGetCell_(row, msHeaderMap["nbr de pièces hors unité de colisage"]),
        poidsG: msGetCell_(row, msHeaderMap["poids (en gramme)"]),
        prix: msGetCell_(row, msHeaderMap["prix"]),
        paysOrigine: msGetCell_(row, msHeaderMap["pays d'origine"]),
        remise: msGetCell_(row, msHeaderMap["remise (%)"]),
        remarque: msGetCell_(row, msHeaderMap["remarque"]),
        dateCreation: msGetCell_(row, msHeaderMap["date de création"])
      };
    }

    var refsMs = [];
    var seen = {};
    for (var t = 0; t < msOrder.length; t++) {
      var refT = msOrder[t];
      if (lastIdxByRef[refT] !== t) continue;
      if (seen[refT]) continue;
      seen[refT] = true;
      refsMs.push(refT);
    }

    for (var r = 0; r < refsMs.length; r++) {
      var k = refsMs[r];
      var rec = msMap[k];
      var merged = msMergeRemarqueNom_(rec.remarque, rec.nom);
      rec.remarqueFinale = msApplyDoublonTag_(merged, doublonCount[k] || 1);
    }

    ss.toast("Lecture MS_IMPORT_DISABLED…", "Microstore", 5);

    var disabledIndex = {};
    var disabledMap = {};
    var disabledOrder = [];
    var disabledLastIdxByRef = {};
    var disabledDoublonCount = {};
    var disLastRow = shMsDisabled.getLastRow();
    var disLastCol = shMsDisabled.getLastColumn();

    if (disLastRow >= 3 && disLastCol > 0) {
      var disHeaders = shMsDisabled.getRange(2, 1, 1, disLastCol).getValues()[0];
      var disHeaderMap = headerMap_(disHeaders);
      var disNeed = [
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
      ensureHeadersExist_(disHeaderMap, disNeed, "MS_IMPORT_DISABLED (ligne 2)");

      var disData = shMsDisabled.getRange(3, 1, disLastRow - 2, disLastCol).getValues();

      for (var d = 0; d < disData.length; d++) {
        var rowDis = disData[d];
        var refRawDis = msGetCell_(rowDis, disHeaderMap["référence"]);
        var refDis = msNormalizeRef_(refRawDis);
        if (!refDis) continue;

        disabledIndex[refDis] = true;
        disabledOrder.push(refDis);
        disabledLastIdxByRef[refDis] = d;
        disabledDoublonCount[refDis] = (disabledDoublonCount[refDis] || 0) + 1;

        disabledMap[refDis] = {
          ref: refDis,
          nom: msGetCell_(rowDis, disHeaderMap["nom"]),
          categorie: msGetCell_(rowDis, disHeaderMap["catégorie"]),
          contenuColis: msGetCell_(rowDis, disHeaderMap["contenu colis"]),
          compo: normalizeUpper_(msGetCell_(rowDis, disHeaderMap["composition matérielle"])),
          marque: msGetCell_(rowDis, disHeaderMap["marque"]),
          annee: msGetCell_(rowDis, disHeaderMap["année"]),
          saison: msGetCell_(rowDis, disHeaderMap["saison"]),
          colisage: msGetCell_(rowDis, disHeaderMap["colisage"]),
          couleur: msNormalizeCouleur_(msGetCell_(rowDis, disHeaderMap["couleur"])),
          stock: msGetCell_(rowDis, disHeaderMap["stock"]),
          horsColisage: msGetCell_(rowDis, disHeaderMap["nbr de pièces hors unité de colisage"]),
          poidsG: msGetCell_(rowDis, disHeaderMap["poids (en gramme)"]),
          prix: msGetCell_(rowDis, disHeaderMap["prix"]),
          paysOrigine: msGetCell_(rowDis, disHeaderMap["pays d'origine"]),
          remise: msGetCell_(rowDis, disHeaderMap["remise (%)"]),
          remarque: msGetCell_(rowDis, disHeaderMap["remarque"]),
          dateCreation: msGetCell_(rowDis, disHeaderMap["date de création"])
        };
      }

      var refsDisabled = [];
      var seenDisabled = {};
      for (var dd = 0; dd < disabledOrder.length; dd++) {
        var refDD = disabledOrder[dd];
        if (disabledLastIdxByRef[refDD] !== dd) continue;
        if (seenDisabled[refDD]) continue;
        seenDisabled[refDD] = true;
        refsDisabled.push(refDD);
      }

      for (var dr = 0; dr < refsDisabled.length; dr++) {
        var dk = refsDisabled[dr];
        var drec = disabledMap[dk];
        var dmerged = msMergeRemarqueNom_(drec.remarque, drec.nom);
        drec.remarqueFinale = msApplyDoublonTag_(dmerged, disabledDoublonCount[dk] || 1);
      }
    }

    ss.toast("Lecture STOCK…", "Microstore", 5);
    var stockLastRow = shStock.getLastRow();
    var stockLastCol = shStock.getLastColumn();
    if (stockLastRow < 2) stockLastRow = 1;

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

    var tplHeaders = shTpl.getRange(1, 1, 1, shTpl.getLastColumn()).getValues()[0];
    var tplHeaderMap = headerMap_(tplHeaders);
    ensureHeadersExist_(tplHeaderMap, formulaCols, "TEMPLATE_STOCK (ligne 1) - colonnes formules");

    var nExisting = Math.max(0, stockLastRow - 1);
    var colRef = stockHeaderMap["货号"];
    var colStatus = stockHeaderMap["ms_statut"];

    var stockRefs = (nExisting > 0) ? shStock.getRange(2, colRef, nExisting, 1).getValues() : [];
    var stockStatus = (nExisting > 0) ? shStock.getRange(2, colStatus, nExisting, 1).getValues() : [];

    var stockIndex = {};
    for (var s = 0; s < stockRefs.length; s++) {
      var stockRef = msNormalizeRef_(stockRefs[s][0]);
      if (!stockRef) continue;
      stockIndex[stockRef] = s;
    }

    var out = msInitOutColumns_(shStock, stockHeaderMap, stockNeed, nExisting);

    ss.toast("Calcul updates + ajouts…", "Microstore", 6);

    var toAdd = [];
    var seenInMs = {};

    for (var j = 0; j < refsMs.length; j++) {
      var refKey = refsMs[j];
      var rec2 = msMap[refKey];
      seenInMs[refKey] = true;

      if (stockIndex.hasOwnProperty(refKey)) {
        var idx = stockIndex[refKey];
        msSetOutRow_(out, idx, {
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
          "MS_STATUT": (disabledIndex[rec2.ref] ? "MS_BOTH" : "MS"),
          "MS_LAST_SEEN": nowText
        });
      } else {
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
          "MS_STATUT": (disabledIndex[rec2.ref] ? "MS_BOTH" : "MS"),
          "MS_LAST_SEEN": nowText
        });
      }
    }

    // ajouter les refs présentes uniquement dans MS_IMPORT_DISABLED
    // si elles n'existent ni dans MS_IMPORT ni dans STOCK
    for (var refDis in disabledIndex) {
      if (!disabledIndex.hasOwnProperty(refDis)) continue;

      // déjà traité via MS_IMPORT
      if (seenInMs[refDis]) continue;

      // existe déjà dans STOCK
      if (stockIndex.hasOwnProperty(refDis)) continue;

      var recDis = disabledMap[refDis] || {};
      toAdd.push({
        "货号": refDis,
        "Nom": recDis.nom || "",
        "Catégorie": recDis.categorie || "",
        "Contenu colis": recDis.contenuColis || "",
        "Composition matérielle": recDis.compo || "",
        "Marque": recDis.marque || "",
        "Année": recDis.annee || "",
        "Saison": recDis.saison || "",
        "Colisage": recDis.colisage || "",
        "Couleur": recDis.couleur || "",
        "Stock": (recDis.stock === null || typeof recDis.stock === "undefined" || recDis.stock === "") ? 0 : recDis.stock,
        "Nbr de pièces hors unité de colisage": recDis.horsColisage || "",
        "Poids (en gramme)": recDis.poidsG || "",
        "Prix": recDis.prix || "",
        "Pays d'origine": recDis.paysOrigine || "",
        "Remise (%)": recDis.remise || "",
        "Remarque": recDis.remarqueFinale || recDis.remarque || "",
        "Date de création": recDis.dateCreation || "",
        "MS_STATUT": "MS_DISABLED",
        "MS_LAST_SEEN": nowText
      });
    }

    if (nExisting > 0) {
      ss.toast("Mise à jour statuts (absents MS)…", "Microstore", 5);
      for (var rr = 0; rr < stockRefs.length; rr++) {
        var rref = msNormalizeRef_(stockRefs[rr][0]);
        if (!rref) continue;

        if (disabledIndex[rref] && !seenInMs[rref]) {
          var recDisExisting = disabledMap[rref] || {};
          msSetOutRow_(out, rr, {
            "货号": rref,
            "Nom": recDisExisting.nom || "",
            "Catégorie": recDisExisting.categorie || "",
            "Contenu colis": recDisExisting.contenuColis || "",
            "Composition matérielle": recDisExisting.compo || "",
            "Marque": recDisExisting.marque || "",
            "Année": recDisExisting.annee || "",
            "Saison": recDisExisting.saison || "",
            "Colisage": recDisExisting.colisage || "",
            "Couleur": recDisExisting.couleur || "",
            "Stock": (recDisExisting.stock === null || typeof recDisExisting.stock === "undefined" || recDisExisting.stock === "") ? 0 : recDisExisting.stock,
            "Nbr de pièces hors unité de colisage": recDisExisting.horsColisage || "",
            "Poids (en gramme)": recDisExisting.poidsG || "",
            "Prix": recDisExisting.prix || "",
            "Pays d'origine": recDisExisting.paysOrigine || "",
            "Remise (%)": recDisExisting.remise || "",
            "Remarque": recDisExisting.remarqueFinale || recDisExisting.remarque || "",
            "Date de création": recDisExisting.dateCreation || "",
            "MS_STATUT": "MS_DISABLED",
            "MS_LAST_SEEN": nowText
          });
          continue;
        }

        if (seenInMs[rref]) continue;

        var old = (stockStatus[rr] && stockStatus[rr][0]) ? String(stockStatus[rr][0]).trim() : "";
        var newStatus = (old === "MS" || old === "MS_DISABLED" || old === "MS_BOTH") ? "MS_SUPPRIME" : "A_CREER";
        msSetOutRow_(out, rr, { "MS_STATUT": newStatus });
      }
    }

    var addCount = toAdd.length;
    var addedStartRow = 0;

    if (addCount > 0) {
      ss.toast("Ajout " + addCount + " nouvelles refs…", "Microstore", 6);
      var lastRowBeforeInsert = shStock.getLastRow();
      if (lastRowBeforeInsert < 1) lastRowBeforeInsert = 1;

      shStock.insertRowsAfter(lastRowBeforeInsert, addCount);
      addedStartRow = lastRowBeforeInsert + 1;
      msExtendOutForAdds_(out, toAdd);
    }

    ss.toast("Écriture des colonnes cibles…", "Microstore", 6);
    msWriteOutColumns_(shStock, stockHeaderMap, stockNeed, out);

    if (addCount > 0) {
      ss.toast("Prolongation formules (TEMPLATE_STOCK)…", "Microstore", 8);
      msApplyTemplateFormulas_(shTpl, shStock, tplHeaderMap, stockHeaderMap, formulaCols, addCount, addedStartRow);
      msApplyTemplateColumnByHeaderNoteKey_(shTpl, shStock, "KEY:TOTAL_BOX", addCount, addedStartRow);
      msApplyTemplateColumnByHeaderNoteKey_(shTpl, shStock, "KEY:TOTAL_PCS", addCount, addedStartRow);
    }

    ss.toast("Rebuild filtre (A→AU)…", "Microstore", 5);
    msRebuildLockedFilter_A_to_AU_(shStock);

    ss.toast("Sync OK (MS_IMPORT → STOCK).", "Microstore", 8);
  } catch (err) {
    var msg = (err && err.message) ? err.message : String(err);
    SpreadsheetApp.getActiveSpreadsheet().toast("Sync échoué: " + msg, "Microstore", 10);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function msGetLastImportDateText_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_LOG_IMPORT_EXPORT);
  if (!sh) return "";
  var lr = sh.getLastRow();
  if (lr < 2) return "";
  return sh.getRange(lr, 1).getDisplayValue() || "";
}

function msInitOutColumns_(shStock, stockHeaderMap, stockNeed, nExisting) {
  var out = {};
  for (var i = 0; i < stockNeed.length; i++) {
    var h = stockNeed[i];
    var col = stockHeaderMap[h.toLowerCase()];

    if (nExisting > 0) {
      out[h] = shStock.getRange(2, col, nExisting, 1).getValues();
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

function msSetOutRow_(out, idx0, patchObj) {
  for (var k in patchObj) {
    if (!patchObj.hasOwnProperty(k)) continue;
    var arr = out[k];
    if (!arr) continue;
    while (arr.length <= idx0) arr.push([""]);
    arr[idx0][0] = (patchObj[k] === null || typeof patchObj[k] === "undefined") ? "" : String(patchObj[k]);
  }
}

function msExtendOutForAdds_(out, toAdd) {
  for (var i = 0; i < toAdd.length; i++) {
    var rec = toAdd[i];
    for (var colName in out) {
      if (!out.hasOwnProperty(colName)) continue;
      out[colName].push([""]);
    }
    var newIdx = out["货号"].length - 1;
    msSetOutRow_(out, newIdx, rec);
  }
}

function msWriteOutColumns_(shStock, stockHeaderMap, stockNeed, out) {
  var totalRows = stockNeed.length > 0 ? out[stockNeed[0]].length : 0;
  if (totalRows === 0) return;

  for (var i = 0; i < stockNeed.length; i++) {
    var h = stockNeed[i];
    var col = stockHeaderMap[h.toLowerCase()];
    var data = out[h];
    var range = shStock.getRange(2, col, totalRows, 1);
    range.setValues(data);
    range.setNumberFormat("@");
  }
}

function msApplyTemplateFormulas_(shTpl, shStock, tplHeaderMap, stockHeaderMap, formulaCols, addCount, startRow) {
  if (!addCount || addCount <= 0) return;
  if (!startRow || startRow <= 0) throw new Error("msApplyTemplateFormulas_: startRow invalide");

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

function msRebuildLockedFilter_A_to_AU_(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow < 1) lastRow = 1;

  var existing = sh.getFilter();
  if (existing) existing.remove();

  sh.getRange(1, 1, lastRow, 46).createFilter();
}

function msApplyTemplateColumnByHeaderNoteKey_(shTpl, shStock, key, addCount, startRow) {
  if (!addCount || addCount <= 0) return;
  if (!startRow || startRow <= 0) throw new Error("msApplyTemplateColumnByHeaderNoteKey_: startRow invalide");

  var stockNotes = shStock.getRange(1, 1, 1, shStock.getLastColumn()).getNotes()[0];
  var tplNotes = shTpl.getRange(1, 1, 1, shTpl.getLastColumn()).getNotes()[0];

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
