/**
 * Export STOCK → MS_EXPORT
 */
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

    var stockLastRow = shStock.getLastRow();
    var stockLastCol = shStock.getLastColumn();
    if (stockLastRow < 2) {
      ss.toast("STOCK vide (rien à exporter).", "Microstore", 6);
      msClearMsExportData_(shExp, expLastCol);
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
      "每箱件数2",
      "包/箱",
      "Couleurs"
    ];
    ensureHeadersExist_(stockMap, stockNeed, "STOCK (ligne 1)");

    var data = shStock.getRange(2, 1, stockLastRow - 1, stockLastCol).getValues();

    var ui = SpreadsheetApp.getUi();
    var msg = "⚠️ L'export va remplacer tout le contenu de MS_EXPORT avec les données de STOCK.\n\nContinuer ?";
    var btn = ui.alert("Confirmer export Microstore", msg, ui.ButtonSet.OK_CANCEL);
    if (btn !== ui.Button.OK) {
      ss.toast("Export annulé.", "Microstore", 4);
      return;
    }

    var cRef = stockMap["货号"] - 1;
    var cNom = stockMap["nom"] - 1;
    var cCat = stockMap["catégorie"] - 1;
    var cComp = stockMap["composition matérielle"] - 1;
    var cMarque = stockMap["marque"] - 1;
    var cAnnee = stockMap["année"] - 1;
    var cSaison = stockMap["saison"] - 1;
    var cStock = stockMap["stock"] - 1;
    var cTotalPqs = 46 - 1; // AT = KEY:TOTAL_PQS
    var cColisage = stockMap["colisage"] - 1;
    var cCouleur = stockMap["couleur"] - 1;
    var cPrix = stockMap["prix"] - 1;
    var cRemise = stockMap["remise (%)"] - 1;
    var cDate = stockMap["date de création"] - 1;
    var cTotalPcs = stockMap["每箱件数2"] - 1;
    var cPacks = stockMap["包/箱"] - 1;
    var cHorsColisage = stockMap["nbr de pièces hors unité de colisage"] - 1;
    var cPoidsG = stockMap["poids (en gramme)"] - 1;
    var cPaysOrigine = stockMap["pays d'origine"] - 1;
    var cRemarque = stockMap["remarque"] ? (stockMap["remarque"] - 1) : -1;
    var cMsStatut = stockMap["ms_statut"] ? (stockMap["ms_statut"] - 1) : -1;
    var cCouleursRemark = stockMap["couleurs"] ? (stockMap["couleurs"] - 1) : cCouleur;

    var rows = [];
    for (var i = 0; i < data.length; i++) {
      var r = data[i];

      if (cMsStatut >= 0 && String(r[cMsStatut]).trim() === "MS_SUPPRIME") continue;

      var refRaw = msString_(r[cRef]);
      if (!refRaw) continue;

      var dtInfo = msNormalizeDateForSort_(r[cDate]);
      rows.push({
        sortEmpty: dtInfo.empty,
        sortTs: dtInfo.ts,
        ref: refRaw,
        msStatut: cMsStatut >= 0 ? String(r[cMsStatut] || "").trim() : "",
        nom: msString_(r[cNom]),
        cat: msString_(r[cCat]),
        contenuColis: msString_(r[stockMap["contenu colis"] - 1]),
        comp: normalizeUpper_(r[cComp]),
        marque: msString_(r[cMarque]),
        annee: msString_(r[cAnnee]),
        saison: msString_(r[cSaison]),
        colisage: (r[cColisage] === null || typeof r[cColisage] === "undefined") ? "" : r[cColisage],
        couleur: msNormalizeCouleur_(r[cCouleur]),
        couleursRaw: msString_(r[cCouleursRemark]),
        prix: (r[cPrix] === null || typeof r[cPrix] === "undefined") ? "" : r[cPrix],
        remise: (r[cRemise] === null || typeof r[cRemise] === "undefined") ? "" : r[cRemise],
        totalPcs: r[cTotalPcs],
        stock: r[cTotalPqs],
        packs: r[cPacks],
        horsColisage: r[cHorsColisage],
        poidsG: r[cPoidsG],
        paysOrigine: msString_(r[cPaysOrigine]),
        remarque: cRemarque >= 0 ? msString_(r[cRemarque]) : ""
      });
    }

    var out = [];
    for (var k = 0; k < rows.length; k++) {
      var it = rows[k];

      var total = toIntSafe_(it.totalPcs);
      var packs = toIntSafe_(it.packs);
      var ppp = toIntSafe_(it.colisage);
      var contenu = (total > 0 && packs > 0 && ppp > 0)
        ? msBuildContenuColis_(total, packs, ppp)
        : "";

      var line = new Array(expLastCol).fill("");
      line[expMap["référence"] - 1] = it.ref;
      line[expMap["nom"] - 1] = it.nom ? String(it.nom).trim() : it.ref;
      line[expMap["catégorie"] - 1] = it.cat;
      line[expMap["contenu colis"] - 1] = it.contenuColis;
      line[expMap["composition matérielle"] - 1] = it.comp;
      line[expMap["marque"] - 1] = it.marque;
      line[expMap["année"] - 1] = it.annee;
      line[expMap["saison"] - 1] = it.saison;
      line[expMap["colisage"] - 1] = it.colisage;
      line[expMap["couleur"] - 1] = it.couleur ? it.couleur : "MIX";
      line[expMap["stock"] - 1] = (it.msStatut === "MS_DISABLED")
        ? 0
        : ((it.stock === null || typeof it.stock === "undefined" || String(it.stock).trim() === "" || Number(it.stock) === 0)
          ? Math.floor(Math.random() * 201) + 100
          : it.stock);
      line[expMap["nbr de pièces hors unité de colisage"] - 1] = it.horsColisage;
      line[expMap["poids (en gramme)"] - 1] = it.poidsG;
      line[expMap["prix"] - 1] = it.prix;
      line[expMap["pays d'origine"] - 1] = it.paysOrigine;
      line[expMap["remise (%)"] - 1] = it.remise;
      var couleursTexte = msBuildCouleursRemark_(it.couleursRaw);
      line[expMap["remarque"] - 1] = msBuildCleanRemarque_(it.remarque, contenu, couleursTexte);

      out.push(line);
    }

    msClearMsExportData_(shExp, expLastCol);
    if (out.length) shExp.getRange(3, 1, out.length, expLastCol).setValues(out);

    SpreadsheetApp.flush();
    Utilities.sleep(1200);

    ss.toast("Export terminé: " + out.length + " lignes.", "Microstore", 6);
    exportMsExportSheetToDriveXlsx();
  } finally {
    lock.releaseLock();
  }
}

function msClearMsExportData_(shExp, expLastCol) {
  var lr = shExp.getLastRow();
  if (lr >= 3) shExp.getRange(3, 1, lr - 2, expLastCol).clearContent();
}

function msNormalizeDateForSort_(v) {
  if (v === null || typeof v === "undefined" || v === "") {
    return { empty: true, ts: 0 };
  }
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    return { empty: false, ts: v.getTime() };
  }

  var s = String(v).trim();
  if (!s) return { empty: true, ts: 0 };

  var d = new Date(s);
  if (!isNaN(d.getTime())) return { empty: false, ts: d.getTime() };

  return { empty: true, ts: 0 };
}

function msBuildContenuColis_(totalPieces, packs, pcsPerPack) {
  var t = totalPieces ? String(totalPieces) : "x";
  var p = packs ? String(packs) : "x";
  var u = pcsPerPack ? String(pcsPerPack) : "x";
  return "Colis: " + t + " pièces avec " + p + " paquets de " + u + " pièces";
}

function msBuildCleanRemarque_(existingRemark, contenu, couleursTexte) {
  var lines = String(existingRemark || "").split(/\r?\n/);
  var otherLines = [];
  var seenOther = {};
  var colisLine = "";
  var couleursLine = "";

  for (var i = 0; i < lines.length; i++) {
    var line = String(lines[i] || "").trim();
    if (!line) continue;

    if (/^Colis\s*:/i.test(line)) {
      if (!colisLine) colisLine = line;
      continue;
    }

    if (/^Couleurs\s*:/i.test(line)) {
      if (!couleursLine) couleursLine = line;
      continue;
    }

    if (seenOther[line]) continue;
    seenOther[line] = true;
    otherLines.push(line);
  }

  if (String(contenu || "").trim()) colisLine = String(contenu).trim();
  if (String(couleursTexte || "").trim()) couleursLine = String(couleursTexte).trim();

  var finalLines = otherLines.slice();
  if (colisLine) finalLines.push(colisLine);
  if (couleursLine) finalLines.push(couleursLine);

  return finalLines.join("\n");
}

function msString_(v) {
  if (v === null || typeof v === "undefined") return "";
  return String(v).trim();
}

function msBuildCouleursRemark_(raw) {
  var s = String(raw || "").trim();
  if (!s) return "";

  var re = /(\d+)\s+([A-Za-zÀ-ÿ]+)/g;
  var parts = [];
  var m;
  while ((m = re.exec(s)) !== null) {
    var qty = m[1];
    var color = msTitleCaseColor_(m[2]);
    parts.push(qty + " " + color);
  }

  if (!parts.length) return "";
  return "Couleurs: " + parts.join(", ");
}

function msTitleCaseColor_(raw) {
  var s = String(raw || "").trim();
  if (!s) return "";
  var low = s.toLowerCase();
  return low.charAt(0).toUpperCase() + low.slice(1);
}

/**
 * Export MS_EXPORT → Drive en .xlsx (overwrite)
 */
function exportMsExportSheetToDriveXlsx() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_MS_EXPORT);
  if (!sheet) throw new Error("Feuille introuvable: " + SHEET_MS_EXPORT);

  SpreadsheetApp.flush();
  Utilities.sleep(1200);

  var folder = DriveApp.getFolderById(MS_EXPORT_DRIVE_FOLDER_ID);
  var spreadsheetId = ss.getId();
  var gid = sheet.getSheetId();

  var url =
    "https://docs.google.com/spreadsheets/d/" +
    spreadsheetId +
    "/export?format=xlsx&gid=" +
    gid;

  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + token }
  });

  var blob = response.getBlob().setName(MS_EXPORT_FILENAME);
  var existing = folder.getFilesByName(MS_EXPORT_FILENAME);
  while (existing.hasNext()) {
    existing.next().setTrashed(true);
  }

  var file = folder.createFile(blob);

  SpreadsheetApp.getActive().toast("Export Drive terminé (overwrite)", "MS Export", 5);
  Logger.log("File created: " + file.getUrl());

  return file.getUrl();
}
