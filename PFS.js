/****************************************************
 * EXPORT STOCK → PFS (Paris Fashion Shops)
 *
 * v1:
 * - Exporte uniquement les lignes STOCK cochées via la colonne 选择 (checkbox TRUE)
 * - Ecrit dans la feuille pfs_export (template PFS)
 * - Génère une ligne par couleur (Pack puis Sub-Pack)
 * - Utilise Colisage comme source de vérité pour le nombre de pièces par taille
 * - Utilise Contenu colis uniquement pour détecter les tailles disponibles
 ****************************************************/

var SHEET_PFS_EXPORT = "pfs_export";
var SHEET_PFS_IMPORT = "PFS_IMPORT";
var SHEET_PFS_DIFF = "PFS_DIFF";

// Drive export settings
var PFS_EXPORT_FOLDER_ID = "1CG_X598c8PPIOyCm3uEg-kzuV1BBa4pI";
var PFS_EXPORT_FILENAME = "PFS_EXPORT.xlsx";
var PFS_TEMPLATE_FOLDER_ID = "1_F3f3_24whJAGzj4RuPr5r7FH0cioKLL";
var PFS_TEMPLATE_FILENAME = "Template_PFS.xlsx";

// Debug
var PFS_DEBUG = true;
var PFS_HEADER_MAX_COLS = 80; // enough for PFS template; avoids getLastColumn()=1 when row has trailing blanks

// Minimum export width to cover fixed-position columns up to Pays (U = 21)
var PFS_MIN_EXPORT_COLS = 21;

function importPFSTemplateToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock();
  let tempFileId = "";
  lock.waitLock(30000);

  ss.toast("Import template PFS : démarrage…", "PFS Import", 5);

  try {
    const templateFile = pfsFindNamedFileInFolder_(PFS_TEMPLATE_FOLDER_ID, PFS_TEMPLATE_FILENAME);
    if (!templateFile) {
      throw new Error("Fichier introuvable dans Drive: " + PFS_TEMPLATE_FILENAME);
    }

    ss.toast("Conversion XLSX → Google Sheet temporaire…", "PFS Import", 5);
    tempFileId = pfsConvertXlsxToGoogleSheet_(templateFile.id);

    ss.toast("Ouverture du fichier converti…", "PFS Import", 5);
    const tempSs = pfsOpenSpreadsheetWithRetry_(tempFileId, 8, 1500);
    const sourceSheet = pfsSelectBestTemplateSheet_(tempSs);
    if (!sourceSheet) {
      throw new Error("Le template PFS converti ne contient aucune feuille exploitable.");
    }

    const analysis = pfsAnalyzeSheetStructure_(sourceSheet);
    ss.toast("Remplacement de " + SHEET_PFS_IMPORT + "…", "PFS Import", 5);
    const importedSheet = pfsReplaceSheetFromTemplate_(ss, SHEET_PFS_IMPORT, sourceSheet);

    SpreadsheetApp.flush();

    const summary = pfsBuildTemplateAnalysisSummary_(analysis, {
      fileName: templateFile.name,
      targetSheetName: importedSheet.getName()
    });

    ss.toast("Template PFS importé dans " + SHEET_PFS_IMPORT, "PFS Import", 6);
    notify_("Import PFS terminé", summary);
    Logger.log(summary);
  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    ss.toast("Import PFS échoué: " + msg, "PFS Import", 10);
    throw err;
  } finally {
    if (tempFileId) {
      try { Drive.Files.remove(tempFileId); } catch (e) { Logger.log("Cleanup temp PFS failed: " + e); }
    }
    lock.releaseLock();
  }
}

function compareStockWithPfsImport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stock = ss.getSheetByName(SHEET_STOCK);
  const pfs = ss.getSheetByName(SHEET_PFS_IMPORT);

  if (!stock) throw new Error("Feuille introuvable: " + SHEET_STOCK);
  if (!pfs) throw new Error("Feuille introuvable: " + SHEET_PFS_IMPORT);

  ss.toast("Comparaison STOCK ↔ PFS_IMPORT : analyse…", "PFS Audit", 5);

  const stockLastRow = stock.getLastRow();
  const stockLastCol = stock.getLastColumn();
  if (stockLastRow < 2 || stockLastCol < 1) {
    ss.toast("STOCK vide", "PFS Audit", 5);
    return;
  }

  const stockHeaders = stock.getRange(1, 1, 1, stockLastCol).getValues()[0];
  const stockMap = headerMap_(stockHeaders);
  ensureHeadersExist_(stockMap, ["货号"], "STOCK");

  const pfsAnalysis = pfsAnalyzeSheetStructure_(pfs);
  if (!pfsAnalysis.headerRow || !pfsAnalysis.usefulWidth) {
    throw new Error("Impossible de détecter les headers de " + SHEET_PFS_IMPORT + ".");
  }

  const pfsHeaders = pfs.getRange(pfsAnalysis.headerRow, 1, 1, pfsAnalysis.usefulWidth).getValues()[0];
  const pfsDataStartRow = pfsAnalysis.headerRow + 1;
  const pfsRowCount = Math.max(0, pfsAnalysis.usefulRows - pfsAnalysis.headerRow);
  if (!pfsRowCount) {
    ss.toast("PFS_IMPORT ne contient aucune ligne à comparer", "PFS Audit", 6);
    return;
  }

  const pfsSkuCol = pfsFindColumnByAliases_(pfsHeaders, ["sku", "sku parent", "reference sku"]);
  const pfsTypeVenteCol = pfsFindColumnByAliases_(pfsHeaders, ["type vente", "type de vente"]);
  const stockRefCol = stockMap["货号"] || 0;

  if (!pfsSkuCol) throw new Error("Colonne 'SKU' introuvable dans " + SHEET_PFS_IMPORT + ".");

  const stockValues = stock.getRange(2, 1, stockLastRow - 1, stockLastCol).getValues();
  const pfsValues = pfs.getRange(pfsDataStartRow, 1, pfsRowCount, pfsAnalysis.usefulWidth).getValues();

  const stockByRef = {};
  const duplicateStockRefs = [];
  for (let i = 0; i < stockValues.length; i++) {
    const ref = pfsNormalizeRef_(stockValues[i][stockRefCol - 1]);
    if (!ref) continue;
    if (stockByRef[ref]) {
      duplicateStockRefs.push(ref);
      continue;
    }
    stockByRef[ref] = stockValues[i];
  }

  const stockRefs = Object.keys(stockByRef);
  if (!stockRefs.length) {
    ss.toast("Aucune ref exploitable trouvée dans STOCK", "PFS Audit", 6);
    return;
  }

  const pfsRowsByRef = {};
  const seenPfsRefs = {};
  for (let i = 0; i < pfsValues.length; i++) {
    const sku = pfsNormalizeScalar_(pfsValues[i][pfsSkuCol - 1]);
    const skuBaseRef = pfsExtractRefFromSku_(sku);
    if (skuBaseRef) seenPfsRefs[skuBaseRef] = true;

    const ref = pfsFindMatchingStockRefForSku_(sku, stockRefs);
    if (!ref) continue;
    if (!pfsRowsByRef[ref]) pfsRowsByRef[ref] = [];
    pfsRowsByRef[ref].push(i);
  }

  const compareDefs = [
    {
      key: "prix",
      label: "Prix",
      stockCol: stockMap["prix@"] || 0,
      pfsCol: pfsFindColumnByAliases_(pfsHeaders, ["prix_vente gros ht unit. eur", "prix vente gros ht unit. eur", "prix_vente gros", "prix vente gros"]),
      type: "number",
      stockValue: function(row) { return row[stockMap["prix@"] - 1]; }
    },
    {
      key: "promo",
      label: "Promo",
      stockCol: stockMap["promo@"] || 0,
      pfsCol: pfsFindColumnByAliases_(pfsHeaders, ["prix_vente réduit ht unit. eur", "prix vente réduit ht unit. eur", "prix_vente reduit ht unit. eur", "prix vente reduit ht unit. eur", "prix réduit", "prix reduit"]),
      type: "number",
      stockValue: function(row) { return row[stockMap["promo@"] - 1]; }
    },
    {
      key: "stock",
      label: "Stock",
      stockCol: stockMap["stock"] || 0,
      pfsCol: pfsFindColumnByAliases_(pfsHeaders, ["quantité total stock pcs", "quantite total stock pcs", "stock pcs", "stock"]),
      type: "number",
      stockValue: function(row) { return row[stockMap["stock"] - 1]; }
    },
    {
      key: "poids",
      label: "Poids",
      stockCol: stockMap["poids (en gramme)"] || 0,
      pfsCol: pfsFindColumnByAliases_(pfsHeaders, ["poids_kg / pc", "poids kg / pc", "poids_kg/pc", "poids kg/pc", "poids_kg", "poids kg"]),
      type: "number",
      stockValue: function(row) { return pfsFormatKgValueFromGrams_(row[stockMap["poids (en gramme)"] - 1]); }
    },
    {
      key: "active",
      label: "Active",
      stockCol: stockMap["ms_statut"] || 0,
      pfsCol: pfsFindColumnByAliases_(pfsHeaders, ["active", "actif"]),
      type: "yesno",
      stockValue: function(row) { return pfsMapActiveFromMsStatus_(row[stockMap["ms_statut"] - 1]); }
    }
  ];

  const availableDefs = compareDefs.filter(function(def) {
    return def.stockCol && def.pfsCol;
  });
  const missingDefs = compareDefs
    .filter(function(def) { return !def.stockCol || !def.pfsCol; })
    .map(function(def) { return def.label; });

  if (!availableDefs.length) {
    throw new Error("Aucune colonne comparable trouvée entre STOCK et " + SHEET_PFS_IMPORT + ".");
  }

  const diffRows = [];
  const matchedRefs = [];
  const absentPfsRefs = [];
  const pfsRefsWithoutStock = [];
  const nonComparableWarnings = [];

  for (let i = 0; i < stockRefs.length; i++) {
    const ref = stockRefs[i];
    const stockRow = stockByRef[ref];
    const targetRows = pfsRowsByRef[ref];

    if (!targetRows || !targetRows.length) {
      absentPfsRefs.push(ref);
      diffRows.push([
        ref,
        "",
        "",
        "REF",
        ref,
        "",
        "ABSENT_PFS",
        "Aucune ligne PFS trouvée pour cette ref"
      ]);
      continue;
    }

    matchedRefs.push(ref);

    for (let t = 0; t < targetRows.length; t++) {
      const targetIndex = targetRows[t];
      const pfsRow = pfsValues[targetIndex];
      const sku = pfsNormalizeScalar_(pfsRow[pfsSkuCol - 1]);
      const typeVente = pfsTypeVenteCol ? pfsNormalizeScalar_(pfsRow[pfsTypeVenteCol - 1]) : "";

      for (let d = 0; d < availableDefs.length; d++) {
        const def = availableDefs[d];
        const stockPrepared = pfsPrepareComparableValue_(def.stockValue(stockRow), def.type);
        const pfsPrepared = pfsPrepareComparableValue_(pfsRow[def.pfsCol - 1], def.type);

        if (!stockPrepared.comparable || !pfsPrepared.comparable) {
          const comment = stockPrepared.reason || pfsPrepared.reason || "Valeur non comparable";
          diffRows.push([
            ref,
            sku,
            typeVente,
            def.label,
            stockPrepared.display,
            pfsPrepared.display,
            "NON_COMPARE",
            comment
          ]);
          nonComparableWarnings.push(ref + " / " + def.label + ": " + comment);
          continue;
        }

        if (pfsComparableValuesEqual_(stockPrepared, pfsPrepared, def.type)) continue;

        diffRows.push([
          ref,
          sku,
          typeVente,
          def.label,
          stockPrepared.display,
          pfsPrepared.display,
          "DIFF",
          ""
        ]);
      }
    }
  }

  Object.keys(seenPfsRefs).forEach(function(ref) {
    if (stockByRef[ref]) return;
    pfsRefsWithoutStock.push(ref);
    const rows = pfsValuesByMatchedBaseRef_(pfsValues, pfsSkuCol, ref);
    for (let i = 0; i < rows.length; i++) {
      const pfsRow = rows[i];
      const sku = pfsNormalizeScalar_(pfsRow[pfsSkuCol - 1]);
      const typeVente = pfsTypeVenteCol ? pfsNormalizeScalar_(pfsRow[pfsTypeVenteCol - 1]) : "";
      diffRows.push([
        ref,
        sku,
        typeVente,
        "REF",
        "",
        ref,
        "ABSENT_STOCK",
        "Ref présente dans PFS_IMPORT mais absente de STOCK"
      ]);
    }
  });

  const diffSheet = pfsRecreateSheet_(ss, SHEET_PFS_DIFF);
  const header = [["Ref", "SKU", "Type vente", "Champ", "Valeur STOCK", "Valeur PFS", "Statut", "Commentaire"]];
  diffSheet.getRange(1, 1, 1, header[0].length).setValues(header);
  diffSheet.getRange(1, 1, 1, header[0].length).setFontWeight("bold");
  diffSheet.setFrozenRows(1);

  if (diffRows.length) {
    diffSheet.getRange(2, 1, diffRows.length, header[0].length).setValues(diffRows);
  }

  for (let c = 1; c <= header[0].length; c++) {
    diffSheet.autoResizeColumn(c);
  }

  const notes = [];
  if (missingDefs.length) notes.push("Colonnes non comparées: " + missingDefs.join(", "));
  if (duplicateStockRefs.length) notes.push("Refs dupliquées dans STOCK ignorées après la première occurrence: " + pfsUniqueList_(duplicateStockRefs).slice(0, 10).join(", "));
  if (absentPfsRefs.length) notes.push("Refs absentes de PFS (aperçu): " + absentPfsRefs.slice(0, 15).join(", ") + (absentPfsRefs.length > 15 ? " …" : ""));
  if (pfsRefsWithoutStock.length) notes.push("Bonus - refs PFS absentes de STOCK (aperçu): " + pfsRefsWithoutStock.slice(0, 15).join(", ") + (pfsRefsWithoutStock.length > 15 ? " …" : ""));
  if (nonComparableWarnings.length) notes.push("Valeurs non comparables (aperçu): " + pfsUniqueList_(nonComparableWarnings).slice(0, 8).join(" | ") + (nonComparableWarnings.length > 8 ? " …" : ""));

  const summary = [
    "Comparaison STOCK ↔ " + SHEET_PFS_IMPORT,
    "",
    "Refs STOCK analysées : " + stockRefs.length,
    "Refs matchées PFS : " + matchedRefs.length,
    "Refs absentes de PFS : " + absentPfsRefs.length,
    "Lignes avec écarts : " + diffRows.length
  ].concat(notes).join("\n");

  if (absentPfsRefs.length) {
    Logger.log("Refs STOCK absentes de PFS (" + absentPfsRefs.length + "):\n" + absentPfsRefs.join("\n"));
  }
  if (pfsRefsWithoutStock.length) {
    Logger.log("Refs PFS absentes de STOCK (" + pfsRefsWithoutStock.length + "):\n" + pfsRefsWithoutStock.join("\n"));
  }
  if (nonComparableWarnings.length) {
    Logger.log("Valeurs non comparables (" + nonComparableWarnings.length + "):\n" + pfsUniqueList_(nonComparableWarnings).join("\n"));
  }
  Logger.log(summary);

  ss.toast("Comparaison PFS terminée", "PFS Audit", 6);
  notify_("Comparaison PFS terminée", summary);
  return {
    analyzedRefs: stockRefs.length,
    matchedRefs: matchedRefs.length,
    absentPfsRefs: absentPfsRefs,
    pfsRefsWithoutStock: pfsRefsWithoutStock,
    diffCount: diffRows.length
  };
}

function exportStockToPFS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stock = ss.getSheetByName(SHEET_STOCK);
  const exp = ss.getSheetByName(SHEET_PFS_EXPORT);

  if (!stock) throw new Error("STOCK introuvable");
  if (!exp) throw new Error("pfs_export introuvable");

  // --- Map STOCK headers
  const stockHeaders = stock.getRange(1, 1, 1, stock.getLastColumn()).getValues()[0];
  const stockMap = headerMap_(stockHeaders);

  // Required for v0 export
  ensureHeadersExist_(stockMap, [
    "选择",
    "货号",
    "Prix@",
    "Catégorie",
    "Contenu colis",
    "Poids (en gramme)",
    "Pays d'origine",
    "Couleurs",
    "Colisage"
  ], "STOCK");

  const lastRow = stock.getLastRow();
  if (lastRow < 2) {
    ss.toast("STOCK vide", "PFS", 4);
    return;
  }

  // Read only what we need (fast)
  const n = lastRow - 1;
  const colSel = stockMap["选择"];
  const colRef = stockMap["货号"];
  const colPrix = stockMap["prix@"]; // headerMap_ lowercases keys
  const colCat = stockMap["catégorie"];
  const colContenu = stockMap["contenu colis"];
  if (!colContenu) {
    throw new Error("PFS export: colonne 'Contenu colis' introuvable dans STOCK.");
  }

  const colCompo = stockMap["composition matérielle"] || 0;
  const colPoids = stockMap["poids (en gramme)"];
  const colPays = stockMap["pays d'origine"];
  const colCouleurs = stockMap["couleurs"];
  const colColisage = stockMap["colisage"];

  if (!colPoids) throw new Error("PFS export: colonne 'Poids (en gramme)' introuvable dans STOCK.");
  if (!colPays) throw new Error("PFS export: colonne 'Pays d'origine' introuvable dans STOCK.");

  const selVals = stock.getRange(2, colSel, n, 1).getValues().flat();
  const refVals = stock.getRange(2, colRef, n, 1).getValues().flat();
  const prixVals = stock.getRange(2, colPrix, n, 1).getValues().flat();
  const catVals = stock.getRange(2, colCat, n, 1).getValues().flat();
  const contenuVals = stock.getRange(2, colContenu, n, 1).getValues().flat();
  const compoVals = colCompo ? stock.getRange(2, colCompo, n, 1).getValues().flat() : null;
  const poidsVals = stock.getRange(2, colPoids, n, 1).getValues().flat();
  const paysVals = stock.getRange(2, colPays, n, 1).getValues().flat();
  const couleursVals = stock.getRange(2, colCouleurs, n, 1).getValues().flat();
  const colisageVals = stock.getRange(2, colColisage, n, 1).getValues().flat();

  // Collect only selected rows
  const selected = [];
  const bad = []; // {row, ref, reason}

  for (let i = 0; i < n; i++) {
    if (selVals[i] !== true) continue; // checkbox TRUE only

    const sheetRow = 2 + i;
    const ref = String(refVals[i] || "").trim();
    const prix = String(prixVals[i] ?? "").trim();
    const colisage = parseColisagePFS_(colisageVals[i]);
    const tailles = colisage.ok
      ? extractTaillesFromContenuColis_(contenuVals[i], colisage.value)
      : "";
    const poids = String(poidsVals[i] ?? "").trim();
    const paysFab = String(paysVals[i] ?? "").trim();
    const couleursRaw = String(couleursVals[i] ?? "").trim();
    const taillesInfo = tailles ? parsePfsTaillesStructure_(tailles) : { ok: false, reason: "Tailles introuvables (Contenu colis)" };
    const colorPack = taillesInfo.ok
      ? parsePfsColorPackFromStock_(couleursRaw, taillesInfo.sizes, colisage.ok ? colisage.value : 0)
      : { ok: false, reason: taillesInfo.reason };

    // Category normalization and validation
    const catRaw = String(catVals[i] || "").trim();
    const catPfs = normalizeCategoriePFS_(catRaw);
    if (!catPfs) {
      bad.push({ row: sheetRow, ref: ref || "(vide)", reason: "Catégorie PFS invalide: " + (catRaw || "(vide)") });
      continue;
    }

    if (!colisage.ok) {
      bad.push({ row: sheetRow, ref: ref || "(vide)", reason: colisage.reason });
      continue;
    }

    if (!ref) {
      bad.push({ row: sheetRow, ref: "(vide)", reason: "货号 vide" });
      continue;
    }
    if (!prix) {
      bad.push({ row: sheetRow, ref: ref, reason: "Prix@ vide" });
      continue;
    }
    if (!tailles) {
      bad.push({ row: sheetRow, ref: ref, reason: "Tailles introuvables (Contenu colis)" });
      continue;
    }
    if (!poids) {
      bad.push({ row: sheetRow, ref: ref || "(vide)", reason: "Poids (en gramme) vide" });
      continue;
    }
    if (!paysFab) {
      bad.push({ row: sheetRow, ref: ref || "(vide)", reason: "Pays d’origine vide" });
      continue;
    }
    if (!taillesInfo.ok) {
      bad.push({ row: sheetRow, ref: ref || "(vide)", reason: "Structure tailles invalide: " + taillesInfo.reason });
      continue;
    }
    if (!colorPack.ok) {
      bad.push({ row: sheetRow, ref: ref || "(vide)", reason: colorPack.reason });
      continue;
    }

    selected.push({
      row: sheetRow,
      ref: ref,
      prix: prix,
      cat: catPfs,
      catRaw: catRaw,
      tailles: tailles,
      compo: compoVals ? String(compoVals[i] || "").trim() : "",
      poids: poids,
      paysFab: paysFab,
      taillesInfo: taillesInfo,
      colorPack: colorPack,
      colisage: colisage.value,
    });
  }

  // If invalid selected rows exist, fail loudly (so tests are deterministic)
  if (bad.length) {
    const details = bad
      .slice(0, 10)
      .map(function (x) {
        return [
          "• Ligne " + x.row,
          "Ref : " + x.ref,
          "Motif : " + appendValidationHint_(x.reason)
        ].join("\n");
      })
      .join("\n\n");

    const more = bad.length > 10
      ? "\n\n(+" + (bad.length - 10) + " autres)"
      : "";

    const message = [
      "PFS export",
      "",
      "Erreurs détectées dans STOCK :",
      "",
      details + more
    ].join("\n");

    SpreadsheetApp.getUi().alert(message);
    throw new Error(message);
  }

  // Nothing selected
  if (!selected.length) {
    // Clear previous export (so user knows it's empty)
    const lr = exp.getLastRow();
    const hinfo = readPfsHeaders_(exp);
    if (lr > 1) exp.getRange(2, 1, lr - 1, hinfo.width).clearContent();

    ss.toast("Aucune ligne cochée (选择=TRUE)", "PFS", 5);
    return;
  }

  // --- Map PFS headers (robust)
  const hinfo = readPfsHeaders_(exp);
  const headers = hinfo.headers;
  const col = buildPfsColumnIndex_(headers);

  // Header-based lookup (preferred)
  let cMarque = col("marque");
  let cGenre = col("genre");
  let cFamille = col("famille");
  let cCategorie = col("catégorie");
  let cRef = col("réf. produit");
  let cNomFr = col("nom du produit* (fr)");
  let cTypeVente = col("type vente");
  let cSaison = col("saison");
  let cTailles = col("tailles");
  let cCouleurs = col("couleurs");
  let cPrix = col("prix_vente gros");
  let cPoids = col("poids_kg");
  let cCompo = col("composition matière");
  let cPays = col("pays de fabrication");

  // Fallback to fixed PFS template positions (A..AA) if headers are missing/empty
  if (!cMarque) cMarque = 1;      // A
  if (!cGenre) cGenre = 2;        // B
  if (!cFamille) cFamille = 3;    // C
  if (!cCategorie) cCategorie = 4;// D
  if (!cRef) cRef = 5;            // E
  if (!cNomFr) cNomFr = 6;        // F
  if (!cTypeVente) cTypeVente = 11; // K
  if (!cSaison) cSaison = 12;     // L
  if (!cTailles) cTailles = 13;   // M
  if (!cCouleurs) cCouleurs = 14; // N
  if (!cPrix) cPrix = 15;         // O
  if (!cPoids) cPoids = 18;       // R
  if (!cCompo) cCompo = 19;       // S
  if (!cPays) cPays = 21;         // U

  if (PFS_DEBUG) {
    Logger.log("PFS header width=" + hinfo.width);
    Logger.log("PFS col indices: marque=" + cMarque + ", genre=" + cGenre + ", famille=" + cFamille + ", categorie=" + cCategorie + ", ref=" + cRef + ", nomFR=" + cNomFr + ", type=" + cTypeVente + ", saison=" + cSaison + ", tailles=" + cTailles + ", couleurs=" + cCouleurs + ", prix=" + cPrix + ", poids=" + cPoids + ", compo=" + cCompo + ", pays=" + cPays);
    Logger.log("PFS headers preview: " + headers.map(h => String(h||"").slice(0,40)).join(" | "));
  }

  const out = [];

  for (const it of selected) {
    const cat = it.cat || ""; // PFS normalized category
    const nomFr = String(it.catRaw || "").trim() || cat || it.ref; // use STOCK category for FR name
    const globalTailles = it.taillesInfo || parsePfsTaillesStructure_(it.tailles);
    const compo = normalizeCompositionPFS_(it.compo) || "90% Viscose - 10% Polyester";

    if (!globalTailles.ok) {
      throw new Error("PFS export: structure tailles invalide pour " + it.ref + ": " + globalTailles.reason);
    }
    if (globalTailles.total !== it.colisage) {
      throw new Error("PFS export: tailles incohérentes avec le colisage pour " + it.ref + ": " + globalTailles.total + " au lieu de " + it.colisage);
    }

    const entries = Array.isArray(it.colorPack && it.colorPack.entries) ? it.colorPack.entries : [];
    for (let idx = 0; idx < entries.length; idx++) {
      const entry = entries[idx];
      const taillesColor = it.colorPack && it.colorPack.mode === "detailed"
        ? buildPfsDetailedColorTailles_(entry.split, globalTailles.sizes)
        : buildPfsColorTailles_(entry.qty, globalTailles.sizes);
      if (!taillesColor.ok) {
        throw new Error("PFS export: tailles/couleurs incompatibles pour " + it.ref + " / " + entry.color + ": " + taillesColor.reason);
      }

      const line = new Array(hinfo.width).fill("");

      // fixed values
      if (cMarque) line[cMarque - 1] = "S.Z FASHION";
      if (cGenre) line[cGenre - 1] = "Femme";
      if (cFamille) line[cFamille - 1] = "Vêtements";

      if (cCategorie) line[cCategorie - 1] = cat;
      if (cRef) line[cRef - 1] = it.ref;
      if (cNomFr) line[cNomFr - 1] = nomFr;

      if (cTypeVente) line[cTypeVente - 1] = idx === 0 ? "Pack" : "Sub-Pack";
      if (cSaison) line[cSaison - 1] = "Printemps/Été 2026";
      if (cTailles) line[cTailles - 1] = taillesColor.value;
      if (cCouleurs) line[cCouleurs - 1] = entry.color;
      if (cPrix) line[cPrix - 1] = String(it.prix ?? "").trim();
      if (cPoids) line[cPoids - 1] = formatWeightKgPFSText_(it.poids);
      if (cCompo) line[cCompo - 1] = compo;
      if (cPays) line[cPays - 1] = it.paysFab;

      out.push(line);
    }
  }

  // Overwrite output
  const lr = exp.getLastRow();
  if (lr > 1) {
    exp.getRange(2, 1, lr - 1, hinfo.width).clearContent();
  }

  // Pre-format as TEXT before writing, to avoid Sheets coercion
  exp.getRange(2, 1, out.length, hinfo.width).setNumberFormat("@");
  exp.getRange(2, 1, out.length, hinfo.width).setValues(out);

  // Ensure data is exported with sane formats (avoid date coercion like 6.8 => 06/08/2026)
  applyPfsDataFormats_(exp, out.length, hinfo.width, { cPrix: cPrix, cPoids: cPoids });

  // Auto-uncheck: décoche 选择 pour les lignes exportées
  try {
    const ranges = selected.map(it => stock.getRange(it.row, colSel));
    const rl = stock.getRangeList(ranges.map(r => r.getA1Notation()));
    // RangeList supports setValue in modern Apps Script
    rl.setValue(false);
  } catch (e) {
    // Fallback: few selections -> loop is fine
    for (const it of selected) {
      stock.getRange(it.row, colSel).setValue(false);
    }
  }

  // Auto-export Drive after sheet generation
  try {
    SpreadsheetApp.flush();
    Utilities.sleep(1200);
    exportPFSToDriveXlsx();
  } catch (e) {
    Logger.log("Auto export Drive PFS failed: " + e);
  }

  // Also generate photo helper files in the export Drive folder
  try {
    exportPfsPhotoDupHelperToDrive_(selected);
  } catch (e) {
    Logger.log("PFS photo helper export failed: " + e);
  }

  ss.toast("PFS_EXPORT généré : " + out.length + " lignes", "PFS", 6);
}

/****************************************************
 * Extract tailles from Contenu colis text, using Colisage as source of truth for qty per size.
 * Input: v = Contenu colis, colisage = integer
 * Output: normalized PFS tailles string or ""
 ****************************************************/
function extractTaillesFromContenuColis_(v, colisage) {
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
    const parsed = parseContenuColisSegmentPFS_(segments[i]);
    if (!parsed || !parsed.size) return "";
    parts.push(qtyPerSize + "*" + parsed.size);
  }

  return parts.join(",");
}

/**
 * Normalisation tailles -> format PFS, using Colisage as source of truth for qty per size.
 * Ex:
 * - "6 x M/L - 6 x XL/XXL" => "6*M/L,6*XL/XXL"
 * - "2*M/L, 2*XL/XXL" => "2*M/L,2*XL/XXL"
 */
function normalizeTaillesPFS_(v, colisage) {
  const raw = String(v || "").trim();
  if (!raw) return "";

  // Parse with the same robust logic as extractTaillesFromContenuColis_
  const extracted = extractTaillesFromContenuColis_(raw, colisage);
  if (extracted) return extracted;

  // Fallback: normalize only spaces/commas, but NEVER break slash-based sizes
  return raw
    .replace(/\u00A0/g, " ")
    .replace(/[／⁄∕]/g, "/")
    .replace(/[–—−]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/\s*,\s*/g, ",")
    .trim();
}
function parseContenuColisSegmentPFS_(segment) {
  const s = String(segment || "").trim();
  if (!s) return null;

  // Examples:
  // "6 x S/M" -> 6*S/M
  // "6 x XL/XXL" -> 6*XL/XXL
  // "S/M" -> keep size token and default qty 6 for current PFS v1
  // "XL/XXL" -> keep size token and default qty 6

  const withQty = s.match(/^(\d+)\s*[xX×*]?\s*(.+)$/);
  if (withQty) {
    const qty = Number(withQty[1]);
    const size = normalizePfsSizeToken_(withQty[2]);
    if (!qty || !size) return null;
    return { qty: qty, size: size };
  }

  const sizeOnly = normalizePfsSizeToken_(s);
  if (sizeOnly) {
    return { qty: 6, size: sizeOnly };
  }

  return null;
}

function normalizePfsSizeToken_(raw) {
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

  // Allow simple single tokens if needed
  if (/^[A-Z0-9]+$/.test(s)) return s;

  return "";
}

/****************************************************
 * Export pfs_export → Google Drive (XLSX)
 * overwrite activé
 ****************************************************/
function exportPFSToDriveXlsx() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_PFS_EXPORT);

  if (!sheet) throw new Error("Feuille introuvable: " + SHEET_PFS_EXPORT);

  const folder = DriveApp.getFolderById(PFS_EXPORT_FOLDER_ID);

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

  const blob = response.getBlob().setName(PFS_EXPORT_FILENAME);

  // overwrite: supprimer anciens fichiers
  const existing = folder.getFilesByName(PFS_EXPORT_FILENAME);
  while (existing.hasNext()) {
    existing.next().setTrashed(true);
  }

  // créer nouveau fichier
  const file = folder.createFile(blob);

  SpreadsheetApp.getActive().toast(
    "Export PFS Drive terminé",
    "PFS",
    5
  );

  Logger.log("File created: " + file.getUrl());
  return file.getUrl();
}

/**
 * Read PFS_EXPORT headers robustly
 * - Doesn't rely on getLastColumn() which may be 1 if only A1 is filled
 */
function readPfsHeaders_(sheet) {
  const maxCols = sheet.getMaxColumns();
  const width = Math.min(maxCols, PFS_HEADER_MAX_COLS);
  const row = sheet.getRange(1, 1, 1, width).getValues()[0];

  // Find last non-empty header cell
  let last = 0;
  for (let i = row.length - 1; i >= 0; i--) {
    if (String(row[i] || "").trim() !== "") { last = i + 1; break; }
  }

  if (last === 0) {
    throw new Error("PFS export: la ligne 1 de 'pfs_export' ne contient aucun header.");
  }

  // Ensure we can write fixed-position columns even if headers are missing/blank past a point
  last = Math.max(last, PFS_MIN_EXPORT_COLS);

  // Keep at least until last header (or enforced minimum)
  const headers = row.slice(0, last);

  if (PFS_DEBUG) {
    Logger.log("readPfsHeaders_: maxCols=" + maxCols + ", scanWidth=" + width + ", lastHeaderCol(enforced)=" + last + ", min=" + PFS_MIN_EXPORT_COLS);
  }

  return { headers: headers, width: last };
}

/****************************************************
 * Formatting helpers to keep PFS_EXPORT XLSX readable
 ****************************************************/
function applyPfsDataFormats_(sheet, numRows, width, cols) {
  if (!numRows || !width) return;

  try {
    const rng = sheet.getRange(2, 1, numRows, width);

    // Force TEXT for all exported cells.
    // This prevents Excel from re-interpreting values like 6.8 as dates on open.
    rng.setNumberFormat("@");

    // Also explicitly force TEXT on key columns (some templates carry weird formats)
    if (cols && cols.cPrix) sheet.getRange(2, cols.cPrix, numRows, 1).setNumberFormat("@");
    if (cols && cols.cPoids) sheet.getRange(2, cols.cPoids, numRows, 1).setNumberFormat("@");
  } catch (e) {
    if (typeof Logger !== "undefined") Logger.log("applyPfsDataFormats_ error: " + e);
  }
}

/****************************************************
 * PFS header utilities (robust to quotes/newlines)
 ****************************************************/
function normalizeHeaderKey_(s) {
  return String(s || "")
    .replace(/^\s*"|"\s*$/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function buildPfsColumnIndex_(headers) {
  const keys = headers.map(h => normalizeHeaderKey_(h));
  return function findCol(needle) {
    const n = normalizeHeaderKey_(needle);
    if (!n) return 0;

    // prefer exact match
    let idx = keys.indexOf(n);
    if (idx >= 0) return idx + 1;

    // fallback: includes match
    for (let i = 0; i < keys.length; i++) {
      if (keys[i] && keys[i].indexOf(n) !== -1) return i + 1;
    }
    return 0;
  };
}

function pfsFindColumnByAliases_(headers, aliases) {
  const findCol = buildPfsColumnIndex_(headers || []);
  for (let i = 0; i < aliases.length; i++) {
    const col = findCol(aliases[i]);
    if (col) return col;
  }
  return 0;
}

function pfsNormalizeRef_(v) {
  return (typeof cleanRef_ === "function" ? cleanRef_(v) : String(v || "").trim()).toUpperCase();
}

function pfsNormalizeScalar_(v) {
  if (v === null || typeof v === "undefined") return "";
  return String(v).trim();
}

function pfsExtractRefFromSku_(sku) {
  const raw = pfsNormalizeRef_(sku);
  if (!raw) return "";
  const pos = raw.indexOf("_");
  return pos >= 0 ? raw.slice(0, pos) : raw;
}

function pfsSkuMatchesRef_(sku, ref) {
  const skuNorm = pfsNormalizeRef_(sku);
  const refNorm = pfsNormalizeRef_(ref);
  if (!skuNorm || !refNorm) return false;
  return skuNorm === refNorm || skuNorm.indexOf(refNorm + "_") === 0;
}

function pfsFindMatchingStockRefForSku_(sku, refs) {
  const stockRefs = Array.isArray(refs) ? refs : [];
  if (!stockRefs.length) return "";

  const direct = pfsExtractRefFromSku_(sku);
  if (direct && stockRefs.indexOf(direct) !== -1 && pfsSkuMatchesRef_(sku, direct)) {
    return direct;
  }

  for (let i = 0; i < stockRefs.length; i++) {
    if (pfsSkuMatchesRef_(sku, stockRefs[i])) return stockRefs[i];
  }
  return "";
}

function pfsResolveSyncValue_(value) {
  if (value && typeof value === "object" && Object.prototype.hasOwnProperty.call(value, "write")) {
    return {
      write: value.write === true,
      value: pfsNormalizeScalar_(value.value),
      reason: value.reason ? String(value.reason) : ""
    };
  }
  return {
    write: true,
    value: pfsNormalizeScalar_(value),
    reason: ""
  };
}

function pfsFormatKgValueFromGrams_(grams) {
  if (grams === null || typeof grams === "undefined" || String(grams).trim() === "") return "";
  return formatWeightKgPFSText_(grams);
}

function pfsMapActiveFromMsStatus_(status) {
  const raw = pfsNormalizeScalar_(status).toUpperCase();
  if (!raw) {
    return { write: true, value: "" };
  }
  if (raw === "MS") {
    return { write: true, value: "Oui" };
  }
  if (raw === "MS_DISABLED") {
    return { write: true, value: "Non" };
  }
  return {
    write: false,
    value: "",
    reason: "MS_STATUT non géré (" + raw + ")"
  };
}

function pfsPrepareComparableValue_(value, compareType) {
  const resolved = pfsResolveSyncValue_(value);
  if (!resolved.write) {
    return {
      comparable: false,
      value: "",
      display: pfsNormalizeScalar_(resolved.value),
      reason: resolved.reason || "Valeur non comparable"
    };
  }

  const raw = pfsNormalizeScalar_(resolved.value);
  if (compareType === "number") {
    if (raw === "") return { comparable: true, value: "", display: "" };
    const num = pfsToComparableNumber_(raw);
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

  if (compareType === "yesno") {
    if (raw === "") return { comparable: true, value: "", display: "" };
    const low = raw.toLowerCase();
    if (low === "oui") return { comparable: true, value: "oui", display: raw };
    if (low === "non") return { comparable: true, value: "non", display: raw };
    return {
      comparable: false,
      value: raw,
      display: raw,
      reason: "Valeur Oui/Non invalide (" + raw + ")"
    };
  }

  return {
    comparable: true,
    value: raw,
    display: raw
  };
}

function pfsComparableValuesEqual_(left, right, compareType) {
  if (compareType === "number") {
    if (left.value === "" && right.value === "") return true;
    if (typeof left.value !== "number" || typeof right.value !== "number") return false;
    return Math.abs(left.value - right.value) < 1e-9;
  }
  return String(left.value || "") === String(right.value || "");
}

function pfsToComparableNumber_(value) {
  const raw = pfsNormalizeScalar_(value);
  if (!raw) return null;
  const normalized = raw.replace(/\s+/g, "").replace(",", ".");
  if (!/^-?\d+(?:\.\d+)?$/.test(normalized)) return null;
  const num = Number(normalized);
  return Number.isFinite(num) ? num : null;
}

function pfsRecreateSheet_(spreadsheet, sheetName) {
  const existing = spreadsheet.getSheetByName(sheetName);
  let index = spreadsheet.getSheets().length + 1;

  if (existing) {
    index = existing.getIndex();
    spreadsheet.deleteSheet(existing);
  }

  const safeIndex = Math.max(1, Math.min(index, spreadsheet.getSheets().length + 1));
  return spreadsheet.insertSheet(sheetName, safeIndex);
}

function pfsValuesByMatchedBaseRef_(rows, skuCol, ref) {
  const out = [];
  const targetRef = pfsNormalizeRef_(ref);
  for (let i = 0; i < rows.length; i++) {
    const sku = pfsNormalizeScalar_(rows[i][skuCol - 1]);
    if (!pfsSkuMatchesRef_(sku, targetRef)) continue;
    out.push(rows[i]);
  }
  return out;
}

function pfsFindNamedFileInFolder_(folderId, fileName) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByName(fileName);
  let best = null;

  while (files.hasNext()) {
    const file = files.next();
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

function pfsConvertXlsxToGoogleSheet_(xlsxFileId) {
  const meta = Drive.Files.get(xlsxFileId, { fields: "name" });
  const resource = {
    name: "TMP_PFS_CONVERT__" + meta.name + "__" + Utilities.getUuid(),
    title: "TMP_PFS_CONVERT__" + meta.name + "__" + Utilities.getUuid(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  const converted = Drive.Files.copy(resource, xlsxFileId);
  return converted.id;
}

function pfsOpenSpreadsheetWithRetry_(fileId, attempts, sleepMs) {
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
  throw lastErr || new Error("Impossible d'ouvrir le fichier converti.");
}

function pfsSelectBestTemplateSheet_(spreadsheet) {
  const sheets = spreadsheet.getSheets();
  if (!sheets || !sheets.length) return null;

  let bestSheet = null;
  let bestScore = -1;
  for (let i = 0; i < sheets.length; i++) {
    const analysis = pfsAnalyzeSheetStructure_(sheets[i]);
    const score = analysis.headerTokenHits * 100 + analysis.usefulRows * 2 + analysis.usefulWidth;
    if (score > bestScore) {
      bestScore = score;
      bestSheet = sheets[i];
    }
  }

  return bestSheet || sheets[0];
}

function pfsAnalyzeSheetStructure_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const values = (lastRow > 0 && lastCol > 0)
    ? sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues()
    : [];
  const dims = (typeof msGetEffectiveDimensions_ === "function")
    ? msGetEffectiveDimensions_(values)
    : pfsGetEffectiveDimensionsLocal_(values);
  const usefulRows = dims.rows;
  const usefulWidth = dims.cols;
  const headerGuess = pfsGuessHeaderRow_(values, usefulRows, usefulWidth);
  const headerRow = headerGuess.row;
  const headerValues = headerRow ? values[headerRow - 1].slice(0, usefulWidth) : [];
  const mergedCount = (usefulRows > 0 && usefulWidth > 0)
    ? sheet.getRange(1, 1, usefulRows, usefulWidth).getMergedRanges().length
    : 0;
  const specialRows = [];

  for (let r = 0; r < Math.min(headerRow - 1, usefulRows); r++) {
    const nonEmpty = pfsCountNonEmptyCells_(values[r], usefulWidth);
    if (!nonEmpty) continue;
    specialRows.push({
      row: r + 1,
      nonEmpty: nonEmpty,
      sample: values[r].slice(0, Math.min(usefulWidth, 6)).join(" | ").trim()
    });
  }

  return {
    sheetName: sheet.getName(),
    lastRow: lastRow,
    lastCol: lastCol,
    usefulRows: usefulRows,
    usefulWidth: usefulWidth,
    headerRow: headerRow,
    headerValues: headerValues,
    headerTokenHits: headerGuess.tokenHits,
    frozenRows: sheet.getFrozenRows(),
    frozenCols: sheet.getFrozenColumns(),
    mergedCount: mergedCount,
    specialRows: specialRows
  };
}

function pfsGuessHeaderRow_(values, usefulRows, usefulWidth) {
  const expectedTokens = [
    "réf",
    "ref",
    "produit",
    "prix",
    "poids",
    "composition",
    "pays",
    "couleurs",
    "tailles",
    "catégorie",
    "categorie"
  ];

  let bestRow = 0;
  let bestScore = -1;
  let bestHits = 0;
  const scanRows = Math.min(usefulRows || values.length || 0, 12);

  for (let r = 0; r < scanRows; r++) {
    const row = values[r] || [];
    const nonEmpty = pfsCountNonEmptyCells_(row, usefulWidth);
    if (!nonEmpty) continue;

    let hits = 0;
    for (let c = 0; c < Math.min(row.length, usefulWidth || row.length); c++) {
      const cell = normalizeHeaderKey_(row[c]);
      if (!cell) continue;
      for (let t = 0; t < expectedTokens.length; t++) {
        if (cell.indexOf(expectedTokens[t]) !== -1) {
          hits++;
          break;
        }
      }
    }

    const score = hits * 20 + nonEmpty;
    if (score > bestScore) {
      bestScore = score;
      bestRow = r + 1;
      bestHits = hits;
    }
  }

  return { row: bestRow, tokenHits: bestHits };
}

function pfsCountNonEmptyCells_(row, usefulWidth) {
  let count = 0;
  const width = Math.min((row || []).length, usefulWidth || (row || []).length);
  for (let i = 0; i < width; i++) {
    if (String(row[i] || "").trim() !== "") count++;
  }
  return count;
}

function pfsGetEffectiveDimensionsLocal_(values) {
  if (!values || !values.length) return { rows: 0, cols: 0 };

  let lastRow = -1;
  let lastCol = -1;
  for (let r = 0; r < values.length; r++) {
    let rowHasData = false;
    for (let c = 0; c < values[r].length; c++) {
      if (String(values[r][c] || "").trim() !== "") {
        rowHasData = true;
        if (c > lastCol) lastCol = c;
      }
    }
    if (rowHasData) lastRow = r;
  }

  return {
    rows: lastRow + 1,
    cols: lastCol + 1
  };
}

function pfsReplaceSheetFromTemplate_(spreadsheet, targetSheetName, sourceSheet) {
  const existing = spreadsheet.getSheetByName(targetSheetName);
  const targetIndex = existing ? existing.getIndex() : spreadsheet.getSheets().length + 1;
  const tempName = targetSheetName + "__TMP__" + Utilities.getUuid().slice(0, 8);
  const copied = sourceSheet.copyTo(spreadsheet).setName(tempName);

  spreadsheet.setActiveSheet(copied);
  spreadsheet.moveActiveSheet(Math.min(targetIndex, spreadsheet.getSheets().length));

  if (existing) {
    spreadsheet.deleteSheet(existing);
  }

  copied.setName(targetSheetName);
  return copied;
}

function pfsBuildTemplateAnalysisSummary_(analysis, context) {
  const headerPreview = (analysis.headerValues || [])
    .slice(0, Math.min((analysis.headerValues || []).length, 12))
    .map(function(v) { return String(v || "").trim(); })
    .filter(Boolean)
    .join(" | ");

  const specialRows = (analysis.specialRows || []).slice(0, 5).map(function(info) {
    return "Ligne " + info.row + " (" + info.nonEmpty + " cellules): " + info.sample;
  });

  return [
    "Fichier importé: " + (context && context.fileName ? context.fileName : PFS_TEMPLATE_FILENAME),
    "Feuille source retenue: " + analysis.sheetName,
    "Feuille cible: " + (context && context.targetSheetName ? context.targetSheetName : SHEET_PFS_IMPORT),
    "Dimensions utiles: " + analysis.usefulRows + " lignes × " + analysis.usefulWidth + " colonnes",
    "Header row détecté: " + analysis.headerRow,
    "Headers (aperçu): " + (headerPreview || "(aucun header lisible)"),
    "Lignes avant header: " + ((analysis.specialRows || []).length ? (analysis.specialRows || []).length : 0),
    "Lignes particulières (aperçu): " + (specialRows.length ? specialRows.join(" / ") : "aucune"),
    "Volets figés: " + analysis.frozenRows + " ligne(s), " + analysis.frozenCols + " colonne(s)",
    "Zones fusionnées détectées: " + analysis.mergedCount
  ].join("\n");
}

function pfsUniqueList_(values) {
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

/****************************************************
 * PFS Accepted Categories + Normalization
 ****************************************************/
var PFS_ALLOWED_CATEGORIES = [
  "Abayas",
  "Blazers & Tailleurs",
  "Blouses",
  "Blousons",
  "Bodys",
  "Capes & Ponchos",
  "Chemises",
  "Combinaisons",
  "Débardeurs & Bustiers",
  "Doudounes",
  "Ensembles",
  "Gilets",
  "Jeans",
  "Jupes",
  "Leggings",
  "Manteaux",
  "Mariage",
  "Pantacourts",
  "Pantalons",
  "Pulls",
  "Robes",
  "Shorts",
  "Soirées-Cocktails",
  "Sport",
  "Sweats",
  "Tops",
  "T-Shirts",
  "Tuniques",
  "Vestes"
];

function normalizeCategoriePFS_(cat) {
  const raw = String(cat || "").trim();
  if (!raw) return "";

  // If already matches one of the accepted values (case-insensitive), keep canonical spelling.
  const low = raw.toLowerCase();
  for (const v of PFS_ALLOWED_CATEGORIES) {
    if (v.toLowerCase() === low) return v;
  }

  // Heuristic mapping from your STOCK categories / variations
  // (ROBES LONGUES, ROBE, DRESSES, etc.)
  const s = low
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const has = (re) => re.test(s);

  if (has(/abaya/)) return "Abayas";
  if (has(/blazer|tailleur|costume/)) return "Blazers & Tailleurs";
  if (has(/blouse/)) return "Blouses";
  if (has(/blouson|bomber/)) return "Blousons";
  if (has(/body/)) return "Bodys";
  if (has(/cape|poncho/)) return "Capes & Ponchos";
  if (has(/chemise|shirt/)) return "Chemises";
  if (has(/combinaison|jumpsuit|combishort/)) return "Combinaisons";
  if (has(/d[ée]bardeur|bustier/)) return "Débardeurs & Bustiers";
  if (has(/doudoune/)) return "Doudounes";
  if (has(/ensemble|set|coordonn/)) return "Ensembles";
  if (has(/gilet|cardigan/)) return "Gilets";
  if (has(/jean/)) return "Jeans";
  if (has(/jupe|skirt/)) return "Jupes";
  if (has(/legging/)) return "Leggings";
  if (has(/manteau|coat/)) return "Manteaux";
  if (has(/mariage|wedding/)) return "Mariage";
  if (has(/pantacourt/)) return "Pantacourts";
  if (has(/pantalon|trouser|pants/)) return "Pantalons";
  if (has(/pull|knit/)) return "Pulls";
  if (has(/robe|dress/)) return "Robes";
  if (has(/short/)) return "Shorts";
  if (has(/soir[ée]e|cocktail|party/)) return "Soirées-Cocktails";
  if (has(/sport|active/)) return "Sport";
  if (has(/sweat|hoodie/)) return "Sweats";
  if (has(/^top\b|\btop\b/)) return "Tops";
  if (has(/t\s*shirt|tee\b|tshirt/)) return "T-Shirts";
  if (has(/tunique/)) return "Tuniques";
  if (has(/veste|jacket|blazer/)) return "Vestes";

  return "";
}

/****************************************************
 * Composition normalization for PFS
 * Input examples:
 * - "35%LIN65%RAYON" => "35% Lin - 65% Rayon"
 * - "37%COTON 60%POLYESTER3%ELASTHANNE" => "37% Coton - 60% Polyester - 3% Élasthanne"
 ****************************************************/
function normalizeCompositionPFS_(v) {
  const raw0 = String(v || "").trim();
  if (!raw0) return "";

  // Uppercase for matching, normalize accents/typos lightly
  let raw = raw0
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  // Common typos
  raw = raw.replace(/POLYSTER/gi, "POLYESTER");
  raw = raw.replace(/RAYONNE/gi, "RAYON");
  raw = raw.replace(/LINEN/gi, "LIN");
  raw = raw.replace(/ELASTHANNE/gi, "ÉLASTHANNE");

  const up = raw.toUpperCase();

  // Extract pairs like 35% COTON, including stuck formats (35%LIN65%RAYON)
  const re = /(\d+)\s*%\s*([A-ZÉÈÊËÀÂÎÏÔÖÛÜÙÇ]+(?:\s+[A-ZÉÈÊËÀÂÎÏÔÖÛÜÙÇ]+)*)/g;
  const parts = [];
  let m;
  while ((m = re.exec(up)) !== null) {
    const pct = m[1];
    const matRaw = m[2].trim();
    const mat = normalizeMaterialName_(matRaw);
    if (mat) parts.push(pct + "% " + mat);
  }

  if (!parts.length) {
    // Fallback: just title-case the string
    return normalizeMaterialName_(raw);
  }

  return parts.join(" - ");
}

function normalizeMaterialName_(matRaw) {
  let s = String(matRaw || "").trim();
  if (!s) return "";

  s = s
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  // If it's already something like "90% Viscose - 10% Polyester", keep structure for non-percentage fallback
  // For single material names, map to canonical French casing.
  const up = s.toUpperCase();

  // Split on separators for fallback strings
  if (/[\-/,]/.test(s) && /%/.test(s)) {
    // If someone passed a full composition string, return as-is (normalized spaces)
    return s.replace(/\s*-\s*/g, " - ").replace(/\s*,\s*/g, ", ");
  }

  // Canonical mapping
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

  // Some inputs may contain multiple words (e.g. "POLY URÉTHANE")
  // Try exact match first
  if (MAP[up]) return MAP[up];

  // Remove spaces to catch variants
  const compact = up.replace(/\s+/g, "");
  if (MAP[compact]) return MAP[compact];

  // Title-case fallback
  const low = s.toLowerCase();
  return low.charAt(0).toUpperCase() + low.slice(1);
}


/****************************************************
 * Parse STOCK Couleurs into a strict shared structure.
 * Supported:
 * - simple:   "4 ROUGE 2 VERT"
 * - detailed: "1-2 ORANGE 2-1 BLEU"
 ****************************************************/
function parsePfsColorPackFromStock_(raw, sizes, expectedTotal) {
  return parseStockColorPackStrict_(raw, {
    normalizeColor: normalizePfsColorName_,
    invalidColorReason: "Couleur non reconnue dans le catalogue PFS",
    sizes: sizes,
    expectedTotal: expectedTotal
  });
}

function parseStockColorPackStrict_(raw, options) {
  const s = String(raw || "").trim();
  if (!s) {
    return { ok: false, reason: "Couleurs vides" };
  }

  const opts = options || {};
  const normalizeColor = typeof opts.normalizeColor === "function"
    ? opts.normalizeColor
    : function (x) { return String(x || "").trim(); };
  const invalidColorReason = String(opts.invalidColorReason || "Couleur non reconnue");
  const sizes = Array.isArray(opts.sizes) ? opts.sizes : [];
  const expectedTotal = Number(opts.expectedTotal);

  if (!sizes.length) {
    return { ok: false, reason: "Structure tailles invalide pour validation couleurs" };
  }

  const tokens = s.split(/\s+/).filter(Boolean);
  if (tokens.length % 2 !== 0) {
    return { ok: false, reason: "Couleurs mal formées (paires quantité/couleur attendues): " + s };
  }

  let mode = "";
  let total = 0;
  const sumsByPosition = new Array(sizes.length).fill(0);
  const entries = [];

  for (let i = 0; i < tokens.length; i += 2) {
    const qtyToken = String(tokens[i] || "").trim();
    const colorRaw = String(tokens[i + 1] || "").trim();
    const color = normalizeColor(colorRaw);

    if (!color || /^\d/.test(colorRaw)) {
      return { ok: false, reason: invalidColorReason + ": '" + colorRaw + "'" };
    }

    const tokenMode = detectStockColorQtyMode_(qtyToken);
    if (!tokenMode) {
      return { ok: false, reason: "Quantité couleur invalide: '" + qtyToken + "' dans '" + s + "'" };
    }
    if (!mode) mode = tokenMode;
    if (mode !== tokenMode) {
      return { ok: false, reason: "Couleurs mixtes non supportées (simple et détaillé mélangés): " + s };
    }

    if (mode === "simple") {
      const qty = Number(qtyToken);
      if (!Number.isFinite(qty) || qty <= 0 || !Number.isInteger(qty)) {
        return { ok: false, reason: "Quantité couleur invalide: '" + qtyToken + "' dans '" + s + "'" };
      }
      if (qty % sizes.length !== 0) {
        return { ok: false, reason: "Quantité couleur " + qty + " non divisible par " + sizes.length + " taille(s)" };
      }

      const qtyPerSize = qty / sizes.length;
      for (let p = 0; p < sizes.length; p++) {
        sumsByPosition[p] += qtyPerSize;
      }

      entries.push({ color: color, qty: qty, split: null });
      total += qty;
      continue;
    }

    const split = qtyToken.split("-").map(function (part) { return Number(part); });
    if (split.length !== sizes.length) {
      return { ok: false, reason: "Détail couleur incompatible avec le nombre de tailles pour '" + colorRaw + "'" };
    }

    let qtyDetailed = 0;
    for (let p = 0; p < split.length; p++) {
      const partQty = split[p];
      if (!Number.isFinite(partQty) || partQty <= 0 || !Number.isInteger(partQty)) {
        return { ok: false, reason: "Quantité détaillée invalide dans '" + qtyToken + " " + colorRaw + "'" };
      }
      sumsByPosition[p] += partQty;
      qtyDetailed += partQty;
    }

    entries.push({ color: color, qty: qtyDetailed, split: split });
    total += qtyDetailed;
  }

  if (!Number.isFinite(expectedTotal) || expectedTotal <= 0) {
    return { ok: false, reason: "Colisage invalide pour validation couleurs" };
  }
  if (total !== expectedTotal) {
    return { ok: false, reason: "Somme des couleurs = " + total + " au lieu de " + expectedTotal };
  }

  for (let p = 0; p < sizes.length; p++) {
    const expectedQty = Number(sizes[p] && sizes[p].qty);
    if (!Number.isFinite(expectedQty) || expectedQty <= 0) {
      return { ok: false, reason: "Structure tailles invalide pour validation couleurs" };
    }
    if (sumsByPosition[p] !== expectedQty) {
      return { ok: false, reason: "Somme des couleurs incohérente pour la taille " + sizes[p].size + ": " + sumsByPosition[p] + " au lieu de " + expectedQty };
    }
  }

  return { ok: true, mode: mode, total: total, sizeCount: sizes.length, entries: entries };
}

function detectStockColorQtyMode_(qtyToken) {
  const s = String(qtyToken || "").trim();
  if (/^\d+$/.test(s)) return "simple";
  if (/^\d+(?:-\d+)+$/.test(s)) return "detailed";
  return "";
}

/****************************************************
 * Map STOCK color names to the fixed PFS color catalog.
 ****************************************************/
function normalizePfsColorName_(raw) {
  const s = String(raw || "").trim();
  if (!s) return "";
  const up = s.toUpperCase();

  const MAP = {
    "ECRU": "Écru",
    "ÉCRU": "Écru",
    "IVOIRE": "Ivoire",
    "NUDE": "Nude",
    "CREME": "Beige",
    "CRÈME": "Beige",
    "VANILLE": "Beige",
    "BEIGE": "Beige",
    "BLANC": "Blanc",
    "TRANSPARENT": "Blanc",
    "BLEU CIEL": "Bleu",
    "BLEU CLAIR": "Bleu",
    "BLEU": "Bleu",
    "CYAN": "Cyan",
    "TURQUOISE": "Turquoise",
    "BLEU ROI": "Bleu",
    "JEANS": "Bleu",
    "DENIM": "Bleu",
    "BLEU CANARD": "Bleu",
    "BLEU PETROLE": "Bleu",
    "BLEU PÉTROLE": "Bleu",
    "BLEU FONCE": "Marine",
    "BLEU FONCÉ": "Marine",
    "BLEU IRISE": "Bleu",
    "BLEU IRISÉ": "Bleu",
    "MARINE": "Marine",
    "CIEL NOCTURNE": "Marine",
    "GRIS CLAIR": "Gris",
    "GRIS PERLE": "Gris",
    "ARGENT": "Argent",
    "GRIS": "Gris",
    "GRIS SOURIS": "Gris",
    "ACIER": "Gris",
    "GRIS FONCE": "Gris",
    "GRIS FONCÉ": "Gris",
    "CARBONE": "Gris",
    "GRIS ARDOISE": "Gris",
    "ANTHRACITE": "Gris",
    "JAUNE CLAIR": "Jaune",
    "JAUNE CITRON": "Jaune",
    "JAUNE": "Jaune",
    "JAUNE FONCE": "Jaune",
    "JAUNE FONCÉ": "Jaune",
    "JAUNE SOLEIL": "Jaune",
    "JAUNE FLUO": "Jaune",
    "OR": "Doré",
    "DORÉ": "Doré",
    "MOUTARDE": "Moutarde",
    "CAMEL": "Camel",
    "CHAMPAGNE": "Champagne",
    "TAUPE": "Taupe",
    "BRUN": "Brun",
    "COGNAC": "Brun",
    "CARAMEL": "Brun",
    "BRONZE": "Brun",
    "TERRACOTTA": "Brun",
    "MARRON": "Brun",
    "MARRON CLAIR": "Brun",
    "MARRON FONCE": "Brun",
    "MARRON FONCÉ": "Brun",
    "CHOCOLAT": "Brun",
    "BRUN FONCE": "Brun",
    "BRUN FONCÉ": "Brun",
    "NOIR IRISE": "Noir",
    "NOIR IRISÉ": "Noir",
    "NOIR": "Noir",
    "ROSE": "Rose",
    "FUCHSIA": "Fuchsia",
    "ROSE FLUO": "Rose",
    "BLUSH": "Rose",
    "VIEUX ROSE": "Rose",
    "MAGENTA": "Fuchsia",
    "FRAMBOISE": "Fuchsia",
    "ROUGE CLAIR": "Rouge",
    "ROUGE": "Rouge",
    "CORAIL": "Corail",
    "SAUMON": "Corail",
    "ABRICOT": "Corail",
    "ORANGE FLUO": "Orange",
    "ORANGE": "Orange",
    "ROUGE ORANGÉ": "Orange",
    "CUIVRE": "Orange",
    "BRIQUE": "Rouge",
    "ROUILLE": "Rouge",
    "CARMIN": "Rouge",
    "ROUGE FONCE": "Rouge",
    "ROUGE FONCÉ": "Rouge",
    "BORDEAUX": "Bordeaux",
    "VERT CLAIR": "Vert",
    "VERT D'EAU": "Vert",
    "VERT D EAU": "Vert",
    "CÉLADON": "Vert",
    "CELADON": "Vert",
    "VERT FLUO": "Vert",
    "VERT POMME": "Vert",
    "VERT": "Vert",
    "VERT FONCE": "Vert",
    "VERT FONCÉ": "Vert",
    "VERT BOUTEILLE": "Vert",
    "VERT SAPIN": "Vert",
    "VERT CANARD": "Vert",
    "OLIVE": "Olive",
    "KAKI": "Kaki",
    "LILAS": "Lilas",
    "LAVANDE": "Lavande",
    "MAUVE": "Mauve",
    "VIOLET": "Violet",
    "INDIGO": "Indigo",
    "PRUNE": "Prune"
  };

  if (MAP[up]) return MAP[up];

  // fallback exact-title if already a PFS catalog color
  for (var i = 0; i < PFS_COLOR_CATALOG.length; i++) {
    if (PFS_COLOR_CATALOG[i].toUpperCase() === up) return PFS_COLOR_CATALOG[i];
  }
  return "";
}

var PFS_COLOR_CATALOG = [
  "Écru","Ivoire","Nude","Beige","Blanc","Bleu","Cyan","Turquoise","Marine","Gris","Argent","Jaune","Moutarde","Doré","Camel","Champagne","Taupe","Brun","Noir","Corail","Orange","Rose","Fuchsia","Rouge","Bordeaux","Vert","Olive","Kaki","Lilas","Lavande","Mauve","Violet","Indigo","Prune"
];

/****************************************************
 * Parse a normalized tailles string like "6*S/M,6*L/XL"
 * into [{size:"S/M", qty:6}, ...]
 ****************************************************/
function parsePfsTaillesStructure_(taillesStr) {
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
    const size = normalizePfsSizeToken_(m[2]);
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

/****************************************************
 * Build per-color tailles string for PFS.
 * Example:
 * - entry qty 6, global sizes [S/M:6, L/XL:6] -> "3*S/M,3*L/XL"
 * - entry qty 4, global sizes [S/M:6, L/XL:6] -> "2*S/M,2*L/XL"
 ****************************************************/
function buildPfsColorTailles_(colorQty, sizes) {
  const qty = Number(colorQty);
  if (!Number.isFinite(qty) || qty <= 0 || !Array.isArray(sizes) || !sizes.length) {
    return { ok: false, reason: "Quantité couleur ou tailles globales invalides" };
  }

  const n = sizes.length;
  if (qty % n !== 0) {
    return { ok: false, reason: "Quantité couleur " + qty + " non divisible par " + n + " taille(s)" };
  }

  const qtyPerSize = qty / n;

  const parts = [];

  for (let i = 0; i < sizes.length; i++) {
    if (qtyPerSize > sizes[i].qty) {
      return { ok: false, reason: "Quantité couleur " + qty + " incompatible avec la taille " + sizes[i].size };
    }
    parts.push(qtyPerSize + "*" + sizes[i].size);
  }
  return { ok: true, value: parts.join(", ") };
}

function buildPfsDetailedColorTailles_(split, sizes) {
  if (!Array.isArray(split) || !Array.isArray(sizes) || split.length !== sizes.length || !sizes.length) {
    return { ok: false, reason: "Répartition détaillée invalide" };
  }

  const parts = [];
  for (let i = 0; i < sizes.length; i++) {
    const qty = Number(split[i]);
    if (!Number.isFinite(qty) || qty <= 0 || !Number.isInteger(qty)) {
      return { ok: false, reason: "Quantité détaillée invalide pour la taille " + sizes[i].size };
    }
    if (qty > sizes[i].qty) {
      return { ok: false, reason: "Quantité détaillée incompatible avec la taille " + sizes[i].size };
    }
    parts.push(qty + "*" + sizes[i].size);
  }

  return { ok: true, value: parts.join(", ") };
}

/****************************************************
 * Convert STOCK weight in grams to PFS kg text.
 ****************************************************/
function formatWeightKgPFSText_(v) {
  const s = String(v ?? "").trim().replace(/\s+/g, "").replace(",", ".");
  const n = Number(s);
  if (!isFinite(n)) return "";
  const kg = n / 1000;
  return String(kg.toFixed(3)).replace(/\.0+$/, "").replace(/(\.\d*[1-9])0+$/, "$1");
}

/****************************************************
 * Generate PFS photo helper files in Drive
 * - pfs_photo_colors.json (always refreshed)
 * - pfs_photo_dup.py (created only if missing)
 ****************************************************/
function exportPfsPhotoDupHelperToDrive_(selectedRows) {
  const folder = DriveApp.getFolderById(PFS_EXPORT_FOLDER_ID);

  const refs = (selectedRows || []).map(function (it) {
    const entries = Array.isArray(it.colorPack && it.colorPack.entries) ? it.colorPack.entries : [];
    return {
      ref: String(it.ref || "").trim().toUpperCase(),
      colors: entries.map(function (e) { return String(e.color || "").trim(); }).filter(Boolean)
    };
  }).filter(function (x) {
    return x.ref && x.colors && x.colors.length;
  });

  const payload = {
    generated_at: new Date().toISOString(),
    refs: refs
  };

  // Always refresh the ref -> colors mapping
  upsertDriveTextFilePFS_(folder, "pfs_photo_colors.json", JSON.stringify(payload, null, 2), MimeType.PLAIN_TEXT);

  // Create the Python helper only once (or if missing)
  ensureDriveTextFileExistsPFS_(folder, "pfs_photo_dup.py", buildPfsPhotoDupPythonScript_(), MimeType.PLAIN_TEXT);
}

function upsertDriveTextFilePFS_(folder, filename, content, mimeType) {
  const existing = folder.getFilesByName(filename);
  while (existing.hasNext()) {
    existing.next().setTrashed(true);
  }
  folder.createFile(filename, content, mimeType || MimeType.PLAIN_TEXT);
}

function ensureDriveTextFileExistsPFS_(folder, filename, content, mimeType) {
  const existing = folder.getFilesByName(filename);
  if (existing.hasNext()) return;
  folder.createFile(filename, content, mimeType || MimeType.PLAIN_TEXT);
}

function buildPfsPhotoDupPythonScript_() {
  return [
    "#!/usr/bin/env python3",
    "# -*- coding: utf-8 -*-",
    "from __future__ import annotations",
    "",
    "import argparse",
    "import json",
    "import re",
    "from collections import defaultdict",
    "from pathlib import Path",
    "",
    "try:",
    "    from PIL import Image",
    "except ImportError as exc:",
    "    raise SystemExit('Pillow n\\'est pas installé. Installe-le avec: pip install pillow') from exc",
    "",
    "IMAGE_EXTS = {'.jpg', '.jpeg', '.png', '.webp'}",
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
    "def load_mapping(mapping_path: Path | None) -> dict[str, list[str]]:",
    "    if mapping_path is None or not mapping_path.exists():",
    "        return {}",
    "    data = json.loads(mapping_path.read_text(encoding='utf-8'))",
    "    out: dict[str, list[str]] = {}",
    "    if isinstance(data, dict) and isinstance(data.get('refs'), list):",
    "        for item in data['refs']:",
    "            if not isinstance(item, dict) or not item.get('ref'):",
    "                continue",
    "            ref = normalize_ref(str(item['ref']))",
    "            colors = item.get('colors') if isinstance(item.get('colors'), list) else []",
    "            out[ref] = [str(c).strip() for c in colors if str(c).strip()]",
    "    return out",
    "",
    "def choose_source_images(input_dir: Path) -> dict[str, list[Path]]:",
    "    by_ref: dict[str, list[tuple[int, Path]]] = defaultdict(list)",
    "    for path in sorted(input_dir.iterdir(), key=lambda p: p.name.lower()):",
    "        if not path.is_file() or path.suffix.lower() not in IMAGE_EXTS:",
    "            continue",
    "        parsed = parse_input_name(path)",
    "        if parsed is None:",
    "            continue",
    "        ref, pos = parsed",
    "        by_ref[ref].append((pos, path))",
    "",
    "    out: dict[str, list[Path]] = {}",
    "    for ref, entries in by_ref.items():",
    "        entries = sorted(entries, key=lambda x: (x[0], x[1].name.lower()))",
    "        out[ref] = [p for _, p in entries]",
    "    return out",
    "",
    "def duplicate_for_colors(input_dir: Path, output_dir: Path, mapping: dict[str, list[str]]) -> None:",
    "    sources = choose_source_images(input_dir)",
    "    output_dir.mkdir(parents=True, exist_ok=True)",
    "",
    "    for ref, colors in mapping.items():",
    "        if not colors:",
    "            print(f'IGNORÉ (pas de couleurs): {ref}')",
    "            continue",
    "        if ref not in sources or not sources[ref]:",
    "            print(f'IGNORÉ (pas de photo source): {ref}')",
    "            continue",
    "",
    "        src = sources[ref][0]",
    "        with Image.open(src) as img:",
    "            rgb = img.convert('RGB')",
    "            for idx, color in enumerate(colors):",
    "                suffix = 0 if idx == 0 else 1",
    "                filename = f'{ref} {color} {suffix}.jpg'",
    "                dst = output_dir / filename",
    "                rgb.save(dst, 'JPEG', quality=95)",
    "        print(f'OK : {ref} -> {len(colors)} fichier(s)')",
    "",
    "def main() -> int:",
    "    parser = argparse.ArgumentParser(description='Duplique les photos PFS selon les couleurs exportées')",
    "    parser.add_argument('--input', default='input', help='Dossier d\\'entrée contenant les photos source')",
    "    parser.add_argument('--output', default='output', help='Dossier de sortie contenant les JPG renommés')",
    "    parser.add_argument('--mapping', default='pfs_photo_colors.json', help='Fichier JSON de mapping ref -> couleurs')",
    "    args = parser.parse_args()",
    "",
    "    input_dir = Path(args.input)",
    "    output_dir = Path(args.output)",
    "    mapping_path = Path(args.mapping) if args.mapping else None",
    "",
    "    if not input_dir.exists() or not input_dir.is_dir():",
    "        raise SystemExit(f'Dossier input introuvable: {input_dir}')",
    "",
    "    mapping = load_mapping(mapping_path)",
    "    if not mapping:",
    "        raise SystemExit('Aucun mapping ref -> couleurs trouvé')",
    "",
    "    duplicate_for_colors(input_dir, output_dir, mapping)",
    "    print(f'Terminé. Fichiers générés dans: {output_dir}')",
    "    return 0",
    "",
    "if __name__ == '__main__':",
    "    raise SystemExit(main())",
    ""
  ].join("\n");
}
function parseColisagePFS_(v) {
  const s = String(v ?? "").trim().replace(/\s+/g, "").replace(",", ".");
  const n = Number(s);

  if (!Number.isFinite(n) || n <= 0 || !Number.isInteger(n)) {
    return { ok: false, reason: "Colisage invalide: '" + String(v ?? "") + "'" };
  }

  return { ok: true, value: n };
}
