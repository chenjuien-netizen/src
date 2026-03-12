/****************************************************
 * Export STOCK → e_export (eFashion)
 ****************************************************/

var SHEET_E_EXPORT = "e_export";

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

    const sheetRow = 2 + i;
    if (!colisage.ok) {
      bad.push({ row: sheetRow, ref: ref || "(vide)", reason: colisage.reason });
      continue;
    }
    if (!tailles) {
      bad.push({ row: sheetRow, ref: ref || "(vide)", reason: "Tailles introuvables ou invalides dans Contenu colis" });
      continue;
    }
    const colorCheck = validateEFashionColorPack_(couleursStock, colisage.value, tailles);
    if (!colorCheck.ok) {
      bad.push({ row: sheetRow, ref: ref || "(vide)", reason: colorCheck.reason });
      continue;
    }
    const couleursEf = formatEFashionMixedColorsFromStock_(couleursStock, tailles);
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
          "  Ref : " + x.ref,
          "  Motif : " + x.reason
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
function formatEFashionMixedColorsFromStock_(raw, taillesStr) {
  const s = String(raw || "").trim();
  if (!s) return "";

  const taillesInfo = parseEFashionTaillesStructure_(taillesStr);
  if (!taillesInfo.ok) return "";

  const sizeCount = taillesInfo.sizes.length;
  if (!sizeCount) return "";

  const re = /(\d+)\s+([A-Za-zÀ-ÿ]+)/g;
  const parts = [];
  let m;
  while ((m = re.exec(s)) !== null) {
    const qty = Number(m[1]);
    const color = normalizeEFashionColorName_(m[2]);
    if (!Number.isFinite(qty) || qty <= 0 || !color) continue;
    if (qty % sizeCount !== 0) return "";
    const qtyPerSize = qty / sizeCount;
    const split = new Array(sizeCount).fill(qtyPerSize);
    parts.push(color + "*" + split.join("-"));
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
function validateEFashionColorPack_(raw, expectedTotal, taillesStr) {
  const s = String(raw || "").trim();
  if (!s) {
    return { ok: false, reason: "Couleurs vides" };
  }

  const taillesInfo = parseEFashionTaillesStructure_(taillesStr);
  if (!taillesInfo.ok) {
    return { ok: false, reason: "Structure tailles invalide: " + taillesInfo.reason };
  }
  const sizeCount = taillesInfo.sizes.length;
  if (!sizeCount) {
    return { ok: false, reason: "Structure tailles vide" };
  }

  const tokens = s.split(/\s+/).filter(Boolean);
  if (tokens.length % 2 !== 0) {
    return { ok: false, reason: "Couleurs mal formées (paires quantité/couleur attendues): " + s };
  }

  let total = 0;
  for (let i = 0; i < tokens.length; i += 2) {
    const qty = Number(tokens[i]);
    const colorRaw = String(tokens[i + 1] || "").trim();
    const color = normalizeEFashionColorName_(colorRaw);

    if (!Number.isFinite(qty) || qty <= 0 || !Number.isInteger(qty)) {
      return { ok: false, reason: "Quantité couleur invalide: '" + tokens[i] + "' dans '" + s + "'" };
    }
    if (!color || /^\d+$/.test(colorRaw)) {
      return { ok: false, reason: "Couleurs mal formées (couleur manquante ou invalide) dans '" + s + "'" };
    }
    if (qty % sizeCount !== 0) {
      return { ok: false, reason: "Quantité couleur " + qty + " non divisible par " + sizeCount + " taille(s)" };
    }

    total += qty;
  }

  const exp = Number(expectedTotal);
  if (!Number.isFinite(exp) || exp <= 0) {
    return { ok: false, reason: "Colisage invalide pour validation couleurs" };
  }

  if (total !== exp) {
    return { ok: false, reason: "Somme des couleurs = " + total + " au lieu de " + exp + " dans '" + s + "'" };
  }

  return { ok: true, total: total };
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
