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
    "Composition matérielle"
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

  const refVals = stock.getRange(2, cRef, n, 1).getValues().flat();
  const prixVals = stock.getRange(2, cPrix, n, 1).getValues().flat();
  const dateVals = stock.getRange(2, cDate, n, 1).getValues().flat();
  const contenuVals = stock.getRange(2, cContenu, n, 1).getValues().flat();
  const poidsVals = stock.getRange(2, cPoids, n, 1).getValues().flat();
  const paysVals = stock.getRange(2, cPays, n, 1).getValues().flat();
  const compoVals = stock.getRange(2, cCompo, n, 1).getValues().flat();
  const selVals = stock.getRange(2, cSel, n, 1).getValues().flat();

  let rows = [];

  for (let i = 0; i < n; i++) {

    if (selVals[i] !== true) continue;

    const ref = String(refVals[i] || "").trim();
    if (!ref) continue;

    const prix = formatPriceEFashionText_(prixVals[i]);
    const dt = normalizeDateForSortEf_(dateVals[i]);

    const tailles = normalizeTaillesEFashion_(extractTaillesFromContenuColis_(contenuVals[i]));
    const poids = String(poidsVals[i] ?? "").trim();
    const pays = String(paysVals[i] ?? "").trim();

    // Composition: prefer shared normalizer from PFS.js if present
    let compo = "";
    if (typeof normalizeCompositionPFS_ === "function") {
      compo = normalizeCompositionPFS_(compoVals[i]);
    } else {
      compo = normalizeCompositionEFashionFallback_(compoVals[i]);
    }

    const sheetRow = 2 + i;

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
    });

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
    line[cSousCategorie - 1] = "Robes & Combinaisons";
    line[cSousSousCategorie - 1] = "Robes longues";

    // User wants mixed-color selling with no minimum quantity for now
    line[cVenduPar - 1] = "melangees";
    line[cQteMin - 1] = "";

    // Extra safety: also write fixed-position H (Qte min) in case header matching is off
    if (hinfo.width >= 8) line[7] = "";

    line[cCollection - 1] = "Printemps / Été";

    // From STOCK
    line[cReference - 1] = it.ref;
    line[cProvenances - 1] = it.pays;

    // For test import, use the mixed-color style from their example
    line[cTailles - 1] = "S,M,L";
    line[cCouleurs - 1] = "Bleu*1-1-1,Noir*1-1-1,Rouge*1-1-1";

    // Extra safety: fixed-position J/K as well
    if (hinfo.width >= 10) line[9] = "S,M,L";
    if (hinfo.width >= 11) line[10] = "Bleu*1-1-1,Noir*1-1-1,Rouge*1-1-1";

    // Prix: write as TEXT "0.00" (dot)
    line[cPrixHT - 1] = it.prix;

    // Poids: write as-is (no conversion)
    line[cPoidsKG - 1] = it.poids;

    // Composition: normalized style with import normalizer
    line[cCompositions - 1] = normalizeCompositionEFashionImport_(it.compo) || "Viscose*90,Polyester*10";

    // Extra safety fallback by fixed template positions (A..U) for test import
    if (hinfo.width >= 1)  line[0] = "J&S FASHION";         // A Marque
    if (hinfo.width >= 2)  line[1] = it.ref;                  // B Référence
    if (hinfo.width >= 3)  line[2] = "Femme";               // C Catégorie
    if (hinfo.width >= 4)  line[3] = "Robes & Combinaisons";// D Sous catégorie
    if (hinfo.width >= 5)  line[4] = "Robes longues";       // E Sous-sous Catégorie
    if (hinfo.width >= 6)  line[5] = it.pays;                 // F Provenances
    if (hinfo.width >= 7)  line[6] = "melangees";          // G Vendu Par
    if (hinfo.width >= 8)  line[7] = "";                   // H Qte min
    if (hinfo.width >= 9)  line[8] = "Printemps / Été";     // I Collection
    if (hinfo.width >= 10) line[9] = "S,M,L";               // J Tailles
    if (hinfo.width >= 11) line[10] = "Bleu,Jaune,vert";    // K Couleurs
    if (hinfo.width >= 12) line[11] = it.prix;                // L Prix HT
    if (hinfo.width >= 14) line[13] = it.poids;               // N Poids KG
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

    // Auto-uncheck: décoche 选择 pour les lignes exportées
    try {
      const a1 = rows.map(it => stock.getRange(it.row, cSel).getA1Notation());
      stock.getRangeList(a1).setValue(false);
    } catch (e) {
      for (const it of rows) {
        stock.getRange(it.row, cSel).setValue(false);
      }
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
 * Tailles eFashion
 * v1 focus: "6 x M/L - 6 x XL/XXL" -> "6*M/L,6*XL/XXL"
 ****************************************************/
function normalizeTaillesEFashion_(v) {
  const s = String(v || "").trim();
  if (!s) return "";
  // keep as star/comma list; remove spaces around commas
  return s.replace(/\s*,\s*/g, ",");
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
 * Extract tailles from Contenu colis (v1 strict parser)
 ****************************************************/
function extractTaillesFromContenuColis_(v) {
  const s0 = String(v || "").trim();
  if (!s0) return "";

  // v1 strict: support principal
  // "6 x M/L - 6 x XL/XXL" -> "6*M/L,6*XL/XXL"
  let s = s0;

  // 1) Convert "6 x" / "6x" / "6×" to "6*"
  s = s.replace(/(\d+)\s*[xX×]\s*/g, "$1*");

  // 2) Convert " - " to comma list separator
  s = s.replace(/\s*-\s*/g, ",");

  // 3) Normalize commas
  s = s.replace(/\s*,\s*/g, ",");

  // Validate: must contain at least one "n*SIZE" block
  const re = /(\d+)\*\s*([A-Za-z0-9]+(?:\/[A-Za-z0-9]+)*)/g;
  const parts = [];
  let m;
  while ((m = re.exec(s)) !== null) {
    parts.push(`${m[1]}*${m[2]}`);
  }

  return parts.length ? parts.join(",") : "";
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