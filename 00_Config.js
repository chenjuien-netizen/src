// 00_Config.gs
const SHEET_UI = "ARRIVAGES";
const SHEET_DB = "ARRIVAGES_DB";


const UI_ID_CELL       = "B1";
const UI_UPDATED_CELL  = "A2";
const UI_CREATED_CELL  = "B2";
const UI_ENTREPOT_CELL = "E1";
const UI_QUICK_RANGE = "G4:G200";    // zone multi-lignes
const UI_QUICK_COL = 7;            
const UI_QUICK_ROW_START = 4;
const UI_QUICK_ROW_END = 200;

const UI_TABLE_RANGE     = "A4:F200";
const UI_TABLE_START_ROW = 4;
const UI_TABLE_ROWS      = 197;
const UI_TABLE_COLS      = 6;

const SUP_RANGE = "H4:K200";
const SUP_START_ROW = 4;

const LABEL_ADD = "新增";
// ===============
// CONFIG
// ===============
const SHEET_TEMPLATE_STOCK = "TEMPLATE_STOCK";
const SHEET_STOCK = "STOCK";
const UI_LINEID_COL = 12;      // L
const UI_LINEID_ROW_START = 4; // L4

var SHEET_MS_EXPORT = "MS_EXPORT";

// Dossier Drive de sortie (export .xlsx) — éviter les collisions de variables globales
var MS_EXPORT_DRIVE_FOLDER_ID = "1B2P4SwJbwbPEXW5XO_LSNdIMYA-aYMZa";