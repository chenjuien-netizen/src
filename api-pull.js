function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = String(params.action || "pull").trim();

  if (action !== "pull") {
    return ApiPull_jsonResponse_({
      ok: false,
      action: action || "",
      errorCode: "INVALID_ACTION",
      message: "Action non supportee."
    });
  }

  try {
    const records = SheetRepository_getInventoryRecords_();
    return ApiPull_jsonResponse_({
      ok: true,
      action: "pull",
      pulledAt: new Date().toISOString(),
      records: records,
      meta: {
        recordCount: records.length
      }
    });
  } catch (error) {
    return ApiPull_jsonResponse_({
      ok: false,
      action: "pull",
      errorCode: "SHEET_READ_ERROR",
      message: error && error.message ? error.message : "Impossible de lire la feuille source."
    });
  }
}

function ApiPull_jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
