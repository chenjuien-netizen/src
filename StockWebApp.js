function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  if (String(params.app || "").trim().toLowerCase() === "v1") {
    return StockWebAppLegacy_doGet_();
  }
  return StockWebAppV2_doGet_();
}
