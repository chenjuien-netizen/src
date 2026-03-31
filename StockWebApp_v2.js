function StockWebAppV2_doGet_() {
  const template = HtmlService.createTemplateFromFile("StockMobile_v2");

  template.StockMobileV2_initialBootJson = StockWebAppV2_safeJsonForTemplate_(StockWebAppV2_buildBoot_());

  return template.evaluate()
    .setTitle("SZFashion | Inventaire V2")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, viewport-fit=cover")
    .addMetaTag("mobile-web-app-capable", "yes")
    .addMetaTag("apple-mobile-web-app-capable", "yes")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function StockWebAppV2_buildBoot_() {
  // Future boot V2: remplacer ce boot mock par le bootstrap reel de la V2.
  return {
    generatedAt: new Date().toISOString(),
    dataConnected: false,
    ui: {
      defaultView: "inventory"
    },
    inventory: {
      items: [
        {
          id: "mock-ref-001",
          reference: "SZ-ALPHA-001",
          stockDisplay: "12 箱 4 件",
          stockState: "positive",
          warehouse: "A1",
          createdAt: "31/03/2026",
          remark: "Mock frontend only"
        },
        {
          id: "mock-ref-002",
          reference: "SZ-BETA-002",
          stockDisplay: "0",
          stockState: "zero",
          warehouse: "B2",
          createdAt: "30/03/2026",
          remark: ""
        },
        {
          id: "mock-ref-003",
          reference: "SZ-GAMMA-003",
          stockDisplay: "3 箱",
          stockState: "positive",
          warehouse: "C4",
          createdAt: "28/03/2026",
          remark: "Fiche mock"
        }
      ]
    },
    history: {
      items: [
        {
          id: "history-001",
          reference: "SZ-ALPHA-001",
          timestampLabel: "31/03 10:45",
          actionType: "modifier",
          beforeDisplay: "10 箱 2 件",
          afterDisplay: "12 箱 4 件",
          remark: "Ajustement mock",
          source: "mock-v2"
        },
        {
          id: "history-002",
          reference: "SZ-GAMMA-003",
          timestampLabel: "30/03 16:20",
          actionType: "sortie_rapide",
          beforeDisplay: "5 箱",
          afterDisplay: "3 箱",
          remark: "Sortie mock",
          source: "mock-v2"
        }
      ]
    },
    detailByReference: {
      "SZ-ALPHA-001": {
        reference: "SZ-ALPHA-001",
        stockDisplay: "12 箱 4 件",
        stockState: "positive",
        warehouse: "A1",
        createdAt: "31/03/2026",
        lastMovementAt: "31/03/2026 10:45",
        summary: "Entrepot A1 · produit mock",
        remark: "Cette fiche est mockee pour la reconstruction V2.",
        history: [
          {
            id: "detail-history-001",
            reference: "SZ-ALPHA-001",
            timestampLabel: "31/03 10:45",
            actionType: "modifier",
            beforeDisplay: "10 箱 2 件",
            afterDisplay: "12 箱 4 件",
            remark: "Ajustement mock",
            source: "mock-v2"
          }
        ]
      },
      "SZ-BETA-002": {
        reference: "SZ-BETA-002",
        stockDisplay: "0",
        stockState: "zero",
        warehouse: "B2",
        createdAt: "30/03/2026",
        lastMovementAt: "",
        summary: "Aucune donnee connectee",
        remark: "",
        history: []
      },
      "SZ-GAMMA-003": {
        reference: "SZ-GAMMA-003",
        stockDisplay: "3 箱",
        stockState: "positive",
        warehouse: "C4",
        createdAt: "28/03/2026",
        lastMovementAt: "30/03/2026 16:20",
        summary: "Reference mock avec historique",
        remark: "Conserver ce rendu pour le futur branchement.",
        history: [
          {
            id: "detail-history-002",
            reference: "SZ-GAMMA-003",
            timestampLabel: "30/03 16:20",
            actionType: "sortie_rapide",
            beforeDisplay: "5 箱",
            afterDisplay: "3 箱",
            remark: "Sortie mock",
            source: "mock-v2"
          }
        ]
      }
    }
  };
}

function StockWebAppV2_include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function StockWebAppV2_safeJsonForTemplate_(value) {
  return JSON.stringify(value)
    .replace(/</g, "\\u003c")
    .replace(/>/g, "\\u003e")
    .replace(/&/g, "\\u0026")
    .replace(/\u2028/g, "\\u2028")
    .replace(/\u2029/g, "\\u2029");
}
