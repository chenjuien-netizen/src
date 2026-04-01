import { getAllRecords, getMetaMap, replaceRecords, setMetaValues } from "./db.js";
import { fetchInventory } from "./api.js";

const state = {
  records: [],
  filteredRecords: [],
  meta: {
    lastPulledAt: "",
    lastPullStatus: "idle",
    lastPullError: ""
  },
  isLoading: true,
  hasLocalData: false,
  searchQuery: ""
};

const elements = {
  networkStatus: document.getElementById("networkStatus"),
  syncStatus: document.getElementById("syncStatus"),
  lastUpdated: document.getElementById("lastUpdated"),
  recordCount: document.getElementById("recordCount"),
  inventoryList: document.getElementById("inventoryList"),
  emptyState: document.getElementById("emptyState"),
  emptyTitle: document.getElementById("emptyTitle"),
  emptyMessage: document.getElementById("emptyMessage"),
  searchInput: document.getElementById("searchInput"),
  refreshButton: document.getElementById("refreshButton")
};

function formatTimestamp(value) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "-";
  return new Intl.DateTimeFormat("fr-FR", {
    dateStyle: "short",
    timeStyle: "short"
  }).format(date);
}

function normalizeRecords(records) {
  return records
    .map((record) => ({
      reference: String(record && record.reference ? record.reference : "").trim()
    }))
    .filter((record) => record.reference)
    .sort((left, right) => left.reference.localeCompare(right.reference, "fr", { sensitivity: "base" }));
}

function applySearch(records, query) {
  const normalizedQuery = String(query || "").trim().toLowerCase();
  if (!normalizedQuery) return records.slice();

  return records.filter((record) => record.reference.toLowerCase().includes(normalizedQuery));
}

function setEmptyState(title, message) {
  elements.emptyTitle.textContent = title;
  elements.emptyMessage.textContent = message;
  elements.emptyState.classList.remove("hidden");
}

function hideEmptyState() {
  elements.emptyState.classList.add("hidden");
}

function renderList() {
  elements.inventoryList.innerHTML = "";

  state.filteredRecords.forEach((record) => {
    const item = document.createElement("li");
    item.className = "inventory-item";

    const row = document.createElement("div");
    row.className = "inventory-row";

    const reference = document.createElement("span");
    reference.className = "reference-text";
    reference.textContent = record.reference;

    row.appendChild(reference);
    item.appendChild(row);
    elements.inventoryList.appendChild(item);
  });
}

function renderStatus() {
  const online = navigator.onLine;
  elements.networkStatus.textContent = online ? "En ligne" : "Hors ligne";
  elements.networkStatus.classList.toggle("network-offline", !online);
  elements.lastUpdated.textContent = formatTimestamp(state.meta.lastPulledAt);
  elements.recordCount.textContent = String(state.filteredRecords.length);

  if (state.isLoading) {
    elements.syncStatus.textContent = "Chargement...";
    return;
  }

  if (!online && state.hasLocalData) {
    elements.syncStatus.textContent = "Consultation locale";
    return;
  }

  if (!online && !state.hasLocalData) {
    elements.syncStatus.textContent = "Aucune donnée locale";
    return;
  }

  if (state.meta.lastPullStatus === "error" && state.hasLocalData) {
    elements.syncStatus.textContent = "Dernière mise à jour conservée";
    return;
  }

  if (state.meta.lastPullStatus === "success") {
    elements.syncStatus.textContent = "Inventaire à jour";
    return;
  }

  elements.syncStatus.textContent = "Prêt";
}

function renderEmptyState() {
  if (state.isLoading) {
    hideEmptyState();
    return;
  }

  if (!state.hasLocalData && !navigator.onLine) {
    setEmptyState(
      "Aucune donnée disponible",
      "Connecte-toi à Internet pour charger l'inventaire une première fois."
    );
    return;
  }

  if (!state.hasLocalData && state.meta.lastPullStatus === "error") {
    setEmptyState(
      "Impossible de charger l'inventaire",
      state.meta.lastPullError || "Réessaie avec une connexion Internet disponible."
    );
    return;
  }

  if (state.hasLocalData && state.filteredRecords.length === 0) {
    setEmptyState(
      "Aucun résultat",
      "Aucune référence ne correspond à la recherche."
    );
    return;
  }

  hideEmptyState();
}

function render() {
  state.filteredRecords = applySearch(state.records, state.searchQuery);
  renderList();
  renderStatus();
  renderEmptyState();
}

async function refreshInventory(reason = "manual") {
  if (!navigator.onLine) {
    render();
    return;
  }

  state.isLoading = true;
  state.meta.lastPullStatus = "loading";
  state.meta.lastPullError = "";
  render();

  try {
    const payload = await fetchInventory();
    const records = normalizeRecords(payload.records);
    const pulledAt = payload.pulledAt || new Date().toISOString();

    await replaceRecords(records);
    await setMetaValues({
      lastPulledAt: pulledAt,
      lastPullStatus: "success",
      lastPullError: ""
    });

    state.records = records;
    state.hasLocalData = records.length > 0;
    state.meta.lastPulledAt = pulledAt;
    state.meta.lastPullStatus = "success";
    state.meta.lastPullError = "";
  } catch (error) {
    state.meta.lastPullStatus = "error";
    state.meta.lastPullError = error instanceof Error ? error.message : "Erreur réseau.";

    await setMetaValues({
      lastPullStatus: "error",
      lastPullError: state.meta.lastPullError
    });

    if (!state.hasLocalData && reason === "startup") {
      state.records = [];
    }
  } finally {
    state.isLoading = false;
    render();
  }
}

async function loadLocalState() {
  const [records, meta] = await Promise.all([getAllRecords(), getMetaMap()]);

  state.records = normalizeRecords(records);
  state.filteredRecords = state.records.slice();
  state.hasLocalData = state.records.length > 0;
  state.meta.lastPulledAt = String(meta.lastPulledAt || "");
  state.meta.lastPullStatus = String(meta.lastPullStatus || "idle");
  state.meta.lastPullError = String(meta.lastPullError || "");
}

function bindEvents() {
  elements.searchInput.addEventListener("input", (event) => {
    state.searchQuery = event.target.value;
    render();
  });

  elements.refreshButton.addEventListener("click", () => {
    refreshInventory("manual");
  });

  window.addEventListener("online", () => {
    render();
    refreshInventory("online");
  });

  window.addEventListener("offline", () => {
    render();
  });
}

async function registerServiceWorker() {
  if (!("serviceWorker" in navigator)) return;
  try {
    await navigator.serviceWorker.register("./sw.js");
  } catch (_error) {
    // Service worker registration should not block the app.
  }
}

async function initializeApp() {
  bindEvents();
  render();

  await Promise.all([
    registerServiceWorker(),
    loadLocalState()
  ]);

  state.isLoading = false;
  render();

  if (navigator.onLine) {
    await refreshInventory("startup");
  }
}

initializeApp().catch(async (error) => {
  state.isLoading = false;
  state.meta.lastPullStatus = "error";
  state.meta.lastPullError = error instanceof Error ? error.message : "Erreur de démarrage.";

  try {
    await setMetaValues({
      lastPullStatus: "error",
      lastPullError: state.meta.lastPullError
    });
  } catch (_error) {
    // Ignore persistence errors in the fallback path.
  }

  render();
});
