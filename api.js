const endpointMeta = document.querySelector('meta[name="szfashion-pull-url"]');

function getPullUrl() {
  const value = endpointMeta ? String(endpointMeta.content || "").trim() : "";
  if (!value || value.includes("__GAS_PULL_URL__")) {
    throw new Error("URL Google Apps Script non configurée.");
  }
  return value;
}

export async function fetchInventory() {
  const response = await fetch(getPullUrl(), {
    method: "GET",
    headers: {
      Accept: "application/json"
    },
    cache: "no-store"
  });

  if (!response.ok) {
    throw new Error("Impossible de récupérer l'inventaire.");
  }

  const payload = await response.json();
  if (!payload || payload.ok !== true || !Array.isArray(payload.records)) {
    throw new Error(payload && payload.message ? payload.message : "Réponse inventaire invalide.");
  }

  return payload;
}
