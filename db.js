const DB_NAME = "szfashion-pwa";
const DB_VERSION = 1;
const RECORDS_STORE = "records";
const META_STORE = "meta";

let dbPromise;

function openDatabase() {
  if (dbPromise) return dbPromise;

  dbPromise = new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = () => {
      const db = request.result;

      if (!db.objectStoreNames.contains(RECORDS_STORE)) {
        db.createObjectStore(RECORDS_STORE, { keyPath: "reference" });
      }

      if (!db.objectStoreNames.contains(META_STORE)) {
        db.createObjectStore(META_STORE, { keyPath: "key" });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error || new Error("IndexedDB unavailable"));
  });

  return dbPromise;
}

function withStore(storeName, mode, callback) {
  return openDatabase().then((db) => new Promise((resolve, reject) => {
    const transaction = db.transaction(storeName, mode);
    const store = transaction.objectStore(storeName);

    let callbackResult;

    transaction.oncomplete = () => resolve(callbackResult);
    transaction.onerror = () => reject(transaction.error || new Error("IndexedDB transaction failed"));
    transaction.onabort = () => reject(transaction.error || new Error("IndexedDB transaction aborted"));

    try {
      callbackResult = callback(store, transaction);
    } catch (error) {
      reject(error);
    }
  }));
}

export async function getAllRecords() {
  return withStore(RECORDS_STORE, "readonly", (store) => new Promise((resolve, reject) => {
    const request = store.getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error || new Error("Failed to read records"));
  }));
}

export async function replaceRecords(records) {
  return withStore(RECORDS_STORE, "readwrite", (store) => {
    store.clear();
    records.forEach((record) => store.put(record));
  });
}

export async function getMetaMap() {
  const entries = await withStore(META_STORE, "readonly", (store) => new Promise((resolve, reject) => {
    const request = store.getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error || new Error("Failed to read meta"));
  }));

  return entries.reduce((accumulator, entry) => {
    accumulator[entry.key] = entry.value;
    return accumulator;
  }, {});
}

export async function setMetaValues(values) {
  const entries = Object.entries(values);

  return withStore(META_STORE, "readwrite", (store) => {
    entries.forEach(([key, value]) => {
      store.put({ key, value });
    });
  });
}
