/* ============================================================================
 * AGL BUDGET — store.js
 * Couche de persistance IndexedDB (snapshot après chaque upload — §5).
 * Permet la reprise sans re-upload en cas de crash navigateur.
 * Stores :
 *   - sources    : métadonnées + snapshot agrégé de chaque base chargée
 *   - aggregates : agrégats calculés (STATCOM par métier × année, etc.)
 *   - rmc        : Référentiel Maître Client construit
 *   - meta       : état de session (sprint, versions, horodatage)
 * Module pur, API Promise. Expose window.AGL.store.
 * ==========================================================================*/
(function (global) {
  "use strict";

  const DB_NAME = "agl_budget";
  const DB_VERSION = 1;
  const STORES = ["sources", "aggregates", "rmc", "meta"];

  let _dbPromise = null;

  function open() {
    if (_dbPromise) return _dbPromise;
    _dbPromise = new Promise((resolve, reject) => {
      if (!global.indexedDB) return reject(new Error("IndexedDB indisponible sur ce navigateur."));
      const req = indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = (e) => {
        const db = e.target.result;
        STORES.forEach((s) => { if (!db.objectStoreNames.contains(s)) db.createObjectStore(s, { keyPath: "key" }); });
      };
      req.onsuccess = (e) => resolve(e.target.result);
      req.onerror = (e) => reject(e.target.error);
    });
    return _dbPromise;
  }

  function _tx(store, mode) {
    return open().then((db) => db.transaction(store, mode).objectStore(store));
  }

  function put(store, key, value) {
    return _tx(store, "readwrite").then((os) => new Promise((resolve, reject) => {
      const r = os.put({ key: key, value: value, ts: Date.now() });
      r.onsuccess = () => resolve(true);
      r.onerror = (e) => reject(e.target.error);
    }));
  }

  function get(store, key) {
    return _tx(store, "readonly").then((os) => new Promise((resolve, reject) => {
      const r = os.get(key);
      r.onsuccess = () => resolve(r.result ? r.result.value : null);
      r.onerror = (e) => reject(e.target.error);
    }));
  }

  function list(store) {
    return _tx(store, "readonly").then((os) => new Promise((resolve, reject) => {
      const r = os.getAll();
      r.onsuccess = () => resolve(r.result || []);
      r.onerror = (e) => reject(e.target.error);
    }));
  }

  function del(store, key) {
    return _tx(store, "readwrite").then((os) => new Promise((resolve, reject) => {
      const r = os.delete(key);
      r.onsuccess = () => resolve(true);
      r.onerror = (e) => reject(e.target.error);
    }));
  }

  function clear(store) {
    return _tx(store, "readwrite").then((os) => new Promise((resolve, reject) => {
      const r = os.clear();
      r.onsuccess = () => resolve(true);
      r.onerror = (e) => reject(e.target.error);
    }));
  }

  function clearAll() { return Promise.all(STORES.map(clear)); }

  // Estimation du quota utilisé (best-effort, §5 quota 500MB+).
  function quota() {
    if (navigator.storage && navigator.storage.estimate) return navigator.storage.estimate();
    return Promise.resolve({ usage: null, quota: null });
  }

  global.AGL = global.AGL || {};
  global.AGL.store = { open, put, get, list, del, clear, clearAll, quota, STORES, DB_NAME };
})(typeof window !== "undefined" ? window : this);
