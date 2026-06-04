/* ============================================================================
 * AGL BUDGET — loader.js
 * Orchestration de l'upload séquentiel (§5) :
 *   FileReader → Web Worker (worker.js) → progression chunked → snapshot IndexedDB.
 * Une base à la fois. Données brutes libérées de la RAM après snapshot ;
 * seuls métadonnées + échantillon + colonnes sont conservés en mémoire vive.
 * Détection mémoire : alerte si RAM dispo < 500 Mo.
 *
 * Dépend de : AGL.utils, AGL.store. Expose window.AGL.loader.
 * ==========================================================================*/
(function (global) {
  "use strict";
  const U = global.AGL.utils;
  const store = global.AGL.store;

  // Les 4 bases sources (+ années STATCOM) — §2.
  const SOURCES = {
    CRM:          { label: "CRM (3 ans)", role: "Référentiel commercial", priorite: "PRIORITÉ 1 — LOCKED" },
    IRIS:         { label: "IRIS (3 ans)", role: "Performance opérat. AGL", priorite: "PRIORITÉ 2" },
    STATCOM_2023: { label: "STATCOM 2023", role: "Marché total CIV", priorite: "Fallback marchandise" },
    STATCOM_2024: { label: "STATCOM 2024", role: "Marché total CIV", priorite: "Fallback marchandise" },
    STATCOM_2025: { label: "STATCOM 2025", role: "Marché total CIV", priorite: "Fallback marchandise" },
    RUBRIKS:      { label: "RUBRIKS", role: "Budget PFA 2026", priorite: "Enrichissement" }
  };

  let _worker = null;
  let _seq = 0;
  const _listeners = {}; // id -> callbacks

  function _getWorker() {
    if (!_worker) {
      _worker = new Worker("js/worker.js");
      _worker.onmessage = (ev) => {
        const m = ev.data || {};
        const cb = _listeners[m.id];
        if (!cb) return;
        if (m.type === "progress") cb.onProgress && cb.onProgress(m);
        else if (m.type === "done") { cb.onDone && cb.onDone(m); delete _listeners[m.id]; }
        else if (m.type === "error") { cb.onError && cb.onError(new Error(m.message)); delete _listeners[m.id]; }
      };
      _worker.onerror = (e) => U.log("loader", "worker error: " + e.message, "error");
    }
    return _worker;
  }

  // Lecture d'un File → ArrayBuffer (Promise).
  function _readFile(file) {
    return new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = () => resolve(fr.result);
      fr.onerror = () => reject(fr.error);
      fr.readAsArrayBuffer(file);
    });
  }

  // Vérifie la mémoire dispo (best-effort, Chrome uniquement).
  function memoryWarning() {
    const pm = global.performance && global.performance.memory;
    if (!pm) return null;
    const freeMB = (pm.jsHeapSizeLimit - pm.usedJSHeapSize) / 1e6;
    if (freeMB < 500) return "RAM JS disponible faible (" + freeMB.toFixed(0) + " Mo). Conseil §5 : charger par année (mode dégradé 1 an).";
    return null;
  }

  /* Charge une base. callbacks = { onProgress, onDone, onError }.
   * Snapshot persistant : on stocke métadonnées + colonnes + échantillon
   * (les lignes brutes complètes ne sont PAS persistées par défaut pour
   *  respecter le quota ; l'agrégation Sprint 3+ stockera les agrégats). */
  function loadFile(sourceKey, file, callbacks) {
    callbacks = callbacks || {};
    if (!SOURCES[sourceKey]) return Promise.reject(new Error("Source inconnue : " + sourceKey));

    const memWarn = memoryWarning();
    if (memWarn) U.log("loader", memWarn, "warn");

    const id = "job_" + (++_seq);
    return _readFile(file).then((buffer) => new Promise((resolve, reject) => {
      _listeners[id] = {
        onProgress: callbacks.onProgress,
        onError: (err) => { callbacks.onError && callbacks.onError(err); reject(err); },
        onDone: (m) => {
          const snapshot = {
            source: sourceKey, name: m.meta.name, rows: m.meta.rows,
            columns: m.columns, sample: m.sample, sheet: m.meta.sheet,
            headerRow: m.meta.headerRow, loadedAt: Date.now()
          };
          // Persiste le snapshot léger ; renvoie les lignes complètes au caller
          // (Sprint 3+ : agrégation immédiate puis libération RAM).
          store.put("sources", sourceKey, snapshot)
            .then(() => { callbacks.onDone && callbacks.onDone(snapshot, m.rows); resolve({ snapshot: snapshot, rows: m.rows }); })
            .catch((e) => { callbacks.onError && callbacks.onError(e); reject(e); });
        }
      };
      // Transfert de l'ArrayBuffer au worker (zéro copie).
      _getWorker().postMessage({ id: id, name: file.name, source: sourceKey, buffer: buffer }, [buffer]);
    }));
  }

  function loadedSources() { return store.list("sources"); }
  function reset() { _listeners; return store.clear("sources"); }

  global.AGL = global.AGL || {};
  global.AGL.loader = { SOURCES, loadFile, loadedSources, memoryWarning, reset };
})(typeof window !== "undefined" ? window : this);
