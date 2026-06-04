/* ============================================================================
 * AGL BUDGET — pipeline.js
 * Chargement des artefacts produits par le pipeline Python (sorties CSV/XLSX
 * de 05_build_rmc, 06_agreg_statcom, 07_calcul_pdm). Persiste les tables dans
 * IndexedDB store "aggregates" et les expose en mémoire pour les onglets.
 *
 * Dépend de : AGL.utils, AGL.store, SheetJS (XLSX) déjà chargé.
 * ==========================================================================*/
(function (global) {
  "use strict";
  const U = global.AGL.utils, store = global.AGL.store;

  // Artefacts attendus (clé store ↔ noms de fichiers reconnus).
  const ARTEFACTS = {
    RMC:           { match: /rmc\.(xlsx|csv)$/i,           label: "RMC (référentiel maître)" },
    MARCHE_TOTAL:  { match: /marche_total\.csv$/i,         label: "Marché total STATCOM" },
    VOLUME_CLIENT: { match: /volume_client\.csv$/i,        label: "Volumes par client" },
    METIERS_3ANS:  { match: /metiers_3ans\.csv$/i,         label: "Pivot métiers 3 ans + CAGR" },
    PDM:           { match: /pdm\.csv$/i,                  label: "PDM AGL (part de marché)" },
    PDM_CLIENTS:   { match: /pdm_clients\.csv$/i,          label: "Volumes clients AGL" },
    MATCH_CRM_IRIS:{ match: /match_crm_iris\.csv$/i,       label: "Flux 1 — matches CRM × IRIS" },
    MATCH_CRM_STAT:{ match: /match_crm_stat\.csv$/i,       label: "Flux 2 — matches CRM × STATCOM" },
    PROJECTIONS:   { match: /projections\.csv$/i,          label: "Projections 2026-2030" },
    WHITESPACES:   { match: /whitespaces\.csv$/i,          label: "White spaces" },
    PIPELINE_CRM:  { match: /opportunites.*\.(xlsx|csv)$/i, label: "Pipeline CRM — opportunités" },
    PND_PROJETS:   { match: /projets_pnd\.csv$/i,           label: "PND — projets (Actions/Produits/Effets)" },
    PND_AXES:      { match: /axes_pnd\.csv$/i,              label: "PND — agrégat par axe" },
    PND_SECTORIELS:{ match: /sectoriels_pnd\.csv$/i,        label: "PND — agrégat par résultat sectoriel" },
  };

  // Cache mémoire des tables désérialisées.
  const _cache = {};

  function detectArtefact(filename) {
    for (const k in ARTEFACTS) if (ARTEFACTS[k].match.test(filename)) return k;
    return null;
  }

  function parseFile(file) {
    return file.arrayBuffer().then((buf) => {
      const wb = XLSX.read(new Uint8Array(buf), { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });
      const columns = rows.length ? Object.keys(rows[0]) : [];
      return { rows, columns };
    });
  }

  function loadFiles(fileList, onProgress) {
    const files = Array.from(fileList);
    const results = [];
    let i = 0;
    function next() {
      if (i >= files.length) return Promise.resolve(results);
      const f = files[i++];
      const key = detectArtefact(f.name);
      if (!key) { results.push({ name: f.name, status: "ignored" }); onProgress && onProgress(i / files.length, f.name, "ignored"); return next(); }
      return parseFile(f).then(({ rows, columns }) => {
        _cache[key] = rows;
        return store.put("aggregates", key, { name: f.name, columns: columns, rows: rows, loadedAt: Date.now() })
          .then(() => { results.push({ name: f.name, key: key, rows: rows.length, status: "ok" }); onProgress && onProgress(i / files.length, f.name, "ok"); return next(); });
      }).catch((err) => { results.push({ name: f.name, key: key, status: "error", error: err.message }); onProgress && onProgress(i / files.length, f.name, "error"); return next(); });
    }
    return next();
  }

  function get(key) {
    if (_cache[key]) return Promise.resolve(_cache[key]);
    return store.get("aggregates", key).then((v) => {
      if (!v) return null;
      _cache[key] = v.rows; return v.rows;
    });
  }

  function listLoaded() { return store.list("aggregates"); }
  function reset() { for (const k in _cache) delete _cache[k]; return store.clear("aggregates"); }

  /* ----- Helpers analytiques sur les tables chargées ----- */
  function toNumber(v) { return U.toNumber(v); }

  // marche_total → [{METIER, ANNEE, VOLUME_TEU, VOLUME_KG, NB_CLIENTS}]
  function marketByMetierYear() {
    return get("MARCHE_TOTAL").then((rows) => rows || []);
  }

  function marketTotalsByMetier(year) {
    return marketByMetierYear().then((rows) => {
      const filtered = rows.filter((r) => !year || +r.ANNEE === +year);
      const byM = {};
      filtered.forEach((r) => {
        const m = r.METIER || "?";
        const isAir = /AERIEN|AÉRIEN/i.test(m);
        const v = toNumber(isAir ? r.VOLUME_KG : r.VOLUME_TEU);
        byM[m] = (byM[m] || 0) + v;
      });
      return Object.keys(byM).map((m) => ({ label: m, value: byM[m] }));
    });
  }

  function pdmByMetier(year) {
    return get("PDM").then((rows) => {
      if (!rows) return [];
      const filtered = rows.filter((r) => !year || +r.ANNEE === +year);
      return filtered.map((r) => ({
        metier: r.METIER, annee: +r.ANNEE,
        volume_agl: toNumber(r.VOLUME_AGL),
        volume_marche: toNumber(r.VOLUME_MARCHE),
        pdm: toNumber(r.PDM),
        cagr_agl: r.CAGR_AGL != null && r.CAGR_AGL !== "" ? toNumber(r.CAGR_AGL) : null,
        cagr_marche: r.CAGR_MARCHE != null && r.CAGR_MARCHE !== "" ? toNumber(r.CAGR_MARCHE) : null,
      }));
    });
  }

  function availableYears() {
    return get("MARCHE_TOTAL").then((rows) => {
      if (!rows) return [];
      return Array.from(new Set(rows.map((r) => +r.ANNEE))).filter((y) => y).sort();
    });
  }

  global.AGL = global.AGL || {};
  global.AGL.pipeline = {
    ARTEFACTS, detectArtefact, loadFiles, get, listLoaded, reset,
    marketByMetierYear, marketTotalsByMetier, pdmByMetier, availableYears
  };
})(typeof window !== "undefined" ? window : this);
