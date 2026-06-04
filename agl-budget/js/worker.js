/* ============================================================================
 * AGL BUDGET — worker.js  (Web Worker)
 * Parsing Excel/CSV en arrière-plan — zéro freeze UI (§5, pattern STATCOM v23).
 * Reçoit un ArrayBuffer + métadonnées, parse avec SheetJS, détecte la ligne
 * d'en-tête, envoie la progression par blocs de CHUNK lignes, puis renvoie
 * les lignes parsées (et un échantillon pour l'aperçu UI).
 *
 * Protocole messages (postMessage) :
 *   IN  : { id, name, source, buffer (ArrayBuffer transférable) }
 *   OUT : { id, type:"progress", phase, pct, rows }
 *         { id, type:"done", meta:{...}, columns:[...], sample:[...], rows:[...] }
 *         { id, type:"error", message }
 * ==========================================================================*/
/* global importScripts, XLSX */
"use strict";

// SheetJS chargé dans le worker (même version que la page).
try {
  importScripts("https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js");
} catch (e) {
  // Échec CDN : signalé au premier message.
}

const CHUNK = 10000; // §5 : blocs de 10 000 lignes.

// Heuristique de détection de la vraie ligne d'en-tête (saute "Filtres appliqués",
// "Exported data limited to..." etc.). Retourne l'index de ligne d'en-tête.
function detectHeaderRow(aoa) {
  const maxScan = Math.min(aoa.length, 15);
  let best = 0, bestScore = -1;
  for (let i = 0; i < maxScan; i++) {
    const row = aoa[i] || [];
    const nonEmpty = row.filter((c) => c !== null && c !== undefined && String(c).trim() !== "");
    // En-tête probable : beaucoup de cellules, majoritairement du texte court non numérique.
    const textCells = nonEmpty.filter((c) => typeof c === "string" && isNaN(Number(c)) && String(c).length < 60);
    const score = nonEmpty.length >= 3 ? textCells.length : -1;
    if (score > bestScore) { bestScore = score; best = i; }
  }
  return best;
}

self.onmessage = function (ev) {
  const msg = ev.data || {};
  const id = msg.id;
  const post = (o) => self.postMessage(Object.assign({ id: id }, o));

  if (typeof XLSX === "undefined") {
    post({ type: "error", message: "SheetJS (XLSX) non chargé dans le worker — vérifier l'accès réseau au CDN." });
    return;
  }

  try {
    post({ type: "progress", phase: "lecture", pct: 5, rows: 0 });
    const wb = XLSX.read(new Uint8Array(msg.buffer), { type: "array", cellDates: true });
    const sheetName = wb.SheetNames[0];
    const sheet = wb.Sheets[sheetName];

    post({ type: "progress", phase: "analyse", pct: 20, rows: 0 });
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, blankrows: false });
    if (!aoa.length) { post({ type: "done", meta: { name: msg.name, source: msg.source, rows: 0, sheet: sheetName }, columns: [], sample: [], rows: [] }); return; }

    const hIdx = detectHeaderRow(aoa);
    const columns = (aoa[hIdx] || []).map((c, i) => (c == null || String(c).trim() === "") ? ("col_" + i) : String(c).trim());
    const dataRows = aoa.slice(hIdx + 1);
    const total = dataRows.length;

    // Construction objet ligne par ligne, en streamant la progression.
    const rows = new Array(total);
    for (let i = 0; i < total; i++) {
      const r = dataRows[i];
      const obj = {};
      for (let c = 0; c < columns.length; c++) obj[columns[c]] = r ? r[c] : null;
      rows[i] = obj;
      if (i % CHUNK === 0) {
        post({ type: "progress", phase: "parsing", pct: 20 + Math.round((i / Math.max(total, 1)) * 75), rows: i });
      }
    }

    post({ type: "progress", phase: "finalisation", pct: 98, rows: total });
    post({
      type: "done",
      meta: { name: msg.name, source: msg.source, rows: total, sheet: sheetName, headerRow: hIdx },
      columns: columns,
      sample: rows.slice(0, 50),
      rows: rows
    });
  } catch (err) {
    post({ type: "error", message: (err && err.message) || String(err) });
  }
};
