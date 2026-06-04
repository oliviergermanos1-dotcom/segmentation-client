/* ============================================================================
 * AGL BUDGET — utils.js
 * Helpers transverses : DOM, formatage (FCFA / TEU / Kg / %), parsing nombres,
 * normalisation basique de noms (fallback JS), logging.
 * Module pur, aucune dépendance externe. Expose window.AGL.utils.
 * ==========================================================================*/
(function (global) {
  "use strict";

  /* ----- DOM ----- */
  const $ = (sel, root) => (root || document).querySelector(sel);
  const $$ = (sel, root) => Array.from((root || document).querySelectorAll(sel));
  const el = (tag, attrs, children) => {
    const node = document.createElement(tag);
    if (attrs) for (const k in attrs) {
      if (k === "class") node.className = attrs[k];
      else if (k === "html") node.innerHTML = attrs[k];
      else if (k === "text") node.textContent = attrs[k];
      else if (k.startsWith("on") && typeof attrs[k] === "function") node.addEventListener(k.slice(2), attrs[k]);
      else node.setAttribute(k, attrs[k]);
    }
    (Array.isArray(children) ? children : children != null ? [children] : [])
      .forEach((c) => node.appendChild(typeof c === "string" ? document.createTextNode(c) : c));
    return node;
  };

  /* ----- Parsing nombres (gère "1 234,56", "1,234.56", "12%", espaces fines) ----- */
  function toNumber(v) {
    if (v == null || v === "") return 0;
    if (typeof v === "number") return isFinite(v) ? v : 0;
    let s = String(v).trim().replace(/ | |\s/g, "").replace(/%$/, "");
    // Si virgule ET point : le dernier séparateur est le décimal.
    if (s.indexOf(",") > -1 && s.indexOf(".") > -1) {
      s = s.lastIndexOf(",") > s.lastIndexOf(".")
        ? s.replace(/\./g, "").replace(",", ".")
        : s.replace(/,/g, "");
    } else if (s.indexOf(",") > -1) {
      s = s.replace(",", ".");
    }
    const n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }

  /* ----- Formatage ----- */
  const _nf = (min, max) => new Intl.NumberFormat("fr-FR", { minimumFractionDigits: min, maximumFractionDigits: max });
  function formatNumber(n, dec) { return _nf(dec || 0, dec || 0).format(toNumber(n)); }
  function formatFCFA(n) {
    const v = toNumber(n);
    const a = Math.abs(v);
    if (a >= 1e9) return _nf(0, 2).format(v / 1e9) + " Mds FCFA";
    if (a >= 1e6) return _nf(0, 1).format(v / 1e6) + " M FCFA";
    if (a >= 1e3) return _nf(0, 0).format(v / 1e3) + " k FCFA";
    return _nf(0, 0).format(v) + " FCFA";
  }
  function formatTEU(n) { return _nf(0, 0).format(toNumber(n)) + " TEU"; }
  function formatKg(n) {
    const v = toNumber(n);
    if (Math.abs(v) >= 1000) return _nf(0, 1).format(v / 1000) + " t";
    return _nf(0, 0).format(v) + " kg";
  }
  function formatPct(n, dec) { return _nf(dec || 1, dec || 1).format(toNumber(n)) + " %"; }
  function formatBytes(b) {
    if (b >= 1e9) return (b / 1e9).toFixed(2) + " Go";
    if (b >= 1e6) return (b / 1e6).toFixed(1) + " Mo";
    if (b >= 1e3) return (b / 1e3).toFixed(0) + " ko";
    return b + " o";
  }

  /* ----- Normalisation basique de noms (fallback léger côté navigateur).
   * La normalisation de référence se fait en Python (01_normalise.py / Cleanco).
   * Ceci ne sert qu'à l'affichage et aux comparaisons rapides dans l'UI. ----- */
  const FORMES_JURIDIQUES = [
    "SARL", "SARLU", "SA", "SAS", "SASU", "GIE", "SNC", "SCS", "SCI",
    "ETS", "ETABLISSEMENTS", "CIE", "COMPAGNIE", "STE", "SOCIETE",
    "NV", "BV", "LTD", "LLC", "INC", "PLC", "GMBH", "CO", "GROUP", "GROUPE"
  ];
  const GEO_STOP = ["COTE", "IVOIRE", "CI", "CIV", "ABIDJAN", "AFRIQUE", "WEST", "AFRICA"];
  function normalizeName(name) {
    if (!name) return "";
    let s = String(name).toUpperCase();
    s = s.normalize("NFD").replace(/[̀-ͯ]/g, "");      // accents
    s = s.replace(/[^A-Z0-9 ]+/g, " ");                          // ponctuation
    let toks = s.split(/\s+/).filter(Boolean)
      .filter((t) => FORMES_JURIDIQUES.indexOf(t) === -1)
      .filter((t) => GEO_STOP.indexOf(t) === -1);
    return toks.join(" ").trim();
  }

  /* ----- Divers ----- */
  function debounce(fn, ms) {
    let t; return function () { const a = arguments, c = this; clearTimeout(t); t = setTimeout(() => fn.apply(c, a), ms || 200); };
  }
  function downloadBlob(filename, blob) {
    const url = URL.createObjectURL(blob);
    const a = el("a", { href: url, download: filename });
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 0);
  }
  function log(scope, msg, level) {
    const tag = "[AGL:" + scope + "]";
    (console[level || "log"] || console.log)(tag, msg);
  }

  global.AGL = global.AGL || {};
  global.AGL.utils = {
    $, $$, el, toNumber, formatNumber, formatFCFA, formatTEU, formatKg,
    formatPct, formatBytes, normalizeName, debounce, downloadBlob, log,
    FORMES_JURIDIQUES, GEO_STOP
  };
})(typeof window !== "undefined" ? window : this);
