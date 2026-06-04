/* ============================================================================
 * AGL BUDGET — prompt.js
 * Générateur de prompts Claude contextuels enrichis (§1 "0 coût additionnel").
 * Construit un prompt riche par thématique en injectant les chiffres et le
 * contexte macro à date juin 2026, prêt à coller dans claude.ai.
 *
 * Module pur — aucune dépendance UI. Expose window.AGL.prompts.
 * ==========================================================================*/
(function (global) {
  "use strict";
  const U = global.AGL.utils;

  const HEADER =
    "Tu es analyste stratégique senior (standard McKinsey / Roland Berger / BCG).\n" +
    "Contexte : AGL Côte d'Ivoire — direction stratégie — juin 2026.\n" +
    "Sources macro à date juin 2026 : OMC, Drewry, IATA, AfCFTA, Banque Mondiale, FMI.\n" +
    "Travaille en français. Sois concis, structure en SCR (Situation / Complication / Resolution).\n";

  function _fmt(n, dec) { return U.formatNumber(n, dec || 0); }
  function _pct(n, dec) { return ((+n || 0) * 100).toFixed(dec == null ? 1 : dec) + " %"; }

  /* ----- Onglet 02 — Marché STATCOM ----- */
  function market(snapshot, year) {
    return HEADER + "\n" +
      "Thème : analyse marché total CIV — métier par métier — année " + year + ".\n\n" +
      "Données STATCOM (volumes principaux) :\n" +
      snapshot.map((s) => "  - " + s.label + " : " + _fmt(s.value)).join("\n") + "\n\n" +
      "Demande :\n" +
      "1. Identifie 2 dynamiques structurantes 2023→" + year + " sur le marché CIV.\n" +
      "2. Replace ces volumes dans le contexte macro 2026 (OMC +1.9 %, AfCFTA +28 % d'ici 2030, IATA Afrique +18 % cargo).\n" +
      "3. Donne 1 recommandation actionnable pour AGL — format SCR — 5 lignes max.\n";
  }

  /* ----- Onglet 03 — Position AGL ----- */
  function position(pdmRows, year) {
    return HEADER + "\n" +
      "Thème : position AGL " + year + " — PDM par métier STATCOM.\n\n" +
      "PDM AGL :\n" +
      pdmRows.map((r) => "  - " + r.metier + " → PDM " + _pct(r.pdm) +
        " (vol AGL " + _fmt(r.volume_agl) + " / marché " + _fmt(r.volume_marche) + ")"
        + (r.cagr_agl != null ? " · CAGR AGL " + _pct(r.cagr_agl) : "")).join("\n") + "\n\n" +
      "Demande :\n" +
      "1. Quels métiers sont en position dominante à PROTÉGER, en croissance à PRIORISER, et fragilisés à RÉORIENTER ?\n" +
      "2. Compare la dynamique AGL à la dynamique marché (CAGR AGL vs CAGR Marché).\n" +
      "3. Produit 3 priorités d'allocation budget pour la PFA — chiffrées.\n";
  }

  /* ----- Onglet 10 — Projections 2026-2030 ----- */
  function projections(rowsForMetier, metier) {
    const base = rowsForMetier.filter((r) => r.SCENARIO === "BASE");
    const pess = rowsForMetier.filter((r) => r.SCENARIO === "PESSIMISTE");
    const opt  = rowsForMetier.filter((r) => r.SCENARIO === "OPTIMISTE");
    const fmtTraj = (rs) => rs.map((r) => r.ANNEE + " : " + _fmt(U.toNumber(r.VOLUME))).join("  |  ");
    return HEADER + "\n" +
      "Thème : projections 2026-2030 — " + metier + " — 3 scénarios pondérés.\n\n" +
      "BASE (P50, 65 %) — " + fmtTraj(base) + "\n" +
      "PESSIMISTE (P10, 15 %) — " + fmtTraj(pess) + "\n" +
      "OPTIMISTE (P90, 20 %) — " + fmtTraj(opt) + "\n\n" +
      "Demande :\n" +
      "1. Quels risques macro à date juin 2026 (Iran, surcapacité conteneurs, ralentissement OMC) peuvent activer le pessimiste ?\n" +
      "2. Quels triggers d'opportunité (AfCFTA $230B, IATA Afrique +18 %, PND CIV 8 400 Mds FCFA) activent l'optimiste ?\n" +
      "3. Donne la trajectoire CAP 2026-2030 espérée (E[V] pondéré par probabilités) et la décision PFA associée.\n";
  }

  /* ----- Onglet 07 — White Spaces ----- */
  function whitespaces(topRows) {
    return HEADER + "\n" +
      "Thème : white spaces — clients STATCOM hors CRM AGL.\n\n" +
      "Top 10 par volume non capturé :\n" +
      topRows.slice(0, 10).map((r, i) =>
        "  " + (i + 1) + ". " + (r.NOM_BASE || "?") + " (" + (r.METIER || "?") + ") · " +
        _fmt(U.toNumber(r.VOLUME_PRINCIPAL)) + " · statut " + (r.STATUT || "?")).join("\n") + "\n\n" +
      "Demande :\n" +
      "1. Estime le potentiel CAP FCFA récupérable (assomption : prix moyen TEU/Kg de marché).\n" +
      "2. Propose une séquence d'attaque commerciale (Top 3 cibles prioritaires + angle).\n" +
      "3. Identifie pour chaque cible un risque relationnel (concurrence installée, contrat captif).\n";
  }

  /* ----- Onglet 11 — Préconisations Client ----- */
  function clientFiche(c) {
    return HEADER + "\n" +
      "Thème : fiche stratégique client.\n\n" +
      "Client : " + (c.nom || "?") + "\n" +
      "Secteur : " + (c.secteur || "?") + " (" + (c.source_secteur || "?") + ")\n" +
      "IDs : CRM=" + (c.id_crm || "—") + " · IRIS=" + (c.id_iris || "—") + " · STATCOM=" + (c.id_statcom || "—") + "\n" +
      "Volume total 3 ans : " + _fmt(c.volume_total) + "\n" +
      "Métier principal : " + c.metier_principal + "\n" +
      "Position BCG : " + c.bcg + "\n" +
      "PDM (métier principal) : " + (c.pdm_principal != null ? _pct(c.pdm_principal) : "—") +
        " · CAGR client : " + (c.cagr_principal != null ? _pct(c.cagr_principal) : "—") + "\n\n" +
      "Demande :\n" +
      "1. Diagnostic en 3 points : santé commerciale, dépendance, exposition macro.\n" +
      "2. 3 actions concrètes pour les 12 prochains mois.\n" +
      "3. Un risque à surveiller + 1 KPI mensuel.\n";
  }

  /* ----- Onglet 12 — Synthèse Executive ----- */
  function executive(snap) {
    let s = HEADER + "\n" +
      "Thème : NOTE COMEX — synthèse stratégique AGL CIV " + (snap.derniere_annee || "") + ".\n\n";
    if (snap.marche_total != null) s += "Marché total CIV (vol principal cumulé) : " + _fmt(snap.marche_total) + "\n";
    if (snap.pdm_globale != null)  s += "PDM AGL globale : " + _pct(snap.pdm_globale) + "\n";
    if (snap.clients_rmc != null)  s += "Clients RMC : " + _fmt(snap.clients_rmc) + " (locked " + (snap.clients_locked || 0) + ")\n";
    if (snap.top_croissance && snap.top_croissance.length) {
      s += "\nTop 3 métiers en croissance :\n";
      snap.top_croissance.forEach((m) => s += "  - " + m.metier + " · CAGR marché " + _pct(m.cagr) + "\n");
    }
    if (snap.top_white_spaces && snap.top_white_spaces.length) {
      s += "\nTop 5 white spaces :\n";
      snap.top_white_spaces.forEach((w) => s += "  - " + w.client + " (" + w.metier + ") · " + _fmt(w.volume) + "\n");
    }
    if (snap.projection_base_2030 && snap.projection_base_2030.length) {
      s += "\nProjection BASE " + snap.horizon + " (volumes attendus) :\n";
      snap.projection_base_2030.forEach((p) => s += "  - " + p.metier + " : " + _fmt(p.volume) + " (+" + _pct(p.cagr) + "/an)\n");
    }
    s += "\nRisques & opportunités macro à date juin 2026 :\n";
    (snap.risques_opportunites || []).slice(0, 6).forEach((x) =>
      s += "  - [" + x.type + "] " + x.titre + " — " + x.detail + "\n");
    s += "\nDemande :\n" +
      "1. Rédige la note SCR (Situation / Complication / Resolution) en 1 page.\n" +
      "2. Donne 3 décisions à acter en COMEX avec deadline.\n" +
      "3. Identifie le risque #1 à mettre sous surveillance immédiate.\n";
    return s;
  }

  function copyToClipboard(text) {
    if (navigator.clipboard) return navigator.clipboard.writeText(text);
    return new Promise((resolve, reject) => {
      const ta = document.createElement("textarea");
      ta.value = text; document.body.appendChild(ta); ta.select();
      try { document.execCommand("copy"); resolve(); } catch (e) { reject(e); }
      finally { ta.remove(); }
    });
  }

  global.AGL = global.AGL || {};
  global.AGL.prompts = { market, position, projections, whitespaces, clientFiche, executive, copyToClipboard };
})(typeof window !== "undefined" ? window : this);
