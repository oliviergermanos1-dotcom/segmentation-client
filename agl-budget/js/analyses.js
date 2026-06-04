/* ============================================================================
 * AGL BUDGET — analyses.js
 * Calculs analytiques côté navigateur (cohortes/LTV/churn, BCG, préconisations).
 * Travaille sur les tables déjà chargées via pipeline.js (VOLUME_CLIENT, PDM, RMC).
 * Module pur — pas de DOM. Expose window.AGL.analyses.
 * ==========================================================================*/
(function (global) {
  "use strict";
  const U = global.AGL.utils, pipe = global.AGL.pipeline;

  function _principal(row) {
    const isAir = /AERIEN|AÉRIEN/i.test(row.METIER || "");
    return U.toNumber(isAir ? row.VOLUME_KG : row.VOLUME_TEU);
  }

  /* ----- Cohortes & rétention (onglet 14 — §1) -----
   * Cohorte = année de première apparition d'un client (tous métiers confondus).
   * Rétention(année N) = % de la cohorte encore présent en année N.
   * Churn(année N) = % des clients présents en N-1 absents en N.
   * LTV proxy = somme du volume principal du client sur la période.
   * --------------------------------------------------------------------- */
  function cohortes() {
    return pipe.get("VOLUME_CLIENT").then((rows) => {
      if (!rows || !rows.length) return null;
      // Map client -> set d'années + total volume.
      const byClient = {};
      rows.forEach((r) => {
        const c = r.NOM_BASE || "?";
        const y = +r.ANNEE;
        if (!byClient[c]) byClient[c] = { years: new Set(), volume: 0 };
        byClient[c].years.add(y);
        byClient[c].volume += _principal(r);
      });
      const years = Array.from(new Set(rows.map((r) => +r.ANNEE))).filter((y) => y).sort();
      const ymin = years[0], ymax = years[years.length - 1];

      // Détermine la cohorte (1ère année observée) par client.
      const cohorts = {};
      for (const c in byClient) {
        const first = Math.min.apply(null, Array.from(byClient[c].years));
        cohorts[first] = cohorts[first] || [];
        cohorts[first].push({ client: c, years: byClient[c].years, volume: byClient[c].volume });
      }

      // Matrice rétention : cohort × année observée → %.
      const retention = [];
      Object.keys(cohorts).sort().forEach((co) => {
        const start = +co;
        const size = cohorts[co].length;
        const row = { cohorte: start, taille_initiale: size };
        for (let y = start; y <= ymax; y++) {
          const present = cohorts[co].filter((x) => x.years.has(y)).length;
          row["y" + y] = present;
          row["r" + y] = size ? present / size : 0;
        }
        retention.push(row);
      });

      // Churn année par année (toutes cohortes confondues).
      const churn = [];
      for (let i = 1; i < years.length; i++) {
        const yPrev = years[i - 1], yNow = years[i];
        const setPrev = new Set(rows.filter((r) => +r.ANNEE === yPrev).map((r) => r.NOM_BASE));
        const setNow  = new Set(rows.filter((r) => +r.ANNEE === yNow).map((r) => r.NOM_BASE));
        const lost = Array.from(setPrev).filter((c) => !setNow.has(c)).length;
        const gained = Array.from(setNow).filter((c) => !setPrev.has(c)).length;
        churn.push({
          annee: yNow, base_n1: setPrev.size, base_n: setNow.size,
          perdus: lost, gagnes: gained,
          churn: setPrev.size ? lost / setPrev.size : 0
        });
      }

      // LTV (3 ans) — top clients par volume cumulé.
      const ltv = Object.keys(byClient)
        .map((c) => ({ client: c, ltv_volume: byClient[c].volume,
                       annees_actives: byClient[c].years.size }))
        .sort((a, b) => b.ltv_volume - a.ltv_volume);

      return { years, retention, churn, ltv, cohorts };
    });
  }

  /* ----- Matrice BCG (onglet 06 — §1) -----
   * 9 cases : axe X = part de marché client dans son métier (PDM client),
   *           axe Y = croissance volume client 3 ans (CAGR).
   * Seuils : 5% / 15% PDM (basse/haute), 0% / 10% CAGR (basse/haute).
   * --------------------------------------------------------------------- */
  function bcg() {
    return Promise.all([pipe.get("VOLUME_CLIENT"), pipe.get("MARCHE_TOTAL")])
      .then(([vc, marche]) => {
        if (!vc || !marche) return null;
        const years = Array.from(new Set(vc.map((r) => +r.ANNEE))).filter((y) => y).sort();
        const ymin = years[0], ymax = years[years.length - 1];
        // Total marché par métier × année.
        const totalMetier = {};
        marche.forEach((r) => {
          const k = (r.METIER || "") + "|" + r.ANNEE;
          totalMetier[k] = U.toNumber(/AERIEN|AÉRIEN/i.test(r.METIER || "") ? r.VOLUME_KG : r.VOLUME_TEU);
        });
        // Volume client × métier × année.
        const byCM = {};
        vc.forEach((r) => {
          const k = (r.NOM_BASE || "") + "|" + (r.METIER || "");
          if (!byCM[k]) byCM[k] = { client: r.NOM_BASE, metier: r.METIER, byYear: {} };
          byCM[k].byYear[+r.ANNEE] = (byCM[k].byYear[+r.ANNEE] || 0) + _principal(r);
        });
        const cells = [];
        for (const k in byCM) {
          const o = byCM[k];
          const v0 = o.byYear[ymin] || 0;
          const v1 = o.byYear[ymax] || 0;
          const period = (ymax - ymin) || 1;
          const cagr = (v0 > 0 && v1 > 0) ? Math.pow(v1 / v0, 1 / period) - 1 : null;
          const total = totalMetier[(o.metier || "") + "|" + ymax] || 0;
          const pdm = total > 0 ? v1 / total : 0;
          // Classification 9 cases (faible / moyen / fort × cagr neg / stable / fort).
          const pdmClass = pdm < 0.05 ? "F" : pdm < 0.15 ? "M" : "H";
          const cagrClass = cagr == null || cagr < 0 ? "L" : cagr < 0.10 ? "M" : "H";
          const matrix = {
            "HH": "★ STAR",       "HM": "VACHE",       "HL": "VACHE-décline",
            "MH": "DILEMME+",     "MM": "DILEMME",     "ML": "DILEMME-",
            "FH": "POIDS MORT+",  "FM": "POIDS MORT",  "FL": "POIDS MORT-"
          };
          cells.push({
            client: o.client, metier: o.metier,
            volume: v1, pdm: pdm, cagr: cagr,
            cellule: matrix[pdmClass + cagrClass] || "?"
          });
        }
        return cells.sort((a, b) => b.volume - a.volume);
      });
  }

  /* ----- Préconisations client (onglet 11) -----
   * Pour chaque client du RMC, agrège : secteur, IRIS/STATCOM, volumes 3 ans
   * par métier, CAGR, tag BCG. Génère une recommandation simple.
   * --------------------------------------------------------------------- */
  function preconisations() {
    return Promise.all([pipe.get("RMC"), pipe.get("VOLUME_CLIENT"), pipe.get("PDM"), bcg()])
      .then(([rmc, vc, pdm, bcgCells]) => {
        if (!rmc) return null;
        // Index volumes par client.
        const volsByClient = {};
        (vc || []).forEach((r) => {
          const c = r.NOM_BASE || ""; if (!c) return;
          if (!volsByClient[c]) volsByClient[c] = { totalVol: 0, byMetier: {} };
          const v = _principal(r);
          volsByClient[c].totalVol += v;
          volsByClient[c].byMetier[r.METIER || "?"] = (volsByClient[c].byMetier[r.METIER || "?"] || 0) + v;
        });
        // Index BCG par client + métier le plus volumineux.
        const bcgByClient = {};
        (bcgCells || []).forEach((c) => {
          if (!bcgByClient[c.client] || c.volume > bcgByClient[c.client].volume)
            bcgByClient[c.client] = c;
        });

        return rmc.map((r) => {
          const alias = (r.ALIAS_STATCOM || r.ALIAS_IRIS || r.ALIAS_CRM || "").toString().toUpperCase().trim();
          const vol = volsByClient[alias] || { totalVol: 0, byMetier: {} };
          const bcg = bcgByClient[alias];
          // Recommandation simple basée sur cellule BCG + secteur.
          const reco = _reco(r, bcg, vol);
          return {
            id_rmc: r.ID_RMC, id_crm: r.ID_CRM, id_iris: r.ID_IRIS, id_statcom: r.ID_STATCOM,
            nom: r.NOM_CANONIQUE, secteur: r.SECTEUR, source_secteur: r.SOURCE_SECTEUR,
            locked: r.LOCKED === true || r.LOCKED === "True" || r.LOCKED === "true",
            iris_score: r.SCORE_MATCH_IRIS, statcom_score: r.SCORE_MATCH_STATCOM,
            volume_total: vol.totalVol,
            metier_principal: Object.keys(vol.byMetier).sort((a, b) => vol.byMetier[b] - vol.byMetier[a])[0] || "—",
            bcg: bcg ? bcg.cellule : "—",
            pdm_principal: bcg ? bcg.pdm : null,
            cagr_principal: bcg ? bcg.cagr : null,
            recommandation: reco
          };
        });
      });
  }
  function _reco(rmc, bcg, vol) {
    if (!bcg) {
      if (!rmc.ID_IRIS) return "AGL n'opère pas ce client : qualifier le potentiel via secteur " + (rmc.SECTEUR || "?") + ".";
      return "Client actif sans donnée STATCOM exploitable — vérifier rattachement marché.";
    }
    const cell = (bcg.cellule || "").split(" ")[0];
    if (cell === "★") return "ÉTOILE — protéger la position : programme dédié, contrats long terme, account team renforcée.";
    if (cell === "VACHE") return "VACHE — maximiser la marge : industrialiser les opérations, contenir les coûts.";
    if (cell.startsWith("DILEMME"))
      return "DILEMME — décider : investir agressivement pour gagner la PDM, ou désengager si CAC trop élevé.";
    if (cell.startsWith("POIDS"))
      return "POIDS MORT — désengagement progressif sauf valeur stratégique sectorielle.";
    return "À évaluer.";
  }

  /* ----- Synthèse Executive (onglet 12) -----
   * Agrégats COMEX : marché total, PDM AGL globale, top métiers en croissance,
   * top white spaces, risques/opportunités macro. Format SCR McKinsey.
   * --------------------------------------------------------------------- */
  function executive() {
    return Promise.all([pipe.get("PDM"), pipe.get("MARCHE_TOTAL"),
                        pipe.get("WHITESPACES"), pipe.get("PROJECTIONS"),
                        pipe.get("RMC")])
      .then(([pdm, marche, ws, proj, rmc]) => {
        const result = { generated_at: new Date().toISOString() };
        if (marche && marche.length) {
          const years = Array.from(new Set(marche.map((r) => +r.ANNEE))).sort();
          const lastY = years[years.length - 1];
          const totalMarche = marche.filter((r) => +r.ANNEE === lastY)
            .reduce((a, r) => a + U.toNumber(/AERIEN|AÉRIEN/i.test(r.METIER || "") ? r.VOLUME_KG : r.VOLUME_TEU), 0);
          result.derniere_annee = lastY;
          result.marche_total = totalMarche;
        }
        if (pdm && pdm.length) {
          const lastY = Math.max.apply(null, pdm.map((r) => +r.ANNEE));
          const snap = pdm.filter((r) => +r.ANNEE === lastY);
          const totalAGL = snap.reduce((a, r) => a + U.toNumber(r.VOLUME_AGL), 0);
          const totalM   = snap.reduce((a, r) => a + U.toNumber(r.VOLUME_MARCHE), 0);
          result.pdm_globale = totalM > 0 ? totalAGL / totalM : 0;
          result.top_croissance = snap.slice()
            .filter((r) => r.CAGR_MARCHE != null && r.CAGR_MARCHE !== "")
            .sort((a, b) => U.toNumber(b.CAGR_MARCHE) - U.toNumber(a.CAGR_MARCHE))
            .slice(0, 3)
            .map((r) => ({ metier: r.METIER, cagr: U.toNumber(r.CAGR_MARCHE) }));
        }
        if (ws && ws.length) {
          result.top_white_spaces = ws.slice()
            .sort((a, b) => U.toNumber(b.VOLUME_PRINCIPAL) - U.toNumber(a.VOLUME_PRINCIPAL))
            .slice(0, 5)
            .map((r) => ({ client: r.NOM_BASE, metier: r.METIER, volume: U.toNumber(r.VOLUME_PRINCIPAL) }));
        }
        if (proj && proj.length) {
          // Projection 2030 BASE par métier.
          const targetY = Math.max.apply(null, proj.map((r) => +r.ANNEE));
          result.horizon = targetY;
          result.projection_base_2030 = proj.filter((r) => r.SCENARIO === "BASE" && +r.ANNEE === targetY)
            .map((r) => ({ metier: r.METIER, volume: U.toNumber(r.VOLUME), cagr: U.toNumber(r.CAGR_APPLIQUE) }));
        }
        if (rmc && rmc.length) {
          result.clients_rmc = rmc.length;
          result.clients_locked = rmc.filter((r) => r.LOCKED === true || r.LOCKED === "True" || r.LOCKED === "true").length;
        }
        // §7 risques/opportunités : on injecte ceux du macro_data.
        result.risques_opportunites = global.MACRO_DATA.risques_opportunites;
        return result;
      });
  }

  global.AGL = global.AGL || {};
  global.AGL.analyses = { cohortes, bcg, preconisations, executive };
})(typeof window !== "undefined" ? window : this);
