/* ============================================================================
 * AGL BUDGET — macro_data.js
 * Sources macro mondiales & CIV — À DATE JUIN 2026 (cf. cahier des charges §7)
 * Données figées, sourcées et datées. Servent de Couche 2 (multiplicateurs)
 * et d'hypothèses pour les 3 scénarios de projection (Couche 3).
 *
 * Module pur : aucune dépendance. Expose window.MACRO_DATA.
 * 1 module = 1 fichier (règle "monolithe interdit").
 * ==========================================================================*/
(function (global) {
  "use strict";

  // Helper de structure d'un indicateur sourcé.
  const M = (indicateur, valeur, source, date, note) => ({
    indicateur, valeur, source, date, note: note || ""
  });

  const MACRO_DATA = {
    meta: {
      version: "2.0",
      a_date: "2026-06",
      confidentiel: true,
      note: "Sources institutionnelles mondiales à date juin 2026."
    },

    /* ----- §7 Macro Côte d'Ivoire ----- */
    civ: [
      M("Croissance PIB CIV 2025", "+6.2%", "Banque Mondiale", "2025-09"),
      M("Croissance PIB CIV 2026E", "+6.4%", "FMI", "2025-12"),
      M("Transport CIV croissance", "+9.3%", "Rapport Éco. Suisse", "2025"),
      M("Port maritime CIV", "+12%", "Rapport Éco. Suisse", "2025"),
      M("Trafic aérien CIV passagers", "+20%", "Rapport Éco. Suisse", "2025"),
      M("CIV — cacao PDM mondiale", "39%", "Trésor.fr", "2025"),
      M("CIV — anacarde production", "40%", "Trésor.fr", "2025"),
      M("Secteur tertiaire CIV", "+8.6%", "Rapport Éco. Suisse", "2025")
    ],

    /* ----- §7 Commerce mondial — OMC ----- */
    commerce_mondial: [
      M("Commerce marchandises 2025", "+4.6%", "OMC Global Trade Outlook", "2026-03"),
      M("Commerce marchandises 2026E", "+1.9%", "OMC", "2026-03", "ralentissement"),
      M("Commerce marchandises 2027E", "+2.6%", "OMC", "2026-03"),
      M("Commerce services 2026E", "+4.8%", "OMC", "2026-03"),
      M("PIB mondial 2026E", "+2.8%", "OMC", "2026-03"),
      M("Afrique + Lat. Am. cargo 2026", "+3.0% et +4.0%", "IATA/OMC", "2025", "> moyenne")
    ],

    /* ----- §7 Maritime conteneurs ----- */
    maritime_conteneurs: [
      M("Taux fret conteneurs 2026E", "-17% vs 2025", "Drewry", "2026-01"),
      M("Demande conteneurs 2026", "+1.8%", "Drewry", "2026-01"),
      M("Demande conteneurs 2027+", "+2-3%/an", "Drewry", "2026-01", "rythme historique"),
      M("Throughput mondial 2025", "+5%", "Drewry", "2026-01", "1er Mrd TEUs dépassé"),
      M("Fleet Afrique nouvelles livr.", "70-80% newbuilds → marchés sec.", "SeaRates/Drewry", "2025-11"),
      M("Demande conteneurs Afrique", "+10-15%/an", "SeaRates", "2025-11"),
      M("Rally taux fret mai 2026", "3ème semaine hausse consécutive", "Drewry", "2026-05-21"),
      M("Blank sailings sem. 22-26", "41 sur 700 départs", "Drewry", "2026-05-22")
    ],

    /* ----- §7 Aérien cargo ----- */
    aerien_cargo: [
      M("Cargo aérien mondial 2026E", "+2.4%", "IATA", "2026-01", "modération"),
      M("Cargo aérien Afrique 2025", "+6% demande, +7.8% capacité", "IATA", "2026-04"),
      M("Cargo aérien Afrique jan. 2026", "+18%", "IATA", "2026-04", "région la + dynamique"),
      M("Cargo aérien mondial fév. 2026", "+11.2% tous corridors", "IATA", "2026-03"),
      M("Afrique % marché mondial", "2.1%", "Aviation Week", "2026-03", "croissance 8-10%/an"),
      M("Drivers Afrique 2026", "E-com, pharma, perishables, énergie", "Air Cargo News", "2026-04"),
      M("Taux aérien impact Iran", "+95% fév-mars 2026", "Drewry", "2026-04")
    ],

    /* ----- §7 Hinterland & AfCFTA ----- */
    hinterland_afcfta: [
      M("Commerce intra-africain 2026E", "$230 milliards", "African Trade Outlook", "2026-05"),
      M("AfCFTA hausse fret d'ici 2030", "+28%", "Brookings", "2026-05"),
      M("Maritime intra-Afrique 2030", "58M → 132M tonnes", "WEF", "2023"),
      M("AfCFTA potentiel commerce", "+52.3% d'ici 2035", "Market Data Forecast", "2026-01"),
      M("Camions Afrique supplémentaires", "2 millions d'ici 2030", "WEF/Brookings", "2026"),
      M("Cross-border road CAGR", "+3.95% jusqu'à 2031", "Mordor Intelligence", "2026-01"),
      M("Output Afrique 2025", "+4.2%", "African Trade Outlook", "2026-05", "vs +3.4% en 2024")
    ],

    /* ----- §7 Risques & opportunités actifs (juin 2026) ----- */
    risques_opportunites: [
      { type: "RISQUE", titre: "GUERRE IRAN (ACTIF)",
        detail: "Taux aérien +95% fév-mars 2026 (Drewry). Détournements maritime via Cap. Volatilité persistante Aérien et Maritime. Scénario pessimiste activé." },
      { type: "RISQUE", titre: "SURCAPACITÉ CONTENEURS",
        detail: "Taux fret -17% attendus 2026 (Drewry). Volumes TEUs en hausse mais revenus/TEU sous pression. Opportunité AGL si positionnement prix agressif." },
      { type: "RISQUE", titre: "COMMERCE MONDIAL RALENTI",
        detail: "OMC : +1.9% seulement en 2026 (vs +4.6% en 2025). Fin frontloading IA + plein effet tarifs. Impact modéré CIV." },
      { type: "OPPORTUNITE", titre: "AfCFTA ACCÉLÉRATION",
        detail: "$230B intra-africain 2026. +28% demande freight d'ici 2030. Fort positif Hinterland AGL (corridors Abidjan stratégiques)." },
      { type: "OPPORTUNITE", titre: "AÉRIEN AFRIQUE BOOM STRUCTUREL",
        detail: "+18% janvier 2026 (IATA). Région la plus dynamique mondiale. Croissance 8-10%/an. Positif Export Aérien CIV." },
      { type: "OPPORTUNITE", titre: "PND 2026-2030 CIV",
        detail: "8 400 Mds FCFA investissements infrastructure. Besoins logistiques massifs (BTP, énergie, mines). AGL sur corridors stratégiques." }
    ],

    /* ----- §6 Couche 2 : multiplicateurs spécifiques par métier ----- */
    multiplicateurs_metier: {
      import_export_maritime_teu: [
        M("OMC commerce mondial 2026E", "+1.9% → +2.6% 2027", "OMC", "2026-03"),
        M("Drewry taux conteneurs 2026", "-17% vs 2025", "Drewry", "2026-01"),
        M("Fleet Afrique nouvelles livr.", "+10-15%/an", "SeaRates", "2025-11")
      ],
      import_export_maritime_bulk: [
        M("Prix cacao — CIV 39% PDM mondiale", "Volatil", "Trésor.fr", "2025", "suivi mensuel"),
        M("Baltic Dry Index tendance", "Volatile", "BDI", "2026", "surcapacité fleet"),
        M("PND matériaux construction", "8 400 Mds FCFA", "PND CIV", "2026", "5 ans")
      ],
      import_aerien: [
        M("IATA Afrique cargo jan. 2026", "+18%", "IATA", "2026-04", "région la + dynamique"),
        M("IATA cargo mondial 2026E", "+2.4%", "IATA", "2026-01", "modération"),
        M("Impact Iran — taux aérien", "+95% fév-mars 2026", "Drewry", "2026-04")
      ],
      export_aerien: [
        M("IATA Afrique cargo 2025", "+6% demande, +7.8% capacité", "IATA", "2026-04"),
        M("Perishables CIV + trafic aérien", "+20% passagers Abidjan", "Rapport Éco. Suisse", "2025")
      ],
      hinterland_import: [
        M("AfCFTA commerce intra-africain 2026", "$230 milliards", "African Trade Outlook", "2026-05"),
        M("AfCFTA hausse fret d'ici 2030", "+28%", "Brookings", "2026-05"),
        M("Cross-border road Africa CAGR", "+3.95% jusqu'à 2031", "Mordor Intelligence", "2026-01")
      ],
      hinterland_export: [
        M("AfCFTA potentiel commerce", "+52.3% d'ici 2035", "Market Data Forecast", "2026-01"),
        M("Maritime intra-Afrique 2030", "58M → 132M tonnes", "WEF", "2023")
      ]
    },

    /* ----- §6 Couche 3 : 3 scénarios de projection par métier ----- */
    scenarios: [
      { code: "PESSIMISTE", percentile: "P10", probabilite: 0.15,
        hypotheses: "Guerre Iran prolongée + surcapacité maritime + instabilité hinterland + PND 50%" },
      { code: "BASE", percentile: "P50", probabilite: 0.65,
        hypotheses: "OMC +1.9% 2026 / +2.6% 2027 + AfCFTA progressif + PND 70% + stabilisation Iran" },
      { code: "OPTIMISTE", percentile: "P90", probabilite: 0.20,
        hypotheses: "AfCFTA accélération $230B + PND 100% + aérien Afrique +18% continue + boom cacao" }
    ]
  };

  global.MACRO_DATA = MACRO_DATA;
})(typeof window !== "undefined" ? window : this);
