#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — 08 : Projections 2026-2030 hybrides
================================================
Spec §6 — architecture 3 couches :
  Couche 1 — STATCOM 3 ans : CAGR réel par métier (base primaire — §6).
  Couche 2 — Multiplicateurs macro par métier (sources §7, juin 2026) :
             OMC, Drewry, IATA, AfCFTA, IATA Afrique, prix cacao…
  Couche 3 — 3 scénarios par métier : P10 PESSIMISTE / P50 BASE / P90 OPTIMISTE.

Méthode :
  v(année) = v(année-1) × (1 + CAGR_ajusté)
  CAGR_ajusté = w_reel · CAGR_reel + w_macro · CAGR_macro_metier_scenario

Pondération par défaut : 50/50 — réajustable via --w-real/--w-macro.

Entrée : agreg/metiers_3ans.csv (sortie de 06)
Sortie : agreg/projections.csv [METIER, SCENARIO, ANNEE, VOLUME, CAGR_APPLIQUE]
"""
from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path

import pandas as pd


# ----------------------------------------------------------------------------
# COUCHE 2 — Multiplicateurs macro par métier (CAGR volume annualisé attendu)
# Dérivés des sources §7 du cahier des charges (juin 2026). Numériques pour le
# moteur. Les sources textuelles restent dans js/macro_data.js (onglet Macro).
# ----------------------------------------------------------------------------
MACRO_MULTIPLIERS = {
    # Maritime — OMC +1.9% mondial + Afrique +10-15% + AfCFTA progressif
    "Import Maritime":   {"PESSIMISTE": -0.03, "BASE": 0.055, "OPTIMISTE": 0.110,
                          "sources": "OMC 2026E +1.9% / Drewry vol +1.8% / SeaRates Afrique +10-15%"},
    "Export Maritime":   {"PESSIMISTE": -0.02, "BASE": 0.060, "OPTIMISTE": 0.120,
                          "sources": "OMC export 2026E +2.6% 2027 / cacao volatile / AfCFTA"},
    # Aérien — IATA Afrique +18% jan-2026, mondial +2.4% modéré
    "Import Aérien":     {"PESSIMISTE":  0.00, "BASE": 0.080, "OPTIMISTE": 0.150,
                          "sources": "IATA Afrique cargo +18% jan-2026 / monde +2.4% / Iran +95% taux"},
    "Export Aérien":     {"PESSIMISTE":  0.00, "BASE": 0.060, "OPTIMISTE": 0.110,
                          "sources": "IATA Afrique +6% demande / perishables CIV / Abidjan +20% pax"},
    # Hinterland — AfCFTA +28% d'ici 2030, cross-border +3.95%/an
    "Hinterland Import": {"PESSIMISTE": -0.02, "BASE": 0.064, "OPTIMISTE": 0.130,
                          "sources": "AfCFTA +28% fret 2030 / Mordor +3.95% CAGR 2031 / $230B intra-Af"},
    "Hinterland Export": {"PESSIMISTE": -0.02, "BASE": 0.072, "OPTIMISTE": 0.140,
                          "sources": "AfCFTA potentiel +52.3% 2035 / WEF maritime intra-Af 58→132M t"},
}

SCENARIOS = ["PESSIMISTE", "BASE", "OPTIMISTE"]
PROBABILITES = {"PESSIMISTE": 0.15, "BASE": 0.65, "OPTIMISTE": 0.20}  # §6 Couche 3


def _scenario_factor(real_cagr, macro_cagr, w_real, w_macro):
    """Pondère le CAGR réel observé et le CAGR macro projeté."""
    if real_cagr is None:
        return macro_cagr
    return w_real * real_cagr + w_macro * macro_cagr


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AGL BUDGET — 08 projections 2026-2030")
    ap.add_argument("--metiers", required=True, help="agreg/metiers_3ans.csv (sortie 06)")
    ap.add_argument("--outdir",  default="agreg")
    ap.add_argument("--horizon", type=int, default=2030, help="année cible (défaut 2030)")
    ap.add_argument("--w-real",  type=float, default=0.5, help="poids CAGR réel STATCOM (défaut 0.5)")
    ap.add_argument("--w-macro", type=float, default=0.5, help="poids CAGR macro (défaut 0.5)")
    args = ap.parse_args(argv)

    if abs((args.w_real + args.w_macro) - 1.0) > 0.001:
        sys.exit("ERREUR : --w-real + --w-macro doit valoir 1.0")

    piv = pd.read_csv(args.metiers)
    if "METIER" not in piv.columns:
        sys.exit("ERREUR : metiers_3ans.csv mal formé (pas de colonne METIER)")
    annees_cols = [c for c in piv.columns if str(c).isdigit()]
    if not annees_cols:
        sys.exit("ERREUR : aucune colonne d'année dans metiers_3ans.csv")
    last_year = max(int(c) for c in annees_cols)
    years_proj = list(range(last_year + 1, args.horizon + 1))

    rows = []
    for _, r in piv.iterrows():
        metier = r["METIER"]
        v_last = float(r.get(str(last_year), 0) or 0)
        cagr_reel = float(r["CAGR"]) if "CAGR" in piv.columns and pd.notna(r.get("CAGR")) else None
        macro = MACRO_MULTIPLIERS.get(metier, {})

        for sc in SCENARIOS:
            macro_cagr = macro.get(sc) if metier in MACRO_MULTIPLIERS else 0.02
            applied = _scenario_factor(cagr_reel, macro_cagr, args.w_real, args.w_macro)
            v = v_last
            for y in years_proj:
                v = v * (1 + applied)
                rows.append({
                    "METIER": metier, "SCENARIO": sc, "ANNEE": y,
                    "VOLUME": round(v, 2),
                    "CAGR_APPLIQUE": round(applied, 4),
                    "PROBABILITE": PROBABILITES[sc],
                    "SOURCES_MACRO": macro.get("sources", "")
                })
        # On ajoute aussi la dernière année réelle (référence visuelle).
        rows.append({"METIER": metier, "SCENARIO": "REEL", "ANNEE": last_year,
                     "VOLUME": round(v_last, 2),
                     "CAGR_APPLIQUE": cagr_reel if cagr_reel is not None else "",
                     "PROBABILITE": 1.0, "SOURCES_MACRO": "STATCOM observé"})

    out = pd.DataFrame(rows).sort_values(["METIER", "SCENARIO", "ANNEE"])
    outdir = Path(args.outdir); outdir.mkdir(parents=True, exist_ok=True)
    out.to_csv(outdir / "projections.csv", index=False, encoding="utf-8")
    # On exporte aussi la table des multiplicateurs pour traçabilité.
    with open(outdir / "macro_multipliers.json", "w", encoding="utf-8") as f:
        json.dump({"horizon": args.horizon, "w_real": args.w_real, "w_macro": args.w_macro,
                   "multiplicateurs": MACRO_MULTIPLIERS, "probabilites": PROBABILITES}, f, indent=2, ensure_ascii=False)

    print(f"[ok] {len(out)} lignes projetées | métiers={out['METIER'].nunique()} | "
          f"horizon {last_year+1}-{args.horizon} | scénarios={SCENARIOS}")
    print(f"[ok] écrit : {outdir/'projections.csv'} & macro_multipliers.json")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
