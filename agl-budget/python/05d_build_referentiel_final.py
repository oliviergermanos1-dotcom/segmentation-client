#!/usr/bin/env python3
"""05d — Référentiel client unifié final (livrable business).

Produit clients_referentiel.xlsx avec :
  - normalisation nomenclature secteurs (mapping marchandise → libellés CRM)
  - score de match global = min(SCORE_IRIS, SCORE_STATCOM) pour les CRM, 0 pour TM
  - 3 onglets segmentés par qualité de match :
        1_match_fort   : score > 80%
        2_match_moyen  : 65 ≤ score ≤ 80%
        3_a_revoir     : score < 65% (white spaces TM_xxx essentiellement)
  - 4e onglet kpi_synthese : compteurs + couverture sectorielle
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

# Normalisation : mapping marchandise → nomenclature CRM officielle AGL
SECTEUR_NORMALISATION = {
    "FMCG / RETAIL": "FAST MOVING CONSUMER GOODS (FMCG)",
    "INFRASTRUCTURE": "INFRASTRUCTURES",
    "AGRI BUSINESS": "AGRICULTURE",
    "OIL & GAZ": "OIL & GAS",
    "HIGHT TECH / TELECOM": "HIGH TECH",
    "ENERGIE": "POWER TECHNOLOGIES",
    "AUTRES": "OTHERS",
    "PRODUITS CHIMIQUES": "PRODUITS CHIMIQUES",  # secteur autonome confirmé
}


def normaliser_secteur(s):
    if pd.isna(s):
        return s
    s = str(s).strip()
    return SECTEUR_NORMALISATION.get(s, s)


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--master-v2", required=True, help="clients_master_v2.xlsx (sortie 05c)")
    ap.add_argument("--rmc", required=True, help="RMC.xlsx (pour récupérer ALIAS et scores fins)")
    ap.add_argument("--out", required=True, help="clients_referentiel.xlsx")
    args = ap.parse_args(argv)

    print("[1/4] Chargement master v2 + RMC…")
    master = pd.read_excel(args.master_v2, sheet_name="clients_master")
    rmc = pd.read_excel(args.rmc)
    print(f"     Master : {len(master):,} lignes")
    print(f"     RMC    : {len(rmc):,} lignes (CRM matchés)")

    # === Enrichir master avec ALIAS + scores fins du RMC =================
    print("[2/4] Enrichissement (alias par base, scores fins)…")
    rmc_keys = rmc[["ID_CRM", "ALIAS_CRM", "ALIAS_IRIS",
                    "SCORE_MATCH_IRIS", "SCORE_MATCH_STATCOM"]].copy()
    rmc_keys = rmc_keys.dropna(subset=["ID_CRM"]).drop_duplicates(subset=["ID_CRM"])
    master = master.merge(rmc_keys, on="ID_CRM", how="left")

    # === Normaliser SECTEUR =============================================
    print("[3/4] Normalisation secteurs (mapping marchandise → CRM)…")
    avant = master["SECTEUR"].value_counts()
    master["SECTEUR"] = master["SECTEUR"].apply(normaliser_secteur)
    apres = master["SECTEUR"].value_counts()
    n_renames = sum(1 for s in avant.index if normaliser_secteur(s) != s)
    print(f"     {n_renames} secteurs renommés sur {len(avant)} uniques")

    # === Score global combiné ===========================================
    # CRM : min(IRIS, STATCOM) pour refléter le maillon faible
    # TM_xxx : score = 0 (pas de match)
    def score_global(row):
        if not str(row["ID_UNIQUE"]).startswith("TM"):
            si = row.get("SCORE_MATCH_IRIS")
            ss = row.get("SCORE_MATCH_STATCOM")
            vals = [v for v in (si, ss) if pd.notna(v)]
            if vals:
                return min(vals)  # le maillon faible domine
            return 1.0  # CRM-only, pas de jointure IRIS/STATCOM → on garde 1.0
        return 0.0
    master["SCORE_GLOBAL"] = master.apply(score_global, axis=1)

    # === Construction des colonnes du livrable ==========================
    print("[4/4] Construction du référentiel + onglets…")
    cols_final = [
        "ID_UNIQUE",
        "ID_CRM",
        "NOM_CLIENT",
        "ID_IRIS",
        "ID_STATCOM",
        "ALIAS_CRM",
        "ALIAS_IRIS",
        "ALIAS_STATCOM",
        "SECTEUR",
        "SECTEUR_SOURCE",
        "SCORE_GLOBAL",
        "SCORE_MATCH_IRIS",
        "SCORE_MATCH_STATCOM",
        "STATUT_MATCHING",
        "PRESENCE_CRM",
        "PRESENCE_IRIS",
        "PRESENCE_STATCOM",
        "PRESENCE_RUBRIKS",
        "DOMINANCE_SECTEUR",
        "NB_BL_STATCOM",
        "TOP2_SECTEURS",
    ]
    cols_final = [c for c in cols_final if c in master.columns]
    ref = master[cols_final].copy()

    # === Buckets par qualité de match ===================================
    onglet_fort  = ref[ref["SCORE_GLOBAL"] > 0.80].sort_values("NOM_CLIENT")
    onglet_moyen = ref[(ref["SCORE_GLOBAL"] >= 0.65) & (ref["SCORE_GLOBAL"] <= 0.80)].sort_values("SCORE_GLOBAL", ascending=False)
    onglet_revoir = ref[ref["SCORE_GLOBAL"] < 0.65].sort_values("NOM_CLIENT")

    # KPI synthèse
    kpi = pd.DataFrame({
        "Indicateur": [
            "Total clients référentiel",
            "Onglet 1 : match fort (>80%)",
            "Onglet 2 : match moyen (65-80%)",
            "Onglet 3 : à revoir (<65% + white spaces)",
            "",
            "Couverture sectorielle",
            "  dont CRM (★ triple match)",
            "  dont INFERE (dominance ≥60%)",
            "  dont INFERE_MOYEN (40-60%)",
            "  dont MIXTE (<40%)",
            "  dont INFERE_DEFAUT (fallback métier)",
            "",
            "Top 10 secteurs",
        ] + [f"  {s}" for s in ref["SECTEUR"].value_counts().head(10).index],
        "Valeur": [
            len(ref),
            len(onglet_fort),
            len(onglet_moyen),
            len(onglet_revoir),
            "",
            f"{ref['SECTEUR'].notna().mean()*100:.1f}%",
            (ref["SECTEUR_SOURCE"] == "CRM").sum(),
            (ref["SECTEUR_SOURCE"] == "INFERE").sum(),
            (ref["SECTEUR_SOURCE"] == "INFERE_MOYEN").sum(),
            (ref["SECTEUR_SOURCE"] == "MIXTE").sum(),
            (ref["SECTEUR_SOURCE"] == "INFERE_DEFAUT").sum(),
            "",
            "",
        ] + [v for v in ref["SECTEUR"].value_counts().head(10).values],
    })

    with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
        kpi.to_excel(xl, sheet_name="kpi_synthese", index=False)
        onglet_fort.to_excel(xl,   sheet_name="1_match_fort_80plus", index=False)
        onglet_moyen.to_excel(xl,  sheet_name="2_match_moyen_65_80", index=False)
        onglet_revoir.to_excel(xl, sheet_name="3_a_revoir_moins65",  index=False)

    print(f"[ok] Référentiel final écrit : {args.out}")
    print()
    print(f"[summary] Répartition par onglet :")
    print(f"  1_match_fort_80plus   : {len(onglet_fort):>6,}")
    print(f"  2_match_moyen_65_80   : {len(onglet_moyen):>6,}")
    print(f"  3_a_revoir_moins65    : {len(onglet_revoir):>6,}")
    print(f"  → total               : {len(ref):>6,}")
    print()
    print(f"[summary] Top 10 secteurs (référentiel complet) :")
    print(ref["SECTEUR"].value_counts().head(10).to_string())

    return 0


if __name__ == "__main__":
    sys.exit(main())
