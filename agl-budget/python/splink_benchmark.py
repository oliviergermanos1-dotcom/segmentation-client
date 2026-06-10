#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Benchmark Splink (croisement probabiliste multi-colonnes)
=====================================================================
Compare le moteur Splink (Fellegi-Sunter, DuckDB) au matching RapidFuzz
mono-colonne actuel, sur le croisement CRM <-> STATCOM (ou CRM <-> IRIS).

Splink lie sur PLUSIEURS colonnes à la fois (nom + secteur) et apprend
le poids de chaque accord/désaccord par EM (non supervisé). On mesure :
  - couverture (combien de liens trouvés)
  - accord avec le pipeline actuel
  - liens NOUVEAUX (que le mono-colonne ratait)
  - sur-matching potentiel (à filtrer par TOKEN_FLOOR + double-check EMM)

⚠️ Splink a tendance à sur-matcher (précision 60-76% en éval publique) :
ce benchmark sert à mesurer le gain de RAPPEL, à encadrer ensuite par les
garde-fous précision (token floor + EMM 05g).

Usage :
    python splink_benchmark.py --left CRM_norm.xlsx --right STATCOM_norm.xlsx \
        --name-col NOM_BASE --sec-left secteur --sec-right secteur \
        --current match_crm_stat.csv --out bench_splink/
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

try:
    from splink import Linker, DuckDBAPI, SettingsCreator, block_on
    import splink.comparison_library as cl
except ImportError:
    sys.exit("Splink requis : pip install splink duckdb")


def _prep(df, name_col, sec_col, prefix):
    """Normalise les colonnes attendues par Splink : unique_id, name, secteur, bloc."""
    out = pd.DataFrame()
    out["unique_id"] = [f"{prefix}{i}" for i in range(len(df))]
    out["name"] = df[name_col].fillna("").astype(str).str.upper().str.strip()
    out["secteur"] = (df[sec_col].fillna("").astype(str).str.upper().str.strip()
                      if sec_col and sec_col in df.columns else "")
    # clé de blocking : 4 premières lettres du 1er token significatif
    def blk(s):
        toks = [t for t in s.split() if len(t) >= 3]
        return toks[0][:4] if toks else (s[:4] if s else "ZZZZ")
    out["bloc"] = out["name"].map(blk)
    out = out[out["name"] != ""]
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--left", required=True, help="table gauche (ex CRM_norm.xlsx)")
    ap.add_argument("--right", required=True, help="table droite (ex STATCOM_norm.xlsx)")
    ap.add_argument("--name-col", default="NOM_BASE")
    ap.add_argument("--sec-left", default="secteur")
    ap.add_argument("--sec-right", default="secteur")
    ap.add_argument("--current", help="match CSV du pipeline actuel (comparaison)")
    ap.add_argument("--threshold", type=float, default=0.9,
                    help="seuil de probabilité de match Splink")
    ap.add_argument("--out", default="bench_splink")
    args = ap.parse_args(argv)

    out = Path(args.out); out.mkdir(parents=True, exist_ok=True)

    L = _prep(pd.read_excel(args.left), args.name_col, args.sec_left, "L")
    R = _prep(pd.read_excel(args.right), args.name_col, args.sec_right, "R")
    print(f"Gauche : {len(L):,} · Droite : {len(R):,}")

    # ===== Modèle Splink : lien probabiliste multi-colonnes =====
    settings = SettingsCreator(
        link_type="link_only",
        comparisons=[
            cl.JaroWinklerAtThresholds("name", [0.9, 0.7]),
            cl.ExactMatch("secteur").configure(term_frequency_adjustments=True),
        ],
        blocking_rules_to_generate_predictions=[block_on("bloc")],
    )
    db = DuckDBAPI()
    linker = Linker([L, R], settings, db_api=db,
                    input_table_aliases=["tbl_l", "tbl_r"])

    # Estimation non supervisée des paramètres (u via random sampling, m via EM)
    print("[splink] estimation des paramètres (EM, non supervisé)…")
    linker.training.estimate_probability_two_random_records_match(
        [block_on("name")], recall=0.7)
    linker.training.estimate_u_using_random_sampling(max_pairs=1e6)
    try:
        linker.training.estimate_parameters_using_expectation_maximisation(block_on("bloc"))
    except Exception as e:
        print(f"[splink] EM partiel ({e}) — on continue avec u-only.")

    print("[splink] prédiction des liens…")
    preds = linker.inference.predict(threshold_match_probability=args.threshold)
    pdf = preds.as_pandas_dataframe()
    print(f"[splink] {len(pdf):,} liens >= proba {args.threshold}")

    # Meilleur lien par enregistrement gauche
    lname = L.set_index("unique_id")["name"].to_dict()
    rname = R.set_index("unique_id")["name"].to_dict()
    # Splink nomme les colonnes unique_id_l / unique_id_r
    idl = "unique_id_l" if "unique_id_l" in pdf.columns else "unique_id_left"
    idr = "unique_id_r" if "unique_id_r" in pdf.columns else "unique_id_right"
    pdf["name_l"] = pdf[idl].map(lname).fillna(pdf[idl].map(rname))
    pdf["name_r"] = pdf[idr].map(rname).fillna(pdf[idr].map(lname))
    best = pdf.sort_values("match_probability", ascending=False).drop_duplicates(idl)
    best[[idl, "name_l", idr, "name_r", "match_probability"]].to_csv(
        out / "splink_matches.csv", index=False)

    print(f"\n=== RÉSULTAT SPLINK ===")
    print(f"  Liens (best/gauche) : {len(best):,}")
    print(f"  proba >= 0.99 : {(best['match_probability']>=0.99).sum():,}")
    print(f"  0.90-0.99     : {((best['match_probability']>=0.90)&(best['match_probability']<0.99)).sum():,}")

    # ===== Comparaison avec le pipeline actuel =====
    if args.current and Path(args.current).exists():
        cur = pd.read_csv(args.current)
        print(f"\n=== vs PIPELINE ACTUEL ===")
        print(f"  Matchs actuels (RapidFuzz) : {len(cur):,}")
        print(f"  Liens Splink               : {len(best):,}")
        print(f"  → Comparer splink_matches.csv avec {Path(args.current).name} "
              f"pour : liens nouveaux (rappel +) ET sur-matchs (à filtrer).")

    print(f"\n[ok] → {out/'splink_matches.csv'}")
    print("Note : encadrer ces liens par TOKEN_FLOOR + double-check EMM (05g) "
          "avant production — Splink sur-matche par nature.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
