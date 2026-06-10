#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Consolidation par GRAPHE (maillage complet 4 bases + clustering)
=============================================================================
Remplace le matching en étoile (tout autour du CRM) par un croisement
"tout-à-tous" : les 4 bases sont versées dans un seul pool, Splink lie
chaque enregistrement à TOUS les autres (multi-colonnes, non supervisé),
puis on regroupe en grappes (connected components). Chaque grappe = une
entité réelle unique.

Avantages :
  - un client présent dans IRIS+STATCOM mais ABSENT du CRM est unifié ;
  - transitivité : IRIS↔STATCOM + STATCOM↔RUBRIKS → 3 fusionnent ;
  - plus de membres par entité = plus de signaux pour certifier le secteur.

Garde-fou anti sur-fusion :
  - seuil de clustering élevé (défaut 0.92) ;
  - alerte sur les grappes anormalement grosses (> --max-cluster membres)
    qui trahissent un faux lien transitif (à éclater manuellement).

Entrées : les *_norm.xlsx (CRM/IRIS/STATCOM) + RUBRIKS. Chaque table doit
exposer une colonne nom normalisé (NOM_BASE) et, si dispo, un secteur.

Sortie : entites_graphe.xlsx
  - entites        : 1 ligne / entité (ID_UNIQUE, présences, noms, secteur voté)
  - membres        : tous les enregistrements avec leur cluster_id
  - grappes_alerte : clusters trop gros (sur-fusion probable)
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


def _prep(path, source, name_col, sec_col):
    """Normalise une base vers le schéma commun du pool."""
    if not path:
        return None
    df = pd.read_excel(path)
    out = pd.DataFrame()
    n = len(df)
    out["unique_id"] = [f"{source}_{i}" for i in range(n)]
    out["source"] = source
    nm = name_col if name_col in df.columns else None
    out["name"] = (df[nm].fillna("").astype(str).str.upper().str.strip()
                   if nm else "")
    out["secteur"] = (df[sec_col].fillna("").astype(str).str.upper().str.strip()
                      if sec_col and sec_col in df.columns else "")
    def blk(s):
        toks = [t for t in s.split() if len(t) >= 3]
        return toks[0][:4] if toks else (s[:4] if s else "ZZZZ")
    out["bloc"] = out["name"].map(blk)
    out = out[out["name"] != ""]
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--crm"); ap.add_argument("--iris")
    ap.add_argument("--statcom"); ap.add_argument("--rubriks")
    ap.add_argument("--name-col", default="NOM_BASE")
    ap.add_argument("--sec-col", default="secteur")
    ap.add_argument("--threshold", type=float, default=0.92,
                    help="seuil de match ET de clustering (strict = anti sur-fusion)")
    ap.add_argument("--max-cluster", type=int, default=8,
                    help="taille de grappe au-delà de laquelle on alerte (sur-fusion)")
    ap.add_argument("--out", default="entites_graphe.xlsx")
    args = ap.parse_args(argv)

    tables = []
    for src, path in [("CRM", args.crm), ("IRIS", args.iris),
                      ("STATCOM", args.statcom), ("RUBRIKS", args.rubriks)]:
        t = _prep(path, src, args.name_col, args.sec_col)
        if t is not None and len(t):
            tables.append(t)
            print(f"  {src:8s} : {len(t):,} enregistrements")
    if len(tables) < 2:
        sys.exit("Au moins 2 bases requises.")

    # ===== Splink : link_and_dedupe sur le pool complet =====
    settings = SettingsCreator(
        link_type="link_and_dedupe",
        comparisons=[
            cl.JaroWinklerAtThresholds("name", [0.92, 0.85]),
            cl.ExactMatch("secteur").configure(term_frequency_adjustments=True),
        ],
        blocking_rules_to_generate_predictions=[block_on("bloc")],
        retain_intermediate_calculation_columns=True,
    )
    db = DuckDBAPI()
    aliases = [f"src_{i}" for i in range(len(tables))]
    linker = Linker(tables, settings, db_api=db, input_table_aliases=aliases)

    print("[splink] estimation des paramètres (EM)…")
    linker.training.estimate_probability_two_random_records_match(
        [block_on("name")], recall=0.7)
    linker.training.estimate_u_using_random_sampling(max_pairs=2e6)
    try:
        linker.training.estimate_parameters_using_expectation_maximisation(block_on("bloc"))
    except Exception as e:
        print(f"[splink] EM partiel ({e})")

    print("[splink] prédiction + clustering…")
    preds = linker.inference.predict(threshold_match_probability=args.threshold)
    clusters = linker.clustering.cluster_pairwise_predictions_at_threshold(
        preds, threshold_match_probability=args.threshold)
    cdf = clusters.as_pandas_dataframe()
    print(f"[splink] {cdf['cluster_id'].nunique():,} entités pour "
          f"{len(cdf):,} enregistrements")

    # ===== Agrégation : 1 ligne / entité =====
    pool = pd.concat(tables, ignore_index=True).set_index("unique_id")
    cdf = cdf.set_index("unique_id").join(pool[["source", "name", "secteur"]],
                                          rsuffix="_p")
    rows = []
    for cid, g in cdf.groupby("cluster_id"):
        srcs = set(g["source"])
        secteurs = [s for s in g["secteur"] if s]
        sec_vote = (pd.Series(secteurs).value_counts().idxmax() if secteurs else "")
        rows.append({
            "ID_UNIQUE": cid,
            "NOM_PRINCIPAL": g["name"].mode().iloc[0] if len(g) else "",
            "N_MEMBRES": len(g),
            "PRESENCE_CRM": "✓" if "CRM" in srcs else "—",
            "PRESENCE_IRIS": "✓" if "IRIS" in srcs else "—",
            "PRESENCE_STATCOM": "✓" if "STATCOM" in srcs else "—",
            "PRESENCE_RUBRIKS": "✓" if "RUBRIKS" in srcs else "—",
            "N_BASES": len(srcs),
            "SECTEUR_VOTE": sec_vote,
            "NOMS_VARIANTES": " | ".join(sorted(set(g["name"]))[:5]),
        })
    ent = pd.DataFrame(rows).sort_values("N_MEMBRES", ascending=False)

    alertes = ent[ent["N_MEMBRES"] > args.max_cluster]
    print(f"\n=== RÉSULTAT GRAPHE ===")
    print(f"  Entités uniques            : {len(ent):,}")
    print(f"  Entités multi-bases (>=2)  : {(ent['N_BASES']>=2).sum():,}")
    print(f"  Entités sur les 4 bases    : {(ent['N_BASES']==4).sum():,}")
    print(f"  ⚠️ Grappes à vérifier (>{args.max_cluster} membres) : {len(alertes):,}")

    with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
        ent.to_excel(xl, sheet_name="entites", index=False)
        cdf.reset_index().to_excel(xl, sheet_name="membres", index=False)
        alertes.to_excel(xl, sheet_name="grappes_alerte", index=False)
    print(f"\n[ok] → {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
