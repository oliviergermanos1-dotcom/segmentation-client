#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Double-check des matchs par 2e moteur indépendant (EMM)
====================================================================
N'utilise PAS EMM pour remplacer le matching RapidFuzz (chaque moteur a ses
angles morts : word-cosine rate les mono-tokens, char-cosine gonfle les
sous-chaînes). On s'en sert comme JUGE INDÉPENDANT pour valider par consensus.

Pour chaque paire acceptée par le pipeline (RMC : CRM ↔ STATCOM / IRIS) :
  - EMM rejoue le matching de son côté (TF-IDF word + char cosine) ;
  - on compare la DÉCISION (pas le score) :

      RapidFuzz → X  ET  EMM → X            : CONFIRME_2MOTEURS  (🟢)
      RapidFuzz → X  ET  EMM → Y (≠)        : DIVERGENCE         (🟠)  worklist
      RapidFuzz → X  ET  EMM → rien proche  : SUSPECT            (🔴)  faux+ probable

Sortie : double_check.xlsx
  - clients              : verdict consensus par paire
  - divergences_suspects : 🟠 + 🔴 triés (worklist de revue, courte)
  - kpi                  : % confirmé par 2 moteurs

Prérequis : pip install emm
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

try:
    from emm import PandasEntityMatching
except ImportError:
    sys.exit("EMM requis : pip install emm")


def emm_best_picks(gt_names: pd.DataFrame, query_names: pd.DataFrame):
    """gt_names[id,name] = référentiel ; query_names[id,name] = à rapprocher.
    Retourne dict {query_name -> (gt_name, score_combine)}."""
    cfg = {"name_only": True, "entity_id_col": "id", "name_col": "name",
           "supervised_on": False, "aggregation_layer": False}
    m = PandasEntityMatching(cfg)
    m.fit(gt_names)
    res = m.transform(query_names).dropna(subset=["gt_entity_id"]).copy()
    # Score combiné = max(word cosine, char-2gram cosine) pour couvrir les 2 angles morts
    w = res.get("cossim_w1")
    c = res.get("cossim_n2")
    res["emm_combined"] = pd.concat(
        [w if w is not None else pd.Series(0, index=res.index),
         c if c is not None else pd.Series(0, index=res.index)], axis=1).max(axis=1)
    best = res.sort_values("emm_combined", ascending=False).drop_duplicates("name")
    return {r["name"]: (r["gt_name"], r["emm_combined"]) for _, r in best.iterrows()}


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--rmc", required=True, help="RMC.xlsx (paires acceptées)")
    ap.add_argument("--side", choices=["statcom", "iris"], default="statcom",
                    help="quelle jointure double-checker")
    ap.add_argument("--emm-floor", type=float, default=0.55,
                    help="sous ce score combiné EMM, la paire est SUSPECT")
    ap.add_argument("--out", default="double_check.xlsx")
    args = ap.parse_args(argv)

    rmc = pd.read_excel(args.rmc)
    alias_col = "ALIAS_STATCOM" if args.side == "statcom" else "ALIAS_IRIS"
    name_col = "NOM_CANONIQUE" if "NOM_CANONIQUE" in rmc.columns else "NOM_CLIENT"

    # Paires acceptées par le pipeline (CRM ↔ side)
    pairs = rmc[rmc[alias_col].notna()][["ID_CRM", name_col, alias_col]].copy()
    pairs.columns = ["id_crm", "crm_name", "side_name"]
    pairs = pairs[pairs["crm_name"].notna()]
    print(f"Paires {args.side} acceptées par le pipeline : {len(pairs):,}")

    # EMM : ground truth = noms CRM ; requêtes = noms du côté (STATCOM/IRIS)
    gt = pairs[["crm_name"]].drop_duplicates().reset_index(drop=True)
    gt.columns = ["name"]; gt["id"] = range(len(gt))
    q = pairs[["side_name"]].drop_duplicates().reset_index(drop=True)
    q.columns = ["name"]; q["id"] = range(len(q))

    print("[emm] 2e moteur en cours…")
    picks = emm_best_picks(gt, q)

    # Verdict de consensus
    def verdict(row):
        emm_pick, emm_score = picks.get(row["side_name"], (None, 0.0))
        if emm_pick is None or emm_score < args.emm_floor:
            return pd.Series([emm_pick, round(emm_score, 3), "SUSPECT"])
        if emm_pick == row["crm_name"]:
            return pd.Series([emm_pick, round(emm_score, 3), "CONFIRME_2MOTEURS"])
        return pd.Series([emm_pick, round(emm_score, 3), "DIVERGENCE"])

    pairs[["emm_pick", "emm_score", "VERDICT"]] = pairs.apply(verdict, axis=1)

    v = pairs["VERDICT"].value_counts()
    pct = (pairs["VERDICT"] == "CONFIRME_2MOTEURS").mean() * 100
    print("\n=== DOUBLE-CHECK ===")
    for k, n in v.items():
        print(f"  {k:18s} : {n:>6,} ({n/len(pairs)*100:4.1f}%)")
    print(f"\n  → % confirmé par 2 moteurs : {pct:.1f}%")
    print(f"  → à revoir (DIVERGENCE+SUSPECT) : "
          f"{(pairs['VERDICT']!='CONFIRME_2MOTEURS').sum():,}")

    suspects = pairs[pairs["VERDICT"] != "CONFIRME_2MOTEURS"].sort_values("emm_score")
    kpi = pd.DataFrame({"Verdict": list(v.index) + ["", "% CONFIRME_2MOTEURS"],
                        "Paires": list(v.values) + ["", f"{pct:.1f}%"]})
    with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
        kpi.to_excel(xl, sheet_name="kpi", index=False)
        pairs.to_excel(xl, sheet_name="clients", index=False)
        suspects.to_excel(xl, sheet_name="divergences_suspects", index=False)

    print(f"\n[ok] → {args.out}")
    if len(suspects):
        print("\nTop 15 SUSPECTS (faux-positifs probables) :")
        for _, r in suspects[suspects["VERDICT"] == "SUSPECT"].head(15).iterrows():
            print(f"  [{r['emm_score']:.2f}] {str(r['crm_name'])[:34]:34s} ↔ {r['side_name']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
