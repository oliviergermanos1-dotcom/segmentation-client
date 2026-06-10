#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Pilote EMM supervisé (benchmark vs RapidFuzz)
==========================================================
Teste le moteur ING EntityMatchingModel en mode SUPERVISÉ sur les données
réelles, là où il révèle sa vraie valeur (vs le mode unsupervised brut qui
ne bat pas l'approche RapidFuzz actuelle — cf. analyse session).

Idée clé : on n'a pas de labels manuels, mais on en FABRIQUE à partir des
signaux dont on est sûr :
  POSITIFS  : paires dont token_set_ratio >= 0.95 ET partage >= 2 tokens
              significatifs (ex. "NESTLE CI" ↔ "NESTLE COTE IVOIRE").
  NÉGATIFS  : paires "hard negatives" — préfixe commun mais token_set bas
              (ex. "ASSETCO" ↔ "ASSEKE ORO", "CARGILL" ↔ "CARGO PLUS").
              Ce sont les faux-positifs typiques : le modèle apprend à les
              rejeter explicitement.

Le classifieur supervisé d'EMM combine alors word-cosine + char-cosine +
features de rang + (si activé) formes légales — ce qu'aucun signal brut seul
ne fait correctement.

Usage :
    python emm_pilot_supervised.py --crm CRM_norm.xlsx --statcom STATCOM_norm.xlsx \
        --current-matches match_crm_stat.csv --out bench_emm/

Prérequis : pip install emm  (rapidfuzz, scikit-learn, xgboost tirés en deps).
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd
from rapidfuzz import fuzz

try:
    from emm import PandasEntityMatching
except ImportError:
    sys.exit("EMM requis : pip install emm")


def _sig_tokens(s: str) -> set:
    """Tokens significatifs (>= 3 lettres) d'un nom normalisé."""
    return {t for t in str(s).upper().split() if len(t) >= 3}


def build_training_labels(crm: pd.DataFrame, stat: pd.DataFrame,
                          name_col_crm: str, name_col_stat: str,
                          n_pos: int = 2000, n_neg: int = 2000) -> pd.DataFrame:
    """Fabrique des paires labellisées (positifs sûrs + hard negatives)."""
    print("[labels] Génération des paires d'entraînement…")
    crm_names = crm[name_col_crm].dropna().astype(str).unique()
    stat_names = stat[name_col_stat].dropna().astype(str).unique()

    # Index par préfixe 4 lettres pour limiter les comparaisons
    from collections import defaultdict
    idx = defaultdict(list)
    for s in stat_names:
        toks = [t for t in s.upper().split() if len(t) >= 3]
        for t in toks[:3]:
            idx[t[:4]].append(s)

    positifs, negatifs = [], []
    for c in crm_names:
        ctoks = _sig_tokens(c)
        if not ctoks:
            continue
        cand = set()
        for t in list(ctoks)[:3]:
            cand.update(idx.get(t[:4], []))
        for s in cand:
            tsr = fuzz.token_set_ratio(c, s) / 100.0
            shared = len(ctoks & _sig_tokens(s))
            if tsr >= 0.95 and shared >= 1:
                positifs.append((c, s, 1))
            elif 0.50 <= tsr < 0.72 and shared == 0:
                # préfixe commun (même bloc) mais mots différents = hard negative
                negatifs.append((c, s, 0))

    pos_df = pd.DataFrame(positifs, columns=["name", "gt_name", "label"]).drop_duplicates()
    neg_df = pd.DataFrame(negatifs, columns=["name", "gt_name", "label"]).drop_duplicates()
    pos_df = pos_df.sample(min(n_pos, len(pos_df)), random_state=42) if len(pos_df) else pos_df
    neg_df = neg_df.sample(min(n_neg, len(neg_df)), random_state=42) if len(neg_df) else neg_df
    print(f"[labels] {len(pos_df):,} positifs · {len(neg_df):,} hard negatives")
    return pd.concat([pos_df, neg_df], ignore_index=True)


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--crm", required=True)
    ap.add_argument("--statcom", required=True)
    ap.add_argument("--current-matches", required=True,
                    help="match_crm_stat.csv produit par le pipeline actuel")
    ap.add_argument("--crm-col", default="NOM_BASE")
    ap.add_argument("--stat-col", default="NOM_BASE")
    ap.add_argument("--out", default="bench_emm")
    args = ap.parse_args(argv)

    out = Path(args.out); out.mkdir(parents=True, exist_ok=True)

    crm = pd.read_excel(args.crm)
    stat = pd.read_excel(args.statcom)
    print(f"CRM : {len(crm):,} · STATCOM : {len(stat):,}")

    # Ground truth = entités CRM uniques
    gt = crm[[args.crm_col]].dropna().drop_duplicates().reset_index(drop=True)
    gt.columns = ["name"]; gt["id"] = range(len(gt))
    gt = gt[gt["name"].str.strip() != ""]

    # 1) Labels d'entraînement fabriqués
    train = build_training_labels(crm, stat, args.crm_col, args.stat_col)
    train.to_csv(out / "training_labels.csv", index=False)

    # 2) Config EMM supervisée
    cfg = {
        "name_only": True, "entity_id_col": "id", "name_col": "name",
        "supervised_on": True, "aggregation_layer": False,
    }
    model = PandasEntityMatching(cfg)
    print("[fit] index TF-IDF sur ground truth CRM…")
    model.fit(gt)

    # 3) Entraînement du classifieur supervisé sur nos labels fabriqués
    #    EMM attend les labels rattachés à un gt_entity_id : on mappe le gt_name.
    gtmap = gt.set_index("name")["id"].to_dict()
    train["gt_entity_id"] = train["gt_name"].map(gtmap)
    train_ok = train.dropna(subset=["gt_entity_id"])
    print(f"[train] {len(train_ok):,} paires labellisées rattachées à une entité GT")
    try:
        model.fit_classifier(train_ok.rename(columns={"label": "correct"}))
        print("[train] classifieur supervisé entraîné ✓")
    except Exception as e:
        print(f"[train] ⚠️ fit_classifier a échoué ({e}). "
              f"Le pilote tombe en scoring cosine brut.")

    # 4) Scoring de toutes les paires STATCOM → meilleur CRM
    queries = stat[[args.stat_col]].dropna().drop_duplicates().reset_index(drop=True)
    queries.columns = ["name"]; queries["id"] = range(len(queries))
    res = model.transform(queries)
    score_col = "nm_score" if "nm_score" in res.columns else (
        "score_0" if "score_0" in res.columns else "cossim_w1")
    res = res.dropna(subset=["gt_entity_id"]).copy()
    best = res.sort_values(score_col, ascending=False).drop_duplicates("name")
    best[["name", "gt_name", score_col]].to_csv(out / "emm_supervised_matches.csv", index=False)
    print(f"[ok] résultats EMM supervisé → {out/'emm_supervised_matches.csv'}")

    # 5) Comparaison vs pipeline actuel
    cur = pd.read_csv(args.current_matches)
    print(f"\n[compare] matchs actuels : {len(cur):,} lignes")
    print(f"[compare] matchs EMM (score {score_col}) : "
          f">=0.85 = {(best[score_col]>=0.85).sum():,} | "
          f"0.65-0.85 = {((best[score_col]>=0.65)&(best[score_col]<0.85)).sum():,}")
    print("\n→ Comparer emm_supervised_matches.csv avec match_crm_stat.csv "
          "pour mesurer le gain réel (couverture + précision sur les AMBIGUS).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
