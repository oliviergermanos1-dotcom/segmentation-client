#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Bibliothèque de matching probabiliste (réutilisée par 02-04)
========================================================================
Implémente le pipeline :
  1. Normalisation (via 01_normalise.normalise) → NOM_NORMALISE, NOM_BASE
  2. Blocking : préfixe de NOM_BASE → réduit l'espace de recherche (sécurité mémoire §4)
  3. Scoring : par défaut Splink (Fellegi-Sunter, cf. §4) ; fallback RapidFuzz
     (token_set_ratio + Jaro-Winkler) si Splink absent.
  4. Classification : auto / zone ambigüe (Claude/manuel) / rejet selon SEUILS.

API principale :
    match_two_tables(left_df, right_df, left_name, right_name,
                     left_id, right_id, secondary_left=None, secondary_right=None)
        → DataFrame [id_left, id_right, score, classe]

`secondary_*` = colonne complémentaire (ex: marchandise pour STATCOM) qui
améliore le score de match — §4 "Splink fonctionne mieux avec 2+ colonnes".
"""
from __future__ import annotations

import sys
from typing import Iterable

try:
    import pandas as pd
except ImportError:
    sys.exit("pandas requis — pip install pandas openpyxl")

try:
    from rapidfuzz import fuzz, distance  # rapide, C-backed
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False

# Splink est l'outil recommandé (§4). Détection optionnelle.
try:
    import splink  # noqa: F401
    HAS_SPLINK = True
except ImportError:
    HAS_SPLINK = False

from config import SEUILS


# ---------------------------------------------------------------------------
# 1. Blocking : générer les paires candidates par préfixe
# ---------------------------------------------------------------------------
def _blocking_keys(name: str, prefix: int):
    """
    Génère plusieurs clés de blocking pour un même nom — évite les ratés
    quand le 1er token diffère ("SIR ..." vs "SOCIETE ..."). Une paire
    candidate est retenue si AU MOINS UNE clé est partagée.

    Clés produites : préfixe des `prefix` premiers caractères de chaque token
    de >= 3 lettres (max 4 tokens pour borner l'explosion).
    """
    if not name:
        return []
    toks = [t for t in name.split() if len(t) >= 3][:4]
    keys = {t[:prefix] for t in toks}
    return [k for k in keys if k]


def candidate_pairs(left, right, key="NOM_BASE", prefix=None, max_pairs=None):
    """Produit les paires (left × right) qui partagent au moins une clé de blocking."""
    prefix = prefix or SEUILS["block_prefix"]
    max_pairs = max_pairs or SEUILS["max_pairs"]

    def explode(df):
        out = df.copy()
        out["_blocks"] = out[key].fillna("").map(lambda s: _blocking_keys(s, prefix))
        out = out[out["_blocks"].map(len) > 0]
        return out.explode("_blocks").rename(columns={"_blocks": "_block"})

    L, R = explode(left), explode(right)
    pairs = L.merge(R, on="_block", suffixes=("_L", "_R"))
    # Dédoublonnage : un (id_L, id_R) peut apparaître par plusieurs clés.
    if "id_L" in pairs.columns and "id_R" in pairs.columns:
        pairs = pairs.drop_duplicates(subset=["id_L", "id_R"])

    if len(pairs) > max_pairs:
        sys.stderr.write(
            f"[warn] {len(pairs):,} paires candidates > seuil {max_pairs:,}. "
            f"Augmente block_prefix dans config.SEUILS pour réduire.\n"
        )
    return pairs.drop(columns=["_block"])


# ---------------------------------------------------------------------------
# 2. Scoring
# ---------------------------------------------------------------------------
def _score_rapidfuzz(name_l: str, name_r: str,
                     sec_l: str = "", sec_r: str = "") -> float:
    """
    Score 0..1 (RapidFuzz). Combine token_set_ratio (insensible à l'ordre)
    et Jaro-Winkler (typos / abréviations) — pondéré 70/30.
    Si une colonne secondaire (marchandise/secteur) est fournie, bonus jusqu'à 10 %.
    """
    if not name_l or not name_r:
        return 0.0
    tsr = fuzz.token_set_ratio(name_l, name_r) / 100.0
    jw = distance.JaroWinkler.normalized_similarity(name_l, name_r)
    base = 0.7 * tsr + 0.3 * jw
    if sec_l and sec_r:
        bonus = 0.10 * (fuzz.token_set_ratio(sec_l, sec_r) / 100.0)
        base = min(1.0, base + bonus)
    return base


def _classify(score: float) -> str:
    if score >= SEUILS["match_haut"]:
        return "AUTO"
    if score >= SEUILS["match_bas"]:
        return "AMBIGU"
    return "REJET"


# ---------------------------------------------------------------------------
# 3. API publique
# ---------------------------------------------------------------------------
def match_two_tables(left, right, left_id, right_id,
                     left_name="NOM_BASE", right_name="NOM_BASE",
                     secondary_left=None, secondary_right=None,
                     flux_label=""):
    """
    Renvoie un DataFrame des correspondances :
      [id_left, id_right, nom_left, nom_right, score, classe, flux]
    Stratégie : RapidFuzz (Splink à activer en production si volumes >>).
    """
    if not HAS_RAPIDFUZZ:
        sys.exit("[err] rapidfuzz requis — pip install rapidfuzz")

    if not HAS_SPLINK:
        sys.stderr.write(
            "[info] Splink absent — fallback RapidFuzz actif. "
            "Pour la production (recommandé §4) : pip install splink duckdb.\n"
        )

    # Sous-ensembles strictement nécessaires.
    L = left[[left_id, left_name] + ([secondary_left] if secondary_left else [])].copy()
    R = right[[right_id, right_name] + ([secondary_right] if secondary_right else [])].copy()
    L = L.rename(columns={left_id: "id_L", left_name: "nom_L",
                          **({secondary_left: "sec_L"} if secondary_left else {})})
    R = R.rename(columns={right_id: "id_R", right_name: "nom_R",
                          **({secondary_right: "sec_R"} if secondary_right else {})})

    pairs = candidate_pairs(
        L.rename(columns={"nom_L": "NOM_BASE"}),
        R.rename(columns={"nom_R": "NOM_BASE"}),
    )
    # Restore real column names after blocking merge.
    pairs = pairs.rename(columns={"NOM_BASE_L": "nom_L", "NOM_BASE_R": "nom_R"})

    if pairs.empty:
        return pd.DataFrame(columns=["id_L", "id_R", "nom_L", "nom_R", "score", "classe", "flux"])

    # Scoring vectorisé en boucle (RapidFuzz, suffisant <10M paires).
    has_sec = secondary_left and secondary_right
    scores = []
    nL = pairs["nom_L"].fillna("").to_numpy()
    nR = pairs["nom_R"].fillna("").to_numpy()
    sL = pairs["sec_L"].fillna("").to_numpy() if has_sec else [""] * len(pairs)
    sR = pairs["sec_R"].fillna("").to_numpy() if has_sec else [""] * len(pairs)
    for i in range(len(pairs)):
        scores.append(_score_rapidfuzz(nL[i], nR[i], sL[i], sR[i]))
    pairs["score"] = scores
    pairs["classe"] = [_classify(s) for s in scores]
    pairs["flux"] = flux_label

    # On garde la meilleure correspondance par id_L (un client gauche → 1 droit).
    pairs = pairs.sort_values(["id_L", "score"], ascending=[True, False])
    best = pairs.drop_duplicates(subset=["id_L"], keep="first")
    return best.reset_index(drop=True)
