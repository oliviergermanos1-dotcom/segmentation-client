#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — 06 : Agrégation STATCOM 3 ans
==========================================
Spec §2 + §3 (champ VOLUME_3ANS) + §6 (Couche 1 base de projection).

Produit, à partir du STATCOM enrichi (sortie de 03_match_crm_stat.py) :
  1. MARCHE_TOTAL  : marché total CIV par MÉTIER × ANNÉE (volumes + nb clients)
  2. VOLUME_CLIENT : volumes par CLIENT × MÉTIER × ANNÉE (pour PDM et RMC)
  3. METIERS_3ANS  : pivot par métier (2023/2024/2025) avec CAGR réel — §2

Volumes :
  - Métiers maritimes/hinterland : TEU principalement (+ Bulk si dispo)
  - Métiers aériens : Kg

Usage :
    python 06_agreg_statcom.py --statcom STATCOM_norm.xlsx --outdir agreg/
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

import pandas as pd

import config


def _load(p):
    p = Path(p)
    if p.suffix.lower() in (".xlsx", ".xls"): return pd.read_excel(p, dtype=str)
    return pd.read_csv(p, dtype=str, sep=None, engine="python")


def _to_num(s):
    """Conversion sûre str→float (gère ',', espaces fines, valeurs vides)."""
    return (pd.to_numeric(s.astype(str).str.replace(r"[   ]", "", regex=True)
                                          .str.replace(",", ".", regex=False),
                          errors="coerce")
            .fillna(0))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AGL BUDGET — 06 agrégation STATCOM")
    ap.add_argument("--statcom", required=True, help="STATCOM normalisé (sortie 03)")
    ap.add_argument("--outdir",  default="agreg", help="Dossier de sortie")
    args = ap.parse_args(argv)

    stat = _load(Path(args.statcom))
    col_metier = config.resolve_column(stat.columns, config.STATCOM["metier"])
    col_annee  = config.resolve_column(stat.columns, config.STATCOM["annee"])
    col_teu    = config.resolve_column(stat.columns, config.STATCOM["volume_teu"])
    col_bulk   = config.resolve_column(stat.columns, config.STATCOM["volume_bulk"])
    col_kg     = config.resolve_column(stat.columns, config.STATCOM["volume_kg"])

    if not col_metier: sys.exit("ERREUR : colonne 'metier' absente — voir config.STATCOM['metier'].")
    if not col_annee:  sys.exit("ERREUR : colonne 'annee' absente.")
    if "NOM_BASE" not in stat:
        sys.exit("ERREUR : STATCOM non normalisé — relance 03 avec --stat-norm-out")

    # Conversion numérique des volumes (ce qui n'existe pas → 0).
    stat["_TEU"]  = _to_num(stat[col_teu])  if col_teu  else 0.0
    stat["_BULK"] = _to_num(stat[col_bulk]) if col_bulk else 0.0
    stat["_KG"]   = _to_num(stat[col_kg])   if col_kg   else 0.0
    stat["_ANNEE"] = pd.to_numeric(stat[col_annee], errors="coerce").astype("Int64")
    stat = stat.dropna(subset=["_ANNEE"])

    outdir = Path(args.outdir); outdir.mkdir(parents=True, exist_ok=True)

    # 1) MARCHE_TOTAL : agrégat marché par métier × année.
    marche = (stat
              .groupby([col_metier, "_ANNEE"], dropna=False)
              .agg(VOLUME_TEU=("_TEU", "sum"),
                   VOLUME_BULK=("_BULK", "sum"),
                   VOLUME_KG=("_KG", "sum"),
                   NB_CLIENTS=("NOM_BASE", "nunique"),
                   NB_OPERATIONS=("NOM_BASE", "size"))
              .reset_index()
              .rename(columns={col_metier: "METIER", "_ANNEE": "ANNEE"}))
    marche.to_csv(outdir / "marche_total.csv", index=False, encoding="utf-8")

    # 2) VOLUME_CLIENT : par client × métier × année.
    vc = (stat
          .groupby(["NOM_BASE", col_metier, "_ANNEE"], dropna=False)
          .agg(VOLUME_TEU=("_TEU", "sum"),
               VOLUME_BULK=("_BULK", "sum"),
               VOLUME_KG=("_KG", "sum"),
               NB_OPERATIONS=("NOM_BASE", "size"),
               ID_STATCOM=("ID_STATCOM", "first") if "ID_STATCOM" in stat else ("NOM_BASE", "first"))
          .reset_index()
          .rename(columns={col_metier: "METIER", "_ANNEE": "ANNEE"}))
    vc.to_csv(outdir / "volume_client.csv", index=False, encoding="utf-8")

    # 3) METIERS_3ANS : pivot par métier 2023/2024/2025 + CAGR réel sur volume principal.
    # Vectorisé pour éviter le bug pandas "apply returns DataFrame" sur certains
    # cas (colonnes dupliquées via groupby+rename, etc.).
    _m_up = marche["METIER"].fillna("").astype(str).str.upper()
    _is_air = _m_up.str.contains("AERIEN") | _m_up.str.contains("AÉRIEN")
    marche["VOLUME_PRINCIPAL"] = marche["VOLUME_TEU"].where(~_is_air, marche["VOLUME_KG"])
    piv = marche.pivot_table(index="METIER", columns="ANNEE", values="VOLUME_PRINCIPAL",
                             aggfunc="sum", fill_value=0).reset_index()
    annees = sorted([c for c in piv.columns if isinstance(c, (int,)) or (hasattr(c, "item") and isinstance(c.item(), int))])
    if len(annees) >= 2:
        a_min, a_max = annees[0], annees[-1]
        years = (a_max - a_min) if (a_max - a_min) > 0 else 1
        def _cagr(r):
            v0, v1 = float(r[a_min]), float(r[a_max])
            if v0 <= 0 or v1 <= 0: return None
            return (v1 / v0) ** (1.0 / years) - 1.0
        piv["CAGR"] = piv.apply(_cagr, axis=1)
    piv.to_csv(outdir / "metiers_3ans.csv", index=False, encoding="utf-8")

    print(f"[ok] {len(marche)} lignes marché_total | "
          f"{len(vc)} lignes volume_client | "
          f"{len(piv)} métiers dans le pivot 3 ans")
    print(f"[ok] sortie : {outdir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
