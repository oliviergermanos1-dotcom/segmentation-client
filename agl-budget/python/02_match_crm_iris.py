#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — FLUX 1 : matching CRM × IRIS
=========================================
Spec §4 : produit la table ID_CRM ↔ ID_IRIS (taux attendu 85-92 %).

Entrée  : CRM_norm.xlsx, IRIS_norm.xlsx  (sortie de 01_normalise.py)
Sortie  : match_crm_iris.csv [id_crm, id_iris, nom_crm, nom_iris, score, classe]

Usage :
    python 02_match_crm_iris.py --crm CRM_norm.xlsx --iris IRIS_norm.xlsx \\
                                --out match_crm_iris.csv
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

import pandas as pd

import config
from matching_lib import match_two_tables


def _load(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path, dtype=str)
    return pd.read_csv(path, dtype=str, sep=None, engine="python")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AGL BUDGET — FLUX 1 matching CRM × IRIS")
    ap.add_argument("--crm",  required=True, help="CRM normalisé (sortie 01_normalise.py)")
    ap.add_argument("--iris", required=True, help="IRIS normalisé (sortie 01_normalise.py)")
    ap.add_argument("--out",  default="match_crm_iris.csv", help="Fichier de sortie")
    args = ap.parse_args(argv)

    crm  = _load(Path(args.crm))
    iris = _load(Path(args.iris))

    # Vérification des colonnes attendues — défauts dans config.py.
    id_crm  = config.resolve_column(crm.columns,  config.CRM["id"])
    id_iris = config.resolve_column(iris.columns, config.IRIS["id"])
    for need in ("NOM_BASE",):
        if need not in crm.columns or need not in iris.columns:
            sys.exit(f"ERREUR : colonne '{need}' absente — lance d'abord 01_normalise.py")
    if not id_crm:  sys.exit(f"ERREUR : colonne id CRM '{config.CRM['id']}' absente dans {args.crm}")
    if not id_iris: sys.exit(f"ERREUR : colonne id IRIS '{config.IRIS['id']}' absente dans {args.iris}")

    print(f"[in] CRM  : {len(crm):>6} lignes  | IRIS : {len(iris):>6} lignes")

    res = match_two_tables(
        left=crm, right=iris,
        left_id=id_crm, right_id=id_iris,
        flux_label="CRM_IRIS",
    )

    n_auto   = int((res["classe"] == "AUTO").sum())
    n_amb    = int((res["classe"] == "AMBIGU").sum())
    n_rej    = int((res["classe"] == "REJET").sum())
    print(f"[ok] {len(res):>6} matches — AUTO={n_auto}  AMBIGU={n_amb}  REJET={n_rej}"
          f"  | couverture CRM = {(n_auto+n_amb)/max(len(crm),1):.1%}")

    out = res.rename(columns={"id_L": "id_crm", "id_R": "id_iris",
                              "nom_L": "nom_crm", "nom_R": "nom_iris"})
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    out.to_csv(args.out, index=False, encoding="utf-8")
    print(f"[ok] écrit : {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
