#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — FLUX 3 (pont) : matching IRIS × STATCOM
====================================================
Spec §4 : filet de sécurité — récupère les clients STATCOM non rattachés au CRM
via le pont IRIS quand IRIS sait les nommer. Permet de propager ID_IRIS sur des
clients STATCOM que le flux 2 (CRM × STATCOM) aurait raté.

Entrée : IRIS_norm.xlsx, STATCOM_norm.xlsx (sortie de 03)
Sortie : match_iris_stat.csv

Usage :
    python 04_match_iris_stat.py --iris IRIS_norm.xlsx --statcom STATCOM_norm.xlsx \\
                                 --out match_iris_stat.csv
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
    ap = argparse.ArgumentParser(description="AGL BUDGET — FLUX 3 (pont) IRIS × STATCOM")
    ap.add_argument("--iris",    required=True)
    ap.add_argument("--statcom", required=True)
    ap.add_argument("--out",     default="match_iris_stat.csv")
    args = ap.parse_args(argv)

    iris = _load(Path(args.iris))
    stat = _load(Path(args.statcom))

    id_iris = config.resolve_column(iris.columns, config.IRIS["id"])
    if not id_iris:           sys.exit(f"ERREUR : colonne id IRIS '{config.IRIS['id']}' absente")
    if "NOM_BASE" not in iris: sys.exit("ERREUR : IRIS non normalisé")
    if "NOM_BASE" not in stat: sys.exit("ERREUR : STATCOM non normalisé (lance 03 avec --stat-norm-out)")
    if "ID_STATCOM" not in stat:
        sys.exit("ERREUR : ID_STATCOM absent — STATCOM doit venir de 03_match_crm_stat.py --stat-norm-out")

    # Côté STATCOM, on dédoublonne par client (les volumes seront agrégés en §6).
    stat_clients = (stat[["ID_STATCOM", "NOM_BASE", "NOM_NORMALISE"]]
                    .dropna(subset=["NOM_BASE"]).drop_duplicates(subset=["ID_STATCOM"]))

    print(f"[in] IRIS : {len(iris):>6}  | STATCOM clients : {len(stat_clients):>6}")

    res = match_two_tables(
        left=iris, right=stat_clients,
        left_id=id_iris, right_id="ID_STATCOM",
        flux_label="IRIS_STATCOM",
    )

    n_auto = int((res["classe"] == "AUTO").sum())
    n_amb  = int((res["classe"] == "AMBIGU").sum())
    print(f"[ok] {len(res):>6} matches — AUTO={n_auto}  AMBIGU={n_amb}"
          f"  | couverture IRIS = {(n_auto+n_amb)/max(len(iris),1):.1%}")

    out = res.rename(columns={"id_L": "id_iris", "id_R": "id_statcom",
                              "nom_L": "nom_iris", "nom_R": "nom_statcom"})
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    out.to_csv(args.out, index=False)
    print(f"[ok] écrit : {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
