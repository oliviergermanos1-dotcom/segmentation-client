#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — 09 : White Spaces (marché non capturé)
====================================================
Spec §4 + onglet 07 — clients STATCOM absents du CRM (= prospects non couverts
par AGL) ou rattachés en zone ambiguë.

Entrée :
  --rmc            RMC.xlsx (sortie 05)
  --volume-client  agreg/volume_client.csv (sortie 06)
  --match-crm-stat fixtures/.../match_crm_stat.csv (sortie 03, pour récupérer score)

Sortie : agreg/whitespaces.csv
  [NOM_BASE, METIER, ANNEE, VOLUME_PRINCIPAL, NB_OPERATIONS, STATUT]
  STATUT ∈ {ABSENT_CRM, MATCH_AMBIGU}
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

import pandas as pd


def _load(p):
    p = Path(p)
    if p.suffix.lower() in (".xlsx", ".xls"): return pd.read_excel(p, dtype=str)
    return pd.read_csv(p, dtype=str, sep=None, engine="python")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AGL BUDGET — 09 white spaces")
    ap.add_argument("--rmc",            required=True)
    ap.add_argument("--volume-client",  required=True)
    ap.add_argument("--match-crm-stat", required=True)
    ap.add_argument("--outdir",         default="agreg")
    args = ap.parse_args(argv)

    rmc = _load(args.rmc)
    vc  = _load(args.volume_client)
    mcs = _load(args.match_crm_stat)

    for col in ("VOLUME_TEU", "VOLUME_BULK", "VOLUME_KG", "NB_OPERATIONS"):
        if col in vc: vc[col] = pd.to_numeric(vc[col], errors="coerce").fillna(0)
    vc["ANNEE"] = pd.to_numeric(vc["ANNEE"], errors="coerce").astype("Int64")

    # Volume principal selon métier (TEU sauf aérien → Kg).
    metier_up = vc["METIER"].fillna("").str.upper()
    vc["VOLUME_PRINCIPAL"] = vc["VOLUME_TEU"].where(
        ~(metier_up.str.contains("AERIEN") | metier_up.str.contains("AÉRIEN")), vc["VOLUME_KG"])

    # Clients STATCOM rattachés en AUTO (score haut).
    rattaches_ok = set()
    if "id_statcom" in mcs and "classe" in mcs:
        rattaches_ok = set(mcs.loc[mcs["classe"] == "AUTO", "id_statcom"].dropna())

    # Tout client STATCOM dont l'ID_STATCOM n'est PAS en AUTO = white space candidat.
    if "ID_STATCOM" in vc.columns:
        vc["EST_AUTO"] = vc["ID_STATCOM"].isin(rattaches_ok)
        candidats = vc[~vc["EST_AUTO"]].copy()
    else:
        # Fallback : on n'a pas l'ID_STATCOM → on prend tout et on signale.
        sys.stderr.write("[warn] volume_client.csv sans ID_STATCOM — fallback : tous les clients sont candidats.\n")
        vc["EST_AUTO"] = False
        candidats = vc.copy()

    # Statut détaillé : ABSENT_CRM si jamais matché, MATCH_AMBIGU sinon.
    ambigus = set()
    if "id_statcom" in mcs and "classe" in mcs:
        ambigus = set(mcs.loc[mcs["classe"] == "AMBIGU", "id_statcom"].dropna())
    def statut(id_stat):
        if id_stat in ambigus: return "MATCH_AMBIGU"
        return "ABSENT_CRM"
    if "ID_STATCOM" in candidats.columns:
        candidats["STATUT"] = candidats["ID_STATCOM"].map(statut)
    else:
        candidats["STATUT"] = "ABSENT_CRM"

    # Agrégat par client × métier (toutes années).
    grp_keys = ["NOM_BASE", "METIER", "STATUT"]
    agg = (candidats.groupby(grp_keys, dropna=False)
           .agg(VOLUME_PRINCIPAL=("VOLUME_PRINCIPAL", "sum"),
                NB_OPERATIONS=("NB_OPERATIONS", "sum"),
                ANNEES_PRESENT=("ANNEE", "nunique"))
           .reset_index()
           .sort_values("VOLUME_PRINCIPAL", ascending=False))

    outdir = Path(args.outdir); outdir.mkdir(parents=True, exist_ok=True)
    agg.to_csv(outdir / "whitespaces.csv", index=False, encoding="utf-8")

    n = len(agg)
    n_abs = int((agg["STATUT"] == "ABSENT_CRM").sum())
    n_amb = int((agg["STATUT"] == "MATCH_AMBIGU").sum())
    top = agg.head(5)[["NOM_BASE", "METIER", "VOLUME_PRINCIPAL"]]
    print(f"[ok] {n} white spaces — ABSENT_CRM={n_abs} | MATCH_AMBIGU={n_amb}")
    if len(top):
        print("[top 5 par volume]")
        for _, r in top.iterrows():
            print(f"    {r['NOM_BASE']:<28} {r['METIER']:<22} {r['VOLUME_PRINCIPAL']:>10,.0f}")
    print(f"[ok] écrit : {outdir/'whitespaces.csv'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
