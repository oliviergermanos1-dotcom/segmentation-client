#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — 11 : Budget PFA (RUBRIKS) vs Réel (IRIS)
=====================================================
Onglet 04 (§1). Compare la CAP budgétée par client/secteur (RUBRIKS PFA 2026)
au CAP réel observé dans IRIS, en s'appuyant sur le RMC pour la jointure.

Méthode :
  1. Normalise les noms RUBRIKS (Cleanco + extension OHADA — partagé avec 01).
  2. Rattache RUBRIKS au RMC via NOM_BASE (best-effort, fallback nom canonique).
  3. Agrège IRIS par CRM × année → cap_reel.
  4. Joint le tout : on obtient pour chaque client RMC l'année 2026
     {cap_pfa, cap_reel, ecart_abs, ecart_pct, secteur}.
  5. Produit deux sorties :
       budget_vs_real.csv   : 1 ligne par client × année
       budget_secteur.csv   : agrégat par secteur × année

Entrée :
  --rmc       RMC.xlsx (sortie 05)
  --iris      IRIS_norm.xlsx (sortie 01) — DOIT contenir cap + periode (ou année)
  --rubriks   RUBRIKS.xlsx — DOIT contenir cap_pfa + secteur + annee
"""
from __future__ import annotations
import argparse
import importlib.util
import re
import sys
from pathlib import Path

import pandas as pd

import config

# Charge 01_normalise pour réutiliser la normalisation OHADA/CIV.
_spec = importlib.util.spec_from_file_location("norm", Path(__file__).parent / "01_normalise.py")
norm = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(norm)


def _load(p):
    p = Path(p)
    if p.suffix.lower() in (".xlsx", ".xls"): return pd.read_excel(p, dtype=str)
    return pd.read_csv(p, dtype=str, sep=None, engine="python")


def _num(s):
    return (pd.to_numeric(s.astype(str).str.replace(r"[   ]", "", regex=True)
                                       .str.replace(",", ".", regex=False),
                          errors="coerce")
            .fillna(0))


def _year(v):
    if v is None or v == "": return None
    s = str(v)
    m = re.search(r"(20\d{2})", s)
    return int(m.group(1)) if m else None


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AGL BUDGET — 11 Budget PFA vs Réel")
    ap.add_argument("--rmc",     required=True)
    ap.add_argument("--iris",    required=True, help="IRIS normalisé (sortie 01)")
    ap.add_argument("--rubriks", required=True)
    ap.add_argument("--outdir",  default="agreg")
    args = ap.parse_args(argv)

    rmc  = _load(args.rmc)
    iris = _load(args.iris)
    ru   = _load(args.rubriks)

    # --- Côté IRIS : cap réel par CRM × année --------------------------------
    iris_id  = config.resolve_column(iris.columns, config.IRIS["id"])
    iris_cap = config.resolve_column(iris.columns, config.IRIS["cap"], ["ca", "ca_fcfa", "revenue"])
    iris_per = config.resolve_column(iris.columns, config.IRIS["periode"], ["annee", "year"])
    if not iris_id:  sys.exit("ERREUR : colonne id IRIS absente (config.IRIS['id']).")
    if not iris_cap: sys.exit("ERREUR : colonne cap IRIS absente — fournis 'cap' ou ajuste config.IRIS['cap'].")

    iris = iris.copy()
    iris["_cap"] = _num(iris[iris_cap])
    iris["_annee"] = iris[iris_per].map(_year) if iris_per else 2026

    iris_agg = (iris.groupby([iris_id, "_annee"], dropna=False)["_cap"].sum()
                .reset_index().rename(columns={iris_id: "ID_IRIS", "_annee": "ANNEE", "_cap": "CAP_REEL"}))

    # --- Côté RUBRIKS : cap PFA par client × année + normalisation pour jointure
    ru_nom = config.resolve_column(ru.columns, config.RUBRIKS["nom"])
    ru_sec = config.resolve_column(ru.columns, config.RUBRIKS["secteur"])
    ru_cap = config.resolve_column(ru.columns, config.RUBRIKS["cap_pfa"], ["cap", "budget", "objectif"])
    ru_ann = config.resolve_column(ru.columns, config.RUBRIKS["annee"], ["periode", "year"])
    if not ru_nom or not ru_cap:
        sys.exit("ERREUR : colonnes nom/cap_pfa RUBRIKS manquantes.")

    res = ru[ru_nom].fillna("").map(norm.normalise)
    ru["NOM_BASE"]      = [r[1] for r in res]
    ru["_cap_pfa"]      = _num(ru[ru_cap])
    ru["_secteur"]      = ru[ru_sec] if ru_sec else ""
    ru["_annee"]        = ru[ru_ann].map(_year) if ru_ann else 2026

    # --- Rattachement RUBRIKS → RMC via NOM_BASE (ALIAS_STATCOM ou normalisation
    # du nom canonique). Best-effort, fallback nom canonique.
    rmc = rmc.copy()
    cn_res = rmc["NOM_CANONIQUE"].fillna("").map(norm.normalise)
    rmc["_NOM_BASE_CANON"] = [r[1] for r in cn_res]

    join_key_rmc = rmc[["ID_CRM", "ID_IRIS", "NOM_CANONIQUE", "SECTEUR", "LOCKED",
                        "ALIAS_STATCOM", "_NOM_BASE_CANON"]].copy()
    # On crée une table 'clés' empilant ALIAS_STATCOM ET le NOM_BASE canonique.
    keys = pd.concat([
        join_key_rmc.rename(columns={"ALIAS_STATCOM": "NOM_BASE"})[["ID_CRM", "ID_IRIS", "NOM_BASE",
                                                                    "NOM_CANONIQUE", "SECTEUR", "LOCKED"]],
        join_key_rmc.rename(columns={"_NOM_BASE_CANON": "NOM_BASE"})[["ID_CRM", "ID_IRIS", "NOM_BASE",
                                                                      "NOM_CANONIQUE", "SECTEUR", "LOCKED"]],
    ], ignore_index=True).dropna(subset=["NOM_BASE"]).query("NOM_BASE != ''")
    keys = keys.drop_duplicates(subset=["NOM_BASE", "ID_CRM"])

    bud = ru.merge(keys, on="NOM_BASE", how="left").copy()
    nb_unmatched = bud["ID_CRM"].isna().sum()

    # --- Joint avec IRIS (cap réel) ------------------------------------------
    bud = bud.merge(iris_agg, left_on=["ID_IRIS", "_annee"],
                    right_on=["ID_IRIS", "ANNEE"], how="left")
    bud["CAP_REEL"] = bud["CAP_REEL"].fillna(0)

    # On préfère le secteur LOCKED du RMC si dispo, sinon le secteur RUBRIKS.
    bud["SECTEUR_FINAL"] = bud["SECTEUR"].where(bud["SECTEUR"].fillna("").str.len() > 0, bud["_secteur"])

    bud["ECART"]     = bud["CAP_REEL"] - bud["_cap_pfa"]
    bud["ECART_PCT"] = bud.apply(
        lambda r: (r["ECART"] / r["_cap_pfa"]) if r["_cap_pfa"] > 0 else 0.0, axis=1)

    out = bud[["ID_CRM", "ID_IRIS", "NOM_CANONIQUE", "SECTEUR_FINAL", "LOCKED",
               "_annee", "_cap_pfa", "CAP_REEL", "ECART", "ECART_PCT"]].rename(
        columns={"_annee": "ANNEE", "_cap_pfa": "CAP_PFA",
                 "NOM_CANONIQUE": "NOM_CLIENT", "SECTEUR_FINAL": "SECTEUR"})

    # On fait remonter les RUBRIKS non rattachés (NOM_CANONIQUE vide) avec le nom RUBRIKS d'origine.
    fallback_name = ru[ru_nom].astype(str)
    out["NOM_CLIENT"] = out["NOM_CLIENT"].fillna("")
    out.loc[out["NOM_CLIENT"] == "", "NOM_CLIENT"] = fallback_name.loc[out["NOM_CLIENT"] == ""]

    outdir = Path(args.outdir); outdir.mkdir(parents=True, exist_ok=True)
    out.to_csv(outdir / "budget_vs_real.csv", index=False)

    # Agrégat par secteur × année.
    sec = (out.groupby(["SECTEUR", "ANNEE"], dropna=False)
           .agg(CAP_PFA=("CAP_PFA", "sum"),
                CAP_REEL=("CAP_REEL", "sum"),
                NB_CLIENTS=("NOM_CLIENT", "nunique"))
           .reset_index())
    sec["ECART"] = sec["CAP_REEL"] - sec["CAP_PFA"]
    sec["ECART_PCT"] = sec.apply(lambda r: r["ECART"] / r["CAP_PFA"] if r["CAP_PFA"] > 0 else 0.0, axis=1)
    sec.to_csv(outdir / "budget_secteur.csv", index=False)

    print(f"[in]  RUBRIKS : {len(ru):>5} lignes | IRIS : {len(iris):>5} lignes | RMC : {len(rmc):>5} clients")
    print(f"[ok] {len(out):>5} lignes RUBRIKS rapprochées "
          f"({len(out) - nb_unmatched} rattachées au RMC, {nb_unmatched} non rattachées)")
    print(f"[ok] {len(sec):>5} lignes par secteur × année")
    print(f"[ok] écrit : {outdir/'budget_vs_real.csv'} & {outdir/'budget_secteur.csv'}")

    # Récap secteur si dispo.
    if len(sec):
        latest = sec["ANNEE"].max()
        snap = sec[sec["ANNEE"] == latest].sort_values("CAP_PFA", ascending=False).head(10)
        print(f"\n[récap {latest}]")
        for _, r in snap.iterrows():
            print(f"    {(r['SECTEUR'] or '—'):<28} "
                  f"PFA={r['CAP_PFA']:>14,.0f} | RÉEL={r['CAP_REEL']:>14,.0f} | "
                  f"écart={r['ECART_PCT']*100:>+6.1f}%")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
