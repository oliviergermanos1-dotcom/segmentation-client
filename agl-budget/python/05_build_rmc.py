#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Construction du Référentiel Maître Client (RMC)
============================================================
Spec §3 : 15 champs. Règle d'or : SECTEUR CRM = LOCKED (jamais écrasé).
Propagation des secteurs selon la hiérarchie §3 (CRM → RUBRIKS → autres).

Entrée :
  --crm CRM_norm.xlsx
  --m12 match_crm_iris.csv  (flux 1)
  --m13 match_crm_stat.csv  (flux 2)
  --m23 match_iris_stat.csv (flux 3 pont, optionnel)
  --rubriks RUBRIKS.xlsx    (optionnel — enrichissement secteur)

Sortie : RMC.xlsx + RMC.csv (15 colonnes selon §3)
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

import pandas as pd

import config


SCHEMA = [
    "ID_RMC", "ID_CRM", "ID_IRIS", "ID_STATCOM",
    "NOM_CANONIQUE", "SECTEUR", "ALIAS_CRM", "ALIAS_IRIS", "ALIAS_STATCOM",
    "MARCHANDISES", "VOLUME_3ANS", "CAP_3ANS",
    "SCORE_MATCH_IRIS", "SCORE_MATCH_STATCOM",
    "SOURCE_SECTEUR", "LOCKED",
]


def _load(p):
    p = Path(p)
    if p.suffix.lower() in (".xlsx", ".xls"): return pd.read_excel(p, dtype=str)
    return pd.read_csv(p, dtype=str, sep=None, engine="python")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AGL BUDGET — construction RMC")
    ap.add_argument("--crm",      required=True)
    ap.add_argument("--m12",      required=True, help="match CRM × IRIS")
    ap.add_argument("--m13",      required=True, help="match CRM × STATCOM")
    ap.add_argument("--m23",      default=None,  help="match IRIS × STATCOM (pont, optionnel)")
    ap.add_argument("--rubriks",  default=None,  help="RUBRIKS (enrichissement secteur)")
    ap.add_argument("--out",      default="RMC.xlsx")
    args = ap.parse_args(argv)

    crm = _load(args.crm)
    id_crm  = config.resolve_column(crm.columns, config.CRM["id"])
    nom_crm = config.resolve_column(crm.columns, config.CRM["nom"])
    sec_crm = config.resolve_column(crm.columns, config.CRM["secteur"])
    if not id_crm or not nom_crm:
        sys.exit("ERREUR : colonnes id/nom CRM manquantes (cf. config.CRM).")

    # 1. Squelette : 1 ligne par client CRM (PRIORITÉ 1).
    rmc = pd.DataFrame({
        "ID_RMC":   ["RMC_" + str(i+1).zfill(6) for i in range(len(crm))],
        "ID_CRM":   crm[id_crm],
        "NOM_CANONIQUE":  crm[nom_crm],
        "SECTEUR":  crm[sec_crm] if sec_crm else "",
        "ALIAS_CRM": crm["NOM_NORMALISE"] if "NOM_NORMALISE" in crm else crm[nom_crm],
    })
    # SECTEUR : LOCKED si fourni par CRM (§3 priorité 1).
    rmc["LOCKED"] = rmc["SECTEUR"].fillna("").str.len().gt(0)
    rmc["SOURCE_SECTEUR"] = rmc["LOCKED"].map({True: "CRM", False: ""})

    # 2. FLUX 1 : ID_IRIS et alias IRIS.
    m12 = _load(args.m12)
    m12_ok = m12[m12["classe"].isin(["AUTO", "AMBIGU"])]
    rmc = rmc.merge(
        m12_ok[["id_crm", "id_iris", "nom_iris", "score"]].rename(
            columns={"id_crm": "ID_CRM", "id_iris": "ID_IRIS",
                     "nom_iris": "ALIAS_IRIS", "score": "SCORE_MATCH_IRIS"}),
        on="ID_CRM", how="left",
    )

    # 3. FLUX 2 : ID_STATCOM et alias STATCOM.
    m13 = _load(args.m13)
    m13_ok = m13[m13["classe"].isin(["AUTO", "AMBIGU"])]
    rmc = rmc.merge(
        m13_ok[["id_crm", "id_statcom", "nom_statcom", "score"]].rename(
            columns={"id_crm": "ID_CRM", "id_statcom": "ID_STATCOM",
                     "nom_statcom": "ALIAS_STATCOM", "score": "SCORE_MATCH_STATCOM"}),
        on="ID_CRM", how="left",
    )

    # 4. FLUX 3 (pont) : récupère ID_STATCOM via IRIS pour les clients où le flux 2 a raté.
    if args.m23:
        m23 = _load(args.m23)
        m23_ok = m23[m23["classe"].isin(["AUTO", "AMBIGU"])]
        m23_map = m23_ok[["id_iris", "id_statcom"]].rename(
            columns={"id_iris": "ID_IRIS", "id_statcom": "ID_STATCOM_PONT"})
        rmc = rmc.merge(m23_map, on="ID_IRIS", how="left")
        # On remplit ID_STATCOM seulement s'il est vide.
        mask = rmc["ID_STATCOM"].isna() & rmc["ID_STATCOM_PONT"].notna()
        rmc.loc[mask, "ID_STATCOM"] = rmc.loc[mask, "ID_STATCOM_PONT"]
        rmc = rmc.drop(columns=["ID_STATCOM_PONT"])

    # 5. RUBRIKS (enrichissement secteur, priorité 2 §3 — uniquement si CRM absent).
    if args.rubriks:
        # Import paresseux de la normalisation pour appliquer les aliases manuels.
        import importlib.util as _iu
        _spec = _iu.spec_from_file_location("norm", Path(__file__).parent / "01_normalise.py")
        norm = _iu.module_from_spec(_spec); _spec.loader.exec_module(norm)

        ru = _load(args.rubriks)
        ru_nom = config.resolve_column(ru.columns, config.RUBRIKS["nom"])
        ru_sec = config.resolve_column(ru.columns, config.RUBRIKS["secteur"])
        if ru_nom and ru_sec:
            # Aliases RUBRIKS → puis upper/strip pour la clé.
            ru_aliased = ru[ru_nom].fillna("").map(lambda n: norm.apply_alias(n, "RUBRIKS"))
            key = rmc["NOM_CANONIQUE"].fillna("").str.upper().str.strip()
            ru_key = ru_aliased.str.upper().str.strip()
            ru_lookup = dict(zip(ru_key, ru[ru_sec].fillna("")))
            for i, k in key.items():
                if rmc.at[i, "LOCKED"]:
                    continue
                if k in ru_lookup and ru_lookup[k]:
                    rmc.at[i, "SECTEUR"] = ru_lookup[k]
                    rmc.at[i, "SOURCE_SECTEUR"] = "RUBRIKS"

    # 6. Colonnes du schéma absentes : à remplir aux sprints suivants (volumes, CAP).
    for c in SCHEMA:
        if c not in rmc.columns:
            rmc[c] = ""

    rmc = rmc[SCHEMA]
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    if out.suffix.lower() == ".csv":
        rmc.to_csv(out, index=False)
    else:
        rmc.to_excel(out, index=False)
        rmc.to_csv(out.with_suffix(".csv"), index=False)

    n = len(rmc)
    n_iris = int(rmc["ID_IRIS"].fillna("").astype(str).str.len().gt(0).sum())
    n_stat = int(rmc["ID_STATCOM"].fillna("").astype(str).str.len().gt(0).sum())
    n_sec  = int(rmc["SECTEUR"].fillna("").str.len().gt(0).sum())
    n_lock = int(rmc["LOCKED"].sum())
    print(f"[ok] RMC : {n} clients | IRIS rattachés {n_iris} ({n_iris/max(n,1):.0%})"
          f" | STATCOM rattachés {n_stat} ({n_stat/max(n,1):.0%})"
          f" | secteur renseigné {n_sec} ({n_sec/max(n,1):.0%}) dont LOCKED {n_lock}")
    print(f"[ok] écrit : {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
