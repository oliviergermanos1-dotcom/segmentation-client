#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — FLUX 2 : matching CRM × STATCOM
============================================
Spec §4 : 75-85 %, Splink + colonne marchandise pour booster le score
("Splink fonctionne mieux avec 2+ colonnes — nom + marchandise = parfait").

⚠ Particularité STATCOM (confirmée utilisateur) :
  - Métiers d'IMPORT (Maritime, Hinterland, Aérien) → client = "destinataires"
  - Métiers d'EXPORT (Maritime, Hinterland, Aérien) → client = "chargeurs"

Le script lit un STATCOM brut (3 ans concaténés possible), reconstruit une
colonne unique NOM_CLIENT_STATCOM selon le sens du métier, normalise, puis
matche contre CRM_norm.

Entrée : CRM_norm.xlsx, STATCOM.xlsx (avec colonne métier + destinataires + chargeurs)
Sortie : match_crm_stat.csv + STATCOM_norm.xlsx (avec NOM_NORMALISE/NOM_BASE/CLIENT_RESOLU)

Usage :
    python 03_match_crm_stat.py --crm CRM_norm.xlsx --statcom STATCOM.xlsx \\
                                --out match_crm_stat.csv
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

import pandas as pd

import config
from matching_lib import match_two_tables
# 01 est non-importable directement (commence par un chiffre) — chargement via importlib.
import importlib.util
_spec = importlib.util.spec_from_file_location("norm", Path(__file__).parent / "01_normalise.py")
norm = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(norm)


def _load(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path, dtype=str)
    return pd.read_csv(path, dtype=str, sep=None, engine="python")


def resolve_statcom_client(df: pd.DataFrame) -> pd.DataFrame:
    """
    Reconstruit une colonne unique CLIENT_RESOLU à partir de la règle métier :
      Import* → destinataires | Export* → chargeurs.

    Mode 1 (priorité) : si la colonne `client` existe déjà (sortie de
    _statcom_consolide.py), on la réutilise telle quelle.
    Mode 2 (fallback) : on applique la règle destinataires/chargeurs sur le
    fichier source brut.
    """
    df = df.copy()

    # Mode 1 : colonne client déjà résolue par la consolidation amont.
    c_client = config.resolve_column(df.columns, "client", ["CLIENT_RESOLU"])
    col_metier = config.resolve_column(df.columns, config.STATCOM["metier"], ["metier_statcom"])
    if col_metier is None:
        sys.exit("ERREUR : colonne métier introuvable dans STATCOM (config.STATCOM['metier']).")

    if c_client is not None:
        df["CLIENT_RESOLU"] = df[c_client]
        df["CLIENT_COL_SOURCE"] = "client (consolidé)"
        return df

    # Mode 2 : fallback historique sur destinataires/chargeurs.
    col_dest  = config.resolve_column(df.columns, "destinataires", ["destinataire"])
    col_charg = config.resolve_column(df.columns, "chargeurs",     ["chargeur", "expediteurs", "expediteur"])
    if col_dest is None and col_charg is None:
        sys.exit("ERREUR : aucune colonne client (`client`, `destinataires` ou `chargeurs`) trouvée.")

    metier_to_col = {m: meta["client_col"] for m, meta in config.STATCOM_METIERS.items()}

    def pick(row):
        m = (row[col_metier] or "").strip()
        col = metier_to_col.get(m)
        if not col:
            up = m.upper()
            if "EXPORT" in up:  col = "chargeurs"
            elif "IMPORT" in up or "HINTERLAND" in up: col = "destinataires"
            else: col = "destinataires"
        real = col_charg if col == "chargeurs" else col_dest
        return row[real] if real else None

    df["CLIENT_RESOLU"] = df.apply(pick, axis=1)
    df["CLIENT_COL_SOURCE"] = df[col_metier].map(metier_to_col).fillna("auto")
    return df


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AGL BUDGET — FLUX 2 matching CRM × STATCOM")
    ap.add_argument("--crm",     required=True)
    ap.add_argument("--statcom", required=True)
    ap.add_argument("--out",     default="match_crm_stat.csv")
    ap.add_argument("--stat-norm-out", default=None,
                    help="Optionnel : chemin où écrire le STATCOM normalisé enrichi")
    args = ap.parse_args(argv)

    crm = _load(Path(args.crm))
    id_crm = config.resolve_column(crm.columns, config.CRM["id"])
    if not id_crm:           sys.exit(f"ERREUR : colonne id CRM '{config.CRM['id']}' absente")
    if "NOM_BASE" not in crm: sys.exit("ERREUR : CRM non normalisé — lance 01_normalise.py d'abord")

    stat = _load(Path(args.statcom))
    stat = resolve_statcom_client(stat)

    # Normalisation côté STATCOM (sur CLIENT_RESOLU).
    res = stat["CLIENT_RESOLU"].fillna("").map(norm.normalise)
    stat["NOM_NORMALISE"] = [r[0] for r in res]
    stat["NOM_BASE"]      = [r[1] for r in res]

    # Agrégation : 1 nom client STATCOM unique (la dédoublonnage volumes vient en §6).
    col_march = config.resolve_column(stat.columns, config.STATCOM["marchandise"])
    keep = ["NOM_BASE", "NOM_NORMALISE", "CLIENT_RESOLU"]
    if col_march: keep.append(col_march)
    stat_clients = stat[stat["NOM_BASE"].str.len() > 0][keep].copy()
    # Agrège la marchandise en chaîne triée (utile pour score secondaire).
    if col_march:
        stat_clients = (stat_clients
                        .groupby("NOM_BASE", as_index=False)
                        .agg({"NOM_NORMALISE": "first",
                              "CLIENT_RESOLU": "first",
                              col_march: lambda s: " ".join(sorted({str(x) for x in s if x}))}))
        sec_right = col_march
    else:
        stat_clients = stat_clients.drop_duplicates(subset=["NOM_BASE"])
        sec_right = None

    # Ajoute un id synthétique STATCOM (pas d'ID natif côté STATCOM, cf. §2).
    stat_clients = stat_clients.reset_index(drop=True)
    stat_clients["ID_STATCOM"] = ["STAT_" + str(i+1).zfill(6) for i in range(len(stat_clients))]

    print(f"[in] CRM  : {len(crm):>6}  | STATCOM clients uniques : {len(stat_clients):>6}")

    res = match_two_tables(
        left=crm, right=stat_clients,
        left_id=id_crm, right_id="ID_STATCOM",
        secondary_left=None,                    # CRM n'a pas de marchandise
        secondary_right=sec_right,              # mais STATCOM oui → bonus
        flux_label="CRM_STATCOM",
    )

    n_auto = int((res["classe"] == "AUTO").sum())
    n_amb  = int((res["classe"] == "AMBIGU").sum())
    n_rej  = int((res["classe"] == "REJET").sum())
    print(f"[ok] {len(res):>6} matches — AUTO={n_auto}  AMBIGU={n_amb}  REJET={n_rej}"
          f"  | couverture CRM = {(n_auto+n_amb)/max(len(crm),1):.1%}")

    out = res.rename(columns={"id_L": "id_crm", "id_R": "id_statcom",
                              "nom_L": "nom_crm", "nom_R": "nom_statcom"})
    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    out.to_csv(args.out, index=False, encoding="utf-8")
    print(f"[ok] écrit : {args.out}")

    if args.stat_norm_out:
        # Rattache l'ID_STATCOM aux lignes brutes via NOM_BASE pour les étapes suivantes.
        stat = stat.merge(stat_clients[["NOM_BASE", "ID_STATCOM"]], on="NOM_BASE", how="left")
        if args.stat_norm_out.endswith(".csv"):
            stat.to_csv(args.stat_norm_out, index=False, encoding="utf-8")
        else:
            stat.to_excel(args.stat_norm_out, index=False)
        print(f"[ok] STATCOM enrichi écrit : {args.stat_norm_out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
