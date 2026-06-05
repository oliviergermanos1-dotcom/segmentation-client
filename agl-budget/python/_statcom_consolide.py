#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Consolidation intelligente des fichiers STATCOM par métier
========================================================================
Lit N fichiers STATCOM (un par métier × direction) et produit UN fichier
unifié avec colonne `client` déjà résolue selon la règle métier validée :
  - Import (Maritime/Hinterland/Aérien) → client = colonne "destinataires"
  - Export (Maritime/Hinterland/Aérien) → client = colonne "chargeurs"

Le métier est déduit du nom de fichier (insensible à la casse) :
  STATCOM AERIEN EXPORT *      → "Export Aérien"
  STATCOM AERIEN IMPORT *      → "Import Aérien"
  STATCOM EXPORT HINTERLAND *  → "Hinterland Export"
  STATCOM IMPORT HINTERLAND *  → "Hinterland Import"
  STATCOM EXPORT MARITIME *    → "Export Maritime"
  STATCOM IMPORT MARITIME *    → "Import Maritime"

Sortie : 1 fichier .xlsx avec colonnes au minimum :
  client, metier, annee, marchandise, volume_teu, volume_bulk, volume_kg
+ toutes les colonnes originales conservées (en bout) pour traçabilité.

Usage :
    python _statcom_consolide.py <dossier_in> <fichier_out.xlsx>
    python _statcom_consolide.py <fichier1.xlsx> <fichier2.xlsx> ... -o <out.xlsx>
"""
from __future__ import annotations
import argparse
import re
import sys
from pathlib import Path

import pandas as pd


# Règle métier ↔ direction (cf. config.STATCOM_METIERS).
METIER_RULES = {
    # (mots-clés à matcher dans le nom de fichier) : (metier_canonique, sens, unite_principale)
    ("AERIEN", "EXPORT"):     ("Export Aérien",     "export", "kg"),
    ("AERIEN", "IMPORT"):     ("Import Aérien",     "import", "kg"),
    ("HINTERLAND", "EXPORT"): ("Hinterland Export", "export", "teu"),
    ("HINTERLAND", "IMPORT"): ("Hinterland Import", "import", "teu"),
    ("MARITIME", "EXPORT"):   ("Export Maritime",   "export", "teu"),
    ("MARITIME", "IMPORT"):   ("Import Maritime",   "import", "teu"),
}


def metier_du_fichier(nom: str):
    """Retourne (metier, sens, unite) ou None."""
    n = nom.upper()
    for (k1, k2), (metier, sens, unite) in METIER_RULES.items():
        if k1 in n and k2 in n:
            return metier, sens, unite
    return None


def resolve_col(df, candidates):
    """Trouve la 1re colonne df dont le nom (lower/stripped) matche un candidate."""
    norm = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.strip().lower()
        if key in norm:
            return norm[key]
    return None


# Codes de conditionnement BULK (confirmé utilisateur). Tout ce qui n'est pas
# dans cet ensemble est considéré comme conteneurisé → volume en TEU.
BULK_CODES = {"SACS", "VRAC", "BRBK", "BREAK BULK", "BREAKBULK"}


def _to_num_series(s):
    """Conversion str → float robuste (gère ',', espaces, NaN)."""
    import pandas as pd
    return pd.to_numeric(s.astype(str).str.replace(r"[ ,]", "", regex=True)
                                       .str.replace(",", ".", regex=False),
                         errors="coerce").fillna(0)


def process_one(path: Path):
    """Traite UN fichier STATCOM par métier → DataFrame normalisé."""
    info = metier_du_fichier(path.name)
    if not info:
        print(f"  [skip] {path.name} : métier non reconnu", file=sys.stderr)
        return None
    metier, sens, unite = info
    df = pd.read_excel(path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    # Colonne client : destinataires (import) ou chargeurs (export).
    if sens == "import":
        c_client = resolve_col(df, ["destinataires", "destinataire", "client", "consignee"])
    else:
        c_client = resolve_col(df, ["chargeurs", "chargeur", "client", "expediteurs", "expediteur", "shipper"])
    if c_client is None:
        sys.stderr.write(f"  [warn] {path.name} : colonne client introuvable "
                         f"({'destinataires' if sens=='import' else 'chargeurs'}), "
                         f"colonnes dispos = {list(df.columns)}\n")

    # Colonnes optionnelles (on tente plusieurs alias usuels).
    # Année : confirmé utilisateur — colonne "années escale" dans les
    # fichiers STATCOM réels (escale = port call).
    c_annee = resolve_col(df, [
        "années escale", "annees escale", "année escale", "annee escale",
        "annee", "année", "year", "exercice",
        "date_escale", "date_operation", "date",
    ])
    c_mois  = resolve_col(df, [
        "mois escale", "mois", "month", "mois_escale",
    ])
    c_march = resolve_col(df, ["marchandise", "marchandises", "produit", "commodity", "designation"])
    # Confirmé utilisateur : colonnes réelles STATCOM
    c_teu    = resolve_col(df, ["NOMBRE_TEU", "nombre_teu", "nombre teu",
                                "nb_teu", "nb teu", "teu", "evp", "volume_teu", "qte_teu"])
    c_poids  = resolve_col(df, ["POIDS_MARCHANDISE", "poids_marchandise",
                                "poids marchandise", "poids", "weight", "poids_brut", "kg",
                                "tonnage", "tonnes"])
    c_condit = resolve_col(df, ["CODE_CONDIT", "code_condit", "code condit",
                                "conditionnement", "condit"])

    # Volumes numériques bruts.
    v_teu   = _to_num_series(df[c_teu])   if c_teu   else 0
    v_poids = _to_num_series(df[c_poids]) if c_poids else 0
    if c_condit is not None:
        condit = df[c_condit].astype(str).str.upper().str.strip()
        is_bulk = condit.isin(BULK_CODES)
    else:
        is_bulk = False  # pas de conditionnement → tout est conteneur si maritime/hinter
    is_air = "AÉRIEN" in metier.upper() or "AERIEN" in metier.upper()

    # Règles métier (confirmées user):
    #  - Aérien        → tout dans volume_kg (POIDS_MARCHANDISE)
    #  - Maritime/Hinterland + CODE_CONDIT ∈ {SACS, VRAC, BRBK} → volume_bulk = POIDS
    #  - Maritime/Hinterland + conteneur (autre conditionnement)→ volume_teu = NOMBRE_TEU
    if is_air:
        vol_teu_final  = 0
        vol_bulk_final = 0
        vol_kg_final   = v_poids
    else:
        # On garde TEU SEULEMENT pour les conteneurs (is_bulk=False).
        vol_teu_final  = v_teu if isinstance(v_teu, int) else v_teu.where(~is_bulk, 0)
        # Et POIDS SEULEMENT pour les bulk (is_bulk=True).
        vol_bulk_final = v_poids if isinstance(v_poids, int) else v_poids.where(is_bulk, 0)
        vol_kg_final   = 0

    out = pd.DataFrame()
    out["client"]      = df[c_client] if c_client else ""
    out["metier"]      = metier
    out["sens"]        = sens
    out["unite_principale"] = unite
    out["annee"]       = df[c_annee] if c_annee else ""
    out["mois"]        = df[c_mois]  if c_mois  else ""
    out["marchandise"] = df[c_march] if c_march else ""
    out["code_condit"] = df[c_condit] if c_condit else ""
    out["volume_teu"]  = vol_teu_final
    out["volume_bulk"] = vol_bulk_final  # en kg ou tonnes selon source
    out["volume_kg"]   = vol_kg_final
    out["fichier_source"] = path.name

    n = len(out)
    n_clients = out["client"].astype(str).str.strip().ne("").sum()
    n_annee   = out["annee"].astype(str).str.strip().ne("").sum()
    teu_sum   = out["volume_teu"].sum()  if not isinstance(out["volume_teu"].iloc[0] if len(out) else 0, int) or out["volume_teu"].dtype != object else 0
    bulk_sum  = out["volume_bulk"].sum() if not isinstance(out["volume_bulk"].iloc[0] if len(out) else 0, int) or out["volume_bulk"].dtype != object else 0
    kg_sum    = out["volume_kg"].sum()   if not isinstance(out["volume_kg"].iloc[0]  if len(out) else 0, int) or out["volume_kg"].dtype != object else 0
    print(f"  [ok] {path.name:<55} → {metier:<20} "
          f"{n:>7} l. · {int(teu_sum):>8,} TEU · {int(bulk_sum):>12,} bulk · {int(kg_sum):>12,} kg")
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AGL BUDGET — consolidation STATCOM multi-fichiers")
    ap.add_argument("inputs", nargs="+",
                    help="Soit un dossier contenant STATCOM*.xlsx, soit la liste des fichiers")
    ap.add_argument("-o", "--out", default="STATCOM.xlsx", help="Fichier consolidé en sortie")
    args = ap.parse_args(argv)

    # Si un seul argument et c'est un dossier → on glob STATCOM*.xlsx dedans.
    # Dédoublonnage : Windows est case-insensitive, donc STATCOM* et Statcom*
    # matchent les mêmes fichiers. On dédoublonne sur le nom absolu en minuscules.
    files = []
    seen = set()
    def _add(p):
        key = str(p.resolve()).lower()
        if key in seen: return
        seen.add(key); files.append(p)

    for inp in args.inputs:
        p = Path(inp)
        if p.is_dir():
            for pattern in ("STATCOM*.xlsx", "Statcom*.xlsx", "statcom*.xlsx"):
                for f in p.glob(pattern):
                    _add(f)
        elif p.is_file():
            _add(p)
        else:
            sys.exit(f"ERREUR : introuvable : {p}")
    files = sorted(files)
    if not files:
        sys.exit("ERREUR : aucun fichier STATCOM*.xlsx à traiter.")

    print(f"[1/2] Lecture de {len(files)} fichier(s) STATCOM…")
    frames = []
    for f in files:
        res = process_one(f)
        if res is not None and len(res):
            frames.append(res)

    if not frames:
        sys.exit("ERREUR : aucune donnée exploitable extraite.")

    print(f"[2/2] Concaténation et écriture…")
    result = pd.concat(frames, ignore_index=True, sort=False)
    # Nettoyage léger : trim sur client / annee.
    for c in ("client", "annee"):
        if c in result: result[c] = result[c].astype(str).str.strip()

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    result.to_excel(out, index=False)
    print(f"\n[ok] CONSOLIDÉ → {out}")
    print(f"     {len(result):,} lignes au total · {result['metier'].nunique()} métiers")
    print("     Métiers couverts :")
    for m, n in result["metier"].value_counts().items():
        print(f"        - {m:<22} {n:>7} lignes")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
