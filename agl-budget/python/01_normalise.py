#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Pipeline de matching — Script 01 : NORMALISATION
================================================================
Sprint 1 — Fondations (cf. cahier des charges §4).

Rôle : nettoyer les noms clients AVANT le matching probabiliste Splink.
  - Cleanco : suppression des formes juridiques mondiales.
  - Extension OHADA / CIV custom : SARL, SA, GIE, SAS, Ets, Cie, SCS, SNC...
  - Strip géographique / stopwords (COTE IVOIRE, ABIDJAN, AFRIQUE...).
  - Casse, accents, ponctuation, espaces multiples normalisés.

Entrée  : un fichier Excel/CSV avec au moins une colonne de nom client.
Sortie  : même table + colonne NOM_NORMALISE (et NOM_BASE sans géo).

Exemple :
    python 01_normalise.py --input CRM.xlsx --col NOM_CLIENT --output CRM_norm.xlsx

"SARL CARGIL COTE D IVOIRE" -> "CARGIL"   (NOM_BASE)
                            -> "CARGIL COTE IVOIRE" (NOM_NORMALISE, géo conservé)

Dépendances : pandas, openpyxl, cleanco (optionnel — fallback regex intégré).
"""
from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from pathlib import Path

try:
    import pandas as pd
except ImportError:  # message clair plutôt que stacktrace
    sys.exit("ERREUR : pandas requis. Installez : pip install -r requirements.txt")

# Cleanco est recommandé mais optionnel : fallback regex si absent.
try:
    from cleanco import basename as cleanco_basename  # type: ignore
    HAS_CLEANCO = True
except ImportError:
    HAS_CLEANCO = False


# --- Formes juridiques OHADA / CIV non couvertes nativement par Cleanco ------
FORMES_OHADA = [
    "SARL", "SARLU", "SA", "SAS", "SASU", "GIE", "SNC", "SCS", "SCI", "SUARL",
    "ETS", "ETABLISSEMENT", "ETABLISSEMENTS", "EURL", "SCOOP", "SCOP",
    "CIE", "COMPAGNIE", "STE", "SOCIETE", "GROUPE", "GROUP", "HOLDING",
    "ENTREPRISE", "ENTREPRISES", "EARL", "SCA",
]

# Stopwords géographiques CIV / Afrique de l'Ouest.
GEO_STOP = [
    "COTE", "DIVOIRE", "IVOIRE", "CI", "CIV", "RCI", "ABIDJAN", "PLATEAU",
    "TREICHVILLE", "YOPOUGON", "AFRIQUE", "AFRICA", "WEST", "OUEST",
    "INTERNATIONAL", "INTERNATIONALE", "SARL",
]

_FORMES_RE = re.compile(
    r"\b(" + "|".join(sorted(map(re.escape, FORMES_OHADA), key=len, reverse=True)) + r")\b"
)
_GEO_RE = re.compile(
    r"\b(" + "|".join(sorted(map(re.escape, GEO_STOP), key=len, reverse=True)) + r")\b"
)
_PUNCT_RE = re.compile(r"[^A-Z0-9 ]+")
_MULTISPACE_RE = re.compile(r"\s+")

# Articles / liaisons FR retirés de la clé de match (NOM_BASE uniquement).
ARTICLES = {"DE", "DU", "DES", "LA", "LE", "LES", "ET", "AUX", "AU"}


def _strip_articles(name: str) -> str:
    """Retire articles FR et tokens d'une seule lettre résiduels."""
    toks = [t for t in name.split() if t not in ARTICLES and len(t) > 1]
    return " ".join(toks)


def strip_accents(text: str) -> str:
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def base_clean(name: str) -> str:
    """Casse haute, sans accents, sans ponctuation, espaces normalisés."""
    if name is None:
        return ""
    s = strip_accents(str(name)).upper()
    s = _PUNCT_RE.sub(" ", s)
    s = _MULTISPACE_RE.sub(" ", s).strip()
    return s


def remove_legal_forms(name: str) -> str:
    """Supprime les formes juridiques. Cleanco si dispo, sinon regex OHADA."""
    if not name:
        return ""
    if HAS_CLEANCO:
        try:
            name = cleanco_basename(name)
        except Exception:
            pass
    # Toujours appliquer l'extension OHADA (Cleanco ne couvre pas tout l'OHADA).
    name = _FORMES_RE.sub(" ", base_clean(name))
    return _MULTISPACE_RE.sub(" ", name).strip()


def normalise(name: str):
    """
    Retourne (nom_normalise, nom_base) :
      - nom_normalise : nettoyé, formes juridiques retirées, géo CONSERVÉ.
      - nom_base      : idem mais stopwords géographiques retirés (clé de match).
    """
    cleaned = remove_legal_forms(name)
    base = _GEO_RE.sub(" ", cleaned)
    base = _MULTISPACE_RE.sub(" ", base).strip()
    base = _strip_articles(base)
    return cleaned, base


def load_table(path: Path) -> "pd.DataFrame":
    if path.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(path, dtype=str)
    if path.suffix.lower() in (".csv", ".txt"):
        return pd.read_csv(path, dtype=str, sep=None, engine="python")
    raise ValueError(f"Format non supporté : {path.suffix}")


def save_table(df: "pd.DataFrame", path: Path) -> None:
    if path.suffix.lower() in (".xlsx", ".xls"):
        df.to_excel(path, index=False)
    else:
        df.to_csv(path, index=False)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AGL BUDGET — 01 normalisation des noms clients")
    ap.add_argument("--input", "-i", required=True, help="Fichier Excel/CSV en entrée")
    ap.add_argument("--col", "-c", required=True, help="Nom de la colonne contenant le nom client")
    ap.add_argument("--output", "-o", help="Fichier de sortie (défaut : <input>_norm.xlsx)")
    ap.add_argument("--preview", action="store_true", help="Affiche un aperçu sans écrire de fichier")
    args = ap.parse_args(argv)

    in_path = Path(args.input)
    if not in_path.exists():
        sys.exit(f"ERREUR : fichier introuvable : {in_path}")

    if not HAS_CLEANCO:
        print("[info] cleanco non installé — fallback regex OHADA actif "
              "(pip install cleanco recommandé).", file=sys.stderr)

    df = load_table(in_path)
    if args.col not in df.columns:
        sys.exit(f"ERREUR : colonne '{args.col}' absente. Colonnes : {list(df.columns)}")

    res = df[args.col].fillna("").map(normalise)
    df["NOM_NORMALISE"] = [r[0] for r in res]
    df["NOM_BASE"] = [r[1] for r in res]

    n = len(df)
    n_vides = int((df["NOM_BASE"].str.len() == 0).sum())
    print(f"[ok] {n} lignes normalisées · {n_vides} noms vides après nettoyage "
          f"· cleanco={'oui' if HAS_CLEANCO else 'non (fallback)'}")

    if args.preview:
        cols = [args.col, "NOM_NORMALISE", "NOM_BASE"]
        print(df[cols].head(20).to_string(index=False))
        return 0

    out_path = Path(args.output) if args.output else in_path.with_name(in_path.stem + "_norm.xlsx")
    save_table(df, out_path)
    print(f"[ok] écrit : {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
