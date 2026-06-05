#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Configuration centralisée du pipeline
====================================================
UN SEUL fichier pour les noms de colonnes des 4 bases. À éditer si les
en-têtes des Excel changent — tous les scripts (02-09) lisent ici.

Décisions colonnes confirmées :
  CRM      : nom_du_compte (client) + crm_id_compte (id)
  IRIS     : client (nom) + id (id)
  RUBRIKS  : client (nom) + secteur (secteur)
  STATCOM  : la colonne CLIENT dépend du métier :
             - import (Maritime/Hinterland/Aérien) → "destinataires"
             - export (Maritime/Hinterland/Aérien) → "chargeurs"
"""

# ----------------------------------------------------------------------------
# CRM — référentiel commercial (PRIORITÉ 1 — LOCKED §3)
# ----------------------------------------------------------------------------
CRM = {
    "id":      "crm_id_compte",     # clé principale SI AGL
    "nom":     "nom_du_compte",     # nom client
    # À confirmer (utilisé si présent, sinon ignoré) :
    "secteur":     "secteur",
    "commercial":  "commercial",
    "pipeline":    "pipeline",
}

# ----------------------------------------------------------------------------
# IRIS — performance opérationnelle AGL (PRIORITÉ 2)
# ----------------------------------------------------------------------------
IRIS = {
    "id":   "id",
    "nom":  "client",
    # À confirmer :
    "cap":      "cap",
    "periode":  "periode",
    "metier":   "metier",
}

# ----------------------------------------------------------------------------
# STATCOM — marché total CIV (3 ans, 7 métiers, ~70k lignes/an)
# ----------------------------------------------------------------------------
# Règle métier ↔ colonne client (validée par l'utilisateur) :
STATCOM_METIERS = {
    "Import Maritime":     {"client_col": "destinataires", "unite_principale": "TEU", "sens": "import"},
    "Export Maritime":     {"client_col": "chargeurs",     "unite_principale": "TEU", "sens": "export"},
    "Hinterland Import":   {"client_col": "destinataires", "unite_principale": "TEU", "sens": "import"},
    "Hinterland Export":   {"client_col": "chargeurs",     "unite_principale": "TEU", "sens": "export"},
    "Import Aérien":       {"client_col": "destinataires", "unite_principale": "Kg",  "sens": "import"},
    "Export Aérien":       {"client_col": "chargeurs",     "unite_principale": "Kg",  "sens": "export"},
}

STATCOM = {
    # À confirmer (l'app détecte aussi la première colonne plausible si absent) :
    "metier":      "metier",
    "annee":       "annee",
    "marchandise": "marchandise",
    "volume_teu":  "volume_teu",
    "volume_bulk": "volume_bulk",
    "volume_kg":   "volume_kg",
}

# ----------------------------------------------------------------------------
# RUBRIKS — budget PFA 2026
# ----------------------------------------------------------------------------
RUBRIKS = {
    "nom":     "client",
    "secteur": "secteur",
    # À confirmer :
    "cap_pfa": "cap_pfa",
    "annee":   "annee",
}

# ----------------------------------------------------------------------------
# Seuils du pipeline (§3 et §4 du cahier des charges)
# ----------------------------------------------------------------------------
SEUILS = {
    # Au-dessus : match auto. Entre haut/bas : zone Claude/manuelle. En-dessous : rejet.
    "match_haut":   0.85,
    "match_bas":    0.65,
    # Blocking : on ne compare que les paires partageant ce préfixe de NOM_BASE.
    "block_prefix": 3,
    # Limite de paires candidates à scorer (sécurité mémoire) :
    "max_pairs":    5_000_000,
}

# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------
def resolve_column(df_columns, expected, alternatives=()):
    """Retrouve une colonne avec tolérance (casse, espaces, alternatives, tokens).

    Stratégie en 3 étapes:
      1. Match exact normalisé (lowercase + espaces→underscores)
      2. Match avec les alternatives passées
      3. Match par tokens : si tous les mots significatifs de `expected` sont
         présents dans la colonne, on accepte.
         Exemple : 'crm_id_compte' (tokens crm/id/compte) matche
         'CRM ID du Compte' (tokens crm/id/du/compte) — le 'du' en plus
         n'empêche pas le match.
    """
    def _norm(s):
        return str(s).strip().lower().replace(" ", "_")

    def _tokens(s):
        # Tokens significatifs : longueur >= 2, sans accents.
        import unicodedata
        s = unicodedata.normalize("NFD", str(s))
        s = "".join(c for c in s if not unicodedata.combining(c))
        s = s.lower().replace(" ", "_")
        return set(t for t in s.split("_") if len(t) >= 2)

    cols_norm = {_norm(c): c for c in df_columns}

    # 1+2. Match exact insensible casse/espaces (expected + alternatives)
    for cand in (expected,) + tuple(alternatives):
        if _norm(cand) in cols_norm:
            return cols_norm[_norm(cand)]

    # 3. Match par tokens (sous-ensemble) — tolère ajouts de mots ("du", "de", etc.)
    expected_tokens = _tokens(expected)
    if expected_tokens:
        for col in df_columns:
            col_tokens = _tokens(col)
            if expected_tokens.issubset(col_tokens):
                return col
    return None
