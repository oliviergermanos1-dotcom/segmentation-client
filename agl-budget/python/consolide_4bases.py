#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — consolide_4bases.py
================================================================
Consolide les 4 bases (CRM/IRIS/RUBRIKS/STATCOM) + opportunites en
UN seul Excel master, avec matching multi-stratégie haute performance
et code couleur sur chaque ligne.

OBJECTIF MATCHING ≥ 90% auto.

Stratégies de matching (en cascade, par ordre décroissant de confiance) :
  1. ID Concerto extrait du nom (RUBRIKS "NOM (1-XXXXX)")        → 100% VERT
  2. Match par ID direct via RMC (déjà fait par le pipeline)     → 100% VERT
  3. Match par alias manuel (aliases.json)                       → 100% VERT
  4. Match par NOM_BASE exact (Cleanco + OHADA + strip)         → 100% VERT
  5. Match par tokens (subset all words)                         → 95% VERT
  6. Match fuzzy RapidFuzz token_set_ratio + Jaro-Winkler        → 75-95% ORANGE
  7. Match fuzzy + bonus secteur identique                       → +5pts
  8. Pas de match → TM_NNNNNN auto                               → 0% ROUGE

Sortie : master_4bases.xlsx avec 5 onglets et coloration des cellules
STATUT_RATTACHEMENT (vert/orange/rouge) + colonne SCORE_CONFIANCE.

Usage:
    python consolide_4bases.py --data-dir ./FINAL --rmc-dir ./AGL-out \\
                               --out ./master_4bases.xlsx
    python consolide_4bases.py --data-dir ./FINAL --rmc-dir ./AGL-out \\
                               --mapping ./mapping.json --out ./master_4bases.xlsx
"""
from __future__ import annotations
import argparse
import datetime as dt
import json
import re
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path

import pandas as pd

try:
    from rapidfuzz import fuzz, process, distance
    HAS_RAPIDFUZZ = True
except ImportError:
    HAS_RAPIDFUZZ = False
    sys.stderr.write("[warn] rapidfuzz non installé — install: pip install rapidfuzz\n")

try:
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL_STYLES = True
except ImportError:
    HAS_OPENPYXL_STYLES = False


# ---------------------------------------------------------------------------
# NORMALISATION DES NOMS (clé du matching)
# ---------------------------------------------------------------------------

FORMES_OHADA = {
    "SARL", "SARLU", "SA", "SAS", "SASU", "GIE", "SNC", "SCS", "SCI", "SUARL",
    "ETS", "ETABLISSEMENT", "ETABLISSEMENTS", "EURL", "SCOOP", "SCOP",
    "CIE", "COMPAGNIE", "STE", "SOCIETE", "GROUPE", "GROUP", "HOLDING",
    "ENTREPRISE", "ENTREPRISES", "EARL", "SCA", "LTD", "LLC", "INC", "PLC",
    "GMBH", "BV", "NV", "AG", "CO", "OY", "SPA", "SRL", "EXIM",
}

GEO_STOP = {
    "COTE", "DIVOIRE", "IVOIRE", "CI", "CIV", "RCI", "ABIDJAN", "PLATEAU",
    "TREICHVILLE", "YOPOUGON", "AFRIQUE", "AFRICA", "WEST", "OUEST",
    "INTERNATIONAL", "INTERNATIONALE", "AFRICAINE", "AFRICAIN",
}

ARTICLES = {"DE", "DU", "DES", "LA", "LE", "LES", "ET", "AUX", "AU",
            "OF", "THE", "AND", "FOR", "IN", "ON", "AT"}

PUNCT_RE = re.compile(r"[^A-Z0-9 ]+")
MULTISPACE_RE = re.compile(r"\s+")


def normalize_name(s):
    """Normalisation ultra-robuste pour matching : retourne (clean, base)."""
    if not s or (isinstance(s, float) and pd.isna(s)):
        return "", ""
    s = str(s)
    # Suppression accents
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.upper()
    # Ponctuation → espace
    s = PUNCT_RE.sub(" ", s)
    s = MULTISPACE_RE.sub(" ", s).strip()
    # Tokens
    tokens = s.split()
    # 1. clean = retire formes juridiques (mais garde géo + articles)
    clean_tokens = [t for t in tokens if t not in FORMES_OHADA]
    clean = " ".join(clean_tokens)
    # 2. base = retire en plus géo + articles + tokens 1 lettre
    base_tokens = [t for t in clean_tokens
                   if t not in GEO_STOP and t not in ARTICLES and len(t) >= 2]
    base = " ".join(base_tokens)
    return clean, base


def name_tokens(s):
    """Tokens significatifs (>= 3 lettres, sans stopwords) pour blocking & subset."""
    _, base = normalize_name(s)
    return {t for t in base.split() if len(t) >= 3}


# ---------------------------------------------------------------------------
# MATCHER MULTI-STRATÉGIE
# ---------------------------------------------------------------------------

class Matcher:
    """
    Matcher haute performance pour rattacher des clients (RUBRIKS, IRIS,
    STATCOM…) à un référentiel CRM via le RMC.

    Préconstruit plusieurs index (par ID, par NOM_BASE, par token) pour des
    lookups O(1) ou O(k) au lieu de O(N).
    """

    # Seuils de confiance (RapidFuzz token_set_ratio sur 100)
    THRESHOLD_GREEN  = 92     # > 92  → vert (match confiant)
    THRESHOLD_ORANGE = 75     # 75-92 → orange (match moyen)
    # < 75 → rouge (pas de match)

    def __init__(self, rmc_df: pd.DataFrame, aliases: dict = None):
        """rmc_df doit avoir colonnes : ID_CRM, NOM_CANONIQUE, SECTEUR."""
        self.rmc = rmc_df.copy()
        # Précompute clés
        self.rmc["_NOM_CLEAN"], self.rmc["_NOM_BASE"] = zip(
            *self.rmc["NOM_CANONIQUE"].fillna("").map(normalize_name))
        # Index par ID
        self.idx_by_id = {str(r).strip(): i for i, r in
                          enumerate(self.rmc["ID_CRM"].fillna(""))
                          if str(r).strip()}
        # Index par NOM_BASE
        self.idx_by_base = defaultdict(list)
        for i, b in enumerate(self.rmc["_NOM_BASE"]):
            if b: self.idx_by_base[b].append(i)
        # Index par premier token (blocking)
        self.idx_by_token = defaultdict(list)
        for i, b in enumerate(self.rmc["_NOM_BASE"]):
            for t in b.split():
                if len(t) >= 3:
                    self.idx_by_token[t].append(i)
        # Aliases
        self.aliases = {k.upper().strip(): v for k, v in (aliases or {}).items()}
        # Pré-tableau des noms base (pour fuzzy matching batch)
        self.base_list = self.rmc["_NOM_BASE"].tolist()
        # Stats
        self.stats = defaultdict(int)

    def match(self, name, id_hint=None, secteur=None):
        """
        Retourne dict {idx, score, method, color, rmc_id_crm, rmc_nom}
        ou None si pas de match exploitable.
        """
        # 1. ID hint (Concerto extrait)
        if id_hint:
            key = str(id_hint).strip()
            if key in self.idx_by_id:
                i = self.idx_by_id[key]
                self.stats["id"] += 1
                return self._result(i, 1.0, "ID Concerto", "green")

        # 2. Aliases manuels
        if name:
            key = str(name).upper().strip()
            if key in self.aliases:
                canonical = self.aliases[key]
                # match canonical name
                _, base = normalize_name(canonical)
                if base in self.idx_by_base:
                    i = self.idx_by_base[base][0]
                    self.stats["alias"] += 1
                    return self._result(i, 1.0, "Alias manuel", "green")

        # 3. NOM_BASE exact
        _, base = normalize_name(name)
        if not base:
            return None
        if base in self.idx_by_base:
            i = self.idx_by_base[base][0]
            self.stats["exact"] += 1
            return self._result(i, 1.0, "Nom exact", "green")

        # 4. Token subset (tous les tokens cibles présents)
        target_tokens = {t for t in base.split() if len(t) >= 3}
        if not target_tokens:
            return None
        candidates = set()
        for t in target_tokens:
            candidates.update(self.idx_by_token.get(t, []))
        if not candidates:
            self.stats["no_candidates"] += 1
            return None

        candidates = list(candidates)
        # Pour chaque candidat, vérifier si tous les tokens cibles sont
        # contenus dans le candidat (token subset)
        for i in candidates:
            cand_tokens = {t for t in self.base_list[i].split() if len(t) >= 3}
            if target_tokens.issubset(cand_tokens):
                self.stats["subset"] += 1
                return self._result(i, 0.95, "Tokens inclus", "green")

        # 5. Fuzzy match (RapidFuzz token_set_ratio + Jaro-Winkler)
        if not HAS_RAPIDFUZZ:
            return None

        best_i, best_score = None, 0
        for i in candidates:
            cand = self.base_list[i]
            if not cand: continue
            s_tsr = fuzz.token_set_ratio(base, cand)
            s_jw  = distance.JaroWinkler.normalized_similarity(base, cand) * 100
            score = 0.6 * s_tsr + 0.4 * s_jw   # pondéré
            # Bonus secteur identique (+5 pts)
            if secteur and secteur.strip():
                sec_rmc = str(self.rmc.iloc[i].get("SECTEUR", "")).strip()
                if sec_rmc and sec_rmc.upper() == secteur.upper():
                    score = min(100, score + 5)
            if score > best_score:
                best_score, best_i = score, i

        if best_i is None:
            self.stats["no_fuzzy"] += 1
            return None
        if best_score >= self.THRESHOLD_GREEN:
            self.stats["fuzzy_high"] += 1
            return self._result(best_i, best_score / 100, "Fuzzy fort", "green")
        if best_score >= self.THRESHOLD_ORANGE:
            self.stats["fuzzy_medium"] += 1
            return self._result(best_i, best_score / 100, "Fuzzy moyen", "orange")
        self.stats["fuzzy_low"] += 1
        return None

    def _result(self, i, score, method, color):
        r = self.rmc.iloc[i]
        return {
            "idx": int(i),
            "score": float(score),
            "method": method,
            "color": color,
            "rmc_id_crm": str(r.get("ID_CRM", "")).strip(),
            "rmc_id_iris": str(r.get("ID_IRIS", "")).strip(),
            "rmc_id_statcom": str(r.get("ID_STATCOM", "")).strip(),
            "rmc_nom": str(r.get("NOM_CANONIQUE", "")).strip(),
            "rmc_secteur": str(r.get("SECTEUR", "")).strip(),
        }


# ---------------------------------------------------------------------------
# CHARGEMENT DES SOURCES
# ---------------------------------------------------------------------------

def _load(path):
    p = Path(path)
    if not p.exists():
        return None
    if p.suffix.lower() in (".xlsx", ".xls"):
        return pd.read_excel(p, dtype=str)
    return pd.read_csv(p, dtype=str)


def _find_file(folder: Path, pattern: str):
    if not folder.exists(): return None
    matches = sorted(folder.glob(pattern))
    return matches[0] if matches else None


def _num_series(s):
    return pd.to_numeric(s.astype(str).str.replace(r"[ ,]", "", regex=True),
                         errors="coerce").fillna(0)


def _extract_id(s, pattern):
    if not s: return None
    m = re.search(pattern, str(s))
    return m.group(1) if m else None


# ---------------------------------------------------------------------------
# RÉFÉRENTIEL PERSISTANT — IDs stables d'un run à l'autre
# ---------------------------------------------------------------------------

def load_referentiel(path: Path) -> dict:
    """Charge le référentiel ou retourne un référentiel vide."""
    if path.exists():
        with open(path, encoding="utf-8") as f:
            ref = json.load(f)
        # Sanity check : structure minimale
        ref.setdefault("clients", {})
        ref.setdefault("tm_counter", 0)
        ref.setdefault("history", [])
        return ref
    return {
        "_doc": "Référentiel client persistant AGL BUDGET. NE PAS RENOMMER LES CLÉS.",
        "_format_version": "1.0",
        "clients": {},
        "tm_counter": 0,
        "last_run": None,
        "history": []
    }


def save_referentiel(path: Path, ref: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(ref, f, indent=2, ensure_ascii=False)


def get_or_assign_id(ref: dict, nom_base: str, id_crm: str, period: str,
                     presence: list) -> tuple:
    """
    Retourne (id_unique, is_new, type_id).
    Garantit la stabilité : un même client garde son ID d'un run à l'autre.

    Stratégie :
      1. Si ID Concerto existe → l'utilise (et le lie au NOM_BASE)
      2. Sinon lookup par NOM_BASE dans le référentiel
      3. Sinon nouveau client → assigne TM_NNNNNN incrémenté
    """
    nom_base = (nom_base or "").strip()

    # 1. Client avec ID Concerto
    if id_crm and id_crm.strip():
        id_crm = id_crm.strip()
        # Cherche si l'ID est déjà connu (sous n'importe quel NOM_BASE)
        for nb, info in ref["clients"].items():
            if info.get("id_unique") == id_crm:
                # Connu — mise à jour
                info["last_seen"] = period
                info.setdefault("presence_history", {})[period] = presence
                # Ajoute le NOM_BASE comme alias si différent
                if nom_base and nom_base != nb:
                    aliases_list = info.setdefault("aliases", [])
                    if nom_base not in aliases_list:
                        aliases_list.append(nom_base)
                return id_crm, False, info.get("type_id", "Concerto")
        # Nouveau client Concerto
        key = nom_base or id_crm
        ref["clients"][key] = {
            "id_unique": id_crm,
            "type_id": "Concerto",
            "first_seen": period,
            "last_seen": period,
            "presence_history": {period: presence},
        }
        return id_crm, True, "Concerto"

    # 2. Match par NOM_BASE
    if nom_base in ref["clients"]:
        info = ref["clients"][nom_base]
        info["last_seen"] = period
        info.setdefault("presence_history", {})[period] = presence
        return info["id_unique"], False, info.get("type_id", "TM auto")

    # 3. Brand new → TM_NNNNNN
    if not nom_base:
        return None, False, None  # impossible à tracker sans nom
    ref["tm_counter"] += 1
    new_tm = f"TM{str(ref['tm_counter']).zfill(6)}"
    ref["clients"][nom_base] = {
        "id_unique": new_tm,
        "type_id": "TM auto",
        "first_seen": period,
        "last_seen": period,
        "presence_history": {period: presence},
    }
    return new_tm, True, "TM auto"


def compute_delta(ref: dict, period: str) -> dict:
    """
    Calcule les nouveautés/disparitions par rapport au run précédent.
    Retourne {new, disappeared, changed_presence, changed_secteur}.
    """
    new = []
    disappeared = []
    changed_presence = []
    for nb, info in ref["clients"].items():
        hist = info.get("presence_history", {})
        last_seen = info.get("last_seen")
        first_seen = info.get("first_seen")
        # NOUVEAU = vu pour la 1ère fois ce run
        if first_seen == period:
            new.append({
                "ID_UNIQUE": info["id_unique"],
                "NOM_BASE": nb,
                "TYPE_ID": info.get("type_id"),
                "PRESENCE_INITIALE": ",".join(hist.get(period, [])),
            })
        # DISPARU = pas vu ce run mais vu avant
        elif last_seen and last_seen != period:
            disappeared.append({
                "ID_UNIQUE": info["id_unique"],
                "NOM_BASE": nb,
                "TYPE_ID": info.get("type_id"),
                "DERNIERE_PRESENCE": last_seen,
                "BASES_OU_VU": ",".join(hist.get(last_seen, [])),
            })
        # CHANGEMENT DE PRÉSENCE = ce run vs run précédent (entrée/sortie d'une base)
        else:
            keys_sorted = sorted([k for k in hist if k <= period])
            if len(keys_sorted) >= 2:
                prev = set(hist[keys_sorted[-2]])
                curr = set(hist[keys_sorted[-1]])
                gained = curr - prev
                lost = prev - curr
                if gained or lost:
                    changed_presence.append({
                        "ID_UNIQUE": info["id_unique"],
                        "NOM_BASE": nb,
                        "GAGNÉ": ",".join(sorted(gained)) or "—",
                        "PERDU": ",".join(sorted(lost)) or "—",
                    })
    return {
        "new": new,
        "disappeared": disappeared,
        "changed_presence": changed_presence,
    }


def write_delta_xlsx(delta: dict, target: Path, period: str, period_prev: str = None) -> None:
    """Génère delta_clients.xlsx avec les changements vs run précédent."""
    synth = pd.DataFrame({
        "Indicateur": [
            "Période courante",
            "Période précédente",
            "Nouveaux clients (1ère apparition)",
            "Clients disparus (absents ce mois)",
            "Changements de présence (gain/perte base)",
        ],
        "Valeur": [
            period,
            period_prev or "—",
            len(delta["new"]),
            len(delta["disappeared"]),
            len(delta["changed_presence"]),
        ],
    })

    with pd.ExcelWriter(target, engine="openpyxl") as writer:
        synth.to_excel(writer, sheet_name="synthese_evolution", index=False)
        if delta["new"]:
            pd.DataFrame(delta["new"]).to_excel(writer, sheet_name="nouveaux_clients", index=False)
        else:
            pd.DataFrame({"Info": ["Aucun nouveau client ce run."]}).to_excel(
                writer, sheet_name="nouveaux_clients", index=False)
        if delta["disappeared"]:
            pd.DataFrame(delta["disappeared"]).to_excel(writer, sheet_name="clients_disparus", index=False)
        else:
            pd.DataFrame({"Info": ["Aucun client disparu ce run."]}).to_excel(
                writer, sheet_name="clients_disparus", index=False)
        if delta["changed_presence"]:
            pd.DataFrame(delta["changed_presence"]).to_excel(writer, sheet_name="changements_presence", index=False)
        else:
            pd.DataFrame({"Info": ["Aucun changement de présence."]}).to_excel(
                writer, sheet_name="changements_presence", index=False)
    print(f"[ok] Delta écrit : {target}")


# ---------------------------------------------------------------------------
# CONSTRUCTION DU MASTER
# ---------------------------------------------------------------------------

def build_master(data_dir: Path, rmc_dir: Path, mapping: dict, aliases: dict,
                 referentiel: dict = None, period: str = None):
    """
    Construit le DataFrame master 4 bases.
    """
    print(f"[1/6] Chargement des sources…")
    rmc = _load(rmc_dir / "RMC.xlsx")
    if rmc is None:
        sys.exit(f"ERREUR : RMC.xlsx introuvable dans {rmc_dir}")

    # Sources
    iris_path = _find_file(rmc_dir, "IRIS_norm.xlsx") or _find_file(data_dir, "IRIS*.xlsx")
    iris = _load(iris_path) if iris_path else None
    rubriks_path = _find_file(data_dir, "RUBRIKS*.xlsx")
    rubriks = _load(rubriks_path) if rubriks_path else None
    statcom_path = (_find_file(rmc_dir, "STATCOM_consolide.xlsx")
                     or _find_file(rmc_dir, "STATCOM_norm.xlsx")
                     or _find_file(data_dir, "STATCOM_consolide.xlsx")
                     or _find_file(data_dir, "STATCOM.xlsx"))
    statcom = _load(statcom_path) if statcom_path else None
    opp_path = _find_file(data_dir, "opportunites*.xlsx")
    opp = _load(opp_path) if opp_path else None

    print(f"     RMC      : {len(rmc):>7,} clients   ({rmc_dir / 'RMC.xlsx'})")
    print(f"     IRIS     : {len(iris) if iris is not None else 0:>7,} lignes    ({iris_path or 'non chargé'})")
    print(f"     RUBRIKS  : {len(rubriks) if rubriks is not None else 0:>7,} lignes    ({rubriks_path or 'non chargé'})")
    print(f"     STATCOM  : {len(statcom) if statcom is not None else 0:>7,} lignes    ({statcom_path or 'non chargé'})")
    print(f"     OPP      : {len(opp) if opp is not None else 0:>7,} lignes    ({opp_path or 'non chargé'})")

    # Initialise le matcher
    print(f"[2/6] Initialisation du matcher haute perf…")
    matcher = Matcher(rmc, aliases)
    print(f"     Index ID    : {len(matcher.idx_by_id):,}")
    print(f"     Index nom   : {len(matcher.idx_by_base):,}")
    print(f"     Index token : {len(matcher.idx_by_token):,}")

    # Préparation IRIS : CAP réel par client × année
    print(f"[3/6] Agrégation IRIS (CAP réel par client × année)…")
    iris_agg = pd.DataFrame()
    if iris is not None:
        f = mapping["fields"]["cap_reel"]
        id_col = f.get("id_column", "id")
        val_col = f.get("value_column", "montant")
        per_col = f.get("year_column", "periode")
        if id_col in iris.columns and val_col in iris.columns:
            iris["_val"] = _num_series(iris[val_col])
            iris["_year"] = iris[per_col].astype(str).str.extract(r"(20\d{2})")[0] if per_col in iris.columns else "2026"
            iris_agg = (iris.groupby([id_col, "_year"], dropna=False)["_val"]
                        .sum().reset_index()
                        .rename(columns={id_col: "ID_IRIS", "_year": "ANNEE", "_val": "CAP_REEL"}))
            iris_agg["ID_IRIS"] = iris_agg["ID_IRIS"].astype(str)
            print(f"     {len(iris_agg):,} lignes agrégées (par client IRIS × année)")

    # Préparation RUBRIKS : extraction ID + match
    print(f"[4/6] Traitement RUBRIKS (matching + extraction ID Concerto)…")
    bud_records = []
    if rubriks is not None:
        f_cap = mapping["fields"]["cap_pfa"]
        id_pat = f_cap.get("id_extract_pattern", r"\(([A-Z0-9\-]+)\)\s*$")
        id_col_extract = f_cap.get("id_extract_column", "CLIENT")
        years_map = f_cap.get("years", {"2026": {"column": "B26", "multiplier": 1000}})
        # Détecte la colonne "client" / "CLIENT" tolérant casse
        client_col_actual = next((c for c in rubriks.columns
                                  if c.strip().upper() == id_col_extract.upper()), None)
        if client_col_actual:
            n_matched = 0
            n_total = 0
            for idx, row in rubriks.iterrows():
                name = row[client_col_actual]
                # Ignore "Total Customer" et lignes vides
                if not name or str(name).strip().lower() == "total customer":
                    continue
                n_total += 1
                # Extrait ID
                id_hint = _extract_id(name, id_pat)
                # Tente match
                m = matcher.match(name, id_hint=id_hint, secteur=row.get("SECTEUR"))
                if m: n_matched += 1
                # Pour chaque année, lit B26/2025/2024
                for year, conf in years_map.items():
                    col = conf.get("column")
                    mult = conf.get("multiplier", 1)
                    if col not in rubriks.columns: continue
                    val = _num_series(pd.Series([row[col]]))[0] * mult
                    bud_records.append({
                        "match": m, "id_hint": id_hint, "name": str(name),
                        "year": year, "cap_pfa": val,
                    })
            print(f"     {n_matched:,} / {n_total:,} clients RUBRIKS matchés ({n_matched/max(n_total,1):.0%})")
            print(f"     {len(bud_records):,} enregistrements PFA (× années)")

    # Préparation STATCOM : volumes par client × année
    print(f"[5/6] Agrégation STATCOM (volumes par client × année)…")
    statcom_agg = pd.DataFrame()
    if statcom is not None:
        cli_col   = next((c for c in statcom.columns if c.lower().strip() == "client"), None)
        year_col  = next((c for c in statcom.columns if c.lower().strip() == "annee"), None)
        teu_col   = next((c for c in statcom.columns if c.lower().strip() == "volume_teu"), None)
        bulk_col  = next((c for c in statcom.columns if c.lower().strip() == "volume_bulk"), None)
        kg_col    = next((c for c in statcom.columns if c.lower().strip() == "volume_kg"), None)
        if cli_col and year_col:
            for c in (teu_col, bulk_col, kg_col):
                if c: statcom[c] = _num_series(statcom[c])
            statcom["_year"] = statcom[year_col].astype(str).str.extract(r"(20\d{2})")[0]
            # NOM_BASE pour matching
            statcom["_nom_base"] = statcom[cli_col].map(lambda x: normalize_name(x)[1])
            agg_cols = {}
            if teu_col:  agg_cols["VOLUME_TEU"]  = (teu_col, "sum")
            if bulk_col: agg_cols["VOLUME_BULK_KG"] = (bulk_col, "sum")
            if kg_col:   agg_cols["VOLUME_KG_AERIEN"] = (kg_col, "sum")
            agg_cols["NB_OPS"] = (cli_col, "count")
            statcom_agg = (statcom.groupby(["_nom_base", "_year"], dropna=False)
                           .agg(**agg_cols).reset_index()
                           .rename(columns={"_nom_base": "NOM_BASE", "_year": "ANNEE"}))
            print(f"     {len(statcom_agg):,} lignes agrégées (par client STATCOM × année)")

    # Pipeline opportunités
    print(f"[6/6] Pipeline opportunités (CAP brut + pondéré)…")
    opp_agg = pd.DataFrame()
    if opp is not None:
        f = mapping["fields"]["pipeline_cap_brut"]
        id_col = f.get("id_column", "CRM ID du Compte")
        val_col = f.get("value_column", "CAP Potentiel Global (€)")
        stat_col = f.get("filter_column", "Statut")
        filter_vals = set(f.get("filter_values", ["Ouvert(e)"]))
        proba_col = mapping["fields"]["pipeline_cap_pondere"].get("weight_column", "Probabilité de succès")
        if id_col in opp.columns and val_col in opp.columns:
            opp["_val"] = _num_series(opp[val_col])
            opp["_proba"] = opp[proba_col].astype(str).str.replace("%", "").str.replace(",", ".") if proba_col in opp.columns else "0"
            opp["_proba"] = pd.to_numeric(opp["_proba"], errors="coerce").fillna(0)
            opp["_proba"] = opp["_proba"].where(opp["_proba"] <= 1, opp["_proba"] / 100)
            opp["_pond"] = opp["_val"] * opp["_proba"]
            mask_open = (opp[stat_col].astype(str).isin(filter_vals)
                         if stat_col in opp.columns else True)
            opp_open = opp[mask_open] if isinstance(mask_open, pd.Series) else opp
            opp_agg = (opp_open.groupby(id_col, dropna=False)
                       .agg(NB_OPP_OUVERTES=(id_col, "count"),
                            PIPELINE_CAP_BRUT=("_val", "sum"),
                            PIPELINE_CAP_POND=("_pond", "sum"))
                       .reset_index()
                       .rename(columns={id_col: "ID_CRM"}))
            opp_agg["ID_CRM"] = opp_agg["ID_CRM"].astype(str)
            print(f"     {len(opp_agg):,} clients CRM avec pipeline ouvert")

    # ----------------------------------------------------------------------
    # CONSTRUCTION DU MASTER
    # ----------------------------------------------------------------------
    print(f"\n[BUILD] Construction du master 4 bases…")

    # On démarre du RMC (CRM-centric, IDs Concerto comme clés primaires)
    master = rmc.copy()
    master["ID_UNIQUE"] = master["ID_CRM"].fillna("").astype(str)
    master["NOM_CLIENT"] = master["NOM_CANONIQUE"].fillna("")
    master["TYPE_ID"] = "Concerto (CRM)"
    master["STATUT_MATCHING"] = master.apply(
        lambda r: ("★ Triple match" if (str(r.get("ID_IRIS", "") or "").strip()
                                         and str(r.get("ID_STATCOM", "") or "").strip())
                   else "CRM + IRIS" if str(r.get("ID_IRIS", "") or "").strip()
                   else "CRM + STATCOM" if str(r.get("ID_STATCOM", "") or "").strip()
                   else "CRM seul"), axis=1)
    master["COULEUR_MATCH"] = "green"  # tous les CRM sont 100% rattachés par défaut
    master["SCORE_MATCH"] = 1.0

    # Pour les RUBRIKS clients : on matche via Matcher (sera utile pour
    # remonter le CAP_PFA aux clients RMC; mais aussi pour identifier d'éventuels
    # clients RUBRIKS hors CRM — peu probable étant donné le format ID Concerto)

    # Pour STATCOM : on a déjà les ALIAS_STATCOM dans le RMC (= NOM_BASE).
    # Les clients STATCOM hors RMC = white spaces, à ajouter en TM_xxx.

    # ----- 1. Enrichissement par année pour les clients RMC ---------------
    rmc_by_id = {row["ID_CRM"]: i for i, row in rmc.iterrows() if str(row.get("ID_CRM", "")).strip()}

    # IRIS : CAP_REEL par année. Jointure via ID_IRIS du RMC.
    cap_reel_by_id_year = {}
    if len(iris_agg):
        for _, row in iris_agg.iterrows():
            cap_reel_by_id_year[(row["ID_IRIS"], str(row["ANNEE"]))] = float(row["CAP_REEL"] or 0)

    # RUBRIKS : CAP_PFA par ID_CRM × année (depuis bud_records via match)
    cap_pfa_by_id_year = defaultdict(float)
    pfa_unmatched = []
    for rec in bud_records:
        m = rec["match"]
        if m:
            cap_pfa_by_id_year[(m["rmc_id_crm"], rec["year"])] += rec["cap_pfa"]
        else:
            pfa_unmatched.append(rec)

    # STATCOM : volumes par NOM_BASE × année.
    # On rattache via ALIAS_STATCOM du RMC (qui contient le NOM_BASE STATCOM).
    vol_by_alias_year = {}
    if len(statcom_agg):
        for _, row in statcom_agg.iterrows():
            key = (row["NOM_BASE"], str(row["ANNEE"]))
            vol_by_alias_year[key] = {
                "VOLUME_TEU":       float(row.get("VOLUME_TEU", 0) or 0),
                "VOLUME_BULK_KG":   float(row.get("VOLUME_BULK_KG", 0) or 0),
                "VOLUME_KG_AERIEN": float(row.get("VOLUME_KG_AERIEN", 0) or 0),
                "NB_OPS":           int(row.get("NB_OPS", 0) or 0),
            }

    # Pipeline opportunités par ID_CRM
    pipe_by_id = {}
    if len(opp_agg):
        for _, row in opp_agg.iterrows():
            pipe_by_id[str(row["ID_CRM"])] = {
                "NB_OPP_OUVERTES":   int(row.get("NB_OPP_OUVERTES", 0)),
                "PIPELINE_CAP_BRUT": float(row.get("PIPELINE_CAP_BRUT", 0)),
                "PIPELINE_CAP_POND": float(row.get("PIPELINE_CAP_POND", 0)),
            }

    # ----- 2. White spaces STATCOM (clients absents du RMC, TM_xxx) -------
    rmc_alias_set = set(rmc["ALIAS_STATCOM"].fillna("").astype(str))
    statcom_clients = set()
    if len(statcom_agg):
        statcom_clients = set(statcom_agg["NOM_BASE"].fillna(""))
    white_spaces = sorted(statcom_clients - rmc_alias_set - {""})
    print(f"[BUILD] {len(white_spaces):,} clients STATCOM absents du RMC (white spaces TM_xxx)")

    # On crée des fiches "client" pour chaque WS, avec ID issu du référentiel
    # PERSISTANT (TM stable d'un run à l'autre).
    ws_rows = []
    n_new_ws = 0
    for alias in white_spaces:
        tm_id, is_new, type_id = get_or_assign_id(
            referentiel, alias, id_crm=None,
            period=period, presence=["STATCOM"])
        if tm_id is None:
            continue
        if is_new: n_new_ws += 1
        ws_rows.append({
            "ID_UNIQUE":      tm_id,
            "TYPE_ID":        "TM auto (STATCOM hors CRM)",
            "ID_CRM":         "",
            "ID_IRIS":        "",
            "ID_STATCOM":     alias,
            "NOM_CANONIQUE":  alias,
            "NOM_CLIENT":     alias,
            "SECTEUR":        "",
            "ALIAS_STATCOM":  alias,
            "ALIAS_CRM":      "",
            "ALIAS_IRIS":     "",
            "SOURCE_SECTEUR": "",
            "LOCKED":         False,
            "SCORE_MATCH_IRIS": "",
            "SCORE_MATCH_STATCOM": "",
            "STATUT_MATCHING": "STATCOM hors AGL (white space)" + (" — NOUVEAU" if is_new else ""),
            "COULEUR_MATCH":   "red",
            "SCORE_MATCH":     0.0,
            "EST_NOUVEAU":     is_new,
        })
    ws_df = pd.DataFrame(ws_rows)
    if n_new_ws:
        print(f"[BUILD] Dont {n_new_ws:,} NOUVEAUX white spaces (TM_xxx jamais vus avant)")

    # Maintenant on enregistre aussi les clients CRM/RMC dans le référentiel
    # (pour garder leur historique de présence par base).
    n_new_crm = 0
    for _, r in rmc.iterrows():
        nom_base = normalize_name(r.get("NOM_CANONIQUE", ""))[1]
        id_crm = str(r.get("ID_CRM", "") or "").strip()
        if not id_crm and not nom_base:
            continue
        presence = ["CRM"]
        if str(r.get("ID_IRIS", "") or "").strip():    presence.append("IRIS")
        if str(r.get("ID_STATCOM", "") or "").strip(): presence.append("STATCOM")
        _, is_new, _ = get_or_assign_id(
            referentiel, nom_base or id_crm, id_crm, period, presence)
        if is_new: n_new_crm += 1
    if n_new_crm:
        print(f"[BUILD] Dont {n_new_crm:,} NOUVEAUX clients CRM jamais vus avant")

    # Concatène RMC + white spaces
    master_full = pd.concat([master, ws_df], ignore_index=True, sort=False)

    # ----- 3. Production de 2 vues : master (cumul) et master_par_annee ---
    years = sorted({y for _, y in cap_reel_by_id_year.keys()} |
                   {y for _, y in cap_pfa_by_id_year.keys()} |
                   {y for _, y in vol_by_alias_year.keys()},
                   key=lambda s: int(s) if s and s.isdigit() else 0)
    print(f"[BUILD] Années couvertes : {years}")

    # Vue par année
    par_annee_rows = []
    for _, row in master_full.iterrows():
        id_crm     = str(row.get("ID_CRM", "") or "")
        id_iris    = str(row.get("ID_IRIS", "") or "")
        alias_stat = str(row.get("ALIAS_STATCOM", "") or "")
        for y in years:
            cap_reel = cap_reel_by_id_year.get((id_iris, y), 0.0) if id_iris else 0.0
            cap_pfa  = cap_pfa_by_id_year.get((id_crm, y), 0.0) if id_crm else 0.0
            vol = vol_by_alias_year.get((alias_stat, y), {}) if alias_stat else {}
            par_annee_rows.append({
                "ID_UNIQUE":       row["ID_UNIQUE"],
                "NOM_CLIENT":      row["NOM_CLIENT"],
                "SECTEUR":         row.get("SECTEUR", ""),
                "STATUT":          row["STATUT_MATCHING"],
                "ANNEE":           int(y) if y and y.isdigit() else None,
                "CAP_REEL":        cap_reel,
                "CAP_PFA":         cap_pfa,
                "ECART":           cap_reel - cap_pfa,
                "ECART_PCT":       (cap_reel - cap_pfa) / cap_pfa if cap_pfa > 0 else None,
                "VOLUME_TEU":      vol.get("VOLUME_TEU", 0.0),
                "VOLUME_BULK_KG":  vol.get("VOLUME_BULK_KG", 0.0),
                "VOLUME_KG_AERIEN":vol.get("VOLUME_KG_AERIEN", 0.0),
                "NB_OPERATIONS":   vol.get("NB_OPS", 0),
            })
    par_annee_df = pd.DataFrame(par_annee_rows)

    # Vue master : agrégat 3 ans (somme)
    if len(par_annee_df):
        master_agg = (par_annee_df.groupby("ID_UNIQUE", dropna=False)
                      .agg(CAP_REEL_TOTAL=("CAP_REEL", "sum"),
                           CAP_PFA_TOTAL=("CAP_PFA", "sum"),
                           ECART_TOTAL=("ECART", "sum"),
                           VOLUME_TEU_TOTAL=("VOLUME_TEU", "sum"),
                           VOLUME_BULK_KG_TOTAL=("VOLUME_BULK_KG", "sum"),
                           VOLUME_KG_AERIEN_TOTAL=("VOLUME_KG_AERIEN", "sum"),
                           NB_OPS_TOTAL=("NB_OPERATIONS", "sum"))
                      .reset_index())
        master_final = master_full.merge(master_agg, on="ID_UNIQUE", how="left")
    else:
        master_final = master_full.copy()
    # Garantit les colonnes attendues même si aucune donnée par année.
    for col in ("CAP_REEL_TOTAL", "CAP_PFA_TOTAL", "ECART_TOTAL",
                "VOLUME_TEU_TOTAL", "VOLUME_BULK_KG_TOTAL",
                "VOLUME_KG_AERIEN_TOTAL", "NB_OPS_TOTAL"):
        if col not in master_final.columns:
            master_final[col] = 0
        else:
            master_final[col] = master_final[col].fillna(0)

    # Pipeline (par ID_CRM)
    if pipe_by_id:
        master_final["NB_OPP_OUVERTES"]   = master_final["ID_CRM"].map(
            lambda i: pipe_by_id.get(str(i), {}).get("NB_OPP_OUVERTES", 0))
        master_final["PIPELINE_CAP_BRUT"] = master_final["ID_CRM"].map(
            lambda i: pipe_by_id.get(str(i), {}).get("PIPELINE_CAP_BRUT", 0.0))
        master_final["PIPELINE_CAP_POND"] = master_final["ID_CRM"].map(
            lambda i: pipe_by_id.get(str(i), {}).get("PIPELINE_CAP_POND", 0.0))
    else:
        master_final["NB_OPP_OUVERTES"] = 0
        master_final["PIPELINE_CAP_BRUT"] = 0
        master_final["PIPELINE_CAP_POND"] = 0

    # Indicateurs synthétiques
    master_final["ECART_PCT_GLOBAL"] = master_final.apply(
        lambda r: (r.get("ECART_TOTAL", 0) / r.get("CAP_PFA_TOTAL", 1)
                   if r.get("CAP_PFA_TOTAL", 0) > 0 else None), axis=1)
    master_final["PRESENCE_CRM"]     = master_final["ID_CRM"].fillna("").astype(str).str.len().gt(0).map({True: "✓", False: "—"})
    master_final["PRESENCE_IRIS"]    = master_final["ID_IRIS"].fillna("").astype(str).str.len().gt(0).map({True: "✓", False: "—"})
    master_final["PRESENCE_STATCOM"] = master_final["ID_STATCOM"].fillna("").astype(str).str.len().gt(0).map({True: "✓", False: "—"})
    master_final["PRESENCE_RUBRIKS"] = master_final["CAP_PFA_TOTAL"].fillna(0).gt(0).map({True: "✓", False: "—"})

    # Stats matching
    n_master = len(master_final)
    n_green  = int((master_final["COULEUR_MATCH"] == "green").sum())
    n_orange = int((master_final["COULEUR_MATCH"] == "orange").sum())
    n_red    = int((master_final["COULEUR_MATCH"] == "red").sum())
    print(f"\n[BUILD] Master final : {n_master:,} lignes")
    print(f"     VERT (match confiant ≥ 92%)  : {n_green:>6,}  ({n_green/n_master:.0%})")
    print(f"     ORANGE (match moyen 75-92%)  : {n_orange:>6,}  ({n_orange/n_master:.0%})")
    print(f"     ROUGE (non rattaché TM_xxx)  : {n_red:>6,}  ({n_red/n_master:.0%})")

    return {
        "master": master_final,
        "par_annee": par_annee_df,
        "matcher_stats": dict(matcher.stats),
        "years": years,
        "n_green": n_green,
        "n_orange": n_orange,
        "n_red": n_red,
        "n_master": n_master,
    }


# ---------------------------------------------------------------------------
# ÉCRITURE EXCEL AVEC COULEURS
# ---------------------------------------------------------------------------

COLOR_FILLS = {
    "green":  PatternFill(start_color="16a34a", end_color="16a34a", fill_type="solid"),
    "orange": PatternFill(start_color="d97706", end_color="d97706", fill_type="solid"),
    "red":    PatternFill(start_color="dc2626", end_color="dc2626", fill_type="solid"),
} if HAS_OPENPYXL_STYLES else {}

WHITE_FONT = Font(color="FFFFFF", bold=True) if HAS_OPENPYXL_STYLES else None


def write_xlsx(result: dict, target: Path, mapping: dict):
    """Écrit master_4bases.xlsx avec coloration des STATUT."""
    master = result["master"]
    par_annee = result["par_annee"]

    # Sélection + ordre des colonnes
    cols_master = [
        "ID_UNIQUE", "TYPE_ID", "ID_CRM", "ID_IRIS", "ID_STATCOM",
        "NOM_CLIENT", "SECTEUR", "ALIAS_STATCOM",
        "STATUT_MATCHING", "SCORE_MATCH",
        "PRESENCE_CRM", "PRESENCE_IRIS", "PRESENCE_STATCOM", "PRESENCE_RUBRIKS",
        "CAP_REEL_TOTAL", "CAP_PFA_TOTAL", "ECART_TOTAL", "ECART_PCT_GLOBAL",
        "VOLUME_TEU_TOTAL", "VOLUME_BULK_KG_TOTAL", "VOLUME_KG_AERIEN_TOTAL",
        "NB_OPS_TOTAL",
        "NB_OPP_OUVERTES", "PIPELINE_CAP_BRUT", "PIPELINE_CAP_POND",
    ]
    cols_master_existing = [c for c in cols_master if c in master.columns]
    master_out = master[cols_master_existing].copy()

    # KPI global
    n_master = result["n_master"]
    n_green = result["n_green"]
    n_orange = result["n_orange"]
    n_red = result["n_red"]
    kpis = pd.DataFrame({
        "Indicateur": [
            "Nb clients consolidés",
            "Clients CRM rattachés (Concerto)",
            "Clients TM auto (white spaces STATCOM)",
            "Match VERT (≥ 92% confiance)",
            "Match ORANGE (75-92%)",
            "Match ROUGE (TM auto, non rattaché)",
            "% rattachés (vert+orange)",
            "CAP RÉEL total (FCFA, toutes années)",
            "CAP PFA total (FCFA, toutes années)",
            "Écart global (FCFA)",
            "Volume TEU total",
            "Volume Bulk total (kg)",
            "Volume Kg total (aérien)",
            "Pipeline opportunités ouvertes (nb)",
            "Pipeline CAP brut (€/FCFA)",
            "Pipeline CAP pondéré",
            "Années couvertes",
        ],
        "Valeur": [
            n_master,
            int((master["TYPE_ID"] == "Concerto (CRM)").sum()),
            int((master["TYPE_ID"].str.startswith("TM")).sum()),
            n_green, n_orange, n_red,
            f"{(n_green + n_orange) / max(n_master, 1):.1%}",
            master["CAP_REEL_TOTAL"].sum() if "CAP_REEL_TOTAL" in master else 0,
            master["CAP_PFA_TOTAL"].sum() if "CAP_PFA_TOTAL" in master else 0,
            (master["CAP_REEL_TOTAL"].sum() - master["CAP_PFA_TOTAL"].sum())
                if "CAP_REEL_TOTAL" in master and "CAP_PFA_TOTAL" in master else 0,
            master["VOLUME_TEU_TOTAL"].sum() if "VOLUME_TEU_TOTAL" in master else 0,
            master["VOLUME_BULK_KG_TOTAL"].sum() if "VOLUME_BULK_KG_TOTAL" in master else 0,
            master["VOLUME_KG_AERIEN_TOTAL"].sum() if "VOLUME_KG_AERIEN_TOTAL" in master else 0,
            master["NB_OPP_OUVERTES"].sum() if "NB_OPP_OUVERTES" in master else 0,
            master["PIPELINE_CAP_BRUT"].sum() if "PIPELINE_CAP_BRUT" in master else 0,
            master["PIPELINE_CAP_POND"].sum() if "PIPELINE_CAP_POND" in master else 0,
            ", ".join(result["years"]),
        ],
    })

    # Diagnostic rattachement
    stats = result["matcher_stats"]
    diag = pd.DataFrame({
        "Méthode de matching": list(stats.keys()),
        "Nb clients": list(stats.values()),
    }).sort_values("Nb clients", ascending=False)

    print(f"\n[WRITE] Écriture : {target}")
    with pd.ExcelWriter(target, engine="openpyxl") as writer:
        kpis.to_excel(writer, sheet_name="kpi_global", index=False)
        master_out.to_excel(writer, sheet_name="clients_master", index=False)
        par_annee.to_excel(writer, sheet_name="clients_par_annee", index=False)
        diag.to_excel(writer, sheet_name="diagnostic_matching", index=False)
        pd.DataFrame({"mapping_json": [json.dumps(mapping, indent=2, ensure_ascii=False)]}).to_excel(
            writer, sheet_name="mapping_utilise", index=False)

        # Coloration cellules STATUT_MATCHING & SCORE_MATCH dans clients_master
        if HAS_OPENPYXL_STYLES:
            ws = writer.sheets["clients_master"]
            # Localise les colonnes
            headers = [cell.value for cell in ws[1]]
            try:
                col_statut = headers.index("STATUT_MATCHING") + 1
            except ValueError:
                col_statut = None
            try:
                col_score = headers.index("SCORE_MATCH") + 1
            except ValueError:
                col_score = None
            for row_idx, (_, mrow) in enumerate(master.iterrows(), start=2):
                color = mrow.get("COULEUR_MATCH", "")
                fill = COLOR_FILLS.get(color)
                if fill is None: continue
                for col in (col_statut, col_score):
                    if col:
                        cell = ws.cell(row=row_idx, column=col)
                        cell.fill = fill
                        cell.font = WHITE_FONT
                        cell.alignment = Alignment(horizontal="center")

            # Ajuste largeur de quelques colonnes clés
            widths = {"ID_UNIQUE": 18, "NOM_CLIENT": 36, "SECTEUR": 22,
                      "STATUT_MATCHING": 28, "TYPE_ID": 22}
            for header, w in widths.items():
                if header in headers:
                    ws.column_dimensions[get_column_letter(headers.index(header) + 1)].width = w

    print(f"[ok] Master 4 bases écrit : {target}")
    print(f"     Onglets : kpi_global, clients_master, clients_par_annee, diagnostic_matching, mapping_utilise")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AGL BUDGET — consolidation 4 bases haute performance avec référentiel persistant")
    ap.add_argument("--data-dir",    required=True, help="Dossier source (CRM, IRIS, RUBRIKS, STATCOMs, opportunites)")
    ap.add_argument("--rmc-dir",     required=True, help="Dossier AGL-out (avec RMC.xlsx, STATCOM_consolide.xlsx)")
    ap.add_argument("--mapping",     default=None,  help="mapping.json (sinon defaults AGL CI 2026)")
    ap.add_argument("--aliases",     default=None,  help="aliases.json (corrections nom)")
    ap.add_argument("--referentiel", default=None,
                    help="Référentiel persistant client.json (créé si absent). "
                         "Défaut: <rmc-dir>/referentiel_clients.json")
    ap.add_argument("--period",      default=None,
                    help="Période YYYY-MM (défaut: mois courant)")
    ap.add_argument("--out",         default=None,
                    help="Fichier master xlsx de sortie. "
                         "Défaut: <rmc-dir>/master_4bases.xlsx")
    ap.add_argument("--delta-out",   default=None,
                    help="Fichier delta xlsx de sortie. "
                         "Défaut: <rmc-dir>/delta_clients_<period>.xlsx")
    args = ap.parse_args(argv)

    data_dir = Path(args.data_dir)
    rmc_dir  = Path(args.rmc_dir)
    target   = Path(args.out) if args.out else rmc_dir / "master_4bases.xlsx"

    # Période courante (YYYY-MM)
    period = args.period or dt.datetime.now().strftime("%Y-%m")
    print(f"[init] Période : {period}\n")

    # Mapping
    if args.mapping:
        with open(args.mapping, encoding="utf-8") as f:
            mapping = json.load(f)
    else:
        default_path = Path(__file__).parent / "mapping.json"
        with open(default_path, encoding="utf-8") as f:
            mapping = json.load(f)

    # Aliases
    aliases = {}
    ap_path = (Path(args.aliases) if args.aliases else Path(__file__).parent / "aliases.json")
    if ap_path.exists():
        with open(ap_path, encoding="utf-8") as f:
            data = json.load(f)
            aliases = data.get("RUBRIKS", {}) if "RUBRIKS" in data else data

    # Référentiel persistant
    ref_path = Path(args.referentiel) if args.referentiel else rmc_dir / "referentiel_clients.json"
    referentiel = load_referentiel(ref_path)
    period_prev = referentiel.get("last_run")
    n_known = len(referentiel["clients"])
    print(f"[init] Référentiel : {ref_path}")
    print(f"       {n_known:,} clients déjà connus (dernier run : {period_prev or '—'})")
    print(f"       Compteur TM actuel : {referentiel['tm_counter']}\n")

    # BUILD
    result = build_master(data_dir, rmc_dir, mapping, aliases, referentiel, period)

    # Sauve le référentiel mis à jour
    referentiel["last_run"] = period
    referentiel["history"].append({
        "period": period,
        "date": dt.datetime.now().isoformat(timespec="seconds"),
        "n_clients_total": len(referentiel["clients"]),
        "tm_counter": referentiel["tm_counter"],
    })
    save_referentiel(ref_path, referentiel)
    print(f"\n[ok] Référentiel mis à jour : {ref_path}")
    print(f"     {len(referentiel['clients']):,} clients au total · "
          f"compteur TM : {referentiel['tm_counter']}")

    # Master
    write_xlsx(result, target, mapping)

    # Delta
    delta = compute_delta(referentiel, period)
    delta_target = (Path(args.delta_out) if args.delta_out
                    else rmc_dir / f"delta_clients_{period}.xlsx")
    write_delta_xlsx(delta, delta_target, period, period_prev)
    print(f"\n[summary] Période {period}")
    print(f"  Nouveaux clients     : {len(delta['new']):>6,}")
    print(f"  Clients disparus     : {len(delta['disappeared']):>6,}")
    print(f"  Changements présence : {len(delta['changed_presence']):>6,}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
