#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Extraction du référentiel : entités qui sont TOP acteurs
=====================================================================
Pour chaque client des tops (IRIS CAP, RUBRIKS CAP, 6 métiers STATCOM AGL),
retrouve sa ligne dans referentiel_vue (par ID puis par nom fuzzy) et produit
une extraction = sous-ensemble du référentiel limité aux top acteurs, enrichi :
  - TOP_SOURCES : dans quels tops figure l'entité
  - CAP_IRIS_2025 / CAP_RUBRIKS_2025 / POIDS_STAT_2025
  - RANG_IRIS / RANG_RUBRIKS
"""
import argparse, sys, importlib.util, os
import pandas as pd
from rapidfuzz import fuzz, process

_spec = importlib.util.spec_from_file_location(
    "norm", os.path.join(os.path.dirname(__file__), "01_normalise.py"))
norm = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(norm)
BASES = ["CRM", "STATCOM", "IRIS", "RUBRIKS"]


def nbase(x):
    try: return norm.normalise(str(x))[1]
    except Exception: return ""


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--vue", required=True)
    ap.add_argument("--tops", required=True)
    ap.add_argument("--out", default="referentiel_top_acteurs.xlsx")
    ap.add_argument("--fuzz-min", type=int, default=82)
    args = ap.parse_args(argv)

    ref = pd.read_excel(args.vue).fillna("").reset_index(drop=True)
    # index par base : id->row et nom normalisé->row
    id_idx = {b: {} for b in BASES}
    name_list = {b: [] for b in BASES}      # (nom_norm, row)
    for i, r in ref.iterrows():
        for b in BASES:
            idv = str(r.get(f"ID_{b}", "")).strip()
            if idv and idv.lower() != "nan":
                id_idx[b].setdefault(idv, i)
            nm = str(r.get(f"NOM_{b}", "")).strip()
            if nm:
                name_list[b].append((nbase(nm), i))
    name_keys = {b: [x[0] for x in name_list[b]] for b in BASES}

    def find_row(base, idv, nom):
        idv = str(idv).strip()
        if idv and idv in id_idx[base]:
            return id_idx[base][idv]
        q = nbase(nom)
        if q and name_keys[base]:
            m = process.extractOne(q, name_keys[base], scorer=fuzz.token_set_ratio)
            if m and m[1] >= args.fuzz_min:
                return name_list[base][m[2]][1]
        return None

    # accumulateurs par ligne du référentiel
    tops = {}   # row -> dict
    xl = pd.ExcelFile(args.tops)

    def feed(sheet, base, nomcol, idcol, tag, capcol=None, rangcol="RANG"):
        d = pd.read_excel(xl, sheet)
        for _, r in d.iterrows():
            row = find_row(base, r.get(idcol, ""), r.get(nomcol, ""))
            if row is None:
                continue
            t = tops.setdefault(row, {"TOP_SOURCES": set()})
            t["TOP_SOURCES"].add(tag)
            if capcol and capcol in d.columns:
                t[f"VAL_{tag}"] = r[capcol]
            if rangcol in d.columns and tag in ("IRIS_CAP", "RUBRIKS_CAP"):
                t[f"RANG_{base}"] = r[rangcol]

    feed("top300_iris_cap", "IRIS", "NOM_IRIS", "ID_IRIS", "IRIS_CAP", "CAP_2025")
    feed("top300_rubriks_cap", "RUBRIKS", "NOM_RUBRIKS", "ID_RUBRIKS", "RUBRIKS_CAP", "CAP_2025")
    for sh in [s for s in xl.sheet_names if s.startswith("top300_stat")]:
        tag = sh.replace("top300_stat_", "STAT_")
        feed(sh, "STATCOM", "NOM_STATCOM", "ID_STATCOM", tag, "POIDS_2025")

    print(f"Entités top distinctes trouvées dans le référentiel : {len(tops):,}")

    # construire l'extraction
    rows = []
    for row, t in tops.items():
        base_row = ref.loc[row].to_dict()
        srcs = sorted(t["TOP_SOURCES"])
        base_row["TOP_SOURCES"] = " ; ".join(srcs)
        base_row["NB_TOPS"] = len(srcs)
        base_row["CAP_IRIS_2025"] = t.get("VAL_IRIS_CAP", "")
        base_row["CAP_RUBRIKS_2025"] = t.get("VAL_RUBRIKS_CAP", "")
        base_row["POIDS_STAT_2025"] = sum(
            v for k, v in t.items() if k.startswith("VAL_STAT_") and isinstance(v, (int, float)))
        base_row["RANG_IRIS"] = t.get("RANG_IRIS", "")
        base_row["RANG_RUBRIKS"] = t.get("RANG_RUBRIKS", "")
        rows.append(base_row)

    cols = [f"{p}_{b}" for b in BASES for p in ("NOM", "ID")] + \
           ["NOM_CONSOLIDE", "ID_RETENU", "SECTEUR", "CONFIANCE_LIEN",
            "TOP_SOURCES", "NB_TOPS", "CAP_IRIS_2025", "CAP_RUBRIKS_2025",
            "POIDS_STAT_2025", "RANG_IRIS", "RANG_RUBRIKS"]
    out = pd.DataFrame(rows)
    cols = [c for c in cols if c in out.columns]
    out = out[cols].sort_values("NB_TOPS", ascending=False)
    out.to_excel(args.out, index=False)
    print(f"[ok] {len(out):,} top acteurs → {args.out}")
    print(f"  multi-tops (≥2 bases/métiers) : {(out['NB_TOPS']>=2).sum():,}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
