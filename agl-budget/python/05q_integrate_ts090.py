#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Intègre les candidats étape 2 (TOKEN_SORT ≥ seuil) dans ref_vue
============================================================================
Lit research_nomatch_large.xlsx, garde les paires TOKEN_SORT ≥ --min (0.90),
meilleur match par client (anti-chaînage), fusionne dans le référentiel,
colore les lignes ajoutées (orange), recalcule les non-matchés.
"""
import argparse, sys
from collections import defaultdict
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

BASES = ["CRM", "STATCOM", "IRIS", "RUBRIKS"]


class UF:
    def __init__(s): s.p = {}
    def find(s, x):
        s.p.setdefault(x, x)
        while s.p[x] != x: s.p[x] = s.p[s.p[x]]; x = s.p[x]
        return x
    def union(s, a, b): s.p[s.find(a)] = s.find(b)


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--vue", required=True)
    ap.add_argument("--research", required=True)
    ap.add_argument("--min", type=float, default=0.90)
    ap.add_argument("--out-vue", default="referentiel_vue_maj2.xlsx")
    ap.add_argument("--out-nonmatch", default="non_matches_maj2.xlsx")
    args = ap.parse_args(argv)

    # 1. liens TOKEN_SORT >= min, meilleur par A et par B
    xl = pd.ExcelFile(args.research)
    raw = []
    for sh in xl.sheet_names:
        if "_x_" not in sh: continue
        ba, bb = sh.split("_x_")
        d = pd.read_excel(xl, sh)
        if "TOKEN_SORT" not in d.columns: continue
        d = d[d["TOKEN_SORT"] >= args.min]
        for _, r in d.iterrows():
            raw.append((float(r["TOKEN_SORT"]), ba, str(r[f"ID_{ba}"]),
                        bb, str(r[f"ID_{bb}"])))
    raw.sort(reverse=True)
    usedA, usedB, links = set(), set(), []
    for ts, ba, ida, bb, idb in raw:        # meilleur par client (1-1)
        ka, kb = (ba, ida), (bb, idb)
        if ka in usedA or kb in usedB:
            continue
        usedA.add(ka); usedB.add(kb); links.append((ba, ida, bb, idb))
    print(f"Liens TOKEN_SORT ≥ {args.min} retenus (1-1) : {len(links):,}")

    # 2. merge dans ref_vue
    v = pd.read_excel(args.vue).fillna("").reset_index(drop=True)
    idx = {b: {} for b in BASES}
    for i, row in v.iterrows():
        for b in BASES:
            val = str(row.get(f"ID_{b}", "")).strip()
            if val and val.lower() != "nan":
                idx[b].setdefault(val, i)
    uf = UF(); fused = set(); n_ok = 0
    for ba, ida, bb, idb in links:
        ra, rb = idx[ba].get(ida), idx[bb].get(idb)
        if ra is None or rb is None or ra == rb: continue
        uf.union(ra, rb); fused.add(ra); fused.add(rb); n_ok += 1
    print(f"Liens appliqués : {n_ok:,}")

    groups = defaultdict(list)
    for i in range(len(v)): groups[uf.find(i)].append(i)

    def pick(ms, col):
        for i in ms:
            x = str(v.at[i, col]).strip()
            if x and x.lower() != "nan": return v.at[i, col]
        return ""

    rows, flags = [], []
    for root, ms in groups.items():
        isf = any(m in fused for m in ms) and len(ms) > 1
        rec = {}
        for b in BASES:
            rec[f"NOM_{b}"] = pick(ms, f"NOM_{b}"); rec[f"ID_{b}"] = pick(ms, f"ID_{b}")
        noms = [rec[f"NOM_{b}"] for b in BASES if rec[f"NOM_{b}"]]
        rec["NOM_CONSOLIDE"] = max(noms, key=len) if noms else pick(ms, "NOM_CONSOLIDE")
        rec["ID_RETENU"] = (rec["ID_CRM"] or rec["ID_RUBRIKS"] or rec["ID_IRIS"]
                            or rec["ID_STATCOM"] or pick(ms, "ID_RETENU"))
        rec["SECTEUR"] = pick(ms, "SECTEUR")
        rec["CONFIANCE_LIEN"] = "RATTRAPE_E2" if isf else pick(ms, "CONFIANCE_LIEN")
        rows.append(rec); flags.append(isf)
    cols = [f"{p}_{b}" for b in BASES for p in ("NOM", "ID")] + \
           ["NOM_CONSOLIDE", "ID_RETENU", "SECTEUR", "CONFIANCE_LIEN"]
    out = pd.DataFrame(rows)
    order = sorted(range(len(out)), key=lambda i: (not flags[i],))
    out = out.iloc[order][cols]; flags = [flags[i] for i in order]
    print(f"Entités : {len(v):,} → {len(out):,}")
    out.to_excel(args.out_vue, index=False)

    wb = load_workbook(args.out_vue); ws = wb.active
    fill = PatternFill("solid", fgColor="C8E6C9")    # vert clair = ajout étape 2
    for r, f in enumerate(flags, start=2):
        if f:
            for c in range(1, len(cols) + 1):
                ws.cell(row=r, column=c).fill = fill
    wb.save(args.out_vue)
    print(f"[ok] {sum(flags):,} lignes colorées (vert, étape 2) → {args.out_vue}")

    pres = {b: out[f"NOM_{b}"].astype(str).str.strip() != "" for b in BASES}
    nb = sum(pres[b].astype(int) for b in BASES)
    with pd.ExcelWriter(args.out_nonmatch, engine="openpyxl") as w:
        rc = []
        for b in BASES:
            sub = out[pres[b] & (nb == 1)][[f"NOM_{b}", f"ID_{b}", "SECTEUR"]]
            sub.columns = ["NOM", "ID", "SECTEUR"]
            sub.sort_values("NOM").to_excel(w, sheet_name=f"non_match_{b}", index=False)
            rc.append({"BASE": b, "NON_MATCHEES": len(sub)})
            print(f"  {b}: {len(sub):,} non matchées")
        pd.DataFrame(rc).to_excel(w, sheet_name="synthese", index=False)
    print(f"[ok] → {args.out_nonmatch}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
