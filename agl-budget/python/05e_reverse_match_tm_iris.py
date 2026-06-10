#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Reverse-match white spaces (TM_xxx) → IRIS
========================================================
Les white spaces STATCOM sont marqués TM_xxx car absents du CRM. MAIS
certains sont de VRAIS clients AGL, simplement facturés dans IRIS sous un
nom différent (et jamais créés proprement dans le CRM). Le pipeline actuel
ne confronte jamais directement TM → IRIS : on répare ça.

Pour chaque TM_xxx :
  - on cherche son meilleur match dans les noms IRIS (NOM_BASE) via la
    même logique sûre que matching_lib (token_set_ratio + TOKEN_FLOOR) ;
  - si match >= seuil AUTO → le client n'est PAS un vrai white space :
    il a une facturation IRIS. On le promeut "RATTACHE_IRIS" avec l'ID IRIS.

Sortie : tm_iris_bridge.csv (TM_xxx ↔ ID_IRIS retrouvés) + stats.
Ces ponts sont à réinjecter dans clients_master (statut + ID_IRIS).
"""
from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd
from rapidfuzz import fuzz, distance

TOKEN_FLOOR = 0.80
AUTO = 0.92
AMBIGU = 0.75


def _blocks(name: str, prefix: int = 4):
    toks = [t for t in str(name).upper().split() if len(t) >= 3][:4]
    return {t[:prefix] for t in toks}


def _score(a: str, b: str) -> float:
    tsr = fuzz.token_set_ratio(a, b) / 100.0
    if tsr < TOKEN_FLOOR:
        return min(tsr, AMBIGU - 0.01)
    jw = distance.JaroWinkler.normalized_similarity(a, b)
    return 0.85 * tsr + 0.15 * jw


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--master", required=True, help="clients_master.xlsx (avec TM_xxx)")
    ap.add_argument("--iris-norm", required=True, help="IRIS_norm.xlsx (NOM_BASE + ID)")
    ap.add_argument("--iris-id-col", default="ID", help="colonne ID client IRIS")
    ap.add_argument("--out", default="tm_iris_bridge.csv")
    args = ap.parse_args(argv)

    master = pd.read_excel(args.master, sheet_name="clients_master")
    tm = master[master["ID_UNIQUE"].astype(str).str.startswith("TM")].copy()
    print(f"White spaces TM_xxx : {len(tm):,}")

    iris = pd.read_excel(args.iris_norm)
    iris = iris.dropna(subset=["NOM_BASE"])
    iris = iris[iris["NOM_BASE"].astype(str).str.strip() != ""]
    # 1 ligne par client IRIS (NOM_BASE)
    iris_u = iris.drop_duplicates("NOM_BASE")[["NOM_BASE"] +
             ([args.iris_id_col] if args.iris_id_col in iris.columns else [])]
    print(f"Clients IRIS uniques : {len(iris_u):,}")

    # Index blocking sur IRIS
    idx = defaultdict(list)
    for _, r in iris_u.iterrows():
        for b in _blocks(r["NOM_BASE"]):
            idx[b].append(r)

    bridges = []
    for _, t in tm.iterrows():
        tname = t.get("ALIAS_STATCOM") or t.get("NOM_CLIENT")
        if not tname:
            continue
        cand_rows, seen = [], set()
        for b in _blocks(tname):
            for r in idx.get(b, []):
                key = r["NOM_BASE"]
                if key not in seen:
                    seen.add(key); cand_rows.append(r)
        best, best_score = None, 0.0
        for r in cand_rows:
            sc = _score(str(tname).upper(), str(r["NOM_BASE"]).upper())
            if sc > best_score:
                best, best_score = r, sc
        if best is not None and best_score >= AMBIGU:
            bridges.append({
                "ID_UNIQUE": t["ID_UNIQUE"],
                "TM_NAME": tname,
                "IRIS_NOM_BASE": best["NOM_BASE"],
                "IRIS_ID": best.get(args.iris_id_col, ""),
                "SCORE": round(best_score, 3),
                "STATUT": "AUTO" if best_score >= AUTO else "AMBIGU",
            })

    br = pd.DataFrame(bridges)
    br.to_csv(args.out, index=False)
    n_auto = (br["STATUT"] == "AUTO").sum() if len(br) else 0
    n_amb = (br["STATUT"] == "AMBIGU").sum() if len(br) else 0
    print(f"\n[ok] {len(br):,} ponts TM→IRIS trouvés "
          f"(AUTO={n_auto:,} · AMBIGU={n_amb:,}) → {args.out}")
    print(f"     => {n_auto:,} 'faux white spaces' : ce sont en réalité des "
          f"clients déjà facturés dans IRIS.")
    if len(br):
        print("\nTop 15 ponts AUTO :")
        for _, r in br[br["STATUT"] == "AUTO"].sort_values(
                "SCORE", ascending=False).head(15).iterrows():
            print(f"  [{r['SCORE']:.2f}] {str(r['TM_NAME'])[:34]:34s} ↔ {r['IRIS_NOM_BASE']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
