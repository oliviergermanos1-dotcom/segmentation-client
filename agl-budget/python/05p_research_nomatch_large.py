#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Étape 2 : re-recherche ÉLARGIE sur les non-matchés 100%
====================================================================
Sur les seules entités mono-base (onglet NON_MATCH_100), on relance le
matching cross-base en LEVANT les restrictions :
  - PAS d'exigence de token rare ;
  - seuil abaissé (token_sort ≥ 0.65 OU trigrammes ≥ 0.55) ;
  - PAS de confirmation nom complet.
SEUL garde-fou conservé : le PAYS (cross-pays = quasi toujours faux).

Aucune fusion automatique : sortie en candidats classés PROBABLE / FAIBLE
pour validation manuelle.
"""
import argparse, os, sys, importlib.util, re
from collections import defaultdict
import pandas as pd
from rapidfuzz import fuzz

_spec = importlib.util.spec_from_file_location(
    "norm", os.path.join(os.path.dirname(__file__), "01_normalise.py"))
norm = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(norm)

BASES = ["CRM", "STATCOM", "IRIS", "RUBRIKS"]
PAYS_KW = {"CI": ["COTE IVOIRE", "COTE D IVOIRE", "IVOIRE", "RCI", " CIV", "ABIDJAN"],
           "BF": [" BF", "BURKINA", "OUAGA"], "ML": [" MALI", "BAMAKO"],
           "SN": ["SENEGAL", "DAKAR"], "GH": ["GHANA"], "TG": [" TOGO", "LOME"],
           "BJ": ["BENIN", "COTONOU"], "NE": [" NIGER ", "NIAMEY"], "GN": ["GUINEE", "CONAKRY"]}


def pays(full):
    s = f" {str(full).upper()} "
    for c, ks in PAYS_KW.items():
        if any(k in s for k in ks):
            return c
    return ""


def trigrams(s):
    s = re.sub(r"[^A-Z0-9]", "", str(s).upper())
    return {s[i:i+3] for i in range(len(s) - 2)} or {s}


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--split", required=True, help="ref_vue_split.xlsx")
    ap.add_argument("--sheet", default="NON_MATCH_100")
    ap.add_argument("--ts-min", type=float, default=0.65)
    ap.add_argument("--tri-min", type=float, default=0.55)
    ap.add_argument("--out", default="research_nomatch_large.xlsx")
    args = ap.parse_args(argv)

    v = pd.read_excel(args.split, sheet_name=args.sheet).fillna("")
    data = defaultdict(list)
    for _, r in v.iterrows():
        for b in BASES:
            nom = str(r.get(f"NOM_{b}", "")).strip()
            if nom:
                nn, nb = norm.normalise(nom)
                data[b].append({"NOM": nom, "ID": str(r.get(f"ID_{b}", "")),
                                "base": nb or nom, "full": nn or nom, "pays": pays(nn or nom)})
                break
    for b in BASES:
        print(f"{b}: {len(data[b])} non-matchés")

    pairs = [("CRM", "STATCOM"), ("CRM", "IRIS"), ("IRIS", "STATCOM"),
             ("RUBRIKS", "STATCOM"), ("IRIS", "RUBRIKS"), ("CRM", "RUBRIKS")]
    sheets, recap = {}, []
    for ba, bb in pairs:
        if not data[ba] or not data[bb]:
            continue
        idxB = defaultdict(list)
        for k, r in enumerate(data[bb]):
            for t in r["base"].split():
                if len(t) >= 3:
                    idxB[t[:4]].append(k)
        rows, seen = [], set()
        for ra in data[ba]:
            cand = set()
            for t in ra["base"].split():
                if len(t) >= 3:
                    cand.update(idxB.get(t[:4], []))
            for k in cand:
                rb = data[bb][k]
                if (ra["NOM"], rb["NOM"]) in seen:
                    continue
                seen.add((ra["NOM"], rb["NOM"]))
                if ra["pays"] and rb["pays"] and ra["pays"] != rb["pays"]:
                    continue                      # SEUL garde-fou conservé
                ts = fuzz.token_sort_ratio(ra["base"], rb["base"]) / 100.0
                tset = fuzz.token_set_ratio(ra["base"], rb["base"]) / 100.0
                A, B = trigrams(ra["base"]), trigrams(rb["base"])
                tri = len(A & B) / (len(A | B) or 1)
                if ts < args.ts_min and tri < args.tri_min and tset < 0.85:
                    continue
                score = max(ts, tri, tset * 0.9)
                niveau = ("SUR" if (ts >= 0.92 or tri >= 0.80)
                          else "PROBABLE" if (ts >= 0.85 or tri >= 0.70)
                          else "FAIBLE")
                rows.append({
                    f"NOM_{ba}": ra["NOM"], f"ID_{ba}": ra["ID"],
                    f"NOM_{bb}": rb["NOM"], f"ID_{bb}": rb["ID"],
                    "TOKEN_SORT": round(ts, 2), "TOKEN_SET": round(tset, 2),
                    "TRIGRAMME": round(tri, 2), "SCORE": round(score, 2),
                    "NIVEAU": niveau,
                    "PAYS": ra["pays"] or rb["pays"] or "—",
                })
        df = pd.DataFrame(rows)
        if len(df):
            df = df.sort_values("SCORE", ascending=False)
        sheets[f"{ba}_x_{bb}"] = df
        nsur = int((df["NIVEAU"] == "SUR").sum()) if len(df) else 0
        nprob = int((df["NIVEAU"] == "PROBABLE").sum()) if len(df) else 0
        recap.append({"PAIRE": f"{ba} × {bb}", "CANDIDATS": len(df),
                      "SUR": nsur, "PROBABLE": nprob})
        print(f"  {ba} × {bb}: {len(df)} candidats ({nsur} sûrs, {nprob} probables)")

    with pd.ExcelWriter(args.out, engine="openpyxl") as w:
        pd.DataFrame(recap).to_excel(w, sheet_name="synthese", index=False)
        for name, df in sheets.items():
            (df if len(df) else pd.DataFrame({"info": ["aucun"]})).to_excel(
                w, sheet_name=name[:31], index=False)
    print(f"[ok] → {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
