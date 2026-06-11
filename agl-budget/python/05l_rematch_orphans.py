#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Recherche approfondie : rattrapage des orphelins entre bases
=========================================================================
Repère les entreprises présentes dans les listes NON MATCHÉES de 2 bases
différentes qui sont en réalité la MÊME entité (matchs ratés en 1ère passe).

Pour chaque paire de bases (CRM/STATCOM/IRIS/RUBRIKS) :
  - normalise les noms orphelins (tronc + nom complet + pays) ;
  - blocking élargi sur préfixes de TOUS les tokens (≥3 lettres) ;
  - score = token_sort_ratio nom complet + token_set tronc ;
  - garde-fou pays ; exige ≥2 tokens communs OU token_set ≥ 0.95 ;
  - 2 niveaux : SUR (≥0.95) / PROBABLE (seuil..0.95) pour revue.

Sortie : orphelins_rattrapes.xlsx — 1 onglet par paire de bases + synthèse.
"""
import argparse, os, sys, importlib.util
from collections import defaultdict
import pandas as pd
from rapidfuzz import fuzz, distance

_spec = importlib.util.spec_from_file_location(
    "norm", os.path.join(os.path.dirname(__file__), "01_normalise.py"))
norm = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(norm)

PAYS_KW = {"CI": ["COTE IVOIRE", "COTE D IVOIRE", "IVOIRE", "RCI", " CIV", "ABIDJAN"],
           "BF": [" BF", "BURKINA", "OUAGA"], "ML": [" MALI", "BAMAKO"],
           "SN": ["SENEGAL", "DAKAR"], "GH": ["GHANA"], "TG": [" TOGO", "LOME"],
           "BJ": ["BENIN", "COTONOU"], "NE": [" NIGER ", "NIAMEY"], "GN": ["GUINEE", "CONAKRY"]}


def pays(full):
    s = f" {str(full).upper()} "
    for code, kws in PAYS_KW.items():
        if any(k in s for k in kws):
            return code
    return ""


def toks(s):
    return [t for t in str(s).split() if len(t) >= 3]


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--non-matches", required=True, help="non_matches_par_base.xlsx")
    ap.add_argument("--threshold", type=float, default=0.88)
    ap.add_argument("--out", default="orphelins_rattrapes.xlsx")
    args = ap.parse_args(argv)

    xl = pd.ExcelFile(args.non_matches)
    bases = {}
    for b in ["CRM", "STATCOM", "IRIS", "RUBRIKS"]:
        sh = f"non_match_{b}"
        if sh not in xl.sheet_names:
            continue
        d = pd.read_excel(xl, sh).fillna("")
        recs = []
        for _, r in d.iterrows():
            nom = str(r["NOM"]).strip()
            if not nom:
                continue
            nn, nb = norm.normalise(nom)
            recs.append({"NOM": nom, "ID": str(r["ID"]), "SECTEUR": r.get("SECTEUR", ""),
                         "tronc": nb, "full": nn or nom, "pays": pays(nn or nom)})
        bases[b] = recs
        print(f"{b:8s}: {len(recs):,} orphelins")

    pairs = [("CRM", "STATCOM"), ("CRM", "IRIS"), ("CRM", "RUBRIKS"),
             ("IRIS", "STATCOM"), ("RUBRIKS", "STATCOM"), ("IRIS", "RUBRIKS")]
    sheets, recap = {}, []
    for ba, bb in pairs:
        if ba not in bases or bb not in bases:
            continue
        # index B par préfixe de tous les tokens
        idxB = defaultdict(list)
        for k, r in enumerate(bases[bb]):
            for t in toks(r["tronc"]):
                idxB[t[:4]].append(k)
        found, seen = [], set()
        for ra in bases[ba]:
            cand = set()
            for t in toks(ra["tronc"]):
                cand.update(idxB.get(t[:4], []))
            for k in cand:
                rb = bases[bb][k]
                if (ra["NOM"], rb["NOM"]) in seen:
                    continue
                seen.add((ra["NOM"], rb["NOM"]))
                if ra["pays"] and rb["pays"] and ra["pays"] != rb["pays"]:
                    continue
                tfull = fuzz.token_sort_ratio(ra["full"], rb["full"]) / 100.0
                if tfull < args.threshold:
                    continue
                shared = len(set(toks(ra["tronc"])) & set(toks(rb["tronc"])))
                tset = fuzz.token_set_ratio(ra["tronc"], rb["tronc"]) / 100.0
                if shared < 2 and tset < 0.95:
                    continue
                found.append({
                    f"NOM_{ba}": ra["NOM"], f"ID_{ba}": ra["ID"],
                    f"NOM_{bb}": rb["NOM"], f"ID_{bb}": rb["ID"],
                    "SCORE": round(tfull, 3),
                    "NIVEAU": "SUR" if tfull >= 0.95 else "PROBABLE",
                    "PAYS": ra["pays"] or rb["pays"] or "—",
                })
        df = pd.DataFrame(found).sort_values("SCORE", ascending=False) if found else pd.DataFrame()
        sheets[f"{ba}_x_{bb}"] = df
        recap.append({"PAIRE": f"{ba} × {bb}", "RATTRAPES": len(df),
                      "SUR": int((df["NIVEAU"] == "SUR").sum()) if len(df) else 0})
        print(f"  {ba} × {bb:8s}: {len(df):,} rattrapés "
              f"({int((df['NIVEAU']=='SUR').sum()) if len(df) else 0} sûrs)")

    with pd.ExcelWriter(args.out, engine="openpyxl") as w:
        pd.DataFrame(recap).to_excel(w, sheet_name="synthese", index=False)
        for name, df in sheets.items():
            (df if len(df) else pd.DataFrame({"info": ["aucun"]})).to_excel(
                w, sheet_name=name[:31], index=False)
    print(f"\n[ok] → {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
