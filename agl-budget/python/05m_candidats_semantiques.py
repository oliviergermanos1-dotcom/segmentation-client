#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Candidats sémantiques (acronyme / lettres communes) pour revue Claude
==================================================================================
Génère des paires d'orphelins cross-base que le matching par tokens a ratées :
  - ACRONYME ↔ EXPANSION : "SICTA" ↔ "STE IVOIRIENNE ..." (initiales) ;
  - sigle entre parenthèses présent dans l'autre nom ;
  - cosine caractères (trigrammes) pour fautes/réordonnancements ;
  - sous-chaîne significative commune.
Sortie : candidats classés multi-signaux → lus et jugés ensuite par Claude.
"""
import argparse, os, sys, importlib.util, re
from collections import defaultdict
import pandas as pd
from rapidfuzz import fuzz

_spec = importlib.util.spec_from_file_location(
    "norm", os.path.join(os.path.dirname(__file__), "01_normalise.py"))
norm = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(norm)

STOP = {"STE", "SOCIETE", "ETS", "ETABLISSEMENT", "ETABLISSEMENTS", "CIE", "COMPAGNIE",
        "GROUPE", "GROUP", "ENTREPRISE", "DE", "DU", "DES", "ET", "LA", "LE", "LES",
        "SARL", "SA", "SAS", "CI", "IVOIRE", "COTE", "INTERNATIONAL"}


def acro(name):
    toks = [t for t in re.sub(r"[^A-Z0-9 ]", " ", name.upper()).split()
            if t not in STOP and len(t) >= 2]
    return "".join(t[0] for t in toks)


def trigrams(s):
    s = re.sub(r"[^A-Z0-9]", "", s.upper())
    return {s[i:i+3] for i in range(len(s) - 2)} or {s}


def tri_cos(a, b):
    A, B = trigrams(a), trigrams(b)
    return len(A & B) / (len(A | B) or 1)


def load(xl, sheet):
    d = pd.read_excel(xl, sheet).fillna("")
    out = []
    for _, r in d.iterrows():
        nom = str(r["NOM"]).strip()
        if not nom:
            continue
        nn, nb = norm.normalise(nom)
        out.append({"NOM": nom, "ID": str(r["ID"]), "SECTEUR": r.get("SECTEUR", ""),
                    "base": nb or nom, "full": nn or nom, "acro": acro(nom)})
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--non-matches", required=True)
    ap.add_argument("--out", default="candidats_semantiques.xlsx")
    ap.add_argument("--tri-min", type=float, default=0.45)
    args = ap.parse_args(argv)

    xl = pd.ExcelFile(args.non_matches)
    data = {b: load(xl, f"non_match_{b}") for b in ["CRM", "STATCOM", "IRIS", "RUBRIKS"]
            if f"non_match_{b}" in xl.sheet_names}
    for b, v in data.items():
        print(f"{b}: {len(v)} orphelins")

    pairs = [("CRM", "STATCOM"), ("CRM", "IRIS"), ("IRIS", "STATCOM"),
             ("RUBRIKS", "STATCOM"), ("IRIS", "RUBRIKS"), ("CRM", "RUBRIKS")]
    sheets = {}
    for ba, bb in pairs:
        if ba not in data or bb not in data:
            continue
        B = data[bb]
        # index B par acronyme et par préfixe de tokens base
        acroB = defaultdict(list); tokB = defaultdict(list)
        for k, r in enumerate(B):
            if r["acro"]:
                acroB[r["acro"]].append(k)
            for t in r["base"].split():
                if len(t) >= 3:
                    tokB[t[:4]].append(k)
        rows, seen = [], set()
        for ra in data[ba]:
            cand = set(acroB.get(ra["acro"], []))            # acronyme identique
            for t in ra["base"].split():                      # + token commun
                if len(t) >= 3:
                    cand.update(tokB.get(t[:4], []))
            for k in cand:
                rb = B[k]
                key = (ra["NOM"], rb["NOM"])
                if key in seen:
                    continue
                seen.add(key)
                tri = tri_cos(ra["base"], rb["base"])
                tset = fuzz.token_set_ratio(ra["base"], rb["base"]) / 100.0
                # acronyme identique (≥3 lettres) = signal sémantique fort
                acro_eq = len(ra["acro"]) >= 3 and ra["acro"] == rb["acro"]
                # sigle : acronyme de l'un (≥4) = compactage de l'autre
                ca = ra["acro"]; cb = rb["acro"]
                sig = ((len(ca) >= 4 and ca == "".join(t[0] for t in rb["base"].split() if len(t) >= 2))
                       or (len(cb) >= 4 and cb == "".join(t[0] for t in ra["base"].split() if len(t) >= 2)))
                # on ne garde QUE le sémantique fort (le textuel est déjà couvert par 05l)
                if (acro_eq and tri >= 0.25) or sig or tri >= 0.72 or tset >= 0.88:
                    rows.append({
                        f"NOM_{ba}": ra["NOM"], f"ID_{ba}": ra["ID"],
                        f"NOM_{bb}": rb["NOM"], f"ID_{bb}": rb["ID"],
                        "TRI_COS": round(tri, 2), "TOKEN_SET": round(tset, 2),
                        "ACRO_A": ra["acro"], "ACRO_B": rb["acro"],
                        "ACRO_EQ": "✓" if acro_eq else "", "SIGLE": "✓" if sig else "",
                    })
        df = pd.DataFrame(rows)
        if len(df):
            df["_s"] = df[["TRI_COS", "TOKEN_SET"]].max(axis=1) + \
                       df["ACRO_EQ"].eq("✓").astype(int) * 0.3 + df["SIGLE"].eq("✓").astype(int) * 0.3
            df = df.sort_values("_s", ascending=False).drop(columns="_s")
        sheets[f"{ba}_x_{bb}"] = df
        print(f"  {ba} × {bb}: {len(df)} candidats")

    with pd.ExcelWriter(args.out, engine="openpyxl") as w:
        for name, df in sheets.items():
            (df if len(df) else pd.DataFrame({"info": ["aucun"]})).to_excel(
                w, sheet_name=name[:31], index=False)
    print(f"[ok] → {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
