#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Entreprises NON matchées, par base (depuis referentiel_vue)
========================================================================
Pour chaque base, liste les entités présentes UNIQUEMENT dans cette base
(aucun lien avec une autre) = orphelines à investiguer.
1 onglet par base + un onglet de synthèse.
"""
import argparse, sys
import pandas as pd

BASES = ["CRM", "STATCOM", "IRIS", "RUBRIKS"]


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--vue", required=True, help="referentiel_vue.xlsx")
    ap.add_argument("--out", default="non_matches_par_base.xlsx")
    args = ap.parse_args(argv)

    v = pd.read_excel(args.vue).fillna("")
    # présence par base = nom non vide
    pres = {b: v[f"NOM_{b}"].astype(str).str.strip() != "" for b in BASES}
    v["_nbases"] = sum(pres[b].astype(int) for b in BASES)

    sheets, recap = {}, []
    for b in BASES:
        # orphelin de b = présent dans b ET dans aucune autre
        mask = pres[b] & (v["_nbases"] == 1)
        sub = v[mask][[f"NOM_{b}", f"ID_{b}", "SECTEUR", "ID_ENTITE"]].copy()
        sub = sub.rename(columns={f"NOM_{b}": "NOM", f"ID_{b}": "ID"})
        sub = sub.sort_values("NOM")
        sheets[f"non_match_{b}"] = sub
        recap.append({"BASE": b, "PRESENTES": int(pres[b].sum()),
                      "NON_MATCHEES": len(sub),
                      "%_NON_MATCH": f"{len(sub)/max(int(pres[b].sum()),1)*100:.1f}%"})
        print(f"{b:8s}: {len(sub):,} non matchées / {int(pres[b].sum()):,} présentes")

    with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
        pd.DataFrame(recap).to_excel(xl, sheet_name="synthese", index=False)
        for name, df in sheets.items():
            df.to_excel(xl, sheet_name=name[:31], index=False)
    print(f"[ok] → {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
