#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Top 200 par base (pour filtrer le référentiel sur les top acteurs)
==============================================================================
- IRIS    : top 200 clients par CAP (Σ MONTANT)
- RUBRIKS : top 200 clients par CAP réel (Σ r2025)
- STATCOM : top 200 clients PAR MÉTIER par poids de marchandise

Chaque ligne porte NOM_BASE (= clé de jointure vers referentiel_vue) + nom
brut + id + la métrique + le rang. Sortie multi-onglets.
"""
import argparse, sys
import pandas as pd

TOPN = 200


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--iris", required=True)
    ap.add_argument("--rubriks", required=True)
    ap.add_argument("--statcom", required=True)
    ap.add_argument("--rubriks-cap-col", default="r2025")
    ap.add_argument("--out", default="tops_par_base.xlsx")
    args = ap.parse_args(argv)

    sheets = {}

    # --- IRIS : CAP = Σ MONTANT par client -------------------------------
    i = pd.read_excel(args.iris, usecols=["CLIENT", "MONTANT", "ID", "NOM_BASE"])
    i["MONTANT"] = pd.to_numeric(i["MONTANT"], errors="coerce").fillna(0)
    gi = (i.groupby("NOM_BASE")
            .agg(NOM_IRIS=("CLIENT", "first"), ID_IRIS=("ID", "first"),
                 CAP_IRIS=("MONTANT", "sum"))
            .reset_index().sort_values("CAP_IRIS", ascending=False).head(TOPN))
    gi.insert(0, "RANG", range(1, len(gi) + 1))
    sheets["top200_iris_cap"] = gi
    print(f"IRIS    : top {len(gi)} (CAP max {gi['CAP_IRIS'].max():,.0f})")

    # --- RUBRIKS : CAP = Σ r2025 par client (depuis le fichier BRUT, car
    #     plusieurs lignes/client à sommer ; clé = ID Concerto) -------------
    import re
    r = pd.read_excel(args.rubriks, sheet_name="Extraction Rubiks")
    cap = args.rubriks_cap_col
    r = r[~r["CLIENT"].astype(str).str.startswith("Total Customer")].copy()
    patid = re.compile(r"\(([A-Z0-9][A-Z0-9\-]+)\)\s*$")
    patnom = re.compile(r"^(.*?)\s*\([A-Z0-9][A-Z0-9\-]+\)\s*$")
    r["ID_RUBRIKS"] = r["CLIENT"].map(lambda s: (patid.search(str(s)).group(1)
                                                 if patid.search(str(s)) else ""))
    r["NOM_RUBRIKS"] = r["CLIENT"].map(lambda s: (patnom.match(str(s)).group(1).strip()
                                                  if patnom.match(str(s)) else str(s).strip()))
    r[cap] = pd.to_numeric(r[cap], errors="coerce").fillna(0)
    r["_key"] = r["ID_RUBRIKS"].where(r["ID_RUBRIKS"] != "", r["NOM_RUBRIKS"])
    gr = (r.groupby("_key")
            .agg(NOM_RUBRIKS=("NOM_RUBRIKS", "first"),
                 ID_RUBRIKS=("ID_RUBRIKS", "first"),
                 CAP_RUBRIKS=(cap, "sum"))
            .reset_index(drop=True).sort_values("CAP_RUBRIKS", ascending=False).head(TOPN))
    gr.insert(0, "RANG", range(1, len(gr) + 1))
    sheets["top200_rubriks_cap"] = gr
    print(f"RUBRIKS : top {len(gr)} (CAP max {gr['CAP_RUBRIKS'].max():,.0f})")

    # --- STATCOM : top 200 par MÉTIER par poids unifié -------------------
    s = pd.read_excel(args.statcom, usecols=["metier", "NOM_BASE", "CLIENT_RESOLU",
                                             "ID_STATCOM", "volume_teu",
                                             "volume_bulk", "volume_kg"])
    s["POIDS"] = (pd.to_numeric(s["volume_kg"], errors="coerce").fillna(0)
                  + pd.to_numeric(s["volume_bulk"], errors="coerce").fillna(0) * 1000
                  + pd.to_numeric(s["volume_teu"], errors="coerce").fillna(0) * 15000)
    abbr = {"Import Maritime": "TIM", "Export Maritime": "TEM",
            "Import Aérien": "TIA", "Export Aérien": "TEA",
            "Hinterland Import": "HINT_IMP", "Hinterland Export": "HINT_EXP"}
    for metier, gm in s.groupby("metier"):
        gt = (gm.groupby("NOM_BASE")
                .agg(NOM_STATCOM=("CLIENT_RESOLU", "first"),
                     ID_STATCOM=("ID_STATCOM", "first"),
                     POIDS=("POIDS", "sum"),
                     TEU=("volume_teu", "sum"), BULK=("volume_bulk", "sum"),
                     KG=("volume_kg", "sum"))
                .reset_index().sort_values("POIDS", ascending=False).head(TOPN))
        gt.insert(0, "RANG", range(1, len(gt) + 1))
        gt.insert(1, "METIER", metier)
        name = f"top200_stat_{abbr.get(metier, metier)[:20]}"
        sheets[name] = gt
        print(f"STATCOM {metier:18s}: top {len(gt)}")

    with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
        for name, df in sheets.items():
            df.to_excel(xl, sheet_name=name[:31], index=False)
    print(f"\n[ok] {len(sheets)} onglets → {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
