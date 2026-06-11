#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Top 300 par base, ventilé 2024/2025/2026 (classement sur 2025)
===========================================================================
- IRIS    : top 300 clients par CAP (Σ MONTANT) — colonnes CAP 2024/25/26,
            rang sur 2025. Comptes internes/inter-agence exclus.
- RUBRIKS : top 300 clients par CAP réel — r2025 (=2025) + B26 (=2026),
            rang sur 2025. (2024 absent du fichier source.)
- STATCOM : top 300 par MÉTIER des clients TRAITÉS PAR AGL (Transitaire =
            Africa Global Logistics) — poids 2024/25/26, rang sur 2025.

Chaque ligne porte NOM_BASE (clé de jointure vers le référentiel).
"""
import argparse, sys, glob, os
import importlib.util
import pandas as pd

TOPN = 300
YEARS = [2024, 2025, 2026]
RANK_YEAR = 2025

_spec = importlib.util.spec_from_file_location(
    "scmod", os.path.join(os.path.dirname(__file__), "_statcom_consolide.py"))
scmod = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(scmod)


def _y4(s):
    s = str(s).strip()
    return int(s[:4]) if s[:4].isdigit() else None


def top_iris(path):
    i = pd.read_excel(path, usecols=["CLIENT", "MONTANT", "ANNEE", "ID", "NOM_BASE"])
    i["MONTANT"] = pd.to_numeric(i["MONTANT"], errors="coerce").fillna(0)
    i["ANNEE"] = pd.to_numeric(i["ANNEE"], errors="coerce")
    INTERNE = ("FACT INTERNE", "INTERNE TRANSIT", "FACTURATION INTERNE", "COMPTANT",
               "INTERCO", "AGL CI FACT", "BOLLORE TL CI FACT", "DIVERS CLIENT",
               "CLIENT DIVERS", "INTRA GROUPE FACT", "INTRAGROUPE")
    i = i[~i["CLIENT"].astype(str).str.upper().str.contains("|".join(INTERNE), na=False, regex=True)]
    i = i[i["ANNEE"].isin(YEARS)]
    piv = (i.pivot_table(index="NOM_BASE", columns="ANNEE", values="MONTANT",
                         aggfunc="sum", fill_value=0))
    for y in YEARS:
        if y not in piv.columns: piv[y] = 0
    piv = piv.rename(columns={y: f"CAP_{y}" for y in YEARS}).reset_index()
    meta = i.groupby("NOM_BASE").agg(NOM_IRIS=("CLIENT", "first"),
                                     ID_IRIS=("ID", "first")).reset_index()
    g = meta.merge(piv, on="NOM_BASE")
    g = g.sort_values(f"CAP_{RANK_YEAR}", ascending=False).head(TOPN)
    g.insert(0, "RANG", range(1, len(g) + 1))
    return g[["RANG", "NOM_IRIS", "ID_IRIS", "NOM_BASE",
              "CAP_2024", "CAP_2025", "CAP_2026"]]


def top_rubriks(path, cap25="r2025", cap26="B26"):
    import re
    r = pd.read_excel(path, sheet_name="Extraction Rubiks")
    r = r[~r["CLIENT"].astype(str).str.startswith("Total Customer")].copy()
    pid = re.compile(r"\(([A-Z0-9][A-Z0-9\-]+)\)\s*$")
    pnom = re.compile(r"^(.*?)\s*\([A-Z0-9][A-Z0-9\-]+\)\s*$")
    r["ID_RUBRIKS"] = r["CLIENT"].map(lambda s: (pid.search(str(s)).group(1) if pid.search(str(s)) else ""))
    r["NOM_RUBRIKS"] = r["CLIENT"].map(lambda s: (pnom.match(str(s)).group(1).strip() if pnom.match(str(s)) else str(s).strip()))
    for c in (cap25, cap26):
        r[c] = pd.to_numeric(r[c], errors="coerce").fillna(0)
    r["_key"] = r["ID_RUBRIKS"].where(r["ID_RUBRIKS"] != "", r["NOM_RUBRIKS"])
    g = (r.groupby("_key").agg(NOM_RUBRIKS=("NOM_RUBRIKS", "first"),
                               ID_RUBRIKS=("ID_RUBRIKS", "first"),
                               CAP_2025=(cap25, "sum"), CAP_2026=(cap26, "sum"))
           .reset_index(drop=True))
    g["CAP_2024"] = ""   # absent du fichier RUBRIKS
    g = g.sort_values("CAP_2025", ascending=False).head(TOPN)
    g.insert(0, "RANG", range(1, len(g) + 1))
    return g[["RANG", "NOM_RUBRIKS", "ID_RUBRIKS", "CAP_2024", "CAP_2025", "CAP_2026"]]


def top_statcom_agl(raw_dir, transitaire, nb_lookup):
    out = {}
    abbr = {"Import Maritime": "TIM", "Export Maritime": "TEM",
            "Import Aérien": "TIA", "Export Aérien": "TEA",
            "Hinterland Import": "HINT_IMP", "Hinterland Export": "HINT_EXP"}
    for path in sorted(glob.glob(os.path.join(raw_dir, "STATCOM*.xlsx"))):
        info = scmod.metier_du_fichier(os.path.basename(path))
        if not info:
            continue
        metier, sens, unite = info
        df = pd.read_excel(path, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]
        c_tr = scmod.resolve_col(df, ["transitaire"])
        if c_tr is None:
            continue
        df = df[df[c_tr].astype(str).str.upper().str.strip() == transitaire]
        if not len(df):
            continue
        c_cli = (scmod.resolve_col(df, ["destinataires", "destinataire", "consignee"])
                 if sens == "import" else
                 scmod.resolve_col(df, ["chargeurs", "chargeur", "expediteur", "shipper"]))
        c_an = scmod.resolve_col(df, ["années escale", "annees escale", "année escale",
                                      "annee escale", "annee", "année", "year"])
        c_poids = scmod.resolve_col(df, ["POIDS_MARCHANDISE", "poids marchandise", "poids"])
        c_teu = scmod.resolve_col(df, ["NOMBRE_TEU", "nombre teu", "teu"])
        c_cond = scmod.resolve_col(df, ["CODE_CONDIT", "code condit", "conditionnement"])
        poids = scmod._to_num_series(df[c_poids]) if c_poids else 0
        teu = scmod._to_num_series(df[c_teu]) if c_teu else 0
        is_air = "AERIEN" in metier.upper() or "AÉRIEN" in metier.upper()
        if is_air:
            w = poids
        else:
            is_bulk = (df[c_cond].astype(str).str.upper().str.strip().isin(scmod.BULK_CODES)
                       if c_cond else False)
            w = (poids * 1000).where(is_bulk, teu * 15000) if c_cond is not None else teu * 15000
        g = pd.DataFrame({"cli": df[c_cli].astype(str).str.upper().str.strip(),
                          "an": df[c_an].map(_y4) if c_an else None, "POIDS": w})
        g = g[(g["cli"] != "") & (g["an"].isin(YEARS))]
        piv = g.pivot_table(index="cli", columns="an", values="POIDS",
                            aggfunc="sum", fill_value=0)
        for y in YEARS:
            if y not in piv.columns: piv[y] = 0
        piv = piv.rename(columns={y: f"POIDS_{y}" for y in YEARS}).reset_index()
        piv["NOM_BASE"] = piv["cli"].map(nb_lookup).fillna("")
        piv = piv.sort_values(f"POIDS_{RANK_YEAR}", ascending=False).head(TOPN)
        piv.insert(0, "RANG", range(1, len(piv) + 1))
        piv.insert(1, "METIER", metier)
        piv = piv.rename(columns={"cli": "NOM_STATCOM"})
        out[f"top300_stat_{abbr.get(metier, metier)[:16]}"] = piv[
            ["RANG", "METIER", "NOM_STATCOM", "NOM_BASE",
             "POIDS_2024", "POIDS_2025", "POIDS_2026"]]
        print(f"STATCOM {metier:18s}: top {len(piv)} clients AGL")
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--iris", required=True)
    ap.add_argument("--rubriks", required=True)
    ap.add_argument("--statcom-raw-dir", required=True)
    ap.add_argument("--statcom-norm", required=True)
    ap.add_argument("--referentiel", required=True,
                    help="referentiel_vue.xlsx (pour l'ID compte consolidé)")
    ap.add_argument("--transitaire", default="AFRICA GLOBAL LOGISTICS")
    ap.add_argument("--out", default="tops_par_base.xlsx")
    args = ap.parse_args(argv)

    sheets = {}
    gi = top_iris(args.iris); sheets["top300_iris_cap"] = gi
    print(f"IRIS    : top {len(gi)} (CAP 2025 max {gi['CAP_2025'].max():,.0f})")
    gr = top_rubriks(args.rubriks); sheets["top300_rubriks_cap"] = gr
    print(f"RUBRIKS : top {len(gr)} (CAP 2025 max {gr['CAP_2025'].max():,.0f})")

    snorm = pd.read_excel(args.statcom_norm, usecols=["CLIENT_RESOLU", "NOM_BASE"])
    nb_lookup = {str(k).upper().strip(): v for k, v in
                 zip(snorm["CLIENT_RESOLU"], snorm["NOM_BASE"]) if pd.notna(k)}
    sheets.update(top_statcom_agl(args.statcom_raw_dir, args.transitaire.upper(), nb_lookup))

    # --- ID compte consolidé (depuis le référentiel) dans chaque onglet ----
    vue = pd.read_excel(args.referentiel)
    def _m(col):  # col id par base → ID_RETENU
        v = vue[vue[col].astype(str) != ""]
        return dict(zip(v[col].astype(str), v["ID_RETENU"].astype(str)))
    m_iris = _m("ID_IRIS"); m_rub = _m("ID_RUBRIKS"); m_stat = _m("ID_STATCOM")
    m_nom = dict(zip(vue["NOM_CONSOLIDE"].astype(str), vue["ID_RETENU"].astype(str)))

    for name, df in sheets.items():
        if name.startswith("top300_iris"):
            df["ID_COMPTE"] = df["ID_IRIS"].astype(str).map(m_iris).fillna("")
        elif name.startswith("top300_rubriks"):
            df["ID_COMPTE"] = df["ID_RUBRIKS"].astype(str).map(m_rub).fillna("")
        else:  # STATCOM : on retient l'ID_STATCOM (= NOM_BASE canonique)
            df["ID_COMPTE"] = df["NOM_BASE"].astype(str)

    with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
        for name, df in sheets.items():
            df.to_excel(xl, sheet_name=name[:31], index=False)
    print(f"\n[ok] {len(sheets)} onglets (+ ID_COMPTE consolidé) → {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
