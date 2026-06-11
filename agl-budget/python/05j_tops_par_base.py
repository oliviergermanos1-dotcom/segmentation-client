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
import argparse, sys, glob, os
import importlib.util
import pandas as pd

TOPN = 200

# Helpers de _statcom_consolide (métier, résolution colonnes, règles poids)
_spec = importlib.util.spec_from_file_location(
    "scmod", os.path.join(os.path.dirname(__file__), "_statcom_consolide.py"))
scmod = importlib.util.module_from_spec(_spec); _spec.loader.exec_module(scmod)


def statcom_tops_agl(raw_dir, transitaire, nb_lookup):
    """Top 200 par métier des clients dont le Transitaire = AGL."""
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
            print(f"  [skip] {os.path.basename(path)} : pas de colonne Transitaire")
            continue
        df = df[df[c_tr].astype(str).str.upper().str.strip() == transitaire]
        if not len(df):
            continue
        c_cli = (scmod.resolve_col(df, ["destinataires", "destinataire", "consignee"])
                 if sens == "import" else
                 scmod.resolve_col(df, ["chargeurs", "chargeur", "expediteur", "shipper"]))
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
        g = pd.DataFrame({"cli": df[c_cli].astype(str).str.upper().str.strip(), "POIDS": w})
        g = g[g["cli"] != ""]
        gt = (g.groupby("cli")["POIDS"].sum().reset_index()
                .sort_values("POIDS", ascending=False).head(TOPN))
        gt["NOM_BASE"] = gt["cli"].map(nb_lookup).fillna("")
        gt.insert(0, "RANG", range(1, len(gt) + 1))
        gt.insert(1, "METIER", metier)
        gt = gt.rename(columns={"cli": "NOM_STATCOM"})
        out[f"top200_stat_{abbr.get(metier, metier)[:18]}"] = gt[
            ["RANG", "METIER", "NOM_STATCOM", "NOM_BASE", "POIDS"]]
        print(f"STATCOM {metier:18s}: top {len(gt)} clients AGL")
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--iris", required=True)
    ap.add_argument("--rubriks", required=True)
    ap.add_argument("--statcom-raw-dir", required=True,
                    help="dossier des STATCOM bruts (avec colonne Transitaire)")
    ap.add_argument("--statcom-norm", required=True,
                    help="STATCOM_norm.xlsx (lookup CLIENT_RESOLU→NOM_BASE)")
    ap.add_argument("--transitaire", default="AFRICA GLOBAL LOGISTICS")
    ap.add_argument("--rubriks-cap-col", default="r2025")
    ap.add_argument("--out", default="tops_par_base.xlsx")
    args = ap.parse_args(argv)

    sheets = {}

    # --- IRIS : CAP = Σ MONTANT par client (hors comptes internes) -------
    i = pd.read_excel(args.iris, usecols=["CLIENT", "MONTANT", "ID", "NOM_BASE"])
    i["MONTANT"] = pd.to_numeric(i["MONTANT"], errors="coerce").fillna(0)
    # Exclure facturation inter-agence / interne / comptant (pas de vrais clients)
    INTERNE = ("FACT INTERNE", "INTERNE TRANSIT", "FACTURATION INTERNE", "COMPTANT",
               "INTERCO", "AGL CI FACT", "BOLLORE TL CI FACT", "DIVERS CLIENT",
               "CLIENT DIVERS", "FACT INTRAGROUPE", "INTRA GROUPE FACT")
    pat_int = "|".join(INTERNE)
    masque_int = i["CLIENT"].astype(str).str.upper().str.contains(pat_int, na=False, regex=True)
    n_int = i.loc[masque_int, "NOM_BASE"].nunique()
    i = i[~masque_int]
    print(f"IRIS    : {n_int} comptes internes/comptant exclus")
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

    # --- STATCOM : top 200 par MÉTIER des clients TRAITÉS PAR AGL --------
    snorm = pd.read_excel(args.statcom_norm, usecols=["CLIENT_RESOLU", "NOM_BASE"])
    nb_lookup = {str(k).upper().strip(): v for k, v in
                 zip(snorm["CLIENT_RESOLU"], snorm["NOM_BASE"]) if pd.notna(k)}
    sheets.update(statcom_tops_agl(args.statcom_raw_dir, args.transitaire.upper(), nb_lookup))

    with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
        for name, df in sheets.items():
            df.to_excel(xl, sheet_name=name[:31], index=False)
    print(f"\n[ok] {len(sheets)} onglets → {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
