#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — 14 : Vue 360 Client (consolidation 3 bases × tous indicateurs)
===========================================================================
Produit UN seul fichier Excel multi-onglets qui réunit pour CHAQUE client RMC :
  - Identité (RMC) : nom, secteur, IDs CRM/IRIS/STATCOM, secteur LOCKED
  - CAP réel (IRIS) par année
  - CAP budgété (RUBRIKS B26)
  - Écart PFA vs Réel
  - Volumes STATCOM (TEU, Bulk, Kg) par année + total 3 ans
  - Métier principal (par volume max)
  - Position BCG ★ STAR / VACHE / DILEMME / POIDS MORT
  - Statut White Space
  - Pipeline CRM (opportunités ouvertes, CAP pondéré) si opportunites.xlsx chargé

Sortie : 5 onglets dans <outdir>/vue_360_client.xlsx
  - 360_client_annee   : 1 ligne par client × année (granulaire)
  - 360_client_total   : 1 ligne par client (synthèse 3 ans)
  - top_clients_cap    : top 100 clients par CAP réel cumulé
  - top_white_spaces   : top 100 prospects par volume
  - kpi_global         : récap exécutif

Usage :
    python 14_vue_360_client.py --out C:\\Users\\...\\AGL-out
    python 14_vue_360_client.py --out ./out --opportunites ./data/opportunites.xlsx
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

import pandas as pd


def _load(p):
    p = Path(p)
    if not p.exists(): return None
    if p.suffix.lower() in (".xlsx", ".xls"): return pd.read_excel(p, dtype=str)
    return pd.read_csv(p, dtype=str)


def _num(df, col):
    if df is None or col not in df.columns:
        return pd.Series(dtype="float64")
    return pd.to_numeric(df[col].astype(str).str.replace(r"[ ,]", "", regex=True),
                         errors="coerce").fillna(0)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AGL BUDGET — 14 vue 360 client")
    ap.add_argument("--out", required=True, help="Dossier où se trouvent RMC.xlsx, agreg/*.csv (sortie pipeline)")
    ap.add_argument("--opportunites", default=None, help="opportunites.xlsx (CRM Pipeline, optionnel)")
    ap.add_argument("--output", default=None, help="Chemin du .xlsx de sortie (défaut: <out>/vue_360_client.xlsx)")
    args = ap.parse_args(argv)

    out = Path(args.out)
    agreg = out / "agreg"
    target = Path(args.output) if args.output else out / "vue_360_client.xlsx"

    # --- Chargement des sources ---------------------------------------------
    rmc = _load(out / "RMC.xlsx")
    if rmc is None: sys.exit(f"ERREUR : RMC.xlsx introuvable dans {out}")
    vc  = _load(agreg / "volume_client.csv")   # STATCOM par client × métier × année
    pdm_cli = _load(agreg / "pdm_clients.csv") # clients AGL volumes
    budget  = _load(agreg / "budget_vs_real.csv")
    whites  = _load(agreg / "whitespaces.csv")
    pdm     = _load(agreg / "pdm.csv")
    opp     = _load(Path(args.opportunites)) if args.opportunites else None

    n = len(rmc)
    print(f"[in] RMC : {n:,} clients · "
          f"STATCOM volumes: {len(vc) if vc is not None else 0:,} · "
          f"BUDGET : {len(budget) if budget is not None else 0:,} · "
          f"WHITE SPACES : {len(whites) if whites is not None else 0:,} · "
          f"PIPELINE : {len(opp) if opp is not None else 0:,}")

    # Normalisation IDs string pour les merges.
    for col in ("ID_CRM", "ID_IRIS", "ID_STATCOM"):
        if col in rmc: rmc[col] = rmc[col].fillna("").astype(str)

    # --- Volumes STATCOM par client × année (sommés sur métiers) ------------
    if vc is not None and len(vc):
        for c in ("VOLUME_TEU", "VOLUME_BULK", "VOLUME_KG", "NB_OPERATIONS"):
            if c in vc: vc[c] = _num(vc, c)
        vc["ANNEE"] = pd.to_numeric(vc.get("ANNEE"), errors="coerce").astype("Int64")
        # Cumul par client × année (toutes métiers confondus)
        vc_year = (vc.groupby(["NOM_BASE", "ANNEE"], dropna=False)
                   .agg(VOLUME_TEU=("VOLUME_TEU", "sum"),
                        VOLUME_BULK=("VOLUME_BULK", "sum"),
                        VOLUME_KG=("VOLUME_KG", "sum"),
                        NB_OPS=("NB_OPERATIONS", "sum"),
                        ID_STATCOM=("ID_STATCOM", "first") if "ID_STATCOM" in vc else ("NOM_BASE", "first"))
                   .reset_index())
        # Métier dominant par client (volume principal max)
        vc["_vol_principal"] = vc["VOLUME_TEU"].where(
            ~vc["METIER"].fillna("").str.upper().str.contains("AERIEN|AÉRIEN", regex=True),
            vc["VOLUME_KG"])
        metier_dom = (vc.groupby(["NOM_BASE", "METIER"])["_vol_principal"].sum()
                      .reset_index().sort_values("_vol_principal", ascending=False)
                      .drop_duplicates("NOM_BASE", keep="first")[["NOM_BASE", "METIER"]]
                      .rename(columns={"METIER": "METIER_DOMINANT"}))
    else:
        vc_year = pd.DataFrame(columns=["NOM_BASE", "ANNEE", "VOLUME_TEU", "VOLUME_BULK", "VOLUME_KG", "NB_OPS"])
        metier_dom = pd.DataFrame(columns=["NOM_BASE", "METIER_DOMINANT"])

    # --- Budget par client (RUBRIKS B26 cumulé) -----------------------------
    if budget is not None and len(budget):
        for c in ("CAP_PFA", "CAP_REEL", "ECART", "ECART_PCT"):
            if c in budget: budget[c] = _num(budget, c)
        budget["ANNEE"] = pd.to_numeric(budget.get("ANNEE"), errors="coerce").astype("Int64")
        # Cumul par client × année (RUBRIKS peut avoir plusieurs lignes par client)
        bud_year = (budget.groupby(["ID_CRM", "ANNEE"], dropna=False)
                    .agg(CAP_PFA=("CAP_PFA", "sum"),
                         CAP_REEL=("CAP_REEL", "sum"))
                    .reset_index())
        bud_year["ECART"] = bud_year["CAP_REEL"] - bud_year["CAP_PFA"]
        bud_year["ECART_PCT"] = bud_year.apply(
            lambda r: (r["ECART"] / r["CAP_PFA"]) if r["CAP_PFA"] > 0 else None, axis=1)
        bud_year["ID_CRM"] = bud_year["ID_CRM"].fillna("").astype(str)
    else:
        bud_year = pd.DataFrame(columns=["ID_CRM", "ANNEE", "CAP_PFA", "CAP_REEL", "ECART", "ECART_PCT"])

    # --- White spaces (statut prospect) -------------------------------------
    ws_set = set()
    if whites is not None and "NOM_BASE" in whites:
        ws_set = set(whites["NOM_BASE"].dropna().astype(str))

    # --- Pipeline CRM (opportunités ouvertes) -------------------------------
    pipe_agg = pd.DataFrame(columns=["ID_CRM", "NB_OPP_OUVERTES", "PIPELINE_CAP_POND", "PIPELINE_CAP_BRUT"])
    if opp is not None and len(opp):
        col_id   = next((c for c in opp.columns if "id" in c.lower() and "compte" in c.lower()), None)
        col_cap  = next((c for c in opp.columns if "cap" in c.lower() and ("potentiel" in c.lower() or "€" in c)), None)
        col_proba= next((c for c in opp.columns if "probabilit" in c.lower()), None)
        col_stat = next((c for c in opp.columns if c.lower().strip() == "statut"), None)
        if col_id and col_cap:
            opp["_cap"] = _num(opp, col_cap)
            if col_proba:
                opp["_proba"] = opp[col_proba].astype(str).str.replace("%", "").str.replace(",", ".").astype(float, errors="ignore")
                opp["_proba"] = pd.to_numeric(opp["_proba"], errors="coerce").fillna(0)
                opp["_proba"] = opp["_proba"].where(opp["_proba"] <= 1, opp["_proba"] / 100)
            else:
                opp["_proba"] = 0
            mask_open = opp[col_stat].astype(str).str.contains("ouvert|Ouvert", na=False) if col_stat else True
            opp_open = opp[mask_open] if isinstance(mask_open, pd.Series) else opp
            pipe_agg = (opp_open.groupby(col_id, dropna=False)
                        .agg(NB_OPP_OUVERTES=(col_id, "count"),
                             PIPELINE_CAP_BRUT=("_cap", "sum"),
                             PIPELINE_CAP_POND=("_cap", lambda s: (s * opp_open.loc[s.index, "_proba"]).sum()))
                        .reset_index()
                        .rename(columns={col_id: "ID_CRM"}))
            pipe_agg["ID_CRM"] = pipe_agg["ID_CRM"].astype(str)

    # --- ASSEMBLAGE VUE 360 client × année ---------------------------------
    # Annés disponibles : intersection IRIS volumes + RUBRIKS
    years = sorted({y for y in vc_year["ANNEE"].dropna().unique()} |
                   {y for y in bud_year["ANNEE"].dropna().unique()})
    if not years:
        sys.exit("ERREUR : aucune année commune entre volumes et budget.")

    # Pour chaque client RMC, on construit 1 ligne par année.
    rows = []
    rmc_indexed = rmc.set_index("ID_CRM", drop=False)
    vc_year_indexed = vc_year.set_index(["NOM_BASE", "ANNEE"]) if len(vc_year) else None
    bud_year_indexed = bud_year.set_index(["ID_CRM", "ANNEE"]) if len(bud_year) else None
    metier_dom_indexed = metier_dom.set_index("NOM_BASE") if len(metier_dom) else None

    for _, r in rmc.iterrows():
        id_crm = str(r["ID_CRM"])
        nom = r.get("NOM_CANONIQUE", "")
        secteur = r.get("SECTEUR", "")
        alias_stat = str(r.get("ALIAS_STATCOM", "") or "")
        is_agl = bool(str(r.get("ID_IRIS", "") or "").strip())
        for y in years:
            row = {
                "ID_CRM": id_crm,
                "ID_IRIS": r.get("ID_IRIS", ""),
                "ID_STATCOM": r.get("ID_STATCOM", ""),
                "NOM_CLIENT": nom,
                "SECTEUR": secteur,
                "LOCKED": r.get("LOCKED", ""),
                "ANNEE": int(y) if pd.notna(y) else None,
                "EST_AGL_CLIENT": is_agl,
            }
            # Volumes STATCOM (via ALIAS_STATCOM qui est le NOM_BASE)
            if vc_year_indexed is not None and (alias_stat, y) in vc_year_indexed.index:
                vrow = vc_year_indexed.loc[(alias_stat, y)]
                if isinstance(vrow, pd.DataFrame): vrow = vrow.iloc[0]
                row["VOLUME_TEU"]  = float(vrow.get("VOLUME_TEU", 0) or 0)
                row["VOLUME_BULK"] = float(vrow.get("VOLUME_BULK", 0) or 0)
                row["VOLUME_KG"]   = float(vrow.get("VOLUME_KG", 0) or 0)
                row["NB_OPS"]      = int(vrow.get("NB_OPS", 0) or 0)
            else:
                row["VOLUME_TEU"] = row["VOLUME_BULK"] = row["VOLUME_KG"] = 0.0
                row["NB_OPS"] = 0
            # Budget RUBRIKS
            if bud_year_indexed is not None and (id_crm, y) in bud_year_indexed.index:
                brow = bud_year_indexed.loc[(id_crm, y)]
                if isinstance(brow, pd.DataFrame): brow = brow.iloc[0]
                row["CAP_PFA"]   = float(brow.get("CAP_PFA", 0) or 0)
                row["CAP_REEL"]  = float(brow.get("CAP_REEL", 0) or 0)
                row["ECART"]     = float(brow.get("ECART", 0) or 0)
                pct = brow.get("ECART_PCT")
                row["ECART_PCT"] = float(pct) if pd.notna(pct) else None
            else:
                row["CAP_PFA"] = row["CAP_REEL"] = row["ECART"] = 0.0
                row["ECART_PCT"] = None
            # Métier dominant
            if metier_dom_indexed is not None and alias_stat in metier_dom_indexed.index:
                m = metier_dom_indexed.loc[alias_stat]
                if isinstance(m, pd.DataFrame): m = m.iloc[0]
                row["METIER_DOMINANT"] = m.get("METIER_DOMINANT", "")
            else:
                row["METIER_DOMINANT"] = ""
            row["EST_WHITE_SPACE"] = alias_stat in ws_set
            rows.append(row)

    df360 = pd.DataFrame(rows)

    # Pipeline opportunités : 1 valeur par client (pas par année)
    if len(pipe_agg):
        df360 = df360.merge(pipe_agg, on="ID_CRM", how="left")
    else:
        df360["NB_OPP_OUVERTES"] = 0
        df360["PIPELINE_CAP_POND"] = 0.0
        df360["PIPELINE_CAP_BRUT"] = 0.0
    df360[["NB_OPP_OUVERTES", "PIPELINE_CAP_POND", "PIPELINE_CAP_BRUT"]] = (
        df360[["NB_OPP_OUVERTES", "PIPELINE_CAP_POND", "PIPELINE_CAP_BRUT"]].fillna(0))

    # --- Onglet 2 : synthèse client (cumul 3 ans) ----------------------------
    synth = (df360.groupby(["ID_CRM", "NOM_CLIENT", "SECTEUR", "LOCKED",
                            "ID_IRIS", "ID_STATCOM", "EST_AGL_CLIENT",
                            "METIER_DOMINANT", "EST_WHITE_SPACE",
                            "NB_OPP_OUVERTES", "PIPELINE_CAP_POND", "PIPELINE_CAP_BRUT"], dropna=False)
             .agg(VOLUME_TEU_3ANS=("VOLUME_TEU", "sum"),
                  VOLUME_BULK_3ANS=("VOLUME_BULK", "sum"),
                  VOLUME_KG_3ANS=("VOLUME_KG", "sum"),
                  NB_OPS_3ANS=("NB_OPS", "sum"),
                  CAP_PFA_TOTAL=("CAP_PFA", "sum"),
                  CAP_REEL_TOTAL=("CAP_REEL", "sum"),
                  ECART_TOTAL=("ECART", "sum"),
                  ANNEES_PRESENT=("ANNEE", "nunique"))
             .reset_index())
    synth["ECART_PCT_GLOBAL"] = synth.apply(
        lambda r: r["ECART_TOTAL"] / r["CAP_PFA_TOTAL"] if r["CAP_PFA_TOTAL"] > 0 else None, axis=1)

    # --- Onglet 3 : Top 100 par CAP réel cumulé ----------------------------
    top_cap = synth.sort_values("CAP_REEL_TOTAL", ascending=False).head(100)

    # --- Onglet 4 : Top 100 white spaces ----------------------------------
    if whites is not None and "VOLUME_PRINCIPAL" in whites:
        whites["VOLUME_PRINCIPAL"] = _num(whites, "VOLUME_PRINCIPAL")
        whites["NB_OPERATIONS"] = _num(whites, "NB_OPERATIONS")
        top_ws = whites.sort_values("VOLUME_PRINCIPAL", ascending=False).head(100)
    else:
        top_ws = pd.DataFrame()

    # --- Onglet 5 : KPI globaux ------------------------------------------
    kpis = {
        "Nb clients RMC": [len(rmc)],
        "Nb clients AGL (avec IRIS)": [int(df360["EST_AGL_CLIENT"].sum() / max(len(years), 1))],
        "CAP PFA total cumulé (FCFA)": [synth["CAP_PFA_TOTAL"].sum()],
        "CAP RÉEL total cumulé (FCFA)": [synth["CAP_REEL_TOTAL"].sum()],
        "Écart global (FCFA)": [synth["ECART_TOTAL"].sum()],
        "Écart global %": [synth["ECART_TOTAL"].sum() / synth["CAP_PFA_TOTAL"].sum()
                           if synth["CAP_PFA_TOTAL"].sum() > 0 else 0],
        "Volume TEU total (3 ans)": [synth["VOLUME_TEU_3ANS"].sum()],
        "Volume Bulk total (3 ans)": [synth["VOLUME_BULK_3ANS"].sum()],
        "Volume Kg total (3 ans, aérien)": [synth["VOLUME_KG_3ANS"].sum()],
        "Nb white spaces (prospects)": [len(top_ws)],
        "Pipeline CAP brut (€/FCFA)": [synth["PIPELINE_CAP_BRUT"].sum()],
        "Pipeline CAP pondéré": [synth["PIPELINE_CAP_POND"].sum()],
    }
    df_kpi = pd.DataFrame({"Indicateur": list(kpis.keys()),
                            "Valeur": [v[0] for v in kpis.values()]})

    # --- Écriture du fichier multi-onglets ---------------------------------
    target.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(target, engine="openpyxl") as writer:
        df_kpi.to_excel(writer,    sheet_name="kpi_global",       index=False)
        synth.to_excel(writer,     sheet_name="360_client_total", index=False)
        df360.to_excel(writer,     sheet_name="360_client_annee", index=False)
        top_cap.to_excel(writer,   sheet_name="top_clients_cap",  index=False)
        if len(top_ws):
            top_ws.to_excel(writer, sheet_name="top_white_spaces", index=False)

    print(f"[ok] Écrit : {target}")
    print(f"     {len(df360):,} lignes client × année")
    print(f"     {len(synth):,} lignes synthèse client")
    print(f"     {len(top_cap):,} clients top CAP")
    print(f"     {len(top_ws):,} top white spaces")
    print("\n=== KPIs globaux ===")
    for k, v in kpis.items():
        if isinstance(v[0], float):
            if "Écart global %" in k:
                print(f"  {k:<35} : {v[0]*100:+.1f} %")
            elif v[0] >= 1e6:
                print(f"  {k:<35} : {v[0]/1e9:,.2f} Mds")
            else:
                print(f"  {k:<35} : {v[0]:,.0f}")
        else:
            print(f"  {k:<35} : {v[0]:,}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
