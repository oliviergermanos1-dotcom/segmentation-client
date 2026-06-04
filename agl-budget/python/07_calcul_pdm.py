#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — 07 : Calcul de la PDM AGL (Part De Marché)
=======================================================
Spec §2 (PDM dynamique gain/perte par métier × année) + §3 (champ VOLUME_3ANS RMC).

Définition :
  PDM_AGL(métier, année) = Σ volumes clients AGL / Σ volumes marché total CIV

Clients AGL = clients du RMC qui ont un ID_IRIS (rattachés à la base AGL).

Entrée :
  --rmc          RMC.xlsx              (sortie de 05)
  --volume-client agreg/volume_client.csv (sortie de 06)
  --marche-total  agreg/marche_total.csv  (sortie de 06)

Sortie :
  --outdir agreg/
    pdm.csv         : MÉTIER × ANNÉE → volume AGL, volume marché, PDM, CAGR
    pdm_clients.csv : volumes clients AGL (NOM_BASE + ID_CRM + ID_IRIS)
"""
from __future__ import annotations
import argparse
import sys
from pathlib import Path

import pandas as pd


def _load(p):
    p = Path(p)
    if p.suffix.lower() in (".xlsx", ".xls"): return pd.read_excel(p, dtype=str)
    return pd.read_csv(p, dtype=str, sep=None, engine="python")


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AGL BUDGET — 07 PDM AGL")
    ap.add_argument("--rmc",           required=True)
    ap.add_argument("--volume-client", required=True)
    ap.add_argument("--marche-total",  required=True)
    ap.add_argument("--outdir",        default="agreg")
    args = ap.parse_args(argv)

    rmc    = _load(args.rmc)
    vc     = _load(args.volume_client)
    marche = _load(args.marche_total)

    # On force les types numériques sur les volumes/années.
    for col in ("VOLUME_TEU", "VOLUME_BULK", "VOLUME_KG"):
        if col in vc:     vc[col]     = pd.to_numeric(vc[col],     errors="coerce").fillna(0)
        if col in marche: marche[col] = pd.to_numeric(marche[col], errors="coerce").fillna(0)
    vc["ANNEE"]     = pd.to_numeric(vc["ANNEE"], errors="coerce").astype("Int64")
    marche["ANNEE"] = pd.to_numeric(marche["ANNEE"], errors="coerce").astype("Int64")

    # Clients AGL = ceux du RMC avec ID_IRIS renseigné (et donc présents dans IRIS).
    agl = rmc[rmc["ID_IRIS"].fillna("").astype(str).str.len().gt(0)]
    agl_keys = set()
    # On rejoint le RMC à STATCOM via ID_STATCOM si disponible, sinon via le nom canonique normalisé.
    if "ID_STATCOM" in agl and "ID_STATCOM" in vc:
        agl_keys = set(agl["ID_STATCOM"].dropna().astype(str)) - {""}
        vc["_KEY_AGL"] = vc["ID_STATCOM"].astype(str)
    else:
        # Fallback : on aligne sur le nom de base normalisé.
        agl_keys = set(agl["ALIAS_STATCOM"].dropna().astype(str).str.upper().str.strip()) - {""}
        vc["_KEY_AGL"] = vc["NOM_BASE"].astype(str).str.upper().str.strip()

    vc["EST_AGL"] = vc["_KEY_AGL"].isin(agl_keys)
    pdm_clients = vc[vc["EST_AGL"]].copy()

    # Volume principal selon métier (TEU sinon Kg pour l'aérien).
    def vol_principal(df):
        m = df["METIER"].fillna("").str.upper()
        return df["VOLUME_TEU"].where(~(m.str.contains("AERIEN") | m.str.contains("AÉRIEN")), df["VOLUME_KG"])

    vc["VOLUME_PRINCIPAL"]          = vol_principal(vc)
    pdm_clients["VOLUME_PRINCIPAL"] = vol_principal(pdm_clients)

    # Agrégat AGL par métier × année + jointure marché total.
    agl_metier = (pdm_clients.groupby(["METIER", "ANNEE"], dropna=False)["VOLUME_PRINCIPAL"]
                  .sum().reset_index().rename(columns={"VOLUME_PRINCIPAL": "VOLUME_AGL"}))

    def vol_principal_marche(df):
        m = df["METIER"].fillna("").str.upper()
        return df["VOLUME_TEU"].where(~(m.str.contains("AERIEN") | m.str.contains("AÉRIEN")), df["VOLUME_KG"])

    marche["VOLUME_MARCHE"] = vol_principal_marche(marche)
    m_metier = marche.groupby(["METIER", "ANNEE"], dropna=False)["VOLUME_MARCHE"].sum().reset_index()

    pdm = m_metier.merge(agl_metier, on=["METIER", "ANNEE"], how="left").fillna({"VOLUME_AGL": 0})
    pdm["PDM"] = pdm.apply(
        lambda r: (r["VOLUME_AGL"] / r["VOLUME_MARCHE"]) if r["VOLUME_MARCHE"] > 0 else 0.0, axis=1)

    # CAGR AGL et CAGR marché par métier (sur la période disponible).
    def add_cagr(df, val_col, name):
        out = []
        for metier, sub in df.groupby("METIER"):
            sub = sub.sort_values("ANNEE")
            if len(sub) >= 2:
                a0, a1 = int(sub["ANNEE"].iloc[0]), int(sub["ANNEE"].iloc[-1])
                v0, v1 = float(sub[val_col].iloc[0]), float(sub[val_col].iloc[-1])
                years = (a1 - a0) or 1
                cagr = (v1 / v0) ** (1 / years) - 1 if (v0 > 0 and v1 > 0) else None
            else:
                cagr = None
            out.append({"METIER": metier, name: cagr})
        return pd.DataFrame(out)

    pdm = pdm.merge(add_cagr(pdm, "VOLUME_AGL",    "CAGR_AGL"),    on="METIER", how="left")
    pdm = pdm.merge(add_cagr(pdm, "VOLUME_MARCHE", "CAGR_MARCHE"), on="METIER", how="left")

    outdir = Path(args.outdir); outdir.mkdir(parents=True, exist_ok=True)
    pdm.to_csv(outdir / "pdm.csv", index=False)
    pdm_clients[["NOM_BASE", "METIER", "ANNEE", "VOLUME_PRINCIPAL", "_KEY_AGL"]]\
        .rename(columns={"_KEY_AGL": "CLE_RMC"}).to_csv(outdir / "pdm_clients.csv", index=False)

    # Récap console.
    last = pdm["ANNEE"].max()
    if pd.notna(last):
        snap = pdm[pdm["ANNEE"] == last].sort_values("PDM", ascending=False)
        print(f"[ok] PDM AGL {int(last)} (résumé) :")
        for _, r in snap.iterrows():
            print(f"    {r['METIER']:<22} → PDM {r['PDM']:>5.1%} "
                  f"| AGL {r['VOLUME_AGL']:>10,.0f} / Marché {r['VOLUME_MARCHE']:>10,.0f}")
    print(f"[ok] écrit : {outdir/'pdm.csv'} & {outdir/'pdm_clients.csv'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
