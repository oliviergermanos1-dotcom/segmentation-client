#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Référentiel client FINAL (maillage + secteurs certifiés)
=====================================================================
Fusionne en un seul livrable, clé = ID_UNIQUE :
  - l'ossature par client (clients_master_v2 : IDs par base, présences) ;
  - le secteur CERTIFIÉ + niveau de fiabilité (05f) ;
  - les conflits tranchés manuellement (conflits_resolus.json) ;
  - le rattachement à l'ENTITÉ du maillage (consolide_graphe_rf : clients
    de toutes bases regroupés en 1 entité réelle), via le nom STATCOM.

Sortie : referentiel_final.xlsx
  - 1_fiable        : secteur fiable (CRM officiel / marchandise forte / croisé / tranché)
  - 2_a_verifier    : PROBABLE (signal faible) — à confirmer
  - 3_non_qualifie  : aucun signal secteur
  - kpi             : synthèse
"""
import argparse, json, sys
from pathlib import Path
import pandas as pd

FIABLES = {"CERTIFIE_CROISE", "CERTIFIE_CRM", "CERTIFIE_MARCH", "TRANCHE_MANUEL"}


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--master-v2", required=True)
    ap.add_argument("--certifies", required=True)
    ap.add_argument("--conflits", help="conflits_resolus.json")
    ap.add_argument("--mesh", help="entites_graphe.xlsx")
    ap.add_argument("--out", default="referentiel_final.xlsx")
    args = ap.parse_args(argv)

    m = pd.read_excel(args.master_v2, sheet_name="clients_master")
    cert = pd.read_excel(args.certifies, sheet_name="clients")[
        ["ID_UNIQUE", "SECTEUR_CERTIFIE", "NIVEAU", "DETAIL_VOTE"]]
    ref = m.merge(cert, on="ID_UNIQUE", how="left")

    # --- Conflits tranchés manuellement -----------------------------------
    if args.conflits and Path(args.conflits).exists():
        reso = json.load(open(args.conflits))
        msk = ref["ID_CRM"].astype(str).isin(reso)
        ref.loc[msk, "SECTEUR_CERTIFIE"] = ref.loc[msk, "ID_CRM"].astype(str).map(reso)
        ref.loc[msk, "NIVEAU"] = "TRANCHE_MANUEL"
        print(f"[conflits] {msk.sum()} secteurs tranchés appliqués")

    # --- Rattachement entité du maillage (via nom STATCOM) ----------------
    if args.mesh and Path(args.mesh).exists():
        mem = pd.read_excel(args.mesh, sheet_name="membres")
        stat = mem[mem["source"] == "STATCOM"][["ID_UNIQUE", "name"]].rename(
            columns={"ID_UNIQUE": "MESH_ENTITY", "name": "_n"})
        stat = stat.drop_duplicates("_n")
        ref = ref.merge(stat, left_on="ID_STATCOM", right_on="_n", how="left").drop(columns=["_n"])
        ref = ref.drop_duplicates("ID_UNIQUE")   # le join nom peut dupliquer
        n_mesh = ref["MESH_ENTITY"].notna().sum()
        print(f"[mesh] {n_mesh:,} clients rattachés à une entité du maillage")

    # --- Secteur final + fiabilité ----------------------------------------
    ref["SECTEUR_FINAL"] = ref["SECTEUR_CERTIFIE"].fillna(ref["SECTEUR"])
    ref["FIABILITE"] = ref["NIVEAU"].where(ref["NIVEAU"].notna(), "NON_QUALIFIE")

    cols = [c for c in ["ID_UNIQUE", "ID_CRM", "ID_IRIS", "ID_STATCOM", "MESH_ENTITY",
            "NOM_CLIENT", "SECTEUR_FINAL", "FIABILITE", "DETAIL_VOTE",
            "PRESENCE_CRM", "PRESENCE_IRIS", "PRESENCE_STATCOM", "PRESENCE_RUBRIKS",
            "NB_BL_STATCOM"] if c in ref.columns]
    ref = ref[cols]

    fiable = ref[ref["FIABILITE"].isin(FIABLES)]
    averif = ref[ref["FIABILITE"].isin(["PROBABLE", "A_TRANCHER"])]
    nonq   = ref[ref["FIABILITE"] == "NON_QUALIFIE"]

    kpi = pd.DataFrame({
        "Catégorie": ["TOTAL", "1_fiable", "  dont croisé (or)", "  dont CRM officiel",
                      "  dont marchandise forte", "  dont tranché manuel",
                      "2_a_verifier", "3_non_qualifie", "", "% FIABLE"],
        "Clients": [len(ref), len(fiable),
                    (ref["FIABILITE"] == "CERTIFIE_CROISE").sum(),
                    (ref["FIABILITE"] == "CERTIFIE_CRM").sum(),
                    (ref["FIABILITE"] == "CERTIFIE_MARCH").sum(),
                    (ref["FIABILITE"] == "TRANCHE_MANUEL").sum(),
                    len(averif), len(nonq), "", f"{len(fiable)/len(ref)*100:.1f}%"]})

    with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
        kpi.to_excel(xl, sheet_name="kpi", index=False)
        fiable.to_excel(xl, sheet_name="1_fiable", index=False)
        averif.to_excel(xl, sheet_name="2_a_verifier", index=False)
        nonq.to_excel(xl, sheet_name="3_non_qualifie", index=False)

    print(f"\n=== RÉFÉRENTIEL FINAL ===")
    print(f"  Total clients   : {len(ref):,}")
    print(f"  1_fiable        : {len(fiable):,} ({len(fiable)/len(ref)*100:.1f}%)")
    print(f"  2_a_verifier    : {len(averif):,}")
    print(f"  3_non_qualifie  : {len(nonq):,}")
    print(f"[ok] → {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
