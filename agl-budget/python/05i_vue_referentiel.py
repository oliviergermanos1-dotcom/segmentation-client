#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Vue référentiel "large" : nom + id de chaque base côte à côte
==========================================================================
1 ligne par entité du maillage, colonnes :
  NOM_CRM | ID_CRM | NOM_STATCOM | ID_STATCOM | NOM_IRIS | ID_IRIS |
  NOM_RUBRIKS | ID_RUBRIKS | NOM_CONSOLIDE | ID_RETENU (+ secteur, confiance)
"""
import argparse, sys
import pandas as pd


def _lookup(path, name_col, id_col):
    """NOM_BASE → (nom brut, id) — 1ère occurrence."""
    df = pd.read_excel(path)
    df = df.dropna(subset=["NOM_BASE"]).drop_duplicates("NOM_BASE")
    nm = df[name_col] if name_col in df.columns else df["NOM_BASE"]
    idv = df[id_col].astype(str) if id_col in df.columns else ""
    return {str(b): (str(n), str(i)) for b, n, i in zip(df["NOM_BASE"], nm, idv if id_col in df.columns else [""]*len(df))}


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--mesh", required=True)
    ap.add_argument("--crm", required=True); ap.add_argument("--iris", required=True)
    ap.add_argument("--statcom", required=True); ap.add_argument("--rubriks", required=True)
    ap.add_argument("--out", default="referentiel_vue.xlsx")
    args = ap.parse_args(argv)

    ent = pd.read_excel(args.mesh, sheet_name="entites").set_index("ID_UNIQUE")
    mem = pd.read_excel(args.mesh, sheet_name="membres")

    look = {
        "CRM":     _lookup(args.crm, "Nom du compte", "CRM ID du Compte"),
        "IRIS":    _lookup(args.iris, "CLIENT", "ID"),
        "STATCOM": _lookup(args.statcom, "CLIENT_RESOLU", "ID_STATCOM"),
        "RUBRIKS": _lookup(args.rubriks, "nom_client", "ID_CONCERTO"),
    }

    rows = []
    for eid, g in mem.groupby("ID_UNIQUE"):
        rec = {s: ("", "") for s in look}
        for _, m in g.iterrows():
            src, nb = m["source"], str(m["name"])
            if src in look and nb in look[src]:
                rec[src] = look[src][nb]
        # id retenu : priorité CRM > RUBRIKS (Concerto) > IRIS > STATCOM
        id_ret = (rec["CRM"][1] or rec["RUBRIKS"][1] or rec["IRIS"][1]
                  or rec["STATCOM"][1] or eid)
        e = ent.loc[eid] if eid in ent.index else None
        rows.append({
            "NOM_CRM": rec["CRM"][0], "ID_CRM": rec["CRM"][1],
            "NOM_STATCOM": rec["STATCOM"][0], "ID_STATCOM": rec["STATCOM"][1],
            "NOM_IRIS": rec["IRIS"][0], "ID_IRIS": rec["IRIS"][1],
            "NOM_RUBRIKS": rec["RUBRIKS"][0], "ID_RUBRIKS": rec["RUBRIKS"][1],
            "NOM_CONSOLIDE": (e["NOM_PRINCIPAL"] if e is not None else ""),
            "ID_RETENU": id_ret,
            "SECTEUR": (e["SECTEUR_VOTE"] if e is not None else ""),
            "CONFIANCE_LIEN": (e["CONFIANCE_LIEN"] if e is not None else ""),
            "ID_ENTITE": eid,
        })
    out = pd.DataFrame(rows)
    # priorité d'affichage : entités multi-bases d'abord
    out["_nb"] = (out[["ID_CRM","ID_STATCOM","ID_IRIS","ID_RUBRIKS"]] != "").sum(axis=1)
    out = out.sort_values("_nb", ascending=False).drop(columns="_nb")
    out.to_excel(args.out, index=False)
    print(f"[ok] {len(out):,} entités → {args.out}")
    print(f"  multi-bases (≥2) : {((out[['ID_CRM','ID_STATCOM','ID_IRIS','ID_RUBRIKS']]!='').sum(axis=1)>=2).sum():,}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
