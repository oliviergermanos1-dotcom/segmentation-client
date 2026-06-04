#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — 13 : Comparaison N vs N-1 entre deux snapshots
============================================================
Confronte deux snapshots (sortie de 12_snapshot.py) pour produire la
matière COMEX : qui a gagné/perdu en PDM, qui est passé d'AMBIGU à AUTO,
quels white spaces nouveaux/comblés, quels clients ont migré dans la matrice
BCG, etc.

Sorties dans --outdir :
  comparison_pdm.csv         : PDM par métier × année, ΔPDM
  comparison_whitespaces.csv : white spaces apparus/disparus
  comparison_rmc.csv         : évolution rattachement IRIS/STATCOM, secteur
  comparison_summary.json    : top changements pour l'onglet 12

Usage :
    python 13_compare.py --n  ./snapshots/2026-06-01_0900 \
                         --n1 ./snapshots/2026-05-01_0900 \
                         --outdir ./out/agreg
"""
from __future__ import annotations
import argparse
import json
import sys
from pathlib import Path

import pandas as pd


def _load_optional(path: Path):
    if not path.exists(): return None
    if path.suffix.lower() in (".xlsx", ".xls"): return pd.read_excel(path, dtype=str)
    return pd.read_csv(path, dtype=str)


def _num(s):
    return (pd.to_numeric(s.astype(str).str.replace(r"[ ,]", "", regex=True),
                          errors="coerce").fillna(0))


def compare_pdm(snap_n: Path, snap_n1: Path):
    a = _load_optional(snap_n / "agreg/pdm.csv")
    b = _load_optional(snap_n1 / "agreg/pdm.csv")
    if a is None or b is None: return None, {}
    for df in (a, b):
        for c in ("PDM", "VOLUME_AGL", "VOLUME_MARCHE"):
            if c in df: df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
        df["ANNEE"] = pd.to_numeric(df["ANNEE"], errors="coerce").astype("Int64")
    # On compare à année commune la plus récente.
    common = sorted(set(a["ANNEE"]).intersection(set(b["ANNEE"])))
    if not common: return None, {}
    y = common[-1]
    A = a[a["ANNEE"] == y][["METIER", "PDM", "VOLUME_AGL", "VOLUME_MARCHE"]].add_suffix("_N")
    B = b[b["ANNEE"] == y][["METIER", "PDM", "VOLUME_AGL", "VOLUME_MARCHE"]].add_suffix("_N1")
    A = A.rename(columns={"METIER_N": "METIER"})
    B = B.rename(columns={"METIER_N1": "METIER"})
    out = A.merge(B, on="METIER", how="outer").fillna(0)
    out["DELTA_PDM"]    = out["PDM_N"] - out["PDM_N1"]
    out["DELTA_VOL_AGL"] = out["VOLUME_AGL_N"] - out["VOLUME_AGL_N1"]
    out = out.sort_values("DELTA_PDM", ascending=False)
    out.insert(0, "ANNEE", int(y))
    summary_top = out.head(3)[["METIER", "DELTA_PDM"]].to_dict(orient="records")
    summary_bot = out.tail(3)[["METIER", "DELTA_PDM"]].to_dict(orient="records")
    return out, {"annee_comparee": int(y),
                 "top_gains_pdm": summary_top,
                 "top_pertes_pdm": summary_bot}


def compare_whitespaces(snap_n: Path, snap_n1: Path):
    a = _load_optional(snap_n / "agreg/whitespaces.csv")
    b = _load_optional(snap_n1 / "agreg/whitespaces.csv")
    if a is None or b is None: return None, {}
    setA = set(a["NOM_BASE"].fillna(""))
    setB = set(b["NOM_BASE"].fillna(""))
    apparus  = setA - setB
    combles  = setB - setA
    communs  = setA & setB
    rows = []
    for n in sorted(apparus):  rows.append({"NOM_BASE": n, "STATUT_CHANGE": "APPARU"})
    for n in sorted(combles):  rows.append({"NOM_BASE": n, "STATUT_CHANGE": "COMBLE"})
    for n in sorted(communs):  rows.append({"NOM_BASE": n, "STATUT_CHANGE": "STABLE"})
    df = pd.DataFrame(rows)
    return df, {"nouveaux_white_spaces": len(apparus),
                "white_spaces_combles":  len(combles),
                "stables":               len(communs)}


def compare_rmc(snap_n: Path, snap_n1: Path):
    a = _load_optional(snap_n / "RMC.xlsx")
    b = _load_optional(snap_n1 / "RMC.xlsx")
    if a is None or b is None: return None, {}
    def keyfn(df):
        df = df.copy()
        df["_K"] = df["ID_CRM"].fillna("").astype(str)
        return df.set_index("_K")
    A = keyfn(a); B = keyfn(b)
    common = A.index.intersection(B.index)
    rows = []
    n_iris_gain = n_stat_gain = 0
    n_iris_perte = n_stat_perte = 0
    for k in common:
        ra, rb = A.loc[k], B.loc[k]
        was_iris  = bool(str(rb.get("ID_IRIS",  "") or "").strip())
        now_iris  = bool(str(ra.get("ID_IRIS",  "") or "").strip())
        was_stat  = bool(str(rb.get("ID_STATCOM", "") or "").strip())
        now_stat  = bool(str(ra.get("ID_STATCOM", "") or "").strip())
        sec_change = str(ra.get("SECTEUR", "")) != str(rb.get("SECTEUR", ""))
        if (now_iris != was_iris) or (now_stat != was_stat) or sec_change:
            rows.append({
                "ID_CRM": k,
                "NOM_CANONIQUE": ra.get("NOM_CANONIQUE", ""),
                "IRIS_N1": "✓" if was_iris else "—",
                "IRIS_N":  "✓" if now_iris else "—",
                "STATCOM_N1": "✓" if was_stat else "—",
                "STATCOM_N":  "✓" if now_stat else "—",
                "SECTEUR_N1": rb.get("SECTEUR", ""),
                "SECTEUR_N":  ra.get("SECTEUR", ""),
            })
        if now_iris and not was_iris: n_iris_gain += 1
        if was_iris and not now_iris: n_iris_perte += 1
        if now_stat and not was_stat: n_stat_gain += 1
        if was_stat and not now_stat: n_stat_perte += 1
    nouv = set(A.index) - set(B.index)
    perdus = set(B.index) - set(A.index)
    for k in sorted(nouv)[:50]:
        rows.append({"ID_CRM": k, "NOM_CANONIQUE": A.loc[k].get("NOM_CANONIQUE", ""),
                     "IRIS_N1": "—", "IRIS_N": "✓" if str(A.loc[k].get("ID_IRIS","")) else "—",
                     "STATCOM_N1": "—", "STATCOM_N": "✓" if str(A.loc[k].get("ID_STATCOM","")) else "—",
                     "SECTEUR_N1": "", "SECTEUR_N": A.loc[k].get("SECTEUR", "")})
    df = pd.DataFrame(rows)
    return df, {
        "clients_communs": len(common),
        "nouveaux_clients": len(nouv),
        "clients_disparus": len(perdus),
        "iris_rattaches_nouveau": n_iris_gain,
        "iris_rattaches_perdus": n_iris_perte,
        "statcom_rattaches_nouveau": n_stat_gain,
        "statcom_rattaches_perdus": n_stat_perte,
    }


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AGL BUDGET — 13 comparaison N vs N-1")
    ap.add_argument("--n",      required=True, help="Snapshot période courante (N)")
    ap.add_argument("--n1",     required=True, help="Snapshot période précédente (N-1)")
    ap.add_argument("--outdir", default="agreg", help="Dossier de sortie")
    args = ap.parse_args(argv)

    n  = Path(args.n);  n1 = Path(args.n1)
    if not n.exists():  sys.exit(f"ERREUR : --n introuvable : {n}")
    if not n1.exists(): sys.exit(f"ERREUR : --n1 introuvable : {n1}")

    outdir = Path(args.outdir); outdir.mkdir(parents=True, exist_ok=True)
    summary = {"snapshot_n": n.name, "snapshot_n1": n1.name}

    pdm, s = compare_pdm(n, n1)
    if pdm is not None:
        pdm.to_csv(outdir / "comparison_pdm.csv", index=False, encoding="utf-8")
        summary["pdm"] = s

    ws, s = compare_whitespaces(n, n1)
    if ws is not None:
        ws.to_csv(outdir / "comparison_whitespaces.csv", index=False, encoding="utf-8")
        summary["whitespaces"] = s

    rmc, s = compare_rmc(n, n1)
    if rmc is not None:
        rmc.to_csv(outdir / "comparison_rmc.csv", index=False, encoding="utf-8")
        summary["rmc"] = s

    with open(outdir / "comparison_summary.json", "w", encoding="utf-8") as f:
        json.dump(summary, f, indent=2, ensure_ascii=False)

    print(f"[ok] comparaison N={n.name} vs N-1={n1.name} écrite dans {outdir}")
    print(json.dumps(summary, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
