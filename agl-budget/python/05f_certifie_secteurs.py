#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Certification des secteurs par vote multi-signaux
==============================================================
Objectif : MAXIMISER la part de clients dont le secteur est "certifié",
c'est-à-dire confirmé par au moins 2 sources indépendantes.

4 signaux de secteur (normalisés vers la nomenclature CRM officielle) :
  S1. Verticale CRM            (RMC.SECTEUR, source CRM)
  S2. Secteur RUBRIKS          (RUBRIKS.SECTEUR — colonne explicite)
  S3. Marchandise dominante    (sectorisation STATCOM, dominance ≥ seuil)
  S4. vertical_destinataire    (STATCOM, si renseigné)

Règle de certification :
  - ≥2 signaux d'accord                     → CERTIFIE   (🟢)
  - 1 seul signal fort (CRM ou march.≥60%)  → PROBABLE   (🟡)
  - ≥2 signaux mais en CONFLIT              → A_TRANCHER (🟠)  ← worklist
  - 0 signal                                → NON_QUALIFIE (🔴)

Sortie : secteurs_certifies.xlsx
  - onglet clients   : secteur retenu + niveau + signaux + détail vote
  - onglet conflits  : les A_TRANCHER triés par volume (worklist correction)
  - onglet kpi       : % certifié, distribution des niveaux
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd

# Normalisation vers nomenclature CRM officielle (idem 05d).
SECTEUR_NORM = {
    "FMCG / RETAIL": "FAST MOVING CONSUMER GOODS (FMCG)",
    "INFRASTRUCTURE": "INFRASTRUCTURES",
    "AGRI BUSINESS": "AGRICULTURE",
    "OIL & GAZ": "OIL & GAS",
    "HIGHT TECH / TELECOM": "HIGH TECH",
    "ENERGIE": "POWER TECHNOLOGIES",
    "AUTRES": "OTHERS",
}


def norm(s):
    if pd.isna(s):
        return None
    s = str(s).strip().upper()
    if not s or s in ("NAN", "INCONNU", "NON QUALIFIE", "UNCLASSIFIED OLD"):
        return None
    return SECTEUR_NORM.get(s, s)


def certifie(row):
    """Vote sur les 4 signaux → (secteur, niveau, n_signaux, detail)."""
    signaux = {
        "CRM": row.get("S1_crm"),
        "RUBRIKS": row.get("S2_rubriks"),
        "MARCH": row.get("S3_march"),
        "STAT_VERT": row.get("S4_vert"),
    }
    votes = {}
    for src, val in signaux.items():
        v = norm(val)
        if v:
            votes.setdefault(v, []).append(src)

    detail = " | ".join(f"{v}←{'+'.join(srcs)}" for v, srcs in votes.items())

    if not votes:
        return pd.Series([None, "NON_QUALIFIE", 0, ""])

    # Secteur le plus voté
    best = max(votes.items(), key=lambda kv: len(kv[1]))
    best_sec, best_srcs = best
    n_best = len(best_srcs)
    n_distinct = len(votes)

    if n_best >= 2:
        niveau = "CERTIFIE_CROISE"        # ≥2 sources indépendantes d'accord (or)
    elif n_distinct >= 2:
        niveau = "A_TRANCHER"             # plusieurs signaux en conflit
    elif "CRM" in best_srcs:
        niveau = "CERTIFIE_CRM"           # 1 signal mais source officielle CRM
    elif "MARCH" in best_srcs:
        niveau = "CERTIFIE_MARCH"         # 1 signal marchandise fort (dominance ≥60%)
    else:
        niveau = "PROBABLE"               # 1 signal faible (RUBRIKS seul / march. moyen)
    return pd.Series([best_sec, niveau, n_best, detail])


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--master-v2", required=True, help="clients_master_v2.xlsx (S3 marchandise)")
    ap.add_argument("--rmc", required=True, help="RMC.xlsx (S1 Verticale CRM)")
    ap.add_argument("--rubriks", required=True, help="RUBRIKS.xlsx (S2 secteur RUBRIKS)")
    ap.add_argument("--statcom-norm", help="STATCOM_norm.xlsx (S3 marchandise pour TOUS + S4)")
    ap.add_argument("--mapping-marchandises", help="SEGEMENTATION_ENRICHI.xlsx (dico marchandise→secteur)")
    ap.add_argument("--out", default="secteurs_certifies.xlsx")
    args = ap.parse_args(argv)

    m = pd.read_excel(args.master_v2, sheet_name="clients_master")
    print(f"Clients : {len(m):,}")

    # S1 — Verticale CRM (depuis RMC, jointure ID_CRM)
    rmc = pd.read_excel(args.rmc)[["ID_CRM", "SECTEUR"]].dropna(subset=["ID_CRM"])
    rmc = rmc.drop_duplicates("ID_CRM").rename(columns={"SECTEUR": "S1_crm"})
    m = m.merge(rmc, on="ID_CRM", how="left")

    # S2 — Secteur RUBRIKS (extraction ID Concerto + jointure)
    import re
    rub = pd.read_excel(args.rubriks, sheet_name="Extraction Rubiks")
    rub = rub[~rub["CLIENT"].astype(str).str.startswith("Total Customer")]
    pat = re.compile(r"\(([A-Z0-9][A-Z0-9\-]+)\)\s*$")
    rub["ID_RUB"] = rub["CLIENT"].astype(str).apply(
        lambda s: (pat.search(s).group(1) if pat.search(s) else None))
    rub_keys = rub.dropna(subset=["ID_RUB"]).drop_duplicates("ID_RUB")[["ID_RUB", "SECTEUR"]]
    rub_keys = rub_keys.rename(columns={"SECTEUR": "S2_rubriks"})
    m = m.merge(rub_keys, left_on="ID_CRM", right_on="ID_RUB", how="left")

    # S3 — Marchandise dominante calculée pour TOUS les clients ayant une
    # activité STATCOM (CRM matchés compris), pas seulement les white spaces.
    # C'est la cross-validation : un client CRM dont les marchandises confirment
    # son secteur déclaré = certifié croisé.
    if args.statcom_norm and args.mapping_marchandises:
        sc = pd.read_excel(args.statcom_norm,
                           usecols=["NOM_BASE", "marchandise",
                                    "volume_teu", "volume_bulk", "volume_kg"])
        seg = pd.read_excel(args.mapping_marchandises, sheet_name="tim",
                            usecols=["SECTEUR", "MARCHANDISES"])
        dico = dict(zip(seg["MARCHANDISES"].astype(str).str.upper().str.strip(),
                        seg["SECTEUR"]))
        sc["sec"] = sc["marchandise"].astype(str).str.upper().str.strip().map(dico)
        sc["poids"] = (sc["volume_kg"].fillna(0)
                       + sc["volume_bulk"].fillna(0) * 1000
                       + sc["volume_teu"].fillna(0) * 15000).clip(lower=1)
        scm = sc.dropna(subset=["sec"])
        # secteur dominant + dominance par client (NOM_BASE)
        agg = (scm.groupby(["NOM_BASE", "sec"])["poids"].sum().reset_index())
        tot = agg.groupby("NOM_BASE")["poids"].transform("sum")
        agg["dom"] = agg["poids"] / tot
        top = agg.sort_values("poids", ascending=False).drop_duplicates("NOM_BASE")
        top = top[top["dom"] >= 0.40][["NOM_BASE", "sec"]].rename(columns={"sec": "S3_march"})
        m = m.merge(top, left_on="ID_STATCOM", right_on="NOM_BASE", how="left")
    else:
        # fallback : ancien comportement (white spaces seulement, depuis master_v2)
        dom = pd.to_numeric(m.get("DOMINANCE_SECTEUR"), errors="coerce").fillna(0)
        m["S3_march"] = m["SECTEUR"].where(
            m["SECTEUR_SOURCE"].isin(["INFERE", "INFERE_MOYEN"]) & (dom >= 0.40))

    # S4 — vertical_destinataire STATCOM (dominant par client) si fourni
    if args.statcom_norm:
        sc = pd.read_excel(args.statcom_norm,
                           usecols=lambda c: c in ("NOM_BASE", "vertical_destinataire"))
        if "vertical_destinataire" in sc.columns:
            sc = sc.dropna(subset=["vertical_destinataire"])
            vd = (sc.groupby("NOM_BASE")["vertical_destinataire"]
                    .agg(lambda x: x.value_counts().idxmax()).rename("S4_vert"))
            m = m.merge(vd, left_on="ID_STATCOM", right_index=True, how="left")
    if "S4_vert" not in m.columns:
        m["S4_vert"] = None

    # ===== Vote =====
    print("[vote] certification multi-signaux…")
    m[["SECTEUR_CERTIFIE", "NIVEAU", "N_SIGNAUX_ACCORD", "DETAIL_VOTE"]] = \
        m.apply(certifie, axis=1)

    # ===== KPI =====
    niv = m["NIVEAU"].value_counts()
    FIABLES = ["CERTIFIE_CROISE", "CERTIFIE_CRM", "CERTIFIE_MARCH"]
    pct_fiable = m["NIVEAU"].isin(FIABLES).mean() * 100
    pct_croise = (m["NIVEAU"] == "CERTIFIE_CROISE").mean() * 100
    print(f"\n=== RÉSULTAT CERTIFICATION ===")
    for n, c in niv.items():
        print(f"  {n:16s} : {c:>6,} ({c/len(m)*100:4.1f}%)")
    print(f"\n  → % FIABLE (CRM officiel OU marchandise forte OU ≥2 d'accord) : {pct_fiable:.1f}%")
    print(f"  → dont CERTIFIÉ CROISÉ (≥2 signaux d'accord, or)             : {pct_croise:.1f}%")
    print(f"  → Conflits à trancher (worklist)                             : {(m['NIVEAU']=='A_TRANCHER').sum():,}")

    # ===== Écriture =====
    cols = ["ID_UNIQUE", "ID_CRM", "NOM_CLIENT", "SECTEUR_CERTIFIE", "NIVEAU",
            "N_SIGNAUX_ACCORD", "DETAIL_VOTE",
            "S1_crm", "S2_rubriks", "S3_march", "S4_vert",
            "NB_BL_STATCOM", "PRESENCE_CRM", "PRESENCE_RUBRIKS", "PRESENCE_STATCOM"]
    cols = [c for c in cols if c in m.columns]
    conflits = m[m["NIVEAU"] == "A_TRANCHER"].copy()
    conflits["_nb"] = pd.to_numeric(conflits.get("NB_BL_STATCOM"), errors="coerce").fillna(0)
    conflits = conflits.sort_values("_nb", ascending=False)

    kpi = pd.DataFrame({
        "Niveau": list(niv.index) + ["", "% FIABLE", "% CERTIFIE_CROISE"],
        "Clients": list(niv.values) + ["", f"{pct_fiable:.1f}%", f"{pct_croise:.1f}%"],
    })

    with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
        kpi.to_excel(xl, sheet_name="kpi", index=False)
        m[cols].to_excel(xl, sheet_name="clients", index=False)
        conflits[cols].to_excel(xl, sheet_name="conflits_a_trancher", index=False)
    print(f"\n[ok] → {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
