#!/usr/bin/env python3
"""05c — Sectorisation des white spaces TM_xxx via marchandises STATCOM.

Pour chaque TM_xxx du clients_master sans secteur :
  1. On joint STATCOM (toutes années, tous flux) sur le nom client
  2. Chaque BL → marchandise → secteur (mapping fourni)
  3. Pondération par POIDS PHYSIQUE (somme kg + bulk + teu*15000)
  4. Décision :
       dominance ≥60%  → INFERE (haute)
       dominance 40-60% → INFERE_MOYEN
       dominance <40%   → MIXTE (top secteur quand même)
       0 marchandise    → INFERE_DEFAUT (top secteur du métier dominant)

Sortie : clients_master_v2.xlsx avec SECTEUR 100% rempli + SECTEUR_SOURCE.
"""

import argparse
import sys
from pathlib import Path

import pandas as pd

GENERIC_MERCHANDISES = {
    "DIVERS PIECES DETACHEES",
    "PRODUITS CHIMIQUES",
    "MACHINES",
    "EMBALLAGES (TOUT GENRES)",
    "MATIERES PLASTIQUES",
    "PRODUITS MANUFACTURE",
    "AUTRES",
    "DIVERS",
}

METIER_TO_DEFAULT_SECTEUR = {
    "Import Aérien": "FAST MOVING CONSUMER GOODS (FMCG)",
    "Export Aérien": "AGRICULTURE",
    "Import Maritime": "FAST MOVING CONSUMER GOODS (FMCG)",
    "Export Maritime": "AGRICULTURE",
    "Hinterland Import": "FAST MOVING CONSUMER GOODS (FMCG)",
    "Hinterland Export": "AGRICULTURE",
}


def poids_unifie(row):
    """Poids physique unifié en kg équivalent (TEU ~15000kg, bulk = tonnes ?)."""
    kg = float(row.get("volume_kg") or 0)
    bulk = float(row.get("volume_bulk") or 0) * 1000  # bulk en tonnes → kg
    teu = float(row.get("volume_teu") or 0) * 15000   # 1 TEU ≈ 15 t
    return max(kg + bulk + teu, 1.0)  # plancher 1kg pour éviter zéro


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--master", required=True, help="clients_master.xlsx en entrée")
    ap.add_argument("--statcom", required=True, help="STATCOM_consolide.xlsx")
    ap.add_argument("--mapping-marchandises", required=True,
                    help="Fichier avec colonnes SECTEUR + MARCHANDISES (feuille 'tim')")
    ap.add_argument("--out", required=True, help="clients_master_v2.xlsx en sortie")
    ap.add_argument("--review-out", required=True,
                    help="whitespaces_mixtes_a_revoir.xlsx (cas <40% dominance)")
    args = ap.parse_args(argv)

    out_path = Path(args.out)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # === 1. Charger les sources ============================================
    print("[1/5] Chargement…")
    master = pd.read_excel(args.master, sheet_name="clients_master")
    print(f"     Master : {len(master):,} clients")

    statcom = pd.read_excel(args.statcom)
    print(f"     STATCOM : {len(statcom):,} lignes")

    seg = pd.read_excel(args.mapping_marchandises, sheet_name="tim",
                        usecols=["SECTEUR", "MARCHANDISES"])
    seg["MARCHANDISES"] = seg["MARCHANDISES"].astype(str).str.upper().str.strip()
    mapping = dict(zip(seg["MARCHANDISES"], seg["SECTEUR"]))
    print(f"     Mapping marchandises : {len(mapping):,} entrées · {seg['SECTEUR'].nunique()} secteurs")

    # === 2. Préparation STATCOM ============================================
    print("[2/5] Pré-calcul poids unifié + jointure marchandise→secteur…")
    statcom["marchandise_up"] = statcom["marchandise"].astype(str).str.upper().str.strip()
    statcom["secteur_inf"] = statcom["marchandise_up"].map(mapping)
    statcom["poids"] = statcom.apply(poids_unifie, axis=1)
    statcom["is_generic"] = statcom["marchandise_up"].isin(GENERIC_MERCHANDISES)

    couv = statcom["secteur_inf"].notna().mean() * 100
    print(f"     Couverture marchandise→secteur : {couv:.1f}%")
    print(f"     Marchandises génériques (ignorées en 1ère passe) : "
          f"{statcom['is_generic'].sum():,} BL")

    # === 3. Cibler les TM_xxx sans secteur =================================
    print("[3/5] Identification des clients à sectoriser…")
    tm = master[master["SECTEUR"].isna()].copy()
    tm_names = set(tm["ID_STATCOM"].dropna().astype(str))
    print(f"     {len(tm):,} TM_xxx à sectoriser")

    # Master.ID_STATCOM = nom canonique court (ex "15MCC")
    # STATCOM_norm.NOM_BASE = même forme canonique → bonne clé de jointure
    join_col = "NOM_BASE" if "NOM_BASE" in statcom.columns else "client"
    tm_names = set(tm["ID_STATCOM"].dropna().astype(str))
    sc = statcom[statcom[join_col].astype(str).isin(tm_names)].copy()
    sc = sc.drop(columns=["client"], errors="ignore").rename(columns={join_col: "client"})
    print(f"     {len(sc):,} BL STATCOM concernés")

    # === 4. Sectorisation par client =======================================
    print("[4/5] Calcul du secteur dominant par client…")

    results = []
    for client_name, grp in sc.groupby("client"):
        # 1ère passe : ignorer les marchandises génériques
        grp_specific = grp[~grp["is_generic"] & grp["secteur_inf"].notna()]
        used_grp = grp_specific if not grp_specific.empty else grp[grp["secteur_inf"].notna()]

        if used_grp.empty:
            # Aucune marchandise classable → fallback métier dominant
            metier_dom = grp.groupby("metier")["poids"].sum().idxmax() if not grp.empty else None
            secteur = METIER_TO_DEFAULT_SECTEUR.get(metier_dom, "OTHERS")
            results.append({
                "client": client_name,
                "secteur": secteur,
                "dominance": 0.0,
                "source": "INFERE_DEFAUT",
                "nb_bl": len(grp),
                "poids_total": grp["poids"].sum() if not grp.empty else 0,
                "metier_dominant": metier_dom,
                "top2": "",
            })
            continue

        dist = used_grp.groupby("secteur_inf")["poids"].sum().sort_values(ascending=False)
        total = dist.sum()
        top_secteur = dist.index[0]
        top_share = dist.iloc[0] / total if total > 0 else 0
        top2_label = " / ".join(f"{s}={p/total*100:.0f}%"
                                for s, p in dist.head(2).items()) if total > 0 else ""

        if top_share >= 0.60:
            source = "INFERE"
        elif top_share >= 0.40:
            source = "INFERE_MOYEN"
        else:
            source = "MIXTE"

        results.append({
            "client": client_name,
            "secteur": top_secteur,
            "dominance": round(top_share, 3),
            "source": source,
            "nb_bl": len(grp),
            "poids_total": round(total, 0),
            "metier_dominant": grp.groupby("metier")["poids"].sum().idxmax() if not grp.empty else None,
            "top2": top2_label,
        })

    res = pd.DataFrame(results)
    print(f"     Inférences calculées : {len(res):,}")
    print(f"     Répartition par confiance :")
    print(res["source"].value_counts().to_string())

    # === 5. Injecter dans le master ========================================
    print("[5/5] Mise à jour master + écriture…")
    res_idx = res.set_index("client")
    master["SECTEUR_SOURCE"] = "CRM"  # par défaut pour les non-TM
    master.loc[master["SECTEUR"].isna(), "SECTEUR_SOURCE"] = None

    # Update colonnes
    def lookup(row, field):
        if row["ID_STATCOM"] in res_idx.index:
            return res_idx.at[row["ID_STATCOM"], field]
        return None

    mask = master["SECTEUR"].isna() & master["ID_STATCOM"].notna()
    master.loc[mask, "SECTEUR"] = master[mask].apply(lambda r: lookup(r, "secteur"), axis=1)
    master.loc[mask, "SECTEUR_SOURCE"] = master[mask].apply(lambda r: lookup(r, "source"), axis=1)
    # colonnes traçabilité
    master["DOMINANCE_SECTEUR"] = master.apply(
        lambda r: lookup(r, "dominance") if r["ID_STATCOM"] in res_idx.index else None,
        axis=1)
    master["NB_BL_STATCOM"] = master.apply(
        lambda r: lookup(r, "nb_bl") if r["ID_STATCOM"] in res_idx.index else None,
        axis=1)
    master["TOP2_SECTEURS"] = master.apply(
        lambda r: lookup(r, "top2") if r["ID_STATCOM"] in res_idx.index else None,
        axis=1)

    # Fallback final : si encore NaN (TM sans aucun BL — improbable), mettre OTHERS
    nan_finals = master["SECTEUR"].isna().sum()
    if nan_finals:
        print(f"     {nan_finals} clients sans aucun BL → 'OTHERS' / INFERE_DEFAUT")
        master.loc[master["SECTEUR"].isna(), "SECTEUR_SOURCE"] = "INFERE_DEFAUT"
        master["SECTEUR"] = master["SECTEUR"].fillna("OTHERS")

    # === Écriture =========================================================
    with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
        master.to_excel(xl, sheet_name="clients_master", index=False)
        # KPI
        kpi = pd.DataFrame({
            "Métrique": [
                "Total clients",
                "Sectorisé via CRM",
                "Sectorisé INFERE (≥60%)",
                "Sectorisé INFERE_MOYEN (40-60%)",
                "Sectorisé MIXTE (<40%, à review)",
                "Sectorisé INFERE_DEFAUT (fallback métier)",
                "% SECTEUR rempli",
            ],
            "Valeur": [
                len(master),
                (master["SECTEUR_SOURCE"] == "CRM").sum(),
                (master["SECTEUR_SOURCE"] == "INFERE").sum(),
                (master["SECTEUR_SOURCE"] == "INFERE_MOYEN").sum(),
                (master["SECTEUR_SOURCE"] == "MIXTE").sum(),
                (master["SECTEUR_SOURCE"] == "INFERE_DEFAUT").sum(),
                f"{master['SECTEUR'].notna().mean()*100:.1f}%",
            ],
        })
        kpi.to_excel(xl, sheet_name="kpi_sectorisation", index=False)

    print(f"[ok] Master v2 écrit : {args.out}")

    # === Liste des MIXTE pour review manuel ===============================
    mixtes = master[master["SECTEUR_SOURCE"] == "MIXTE"].copy()
    if not mixtes.empty:
        mixtes = mixtes.sort_values("NB_BL_STATCOM", ascending=False, na_position="last")
        with pd.ExcelWriter(args.review_out, engine="openpyxl") as xl:
            mixtes[[
                "ID_UNIQUE", "NOM_CLIENT", "SECTEUR", "DOMINANCE_SECTEUR",
                "TOP2_SECTEURS", "NB_BL_STATCOM"
            ]].to_excel(xl, sheet_name="mixtes_a_revoir", index=False)
        print(f"[ok] Liste cas MIXTE écrite : {args.review_out} ({len(mixtes):,} lignes)")

    # === Summary ==========================================================
    print()
    print(f"[summary] Couverture finale :")
    print(f"  SECTEUR rempli : {master['SECTEUR'].notna().sum():,} / {len(master):,} "
          f"({master['SECTEUR'].notna().mean()*100:.1f}%)")
    print(f"  Top 10 secteurs :")
    print(master["SECTEUR"].value_counts().head(10).to_string())
    print(f"  Sources :")
    print(master["SECTEUR_SOURCE"].value_counts().to_string())

    return 0


if __name__ == "__main__":
    sys.exit(main())
