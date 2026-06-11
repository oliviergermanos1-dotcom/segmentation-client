#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Maillage 4 bases déterministe (RapidFuzz + union-find)
===================================================================
Alternative robuste à consolide_graphe.py (Splink non supervisé peu fiable :
EM ne converge pas toujours → 0 lien). Ici on réutilise le scorer ÉPROUVÉ du
pipeline (token_set_ratio + Jaro-Winkler + barrière TOKEN_FLOOR=0.80) et on
regroupe par union-find (connected components).

Étapes :
  1. chaque base → clients uniques (nom normalisé, secteur dominant)
  2. blocking par préfixe 4 lettres de chaque token significatif (multi-clé)
  3. scoring des paires candidates intra-bloc (toutes bases confondues)
  4. lien retenu si score >= --threshold (avec garde TOKEN_FLOOR)
  5. union-find → 1 grappe = 1 entité ; garde anti sur-fusion (taille max)

Sortie identique à consolide_graphe.py : entites / membres / grappes_alerte.
"""
from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path

import pandas as pd
from rapidfuzz import fuzz, distance

TOKEN_FLOOR = 0.80


# Pays / régions (UEMOA + voisins) détectés dans le nom complet.
PAYS_KW = {
    "CI": ["COTE IVOIRE", "COTE D IVOIRE", "IVOIRE", "RCI", " CIV", "ABIDJAN"],
    "BF": [" BF", "BURKINA", "OUAGA"],
    "ML": [" MALI", "BAMAKO"],
    "SN": ["SENEGAL", "DAKAR"],
    "GH": ["GHANA", "ACCRA"],
    "TG": [" TOGO", "LOME"],
    "BJ": ["BENIN", "COTONOU"],
    "NE": [" NIGER ", "NIAMEY"],
    "GN": ["GUINEE", "CONAKRY"],
    "NG": ["NIGERIA", "LAGOS"],
}


def detect_pays(full):
    s = f" {str(full).upper()} "
    for code, kws in PAYS_KW.items():
        for kw in kws:
            if kw in s:
                return code
    return ""   # pays inconnu


def _prep(path, source, name_col, sec_col, full_col="NOM_NORMALISE", id_col=None):
    if not path:
        return None
    df = pd.read_excel(path)
    nm = name_col if name_col in df.columns else None
    if not nm:
        return None
    fc = full_col if full_col in df.columns else nm
    out = pd.DataFrame()
    out["name"] = df[nm].fillna("").astype(str).str.upper().str.strip()
    out["full"] = df[fc].fillna("").astype(str).str.upper().str.strip()
    out["extid"] = (df[id_col].fillna("").astype(str).str.upper().str.strip()
                    if id_col and id_col in df.columns else "")
    out["secteur"] = (df[sec_col].fillna("").astype(str).str.upper().str.strip()
                      if sec_col and sec_col in df.columns else "")
    out = out[out["name"] != ""]

    def _mode_sec(s):
        nz = s[s != ""]
        return nz.mode().iloc[0] if len(nz) else ""
    out = out.groupby("name", as_index=False).agg(
        full=("full", "first"), extid=("extid", "first"),
        secteur=("secteur", _mode_sec))
    out["pays"] = out["full"].map(detect_pays)
    out["source"] = source
    out["uid"] = [f"{source}_{i}" for i in range(len(out))]
    return out


def _toks(name):
    return [t for t in name.split() if len(t) >= 3]


def _score(a, b, rare_a, rare_b, full_a="", full_b="", pays_a="", pays_b=""):
    """Matching à 2 niveaux + garde-fou pays.

    1. tronc (NOM_BASE) : token_sort_ratio (length-aware) + ≥1 token rare partagé.
    2. nom COMPLET (NOM_NORMALISE) : doit aussi concorder → évite de fusionner
       deux entités dont seuls les tokens amputés (pays/forme) différaient.
    3. garde-fou PAYS : 2 pays explicites différents → JAMAIS de fusion.
    """
    if not (rare_a & rare_b):
        return 0.0
    # garde-fou pays : si les deux ont un pays explicite et qu'ils diffèrent → refus
    if pays_a and pays_b and pays_a != pays_b:
        return 0.0
    tsort = fuzz.token_sort_ratio(a, b) / 100.0
    if tsort < TOKEN_FLOOR:
        return 0.0
    # confirmation sur le nom complet (lit TOUT le nom)
    if full_a and full_b:
        tfull = fuzz.token_sort_ratio(full_a, full_b) / 100.0
        if tfull < TOKEN_FLOOR - 0.05:   # léger jeu pour formes juridiques
            return 0.0
    jw = distance.JaroWinkler.normalized_similarity(a, b)
    return 0.85 * tsort + 0.15 * jw


class UF:
    """Union-find conscient du PAYS : un cluster porte son pays explicite ;
    une fusion qui joindrait 2 pays différents (même via un nœud sans pays)
    est refusée → empêche le chaînage transitif CI—(inconnu)—BF."""
    def __init__(self):
        self.p = {}
        self.cp = {}   # root -> pays explicite du cluster
    def find(self, x):
        self.p.setdefault(x, x)
        while self.p[x] != x:
            self.p[x] = self.p[self.p[x]]; x = self.p[x]
        return x
    def set_pays(self, x, pays):
        if pays:
            r = self.find(x)
            self.cp.setdefault(r, pays)
    def country_ok(self, a, c):
        pa, pc = self.cp.get(self.find(a), ""), self.cp.get(self.find(c), "")
        return not (pa and pc and pa != pc)
    def union(self, a, c, force=False):
        ra, rb = self.find(a), self.find(c)
        if ra == rb:
            return True
        pa, pb = self.cp.get(ra, ""), self.cp.get(rb, "")
        if not force and pa and pb and pa != pb:
            return False   # refus : pays explicites différents
        self.p[ra] = rb
        merged = pb or pa
        if merged:
            self.cp[rb] = merged
        return True


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--crm"); ap.add_argument("--iris")
    ap.add_argument("--statcom"); ap.add_argument("--rubriks")
    ap.add_argument("--name-col", default="NOM_BASE")
    ap.add_argument("--sec-col", default="secteur")
    ap.add_argument("--crm-id-col", default="crm_id_compte")
    ap.add_argument("--rubriks-id-col", default="ID_CONCERTO")
    ap.add_argument("--threshold", type=float, default=0.90)
    ap.add_argument("--verify-band", type=float, default=0.95,
                    help="liens fuzzy entre threshold et cette valeur = à vérifier")
    ap.add_argument("--max-cluster", type=int, default=10)
    ap.add_argument("--out", default="entites_graphe.xlsx")
    args = ap.parse_args(argv)

    id_cols = {"CRM": args.crm_id_col, "RUBRIKS": args.rubriks_id_col}
    parts = []
    for src, p in [("CRM", args.crm), ("IRIS", args.iris),
                   ("STATCOM", args.statcom), ("RUBRIKS", args.rubriks)]:
        t = _prep(p, src, args.name_col, args.sec_col, id_col=id_cols.get(src))
        if t is not None and len(t):
            parts.append(t); print(f"  {src:8s} : {len(t):,} clients uniques")
    pool = pd.concat(parts, ignore_index=True)
    print(f"  POOL     : {len(pool):,} enregistrements")

    rec = pool.set_index("uid")[["name", "full", "pays", "extid", "source", "secteur"]].to_dict("index")

    # --- Fréquence des tokens (DF) : un token fréquent = générique ---------
    df_tok = defaultdict(int)
    for r in rec.values():
        for t in set(_toks(r["name"])):
            df_tok[t] += 1
    # seuil "générique" : token présent dans > GEN_CAP noms (≈ stopword métier)
    GEN_CAP = max(25, int(0.003 * len(rec)))
    rare_of = {uid: {t for t in _toks(rec[uid]["name"]) if df_tok[t] <= GEN_CAP}
               for uid in rec}
    n_generic = sum(1 for t, c in df_tok.items() if c > GEN_CAP)
    print(f"[match] {n_generic:,} tokens génériques ignorés "
          f"(présents dans > {GEN_CAP} noms) ; blocking sur tokens RARES")

    # --- Blocking sur préfixe 5 des tokens RARES uniquement ----------------
    idx = defaultdict(list)
    for uid in rec:
        for t in rare_of[uid]:
            idx[t[:5]].append(uid)

    uf = UF()
    for uid in rec:
        uf.find(uid)
        uf.set_pays(uid, rec[uid]["pays"])
    links = []   # (a, c, score, type)

    # --- PASSE 1 : ANCRES PAR ID EXACT (Concerto) — 0 erreur ---------------
    # RUBRIKS et CRM partagent l'ID Concerto : lien certain (force=True : un
    # même ID Concerto = même entité légale, même si succursale autre pays).
    id_index = defaultdict(list)
    for uid in rec:
        eid = rec[uid]["extid"]
        if eid and len(eid) >= 5:
            id_index[eid].append(uid)
    n_anchor = 0
    for eid, uids in id_index.items():
        for k in range(1, len(uids)):
            if uf.union(uids[0], uids[k], force=True):
                links.append((uids[0], uids[k], 1.0, "ANCRE_ID"))
                n_anchor += 1
    print(f"[match] passe 1 — {n_anchor:,} liens par ID exact (Concerto)")

    # --- PASSE 2 : FUZZY (tronc + nom complet + garde-fou pays) ------------
    print("[match] passe 2 — scoring fuzzy…")
    seen = set()
    cand = []
    for b, uids in idx.items():
        if len(uids) < 2 or len(uids) > 2000:
            continue
        for i in range(len(uids)):
            for j in range(i + 1, len(uids)):
                a, c = uids[i], uids[j]
                key = (a, c) if a < c else (c, a)
                if key in seen:
                    continue
                seen.add(key)
                sc = _score(rec[a]["name"], rec[c]["name"], rare_of[a], rare_of[c],
                            rec[a]["full"], rec[c]["full"],
                            rec[a]["pays"], rec[c]["pays"])
                if sc >= args.threshold:
                    cand.append((sc, a, c))
    # meilleurs liens d'abord : ils fixent le pays du cluster en priorité
    cand.sort(reverse=True)
    n_fuzzy = n_verif = n_blocked = 0
    for sc, a, c in cand:
        if uf.union(a, c):                       # refuse si pays cluster diffère
            typ = "FORT" if sc >= args.verify_band else "A_VERIFIER"
            links.append((a, c, round(sc, 3), typ))
            n_fuzzy += 1
            n_verif += (typ == "A_VERIFIER")
        else:
            n_blocked += 1
    print(f"[match] passe 2 — {n_fuzzy:,} liens fuzzy "
          f"(dont {n_verif:,} à vérifier) · {n_blocked:,} bloqués par garde-fou pays")

    # Grappes
    groups = defaultdict(list)
    for uid in rec:
        groups[uf.find(uid)].append(uid)

    cid_to_eid = {cid: f"E{i:08d}" for i, cid in enumerate(sorted(groups))}

    # confiance par entité : a-t-elle une ancre ID ? un lien à vérifier ?
    ent_has_anchor, ent_min_fuzzy = defaultdict(bool), {}
    for a, c, sc, typ in links:
        e = cid_to_eid[uf.find(a)]
        if typ == "ANCRE_ID":
            ent_has_anchor[e] = True
        else:
            ent_min_fuzzy[e] = min(ent_min_fuzzy.get(e, 1.0), sc)

    rows = []
    for cid, members in groups.items():
        eid = cid_to_eid[cid]
        srcs = {rec[m]["source"] for m in members}
        secs = [rec[m]["secteur"] for m in members if rec[m]["secteur"]]
        sec_vote = pd.Series(secs).value_counts().idxmax() if secs else ""
        names = sorted({rec[m]["name"] for m in members})
        mn = ent_min_fuzzy.get(eid)
        if len(members) == 1:
            conf = "SINGLETON"
        elif ent_has_anchor[eid]:
            conf = "ANCRE_ID"
        elif mn is None or mn >= args.verify_band:
            conf = "FORT"
        else:
            conf = "A_VERIFIER"
        rows.append({
            "ID_UNIQUE": eid, "NOM_PRINCIPAL": max(names, key=len) if names else "",
            "N_MEMBRES": len(members), "N_BASES": len(srcs),
            "CONFIANCE_LIEN": conf,
            "SCORE_MIN_FUSION": round(mn, 3) if mn is not None else "",
            "PRESENCE_CRM": "✓" if "CRM" in srcs else "—",
            "PRESENCE_IRIS": "✓" if "IRIS" in srcs else "—",
            "PRESENCE_STATCOM": "✓" if "STATCOM" in srcs else "—",
            "PRESENCE_RUBRIKS": "✓" if "RUBRIKS" in srcs else "—",
            "SECTEUR_VOTE": sec_vote, "NOMS_VARIANTES": " | ".join(names[:6]),
        })
    ent = pd.DataFrame(rows).sort_values(["N_BASES", "N_MEMBRES"], ascending=False)
    alertes = ent[ent["N_MEMBRES"] > args.max_cluster]

    print(f"\n=== RÉSULTAT MAILLAGE (max fiabilité) ===")
    print(f"  Entités uniques            : {len(ent):,}")
    print(f"  Entités multi-bases (>=2)  : {(ent['N_BASES']>=2).sum():,}")
    print(f"    dont ancrées par ID exact: {(ent['CONFIANCE_LIEN']=='ANCRE_ID').sum():,}")
    print(f"    dont fuzzy fort          : {(ent['CONFIANCE_LIEN']=='FORT').sum():,}")
    print(f"    dont à vérifier          : {(ent['CONFIANCE_LIEN']=='A_VERIFIER').sum():,}")
    print(f"  Entités sur 4 bases        : {(ent['N_BASES']==4).sum():,}")
    print(f"  ⚠️ Grappes anormales (>{args.max_cluster}) : {len(alertes):,}")

    membres = pd.DataFrame([
        {"ID_UNIQUE": cid_to_eid[uf.find(uid)], "source": rec[uid]["source"],
         "name": rec[uid]["name"], "secteur": rec[uid]["secteur"]} for uid in rec])

    liens = pd.DataFrame([
        {"ID_UNIQUE": cid_to_eid[uf.find(a)], "TYPE": typ, "SCORE": sc,
         "NOM_A": rec[a]["name"], "NOM_B": rec[c]["name"],
         "BASE_A": rec[a]["source"], "BASE_B": rec[c]["source"]}
        for a, c, sc, typ in links])
    a_verif = liens[liens["TYPE"] == "A_VERIFIER"].sort_values("SCORE")

    with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
        ent.to_excel(xl, sheet_name="entites", index=False)
        membres.to_excel(xl, sheet_name="membres", index=False)
        liens.to_excel(xl, sheet_name="liens", index=False)
        a_verif.to_excel(xl, sheet_name="liens_a_verifier", index=False)
        alertes.to_excel(xl, sheet_name="grappes_alerte", index=False)
    print(f"  Worklist liens à vérifier  : {len(a_verif):,}")
    print(f"[ok] → {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
