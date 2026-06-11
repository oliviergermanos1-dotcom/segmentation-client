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


def _prep(path, source, name_col, sec_col):
    if not path:
        return None
    df = pd.read_excel(path)
    nm = name_col if name_col in df.columns else None
    if not nm:
        return None
    out = pd.DataFrame()
    out["name"] = df[nm].fillna("").astype(str).str.upper().str.strip()
    out["secteur"] = (df[sec_col].fillna("").astype(str).str.upper().str.strip()
                      if sec_col and sec_col in df.columns else "")
    out = out[out["name"] != ""]
    # client unique : secteur dominant par nom
    def _mode_sec(s):
        nz = s[s != ""]
        return nz.mode().iloc[0] if len(nz) else ""
    out = out.groupby("name", as_index=False).agg(secteur=("secteur", _mode_sec))
    out["source"] = source
    out["uid"] = [f"{source}_{i}" for i in range(len(out))]
    return out


def _blocks(name, prefix=4):
    toks = [t for t in name.split() if len(t) >= 3][:4]
    return {t[:prefix] for t in toks} or ({name[:prefix]} if name else set())


def _score(a, b):
    tsr = fuzz.token_set_ratio(a, b) / 100.0
    if tsr < TOKEN_FLOOR:
        return 0.0
    jw = distance.JaroWinkler.normalized_similarity(a, b)
    return 0.85 * tsr + 0.15 * jw


class UF:
    def __init__(self): self.p = {}
    def find(self, x):
        self.p.setdefault(x, x)
        while self.p[x] != x:
            self.p[x] = self.p[self.p[x]]; x = self.p[x]
        return x
    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb: self.p[ra] = rb


def main(argv=None):
    ap = argparse.ArgumentParser(__doc__)
    ap.add_argument("--crm"); ap.add_argument("--iris")
    ap.add_argument("--statcom"); ap.add_argument("--rubriks")
    ap.add_argument("--name-col", default="NOM_BASE")
    ap.add_argument("--sec-col", default="secteur")
    ap.add_argument("--threshold", type=float, default=0.90)
    ap.add_argument("--max-cluster", type=int, default=10)
    ap.add_argument("--out", default="entites_graphe.xlsx")
    args = ap.parse_args(argv)

    parts = []
    for src, p in [("CRM", args.crm), ("IRIS", args.iris),
                   ("STATCOM", args.statcom), ("RUBRIKS", args.rubriks)]:
        t = _prep(p, src, args.name_col, args.sec_col)
        if t is not None and len(t):
            parts.append(t); print(f"  {src:8s} : {len(t):,} clients uniques")
    pool = pd.concat(parts, ignore_index=True)
    print(f"  POOL     : {len(pool):,} enregistrements")

    # Index blocking
    idx = defaultdict(list)
    rec = pool.set_index("uid")[["name", "source", "secteur"]].to_dict("index")
    for uid, r in rec.items():
        for b in _blocks(r["name"]):
            idx[b].append(uid)

    # Scoring intra-bloc + union-find
    print("[match] scoring des paires candidates…")
    uf = UF()
    for uid in rec: uf.find(uid)
    seen = set()
    n_links = 0
    for b, uids in idx.items():
        if len(uids) < 2:
            continue
        for i in range(len(uids)):
            for j in range(i + 1, len(uids)):
                a, c = uids[i], uids[j]
                key = (a, c) if a < c else (c, a)
                if key in seen:
                    continue
                seen.add(key)
                if _score(rec[a]["name"], rec[c]["name"]) >= args.threshold:
                    uf.union(a, c); n_links += 1
    print(f"[match] {n_links:,} liens retenus (>= {args.threshold})")

    # Grappes
    groups = defaultdict(list)
    for uid in rec:
        groups[uf.find(uid)].append(uid)

    rows = []
    for cid, members in groups.items():
        srcs = {rec[m]["source"] for m in members}
        secs = [rec[m]["secteur"] for m in members if rec[m]["secteur"]]
        sec_vote = pd.Series(secs).value_counts().idxmax() if secs else ""
        names = sorted({rec[m]["name"] for m in members})
        rows.append({
            "ID_UNIQUE": f"E{abs(hash(cid)) % 10**8:08d}",
            "NOM_PRINCIPAL": max(names, key=len) if names else "",
            "N_MEMBRES": len(members),
            "N_BASES": len(srcs),
            "PRESENCE_CRM": "✓" if "CRM" in srcs else "—",
            "PRESENCE_IRIS": "✓" if "IRIS" in srcs else "—",
            "PRESENCE_STATCOM": "✓" if "STATCOM" in srcs else "—",
            "PRESENCE_RUBRIKS": "✓" if "RUBRIKS" in srcs else "—",
            "SECTEUR_VOTE": sec_vote,
            "NOMS_VARIANTES": " | ".join(names[:6]),
        })
    ent = pd.DataFrame(rows).sort_values(["N_BASES", "N_MEMBRES"], ascending=False)
    alertes = ent[ent["N_MEMBRES"] > args.max_cluster]

    print(f"\n=== RÉSULTAT MAILLAGE (déterministe) ===")
    print(f"  Entités uniques            : {len(ent):,}")
    print(f"  Entités multi-bases (>=2)  : {(ent['N_BASES']>=2).sum():,}")
    print(f"  Entités sur 3 bases        : {(ent['N_BASES']==3).sum():,}")
    print(f"  Entités sur 4 bases        : {(ent['N_BASES']==4).sum():,}")
    print(f"  ⚠️ Grappes à vérifier (>{args.max_cluster}) : {len(alertes):,}")

    membres = pd.DataFrame([
        {"ID_UNIQUE": f"E{abs(hash(uf.find(uid))) % 10**8:08d}",
         "source": rec[uid]["source"], "name": rec[uid]["name"],
         "secteur": rec[uid]["secteur"]}
        for uid in rec])

    with pd.ExcelWriter(args.out, engine="openpyxl") as xl:
        ent.to_excel(xl, sheet_name="entites", index=False)
        membres.to_excel(xl, sheet_name="membres", index=False)
        alertes.to_excel(xl, sheet_name="grappes_alerte", index=False)
    print(f"\n[ok] → {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
