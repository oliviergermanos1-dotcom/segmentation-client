#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Orchestrateur du pipeline complet
==============================================
Enchaîne 01 → 09 en une commande. Conçu pour la routine mensuelle de la
Direction Stratégie.

Conventions de nommage des entrées dans --data-dir :
  CRM.xlsx | IRIS.xlsx | STATCOM*.xlsx | RUBRIKS.xlsx
(STATCOM* peut être un fichier consolidé ou plusieurs — premier match retenu.)

Toutes les sorties vont dans --out (par défaut : ./out/).
Chaque étape est skippable via --skip 06,07.

Usage :
    python run_all.py --data-dir ./fixtures --out ./out
    python run_all.py --data-dir ./data --out ./out --skip 08,09
"""
from __future__ import annotations
import argparse
import importlib.util
import sys
import time
from pathlib import Path

ROOT = Path(__file__).parent


def _load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _find(data_dir: Path, pattern: str):
    for p in sorted(data_dir.glob(pattern)):
        return p
    return None


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AGL BUDGET — orchestrateur pipeline 01→09")
    ap.add_argument("--data-dir", required=True, help="Dossier contenant CRM.xlsx / IRIS.xlsx / STATCOM*.xlsx / RUBRIKS.xlsx")
    ap.add_argument("--out",      default="out",  help="Dossier de sortie (créé si absent)")
    ap.add_argument("--skip",     default="",     help="Étapes à passer, ex: 08,09")
    ap.add_argument("--horizon",  type=int, default=2030, help="Horizon projections (défaut 2030)")
    args = ap.parse_args(argv)

    data = Path(args.data_dir)
    out  = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    skip = {s.strip().zfill(2) for s in args.skip.split(",") if s.strip()}

    crm     = _find(data, "CRM*.xlsx")
    iris    = _find(data, "IRIS*.xlsx")
    statcom = _find(data, "STATCOM*.xlsx")
    rubriks = _find(data, "RUBRIKS*.xlsx")
    for name, p in [("CRM", crm), ("IRIS", iris), ("STATCOM", statcom)]:
        if not p: sys.exit(f"ERREUR : {name} introuvable dans {data} (attendu : {name}*.xlsx)")

    # Charge les modules par chemin (les noms commencent par un chiffre).
    norm = _load("norm", ROOT / "01_normalise.py")
    m02  = _load("m02",  ROOT / "02_match_crm_iris.py")
    m03  = _load("m03",  ROOT / "03_match_crm_stat.py")
    m04  = _load("m04",  ROOT / "04_match_iris_stat.py")
    m05  = _load("m05",  ROOT / "05_build_rmc.py")
    m06  = _load("m06",  ROOT / "06_agreg_statcom.py")
    m07  = _load("m07",  ROOT / "07_calcul_pdm.py")
    m08  = _load("m08",  ROOT / "08_projections.py")
    m09  = _load("m09",  ROOT / "09_whitespaces.py")
    m11  = _load("m11",  ROOT / "11_budget_real.py")

    def step(num, name, fn):
        if num in skip:
            print(f"[skip] {num} — {name}")
            return
        t0 = time.time()
        print(f"\n=== {num} — {name} ===")
        rc = fn()
        dt = time.time() - t0
        if rc and rc != 0:
            sys.exit(f"[STOP] {num} a renvoyé {rc}")
        print(f"[done] {num} en {dt:.1f}s")

    crm_norm  = out / "CRM_norm.xlsx"
    iris_norm = out / "IRIS_norm.xlsx"
    stat_norm = out / "STATCOM_norm.xlsx"
    m_ci      = out / "match_crm_iris.csv"
    m_cs      = out / "match_crm_stat.csv"
    m_is      = out / "match_iris_stat.csv"
    rmc       = out / "RMC.xlsx"
    agreg     = out / "agreg"

    step("01a", "normalise CRM",  lambda: norm.main(["-i", str(crm),  "-c", "nom_du_compte", "-o", str(crm_norm)]))
    step("01b", "normalise IRIS", lambda: norm.main(["-i", str(iris), "-c", "client",         "-o", str(iris_norm)]))
    step("02",  "match CRM x IRIS",
         lambda: m02.main(["--crm", str(crm_norm), "--iris", str(iris_norm), "--out", str(m_ci)]))
    step("03",  "match CRM x STATCOM",
         lambda: m03.main(["--crm", str(crm_norm), "--statcom", str(statcom),
                           "--out", str(m_cs), "--stat-norm-out", str(stat_norm)]))
    step("04",  "match IRIS x STATCOM (pont)",
         lambda: m04.main(["--iris", str(iris_norm), "--statcom", str(stat_norm), "--out", str(m_is)]))
    rubriks_args = ["--rubriks", str(rubriks)] if rubriks else []
    step("05",  "build RMC",
         lambda: m05.main(["--crm", str(crm_norm), "--m12", str(m_ci), "--m13", str(m_cs),
                           "--m23", str(m_is), *rubriks_args, "--out", str(rmc)]))
    step("06",  "agrégats STATCOM 3 ans",
         lambda: m06.main(["--statcom", str(stat_norm), "--outdir", str(agreg)]))
    step("07",  "PDM AGL",
         lambda: m07.main(["--rmc", str(rmc),
                           "--volume-client", str(agreg / "volume_client.csv"),
                           "--marche-total",  str(agreg / "marche_total.csv"),
                           "--outdir", str(agreg)]))
    step("08",  "projections 2026-" + str(args.horizon),
         lambda: m08.main(["--metiers", str(agreg / "metiers_3ans.csv"),
                           "--outdir", str(agreg), "--horizon", str(args.horizon)]))
    step("09",  "white spaces",
         lambda: m09.main(["--rmc", str(rmc),
                           "--volume-client", str(agreg / "volume_client.csv"),
                           "--match-crm-stat", str(m_cs),
                           "--outdir", str(agreg)]))
    if rubriks:
        step("11",  "budget PFA vs réel",
             lambda: m11.main(["--rmc", str(rmc), "--iris", str(iris_norm),
                               "--rubriks", str(rubriks), "--outdir", str(agreg)]))

    print(f"\n[ok] PIPELINE COMPLET TERMINÉ → {out}")
    print("Charger ces artefacts dans l'app HTML (onglet 01 → zone pipeline) :")
    for f in sorted([rmc, agreg / "marche_total.csv", agreg / "volume_client.csv",
                     agreg / "metiers_3ans.csv", agreg / "pdm.csv",
                     agreg / "projections.csv", agreg / "whitespaces.csv",
                     m_ci, m_cs]):
        if f.exists():
            print(f"    {f.relative_to(out.parent) if out.parent in f.parents else f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
