#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — 12 : Snapshot horodaté d'une exécution du pipeline
================================================================
Copie les sorties d'un `run_all.py` vers un dossier daté
`snapshots/AAAA-MM-JJ/` avec un manifest JSON. Cela permet de
construire un historique multi-périodes exploitable par 13_compare.

Usage :
    python 12_snapshot.py --src ./out --snapshots ./snapshots
    python 12_snapshot.py --src ./out --snapshots ./snapshots --label "2026-06_PFA-mois"
"""
from __future__ import annotations
import argparse
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path


# Artefacts ESSENTIELS pour comparer 2 périodes — copiés systématiquement.
ARTEFACTS_REQUIRED = [
    "RMC.xlsx",
    "agreg/marche_total.csv",
    "agreg/volume_client.csv",
    "agreg/pdm.csv",
    "agreg/projections.csv",
    "agreg/whitespaces.csv",
]
ARTEFACTS_OPTIONAL = [
    "agreg/metiers_3ans.csv",
    "agreg/pdm_clients.csv",
    "agreg/macro_multipliers.json",
    "agreg/budget_vs_real.csv",
    "agreg/budget_secteur.csv",
    "agreg/projets_pnd.csv",
    "agreg/axes_pnd.csv",
    "agreg/sectoriels_pnd.csv",
    "match_crm_iris.csv",
    "match_crm_stat.csv",
    "match_iris_stat.csv",
]


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="AGL BUDGET — 12 snapshot d'une exécution")
    ap.add_argument("--src",       required=True, help="Dossier --out de run_all.py")
    ap.add_argument("--snapshots", required=True, help="Dossier des snapshots historisés")
    ap.add_argument("--label",     default=None,  help="Suffixe optionnel (ex: '2026-06_PFA-mois')")
    args = ap.parse_args(argv)

    src = Path(args.src); root = Path(args.snapshots)
    if not src.exists():
        sys.exit(f"ERREUR : --src introuvable : {src}")

    name = datetime.now().strftime("%Y-%m-%d_%H%M")
    if args.label: name = f"{name}_{args.label}"
    dest = root / name
    dest.mkdir(parents=True, exist_ok=False)

    copied, missing = [], []
    for f in ARTEFACTS_REQUIRED:
        p = src / f
        if p.exists():
            target = dest / f
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(p, target)
            copied.append(f)
        else:
            missing.append(f)
    for f in ARTEFACTS_OPTIONAL:
        p = src / f
        if p.exists():
            target = dest / f
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(p, target)
            copied.append(f)

    if missing:
        sys.stderr.write(f"[warn] {len(missing)} artefacts requis manquants : "
                         f"{', '.join(missing)}\n"
                         "       Le snapshot reste utilisable mais incomplet.\n")

    manifest = {
        "snapshot": name,
        "source": str(src.resolve()),
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "label": args.label or "",
        "n_files": len(copied),
        "files": copied,
        "missing_required": missing,
    }
    with open(dest / "manifest.json", "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)

    print(f"[ok] snapshot écrit : {dest}")
    print(f"     {len(copied)} fichiers copiés"
          + (f", {len(missing)} requis manquants" if missing else ""))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
