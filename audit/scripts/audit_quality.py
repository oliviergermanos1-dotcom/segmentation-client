#!/usr/bin/env python3
"""Audit qualité par fichier Excel — produit un rapport markdown par fichier."""
import os
import re
import json
import unicodedata
from pathlib import Path
from collections import Counter, defaultdict

import openpyxl
from openpyxl.utils import get_column_letter

EXCEL_DIR = Path("/home/user/segmentation-client/audit/excel")
OUT_DIR = Path("/home/user/segmentation-client/audit/reports")

FILES = [
    "01_LOYERS.xlsx",
    "02_CHARGES.xlsx",
    "03_PLANNING.xlsx",
    "04_POINTAGE.xlsx",
    "05_CA.xlsx",
    "06_ACCOSTAGE.xlsx",
    "07_QHSE.xlsx",
    "08_ACTIONS.xlsx",
    "09_CA_MAI.xlsx",
    "10_STATS_BASE.xlsx",
]

# Patterns
ERROR_VALUES = {"#REF!", "#DIV/0!", "#N/A", "#NAME?", "#VALUE!", "#NULL!", "#NUM!"}
EXTERNAL_REF_RE = re.compile(r"\[(\d+)\]|'?\[[^\]]+\.xlsx?\]")
TIME_STR_RE = re.compile(r"^\s*\d{1,2}\s*[hH:]\s*\d{0,2}\s*$")

def normalize_label(s):
    if s is None:
        return ""
    s = str(s).strip()
    nfd = unicodedata.normalize("NFD", s)
    return "".join(c for c in nfd if unicodedata.category(c) != "Mn").lower()

def safe_cell_value(v):
    if v is None:
        return ""
    return str(v)

def analyze_file(path, sheet_limit=None):
    """Analyse complète d'un fichier. Renvoie un dict de stats."""
    result = {
        "path": str(path),
        "size_kb": round(path.stat().st_size / 1024, 1),
        "sheets": [],
        "n_sheets": 0,
        "n_cells_data": 0,
        "errors": [],            # [(sheet, cell, value)]
        "external_refs": [],     # [(sheet, cell, formula)]
        "merged_in_data": [],    # [(sheet, range)]
        "header_issues": [],     # [(sheet, issue)]
        "label_variants": defaultdict(set),  # normalized -> {variants}
        "time_strings": [],      # [(sheet, cell, value)]
        "trailing_space_cells": [],  # [(sheet, cell, value)]
        "text_in_number_col": [],
        "empty_rows_intercalated": [],  # [(sheet, n_empty_between_data)]
        "data_validations": [],         # sheets w/ DV
        "frozen_panes": [],
        "duplicate_sheets": [],
        "sheet_name_patterns": [],
        "formulas_count": 0,
        "rows_total": 0,
        "cols_total": 0,
    }

    # Pass 1: data_only=True for evaluated values (detects #REF! etc.)
    try:
        wb_data = openpyxl.load_workbook(path, data_only=True, read_only=False)
    except Exception as e:
        result["load_error"] = f"data_only=True: {e}"
        return result

    sheet_names = wb_data.sheetnames
    result["n_sheets"] = len(sheet_names)
    result["sheet_name_patterns"] = sheet_names[:]

    sheets_to_scan = sheet_names if sheet_limit is None else sheet_names[:sheet_limit]

    for sn in sheets_to_scan:
        ws = wb_data[sn]
        sheet_info = {
            "name": sn,
            "max_row": ws.max_row or 0,
            "max_col": ws.max_column or 0,
            "n_cells": 0,
            "n_empty_rows": 0,
            "header_row1": [],
        }
        # Iter limited region
        max_r = min(ws.max_row or 0, 2000)
        max_c = min(ws.max_column or 0, 60)
        result["rows_total"] += max_r
        result["cols_total"] = max(result["cols_total"], max_c)

        prev_empty_streak = 0
        data_started = False
        for r in range(1, max_r + 1):
            row_has_data = False
            for c in range(1, max_c + 1):
                v = ws.cell(row=r, column=c).value
                if v is None or (isinstance(v, str) and not v.strip()):
                    continue
                row_has_data = True
                sheet_info["n_cells"] += 1
                # error values
                if isinstance(v, str) and v.strip().upper() in ERROR_VALUES:
                    if len(result["errors"]) < 200:
                        result["errors"].append((sn, f"{get_column_letter(c)}{r}", v))
                # header row1 capture
                if r == 1 and isinstance(v, str):
                    sheet_info["header_row1"].append(v)
                # time strings
                if isinstance(v, str) and TIME_STR_RE.match(v):
                    if len(result["time_strings"]) < 50:
                        result["time_strings"].append((sn, f"{get_column_letter(c)}{r}", v))
                # trailing/leading spaces
                if isinstance(v, str) and v != v.strip() and len(result["trailing_space_cells"]) < 30:
                    result["trailing_space_cells"].append((sn, f"{get_column_letter(c)}{r}", repr(v)))
                # label variants in row 1
                if r == 1 and isinstance(v, str):
                    n = normalize_label(v)
                    if n:
                        result["label_variants"][n].add(v.strip())
            if row_has_data:
                if data_started and prev_empty_streak >= 1:
                    if len(result["empty_rows_intercalated"]) < 50:
                        result["empty_rows_intercalated"].append((sn, r, prev_empty_streak))
                data_started = True
                prev_empty_streak = 0
            else:
                if data_started:
                    prev_empty_streak += 1
                    sheet_info["n_empty_rows"] += 1

        # Merged cells
        for mr in ws.merged_cells.ranges:
            # ignore merges fully in row 1 (header) or column A header
            if mr.min_row > 1:
                if len(result["merged_in_data"]) < 80:
                    result["merged_in_data"].append((sn, str(mr)))

        # Frozen panes
        if ws.freeze_panes:
            result["frozen_panes"].append((sn, ws.freeze_panes))

        # Data validations
        try:
            dvs = ws.data_validations.dataValidation
            if dvs:
                result["data_validations"].append((sn, len(dvs)))
        except Exception:
            pass

        result["sheets"].append(sheet_info)
        result["n_cells_data"] += sheet_info["n_cells"]

    wb_data.close()

    # Pass 2: data_only=False for formulas (external refs, formula count)
    try:
        wb_f = openpyxl.load_workbook(path, data_only=False, read_only=True)
        for sn in sheets_to_scan:
            if sn not in wb_f.sheetnames:
                continue
            ws = wb_f[sn]
            max_r = min(ws.max_row or 0, 2000)
            max_c = min(ws.max_column or 0, 60)
            for r in range(1, max_r + 1):
                for c in range(1, max_c + 1):
                    v = ws.cell(row=r, column=c).value
                    if isinstance(v, str) and v.startswith("="):
                        result["formulas_count"] += 1
                        if EXTERNAL_REF_RE.search(v) and len(result["external_refs"]) < 50:
                            result["external_refs"].append((sn, f"{get_column_letter(c)}{r}", v[:200]))
        wb_f.close()
    except Exception as e:
        result["formula_scan_error"] = str(e)

    # Duplicate sheets — heuristic by name pattern
    name_norm = [normalize_label(s) for s in sheet_names]
    dup_counter = Counter(name_norm)
    result["duplicate_sheets"] = [(n, c) for n, c in dup_counter.items() if c > 1]

    return result


def detect_sheet_naming_inconsistency(names):
    """Detect inconsistent month naming styles."""
    styles = set()
    month_pat_fr = re.compile(r"(janv|fevr|mars|avr|mai|juin|juil|aout|sept|oct|nov|dec)", re.I)
    iso_pat = re.compile(r"^\d{2}-\d{2}$")
    long_pat = re.compile(r"\b(JANVIER|FEVRIER|MARS|AVRIL|MAI|JUIN|JUILLET|AOUT|SEPTEMBRE|OCTOBRE|NOVEMBRE|DECEMBRE)\b", re.I)
    short_pat = re.compile(r"\b(janv|fevr|mars|avr|mai|juin|juil|aout|sept|oct|nov|dec)[-./ ]?\d{2,4}\b", re.I)
    for n in names:
        if iso_pat.match(n.strip()):
            styles.add("iso (MM-YY)")
        elif long_pat.search(n):
            styles.add("long FR (JANVIER 26)")
        elif short_pat.search(n):
            styles.add("court (janv-26)")
        elif month_pat_fr.search(n):
            styles.add("autre fr")
    return styles


def quality_grade(r):
    """Note 🟢/🟠/🔴 selon critères."""
    score = 100
    critical = 0
    if r.get("errors"):
        score -= min(30, len(r["errors"]))
        if len(r["errors"]) >= 5:
            critical += 1
    if r.get("external_refs"):
        score -= 15
        critical += 1
    if len(r.get("merged_in_data", [])) > 5:
        score -= 15
        critical += 1
    if len(r.get("time_strings", [])) > 5:
        score -= 5
    if len(r.get("trailing_space_cells", [])) > 0:
        score -= 5
    # variants of labels (different casing/accent for same field)
    var_count = sum(1 for v in r.get("label_variants", {}).values() if len(v) > 1)
    if var_count:
        score -= min(15, var_count * 3)
    if len(r.get("empty_rows_intercalated", [])) > 10:
        score -= 10
    if not r.get("frozen_panes"):
        score -= 3
    if not r.get("data_validations"):
        score -= 5

    if score >= 80 and critical == 0:
        return "🟢", score
    if score >= 55:
        return "🟠", score
    return "🔴", score


def write_report(filename, r):
    grade, score = quality_grade(r)
    n_err = len(r["errors"])
    n_ext = len(r["external_refs"])
    n_merge = len(r["merged_in_data"])
    n_time = len(r["time_strings"])
    n_space = len(r["trailing_space_cells"])
    n_empty = len(r["empty_rows_intercalated"])
    n_dv = len(r["data_validations"])
    n_frozen = len(r["frozen_panes"])
    n_dup = sum(1 for _, c in r["duplicate_sheets"] if c > 1)

    naming_styles = detect_sheet_naming_inconsistency(r["sheet_name_patterns"])

    out = []
    out.append(f"# Audit qualité — {filename}")
    out.append("**Date** : 2026-06-03")
    out.append("")
    out.append("## Vue d'ensemble")
    out.append(f"- Taille : {r['size_kb']} Ko")
    out.append(f"- Onglets : {r['n_sheets']}")
    out.append(f"- Cellules de données estimées : {r['n_cells_data']:,}".replace(",", " "))
    out.append(f"- Formules : {r['formulas_count']:,}".replace(",", " "))
    out.append(f"- Note qualité globale : {grade} ({score}/100)")
    out.append("")

    # Critical
    out.append("## 🔴 Problèmes critiques (bloquants pour parsing)")
    crit = []
    if n_err:
        ex = ", ".join(f"`{s}!{c}`" for s, c, _ in r["errors"][:3])
        crit.append(f"- **{n_err} cellules d'erreur** (#REF!, #DIV/0!, #N/A) — ex : {ex}")
    if n_ext:
        ex = ", ".join(f"`{s}!{c}`" for s, c, _ in r["external_refs"][:3])
        crit.append(f"- **{n_ext} formules avec références externes** (cassent à l'export) — ex : {ex}")
    if n_merge > 5:
        ex = ", ".join(f"`{s}!{rng}`" for s, rng in r["merged_in_data"][:3])
        crit.append(f"- **{n_merge} plages fusionnées dans la zone de données** (hors ligne 1) — ex : {ex}")
    # multi-row headers heuristic
    multi_header = 0
    for sh in r["sheets"]:
        if sh["max_row"] >= 2 and sh.get("header_row1"):
            # if row 1 contains few non-null vs row 2 has more, header likely multi-row
            pass
    if r.get("duplicate_sheets") and any(c > 1 for _, c in r["duplicate_sheets"]):
        dups = ", ".join(f"`{n}` x{c}" for n, c in r["duplicate_sheets"] if c > 1)
        crit.append(f"- **Onglets dupliqués (ou collisions de normalisation)** : {dups}")
    if not crit:
        crit.append("- Aucun problème critique détecté.")
    out.extend(crit)
    out.append("")

    # Medium
    out.append("## 🟠 Problèmes moyens (qualité dégradée)")
    med = []
    var_dupes = [(k, sorted(v)) for k, v in r["label_variants"].items() if len(v) > 1]
    if var_dupes:
        ex = "; ".join(f"« {' / '.join(vs)} »" for _, vs in var_dupes[:4])
        med.append(f"- **{len(var_dupes)} en-têtes avec variantes de casse/accent** : {ex}")
    if n_time:
        ex = ", ".join(f"`{s}!{c}`='{v}'" for s, c, v in r["time_strings"][:3])
        med.append(f"- **{n_time} cellules horaires en texte** (ex : '11H45' au lieu d'un type Heure) — {ex}")
    if n_space:
        ex = ", ".join(f"`{s}!{c}`={v}" for s, c, v in r["trailing_space_cells"][:3])
        med.append(f"- **{n_space} cellules avec espaces parasites** (début/fin) — {ex}")
    if n_empty > 5:
        med.append(f"- **{n_empty} ruptures par lignes vides** intercalées entre lignes de données (rend la détection de plage fragile)")
    if naming_styles and len(naming_styles) > 1:
        med.append(f"- **Mélange de styles de nommage d'onglets** : {', '.join(sorted(naming_styles))}")
    if not med:
        med.append("- Aucun problème moyen significatif.")
    out.extend(med)
    out.append("")

    # Comfort
    out.append("## 🟢 Points d'amélioration (confort utilisateur)")
    com = []
    if not r["data_validations"]:
        com.append("- **Aucune validation Excel** (listes déroulantes Client/Statut/Mode Paiement absentes) : risque de fautes de frappe et nouvelles valeurs non normées.")
    else:
        com.append(f"- Validations présentes sur {n_dv} onglet(s) — étendre aux autres.")
    if not r["frozen_panes"]:
        com.append("- **Aucune cellule figée** : navigation fastidieuse dans les longs tableaux.")
    else:
        com.append(f"- Cellules figées présentes sur {n_frozen} onglet(s) — bon point à généraliser.")
    com.append("- Pas d'onglet « Mode d'emploi / Légende » documentant les colonnes, conventions et statuts.")
    com.append("- Pas de feuille « Référentiels » pour les listes (clients, statuts, types) — chaque saisie ressaisit le texte.")
    out.extend(com)
    out.append("")

    # Duplicates & redundancies
    out.append("## 📊 Doublons et redondances")
    if r["duplicate_sheets"] and any(c > 1 for _, c in r["duplicate_sheets"]):
        for n, c in r["duplicate_sheets"]:
            if c > 1:
                out.append(f"- Nom normalisé `{n}` apparaît {c} fois")
    if r["n_sheets"] >= 10:
        out.append(f"- Présence d'onglets historiques/archives multiples ({r['n_sheets']}) — risque de redondance et de versions concurrentes.")
    if not r["duplicate_sheets"] or not any(c > 1 for _, c in r["duplicate_sheets"]):
        out.append("- Pas de doublon d'onglet flagrant détecté (au niveau du nom normalisé).")
    out.append("")

    # Structure & naming
    out.append("## 🧱 Structure et nommage")
    out.append(f"- Convention de nommage : { 'cohérente' if len(naming_styles) <= 1 else 'INCOHÉRENTE — ' + ', '.join(sorted(naming_styles)) }")
    sample = r["sheet_name_patterns"][:8]
    out.append(f"- Échantillon d'onglets : {', '.join(repr(s) for s in sample)}" + (" …" if r["n_sheets"] > 8 else ""))
    if var_dupes:
        out.append(f"- Libellés de colonnes : {len(var_dupes)} variantes détectées (casse/accents).")
    out.append("")

    # Recommandations
    out.append("## 🎯 Recommandations prioritaires (top 5)")
    recs = []
    if n_err:
        recs.append(f"Corriger les {n_err} cellules d'erreur (#REF!/#DIV/0!/#N/A) avant tout parsing automatisé.")
    if n_ext:
        recs.append(f"Remplacer les {n_ext} formules à références externes par des valeurs ou liens internes.")
    if n_merge > 5:
        recs.append(f"Démerger les {n_merge} plages fusionnées de la zone de données (1 valeur = 1 cellule).")
    if var_dupes:
        recs.append("Harmoniser les en-têtes (casse + accents) et figer une convention unique par feuille.")
    if n_time:
        recs.append("Convertir les cellules horaires texte (« 11H45 ») en format Heure (hh:mm) pour calculs.")
    if not r["data_validations"]:
        recs.append("Ajouter des validations Excel (listes déroulantes) sur Client/Statut/Type/Mode pour éviter les variantes.")
    if not r["frozen_panes"]:
        recs.append("Figer la ligne d'en-tête (View > Freeze Panes) sur chaque feuille de saisie.")
    if naming_styles and len(naming_styles) > 1:
        recs.append("Unifier le nommage des onglets mensuels (recommandé : `YYYY-MM` ISO).")
    if n_empty > 5:
        recs.append("Supprimer les lignes vides intercalées qui faussent la détection des tableaux.")
    recs.append("Ajouter un onglet « Mode d'emploi » documentant colonnes, statuts et exceptions.")
    for i, r_ in enumerate(recs[:5], 1):
        out.append(f"{i}. {r_}")
    out.append("")

    # Annexe
    out.append("## Annexe : exemples concrets")
    samples = []
    for s, c, v in r["errors"][:3]:
        samples.append(f"- `{s}!{c}` = `{v}` → recalculer la formule ou remplacer par une valeur figée.")
    for s, c, f in r["external_refs"][:2]:
        f_short = f[:100] + ("…" if len(f) > 100 else "")
        samples.append(f"- `{s}!{c}` = `{f_short}` → références externe : casser le lien.")
    for s, rng in r["merged_in_data"][:3]:
        samples.append(f"- `{s}!{rng}` fusionnée → démerger et propager la valeur.")
    for s, c, v in r["time_strings"][:2]:
        samples.append(f"- `{s}!{c}` = `'{v}'` (texte) → convertir en Heure native.")
    for s, c, v in r["trailing_space_cells"][:2]:
        samples.append(f"- `{s}!{c}` = {v} → TRIM().")
    if var_dupes:
        for k, vs in var_dupes[:3]:
            samples.append(f"- En-tête `{k}` apparaît sous : {', '.join(repr(x) for x in vs)} → harmoniser.")
    if not samples:
        samples.append("- Aucun cas problématique précis à signaler.")
    out.extend(samples[:8])
    out.append("")
    return "\n".join(out), grade, score, {
        "errors": n_err, "ext": n_ext, "merge": n_merge,
        "time": n_time, "space": n_space, "empty": n_empty,
        "var": len(var_dupes), "dv": n_dv, "frozen": n_frozen,
        "naming_styles": len(naming_styles),
    }


def main():
    summary_rows = []
    transverse = Counter()
    total_size = 0

    for fname in FILES:
        path = EXCEL_DIR / fname
        if not path.exists():
            print(f"missing: {path}")
            continue
        print(f"analyzing {fname} ...")
        # Limit for large files
        limit = None
        if path.stat().st_size > 5_000_000:
            limit = 12
        r = analyze_file(path, sheet_limit=limit)
        if "load_error" in r:
            print(f"  load error: {r['load_error']}")
            continue
        report, grade, score, stats = write_report(fname, r)
        out_path = OUT_DIR / f"QUALITE_{fname.replace('.xlsx', '')}.md"
        out_path.write_text(report, encoding="utf-8")
        total_size += out_path.stat().st_size
        summary_rows.append({
            "file": fname,
            "size_kb": r["size_kb"],
            "n_sheets": r["n_sheets"],
            "grade": grade,
            "score": score,
            "stats": stats,
            "label_variants_count": stats["var"],
            "errors": stats["errors"],
            "external": stats["ext"],
            "merge": stats["merge"],
            "time": stats["time"],
            "space": stats["space"],
            "empty": stats["empty"],
            "naming_styles": stats["naming_styles"],
        })

        # transverse counters
        if stats["errors"]:
            transverse["cellules d'erreur (#REF!/#DIV!/#N/A)"] += 1
        if stats["ext"]:
            transverse["références externes dans formules"] += 1
        if stats["merge"] > 5:
            transverse["plages fusionnées en zone de données"] += 1
        if stats["var"]:
            transverse["en-têtes variant en casse/accent"] += 1
        if stats["time"]:
            transverse["horaires en texte (ex '11H45')"] += 1
        if stats["space"]:
            transverse["espaces parasites en cellules"] += 1
        if stats["empty"] > 5:
            transverse["lignes vides intercalées"] += 1
        if not stats["dv"]:
            transverse["absence de validations / listes déroulantes"] += 1
        if not stats["frozen"]:
            transverse["pas de cellule figée"] += 1
        if stats["naming_styles"] > 1:
            transverse["nommage d'onglets incohérent"] += 1

    # Synthèse
    s = []
    s.append("# Synthèse — Audit qualité des 10 fichiers Excel")
    s.append("**Date** : 2026-06-03")
    s.append("")
    s.append("## Tableau récapitulatif")
    s.append("")
    s.append("| Fichier | Taille (Ko) | Onglets | Note | Score | Top 3 problèmes |")
    s.append("|---|---:|---:|:--:|---:|---|")
    for row in summary_rows:
        st = row["stats"]
        top = []
        candidates = [
            (st["errors"], f"{st['errors']} cellules d'erreur"),
            (st["ext"], f"{st['ext']} réfs externes"),
            (st["merge"], f"{st['merge']} merges en zone data"),
            (st["var"], f"{st['var']} en-têtes en variantes"),
            (st["time"], f"{st['time']} horaires texte"),
            (st["empty"], f"{st['empty']} ruptures vides"),
            (st["space"], f"{st['space']} cellules espacées"),
        ]
        candidates.sort(reverse=True)
        for n, label in candidates[:3]:
            if n > 0:
                top.append(label)
        if not top:
            top.append("RAS notable")
        s.append(f"| `{row['file']}` | {row['size_kb']} | {row['n_sheets']} | {row['grade']} | {row['score']} | {' ; '.join(top)} |")
    s.append("")

    s.append("## Problèmes transverses (fréquence par fichier)")
    s.append("")
    s.append("| Problème | Fichiers concernés (/10) |")
    s.append("|---|---:|")
    for k, c in transverse.most_common(12):
        s.append(f"| {k} | {c} |")
    s.append("")

    # Roadmap
    s.append("## Roadmap qualité priorisée (impact ↗)")
    s.append("")
    s.append("1. **Casser les références externes & corriger les erreurs** — bloquantes pour tout export et tout parsing : à faire avant l'industrialisation.")
    s.append("2. **Démerger la zone de données** sur les fichiers à fort taux de fusion (PLANNING, ACCOSTAGE, QHSE) : 1 valeur = 1 cellule.")
    s.append("3. **Harmoniser les en-têtes** (casse + accents) avec une convention unique par feuille — base pour le matching des templates stricts.")
    s.append("4. **Unifier le nommage des onglets** mensuels (`YYYY-MM`) — élimine la confusion entre 'JANVIER 26' / 'janv-26' / '01-26'.")
    s.append("5. **Ajouter des validations Excel** (listes déroulantes) sur les colonnes énumérées (Client, Statut, Mode Paiement, Type) à partir d'un onglet « Référentiels ».")
    s.append("6. **Normaliser les horaires** au type Heure natif (`hh:mm`) au lieu des chaînes `11H45`.")
    s.append("7. **Figer la ligne d'en-tête** sur chaque feuille de saisie (View > Freeze Panes).")
    s.append("8. **Supprimer les lignes vides intercalées** et bannir les sous-totaux mêlés aux lignes de données.")
    s.append("9. **Ajouter un onglet « Mode d'emploi »** : colonnes attendues, statuts admis, exceptions, contact référent.")
    s.append("10. **Archiver les anciens onglets** dans un fichier séparé pour ne pas polluer la détection automatique.")
    s.append("")

    # Worst 3
    worst = sorted(summary_rows, key=lambda x: x["score"])[:3]
    s.append("## Top 3 des fichiers les plus problématiques")
    s.append("")
    for w in worst:
        s.append(f"- **{w['file']}** — {w['grade']} ({w['score']}/100) : "
                 f"{w['errors']} erreurs, {w['external']} réfs externes, {w['merge']} merges data, {w['label_variants_count']} en-têtes variantes.")
    s.append("")

    syn_path = OUT_DIR / "QUALITE_SYNTHESE.md"
    syn_path.write_text("\n".join(s), encoding="utf-8")
    total_size += syn_path.stat().st_size

    print(f"\nTotal: {total_size/1024:.1f} Ko")
    print("Worst 3:")
    for w in worst:
        print(f"  {w['file']}: {w['grade']} {w['score']}/100")

    return summary_rows, worst, total_size


if __name__ == "__main__":
    main()
