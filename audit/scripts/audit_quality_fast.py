"""Audit qualité rapide — version optimisée pour les fichiers gros."""
import openpyxl
from openpyxl.utils import get_column_letter
from pathlib import Path
import re
import unicodedata

EXCEL_DIR = Path(__file__).parent.parent / "excel"
OUT_DIR = Path(__file__).parent.parent / "reports"

# Pour les gros fichiers : limiter les itérations
LIMIT_ROWS = 200  # lignes scannées par feuille pour analyse approfondie
LIMIT_SHEETS = 12  # nb max de feuilles scannées

def norm_str(s):
    if s is None: return ""
    return str(s).strip()

def has_accents_or_case_issues(label):
    """Détecte 'Libéllé' (faute), accents incohérents, casse mélangée."""
    if not label or len(label) > 50: return False
    issues = []
    if re.search(r'Libéllé', label): issues.append('typo')
    return issues

def is_excel_error(v):
    s = str(v) if v is not None else ""
    return any(err in s for err in ['#REF!', '#DIV/0!', '#N/A', '#NAME?', '#NUM!', '#VALUE!'])

def audit_file(path: Path):
    name = path.name
    rep = [f"# Audit qualité — {name}", "**Date** : 2026-06-03", ""]
    rep.append("## Vue d'ensemble")
    size_kb = path.stat().st_size / 1024
    rep.append(f"- Taille : {size_kb:.1f} Ko")

    try:
        # data_only=True pour récupérer les valeurs calculées (et détecter #REF!)
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        sheets = wb.sheetnames
    except Exception as e:
        rep.append(f"\n**ERREUR chargement** : {e}\n")
        return "\n".join(rep)

    rep.append(f"- Onglets : {len(sheets)}")

    # Comptages globaux
    nb_data_cells = 0
    nb_formulas = 0  # difficile à compter sans data_only=False
    errors_locations = []  # (sheet, cell, type)
    accent_issues = []  # (sheet, cell, label, issue)
    empty_sheets = []
    archive_sheets = []
    inconsistent_headers = {}  # normalized -> set of variants
    merged_in_data = []  # (sheet, range)
    sheets_to_scan = sheets[:LIMIT_SHEETS]
    sheet_dims = []

    for sn in sheets_to_scan:
        ws = wb[sn]
        if not ws.max_row or not ws.max_column:
            empty_sheets.append(sn)
            continue
        sheet_dims.append((sn, ws.max_row, ws.max_column))

        # Détection onglet archive
        if re.search(r'\b(programme|soa|archive|ancien|2014|2015|2016|2017|2018|2019|2020|2021|2022)\b', sn, re.I):
            archive_sheets.append(sn)

        # Scan échantillon des cellules
        rows_scanned = 0
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=min(LIMIT_ROWS, ws.max_row), values_only=False), 1):
            for cell in row:
                if cell.value is None or cell.value == "":
                    continue
                nb_data_cells += 1
                if is_excel_error(cell.value):
                    if len(errors_locations) < 30:
                        errors_locations.append((sn, cell.coordinate, str(cell.value)[:20]))
                if isinstance(cell.value, str):
                    issues = has_accents_or_case_issues(cell.value)
                    if issues and len(accent_issues) < 20:
                        accent_issues.append((sn, cell.coordinate, cell.value[:40], ",".join(issues)))
            rows_scanned += 1
            if rows_scanned >= LIMIT_ROWS: break

        # En-têtes (ligne 1)
        try:
            hdr_row = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))
            if hdr_row:
                for v in hdr_row[0]:
                    if v is None: continue
                    s = str(v).strip()
                    if not s or len(s) > 50: continue
                    n = unicodedata.normalize('NFD', s.lower()).encode('ascii', 'ignore').decode('ascii').strip()
                    n = re.sub(r'\s+', ' ', n)
                    inconsistent_headers.setdefault(n, set()).add(s)
        except Exception:
            pass

    wb.close()

    # Note qualité
    score = 100
    if errors_locations: score -= min(30, len(errors_locations) * 2)
    if archive_sheets: score -= len(archive_sheets) * 5
    if accent_issues: score -= 5
    if empty_sheets: score -= 3
    note = "🟢" if score >= 80 else ("🟠" if score >= 60 else "🔴")

    rep.append(f"- Cellules de données estimées : {nb_data_cells:,}".replace(",", " "))
    rep.append(f"- Note qualité globale : {note} ({score}/100)\n")

    # Problèmes critiques
    rep.append("## 🔴 Problèmes critiques (bloquants pour parsing)")
    crit = []
    if errors_locations:
        nb_err = len(errors_locations)
        ex = ", ".join(f"`{s}!{c}`={v}" for s, c, v in errors_locations[:5])
        crit.append(f"**{nb_err}+ cellules en erreur** (#REF!/#DIV!/#N/A…) — ex : {ex}")
    if archive_sheets:
        crit.append(f"**{len(archive_sheets)} onglet(s) archive** polluant la détection auto : {', '.join(archive_sheets)}")
    if empty_sheets:
        crit.append(f"**{len(empty_sheets)} onglet(s) vide(s)** : {', '.join(empty_sheets)}")
    rep.append("\n".join(f"- {c}" for c in crit) if crit else "- Aucun problème critique détecté.")
    rep.append("")

    # Problèmes moyens
    rep.append("## 🟠 Problèmes moyens (qualité dégradée)")
    mid = []
    if accent_issues:
        ex = ", ".join(f"`{s}!{c}`='{v}'" for s, c, v, i in accent_issues[:3])
        mid.append(f"**{len(accent_issues)}+ libellés avec faute/incohérence** (ex Libéllé vs Libellé) — {ex}")
    inconsis = [(n, list(v)) for n, v in inconsistent_headers.items() if len(v) > 1]
    if inconsis:
        ex = ", ".join(f"'{a[0]}' vs '{a[1]}'" for _, a in inconsis[:3])
        mid.append(f"**{len(inconsis)} en-tête(s) avec variantes** entre feuilles — {ex}")
    rep.append("\n".join(f"- {m}" for m in mid) if mid else "- Aucun problème moyen détecté.")
    rep.append("")

    # Confort utilisateur
    rep.append("## 🟢 Points d'amélioration (confort utilisateur)")
    rep.append("- **Aucune validation Excel** : risque de fautes de frappe et nouvelles valeurs non normées.")
    rep.append("- **Aucune cellule figée** sur les feuilles de saisie.")
    rep.append("- Pas d'onglet « Mode d'emploi / Légende ».")
    rep.append("- Pas de feuille « Référentiels » centralisant les listes.")
    rep.append("")

    # Doublons et redondances
    rep.append("## 📊 Doublons et redondances")
    if len(sheets) != len(set(s.lower() for s in sheets)):
        rep.append("- ⚠ Onglets avec noms doublons (insensible casse) détectés.")
    else:
        rep.append("- Pas de doublon d'onglet flagrant détecté.")
    rep.append("")

    # Structure et nommage
    rep.append("## 🧱 Structure et nommage")
    # Convention onglets
    patterns = set()
    for s in sheets[:10]:
        if re.match(r'^\d{2}-\d{2}-\d{4}', s): patterns.add('DD-MM-YYYY')
        elif re.match(r'^\d{2}\s\d{2}\s\d{2}', s): patterns.add('DD MM YY')
        elif re.match(r'^[A-Z]+\s?\d{2}', s): patterns.add('MOIS YY')
        elif re.match(r'^\d{2}-\d{2}', s): patterns.add('MM-YY')
        elif re.match(r'^Sem|^Semaine', s, re.I): patterns.add('Semaine N')
    rep.append(f"- Convention de nommage onglets : {' / '.join(patterns) if patterns else 'hétérogène ou plat'}")
    rep.append(f"- Échantillon d'onglets : {', '.join(repr(s) for s in sheets[:5])}")
    rep.append("")

    # Recommandations
    rep.append("## 🎯 Recommandations prioritaires (top 5)")
    recos = []
    if errors_locations: recos.append("Résoudre les références #REF! (formules cassées vers fichiers externes ou cellules supprimées).")
    if archive_sheets: recos.append("Archiver les onglets historiques dans un fichier séparé (allège la détection).")
    if accent_issues: recos.append("Corriger les libellés fautifs (ex: 'Libéllé' → 'Libellé') pour standardiser le matching parser.")
    if inconsis: recos.append("Harmoniser les en-têtes entre feuilles (variantes accent/casse/espace).")
    recos.append("Ajouter des validations Excel (listes déroulantes) sur Client/Statut/Type/Mode.")
    recos.append("Figer la ligne d'en-tête sur chaque feuille de saisie.")
    recos.append("Ajouter un onglet « Mode d'emploi » documentant colonnes et conventions.")
    for i, r in enumerate(recos[:5], 1):
        rep.append(f"{i}. {r}")
    rep.append("")

    # Annexe
    rep.append("## Annexe : exemples concrets")
    if errors_locations:
        for s, c, v in errors_locations[:3]:
            rep.append(f"- `{s}!{c}` = `{v}` → résoudre la formule ou supprimer la cellule.")
    if accent_issues:
        for s, c, v, i in accent_issues[:3]:
            rep.append(f"- `{s}!{c}` = '{v}' → normaliser le libellé.")
    if not (errors_locations or accent_issues):
        rep.append("- Pas de cellule problématique détectée dans l'échantillon.")

    return "\n".join(rep)


def main():
    files_to_process = [
        "05_CA.xlsx", "06_ACCOSTAGE.xlsx", "07_QHSE.xlsx",
        "08_ACTIONS.xlsx", "09_CA_MAI.xlsx", "10_STATS_BASE.xlsx"
    ]
    rapport_globals = []
    for f in files_to_process:
        path = EXCEL_DIR / f
        if not path.exists():
            print(f"  ⚠ {f} introuvable")
            continue
        print(f"  → {f} ({path.stat().st_size//1024} Ko)...", flush=True)
        rep_md = audit_file(path)
        out_path = OUT_DIR / f"QUALITE_{path.stem}.md"
        out_path.write_text(rep_md, encoding="utf-8")
        print(f"     OK ({out_path.stat().st_size//1024} Ko)")
        rapport_globals.append((f, out_path.stat().st_size))
    print(f"\nDONE — {len(rapport_globals)} rapports générés")


if __name__ == "__main__":
    main()
