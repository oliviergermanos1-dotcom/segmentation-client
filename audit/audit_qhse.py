"""Audit approfondi QHSE : détecte sections, blocs de données, colonnes utilisées."""
import openpyxl
from openpyxl.utils import get_column_letter
from pathlib import Path

SRC = Path(__file__).parent / "excel" / "07_QHSE.xlsx"
OUT = Path(__file__).parent / "reports" / "07_QHSE_DEEP.md"

def norm(v):
    if v is None: return ""
    return str(v).strip()

def is_section_header(cells):
    """Une ligne est un header de section si col A ou B contient un texte long en majuscules
    ou un libellé typique (PERSONNEL, EQUIPEMENTS, METEO, INCIDENTS, etc.)."""
    txt = " ".join(norm(c) for c in cells[:4])
    if not txt: return False
    keywords = ['PERSONNEL','EFFECTIF','EQUIPEMENT','MATERIEL','METEO','MÉTÉO','OPERATION','INCIDENT',
                'ACCIDENT','OBSERVATION','SECURITE','SÉCURITÉ','FORMATION','EPI','ACTION','ANALYSE',
                'TRIF','LTIF','IOGP','HEURE','JSA','PRESQUE','NEAR','RAPPORT','SEMAINE','PAYS','ENTITE',
                'TOTAL','SOUS-TRAITANT','JOURNALIERS','CADRE','CERTIFICATION']
    for k in keywords:
        if k in txt.upper(): return True
    return False

def cell_density(row):
    """% de cellules non vides dans la ligne"""
    if not row: return 0
    non_empty = sum(1 for c in row if norm(c))
    return non_empty / len(row) * 100

def audit_sheet(ws, max_cols=27):
    rep = []
    rep.append(f"\n## Feuille : `{ws.title}`")
    rep.append(f"- Dimensions : **{ws.max_row} lignes × {ws.max_column} colonnes**")

    # Lire toutes les lignes une fois
    rows = []
    for r in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=min(ws.max_column, max_cols), values_only=True):
        rows.append(list(r))

    # Détecter sections
    sections = []
    current = None
    for i, row in enumerate(rows):
        if is_section_header(row):
            if current:
                current['end'] = i - 1
                sections.append(current)
            current = {'start': i, 'header': " | ".join(norm(c) for c in row[:6] if norm(c))[:120], 'rows': []}
        elif current and any(norm(c) for c in row):
            current['rows'].append((i, row))
    if current:
        current['end'] = len(rows) - 1
        sections.append(current)

    rep.append(f"- Sections détectées : **{len(sections)}**\n")

    for sec in sections:
        rep.append(f"\n### 🔹 Section ligne {sec['start']+1} → {sec['end']+1} ({len(sec['rows'])} lignes données)")
        rep.append(f"**Header** : {sec['header']}")
        # Colonnes utilisées dans cette section
        used_cols = set()
        for _, row in sec['rows']:
            for ci, v in enumerate(row):
                if norm(v): used_cols.add(ci)
        if used_cols:
            cols_str = ", ".join(get_column_letter(c+1) for c in sorted(used_cols))
            rep.append(f"**Colonnes utilisées** : {cols_str} ({len(used_cols)} cols)")
        # Échantillon : 3 lignes représentatives (la plus dense)
        if sec['rows']:
            sorted_rows = sorted(sec['rows'], key=lambda x: -cell_density(x[1]))[:3]
            rep.append("\n**Exemples (lignes les plus denses)** :")
            for ri, row in sorted_rows:
                content = " | ".join(f"`{get_column_letter(ci+1)}={norm(v)[:30]}`" for ci, v in enumerate(row) if norm(v))
                rep.append(f"- L{ri+1} : {content[:400]}")

    # Lignes hors sections détectées
    in_sec = set()
    for sec in sections:
        for r in range(sec['start'], sec['end']+1):
            in_sec.add(r)
    orphan_data = []
    for i, row in enumerate(rows):
        if i not in in_sec and any(norm(c) for c in row):
            orphan_data.append((i, row))
    if orphan_data:
        rep.append(f"\n### ⚠️ Lignes hors section ({len(orphan_data)}) — premières 10")
        for ri, row in orphan_data[:10]:
            content = " | ".join(f"`{get_column_letter(ci+1)}={norm(v)[:30]}`" for ci, v in enumerate(row) if norm(v))
            rep.append(f"- L{ri+1} : {content[:300]}")

    return "\n".join(rep)

def main():
    wb = openpyxl.load_workbook(SRC, data_only=False)
    out = [f"# QHSE — Audit approfondi (toutes lignes)\n"]
    out.append(f"Source : `{SRC.name}` — {len(wb.sheetnames)} onglet(s)\n")
    for sn in wb.sheetnames:
        out.append(audit_sheet(wb[sn]))
    OUT.write_text("\n".join(out), encoding="utf-8")
    wb.close()
    print(f"DONE — {OUT} ({OUT.stat().st_size//1024} Ko)")

if __name__ == "__main__":
    main()
