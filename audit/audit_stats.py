"""Audit profond STATS_BASE_LOGISTIQUE : structure mois × KPI, sections, mapping colonnes."""
import openpyxl
from openpyxl.utils import get_column_letter
from pathlib import Path
from datetime import datetime

SRC = Path(__file__).parent / "excel" / "10_STATS_BASE.xlsx"
OUT = Path(__file__).parent / "reports" / "10_STATS_DEEP.md"

def norm(v):
    if v is None: return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()

def is_section_label(s):
    """Détecte une section majeure (texte court en majuscules)"""
    s = s.strip()
    if not s or len(s) > 60: return False
    upper_ratio = sum(1 for c in s if c.isupper()) / max(len(s),1)
    has_alpha = any(c.isalpha() for c in s)
    return has_alpha and (upper_ratio > 0.5 or s.endswith(":"))

def audit_sheet(ws, max_rows=200, max_cols=20):
    rep = []
    rep.append(f"\n## Feuille : `{ws.title}`")
    rep.append(f"- Dimensions : **{ws.max_row} lignes × {ws.max_column} colonnes**\n")

    # Détecter ligne d'en-tête mois (cherche des dates ou nombres 1-12 dans les premières 10 lignes)
    header_row = None
    for r in range(1, min(15, ws.max_row+1)):
        date_count = 0
        for c in range(2, min(ws.max_column+1, max_cols+1)):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, datetime):
                date_count += 1
            elif isinstance(v, (int,float)) and 1 <= v <= 12:
                date_count += 1
        if date_count >= 3:
            header_row = r
            break
    if header_row:
        rep.append(f"- **Ligne d'en-tête mois détectée : L{header_row}**")
        months = []
        for c in range(2, min(ws.max_column+1, max_cols+5)):
            v = ws.cell(row=header_row, column=c).value
            if v is not None:
                months.append(f"{get_column_letter(c)}={norm(v)[:15]}")
        rep.append(f"  - Colonnes : {' | '.join(months[:15])}")
    else:
        rep.append("- (pas de ligne d'en-tête mois claire dans les 15 premières lignes)")

    # Lignes avec contenu en col A : KPIs candidats
    rep.append(f"\n### KPIs (colonne A) — toutes lignes non vides")
    sections = []
    current_sec = None
    kpi_count = 0
    for r in range(1, min(max_rows+1, ws.max_row+1)):
        a = norm(ws.cell(row=r, column=1).value)
        b = norm(ws.cell(row=r, column=2).value)
        if not a and not b: continue
        # Check if it's a section header (ALL CAPS or specific keywords)
        if a and is_section_label(a):
            current_sec = a
            sections.append((r, a))
            rep.append(f"\n#### 🔹 L{r} — **{a}**")
            continue
        if a:
            kpi_count += 1
            # Échantillon valeurs : 3 premières colonnes mois
            vals = []
            for c in range(2, min(ws.max_column+1, 8)):
                v = ws.cell(row=r, column=c).value
                if v is not None:
                    s = norm(v)[:20]
                    if s and s != "0": vals.append(f"{get_column_letter(c)}=`{s}`")
            sample = (" | ".join(vals[:5])) if vals else "_(vide)_"
            rep.append(f"- L{r} `{a[:60]}` → {sample}")

    rep.append(f"\n**Total : {kpi_count} KPIs · {len(sections)} sections**")

    return "\n".join(rep)

def main():
    wb = openpyxl.load_workbook(SRC, data_only=False)
    out = [f"# STATS_BASE_LOGISTIQUE — Audit profond\n"]
    out.append(f"Source : `{SRC.name}` — {len(wb.sheetnames)} onglet(s)\n")
    out.append("Objectif : identifier les KPIs disponibles pour le parser parseStats() à venir.\n")
    for sn in wb.sheetnames:
        out.append(audit_sheet(wb[sn]))
    OUT.write_text("\n".join(out), encoding="utf-8")
    wb.close()
    print(f"DONE — {OUT} ({OUT.stat().st_size//1024} Ko)")

if __name__ == "__main__":
    main()
