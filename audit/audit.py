"""Audit structurel des 8 fichiers Excel : onglets, colonnes, validations, échantillons."""
import sys, json, os, re
from pathlib import Path
import openpyxl
from openpyxl.utils import get_column_letter

EXCEL_DIR = Path(__file__).parent / "excel"
REPORT_DIR = Path(__file__).parent / "reports"
REPORT_DIR.mkdir(exist_ok=True)

def norm(v):
    if v is None: return ""
    s = str(v).strip()
    return s

def fmt_val(v, maxlen=40):
    s = norm(v)
    if len(s) > maxlen: s = s[:maxlen-1]+"…"
    return s

def audit_file(path: Path):
    rep = []
    rep.append(f"# Audit — {path.name}\n")
    rep.append(f"Taille : {path.stat().st_size/1024:.1f} Ko\n")
    try:
        wb = openpyxl.load_workbook(path, data_only=False, read_only=False)
    except Exception as e:
        rep.append(f"\n**ERREUR chargement** : {e}\n")
        return "\n".join(rep)

    rep.append(f"\n## Onglets ({len(wb.sheetnames)})\n")
    for sn in wb.sheetnames:
        rep.append(f"- `{sn}`")
    rep.append("")

    # Named ranges
    try:
        names = list(wb.defined_names)
        if names:
            rep.append(f"## Plages nommées ({len(names)})\n")
            for n in names[:30]:
                try:
                    dest = list(wb.defined_names[n].destinations) if hasattr(wb.defined_names[n],'destinations') else []
                    rep.append(f"- `{n}` → {dest}")
                except: rep.append(f"- `{n}`")
            rep.append("")
    except Exception: pass

    for sn in wb.sheetnames:
        ws = wb[sn]
        rep.append(f"\n## Feuille : `{sn}`\n")
        rep.append(f"- Dimensions : {ws.max_row} lignes × {ws.max_column} colonnes (réf : `{ws.dimensions}`)")
        try:
            mc = list(ws.merged_cells.ranges)
            if mc:
                rep.append(f"- Cellules fusionnées : {len(mc)} plages — ex : {[str(x) for x in mc[:5]]}")
        except Exception: pass

        # Data validations / listes déroulantes
        try:
            dvs = list(ws.data_validations.dataValidation) if hasattr(ws,'data_validations') else []
            if dvs:
                rep.append(f"- Validations de données : {len(dvs)}")
                for dv in dvs[:8]:
                    f = (dv.formula1 or "")[:120] if hasattr(dv,'formula1') else ""
                    rep.append(f"  - type={dv.type} formula1=`{f}` cells={list(dv.sqref.ranges)[:3] if dv.sqref else []}")
        except Exception as e:
            rep.append(f"- (validations non lisibles : {e})")

        # Echantillon : premières 25 lignes (toutes colonnes)
        rep.append(f"\n### Échantillon — 25 premières lignes\n")
        max_r = min(25, ws.max_row or 0)
        max_c = min(20, ws.max_column or 0)  # cap cols pour lisibilité
        if max_r == 0 or max_c == 0:
            rep.append("_(feuille vide)_")
            continue
        # En-têtes colonnes Excel
        header = "| # | " + " | ".join(f"col{get_column_letter(c)}" for c in range(1, max_c+1)) + " |"
        sep    = "|---|" + "|".join("---" for _ in range(max_c)) + "|"
        rep.append(header); rep.append(sep)
        for r in range(1, max_r+1):
            cells = []
            for c in range(1, max_c+1):
                try:
                    v = ws.cell(row=r, column=c).value
                except Exception: v = "?"
                cells.append(fmt_val(v).replace("|","\\|"))
            rep.append(f"| {r} | " + " | ".join(cells) + " |")
        if ws.max_row > 25:
            rep.append(f"\n_… {ws.max_row - 25} lignes supplémentaires non affichées_")
        if ws.max_column > 20:
            rep.append(f"\n_… {ws.max_column - 20} colonnes supplémentaires (au-delà de col T) non affichées_")

        # Stats numériques rapides
        try:
            num_count = 0
            for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 1000), values_only=True):
                for v in row:
                    if isinstance(v, (int, float)) and v != 0:
                        num_count += 1
            rep.append(f"\n- Cellules numériques non nulles (1000 premières lignes) : {num_count}")
        except Exception: pass

    wb.close()
    return "\n".join(rep)

def main():
    index = ["# Index audit Excel — 8 fichiers\n"]
    for f in sorted(EXCEL_DIR.glob("*.xlsx")):
        print(f"-> {f.name}", flush=True)
        out = audit_file(f)
        rep_path = REPORT_DIR / (f.stem + ".md")
        rep_path.write_text(out, encoding="utf-8")
        index.append(f"- [{f.name}]({rep_path.name}) — {f.stat().st_size/1024:.0f} Ko")
    (REPORT_DIR / "00_INDEX.md").write_text("\n".join(index), encoding="utf-8")
    print("DONE — rapports dans", REPORT_DIR)

if __name__ == "__main__":
    main()
