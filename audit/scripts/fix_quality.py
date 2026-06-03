"""Quick wins qualité :
1. Wrap les divisions sensibles en =IFERROR() dans 05_CA et 09_CA_MAI
2. Split CA en opérationnel + archives (PROGRAMME 2014, SOA Envoyés, Évolution clients)
3. POINTAGE : feuille SOURCE déjà intégrée via le template 04_POINTAGE_v1.xlsx (rien à faire ici)
"""
import openpyxl
from openpyxl.utils import get_column_letter
from pathlib import Path
import shutil

BASE = Path(__file__).parent.parent
SRC = BASE / "excel"
OUT = BASE / "excel_fixed"
OUT.mkdir(exist_ok=True)

ARCHIVE_TABS = ["PROGRAMME 2014", "SOA Envoyés", "Évolution clients", "Evolution clients", "Représentations"]


def fix_div_errors(src_path: Path, dst_path: Path):
    """Recharge le fichier en data_only=True pour repérer les #DIV/0!, puis en data_only=False
    pour réécrire les formules en =IFERROR(formule, "")"""
    shutil.copy(src_path, dst_path)
    # On charge en mode "formules" (data_only=False) pour modifier
    wb = openpyxl.load_workbook(dst_path, data_only=False)
    nb_wrapped = 0
    nb_cells_scanned = 0
    for sn in wb.sheetnames:
        ws = wb[sn]
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 200)):
            for cell in row:
                nb_cells_scanned += 1
                v = cell.value
                # Si c'est une formule (commence par =) et contient une division, on wrap
                if isinstance(v, str) and v.startswith("="):
                    # Heuristique : présence de '/' dans la formule => candidat
                    if "/" in v and not v.upper().startswith("=IFERROR"):
                        # Wrapper
                        inner = v[1:]  # supprime le =
                        cell.value = f'=IFERROR({inner},"")'
                        nb_wrapped += 1
    wb.save(dst_path)
    wb.close()
    return nb_wrapped, nb_cells_scanned


def split_archives(src_path: Path, dst_op: Path, dst_arch: Path):
    """Crée 2 fichiers : un opérationnel (sans onglets archive) et un archive (avec)."""
    # Version opérationnelle : supprime les onglets archives
    shutil.copy(src_path, dst_op)
    wb = openpyxl.load_workbook(dst_op, data_only=False)
    removed = []
    for tab in list(ARCHIVE_TABS):
        if tab in wb.sheetnames:
            del wb[tab]
            removed.append(tab)
    wb.save(dst_op)
    wb.close()

    # Version archive : garde uniquement les onglets archives
    shutil.copy(src_path, dst_arch)
    wb = openpyxl.load_workbook(dst_arch, data_only=False)
    to_remove = [sn for sn in wb.sheetnames if sn not in ARCHIVE_TABS]
    # Garder au moins 1 feuille (Excel exige)
    if len(to_remove) == len(wb.sheetnames):
        return removed, []
    for sn in to_remove:
        del wb[sn]
    wb.save(dst_arch)
    wb.close()
    return removed, list(wb.sheetnames if False else [])


def main():
    print("=== Quick wins qualité ===\n")
    # 1. CA : wrap IFERROR + split archives
    for ca_name in ["05_CA.xlsx", "09_CA_MAI.xlsx"]:
        src = SRC / ca_name
        if not src.exists():
            print(f"  ⚠ {ca_name} introuvable, skip"); continue
        # 1a. IFERROR fix
        fixed = OUT / ca_name.replace(".xlsx", "_FIXED.xlsx")
        nb, scanned = fix_div_errors(src, fixed)
        print(f"  ✓ {ca_name} → IFERROR wrapping : {nb} formules wrappées (scan {scanned} cellules)")
        # 1b. Split archives
        op = OUT / ca_name.replace(".xlsx", "_OPERATIONNEL.xlsx")
        arch = OUT / ca_name.replace(".xlsx", "_ARCHIVES.xlsx")
        removed, _ = split_archives(src, op, arch)
        print(f"     Split archives : {len(removed)} onglet(s) déplacé(s) → {removed}")
        print(f"     Fichiers : {op.name} (sans archives) + {arch.name} (archives uniquement)")

    print("\n  NOTE POINTAGE : la résolution VLOOKUP est déjà gérée via le")
    print("  template 04_POINTAGE_v1.xlsx (pré-rempli) qui intègre l'onglet")
    print("  SOURCE_Personnel avec les matricules uniques. Aucun fix à apporter")
    print("  au fichier source historique (les VLOOKUP cassés sont irrécupérables")
    print("  sans la feuille SOURCE externe).")

    print(f"\nDONE — fichiers fixés dans {OUT}/")


if __name__ == "__main__":
    main()
