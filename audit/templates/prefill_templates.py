"""
Phase 4 — Pré-remplissage des 11 templates Excel "stricts" AGL.

Charge chaque template dans `output_prefilled/`, copie les données du source
correspondant dans les bonnes cellules en respectant les en-têtes verrouillées,
puis sauvegarde par-dessus.

Règles strictes (cf. brief) :
- NE PAS toucher à la ligne 1 (en-têtes verrouillées) ni à l'onglet "📖 Mode d'emploi"
  ni à l'onglet caché "_SYS".
- Écrire à partir de la ligne 2.
- Pour chaque cellule écrite : Protection(locked=False) pour rester éditable.
- Tolérer les erreurs source : try/catch autour de chaque fichier source.
- Si une donnée n'est pas mappable proprement, on l'ignore (pas de plantage).

Notes pragmatiques :
- 03_PLANNING : on copie 1 onglet jour + 1 onglet nuit (les onglets MODELE-*).
- 04_POINTAGE : on extrait les matricules UNIQUES dans SOURCE_Personnel +
  1 exemple JOUR/NUIT/WK dans les onglets MODELE-*.
- 05B_CA_PNL : le source est en format MATRICE compliqué ; on extrait seulement
  les lignes principales lisibles. Documenté dans la fonction.
- 07_QHSE : on prend la 1ère semaine source et on remplit les sous-onglets thématiques.
"""
from __future__ import annotations

import re
import sys
import warnings
from datetime import date, datetime, time
from pathlib import Path

import openpyxl
from openpyxl.styles import Protection

warnings.simplefilter("ignore")

ROOT = Path(__file__).resolve().parent
TPL_DIR = ROOT / "output_prefilled"
SRC_DIR = ROOT.parent / "excel"

UNLOCKED = Protection(locked=False)

# Compteurs
STATS = {"files": 0, "cells_total": 0}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def set_cell(ws, row, col, value, number_format=None):
    """Écrit une cellule, déverrouille pour saisie, retourne 1 si écrit."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return 0
    c = ws.cell(row=row, column=col, value=value)
    c.protection = UNLOCKED
    if number_format:
        c.number_format = number_format
    return 1


def open_src(name):
    """Ouvre un fichier source avec warnings filtrés."""
    return openpyxl.load_workbook(SRC_DIR / name, data_only=True)


def open_tpl(name):
    """Ouvre un template à pré-remplir."""
    return openpyxl.load_workbook(TPL_DIR / name)


def parse_time_str(s):
    """Convertit '11H45', '11h45', '04H 00' etc. en time(h, m) ; sinon None."""
    if s is None:
        return None
    if isinstance(s, time):
        return s
    if isinstance(s, datetime):
        return s.time()
    if not isinstance(s, str):
        return None
    s = s.strip().replace(" ", "")
    m = re.match(r"^(\d{1,2})[Hh:](\d{2})$", s)
    if m:
        try:
            return time(int(m.group(1)) % 24, int(m.group(2)))
        except ValueError:
            return None
    return None


def safe_save(wb, path):
    """Sauvegarde robuste."""
    wb.save(path)


# ---------------------------------------------------------------------------
# 00 — MASTER REFERENTIELS : enrichir avec valeurs trouvées dans sources
# ---------------------------------------------------------------------------
def prefill_master():
    tpl_name = "00_MASTER_REFERENTIELS.xlsx"
    wb = open_tpl(tpl_name)
    written = 0
    tabs_touched = set()

    # --- 1. Collecter clients de toutes les sources ---
    clients_uniques = set()

    # Planning : col B (CLIENT)
    try:
        src = open_src("03_PLANNING.xlsx")
        for sname in src.sheetnames[:5]:  # échantillon (51 onglets)
            ws = src[sname]
            for r in range(1, min(ws.max_row + 1, 60)):
                v = ws.cell(row=r, column=2).value
                if isinstance(v, str) and v.strip() and len(v.strip()) > 1:
                    clients_uniques.add(v.strip())
        src.close()
    except Exception as e:
        print(f"   ! Planning clients: {e}")

    # Accostage : col D (CLIENTS)
    try:
        src = open_src("06_ACCOSTAGE.xlsx")
        for sname in src.sheetnames:
            ws = src[sname]
            for r in range(5, ws.max_row + 1):
                v = ws.cell(row=r, column=4).value
                if isinstance(v, str) and v.strip():
                    clients_uniques.add(v.strip())
        src.close()
    except Exception as e:
        print(f"   ! Accostage clients: {e}")

    # Loyers : col A
    try:
        src = open_src("01_LOYERS.xlsx")
        ws = src["Jan"]
        for r in range(3, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if isinstance(v, str) and v.strip():
                clients_uniques.add(v.strip())
        src.close()
    except Exception as e:
        print(f"   ! Loyers clients: {e}")

    # CA Janvier : col A
    try:
        src = open_src("09_CA_MAI.xlsx")
        ws = src["Janvier"]
        for r in range(4, ws.max_row + 1):
            v = ws.cell(row=r, column=1).value
            if isinstance(v, str) and v.strip() and v.strip().upper() != "TOTAL":
                clients_uniques.add(v.strip())
        src.close()
    except Exception as e:
        print(f"   ! CA clients: {e}")

    # nettoyer : retirer mots/notes qui ne sont pas des clients
    bad_keywords = ("PLANNING", "OPERATION", "CLIENT", "PROJET", "OBSERV",
                    "TOTAL", "CLIENTS", "SECTION", "NATURE", "DOCKER")
    clients_uniques = {
        c for c in clients_uniques
        if not any(b in c.upper() for b in bad_keywords)
        and len(c) > 1 and len(c) < 80
    }

    # Ajouter dans onglet Clients (sans doublon avec existants)
    ws = wb["Clients"]
    existants = {ws.cell(row=r, column=2).value for r in range(2, ws.max_row + 1)
                 if ws.cell(row=r, column=2).value}
    next_row = ws.max_row + 1
    # trouver première ligne vide réelle
    while ws.cell(row=next_row - 1, column=1).value is None and next_row > 2:
        next_row -= 1
    for c in sorted(clients_uniques):
        if c in existants:
            continue
        w = 0
        w += set_cell(ws, next_row, 1, c)
        w += set_cell(ws, next_row, 2, c)
        w += set_cell(ws, next_row, 4, "Service")
        w += set_cell(ws, next_row, 5, "OUI")
        written += w
        if w:
            next_row += 1
            tabs_touched.add("Clients")

    # --- 2. Espaces : remplir avec Feuil1 de 01_LOYERS (Client + Libellé) ---
    try:
        ws = wb["Espaces"]
        existants_lib = {ws.cell(row=r, column=2).value for r in range(2, ws.max_row + 1)
                         if ws.cell(row=r, column=2).value}
        next_row = ws.max_row + 1
        while ws.cell(row=next_row - 1, column=1).value is None and next_row > 2:
            next_row -= 1
        src = open_src("01_LOYERS.xlsx")
        sws = src["Feuil1"]
        idx = 1
        for r in range(3, sws.max_row + 1):
            client = sws.cell(row=r, column=1).value
            libelle = sws.cell(row=r, column=2).value
            if not libelle:
                continue
            if libelle in existants_lib:
                continue
            w = 0
            w += set_cell(ws, next_row, 1, f"ESP-{idx:03d}")
            w += set_cell(ws, next_row, 2, str(libelle))
            w += set_cell(ws, next_row, 3, "Mixte")
            w += set_cell(ws, next_row, 5, str(client) if client else "")
            written += w
            if w:
                next_row += 1
                idx += 1
                tabs_touched.add("Espaces")
        src.close()
    except Exception as e:
        print(f"   ! Master.Espaces: {e}")

    # --- 3. Fournisseurs : extraire col C de 02_CHARGES ---
    try:
        fourn_uniques = set()
        src = open_src("02_CHARGES.xlsx")
        for sname in src.sheetnames:
            sws = src[sname]
            for r in range(4, sws.max_row + 1):
                v = sws.cell(row=r, column=3).value
                if isinstance(v, str) and v.strip():
                    fourn_uniques.add(v.strip())
        src.close()
        ws = wb["Fournisseurs"]
        existants = {ws.cell(row=r, column=2).value for r in range(2, ws.max_row + 1)
                     if ws.cell(row=r, column=2).value}
        next_row = ws.max_row + 1
        while ws.cell(row=next_row - 1, column=1).value is None and next_row > 2:
            next_row -= 1
        idx = len(existants) + 1
        for f in sorted(fourn_uniques):
            if f in existants:
                continue
            w = 0
            w += set_cell(ws, next_row, 1, f"F-{idx:03d}")
            w += set_cell(ws, next_row, 2, f)
            w += set_cell(ws, next_row, 3, "Virement")
            written += w
            if w:
                next_row += 1
                idx += 1
                tabs_touched.add("Fournisseurs")
    except Exception as e:
        print(f"   ! Master.Fournisseurs: {e}")

    # --- 4. Equipements : enrichir avec QHSE L26-50 ---
    try:
        src = open_src("07_QHSE.xlsx")
        sws = src[src.sheetnames[0]]
        equips = []
        for r in range(27, 60):
            typ = sws.cell(row=r, column=3).value
            ton = sws.cell(row=r, column=4).value
            nb = sws.cell(row=r, column=5).value
            cert = sws.cell(row=r, column=6).value
            if typ and ton:
                equips.append((str(typ).strip(), str(ton).strip(),
                               nb if isinstance(nb, (int, float)) else 0,
                               str(cert).strip() if cert else "Certifié"))
        src.close()
        ws = wb["Equipements"]
        existants = {(ws.cell(row=r, column=2).value, ws.cell(row=r, column=3).value)
                     for r in range(2, ws.max_row + 1)}
        next_row = ws.max_row + 1
        while ws.cell(row=next_row - 1, column=1).value is None and next_row > 2:
            next_row -= 1
        idx = 100
        for typ, ton, nb, cert in equips:
            if (typ, ton) in existants:
                continue
            w = 0
            w += set_cell(ws, next_row, 1, f"EQ-{idx:03d}")
            w += set_cell(ws, next_row, 2, typ)
            w += set_cell(ws, next_row, 3, ton)
            w += set_cell(ws, next_row, 4, nb)
            w += set_cell(ws, next_row, 5, cert if cert in ("Certifié", "Étalonné", "En cours", "À réaliser") else "Certifié")
            written += w
            if w:
                next_row += 1
                idx += 1
                tabs_touched.add("Equipements")
    except Exception as e:
        print(f"   ! Master.Equipements: {e}")

    # --- 5. Sections_Planning : extraire de 03_PLANNING col C ---
    try:
        sections = set()
        src = open_src("03_PLANNING.xlsx")
        for sname in src.sheetnames[:8]:
            sws = src[sname]
            for r in range(4, min(sws.max_row + 1, 40)):
                v = sws.cell(row=r, column=3).value
                if isinstance(v, str) and v.strip() and len(v.strip()) < 30:
                    s = v.strip().rstrip()
                    # garder seulement les codes courts
                    if len(s) <= 20:
                        sections.add(s)
        src.close()
        ws = wb["Sections_Planning"]
        existants = {ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)
                     if ws.cell(row=r, column=1).value}
        next_row = ws.max_row + 1
        while ws.cell(row=next_row - 1, column=1).value is None and next_row > 2:
            next_row -= 1
        for s in sorted(sections):
            if s in existants:
                continue
            w = 0
            w += set_cell(ws, next_row, 1, s)
            w += set_cell(ws, next_row, 2, s)
            w += set_cell(ws, next_row, 3, "OUI")
            written += w
            if w:
                next_row += 1
                tabs_touched.add("Sections_Planning")
    except Exception as e:
        print(f"   ! Master.Sections: {e}")

    safe_save(wb, TPL_DIR / tpl_name)
    print(f"  OK {tpl_name:38s} — {written} cellules sur {len(tabs_touched)} onglets")
    STATS["files"] += 1
    STATS["cells_total"] += written


# ---------------------------------------------------------------------------
# 01 LOYERS
# ---------------------------------------------------------------------------
def prefill_loyers():
    tpl_name = "01_LOYERS_v1.xlsx"
    wb = open_tpl(tpl_name)
    ws = wb["Loyers"]
    written = 0
    try:
        src = open_src("01_LOYERS.xlsx")
        sws = src["Jan"]
        # Source: R1 titre, R2 headers, R3+ data
        # Source cols: A=Client, B=Libellé, C=Montant, D-O = 12 mois (M1-M12)
        # Template cols: A=Client, B=ID Espace (laisse vide), C=Libellé, D=Montant, E-P = 12 mois, Q=Commentaire
        out_row = 2
        for r in range(3, sws.max_row + 1):
            client = sws.cell(row=r, column=1).value
            libelle = sws.cell(row=r, column=2).value
            montant = sws.cell(row=r, column=3).value
            if not client and not libelle:
                continue
            w = 0
            w += set_cell(ws, out_row, 1, client)
            w += set_cell(ws, out_row, 3, libelle)
            w += set_cell(ws, out_row, 4, montant, "#,##0")
            for m_idx in range(12):
                v = sws.cell(row=r, column=4 + m_idx).value
                if v is not None:
                    w += set_cell(ws, out_row, 5 + m_idx, v, "#,##0")
            if w:
                out_row += 1
                written += w
        src.close()
    except Exception as e:
        print(f"   ! 01_LOYERS: {e}")

    safe_save(wb, TPL_DIR / tpl_name)
    print(f"  OK {tpl_name:38s} — {written} cellules sur 1 onglet")
    STATS["files"] += 1
    STATS["cells_total"] += written


# ---------------------------------------------------------------------------
# 02 CHARGES — 5 mois disponibles (01-26 à 05-26)
# ---------------------------------------------------------------------------
def prefill_charges():
    tpl_name = "02_CHARGES_v1.xlsx"
    wb = open_tpl(tpl_name)
    written = 0
    tabs = 0
    try:
        src = open_src("02_CHARGES.xlsx")
        # Map: source name -> target name (both same name pattern)
        for sname in src.sheetnames:
            if sname not in wb.sheetnames:
                continue
            sws = src[sname]
            tws = wb[sname]
            # source: R3 = headers, R4+ = data (parfois R3 est le titre)
            # détecter ligne header
            header_row = None
            for r in range(1, 8):
                v = sws.cell(row=r, column=1).value
                if isinstance(v, str) and v.strip().lower() == "libellé":
                    header_row = r
                    break
            if header_row is None:
                continue
            n_cols_src = sws.max_column
            out_row = 2
            for r in range(header_row + 1, sws.max_row + 1):
                libelle = sws.cell(row=r, column=1).value
                if not libelle:
                    continue
                w = 0
                # A=Libellé, B=N°UM, C=Fournisseur, D=Affectation, E=Montant
                w += set_cell(tws, out_row, 1, sws.cell(row=r, column=1).value)
                w += set_cell(tws, out_row, 2, sws.cell(row=r, column=2).value)
                w += set_cell(tws, out_row, 3, sws.cell(row=r, column=3).value)
                w += set_cell(tws, out_row, 4, sws.cell(row=r, column=4).value)
                w += set_cell(tws, out_row, 5, sws.cell(row=r, column=5).value, "#,##0")
                if n_cols_src >= 6:
                    mois_val = sws.cell(row=r, column=6).value
                    if isinstance(mois_val, str):
                        mois_val = mois_val.strip().rstrip()
                    w += set_cell(tws, out_row, 6, mois_val)
                if n_cols_src >= 7:
                    w += set_cell(tws, out_row, 7, sws.cell(row=r, column=7).value)
                if n_cols_src >= 8:
                    w += set_cell(tws, out_row, 8, sws.cell(row=r, column=8).value)
                if w:
                    out_row += 1
                    written += w
            if out_row > 2:
                tabs += 1
        src.close()
    except Exception as e:
        print(f"   ! 02_CHARGES: {e}")

    safe_save(wb, TPL_DIR / tpl_name)
    print(f"  OK {tpl_name:38s} — {written} cellules sur {tabs} onglets")
    STATS["files"] += 1
    STATS["cells_total"] += written


# ---------------------------------------------------------------------------
# 03 PLANNING — 1 onglet jour + 1 onglet nuit
# ---------------------------------------------------------------------------
def prefill_planning():
    tpl_name = "03_PLANNING_v1.xlsx"
    wb = open_tpl(tpl_name)
    written = 0
    try:
        src = open_src("03_PLANNING.xlsx")
        # Choisir : 1 onglet JOUR significatif + 1 onglet NUIT significatif
        # Sélection : prendre des sheets avec plus de lignes
        chosen_jour = None
        chosen_nuit = None
        for sname in src.sheetnames:
            if "JOUR" in sname.upper() and "NUIT" not in sname.upper():
                if chosen_jour is None or src[sname].max_row > src[chosen_jour].max_row:
                    chosen_jour = sname
            elif "NUIT" in sname.upper():
                if chosen_nuit is None or src[sname].max_row > src[chosen_nuit].max_row:
                    chosen_nuit = sname

        def copy_planning(src_sname, target_sheet, shift):
            nonlocal written
            sws = src[src_sname]
            tws = wb[target_sheet]
            # Trouver ligne header (cherche "CLIENT" en col B ou C)
            header_row = None
            client_col = 2
            for r in range(1, 12):
                for c in (2, 3):
                    v = sws.cell(row=r, column=c).value
                    if isinstance(v, str) and v.strip().upper() == "CLIENT":
                        header_row = r
                        client_col = c
                        break
                if header_row:
                    break
            if header_row is None:
                return
            # Mapping cols source : si client en B (col2), shift JOUR : B=client,C=section,D=demande,E=nature,F=positif,G=nega,H=obs
            # Si client en C (col3), shift NUIT : C=client,D=section,E=demande,F=nature,G=positif,H=nega,I=obs
            out_row = 2
            n = 0
            for r in range(header_row + 1, sws.max_row + 1):
                client = sws.cell(row=r, column=client_col).value
                section = sws.cell(row=r, column=client_col + 1).value
                if not client and not section:
                    continue
                if isinstance(client, str) and client.strip().upper() in ("LEGENDE", "TOTAL"):
                    continue
                w = 0
                n += 1
                w += set_cell(tws, out_row, 1, n)
                w += set_cell(tws, out_row, 2, client)
                w += set_cell(tws, out_row, 3, section)
                w += set_cell(tws, out_row, 4, sws.cell(row=r, column=client_col + 2).value)
                w += set_cell(tws, out_row, 5, sws.cell(row=r, column=client_col + 3).value)
                w += set_cell(tws, out_row, 6, sws.cell(row=r, column=client_col + 4).value)
                w += set_cell(tws, out_row, 7, sws.cell(row=r, column=client_col + 5).value)
                w += set_cell(tws, out_row, 8, sws.cell(row=r, column=client_col + 6).value)
                w += set_cell(tws, out_row, 10, shift)
                if w:
                    out_row += 1
                    written += w

        if chosen_jour:
            copy_planning(chosen_jour, "MODELE-JOUR", "JOUR")
        if chosen_nuit:
            copy_planning(chosen_nuit, "MODELE-NUIT", "NUIT")
        src.close()
    except Exception as e:
        print(f"   ! 03_PLANNING: {e}")

    safe_save(wb, TPL_DIR / tpl_name)
    print(f"  OK {tpl_name:38s} — {written} cellules sur 2 onglets")
    STATS["files"] += 1
    STATS["cells_total"] += written


# ---------------------------------------------------------------------------
# 04 POINTAGE — SOURCE_Personnel + 1 exemple par modèle
# ---------------------------------------------------------------------------
def prefill_pointage():
    tpl_name = "04_POINTAGE_v1.xlsx"
    wb = open_tpl(tpl_name)
    written = 0
    tabs = set()
    try:
        src = open_src("04_POINTAGE.xlsx")
        # Collecter matricules uniques : col C, fonction col E, code col F, nom col D
        personnels = {}
        for sname in src.sheetnames:
            sws = src[sname]
            # Trouver header row (cherche "Matricule" en col C)
            header_row = None
            for r in range(1, 15):
                v = sws.cell(row=r, column=3).value
                if isinstance(v, str) and "matricule" in v.lower():
                    header_row = r
                    break
            if header_row is None:
                continue
            for r in range(header_row + 1, sws.max_row + 1):
                mat = sws.cell(row=r, column=3).value
                if mat is None:
                    continue
                # convertir en str/int
                if isinstance(mat, (int, float)):
                    mat = str(int(mat))
                elif isinstance(mat, str):
                    mat = mat.strip()
                    if not mat:
                        continue
                else:
                    continue
                if mat in personnels:
                    continue
                nom = sws.cell(row=r, column=4).value
                fonct = sws.cell(row=r, column=5).value
                code = sws.cell(row=r, column=6).value
                chantier = sws.cell(row=r, column=7).value
                produit = sws.cell(row=r, column=8).value
                personnels[mat] = (nom, fonct, code, chantier, produit)
        src.close()

        # Remplir SOURCE_Personnel (en gardant les 2 exemples existants si différents)
        ws = wb["SOURCE_Personnel"]
        existants = {ws.cell(row=r, column=1).value for r in range(2, ws.max_row + 1)
                     if ws.cell(row=r, column=1).value}
        # Identifier première ligne vide
        next_row = 2
        while ws.cell(row=next_row, column=1).value is not None:
            next_row += 1

        for mat, (nom, fonct, code, chantier, produit) in sorted(personnels.items()):
            if mat in existants or (isinstance(mat, str) and any(m == mat for m in existants)):
                continue
            w = 0
            w += set_cell(ws, next_row, 1, mat)
            w += set_cell(ws, next_row, 2, nom)
            w += set_cell(ws, next_row, 3, fonct)
            w += set_cell(ws, next_row, 4, code)
            w += set_cell(ws, next_row, 5, chantier)
            w += set_cell(ws, next_row, 6, produit)
            w += set_cell(ws, next_row, 7, "OUI")
            if w:
                next_row += 1
                written += w
                tabs.add("SOURCE_Personnel")

        # 1 exemple par modèle (JOUR/NUIT/WK) — copier qq lignes du source
        src = open_src("04_POINTAGE.xlsx")

        def find_sheet(keyword_pos, keyword_neg=None):
            for s in src.sheetnames:
                up = s.upper()
                if keyword_pos in up:
                    if keyword_neg and keyword_neg in up:
                        continue
                    return s
            return None

        sample_jour = find_sheet("JOUR", "WK")
        sample_nuit = find_sheet("NUIT")
        sample_wk = find_sheet("WK")

        def copy_sample(src_sname, target_name, shift):
            nonlocal written
            if src_sname is None:
                return
            sws = src[src_sname]
            tws = wb[target_name]
            # Trouver header (cherche "Matricule")
            header_row = None
            for r in range(1, 15):
                v = sws.cell(row=r, column=3).value
                if isinstance(v, str) and "matricule" in v.lower():
                    header_row = r
                    break
            if header_row is None:
                return
            # Trouver une date dans la sheet
            today = date(2026, 5, 1)
            for r in range(1, header_row):
                for c in range(1, 12):
                    v = sws.cell(row=r, column=c).value
                    if isinstance(v, datetime):
                        today = v.date()
                        break
            out_row = 2
            n = 0
            for r in range(header_row + 1, min(sws.max_row + 1, header_row + 30)):
                mat = sws.cell(row=r, column=3).value
                if mat is None:
                    continue
                if isinstance(mat, float):
                    mat = int(mat)
                w = 0
                n += 1
                w += set_cell(tws, out_row, 1, n)
                w += set_cell(tws, out_row, 2, today, "yyyy-mm-dd")
                w += set_cell(tws, out_row, 3, mat)
                # Colonnes D/E/F gérées par VLOOKUP du template
                w += set_cell(tws, out_row, 7, sws.cell(row=r, column=7).value)
                w += set_cell(tws, out_row, 8, sws.cell(row=r, column=8).value)
                w += set_cell(tws, out_row, 9, shift)
                if w:
                    out_row += 1
                    written += w
            if out_row > 2:
                tabs.add(target_name)

        copy_sample(sample_jour, "MODELE-JOUR", "JOUR")
        copy_sample(sample_nuit, "MODELE-NUIT", "NUIT")
        copy_sample(sample_wk, "MODELE-WK", "WK")
        src.close()
    except Exception as e:
        import traceback; traceback.print_exc()
        print(f"   ! 04_POINTAGE: {e}")

    safe_save(wb, TPL_DIR / tpl_name)
    print(f"  OK {tpl_name:38s} — {written} cellules sur {len(tabs)} onglets")
    STATS["files"] += 1
    STATS["cells_total"] += written


# ---------------------------------------------------------------------------
# 05A CA FLASH
# ---------------------------------------------------------------------------
def prefill_ca_flash():
    tpl_name = "05A_CA_FLASH_v1.xlsx"
    wb = open_tpl(tpl_name)
    written = 0
    tabs = set()
    try:
        src = open_src("09_CA_MAI.xlsx")

        # === Récap Global ===
        # Source: R3 headers=[None, None, 'Base O&G', 'Import Maritime', 'Hinterland', 'Aérien', 'Export Maritime', 'Afritramp', 'Meet & Greet', 'SIRT / LEVAGE', 'TOTAL GENERE', None]
        # Source: R4+ data: [None, 'Janvier', BASE, IMAR, HINT, AERIEN, EXPORT, Afritramp, Meet&Greet, SIRT, TOTAL]
        # Template Récap Global headers: A=Type, B=Mois, C=BASE LOYER, D=Plog&Divers, E=Total BASE,
        #   F=IMAR, G=HINT, H=IAER, I=EAER, J=AÉRIEN, K=EXPORT, L=SIRT, M=Afritramp, N=Meet&Greet,
        #   O=CA GENERE Autres, P=CA TOTAL GENERE
        sws = src["Récap Global"]
        tws = wb["Récap Global"]
        out_row = 2
        for r in range(4, sws.max_row + 1):
            mois = sws.cell(row=r, column=2).value
            if not isinstance(mois, str):
                continue
            base_og = sws.cell(row=r, column=3).value
            imar = sws.cell(row=r, column=4).value
            hint = sws.cell(row=r, column=5).value
            aerien = sws.cell(row=r, column=6).value
            export = sws.cell(row=r, column=7).value
            afritramp = sws.cell(row=r, column=8).value
            meet = sws.cell(row=r, column=9).value
            sirt = sws.cell(row=r, column=10).value
            total = sws.cell(row=r, column=11).value
            w = 0
            w += set_cell(tws, out_row, 1, "Mois")
            w += set_cell(tws, out_row, 2, mois)
            w += set_cell(tws, out_row, 5, base_og, "#,##0")  # Total BASE
            w += set_cell(tws, out_row, 6, imar, "#,##0")
            w += set_cell(tws, out_row, 7, hint, "#,##0")
            w += set_cell(tws, out_row, 10, aerien, "#,##0")
            w += set_cell(tws, out_row, 11, export, "#,##0")
            w += set_cell(tws, out_row, 12, sirt, "#,##0")
            w += set_cell(tws, out_row, 13, afritramp, "#,##0")
            w += set_cell(tws, out_row, 14, meet, "#,##0")
            w += set_cell(tws, out_row, 16, total, "#,##0")
            if w:
                out_row += 1
                written += w
                tabs.add("Récap Global")

        # === Cumul Annuel ===
        # Source: R3 headers=[None, 'LOYERS', 'Plog & Divers', 'Total BASE', 'IMAR', 'HINT', 'IAER', 'EAER', 'AERIEN', 'EXPORT', 'SIRT (dédié)', 'CA Dpt Ext', ...]
        # Source: R4+ : [Client, LOYERS, Plog&Divers, Total BASE, ...]
        # Template Cumul Annuel headers: A=Client, B=Loyers, C=CA Généré (hors loyers), D=CA TOTAL GENERE,
        #   E=Réel N-1, F=Budget annuel, G=VAR (FCFA), H=VAR (%)
        sws = src["Cumul annuel"]
        tws = wb["Cumul Annuel"]
        out_row = 2
        for r in range(4, sws.max_row + 1):
            client = sws.cell(row=r, column=1).value
            if not isinstance(client, str) or not client.strip():
                continue
            if client.strip().upper() == "TOTAL":
                continue
            loyers = sws.cell(row=r, column=2).value
            total_base = sws.cell(row=r, column=4).value
            # CA Généré = sommes autres (utiliser CA Dpt Ext col 12 si dispo)
            ca_gen = sws.cell(row=r, column=12).value
            w = 0
            w += set_cell(tws, out_row, 1, client.strip())
            w += set_cell(tws, out_row, 2, loyers, "#,##0")
            w += set_cell(tws, out_row, 3, ca_gen, "#,##0")
            # Total = loyers + ca_gen + others — on prend total_base + ca_gen
            try:
                total_val = (loyers or 0) + (ca_gen or 0)
                w += set_cell(tws, out_row, 4, total_val, "#,##0")
            except Exception:
                pass
            if w:
                out_row += 1
                written += w
                tabs.add("Cumul Annuel")

        src.close()
    except Exception as e:
        print(f"   ! 05A_CA_FLASH: {e}")

    safe_save(wb, TPL_DIR / tpl_name)
    print(f"  OK {tpl_name:38s} — {written} cellules sur {len(tabs)} onglets")
    STATS["files"] += 1
    STATS["cells_total"] += written


# ---------------------------------------------------------------------------
# 05B CA PNL — extraction LONG depuis matrice 'Gestion base logistique'
# Format source : MATRICE complexe (35 lignes × 40 cols). On extrait au mieux
# les lignes lisibles et on dépivote ligne x mois en (Mois, Code, Libellé, ...).
# ---------------------------------------------------------------------------
def prefill_ca_pnl():
    tpl_name = "05B_CA_PNL_v1.xlsx"
    wb = open_tpl(tpl_name)
    ws = wb["PnL_Base"]
    written = 0
    try:
        src = open_src("09_CA_MAI.xlsx")
        sws = src["Gestion base logistique"]
        # Détecter ligne header (chercher 'Janvier' ou nom mois)
        mois_noms = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                     "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
        header_row = None
        mois_cols = {}  # mois -> col_index
        for r in range(1, min(sws.max_row + 1, 10)):
            for c in range(1, sws.max_column + 1):
                v = sws.cell(row=r, column=c).value
                if isinstance(v, str) and v.strip() in mois_noms:
                    header_row = r
                    mois_cols[v.strip()] = c
        if header_row is None:
            # rien à faire
            src.close()
            safe_save(wb, TPL_DIR / tpl_name)
            print(f"  OK {tpl_name:38s} — 0 cellules sur 0 onglets (matrice non détectée)")
            STATS["files"] += 1
            return

        # Parcourir les lignes data : col A = libellé poste
        out_row = 2
        idx_poste = 1
        for r in range(header_row + 1, sws.max_row + 1):
            libelle = sws.cell(row=r, column=1).value
            if not libelle:
                # tester col B
                libelle = sws.cell(row=r, column=2).value
            if not isinstance(libelle, str) or not libelle.strip():
                continue
            libelle = libelle.strip()
            # dépivoter sur les mois trouvés
            wrote_any = False
            for mois, col in mois_cols.items():
                v = sws.cell(row=r, column=col).value
                if v is None or v == 0:
                    continue
                w = 0
                w += set_cell(ws, out_row, 1, mois.upper())
                w += set_cell(ws, out_row, 2, f"P{idx_poste:03d}")
                w += set_cell(ws, out_row, 3, libelle)
                w += set_cell(ws, out_row, 4, "Charges")
                # Pas de distinction Prévu/Réel dans source matrice → mettre tout en Réel (col F)
                w += set_cell(ws, out_row, 6, v, "#,##0")
                if w:
                    out_row += 1
                    written += w
                    wrote_any = True
            if wrote_any:
                idx_poste += 1
        src.close()
    except Exception as e:
        print(f"   ! 05B_CA_PNL: {e}")

    safe_save(wb, TPL_DIR / tpl_name)
    print(f"  OK {tpl_name:38s} — {written} cellules sur 1 onglet")
    STATS["files"] += 1
    STATS["cells_total"] += written


# ---------------------------------------------------------------------------
# 05C CA DETAIL — copier positions Q/R/U (16/17/20) + col A/B
# ---------------------------------------------------------------------------
def prefill_ca_detail():
    tpl_name = "05C_CA_DETAIL_v1.xlsx"
    wb = open_tpl(tpl_name)
    written = 0
    tabs = 0
    try:
        src = open_src("09_CA_MAI.xlsx")
        # Mapping source mois name -> template mois name (mêmes noms)
        mois_map = {
            "Janvier": "Janvier", "Février": "Février", "Mars": "Mars",
            "Avril": "Avril", "Mai": "Mai", "Juin": "Juin",
            "Juillet": "Juillet", "Aout": "Août", "Août": "Août",
            "Septembre": "Septembre", "Octobre": "Octobre",
            "Novembre": "Novembre", "Décembre": "Décembre",
        }
        for src_name, tpl_name_sheet in mois_map.items():
            if src_name not in src.sheetnames:
                continue
            if tpl_name_sheet not in wb.sheetnames:
                continue
            sws = src[src_name]
            tws = wb[tpl_name_sheet]
            # Détecter header (R3 dans source ; on cherche 'REEL BASE 2026' ou similaire)
            header_row = None
            for r in range(1, 10):
                v = sws.cell(row=r, column=17).value
                if isinstance(v, str) and "REEL" in v.upper() and "2026" in v:
                    header_row = r
                    break
            if header_row is None:
                header_row = 3
            out_row = 2
            for r in range(header_row + 1, sws.max_row + 1):
                client = sws.cell(row=r, column=1).value
                if not isinstance(client, str) or not client.strip():
                    continue
                if client.strip().upper() in ("TOTAL", "TOTAL GENERAL"):
                    continue
                w = 0
                w += set_cell(tws, out_row, 1, client.strip())
                # Sources revenus (col 16 dans source = col P → idx 16)
                sources = sws.cell(row=r, column=16).value
                w += set_cell(tws, out_row, 2, sources)
                # Positions hardcodées Q=17, R=18, U=21 dans template
                # Source idx Q=17 (REEL BASE 2026), R=18 (Budget 2026), U=21 (Réel BASE 2025)
                reel26 = sws.cell(row=r, column=17).value
                budget26 = sws.cell(row=r, column=18).value
                reel25 = sws.cell(row=r, column=21).value
                w += set_cell(tws, out_row, 17, reel26, "#,##0")
                w += set_cell(tws, out_row, 18, budget26, "#,##0")
                w += set_cell(tws, out_row, 21, reel25, "#,##0")
                # VAR S = R-B, T = S/B
                try:
                    if isinstance(reel26, (int, float)) and isinstance(budget26, (int, float)):
                        var_fcfa = reel26 - budget26
                        w += set_cell(tws, out_row, 19, var_fcfa, "#,##0")
                        if budget26 != 0:
                            w += set_cell(tws, out_row, 20, var_fcfa / budget26, "0.00%")
                except Exception:
                    pass
                if w:
                    out_row += 1
                    written += w
            if out_row > 2:
                tabs += 1
        src.close()
    except Exception as e:
        print(f"   ! 05C_CA_DETAIL: {e}")

    safe_save(wb, TPL_DIR / tpl_name)
    print(f"  OK {tpl_name:38s} — {written} cellules sur {tabs} onglets")
    STATS["files"] += 1
    STATS["cells_total"] += written


# ---------------------------------------------------------------------------
# 06 ACCOSTAGE — 4 mois disponibles
# ---------------------------------------------------------------------------
def prefill_accostage():
    tpl_name = "06_ACCOSTAGE_v1.xlsx"
    wb = open_tpl(tpl_name)
    written = 0
    tabs = 0
    try:
        src = open_src("06_ACCOSTAGE.xlsx")
        # Source: 'JANVIER 26' etc. — template aussi
        # Source headers : R3 = [N°, DATE, NAVIRES, CLIENTS, ETA, ETD, NBRE HEURES, QUAI 25, QUAI 21, MONTANT, ..., GOTTWALD, MAFI]
        # R4 = sous-headers
        # R5+ = data
        # Source cols : 1=N°, 2=DATE, 3=NAVIRE, 4=CLIENT, 5=ETA, 6=ETD, 7=NBRE HEURES,
        #   8=QUAI25 (1 ou vide), 9=QUAI21 (1 ou vide), 10=Montant
        #   12=GOTTWALD MOB/H, 13=DEMOB/H, 14=Nbre H, 15=BL
        #   17=MAFI SHIFT JOUR, 18=NUIT, 19=TOTAL, 20=CLIENT, 21=BL
        # Template cols : A=N°, B=DATE, C=NAVIRE, D=CLIENT, E=ETA, F=ETD, G=NBRE HEURES,
        #   H=QUAI POSTE (Q25 ou Q21), I=MONTANT,
        #   J=GOTTWALD DATE, K=MOB/H, L=DEMOB/H, M=Nbre H, N=BL,
        #   O=MAFI DATE, P=SHIFT JOUR, Q=SHIFT NUIT, R=TOTAL, S=Obs
        for sname in src.sheetnames:
            if sname not in wb.sheetnames:
                continue
            sws = src[sname]
            tws = wb[sname]
            out_row = 2
            for r in range(5, sws.max_row + 1):
                num = sws.cell(row=r, column=1).value
                if num is None:
                    continue
                w = 0
                w += set_cell(tws, out_row, 1, num)
                # DATE
                d = sws.cell(row=r, column=2).value
                if isinstance(d, datetime):
                    w += set_cell(tws, out_row, 2, d.date(), "yyyy-mm-dd")
                elif d is not None:
                    w += set_cell(tws, out_row, 2, d)
                w += set_cell(tws, out_row, 3, sws.cell(row=r, column=3).value)  # NAVIRE
                w += set_cell(tws, out_row, 4, sws.cell(row=r, column=4).value)  # CLIENT
                # ETA / ETD
                eta = parse_time_str(sws.cell(row=r, column=5).value)
                etd = parse_time_str(sws.cell(row=r, column=6).value)
                if eta:
                    w += set_cell(tws, out_row, 5, eta, "hh:mm")
                else:
                    w += set_cell(tws, out_row, 5, sws.cell(row=r, column=5).value)
                if etd:
                    w += set_cell(tws, out_row, 6, etd, "hh:mm")
                else:
                    w += set_cell(tws, out_row, 6, sws.cell(row=r, column=6).value)
                # NBRE HEURES
                nb = sws.cell(row=r, column=7).value
                if isinstance(nb, str):
                    pt = parse_time_str(nb)
                    if pt:
                        w += set_cell(tws, out_row, 7, pt.hour + pt.minute / 60.0)
                    else:
                        w += set_cell(tws, out_row, 7, nb)
                else:
                    w += set_cell(tws, out_row, 7, nb)
                # QUAI POSTE : Q25 si col8=1, Q21 si col9=1
                q25 = sws.cell(row=r, column=8).value
                q21 = sws.cell(row=r, column=9).value
                quai = None
                if q25 == 1 or q25 == "1":
                    quai = "Q25"
                elif q21 == 1 or q21 == "1":
                    quai = "Q21"
                if quai:
                    w += set_cell(tws, out_row, 8, quai)
                # MONTANT
                w += set_cell(tws, out_row, 9, sws.cell(row=r, column=10).value, "#,##0")
                # GOTTWALD : pas de date isolée en source, on saisit MOB/DEMOB/Nbre H/BL
                gott_mob = parse_time_str(sws.cell(row=r, column=12).value)
                gott_demob = parse_time_str(sws.cell(row=r, column=13).value)
                if gott_mob:
                    w += set_cell(tws, out_row, 11, gott_mob, "hh:mm")
                if gott_demob:
                    w += set_cell(tws, out_row, 12, gott_demob, "hh:mm")
                w += set_cell(tws, out_row, 13, sws.cell(row=r, column=14).value)
                w += set_cell(tws, out_row, 14, sws.cell(row=r, column=15).value)
                # MAFI
                w += set_cell(tws, out_row, 16, sws.cell(row=r, column=17).value)
                w += set_cell(tws, out_row, 17, sws.cell(row=r, column=18).value)
                w += set_cell(tws, out_row, 18, sws.cell(row=r, column=19).value)
                if w:
                    out_row += 1
                    written += w
            if out_row > 2:
                tabs += 1
        src.close()
    except Exception as e:
        print(f"   ! 06_ACCOSTAGE: {e}")

    safe_save(wb, TPL_DIR / tpl_name)
    print(f"  OK {tpl_name:38s} — {written} cellules sur {tabs} onglets")
    STATS["files"] += 1
    STATS["cells_total"] += written


# ---------------------------------------------------------------------------
# 07 QHSE — 1ère semaine source → sous-onglets thématiques
# ---------------------------------------------------------------------------
def prefill_qhse():
    tpl_name = "07_QHSE_v1.xlsx"
    wb = open_tpl(tpl_name)
    written = 0
    tabs = set()
    try:
        src = open_src("07_QHSE.xlsx")
        sws = src[src.sheetnames[0]]

        # === Entete ===
        # Source R5 : Pays D=Côte d'ivoire, Entité G7=AFRICA GLOBAL LOGISTICS, Semaine N=18
        tws = wb["Entete"]
        pays = sws.cell(row=5, column=4).value
        entite = sws.cell(row=5, column=7).value
        semaine = sws.cell(row=5, column=13).value
        w = 0
        w += set_cell(tws, 2, 1, pays)
        w += set_cell(tws, 2, 2, entite)
        w += set_cell(tws, 2, 3, semaine)
        # Dates approximatives (semaine 18 2026 ≈ 27/04 - 03/05)
        w += set_cell(tws, 2, 4, date(2026, 4, 27), "yyyy-mm-dd")
        w += set_cell(tws, 2, 5, date(2026, 5, 3), "yyyy-mm-dd")
        rédacteur = sws.cell(row=9, column=10).value
        w += set_cell(tws, 2, 6, rédacteur)
        if w:
            written += w
            tabs.add("Entete")

        # === Météo (L9-10 source : header + valeurs) ===
        tws = wb["QHSE_Meteo"]
        # Pré-rempli : 7 jours en col A. On met les valeurs (Ensoleillé/Nuageux/Venteux/Pluie)
        # de la ligne 10 source pour le 1er jour seulement (le source ne distingue pas par jour)
        meteo_vals = [sws.cell(row=10, column=c).value for c in [3, 4, 5, 6]]
        for day_row in range(2, 9):
            w = 0
            for ci, v in enumerate(meteo_vals, start=2):
                if v:
                    w += set_cell(tws, day_row, ci, v)
            written += w
            if w:
                tabs.add("QHSE_Meteo")

        # === Personnel (L16-18 source) ===
        tws = wb["QHSE_Personnel"]
        # Template pré-rempli avec EFFECTIFS noms ; on remplit CADRES (B) / NON CADRES (C)
        for src_r, tgt_r in [(16, 2), (17, 3), (18, 4)]:
            cadres = sws.cell(row=src_r, column=4).value
            non_cadres = sws.cell(row=src_r, column=5).value
            w = 0
            if isinstance(cadres, (int, float)):
                w += set_cell(tws, tgt_r, 2, cadres)
            if isinstance(non_cadres, (int, float)):
                w += set_cell(tws, tgt_r, 3, non_cadres)
            written += w
            if w:
                tabs.add("QHSE_Personnel")

        # === Equipements (L27-50 source) ===
        tws = wb["QHSE_Equipements"]
        out_row = 2
        for r in range(27, 60):
            typ = sws.cell(row=r, column=3).value
            ton = sws.cell(row=r, column=4).value
            nb = sws.cell(row=r, column=5).value
            cert = sws.cell(row=r, column=6).value
            if not typ:
                continue
            w = 0
            w += set_cell(tws, out_row, 1, f"EQ-{out_row:03d}")
            w += set_cell(tws, out_row, 2, str(typ).strip())
            w += set_cell(tws, out_row, 3, str(ton).strip() if ton else "")
            w += set_cell(tws, out_row, 4, nb)
            w += set_cell(tws, out_row, 5, str(cert).strip() if cert else "")
            if w:
                out_row += 1
                written += w
                tabs.add("QHSE_Equipements")

        # === Incidents (L78-126 source col C/E) ===
        tws = wb["QHSE_Incidents"]
        out_row = 2
        for r in range(78, 140):
            cat = sws.cell(row=r, column=3).value
            nb = sws.cell(row=r, column=5).value
            if not isinstance(cat, str) or not cat.strip():
                continue
            # filtrer headers/totals
            cat_clean = cat.strip()
            if cat_clean.upper() in ("ACCIDENTS / INCIDENTS", "REPORTING", "MEDICAL", "MÉDICAL", "DIVERS", "TOTAL", "COMMENTAIRES HSE", "SITUATIONS DANGEREUSES"):
                continue
            w = 0
            w += set_cell(tws, out_row, 1, cat_clean)
            if isinstance(nb, (int, float)):
                w += set_cell(tws, out_row, 2, nb)
            if w:
                out_row += 1
                written += w
                tabs.add("QHSE_Incidents")
        # Ajouter nb jours sans accident en G2
        try:
            jsa = sws.cell(row=86, column=5).value
            if isinstance(jsa, (int, float)):
                w = set_cell(tws, 2, 7, jsa)
                written += w
        except Exception:
            pass

        # === Securite (L78-126 col J/L) ===
        tws = wb["QHSE_Securite"]
        out_row = 2
        for r in range(78, 128):
            act = sws.cell(row=r, column=10).value
            nb = sws.cell(row=r, column=12).value
            if not isinstance(act, str) or not act.strip():
                continue
            act_clean = act.strip()
            if act_clean.upper() in ("ACTIVITÉS HSE", "PARCOURS D'ACCUEIL", "PERSONNEL FORMÉ",
                                       "HABILITATIONS / FORMATIONS AU POSTE DE TRAVAIL",
                                       "EXTINCTEURS", "LOCALISATION", "NOMBRE"):
                continue
            w = 0
            w += set_cell(tws, out_row, 1, act_clean)
            if isinstance(nb, (int, float)):
                w += set_cell(tws, out_row, 2, nb)
            if w:
                out_row += 1
                written += w
                tabs.add("QHSE_Securite")

        # === Risques (L149-167 col C/E) ===
        tws = wb["QHSE_Risques"]
        out_row = 2
        for r in range(149, 170):
            fam = sws.cell(row=r, column=3).value
            nb = sws.cell(row=r, column=5).value
            if not isinstance(fam, str) or not fam.strip():
                continue
            fam_clean = fam.strip()
            if fam_clean.upper() in ("BONNE PERFORMANCE", "MAUVAISE PERFORMANCE"):
                continue
            w = 0
            w += set_cell(tws, out_row, 1, fam_clean)
            if isinstance(nb, (int, float)):
                w += set_cell(tws, out_row, 2, nb)
            if w:
                out_row += 1
                written += w
                tabs.add("QHSE_Risques")

        # === Actions à valoriser (L144+ col D) ===
        tws = wb["QHSE_Actions"]
        out_row = 2
        for r in range(144, 150):
            action = sws.cell(row=r, column=4).value
            if isinstance(action, str) and action.strip():
                w = 0
                w += set_cell(tws, out_row, 1, action.strip())
                w += set_cell(tws, out_row, 4, "Planifiée")
                w += set_cell(tws, out_row, 5, "Moyenne")
                if w:
                    out_row += 1
                    written += w
                    tabs.add("QHSE_Actions")

        src.close()
    except Exception as e:
        import traceback; traceback.print_exc()
        print(f"   ! 07_QHSE: {e}")

    safe_save(wb, TPL_DIR / tpl_name)
    print(f"  OK {tpl_name:38s} — {written} cellules sur {len(tabs)} onglets")
    STATS["files"] += 1
    STATS["cells_total"] += written


# ---------------------------------------------------------------------------
# 08 ACTIONS — 3 semaines source vers 3 onglets
# ---------------------------------------------------------------------------
def prefill_actions():
    tpl_name = "08_ACTIONS_v1.xlsx"
    wb = open_tpl(tpl_name)
    written = 0
    tabs = 0
    try:
        src = open_src("08_ACTIONS.xlsx")
        # Mapping source sheet -> template sheet (3 first)
        src_sheets = src.sheetnames[:3]
        tgt_sheets = ["2026-S01", "2026-S02", "2026-S03"]
        for sname, tname in zip(src_sheets, tgt_sheets):
            sws = src[sname]
            tws = wb[tname]
            out_row = 2
            n = 0
            current_section = ""
            for r in range(2, sws.max_row + 1):
                a = sws.cell(row=r, column=1).value
                b = sws.cell(row=r, column=2).value
                if a and isinstance(a, str) and a.strip():
                    current_section = a.strip().rstrip()
                if not b or not isinstance(b, str):
                    continue
                action = b.strip().replace("\xa0", " ")
                # Ignorer actions vides comme "Act N : "
                m = re.match(r"^[·\.\s]*(Act|Proj|Rec|Presque accident|Dysfonctionnement|Incident)\s*\d+\s*:\s*(.*)$", action, re.IGNORECASE)
                if m:
                    body = m.group(2).strip()
                    if not body:
                        continue
                w = 0
                n += 1
                w += set_cell(tws, out_row, 1, n)
                w += set_cell(tws, out_row, 2, current_section)
                w += set_cell(tws, out_row, 3, action)
                w += set_cell(tws, out_row, 4, "Ouverte")
                w += set_cell(tws, out_row, 8, "Moyenne")
                if w:
                    out_row += 1
                    written += w
            if out_row > 2:
                tabs += 1
        src.close()
    except Exception as e:
        print(f"   ! 08_ACTIONS: {e}")

    safe_save(wb, TPL_DIR / tpl_name)
    print(f"  OK {tpl_name:38s} — {written} cellules sur {tabs} onglets")
    STATS["files"] += 1
    STATS["cells_total"] += written


# ---------------------------------------------------------------------------
# Validation finale : ré-ouvrir chaque template sans warning
# ---------------------------------------------------------------------------
def validate_all():
    print("\n=== Validation : ré-ouverture des templates ===")
    for f in sorted(TPL_DIR.glob("*.xlsx")):
        try:
            with warnings.catch_warnings(record=True) as w:
                warnings.simplefilter("always")
                wb = openpyxl.load_workbook(f)
                n_sheets = len(wb.sheetnames)
                wb.close()
                w_filtered = [x for x in w if "Data Validation" not in str(x.message)]
                tag = "OK" if not w_filtered else f"WARN ({len(w_filtered)})"
            size_ko = f.stat().st_size / 1024
            print(f"  {tag:6s} {f.name:40s}  {size_ko:7.1f} Ko  {n_sheets:>2} onglets")
        except Exception as e:
            print(f"  KO   {f.name} : {e}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    print("=== Pré-remplissage des 11 templates AGL ===\n")
    sizes_before = {f.name: f.stat().st_size for f in TPL_DIR.glob("*.xlsx")}

    steps = [
        prefill_master,
        prefill_loyers,
        prefill_charges,
        prefill_planning,
        prefill_pointage,
        prefill_ca_flash,
        prefill_ca_pnl,
        prefill_ca_detail,
        prefill_accostage,
        prefill_qhse,
        prefill_actions,
    ]
    for fn in steps:
        try:
            fn()
        except Exception as e:
            import traceback
            traceback.print_exc()
            print(f"   !! Erreur dans {fn.__name__}: {e}")

    # Recap tailles
    print("\n=== Tailles avant/après ===")
    total_before = 0
    total_after = 0
    for f in sorted(TPL_DIR.glob("*.xlsx")):
        before = sizes_before.get(f.name, 0)
        after = f.stat().st_size
        total_before += before
        total_after += after
        delta = (after - before) / 1024
        sign = "+" if delta >= 0 else ""
        print(f"  {f.name:40s}  {before/1024:7.1f} → {after/1024:7.1f} Ko  ({sign}{delta:+.1f} Ko)")
    print(f"  {'TOTAL':40s}  {total_before/1024:7.1f} → {total_after/1024:7.1f} Ko")

    validate_all()

    print(f"\nOK — {STATS['files']} templates pré-remplis dans audit/templates/output_prefilled/")
    print(f"Total cellules remplies : {STATS['cells_total']}")


if __name__ == "__main__":
    main()
