"""
Phase 3 — Génération des 10 templates Excel "stricts" AGL.

Conformément à DESIGN.md (748 lignes) :
- Police : Calibri 11 corps / Bold 12 en-têtes / Bold 14 titres mode d'emploi
- Couleurs : Navy #0A2540 fond en-têtes, texte blanc, bordures Gold #C9A14A
- Champs obligatoires : fond gold #FFF4CC sur cellule en-tête
- Ligne 1 verrouillée (mot de passe AGL-OG-2026), data déverrouillée
- Cellules figées : B2 (ligne 1 + col A)
- Format dates yyyy-mm-dd, heures hh:mm, montants #,##0
- Onglet 📖 Mode d'emploi en position 1
- Onglet _SYS caché (listes nommées par fichier — pas d'inter-fichier)
- Versioning : _SYS!A1 = "v1.0 — <isoformat>"

DÉCISIONS DE CONCEPTION (cas ambigus DESIGN.md résolus ici) :
1. Les références listes déroulantes pointent toutes vers _SYS local (jamais inter-fichier)
   même quand DESIGN dit "Master.Clients" — pour stabilité au déplacement de fichier.
   Le Master existe comme référentiel humain ; pour les listes déroulantes opérationnelles
   on duplique les valeurs dans _SYS de chaque fichier (acceptable car volumes faibles).
2. Pour CHARGES, ACCOSTAGE, CA_DETAIL : on génère les 12 onglets mensuels (JAN..DEC).
3. PLANNING / POINTAGE : on génère 2 onglets MODELE-JOUR / MODELE-NUIT (et MODELE-WK pour POINTAGE)
   plutôt que de pré-générer 365 jours.
4. ACTIONS : on génère 4 onglets exemples 2026-S01..S04.
5. QHSE : 1 fichier = 1 modèle hebdo (semaine type), sous-onglets pré-créés.
6. Pour Master référentiels : pré-rempli avec valeurs initiales documentées DESIGN §1.
7. _SYS A1 contient le tag version ; les listes nommées démarrent ligne 3+ (col A=Clients, B=Espaces, etc.).
8. Largeurs colonnes : auto-fit raisonnable, max 30.
"""
from __future__ import annotations

import sys
from datetime import datetime, date
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Protection, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------------------
# Constantes de style
# ---------------------------------------------------------------------------
OUT_DIR = Path(__file__).parent / "output"
OUT_DIR.mkdir(parents=True, exist_ok=True)

NAVY = "FF0A2540"
GOLD = "FFC9A14A"
GOLD_LIGHT = "FFFFF4CC"
WHITE = "FFFFFFFF"
GREY_LIGHT = "FFF2F2F2"

PASSWORD = "AGL-OG-2026"
VERSION_TAG = f"v1.0 — {datetime.today().date().isoformat()}"

FONT_BODY = Font(name="Calibri", size=11)
FONT_HEADER = Font(name="Calibri", size=12, bold=True, color=WHITE)
FONT_TITLE = Font(name="Calibri", size=14, bold=True, color=NAVY)
FONT_SUBTITLE = Font(name="Calibri", size=12, bold=True, color=NAVY)

FILL_HEADER = PatternFill("solid", fgColor=NAVY)
FILL_HEADER_REQ = PatternFill("solid", fgColor=GOLD)  # en-tête obligatoire = fond gold
FILL_GREY = PatternFill("solid", fgColor=GREY_LIGHT)
FILL_GOLD_LIGHT = PatternFill("solid", fgColor=GOLD_LIGHT)

THIN_GOLD = Side(style="thin", color=GOLD)
BORDER_HEADER = Border(left=THIN_GOLD, right=THIN_GOLD, top=THIN_GOLD, bottom=THIN_GOLD)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

PROT_UNLOCKED = Protection(locked=False)
PROT_LOCKED = Protection(locked=True)

# ---------------------------------------------------------------------------
# Données de référence (valeurs initiales DESIGN §1)
# ---------------------------------------------------------------------------
CLIENTS_INITIAUX = [
    "FOXTROT", "MURPHY", "SLB", "SAIPEM SEI", "SAIPEM OM", "EXPRO", "IVOS",
    "DRIL QUIP", "HALLIBURTON", "GEV", "BAKER HUGHES", "BOSKALIS", "NAVITRANS",
    "AFRITRAMP", "CNR", "PAA", "FRENCH ARMY", "ANTRAK GHANA", "ANADARKO",
]

# Extrait de LOYERS.Feuil1 (échantillon documenté)
ESPACES_INITIAUX = [
    # (id, libelle, type, superficie_m2, client_dedie, tarif_m2)
    ("MAG-T10", "Magasin T10", "Magasin", 4500, "FOXTROT", 6000),
    ("MAG-T20", "Magasin T20", "Magasin", 3200, "MURPHY", 6000),
    ("YARD-3262", "Yard 3262 m²", "Yard", 3262, "FOXTROT", 4500),
    ("YARD-OPS", "Yard Opérations", "Yard", 5000, "", 4500),
    ("BUR-1", "Bureau Base 1", "Bureau", 250, "", 12000),
    ("ATEL-MAINT", "Atelier Maintenance", "Atelier", 800, "", 7500),
    ("PROT-MURPHY", "Protocole MURPHY", "Protocole", 1500, "MURPHY", 6000),
    ("PROT-FOXTROT", "Protocole FOXTROT", "Protocole", 2200, "FOXTROT", 6000),
    ("MIXTE-01", "Zone mixte 01", "Mixte", 1000, "", 5500),
]

EQUIPEMENTS_INITIAUX = [
    ("CE-16T-01", "CHARIOT ELEVATEUR", "16T", 3, "Certifié", date(2026, 9, 30), "402301"),
    ("CE-13T-01", "CHARIOT ELEVATEUR", "13T", 2, "Certifié", date(2026, 8, 15), "402302"),
    ("CE-12T-01", "CHARIOT ELEVATEUR", "12T", 2, "Étalonné", date(2026, 7, 20), "402303"),
    ("CE-7T-01",  "CHARIOT ELEVATEUR", "7T",  4, "Certifié", date(2026, 10, 5), "402304"),
    ("CE-3T-01",  "CHARIOT ELEVATEUR", "3T",  5, "Certifié", date(2026, 11, 12), "402305"),
    ("GOTTWALD-01", "GOTTWALD", "Autre", 1, "Certifié", date(2026, 12, 1), "402401"),
    ("MAFI-01", "MAFI", "45T", 2, "Certifié", date(2026, 6, 30), "402501"),
    ("TRACT-01", "TRACTEUR", "Autre", 3, "Certifié", date(2026, 9, 1), "402601"),
    ("GRUE-25T", "GRUE MOBILE", "25T", 1, "En cours", date(2026, 5, 30), "402701"),
    ("CAM-01", "CAMION", "Autre", 2, "Certifié", date(2026, 8, 1), "402801"),
]

SECTIONS_PLANNING = [
    ("CWI", "Container Wash & Inspection"),
    ("OIC", "Operations In Charge"),
    ("TRS", "Transit / Transport"),
    ("OFE", "Off-shore Equipment"),
    ("MAG", "Magasin"),
    ("YARD", "Yard"),
    ("QUAI", "Quai"),
]

NATURES_OPERATIONS = [
    ("NAT-01", "Activités en magasin et sur yard", "Manutention"),
    ("NAT-02", "Chargement / déchargement camion", "Manutention"),
    ("NAT-03", "Inspection conteneur", "Inspection"),
    ("NAT-04", "Lavage conteneur (CWI)", "Lavage"),
    ("NAT-05", "Transit douanier", "Douane"),
    ("NAT-06", "Levage offshore", "Levage"),
    ("NAT-07", "Stockage produits chimiques", "Stockage"),
    ("NAT-08", "Maintenance préventive", "Maintenance"),
]

AFFECTATIONS = [
    ("40B00", "Base logistique commune"),
    ("40B01", "Magasin T10"),
    ("40B02", "Magasin T20"),
    ("40Y01", "Yards"),
    ("40Q01", "Quai 25"),
    ("40Q02", "Quai 21"),
    ("40A01", "Atelier mécanique"),
    ("40S01", "Sous-traitance générale"),
]

FOURNISSEURS = [
    ("F-AYTHAM", "AYTHAM SERVICES", "Virement"),
    ("F-LOC", "LOCATION GÉNÉRALE CI", "Virement"),
    ("F-CARB", "TOTAL ENERGIES CI", "Virement"),
    ("F-MAINT", "ATELIERS DE MAINTENANCE CI", "Chèque"),
    ("F-SECU", "SECURITAS CI", "Virement"),
    ("F-NETT", "NETTOYAGE PRO", "Espèces"),
    ("F-MAT", "MATÉRIEL O&G CI", "Virement"),
]

DEPARTEMENTS_CA = [
    "BASE LOYER", "Plog & Divers", "IMAR", "HINT", "IAER", "EAER",
    "AÉRIEN", "EXPORT", "SIRT", "Afritramp", "Meet&Greet", "Autres", "TOTAL",
]

POSTES_PNL = [
    "CAP NET", "Salaires", "Charges sociales", "Frais généraux", "Amortissements",
    "Assurances", "Impôts & taxes", "Sous-traitance", "Carburant", "Maintenance",
    "Loyers", "Autres charges", "RÉSULTAT D'EXPLOITATION",
]

QUAIS = [("Q25", "Quai 25"), ("Q21", "Quai 21")]

MOIS_FR = [
    "JANVIER", "FÉVRIER", "MARS", "AVRIL", "MAI", "JUIN",
    "JUILLET", "AOÛT", "SEPTEMBRE", "OCTOBRE", "NOVEMBRE", "DÉCEMBRE",
]
MOIS_FR_SHORT = ["01-26", "02-26", "03-26", "04-26", "05-26", "06-26",
                  "07-26", "08-26", "09-26", "10-26", "11-26", "12-26"]
MOIS_FR_ACCOSTAGE = [f"{m} 26" for m in MOIS_FR]
MOIS_CA_DETAIL = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre",
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def safe_title(title: str) -> str:
    """Tronque/sanitize les noms d'onglets Excel (max 31, sans :\\/?*[])."""
    bad = ':\\/?*[]'
    t = "".join("_" if c in bad else c for c in title)
    return t[:31]


def style_header_row(ws, headers, required_idx=None):
    """
    headers : liste de libellés
    required_idx : set d'indices 1-based qui doivent être en gold (obligatoires)
    """
    required_idx = required_idx or set()
    for col_idx, label in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.font = FONT_HEADER
        # Si obligatoire : fond gold avec texte navy (sinon navy/blanc standard)
        if col_idx in required_idx:
            cell.fill = FILL_HEADER_REQ
            cell.font = Font(name="Calibri", size=12, bold=True, color=NAVY)
        else:
            cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER
        cell.protection = PROT_LOCKED
    ws.row_dimensions[1].height = 32


def apply_column_widths(ws, headers):
    for i, h in enumerate(headers, start=1):
        w = max(12, min(30, len(str(h)) + 4))
        ws.column_dimensions[get_column_letter(i)].width = w


def unlock_data_cells(ws, n_rows=200, n_cols=None):
    """Déverrouille les cellules de saisie (ligne 2 -> n_rows)."""
    if n_cols is None:
        n_cols = ws.max_column or 1
    for r in range(2, n_rows + 1):
        for c in range(1, n_cols + 1):
            ws.cell(row=r, column=c).protection = PROT_UNLOCKED


def freeze_and_protect(ws):
    ws.freeze_panes = "B2"
    ws.protection.sheet = True
    ws.protection.password = PASSWORD
    ws.protection.enable()


def add_mode_emploi(wb: Workbook, title: str, body_lines: list[str]) -> None:
    """Crée et place en position 0 un onglet '📖 Mode d'emploi'."""
    name = "📖 Mode d'emploi"
    if name in wb.sheetnames:
        ws = wb[name]
    else:
        ws = wb.create_sheet(name)
    ws.column_dimensions["A"].width = 110

    ws["A1"] = title
    ws["A1"].font = FONT_TITLE
    ws["A1"].alignment = ALIGN_LEFT
    ws["A2"] = f"Version : {VERSION_TAG}"
    ws["A2"].font = Font(name="Calibri", size=10, italic=True)

    row = 4
    for line in body_lines:
        c = ws.cell(row=row, column=1, value=line)
        c.font = FONT_BODY
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        ws.row_dimensions[row].height = max(18, 14 * (1 + line.count("\n")))
        row += 1

    # placer en première position
    wb.move_sheet(ws, offset=-(wb.sheetnames.index(name)))


def add_sys_sheet(wb: Workbook, lists: dict[str, list[str]]) -> None:
    """
    Crée un onglet _SYS caché contenant les listes nommées.
    lists : dict {nom_liste: [valeurs]}
    Les noms définis pointent sur _SYS!$X$2:$X$N.
    """
    name = "_SYS"
    if name in wb.sheetnames:
        ws = wb[name]
    else:
        ws = wb.create_sheet(name)
    ws["A1"] = VERSION_TAG
    ws["A1"].font = Font(name="Calibri", size=10, bold=True)

    # Chaque liste = une colonne
    col = 1
    for list_name, values in lists.items():
        col_letter = get_column_letter(col)
        # ligne 2 = libellé liste
        ws.cell(row=2, column=col, value=list_name).font = Font(
            name="Calibri", size=10, bold=True, color=NAVY
        )
        for i, v in enumerate(values, start=3):
            ws.cell(row=i, column=col, value=v)
        last_row = max(3, 2 + len(values))
        # définir un nom global pointant vers la plage
        ref = f"_SYS!${col_letter}$3:${col_letter}${last_row}"
        # nom sans caractères spéciaux
        clean_name = "".join(c if (c.isalnum() or c == "_") else "_" for c in list_name)
        if clean_name and not clean_name[0].isalpha():
            clean_name = "L_" + clean_name
        try:
            dn = DefinedName(clean_name, attr_text=ref)
            # supprimer s'il existe déjà
            if clean_name in wb.defined_names:
                del wb.defined_names[clean_name]
            wb.defined_names[clean_name] = dn
        except Exception:
            pass
        col += 1

    ws.sheet_state = "hidden"


def add_dv_list(ws, col_letter: str, values_ref_or_list, n_rows=200):
    """
    Ajoute une DataValidation type liste sur la colonne col_letter (lignes 2..n_rows).
    values_ref_or_list : soit une chaîne de référence (ex '=_SYS!$A$3:$A$22'), soit une liste -> inline.
    """
    if isinstance(values_ref_or_list, str):
        formula = values_ref_or_list
    else:
        # liste inline (limite Excel ~255 chars)
        joined = ",".join(str(v) for v in values_ref_or_list)
        formula = f'"{joined}"'
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    dv.error = "Valeur non autorisée — choisir dans la liste."
    dv.errorTitle = "Saisie invalide"
    dv.prompt = "Sélectionner une valeur dans la liste."
    dv.promptTitle = "Liste déroulante"
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}2:{col_letter}{n_rows}")


def add_dv_date(ws, col_letter: str, n_rows=200, min_date="2020-01-01"):
    dv = DataValidation(
        type="date",
        operator="greaterThanOrEqual",
        formula1=min_date,
        allow_blank=True,
    )
    dv.error = "Date invalide (>= 2020-01-01)."
    dv.errorTitle = "Date invalide"
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}2:{col_letter}{n_rows}")


def add_dv_time(ws, col_letter: str, n_rows=200):
    dv = DataValidation(type="time", allow_blank=True)
    dv.error = "Heure invalide (HH:MM)."
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}2:{col_letter}{n_rows}")


def add_dv_number(ws, col_letter: str, min_v=0, n_rows=200):
    dv = DataValidation(
        type="decimal",
        operator="greaterThanOrEqual",
        formula1=str(min_v),
        allow_blank=True,
    )
    dv.error = f"Valeur numérique >= {min_v} attendue."
    ws.add_data_validation(dv)
    dv.add(f"{col_letter}2:{col_letter}{n_rows}")


def set_col_format(ws, col_letter: str, number_format: str, n_rows=200):
    for r in range(2, n_rows + 1):
        ws[f"{col_letter}{r}"].number_format = number_format


# ---------------------------------------------------------------------------
# Fichier 00 — Master Référentiels
# ---------------------------------------------------------------------------
def build_master() -> Path:
    wb = Workbook()
    # Supprime le sheet par défaut, on créera tout
    default = wb.active
    wb.remove(default)

    # === Clients ===
    ws = wb.create_sheet("Clients")
    headers = ["ID Client", "Nom Client", "Raison sociale", "Segment", "Actif", "Alias (parsing)"]
    required = {1, 2, 4, 5}
    style_header_row(ws, headers, required)
    apply_column_widths(ws, headers)
    for i, c in enumerate(CLIENTS_INITIAUX, start=2):
        ws.cell(row=i, column=1, value=c)
        ws.cell(row=i, column=2, value=c)
        ws.cell(row=i, column=4, value="Forage" if c in ("FOXTROT", "MURPHY", "CNR", "ANADARKO") else "Service")
        ws.cell(row=i, column=5, value="OUI")
    add_dv_list(ws, "D", '"Forage,Production,Service,Sismique,Construction,Marine"')
    add_dv_list(ws, "E", '"OUI,NON"')
    unlock_data_cells(ws, n_rows=200, n_cols=len(headers))
    freeze_and_protect(ws)

    # === Espaces ===
    ws = wb.create_sheet("Espaces")
    headers = ["ID Espace", "Libellé", "Type", "Superficie (m²)", "Client dédié", "Tarif m²/mois (XOF)"]
    required = {1, 2, 3, 4}
    style_header_row(ws, headers, required)
    apply_column_widths(ws, headers)
    for i, e in enumerate(ESPACES_INITIAUX, start=2):
        for j, val in enumerate(e, start=1):
            ws.cell(row=i, column=j, value=val)
    add_dv_list(ws, "C", '"Magasin,Yard,Bureau,Atelier,Protocole,Mixte"')
    add_dv_number(ws, "D", 0.01)
    add_dv_number(ws, "F", 0)
    set_col_format(ws, "F", "#,##0")
    set_col_format(ws, "D", "#,##0")
    unlock_data_cells(ws, n_rows=200, n_cols=len(headers))
    freeze_and_protect(ws)

    # === Equipements ===
    ws = wb.create_sheet("Equipements")
    headers = ["ID Équipement", "Type", "Tonnage", "Nb disponible", "Status certification",
               "Prochain contrôle", "N° UM (immat.)"]
    required = {1, 2, 3, 4, 5}
    style_header_row(ws, headers, required)
    apply_column_widths(ws, headers)
    for i, e in enumerate(EQUIPEMENTS_INITIAUX, start=2):
        for j, val in enumerate(e, start=1):
            ws.cell(row=i, column=j, value=val)
    add_dv_list(ws, "B", '"CHARIOT ELEVATEUR,GOTTWALD,MAFI,TRACTEUR,GRUE MOBILE,CAMION,Autre"')
    add_dv_list(ws, "C", '"3T,7T,12T,13T,16T,25T,45T,Autre"')
    add_dv_number(ws, "D", 0)
    add_dv_list(ws, "E", '"Certifié,Étalonné,En cours,À réaliser"')
    add_dv_date(ws, "F")
    set_col_format(ws, "F", "yyyy-mm-dd")
    unlock_data_cells(ws, n_rows=200, n_cols=len(headers))
    freeze_and_protect(ws)

    # === Sections_Planning ===
    ws = wb.create_sheet("Sections_Planning")
    headers = ["Code", "Libellé long", "Actif"]
    style_header_row(ws, headers, {1, 2, 3})
    apply_column_widths(ws, headers)
    for i, (code, lib) in enumerate(SECTIONS_PLANNING, start=2):
        ws.cell(row=i, column=1, value=code)
        ws.cell(row=i, column=2, value=lib)
        ws.cell(row=i, column=3, value="OUI")
    add_dv_list(ws, "C", '"OUI,NON"')
    unlock_data_cells(ws, n_rows=100, n_cols=3)
    freeze_and_protect(ws)

    # === Natures_Operations ===
    ws = wb.create_sheet("Natures_Operations")
    headers = ["ID", "Libellé", "Famille"]
    style_header_row(ws, headers, {1, 2, 3})
    apply_column_widths(ws, headers)
    for i, n in enumerate(NATURES_OPERATIONS, start=2):
        for j, v in enumerate(n, start=1):
            ws.cell(row=i, column=j, value=v)
    add_dv_list(ws, "C", '"Manutention,Inspection,Lavage,Douane,Levage,Stockage,Maintenance,Autre"')
    unlock_data_cells(ws, n_rows=100, n_cols=3)
    freeze_and_protect(ws)

    # === Affectations ===
    ws = wb.create_sheet("Affectations")
    headers = ["Code", "Libellé"]
    style_header_row(ws, headers, {1, 2})
    apply_column_widths(ws, headers)
    for i, a in enumerate(AFFECTATIONS, start=2):
        for j, v in enumerate(a, start=1):
            ws.cell(row=i, column=j, value=v)
    unlock_data_cells(ws, n_rows=100, n_cols=2)
    freeze_and_protect(ws)

    # === Fournisseurs ===
    ws = wb.create_sheet("Fournisseurs")
    headers = ["ID", "Nom", "Mode paiement par défaut"]
    style_header_row(ws, headers, {1, 2})
    apply_column_widths(ws, headers)
    for i, f in enumerate(FOURNISSEURS, start=2):
        for j, v in enumerate(f, start=1):
            ws.cell(row=i, column=j, value=v)
    add_dv_list(ws, "C", '"Virement,Chèque,Espèces,Carte,Compensation,Autre"')
    unlock_data_cells(ws, n_rows=200, n_cols=3)
    freeze_and_protect(ws)

    # === Departements_CA ===
    ws = wb.create_sheet("Departements_CA")
    headers = ["Libellé département"]
    style_header_row(ws, headers, {1})
    apply_column_widths(ws, headers)
    for i, d in enumerate(DEPARTEMENTS_CA, start=2):
        ws.cell(row=i, column=1, value=d)
    unlock_data_cells(ws, n_rows=50, n_cols=1)
    freeze_and_protect(ws)

    # === Postes_Charges_PNL ===
    ws = wb.create_sheet("Postes_Charges_PNL")
    headers = ["Libellé poste"]
    style_header_row(ws, headers, {1})
    apply_column_widths(ws, headers)
    for i, p in enumerate(POSTES_PNL, start=2):
        ws.cell(row=i, column=1, value=p)
    unlock_data_cells(ws, n_rows=50, n_cols=1)
    freeze_and_protect(ws)

    # === Quais ===
    ws = wb.create_sheet("Quais")
    headers = ["Code", "Libellé"]
    style_header_row(ws, headers, {1, 2})
    apply_column_widths(ws, headers)
    for i, q in enumerate(QUAIS, start=2):
        ws.cell(row=i, column=1, value=q[0])
        ws.cell(row=i, column=2, value=q[1])
    unlock_data_cells(ws, n_rows=20, n_cols=2)
    freeze_and_protect(ws)

    # Mode d'emploi
    add_mode_emploi(
        wb,
        "Master Référentiels AGL — Mode d'emploi",
        [
            "Ce fichier est le référentiel partagé. Le KAM le met à jour 1×/mois.",
            "Ne pas modifier la structure des onglets.",
            "Pour ajouter une valeur : aller dans l'onglet concerné, ajouter une ligne en bas, sauvegarder.",
            "Les autres fichiers se synchronisent à l'ouverture.",
            "",
            "Onglets : Clients · Espaces · Equipements · Sections_Planning · Natures_Operations · "
            "Affectations · Fournisseurs · Departements_CA · Postes_Charges_PNL · Quais",
            "",
            "Mot de passe de protection des en-têtes : AGL-OG-2026",
        ],
    )

    # _SYS (caché)
    add_sys_sheet(wb, {
        "Clients": CLIENTS_INITIAUX,
        "Segments": ["Forage", "Production", "Service", "Sismique", "Construction", "Marine"],
        "OuiNon": ["OUI", "NON"],
    })

    path = OUT_DIR / "00_MASTER_REFERENTIELS.xlsx"
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# 01 LOYERS
# ---------------------------------------------------------------------------
def build_loyers() -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("Loyers")
    # Colonnes : A client, B id_espace, C libelle_facturation, D montant_ht_annuel,
    # E-P M1-M12 (dates Excel 1er du mois), Q commentaire
    headers = ["Client", "ID Espace", "Libellé facturation", "Montant H.T annuel (XOF)"]
    for m in range(1, 13):
        headers.append(date(2026, m, 1))  # dates Excel comme en-têtes
    headers.append("Commentaire")
    required = {1, 2, 3, 4}
    # On gère manuellement car header date doit être Date Excel
    for col_idx, label in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.font = FONT_HEADER
        if col_idx in required:
            cell.fill = FILL_HEADER_REQ
            cell.font = Font(name="Calibri", size=12, bold=True, color=NAVY)
        else:
            cell.fill = FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_HEADER
        cell.protection = PROT_LOCKED
        if isinstance(label, date):
            cell.number_format = "mmm yy"
    ws.row_dimensions[1].height = 32
    apply_column_widths(ws, [str(h) for h in headers])

    # _SYS local
    add_sys_sheet(wb, {
        "Clients": CLIENTS_INITIAUX,
        "Espaces": [e[0] for e in ESPACES_INITIAUX],
    })

    add_dv_list(ws, "A", "=_SYS!$A$3:$A$" + str(2 + len(CLIENTS_INITIAUX)))
    add_dv_list(ws, "B", "=_SYS!$B$3:$B$" + str(2 + len(ESPACES_INITIAUX)))
    add_dv_number(ws, "D", 1)
    set_col_format(ws, "D", "#,##0")
    for col_idx in range(5, 17):
        cl = get_column_letter(col_idx)
        add_dv_number(ws, cl, 0)
        set_col_format(ws, cl, "#,##0")

    unlock_data_cells(ws, n_rows=200, n_cols=len(headers))
    freeze_and_protect(ws)

    add_mode_emploi(
        wb,
        "LOYERS — Mode d'emploi",
        [
            "Le KAM saisit 1 ligne par contrat client × espace.",
            "Le montant annuel HT (col D) se ventile mensuellement en M1-M12 (col E-P, 1er de chaque mois).",
            "Si avenant en cours d'année, créer une 2e ligne avec libellé \"Avenant\" en commentaire.",
            "Les en-têtes M1-M12 sont des DATES Excel (pas du texte) — ne pas les modifier.",
            "",
            f"Mot de passe en-têtes : {PASSWORD}",
        ],
    )

    path = OUT_DIR / "01_LOYERS_v1.xlsx"
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# 02 CHARGES — 12 onglets mensuels
# ---------------------------------------------------------------------------
def build_charges() -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    headers = [
        "Libellé", "N° UM", "Fournisseur", "Affectation", "Montant H.T (XOF)",
        "Mois", "Statut", "Mode Paiement", "Date facture", "N° Facture", "Commentaire",
    ]
    required = {1, 3, 4, 5, 6, 7, 8}

    add_sys_sheet(wb, {
        "Fournisseurs": [f[1] for f in FOURNISSEURS],
        "Affectations": [a[0] for a in AFFECTATIONS],
        "Mois": MOIS_FR,
        "Statuts": ["À justifier", "Justifiée", "Validée", "Litigieuse", "Annulée"],
        "ModesPaiement": ["Virement", "Chèque", "Espèces", "Carte", "Compensation", "Autre"],
    })

    n_four = len(FOURNISSEURS)
    n_aff = len(AFFECTATIONS)
    n_mois = len(MOIS_FR)
    n_stat = 5
    n_mp = 6

    for mois_label in MOIS_FR_SHORT:
        ws = wb.create_sheet(safe_title(mois_label))
        style_header_row(ws, headers, required)
        apply_column_widths(ws, headers)

        # validations
        add_dv_list(ws, "C", f"=_SYS!$A$3:$A${2+n_four}")
        add_dv_list(ws, "D", f"=_SYS!$B$3:$B${2+n_aff}")
        add_dv_number(ws, "E", 1)
        set_col_format(ws, "E", "#,##0")
        add_dv_list(ws, "F", f"=_SYS!$C$3:$C${2+n_mois}")
        add_dv_list(ws, "G", f"=_SYS!$D$3:$D${2+n_stat}")
        add_dv_list(ws, "H", f"=_SYS!$E$3:$E${2+n_mp}")
        add_dv_date(ws, "I")
        set_col_format(ws, "I", "yyyy-mm-dd")

        unlock_data_cells(ws, n_rows=300, n_cols=len(headers))
        freeze_and_protect(ws)

    add_mode_emploi(
        wb,
        "CHARGES — Mode d'emploi",
        [
            "Une ligne = une charge. Saisir d'abord Libellé / Montant / Fournisseur / Affectation.",
            "Le statut \"Validée\" est requis avant tout paiement.",
            "Le N° UM permet de joindre avec l'équipement (Master.Equipements).",
            "Reproduire la même structure tous les mois — ne pas supprimer les colonnes F/G/H même si vides.",
            "Un onglet par mois (01-26 à 12-26).",
            "",
            f"Mot de passe en-têtes : {PASSWORD}",
        ],
    )

    path = OUT_DIR / "02_CHARGES_v1.xlsx"
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# 03 PLANNING — modèles JOUR & NUIT
# ---------------------------------------------------------------------------
def build_planning() -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    headers = [
        "N°", "CLIENT", "SECTION", "DEMANDE EMISE", "NATURE DE L'OPÉRATION",
        "POSITIF", "NEGA", "OBSERVATIONS", "Équipement principal", "Shift",
    ]
    required = {1, 2, 3, 4, 5, 6, 7, 10}

    add_sys_sheet(wb, {
        "Clients": CLIENTS_INITIAUX,
        "Sections": [s[0] for s in SECTIONS_PLANNING],
        "Natures": [n[1] for n in NATURES_OPERATIONS],
        "Equipements": [e[0] for e in EQUIPEMENTS_INITIAUX],
        "Shift": ["JOUR", "NUIT"],
    })

    n_cli = len(CLIENTS_INITIAUX)
    n_sec = len(SECTIONS_PLANNING)
    n_nat = len(NATURES_OPERATIONS)
    n_equip = len(EQUIPEMENTS_INITIAUX)

    for shift in ("JOUR", "NUIT"):
        sheet_name = f"MODELE-{shift}"
        ws = wb.create_sheet(safe_title(sheet_name))
        style_header_row(ws, headers, required)
        apply_column_widths(ws, headers)

        add_dv_number(ws, "A", 1)
        add_dv_list(ws, "B", f"=_SYS!$A$3:$A${2+n_cli}")
        add_dv_list(ws, "C", f"=_SYS!$B$3:$B${2+n_sec}")
        add_dv_list(ws, "E", f"=_SYS!$C$3:$C${2+n_nat}")
        add_dv_number(ws, "F", 0)
        add_dv_number(ws, "G", 0)
        add_dv_list(ws, "I", f"=_SYS!$D$3:$D${2+n_equip}")
        add_dv_list(ws, "J", '"JOUR,NUIT"')

        # pré-remplir shift colonne J ligne 2
        ws.cell(row=2, column=10, value=shift)

        unlock_data_cells(ws, n_rows=300, n_cols=len(headers))
        freeze_and_protect(ws)

    add_mode_emploi(
        wb,
        "PLANNING — Mode d'emploi",
        [
            "Chef de quart : 1 ligne = 1 opération.",
            "JOUR et NUIT ont la MÊME structure (10 colonnes).",
            "Saisir Client / Section / Demande émise / Nature obligatoirement.",
            "POSITIF = 1 si l'opération a abouti, NEGA = 1 si refusée/abandonnée.",
            "Observations = flags urgence/contrats.",
            "Pour la production : dupliquer l'onglet modèle et renommer DD-MM-YYYY-JOUR ou DD-MM-YYYY-NUIT.",
            "",
            f"Mot de passe en-têtes : {PASSWORD}",
        ],
    )

    path = OUT_DIR / "03_PLANNING_v1.xlsx"
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# 04 POINTAGE
# ---------------------------------------------------------------------------
def build_pointage() -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    # SOURCE_Personnel — 5 colonnes vides (Matricule, Nom, Fonction, Code Fonction, Chantier) + 2 lignes d'exemple
    ws_src = wb.create_sheet("SOURCE_Personnel")
    src_headers = ["Matricule", "Nom et Prénoms", "Fonction", "Code Fonction", "Chantier",
                   "Produit par défaut", "Actif"]
    style_header_row(ws_src, src_headers, {1, 2, 3, 4, 7})
    apply_column_widths(ws_src, src_headers)
    # 2 lignes exemple
    ws_src.cell(row=2, column=1, value="11580")
    ws_src.cell(row=2, column=2, value="JOHNSON Kouakouvi Nelson")
    ws_src.cell(row=2, column=3, value="CARISTE")
    ws_src.cell(row=2, column=4, value="C-CAR")
    ws_src.cell(row=2, column=5, value="FOXTROT-MAG")
    ws_src.cell(row=2, column=6, value="Manutention")
    ws_src.cell(row=2, column=7, value="OUI")
    ws_src.cell(row=3, column=1, value="11581")
    ws_src.cell(row=3, column=2, value="KONAN Yao Patrick")
    ws_src.cell(row=3, column=3, value="MAGASINIER")
    ws_src.cell(row=3, column=4, value="M-MAG")
    ws_src.cell(row=3, column=5, value="MURPHY-T20")
    ws_src.cell(row=3, column=6, value="Inspection")
    ws_src.cell(row=3, column=7, value="OUI")
    add_dv_list(ws_src, "C", '"CARISTE,MAGASINIER,GRUTIER,CHEF DE QUART,AGENT QHSE,DOCUMENTALISTE,Autre"')
    add_dv_list(ws_src, "G", '"OUI,NON"')
    unlock_data_cells(ws_src, n_rows=500, n_cols=len(src_headers))
    freeze_and_protect(ws_src)

    # _SYS
    add_sys_sheet(wb, {
        "Chantiers": ["FOXTROT-MAG", "MURPHY-T20", "BAKER-YARD", "SLB-OFE", "CWI-SHAN", "QUAI-25", "QUAI-21", "Autre"],
        "Shifts": ["JOUR", "NUIT"],
        "Validations": ["Validé", "Refusé", "À revoir"],
        "OuiNon": ["OUI", "NON"],
    })

    # Modèles journaliers (JOUR/NUIT/WK)
    daily_headers = [
        "N°", "Date", "Matricule", "Nom et Prénoms", "Fonction", "Code Fonction",
        "Chantiers/Activités", "Produits", "Shift", "Heures travaillées",
        "Validation chef", "Signature chef",
    ]
    required = {1, 2, 3, 4, 5, 6, 7, 9, 11}

    for shift in ("JOUR", "NUIT", "WK"):
        sheet_name = f"MODELE-{shift}"
        ws = wb.create_sheet(safe_title(sheet_name))
        style_header_row(ws, daily_headers, required)
        apply_column_widths(ws, daily_headers)

        add_dv_number(ws, "A", 1)
        add_dv_date(ws, "B")
        set_col_format(ws, "B", "yyyy-mm-dd")

        # Matricule : liste sur SOURCE_Personnel.A
        dv_mat = DataValidation(type="list",
                                formula1="=SOURCE_Personnel!$A$2:$A$501",
                                allow_blank=True)
        dv_mat.error = "Matricule absent de SOURCE_Personnel."
        ws.add_data_validation(dv_mat)
        dv_mat.add("C2:C300")

        # Formules VLOOKUP D/E/F → SOURCE_Personnel
        for r in range(2, 301):
            ws.cell(row=r, column=4,
                    value=f'=IFERROR(VLOOKUP(C{r},SOURCE_Personnel!$A$2:$G$501,2,FALSE),"")')
            ws.cell(row=r, column=5,
                    value=f'=IFERROR(VLOOKUP(C{r},SOURCE_Personnel!$A$2:$G$501,3,FALSE),"")')
            ws.cell(row=r, column=6,
                    value=f'=IFERROR(VLOOKUP(C{r},SOURCE_Personnel!$A$2:$G$501,4,FALSE),"")')

        add_dv_list(ws, "G", "=_SYS!$A$3:$A$10")
        add_dv_list(ws, "I", '"JOUR,NUIT,WK"')
        ws.cell(row=2, column=9, value=shift)
        add_dv_number(ws, "J", 0)
        add_dv_list(ws, "K", "=_SYS!$C$3:$C$5")

        unlock_data_cells(ws, n_rows=300, n_cols=len(daily_headers))
        freeze_and_protect(ws)

    add_mode_emploi(
        wb,
        "POINTAGE — Mode d'emploi",
        [
            "Le RH maintient l'onglet SOURCE_Personnel (matricule unique par agent).",
            "Le chef de chantier ne saisit que : Matricule (C), Chantier (G), Produit (H), Heures (J), Validation (K).",
            "Les colonnes D/E/F (Nom / Fonction / Code Fonction) se remplissent automatiquement via VLOOKUP",
            "interne au fichier (ciblage SOURCE_Personnel). Pas de fichier externe.",
            "Si le VLOOKUP ne renvoie rien, vérifier le matricule dans SOURCE_Personnel.",
            "Onglets MODELE-JOUR / MODELE-NUIT / MODELE-WK : à dupliquer et renommer DD-MM-YYYY-{JOUR|NUIT|WK}.",
            "",
            f"Mot de passe en-têtes : {PASSWORD}",
        ],
    )

    path = OUT_DIR / "04_POINTAGE_v1.xlsx"
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# 05A CA FLASH
# ---------------------------------------------------------------------------
def build_ca_flash() -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    # Récap Global
    ws = wb.create_sheet("Récap Global")
    headers = ["Type", "Mois", "BASE LOYER", "Plog & Divers", "Total BASE",
               "IMAR", "HINT", "IAER", "EAER", "AÉRIEN", "EXPORT", "SIRT",
               "Afritramp", "Meet&Greet", "CA GENERE Autres", "CA TOTAL GENERE"]
    required = {1, 2, 5, 16}
    style_header_row(ws, headers, required)
    apply_column_widths(ws, headers)

    add_sys_sheet(wb, {
        "Type": ["Mois", "Cumul", "Part %"],
        "Mois": MOIS_FR,
        "Clients": CLIENTS_INITIAUX,
    })

    add_dv_list(ws, "A", "=_SYS!$A$3:$A$5")
    add_dv_list(ws, "B", f"=_SYS!$B$3:$B${2+len(MOIS_FR)}")
    for col_letter in ["C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]:
        add_dv_number(ws, col_letter, 0)
        set_col_format(ws, col_letter, "#,##0")
    unlock_data_cells(ws, n_rows=200, n_cols=len(headers))
    freeze_and_protect(ws)

    # Cumul Annuel
    ws2 = wb.create_sheet("Cumul Annuel")
    headers2 = ["Client", "Loyers", "CA Généré (hors loyers)", "CA TOTAL GENERE",
                "Réel N-1", "Budget annuel", "VAR (FCFA)", "VAR (%)"]
    required2 = {1, 2, 3, 4}
    style_header_row(ws2, headers2, required2)
    apply_column_widths(ws2, headers2)

    add_dv_list(ws2, "A", f"=_SYS!$C$3:$C${2+len(CLIENTS_INITIAUX)}")
    for cl in ["B", "C", "D", "E", "F", "G"]:
        add_dv_number(ws2, cl, 0)
        set_col_format(ws2, cl, "#,##0")
    set_col_format(ws2, "H", "0.00%")
    unlock_data_cells(ws2, n_rows=200, n_cols=len(headers2))
    freeze_and_protect(ws2)

    add_mode_emploi(
        wb,
        "CA FLASH — Mode d'emploi",
        [
            "05A = vue pilotage rapide (mois × département + client).",
            "Onglet 'Récap Global' : matrice mensuelle (lignes = Mois / Cumul / Part %).",
            "Onglet 'Cumul Annuel' : 1 ligne par client (Loyers + CA Généré + comparaisons N-1/Budget).",
            "",
            f"Mot de passe en-têtes : {PASSWORD}",
        ],
    )

    path = OUT_DIR / "05A_CA_FLASH_v1.xlsx"
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# 05B CA PNL
# ---------------------------------------------------------------------------
def build_ca_pnl() -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("PnL_Base")
    headers = ["Mois", "Code poste", "Libellé poste", "Catégorie",
               "Prévu (XOF)", "Réel (XOF)", "Variation (XOF)", "Variation (%)", "Commentaire"]
    required = {1, 2, 3, 4, 5, 6, 7}
    style_header_row(ws, headers, required)
    apply_column_widths(ws, headers)

    add_sys_sheet(wb, {
        "Mois": MOIS_FR,
        "Postes": POSTES_PNL,
        "Categorie": ["Produits", "Charges", "Résultat"],
    })

    add_dv_list(ws, "A", f"=_SYS!$A$3:$A${2+len(MOIS_FR)}")
    add_dv_list(ws, "B", f"=_SYS!$B$3:$B${2+len(POSTES_PNL)}")
    add_dv_list(ws, "D", "=_SYS!$C$3:$C$5")
    for cl in ["E", "F", "G"]:
        set_col_format(ws, cl, "#,##0")
    set_col_format(ws, "H", "0.00%")
    unlock_data_cells(ws, n_rows=400, n_cols=len(headers))
    freeze_and_protect(ws)

    add_mode_emploi(
        wb,
        "CA PnL Base logistique — Mode d'emploi",
        [
            "05B = compte d'exploitation détaillé.",
            "Format LONG : 1 ligne = 1 poste × 1 mois × Prévu/Réel.",
            "Variation (G) = E - F (à calculer ou saisir).",
            "Variation % (H) = G/E (format 0.00%).",
            "",
            f"Mot de passe en-têtes : {PASSWORD}",
        ],
    )

    path = OUT_DIR / "05B_CA_PNL_v1.xlsx"
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# 05C CA DETAIL — 12 onglets mensuels, positions Q/R/U respectées (idx 16/17/20)
# ---------------------------------------------------------------------------
def build_ca_detail() -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    add_sys_sheet(wb, {"Clients": CLIENTS_INITIAUX})

    # Colonnes A..U (21 colonnes) avec positions hardcodées :
    # A=client, B=sources_revenus, ..., Q=reel_base_2026 (idx16), R=budget_2026 (idx17),
    # S=var_fcfa, T=var_pct, U=reel_base_2025 (idx20)
    headers = ["Client", "Sources revenus", "C", "D", "E", "F", "G", "H", "I", "J",
               "K", "L", "M", "N", "O", "P",
               "REEL BASE 2026 (XOF)", "Budget 2026 (XOF)", "VAR (FCFA)", "Var (%)",
               "REEL BASE 2025 (XOF)"]
    required = {1, 17, 18, 21}

    for mois_label in MOIS_CA_DETAIL:
        ws = wb.create_sheet(safe_title(mois_label))
        style_header_row(ws, headers, required)
        apply_column_widths(ws, headers)

        add_dv_list(ws, "A", f"=_SYS!$A$3:$A${2+len(CLIENTS_INITIAUX)}")
        add_dv_number(ws, "Q", 0)
        add_dv_number(ws, "R", 0)
        add_dv_number(ws, "U", 0)
        for cl in ["Q", "R", "S", "U"]:
            set_col_format(ws, cl, "#,##0")
        set_col_format(ws, "T", "0.00%")

        unlock_data_cells(ws, n_rows=200, n_cols=len(headers))
        freeze_and_protect(ws)

    add_mode_emploi(
        wb,
        "CA Détail mensuel — Mode d'emploi",
        [
            "05C = backup historique des détails clients (12 onglets mensuels).",
            "ATTENTION : les positions de colonnes Q (REEL BASE 2026 idx 16), R (Budget 2026 idx 17)",
            "et U (REEL BASE 2025 idx 20) sont HARDCODÉES côté parser. NE PAS DÉPLACER ces colonnes.",
            "",
            f"Mot de passe en-têtes : {PASSWORD}",
        ],
    )

    path = OUT_DIR / "05C_CA_DETAIL_v1.xlsx"
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# 06 ACCOSTAGE — 12 onglets mensuels avec blocs GOTTWALD + MAFI
# ---------------------------------------------------------------------------
def build_accostage() -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    headers = [
        "N°", "DATE", "NAVIRE", "CLIENT", "ETA", "ETD", "NBRE HEURES",
        "QUAI POSTE", "MONTANT (XOF)",
        # Bloc GOTTWALD
        "GOTTWALD DATE", "GOTTWALD MOB/H", "GOTTWALD DEMOB/H", "GOTTWALD Nbre H", "GOTTWALD BL",
        # Bloc MAFI
        "MAFI DATE", "MAFI SHIFT JOUR", "MAFI SHIFT NUIT", "MAFI TOTAL",
        "Observations",
    ]
    required = {1, 2, 3, 4, 5, 6, 7, 8, 9}

    add_sys_sheet(wb, {
        "Clients": CLIENTS_INITIAUX,
        "Quais": [q[0] for q in QUAIS],
    })

    for mois_label in MOIS_FR_ACCOSTAGE:
        ws = wb.create_sheet(safe_title(mois_label))
        style_header_row(ws, headers, required)
        apply_column_widths(ws, headers)

        add_dv_number(ws, "A", 1)
        add_dv_date(ws, "B")
        set_col_format(ws, "B", "yyyy-mm-dd")
        add_dv_list(ws, "D", f"=_SYS!$A$3:$A${2+len(CLIENTS_INITIAUX)}")
        add_dv_time(ws, "E"); set_col_format(ws, "E", "hh:mm")
        add_dv_time(ws, "F"); set_col_format(ws, "F", "hh:mm")
        add_dv_number(ws, "G", 0)
        add_dv_list(ws, "H", f"=_SYS!$B$3:$B${2+len(QUAIS)}")
        set_col_format(ws, "I", "#,##0")
        # bloc GOTTWALD
        add_dv_date(ws, "J"); set_col_format(ws, "J", "yyyy-mm-dd")
        add_dv_time(ws, "K"); set_col_format(ws, "K", "hh:mm")
        add_dv_time(ws, "L"); set_col_format(ws, "L", "hh:mm")
        add_dv_number(ws, "M", 0)
        # bloc MAFI
        add_dv_date(ws, "O"); set_col_format(ws, "O", "yyyy-mm-dd")
        add_dv_number(ws, "P", 0)
        add_dv_number(ws, "Q", 0)
        add_dv_number(ws, "R", 0)

        unlock_data_cells(ws, n_rows=200, n_cols=len(headers))
        freeze_and_protect(ws)

    add_mode_emploi(
        wb,
        "ACCOSTAGE — Mode d'emploi",
        [
            "Une ligne = une escale.",
            "Renseigner DATE / NAVIRE / CLIENT / ETA / ETD / QUAI / MONTANT.",
            "QUAI POSTE : obligatoire (Q25 ou Q21).",
            "Bloc GOTTWALD (J-N) : remplir uniquement si la grue Gottwald a été mobilisée.",
            "Bloc MAFI (O-R) : remplir uniquement si une remorque MAFI a été utilisée.",
            "MONTANT = FREE autorisé si gratuit (saisir 0 + commentaire).",
            "Un onglet par mois (JANVIER 26 → DÉCEMBRE 26).",
            "",
            f"Mot de passe en-têtes : {PASSWORD}",
        ],
    )

    path = OUT_DIR / "06_ACCOSTAGE_v1.xlsx"
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# 07 QHSE
# ---------------------------------------------------------------------------
def build_qhse() -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    add_sys_sheet(wb, {
        "Pays": ["Côte d'Ivoire", "Ghana", "Sénégal", "Cameroun"],
        "Jours": ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"],
        "OuiNon": ["OUI", "NON"],
        "Effectifs": ["AGL", "Sous-traitant", "Journaliers"],
        "Incidents": [
            "Accident avec arrêt", "Accident sans arrêt", "Incident", "Presque accident",
            "Premiers soins", "Maladie", "Déces", "Postes aménagés", "Situation dangereuse",
            "Déversement", "Départ de feu", "Evasan jour", "Evasan nuit",
            "Malaria/Paludisme", "Crise d'ulcère", "Pansement", "Divers"
        ],
        "Securite": [
            "Toolbox Meeting", "Briefing HSE", "Induction HSE", "Permis de travail",
            "SET CARD", "JSA", "Audit CIAPOL", "Check-list Camion",
            "Check-list Inspection Camion & Engins", "Check-list Opérations Navire",
            "Test d'alcoolémie", "Reporting", "Coaching Permis à Points",
            "Prévention Risques Routiers", "Formation Grutier", "Formation Cariste"
        ],
        "Risques": [
            "Mauvaise performance", "Chute ou projection", "Équipement sous pression",
            "Travail en hauteur", "Opération de levage", "Communication", "EPI",
            "Accessoires de levage", "Housekeeping", "Outils/Équipement",
            "Matières dangereuses", "Procédures & Documents", "Équipement électrique",
            "Zone de travail", "Instruments de mesure", "Manutention manuelle/Mécanisée"
        ],
        "Gravite": ["Faible", "Moyenne", "Grave", "Critique"],
        "Statuts": ["Planifiée", "En cours", "Réalisée", "En retard"],
        "Priorite": ["Basse", "Moyenne", "Haute", "Critique"],
        "Equipements": [e[0] for e in EQUIPEMENTS_INITIAUX],
        "Tonnage": ["3T", "7T", "12T", "13T", "16T", "25T", "45T", "Autre"],
        "StatusCert": ["Certifié", "Étalonné", "En cours", "À réaliser"],
    })

    def col_ref(letter, n):
        return f"=_SYS!${letter}$3:${letter}${2+n}"

    # === Entete ===
    ws = wb.create_sheet("Entete")
    headers = ["Pays", "Entité", "Semaine N°", "Date début", "Date fin",
               "Rédacteur QHSE", "Référence document"]
    style_header_row(ws, headers, {1, 2, 3, 4, 5, 6})
    apply_column_widths(ws, headers)
    add_dv_list(ws, "A", col_ref("A", 4))
    add_dv_number(ws, "C", 1)
    add_dv_date(ws, "D"); set_col_format(ws, "D", "yyyy-mm-dd")
    add_dv_date(ws, "E"); set_col_format(ws, "E", "yyyy-mm-dd")
    # Note parser : "semaine" en col L (idx 12) — on duplique en L1/M1 pour compat
    ws.cell(row=1, column=12, value="Semaine")
    ws.cell(row=1, column=12).font = FONT_HEADER
    ws.cell(row=1, column=12).fill = FILL_HEADER
    ws.cell(row=2, column=12, value="=C2")
    unlock_data_cells(ws, n_rows=10, n_cols=13)
    freeze_and_protect(ws)

    # === QHSE_Meteo ===
    ws = wb.create_sheet("QHSE_Meteo")
    headers = ["Jour", "Ensoleillé", "Nuageux", "Venteux", "Pluie", "Commentaire"]
    style_header_row(ws, headers, {1, 2, 3, 4, 5})
    apply_column_widths(ws, headers)
    add_dv_list(ws, "A", col_ref("B", 7))
    for c in ["B", "C", "D", "E"]:
        add_dv_list(ws, c, col_ref("C", 2))
    # pré-remplir 7 jours
    for i, j in enumerate(["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"], start=2):
        ws.cell(row=i, column=1, value=j)
    unlock_data_cells(ws, n_rows=20, n_cols=len(headers))
    freeze_and_protect(ws)

    # === QHSE_Personnel ===
    ws = wb.create_sheet("QHSE_Personnel")
    headers = ["EFFECTIFS", "CADRES", "NON CADRES", "TOTAL", "Déploiement Personnel QHSE"]
    style_header_row(ws, headers, {1, 2, 3, 4})
    apply_column_widths(ws, headers)
    add_dv_list(ws, "A", col_ref("D", 3))
    add_dv_number(ws, "B", 0)
    add_dv_number(ws, "C", 0)
    for i, eff in enumerate(["AGL", "Sous-traitant", "Journaliers"], start=2):
        ws.cell(row=i, column=1, value=eff)
        ws.cell(row=i, column=4, value=f"=B{i}+C{i}")
    unlock_data_cells(ws, n_rows=20, n_cols=len(headers))
    freeze_and_protect(ws)

    # === QHSE_Equipements ===
    ws = wb.create_sheet("QHSE_Equipements")
    headers = ["ID Équipement", "Type", "Tonnage", "Nb Disponible",
               "Status Certif./Étalonn.", "Prochain contrôle",
               "Affectation client", "Observations"]
    style_header_row(ws, headers, {1, 2, 3, 4, 5})
    apply_column_widths(ws, headers)
    add_dv_list(ws, "A", col_ref("M", len(EQUIPEMENTS_INITIAUX)))
    add_dv_list(ws, "C", col_ref("N", 8))
    add_dv_number(ws, "D", 0)
    add_dv_list(ws, "E", col_ref("O", 4))
    add_dv_date(ws, "F"); set_col_format(ws, "F", "yyyy-mm-dd")
    unlock_data_cells(ws, n_rows=50, n_cols=len(headers))
    freeze_and_protect(ws)

    # === QHSE_Incidents ===
    ws = wb.create_sheet("QHSE_Incidents")
    headers = ["Catégorie", "Nombre", "Description / Contexte", "Action correctrice",
               "Responsable suivi", "Nombre jours d'arrêt",
               "Nombre jours sans accident (cumul)", "Record N jours sans accident"]
    style_header_row(ws, headers, {1, 2, 7})
    apply_column_widths(ws, headers)
    add_dv_list(ws, "A", col_ref("E", 17))
    add_dv_number(ws, "B", 0)
    add_dv_number(ws, "F", 0)
    add_dv_number(ws, "G", 0)
    add_dv_number(ws, "H", 0)
    unlock_data_cells(ws, n_rows=50, n_cols=len(headers))
    freeze_and_protect(ws)

    # === QHSE_Securite ===
    ws = wb.create_sheet("QHSE_Securite")
    headers = ["Activité Q-HSE", "Nombre", "Durée (heures)", "Participants",
               "Animateur", "Observations"]
    style_header_row(ws, headers, {1, 2})
    apply_column_widths(ws, headers)
    add_dv_list(ws, "A", col_ref("F", 16))
    add_dv_number(ws, "B", 0)
    add_dv_number(ws, "C", 0)
    add_dv_number(ws, "D", 0)
    unlock_data_cells(ws, n_rows=50, n_cols=len(headers))
    freeze_and_protect(ws)

    # === QHSE_Risques ===
    ws = wb.create_sheet("QHSE_Risques")
    headers = ["Famille de risque", "Nombre occurrences", "Gravité", "Action immédiate"]
    style_header_row(ws, headers, {1, 2})
    apply_column_widths(ws, headers)
    add_dv_list(ws, "A", col_ref("G", 16))
    add_dv_number(ws, "B", 0)
    add_dv_list(ws, "C", col_ref("H", 4))
    unlock_data_cells(ws, n_rows=50, n_cols=len(headers))
    freeze_and_protect(ws)

    # === QHSE_Actions ===
    ws = wb.create_sheet("QHSE_Actions")
    headers = ["Action à valoriser", "Responsable", "Date prévue", "Statut", "Priorité"]
    style_header_row(ws, headers, {1, 2, 3, 4, 5})
    apply_column_widths(ws, headers)
    add_dv_date(ws, "C"); set_col_format(ws, "C", "yyyy-mm-dd")
    add_dv_list(ws, "D", col_ref("I", 4))
    add_dv_list(ws, "E", col_ref("J", 4))
    unlock_data_cells(ws, n_rows=50, n_cols=len(headers))
    freeze_and_protect(ws)

    add_mode_emploi(
        wb,
        "QHSE — Mode d'emploi",
        [
            "Le QHSE Manager remplit en début de semaine l'Entête puis tous les sous-onglets.",
            "Garder les libellés EXACTS dans les colonnes 'Catégorie' (listes déroulantes)",
            "— ils servent aux KPI mensuels.",
            "Le 'Nombre jours sans accident (cumul)' est obligatoire (saisir 0 si reset).",
            "",
            "Sous-onglets : Entete · QHSE_Meteo · QHSE_Personnel · QHSE_Equipements · ",
            "QHSE_Incidents · QHSE_Securite · QHSE_Risques · QHSE_Actions",
            "",
            f"Mot de passe en-têtes : {PASSWORD}",
        ],
    )

    path = OUT_DIR / "07_QHSE_v1.xlsx"
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# 08 ACTIONS — modèles hebdo
# ---------------------------------------------------------------------------
def build_actions() -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    headers = ["N°", "Section", "Action", "Statut", "Responsable",
               "Date prévue", "Date réalisée", "Priorité",
               "Lien Incident QHSE", "Progrès (%)", "Commentaire"]
    required = {1, 2, 3, 4, 5, 6, 8}

    SECTIONS = [
        "I. Actions menées Semaine S-1",
        "II. Actions à mener Semaine en cours",
        "III. Difficultés / Dysfonctionnements / Incidents",
        "IV. Apprentissages / Bonnes pratiques",
        "V. Points d'attention / Risques émergents",
    ]

    add_sys_sheet(wb, {
        "Sections": SECTIONS,
        "Statuts": ["Ouverte", "En cours", "Réalisée", "Fermée", "Annulée", "Bloquée"],
        "Priorites": ["Basse", "Moyenne", "Haute", "Critique"],
    })

    for week_label in ["2026-S01", "2026-S02", "2026-S03", "2026-S04"]:
        ws = wb.create_sheet(safe_title(week_label))
        style_header_row(ws, headers, required)
        apply_column_widths(ws, headers)

        add_dv_number(ws, "A", 1)
        add_dv_list(ws, "B", f"=_SYS!$A$3:$A${2+len(SECTIONS)}")
        add_dv_list(ws, "D", "=_SYS!$B$3:$B$8")
        add_dv_date(ws, "F"); set_col_format(ws, "F", "yyyy-mm-dd")
        add_dv_date(ws, "G"); set_col_format(ws, "G", "yyyy-mm-dd")
        add_dv_list(ws, "H", "=_SYS!$C$3:$C$6")
        add_dv_number(ws, "J", 0)

        unlock_data_cells(ws, n_rows=200, n_cols=len(headers))
        freeze_and_protect(ws)

    add_mode_emploi(
        wb,
        "ACTIONS — Mode d'emploi",
        [
            "Une ligne = une action.",
            "Sélectionner la Section (I à V), saisir l'action sous le format `Act N : description`.",
            "Renseigner Statut / Responsable / Date prévue obligatoirement.",
            "À la clôture, remplir Date réalisée et passer Statut à 'Réalisée' ou 'Fermée'.",
            "Lier aux incidents QHSE via la référence QHSE-S<n>-INC-<m> quand applicable.",
            "Un onglet par semaine (YYYY-Sxx).",
            "",
            f"Mot de passe en-têtes : {PASSWORD}",
        ],
    )

    path = OUT_DIR / "08_ACTIONS_v1.xlsx"
    wb.save(path)
    return path


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    builders = [
        ("00_MASTER_REFERENTIELS", build_master),
        ("01_LOYERS", build_loyers),
        ("02_CHARGES", build_charges),
        ("03_PLANNING", build_planning),
        ("04_POINTAGE", build_pointage),
        ("05A_CA_FLASH", build_ca_flash),
        ("05B_CA_PNL", build_ca_pnl),
        ("05C_CA_DETAIL", build_ca_detail),
        ("06_ACCOSTAGE", build_accostage),
        ("07_QHSE", build_qhse),
        ("08_ACTIONS", build_actions),
    ]

    produced = []
    for name, fn in builders:
        try:
            path = fn()
            produced.append(path)
            print(f"  OK  {path.name}")
        except Exception as e:
            import traceback
            print(f"  KO  {name} : {e}")
            traceback.print_exc()

    # Récap
    total_ko = 0
    print("\n=== Récapitulatif ===")
    for p in produced:
        try:
            wb = openpyxl.load_workbook(p, read_only=True)
            n_sheets = len(wb.sheetnames)
            wb.close()
        except Exception as e:
            n_sheets = -1
            print(f"   ! Re-ouverture échouée pour {p.name} : {e}")
        size_ko = p.stat().st_size / 1024
        total_ko += size_ko
        print(f"  - {p.name:40s}  {size_ko:6.1f} Ko  {n_sheets:>2} onglets")

    n = len(produced)
    print(f"\nGenerated {n} files in audit/templates/output/ (total {total_ko:.1f} Ko)")

    # Validation : ré-ouverture 1 fichier et lecture du Mode d'emploi
    if produced:
        sample = produced[0]
        print(f"\n=== Validation : lecture du Mode d'emploi de {sample.name} ===")
        wb = openpyxl.load_workbook(sample)
        sheet_name = "📖 Mode d'emploi"
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for r in range(1, min(ws.max_row, 15) + 1):
                v = ws.cell(row=r, column=1).value
                if v:
                    print(f"  L{r}: {v}")
        wb.close()


if __name__ == "__main__":
    main()
