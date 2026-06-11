#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Sectorisation par MOTS-CLÉS du nom d'entreprise
============================================================
Signal complémentaire : le nom d'une société révèle souvent son secteur
(ex. "DEVELOP CAOUTCHOUC IVOIRE" → AGRICULTURE, "IVOIRIENNE RAFFINAGE" →
OIL & GAS). Règles à haute précision (mots entiers), nomenclature CRM.

Sert à :
  - sectoriser les white spaces sans signal marchandise ;
  - re-vérifier les certifiés croisés (si le nom contredit nettement le
    secteur voté, on rétrograde le client en A_TRANCHER).

secteur_par_nom(nom) → secteur (str) ou None si aucun mot-clé probant.
particulier_ou_sale(nom) → True si "PARTICULIER" ou nom non exploitable.
"""
import re

# Règles HAUTE CONFIANCE : mots-clés sans ambiguïté sur le secteur.
# Utilisées pour sectoriser ET pour re-vérifier/rétrograder un croisé.
RULES_HIGH = [
    ("OIL & GAS", ["RAFFINAGE", "PETROL", "PETROLE", "HYDROCARBUR",
                   "OILFIELD", "FORAGE PETROL"]),
    ("MINING", ["MINIER", "MINING", " MINES", "MINERAI", "MANGANESE",
                "SPODUMENE", "LITHIUM", "BAUXITE"]),
    ("POWER TECHNOLOGIES", ["SOLAIRE", "PHOTOVOLTA", "CENTRALE ELECTRIQUE"]),
    ("HEALTHCARE", ["PHARMA", "CLINIQUE", "HOPITAL", "HOSPITAL",
                    "MEDICAL", "DENTAIRE", "LABORATOIRE MEDIC"]),
    ("TELECOM", ["TELECOM", "TELEPHONIE", "RESEAU MOBILE"]),
    ("AUTOMOTIVE", ["AUTOMOBILE", "MOTORS", "PNEUMATIQUE", "PIECES AUTO"]),
    ("INFRASTRUCTURES", ["BTP", "CIMENT", "BETON", "TRAVAUX PUBLIC",
                         "TERRASSEMENT", "GENIE CIVIL"]),
    ("CHEMICALS & PETROCHEMICALS", ["CHIMIE", "CHIMIQUE", "CHEMICAL",
                                    "POLYMER", "PEINTURE", "ENGRAIS"]),
    ("GROUP - FREIGHT FORWARDER", ["TRANSIT", "CONSIGNATION", "ACCONAGE",
                                   "FREIGHT FORWARD"]),
    ("AGRICULTURE", ["CAOUTCHOUC", "HEVEA", "CACAO", "ANACARDE", "CASHEW",
                     "KARITE", "PALMISTE", "PLANTATION", "GRUME", "HORTICOLE"]),
]

# Règles BASSE CONFIANCE : indices plausibles mais ambigus. Servent SEULEMENT
# à sectoriser un white space sans autre signal ; ne rétrogradent jamais.
RULES_LOW = [
    ("AGRICULTURE", ["AGRO", "AGRI", "SEMENCE", "BOIS", "COTON", "CAFE",
                     "CAJOU", "PALM", "FRUIT", "BANANE", "MANGUE", "ALIMENT",
                     "CEREALE", "RIZ ", "NOIX", "SESAME", "RUBBER"]),
    ("FAST MOVING CONSUMER GOODS (FMCG)", ["DISTRIBUTION", "COMMERCE", "NEGOCE",
                                           "SUPERMARCHE", "BOUTIQUE", "RETAIL",
                                           "COSMETIQUE", "BRASSERIE"]),
    ("INDUSTRIES & SERVICES", ["BANQUE", "BANK", "ASSURANCE", "FINANCE", "BOURSE",
                               "CONSULTING", "SERVICES", "INDUSTRIE", "INGENIERIE"]),
    ("AUTOMOTIVE", ["GARAGE", " AUTO "]),
    ("POWER TECHNOLOGIES", ["ENERGIE", "ENERGY", "ELECTRIQUE"]),
]

_SALE = re.compile(r"^(PARTICULIER|CLIENT|DIVERS|INCONNU|N/?A|XXX|TEST|"
                   r"COMPTANT|SANS NOM|NON DEFINI|\W*)\W*$", re.I)


def particulier_ou_sale(nom) -> bool:
    """True si nom = particulier ou non exploitable (vide, code, générique)."""
    if nom is None:
        return True
    s = str(nom).strip().upper()
    if not s or len(s.replace(" ", "")) < 3:
        return True
    if _SALE.match(s):
        return True
    if "PARTICULIER" in s:
        return True
    return False


def secteur_par_nom(nom, high_only=False):
    """Secteur déduit du nom, ou None.
    high_only=True → seulement les règles haute confiance (pour rétrograder)."""
    if nom is None:
        return None
    s = f" {str(nom).upper()} "
    for secteur, kws in RULES_HIGH:
        for kw in kws:
            if kw in s:
                return secteur
    if high_only:
        return None
    for secteur, kws in RULES_LOW:
        for kw in kws:
            if kw in s:
                return secteur
    return None
