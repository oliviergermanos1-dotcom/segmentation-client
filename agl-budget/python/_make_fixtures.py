#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Génère des fixtures réalistes pour tester le pipeline Sprint 2."""
from pathlib import Path
import pandas as pd

OUT = Path(__file__).parent / "fixtures"
OUT.mkdir(exist_ok=True)

# 8 clients "vérité terrain" avec variantes typiques (formes juridiques OHADA, géo, typos).
CRM = pd.DataFrame([
    {"crm_id_compte": "C001", "nom_du_compte": "NESTLE CI",                       "secteur": "Agroalimentaire"},
    {"crm_id_compte": "C002", "nom_du_compte": "SARL CARGIL COTE D'IVOIRE",       "secteur": "Agroalimentaire"},
    {"crm_id_compte": "C003", "nom_du_compte": "Société Ivoirienne de Raffinage", "secteur": "Energie"},
    {"crm_id_compte": "C004", "nom_du_compte": "Ets KOUASSI & Frères",            "secteur": "BTP"},
    {"crm_id_compte": "C005", "nom_du_compte": "Cie Fruitière de l'Ouest",        "secteur": "Agroalimentaire"},
    {"crm_id_compte": "C006", "nom_du_compte": "SIFCA SA",                        "secteur": "Agroalimentaire"},
    {"crm_id_compte": "C007", "nom_du_compte": "TOTAL ENERGIES CI",               "secteur": ""},  # pas de secteur CRM
    {"crm_id_compte": "C008", "nom_du_compte": "ORANGE COTE D IVOIRE",            "secteur": "Telecom"},
])

# IRIS : 7 clients (1 manquant, des variantes orthographiques).
IRIS = pd.DataFrame([
    {"id": "I100", "client": "NESTLÉ COTE D IVOIRE",          "cap": "1200000000"},
    {"id": "I101", "client": "CARGILL CI",                    "cap": "850000000"},
    {"id": "I102", "client": "SIR Société Ivoirienne Raff.",  "cap": "3200000000"},
    {"id": "I103", "client": "KOUASSI FRERES ETS",            "cap": "120000000"},
    {"id": "I104", "client": "COMPAGNIE FRUITIERE OUEST",     "cap": "560000000"},
    {"id": "I105", "client": "SIFCA",                         "cap": "2900000000"},
    {"id": "I106", "client": "ORANGE CI",                     "cap": "780000000"},
])

# STATCOM : 6 clients × variantes, plusieurs lignes par client, métiers variés.
STATCOM = pd.DataFrame([
    # Imports → colonne destinataires renseignée
    {"metier": "Import Maritime", "annee": 2024, "destinataires": "NESTLE COTE D'IVOIRE",     "chargeurs": "",                       "marchandise": "Cacao", "volume_teu": "120"},
    {"metier": "Import Maritime", "annee": 2024, "destinataires": "CARGILL COTE IVOIRE",      "chargeurs": "",                       "marchandise": "Soja",  "volume_teu": "180"},
    {"metier": "Import Aérien",   "annee": 2024, "destinataires": "ORANGE COTE D'IVOIRE SA",  "chargeurs": "",                       "marchandise": "Equipement réseau", "volume_kg": "4200"},
    {"metier": "Hinterland Import","annee": 2025,"destinataires": "ETS KOUASSI FRERES",       "chargeurs": "",                       "marchandise": "Matériaux BTP",     "volume_teu": "45"},
    # Exports → colonne chargeurs renseignée
    {"metier": "Export Maritime", "annee": 2024, "destinataires": "",                         "chargeurs": "SIFCA SA",               "marchandise": "Cacao", "volume_teu": "950"},
    {"metier": "Export Maritime", "annee": 2025, "destinataires": "",                         "chargeurs": "CIE FRUITIERE OUEST CI", "marchandise": "Banane","volume_teu": "780"},
    {"metier": "Export Aérien",   "annee": 2024, "destinataires": "",                         "chargeurs": "SIFCA",                  "marchandise": "Anacarde","volume_kg": "12000"},
    # Client inconnu du CRM (white space attendu)
    {"metier": "Import Maritime", "annee": 2024, "destinataires": "OLAM IVOIRE",              "chargeurs": "",                       "marchandise": "Riz",   "volume_teu": "300"},
])

# RUBRIKS : enrichit le secteur de TOTAL (manquant côté CRM).
RUBRIKS = pd.DataFrame([
    {"client": "TOTAL ENERGIES CI",     "secteur": "Energie",         "cap_pfa": "2400000000"},
    {"client": "ORANGE COTE D'IVOIRE",  "secteur": "Telecom",         "cap_pfa": "900000000"},
    {"client": "NESTLE CI",             "secteur": "Agroalimentaire", "cap_pfa": "1500000000"},
])

CRM.to_excel(OUT / "CRM.xlsx", index=False)
IRIS.to_excel(OUT / "IRIS.xlsx", index=False)
STATCOM.to_excel(OUT / "STATCOM.xlsx", index=False)
RUBRIKS.to_excel(OUT / "RUBRIKS.xlsx", index=False)
print(f"[ok] fixtures écrits dans {OUT}")
