#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AGL BUDGET — Dico marchandise → secteur (217 user + 66 ajouts STATCOM)
=====================================================================
Construit SEGEMENTATION_ENRICHI.xlsx : le mapping marchandise→secteur de
l'utilisateur (217 entrées, feuille 'tim') complété par 66 marchandises
présentes dans STATCOM mais absentes du dico initial (OEUFS, BILLETS DE
BANQUE, LINGOT D'OR, etc.), classées au jugement métier.

Usage : python mapping_marchandises.py SEG.xlsx SEGEMENTATION_ENRICHI.xlsx
"""
import sys
import pandas as pd

# 66 marchandises STATCOM non couvertes par le dico user (217) — secteurs
# attribués au jugement métier (nomenclature du dico user, à valider).
COMPLEMENT = {
    'OEUFS': 'AGRI BUSINESS', 'CARTE BANCAIRE': 'AUTRES',
    'COLIS DIPLOMATIQUE': 'AUTRES', 'DOCUMENT ET COLIS EXPRESS': 'AUTRES',
    'COLIS EXPRESS': 'AUTRES', 'DOCUMENTS': 'AUTRES', 'DEPOUILLE MORTELLE': 'AUTRES',
    'VERRES OPTIQUES': 'HEALTHCARE', 'BIJOUX': 'FMCG / RETAIL',
    'BILLETS DE BANQUE': 'AUTRES', 'ANIMAL': 'AGRI BUSINESS',
    "LINGOT D'OR": 'MINING', 'PIECES DETAILLEES MACHINES A OUTIL': 'INFRASTRUCTURE',
    'MANGUE': 'AGRI BUSINESS', 'ONDULEUR': 'HIGHT TECH / TELECOM',
    'POUSSINS VIVANTS': 'AGRI BUSINESS', 'PIECES DETACHEES NEUVES': 'AUTOMOTIVE',
    'NOIX DE COLA': 'AGRI BUSINESS', 'HUILES ESSENTIELLES': 'FMCG / RETAIL',
    'IMPRIMANTE': 'HIGHT TECH / TELECOM', 'SYSTEME DE SECURITE': 'INFRASTRUCTURE',
    'INTERRUPTEURS': 'INFRASTRUCTURE', 'PRODUITS VETERINAIRES': 'HEALTHCARE',
    'TELEPHONE (MOBILE PHONE)': 'HIGHT TECH / TELECOM', 'CAMERA': 'HIGHT TECH / TELECOM',
    'VALVE': 'INFRASTRUCTURE', 'DATTES (FRUIT)': 'AGRI BUSINESS',
    'HORLOGERIE ET PIECES': 'FMCG / RETAIL', 'ANANAS': 'AGRI BUSINESS',
    'ANTENNE': 'HIGHT TECH / TELECOM', 'DISJONCTEURS': 'INFRASTRUCTURE',
    'GANTS POUR INDUSTRIE': 'INFRASTRUCTURE', 'MATERIEL DE PHOTO': 'HIGHT TECH / TELECOM',
    'SOLVICOL (ADHESIFS INDUSTRIELS)': 'INFRASTRUCTURE', 'CHOCOLAT': 'FMCG / RETAIL',
    'CARTE SIM': 'HIGHT TECH / TELECOM', 'PIGMENT': 'PRODUITS CHIMIQUES',
    'MOULES': 'INFRASTRUCTURE', 'NOTE BOOK (CARNET)': 'FMCG / RETAIL',
    'SCIE ET TRONCONNEUSE': 'INFRASTRUCTURE', 'NOIX DE CAJOU': 'AGRI BUSINESS',
    'POLYESTER': 'PRODUITS CHIMIQUES', 'REVETEMENTS DE SOL': 'INFRASTRUCTURE',
    'DETENDEUR': 'INFRASTRUCTURE', 'DERIVES DE CACAO': 'AGRI BUSINESS',
    'ENVELOPPES': 'FMCG / RETAIL', 'COFFRE FORT': 'INFRASTRUCTURE',
    'CAFE': 'AGRI BUSINESS', 'CACAO': 'AGRI BUSINESS', 'ANSES (POIGNEE)': 'FMCG / RETAIL',
    'HUILE VEGETALE': 'FMCG / RETAIL', 'HAUT PARLEUR BAFFLE': 'HIGHT TECH / TELECOM',
    'GLACIAL ACETIC ACID': 'PRODUITS CHIMIQUES', 'COMPRESSEURS': 'INFRASTRUCTURE',
    'COTON': 'AGRI BUSINESS', 'CASQUE DE SECURITE': 'INFRASTRUCTURE',
    'CHLORURE DE SODIUM': 'PRODUITS CHIMIQUES', 'CITRIC ACID MONOHYDRATE': 'PRODUITS CHIMIQUES',
    'CAFÉ SOLUBLE': 'FMCG / RETAIL', 'MONOSODIUM GLUTAMATE': 'PRODUITS CHIMIQUES',
    'MANNEQUIN PLASTIQUE': 'FMCG / RETAIL', 'POLYMERE ACRYLIC': 'PRODUITS CHIMIQUES',
    'SAC A MAIN': 'FMCG / RETAIL', 'PULVERISATEUR AGRICOLE': 'AGRI BUSINESS',
    'SOLVENT NAPHTHA': 'OIL & GAZ', 'SHOE SOLE ( SEMELLE DE CHAUSSURE )': 'FMCG / RETAIL',
}


def build(seg_path, out_path):
    user = pd.read_excel(seg_path, sheet_name='tim', usecols=['SECTEUR', 'MARCHANDISES'])
    user['MARCHANDISES'] = user['MARCHANDISES'].astype(str).str.upper().str.strip()
    comp = pd.DataFrame([(s, m) for m, s in COMPLEMENT.items()],
                        columns=['SECTEUR', 'MARCHANDISES'])
    full = pd.concat([user, comp], ignore_index=True).drop_duplicates('MARCHANDISES')
    with pd.ExcelWriter(out_path, engine='openpyxl') as xl:
        full.to_excel(xl, sheet_name='tim', index=False)
    print(f"User: {len(user)} · ajouts: {len(comp)} · total: {len(full)} → {out_path}")
    return full


if __name__ == '__main__':
    build(sys.argv[1], sys.argv[2])
