# Audit qualité — 01_LOYERS.xlsx
**Date** : 2026-06-03

## Vue d'ensemble
- Taille : 24.2 Ko
- Onglets : 2
- Cellules de données estimées : 882
- Formules : 220
- Note qualité globale : 🟢 (87/100)

## 🔴 Problèmes critiques (bloquants pour parsing)
- Aucun problème critique détecté.

## 🟠 Problèmes moyens (qualité dégradée)
- **19 cellules avec espaces parasites** (début/fin) — `Jan!B16`='Magasins SICC (4000 m²)  ', `Jan!B18`='Yard 600 ', `Jan!B19`='bureau 228 m² '

## 🟢 Points d'amélioration (confort utilisateur)
- **Aucune validation Excel** (listes déroulantes Client/Statut/Mode Paiement absentes) : risque de fautes de frappe et nouvelles valeurs non normées.
- **Aucune cellule figée** : navigation fastidieuse dans les longs tableaux.
- Pas d'onglet « Mode d'emploi / Légende » documentant les colonnes, conventions et statuts.
- Pas de feuille « Référentiels » pour les listes (clients, statuts, types) — chaque saisie ressaisit le texte.

## 📊 Doublons et redondances
- Pas de doublon d'onglet flagrant détecté (au niveau du nom normalisé).

## 🧱 Structure et nommage
- Convention de nommage : cohérente
- Échantillon d'onglets : 'Jan', 'Feuil1'

## 🎯 Recommandations prioritaires (top 5)
1. Ajouter des validations Excel (listes déroulantes) sur Client/Statut/Type/Mode pour éviter les variantes.
2. Figer la ligne d'en-tête (View > Freeze Panes) sur chaque feuille de saisie.
3. Ajouter un onglet « Mode d'emploi » documentant colonnes, statuts et exceptions.

## Annexe : exemples concrets
- `Jan!B16` = 'Magasins SICC (4000 m²)  ' → TRIM().
- `Jan!B18` = 'Yard 600 ' → TRIM().
