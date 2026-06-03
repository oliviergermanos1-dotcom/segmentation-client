# Audit qualité — 02_CHARGES.xlsx
**Date** : 2026-06-03

## Vue d'ensemble
- Taille : 36.8 Ko
- Onglets : 5
- Cellules de données estimées : 1 652
- Formules : 12
- Note qualité globale : 🟢 (87/100)

## 🔴 Problèmes critiques (bloquants pour parsing)
- Aucun problème critique détecté.

## 🟠 Problèmes moyens (qualité dégradée)
- **30 cellules avec espaces parasites** (début/fin) — `01-26!A30`='Location véhicule électrique ', `01-26!B30`='BORR DRILLING ', `01-26!A31`='Location Elevateur 16T '

## 🟢 Points d'amélioration (confort utilisateur)
- **Aucune validation Excel** (listes déroulantes Client/Statut/Mode Paiement absentes) : risque de fautes de frappe et nouvelles valeurs non normées.
- **Aucune cellule figée** : navigation fastidieuse dans les longs tableaux.
- Pas d'onglet « Mode d'emploi / Légende » documentant les colonnes, conventions et statuts.
- Pas de feuille « Référentiels » pour les listes (clients, statuts, types) — chaque saisie ressaisit le texte.

## 📊 Doublons et redondances
- Pas de doublon d'onglet flagrant détecté (au niveau du nom normalisé).

## 🧱 Structure et nommage
- Convention de nommage : cohérente
- Échantillon d'onglets : '01-26', '02-26', '03-26', '04-26', '05-26'

## 🎯 Recommandations prioritaires (top 5)
1. Ajouter des validations Excel (listes déroulantes) sur Client/Statut/Type/Mode pour éviter les variantes.
2. Figer la ligne d'en-tête (View > Freeze Panes) sur chaque feuille de saisie.
3. Ajouter un onglet « Mode d'emploi » documentant colonnes, statuts et exceptions.

## Annexe : exemples concrets
- `01-26!A30` = 'Location véhicule électrique ' → TRIM().
- `01-26!B30` = 'BORR DRILLING ' → TRIM().
