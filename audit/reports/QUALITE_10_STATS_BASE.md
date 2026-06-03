# Audit qualité — 10_STATS_BASE.xlsx
**Date** : 2026-06-03

## Vue d'ensemble
- Taille : 845.8 Ko
- Onglets : 6
- Cellules de données estimées : 16 899
- Note qualité globale : 🟠 (70/100)

## 🔴 Problèmes critiques (bloquants pour parsing)
- **30+ cellules en erreur** (#REF!/#DIV!/#N/A…) — ex : `Conso AGL CI ABJ & SPY!B7`=#REF!, `Conso AGL CI ABJ & SPY!C7`=#REF!, `Conso AGL CI ABJ & SPY!D7`=#REF!, `Conso AGL CI ABJ & SPY!E7`=#REF!, `Conso AGL CI ABJ & SPY!F7`=#REF!

## 🟠 Problèmes moyens (qualité dégradée)
- Aucun problème moyen détecté.

## 🟢 Points d'amélioration (confort utilisateur)
- **Aucune validation Excel** : risque de fautes de frappe et nouvelles valeurs non normées.
- **Aucune cellule figée** sur les feuilles de saisie.
- Pas d'onglet « Mode d'emploi / Légende ».
- Pas de feuille « Référentiels » centralisant les listes.

## 📊 Doublons et redondances
- Pas de doublon d'onglet flagrant détecté.

## 🧱 Structure et nommage
- Convention de nommage onglets : hétérogène ou plat
- Échantillon d'onglets : 'Conso AGL CI ABJ & SPY', 'BASE LOGISTIQUE', 'Détail Stats Exp.', 'Complément Stats', 'Contrôle'

## 🎯 Recommandations prioritaires (top 5)
1. Résoudre les références #REF! (formules cassées vers fichiers externes ou cellules supprimées).
2. Ajouter des validations Excel (listes déroulantes) sur Client/Statut/Type/Mode.
3. Figer la ligne d'en-tête sur chaque feuille de saisie.
4. Ajouter un onglet « Mode d'emploi » documentant colonnes et conventions.

## Annexe : exemples concrets
- `Conso AGL CI ABJ & SPY!B7` = `#REF!` → résoudre la formule ou supprimer la cellule.
- `Conso AGL CI ABJ & SPY!C7` = `#REF!` → résoudre la formule ou supprimer la cellule.
- `Conso AGL CI ABJ & SPY!D7` = `#REF!` → résoudre la formule ou supprimer la cellule.