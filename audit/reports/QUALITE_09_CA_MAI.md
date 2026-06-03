# Audit qualité — 09_CA_MAI.xlsx
**Date** : 2026-06-03

## Vue d'ensemble
- Taille : 364.6 Ko
- Onglets : 21
- Cellules de données estimées : 9 303
- Note qualité globale : 🟠 (60/100)

## 🔴 Problèmes critiques (bloquants pour parsing)
- **30+ cellules en erreur** (#REF!/#DIV!/#N/A…) — ex : `Prévisions Mois N!T4`=#DIV/0!, `Prévisions Mois N!W4`=#DIV/0!, `Prévisions Mois N!T5`=#DIV/0!, `Prévisions Mois N!W5`=#DIV/0!, `Prévisions Mois N!T8`=#DIV/0!
- **2 onglet(s) archive** polluant la détection auto : PROGRAMME 2014, SOA Envoyés

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
- Convention de nommage onglets : MOIS YY
- Échantillon d'onglets : 'PROGRAMME 2014', 'SOA Envoyés', 'Prévisions Mois N', 'Occupation (2)', 'Représentations'

## 🎯 Recommandations prioritaires (top 5)
1. Résoudre les références #REF! (formules cassées vers fichiers externes ou cellules supprimées).
2. Archiver les onglets historiques dans un fichier séparé (allège la détection).
3. Ajouter des validations Excel (listes déroulantes) sur Client/Statut/Type/Mode.
4. Figer la ligne d'en-tête sur chaque feuille de saisie.
5. Ajouter un onglet « Mode d'emploi » documentant colonnes et conventions.

## Annexe : exemples concrets
- `Prévisions Mois N!T4` = `#DIV/0!` → résoudre la formule ou supprimer la cellule.
- `Prévisions Mois N!W4` = `#DIV/0!` → résoudre la formule ou supprimer la cellule.
- `Prévisions Mois N!T5` = `#DIV/0!` → résoudre la formule ou supprimer la cellule.