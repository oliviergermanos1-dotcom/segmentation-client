# Audit qualité — 04_POINTAGE.xlsx
**Date** : 2026-06-03

## Vue d'ensemble
- Taille : 508.0 Ko
- Onglets : 55
- Cellules de données estimées : 12 900
- Formules : 8 515
- Note qualité globale : 🔴 (22/100)

## 🔴 Problèmes critiques (bloquants pour parsing)
- **85 cellules d'erreur** (#REF!, #DIV/0!, #N/A) — ex : `06052026-JOUR!E75`, `06052026-JOUR!F75`, `06052026-JOUR!G75`
- **50 formules avec références externes** (cassent à l'export) — ex : `01052026-JOUR !M7`, `01052026-JOUR !D8`, `01052026-JOUR !E8`
- **55 plages fusionnées dans la zone de données** (hors ligne 1) — ex : `01052026-JOUR !B3:G3`, `01052026-NUIT!B3:G3`, `02052026-JOUR!B3:G3`

## 🟠 Problèmes moyens (qualité dégradée)
- **30 cellules avec espaces parasites** (début/fin) — `01052026-JOUR !B5`='Shift : jour ', `01052026-NUIT!B5`='Shift : NUIT ', `02052026-JOUR!B5`='Shift : jour '
- **50 ruptures par lignes vides** intercalées entre lignes de données (rend la détection de plage fragile)

## 🟢 Points d'amélioration (confort utilisateur)
- Validations présentes sur 55 onglet(s) — étendre aux autres.
- **Aucune cellule figée** : navigation fastidieuse dans les longs tableaux.
- Pas d'onglet « Mode d'emploi / Légende » documentant les colonnes, conventions et statuts.
- Pas de feuille « Référentiels » pour les listes (clients, statuts, types) — chaque saisie ressaisit le texte.

## 📊 Doublons et redondances
- Présence d'onglets historiques/archives multiples (55) — risque de redondance et de versions concurrentes.
- Pas de doublon d'onglet flagrant détecté (au niveau du nom normalisé).

## 🧱 Structure et nommage
- Convention de nommage : cohérente
- Échantillon d'onglets : '01052026-JOUR ', '01052026-NUIT', '02052026-JOUR', '02052026-NUIT', '03052026-JOUR WK', '03052026-NUIT', '04052026-JOUR', '04052026-NUIT' …

## 🎯 Recommandations prioritaires (top 5)
1. Corriger les 85 cellules d'erreur (#REF!/#DIV/0!/#N/A) avant tout parsing automatisé.
2. Remplacer les 50 formules à références externes par des valeurs ou liens internes.
3. Démerger les 55 plages fusionnées de la zone de données (1 valeur = 1 cellule).
4. Figer la ligne d'en-tête (View > Freeze Panes) sur chaque feuille de saisie.
5. Supprimer les lignes vides intercalées qui faussent la détection des tableaux.

## Annexe : exemples concrets
- `06052026-JOUR!E75` = `#N/A` → recalculer la formule ou remplacer par une valeur figée.
- `06052026-JOUR!F75` = `#N/A` → recalculer la formule ou remplacer par une valeur figée.
- `06052026-JOUR!G75` = `#N/A` → recalculer la formule ou remplacer par une valeur figée.
- `01052026-JOUR !M7` = `=VLOOKUP(L7,[1]SOURCE!R:S,2,0)` → références externe : casser le lien.
- `01052026-JOUR !D8` = `=VLOOKUP(C8,[2]SOURCE!$B$6:$G$105,2,0)` → références externe : casser le lien.
- `01052026-JOUR !B3:G3` fusionnée → démerger et propager la valeur.
- `01052026-NUIT!B3:G3` fusionnée → démerger et propager la valeur.
- `02052026-JOUR!B3:G3` fusionnée → démerger et propager la valeur.
