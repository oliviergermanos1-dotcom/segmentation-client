# Audit qualité — 03_PLANNING.xlsx
**Date** : 2026-06-03

## Vue d'ensemble
- Taille : 23657.7 Ko
- Onglets : 51
- Cellules de données estimées : 1 114
- Formules : 24
- Note qualité globale : 🟠 (62/100)

## 🔴 Problèmes critiques (bloquants pour parsing)
- **80 plages fusionnées dans la zone de données** (hors ligne 1) — ex : `30042026-JOUR!E36:E38`, `30042026-JOUR!B31:C31`, `30042026-JOUR!B27:C27`

## 🟠 Problèmes moyens (qualité dégradée)
- **30 cellules avec espaces parasites** (début/fin) — `30042026-JOUR!B2`='                 Planning du 30/04/2026 -JOUR ', `30042026-JOUR!E5`='Activités en magasin et sur yard ', `30042026-JOUR!E6`='Chargement /transfert de 02 ISOTANKS pour MOVIS '
- **18 ruptures par lignes vides** intercalées entre lignes de données (rend la détection de plage fragile)

## 🟢 Points d'amélioration (confort utilisateur)
- **Aucune validation Excel** (listes déroulantes Client/Statut/Mode Paiement absentes) : risque de fautes de frappe et nouvelles valeurs non normées.
- **Aucune cellule figée** : navigation fastidieuse dans les longs tableaux.
- Pas d'onglet « Mode d'emploi / Légende » documentant les colonnes, conventions et statuts.
- Pas de feuille « Référentiels » pour les listes (clients, statuts, types) — chaque saisie ressaisit le texte.

## 📊 Doublons et redondances
- Présence d'onglets historiques/archives multiples (51) — risque de redondance et de versions concurrentes.
- Pas de doublon d'onglet flagrant détecté (au niveau du nom normalisé).

## 🧱 Structure et nommage
- Convention de nommage : cohérente
- Échantillon d'onglets : '30042026-JOUR', '01052026-JOUR', '01052026-NUIT', '02052026-JOUR WK', '02052026-NUIT', '03052026-NUIT', '04052026-JOUR', '04052026-NUIT' …

## 🎯 Recommandations prioritaires (top 5)
1. Démerger les 80 plages fusionnées de la zone de données (1 valeur = 1 cellule).
2. Ajouter des validations Excel (listes déroulantes) sur Client/Statut/Type/Mode pour éviter les variantes.
3. Figer la ligne d'en-tête (View > Freeze Panes) sur chaque feuille de saisie.
4. Supprimer les lignes vides intercalées qui faussent la détection des tableaux.
5. Ajouter un onglet « Mode d'emploi » documentant colonnes, statuts et exceptions.

## Annexe : exemples concrets
- `30042026-JOUR!E36:E38` fusionnée → démerger et propager la valeur.
- `30042026-JOUR!B31:C31` fusionnée → démerger et propager la valeur.
- `30042026-JOUR!B27:C27` fusionnée → démerger et propager la valeur.
- `30042026-JOUR!B2` = '                 Planning du 30/04/2026 -JOUR ' → TRIM().
- `30042026-JOUR!E5` = 'Activités en magasin et sur yard ' → TRIM().
