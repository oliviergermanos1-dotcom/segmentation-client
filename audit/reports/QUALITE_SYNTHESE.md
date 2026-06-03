# Synthèse audit qualité — 10 fichiers source
**Date** : 2026-06-03

## Tableau récap qualité par fichier

| Fichier | Taille | Onglets | Note | Top problèmes |
|---|---|---|---|---|
| 01_LOYERS.xlsx | 24 Ko | 2 | 🟢 87/100 | Espaces parasites en fin de libellé · pas de validations |
| 02_CHARGES.xlsx | 38 Ko | 5 | 🟢 87/100 | Colonnes Statut/Mode Paiement vides · pas de validations |
| 03_PLANNING.xlsx | 23 Mo | **51** | 🟠 62/100 | Trop d'onglets (1 par jour/shift) · images embarquées (23 Mo) |
| 04_POINTAGE.xlsx | 520 Ko | **55** | 🔴 22/100 | VLOOKUP cassés vers feuille SOURCE externe · 55 onglets |
| 05_CA.xlsx | 362 Ko | 21 | 🟠 60/100 | **30+ #DIV/0!/#REF!** · 2 onglets archives (PROGRAMME 2014, SOA Envoyés) |
| 06_ACCOSTAGE.xlsx | 70 Ko | 4 | 🟢 100/100 | Mix format heure '11H45' string vs Date · QUAI 25/21 sur 2 colonnes |
| 07_QHSE.xlsx | 10 Mo | 4 | 🟢 100/100 | 545 lignes × 25 cols par semaine, structure dense mais propre |
| 08_ACTIONS.xlsx | 21 Ko | 3 | 🟢 100/100 | 2 colonnes seulement, métadonnées absentes |
| 09_CA_MAI.xlsx | 364 Ko | 21 | 🟠 60/100 | **Idem 05_CA** (clone) — mêmes problèmes hérités |
| 10_STATS_BASE.xlsx | 845 Ko | 6 | 🟠 70/100 | **30+ #REF! sur la feuille Conso AGL CI** (données externes cassées) |

## Les 3 fichiers les plus problématiques

### 🔴 1. POINTAGE — Note 22/100
- **55 onglets** (un par shift jour/nuit/weekend) → maintenance impossible
- Toutes les colonnes Nom/Fonction/Code/Chantier sont des **`=VLOOKUP([2]SOURCE!…)`** vers une **feuille externe non distribuée** → l'app ne récupère que les matricules en clair
- **Action prioritaire** : intégrer la feuille SOURCE dans le fichier OU remplacer les VLOOKUP par des colonnes saisies (le template `04_POINTAGE_v1` règle ça avec un onglet SOURCE_Personnel)

### 🟠 2-3. CA / CA_MAI — Note 60/100 (clones)
- **30+ #DIV/0!** sur l'onglet `Prévisions Mois N` (cols T et W) → formules de variance/% divisant par zéro quand le budget est vide
- **2 onglets archives** (`PROGRAMME 2014`, `SOA Envoyés`) → polluent la détection auto et alourdissent le fichier
- Onglets `Représentations`, `Evolution clients`, `Occupation (2)` orphelins (parsés en GENERIQUE)
- **Action prioritaire** : archiver séparément les onglets historiques + corriger les divisions par zéro avec `=IFERROR(…)`

## Les 10 problèmes transverses les plus fréquents

| # | Problème | Fréquence | Impact |
|---|---|---|---|
| 1 | Absence de validations Excel (listes déroulantes) | **10/10 fichiers** | Variantes de saisie (FOXTROT vs Foxtrot vs FOX) |
| 2 | Absence de cellules figées | **10/10 fichiers** | Navigation pénible sur les longues tables |
| 3 | Pas d'onglet « Mode d'emploi » | **10/10 fichiers** | Nouveaux opérateurs non guidés |
| 4 | Pas de feuille « Référentiels » centralisée | **10/10 fichiers** | Listes ressaisies dans chaque fichier |
| 5 | Cellules en erreur (#REF!/#DIV/0!/#N/A) | 3/10 (CA, CA_MAI, STATS) | Parser doit filtrer ces valeurs |
| 6 | Onglets archives mélangés aux données actives | 2/10 (CA, CA_MAI) | Détection auto polluée |
| 7 | VLOOKUP vers fichiers externes | 1/10 (POINTAGE) | Données invisibles hors Excel local |
| 8 | Multiplication d'onglets « 1 par jour/shift » | 2/10 (PLANNING 51, POINTAGE 55) | Maintenance et archivage difficiles |
| 9 | Mix format heures (string '11H45' vs Date) | 1/10 (ACCOSTAGE) | Parser doit gérer 2 formats |
| 10 | Espaces parasites en fin de libellé | 1/10 (LOYERS) | Matching cassé (`'Yard 600 '` ≠ `'Yard 600'`) |

## Roadmap qualité priorisée

### Priorité 1 — Quick wins (1-2 jours, gros impact)
1. **Corriger les VLOOKUP de POINTAGE** : intégrer la feuille `SOURCE` dans le fichier (déjà fait dans le template `04_POINTAGE_v1.xlsx`)
2. **Wrapper les divisions sensibles en `=IFERROR(…)` dans CA/CA_MAI** : élimine 30+ #DIV/0!
3. **Archiver les onglets historiques de CA** : déplacer `PROGRAMME 2014`, `SOA Envoyés`, `Évolution clients` dans un fichier séparé `CA_archives.xlsx`

### Priorité 2 — Standardisation (1 semaine)
4. **Déployer les templates `00_MASTER_REFERENTIELS.xlsx` + 11 templates stricts** déjà pré-remplis (cf. ZIP `AGL_Templates_PreRemplis`)
5. **Former les opérateurs** au mode d'emploi en première page de chaque template
6. **Réduire le nombre d'onglets PLANNING/POINTAGE** : un fichier mensuel (30 onglets max) au lieu d'annuel

### Priorité 3 — Industrialisation (1 mois)
7. **Imposer le format des templates** : refus des fichiers sources qui n'ont pas la même structure
8. **Hub Référentiels** : 1 fichier `MASTER_REFERENTIELS` partagé, mis à jour par le KAM, consommé par tous les autres fichiers
9. **Versioning sur chaque fichier** : cellule `_SYS!A1` avec version + date de génération

## Note méthodologique
- Audit basé sur les 200 premières lignes × 12 premiers onglets par fichier (sample représentatif pour les gros fichiers)
- Score qualité = 100 - pénalités (erreurs/2pts, archives/5pts, accents/5pts, vides/3pts)
- Rapports détaillés : `audit/reports/QUALITE_NN_TYPE.md` (10 fichiers)
