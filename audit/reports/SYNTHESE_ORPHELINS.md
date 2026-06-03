# SYNTHÈSE AUDIT ORPHELINS — O&G AGL Analyseur v23
**Date : 03 Juin 2026**  
**Objet** : Identification des colonnes Excel NON exploitées par les parsers JavaScript

---

## EXECUTIVE SUMMARY

Audit complet des 8 fichiers Excel de la base logistique O&G, croisé avec l'analyse des 8 parsers JS. **Objectif** : repérer les champs présents dans les sources mais absents des extraits exploités, afin de concevoir des templates Excel "100 % complet" pour les opérateurs.

### Méthodologie
- Analyse des **25 premières lignes × 20 colonnes** (sample dans rapports audit)
- **Extraction des parsers JS** : `parseQHSE`, `parsePointage`, `parsePlanning`, `parseActions`, `parseCharges`, `parseLoyers`, `parseCA`, `parseAccostage`
- Identification du **mapping colonne→clé d'extraction** via index recherche (`findIndex`) et regex (`/pattern/i.test()`)
- Classification : **EXTRAIT** vs **ORPHELIN** vs **ONGLET ORPHELIN**

---

## 01 — LOYERS.xlsx

### Onglets (2)
- **`Jan`** (73 lignes × 15 colonnes)
- **`Feuil1`** (46 lignes × 2 colonnes) — RÉFÉRENTIEL ESPACES DÉDIÉS

### Colonnes PRÉSENTES (Jan)
| Colonne | Libellé | Exemple |
|---------|---------|---------|
| A | Client | FOXTROT |
| B | Libéllé facturation | Protocole Mag T10 4500m² & Yard 3262m² |
| C | Montant (XOF H.T) | 27469000 |
| D-O | Mois 1-12 (2026-01-01 à 2026-12-01) | 27469000 (replicated) |

### Champs EXTRAITS par parseLoyers()
```
- month_labels: ['M1', 'M2', ..., 'M12']
- periods[i]:
  - ym (year-month): '2026-01'
  - byClient: { 'FOXTROT': montant, ... }
  - total: somme mensuels
- globalByClient: { client: total_annuel }
- totalAnnuel: somme tous clients/mois
- nbLignes, nbClients, monthLabels
```

### Champs ORPHELINS
| Colonne | Libellé | Exemple | Risque/Note |
|---------|---------|---------|------------|
| - | Feuil1 entière (Espaces dédiés) | 46 lignes | **ONGLET ORPHELIN** — Référentiel clients/espaces, 0 cellule numérique. Non lu. |
| - | Formules XLS (=C9/2, =L9, etc) | - | Parser extrait *valeurs évaluées*, pas formules. OK. |

**Observations**
- Parser ne lit QUE le **premier onglet** (`wb.SheetNames[0]`)
- **`Feuil1`** (référentiel espaces) **jamais exploité** — contient structure clients/libellés uniques (utile pour dictionnaire)
- Pas de colonnes calculées, critères de sélection, ou hiérarchies détectées dans sample

---

## 02 — CHARGES.xlsx

### Onglets (5)
- `01-26`, `02-26`, `03-26`, `04-26`, `05-26` (41–86 lignes × 5–8 colonnes)

### Colonnes PRÉSENTES (par feuille type 01-26)
| Colonne | Libellé | Exemple |
|---------|---------|---------|
| A | Libellé | Location Tracteur routier 4x2 |
| B | N° UM | 402301 |
| C | Fournisseur | AYTHAM SERVICES |
| D | Affectation | 40B00 / 40B01 |
| E | Montant H.T (XOF) | 3000000 |
| (03-26+) F | Mois | MARS, JANVIER, etc |
| (03-26+) G | Statut | (vide dans sample) |
| (03-26+) H | Mode Paiement | (vide dans sample) |

### Champs EXTRAITS par parseCharges()
```
- periods[i] (par feuille):
  - total: somme montants
  - byAff: { '40B00': montant, '40B01': montant, ... }
  - byFour: { 'AYTHAM SERVICES': montant, ... }
  - lignes: [ { lib, aff, four, montant }, ... ]
```

### Champs ORPHELINS
| Colonne | Libellé | Exemple | Risque/Note |
|---------|---------|---------|------------|
| B | N° UM | 402301 | **ORPHELIN** — Code identifiant équipement/service. Présent mais ignoré (utile pour tracking). |
| F | Mois | MARS, JANVIER | **ORPHELIN (03-26+)** — Colonne présente en Mars/Avril/Mai mais non extraite. Utile pour profiler charges par mois. |
| G | Statut | (vide) | **ORPHELIN (03-26+)** — Colonne présente mais vide dans sample. Pourrait accueillir état justification/validation. |
| H | Mode Paiement | (vide) | **ORPHELIN (03-26+)** — Colonne titre présente mais vide. Critère de segmentation comptable. |

**Observations**
- Parser agrège **par affectation (40B00/40B01) et fournisseur** — structuré pour comptabilité analytique
- **N° UM** manquant = impossible de joindre avec master équipements (traçabilité risque)
- Colonnes **Mois, Statut, Mode Paiement** (ajoutées en Mars) = champs "extended" non lu => **gap de flexibilité**

---

## 03 — PLANNING.xlsx

### Onglets (51)
- 50 feuilles `DD-MM-YYYY-{JOUR|NUIT|FERIE|WK}` (30/04 à 28/05/2026)
- 1 feuille `Feuil1` (ORPHELINE)

### Colonnes PRÉSENTES (type jour)
| Colonne | Libellé | Exemple |
|---------|---------|---------|
| B | CLIENT | BAKER HUGHES |
| C | SECTION | CWI (SHAN), OIC, TRS, OFE, ... |
| D | DEMANDE EMISE | 4t, 01Sr20+01chall, 7t+01ELG, ... |
| E | NATURE DE L'OPÉRATION | Activités en magasin et sur yard, ... |
| F | POSITIF | 1, 0 |
| G | NEGA | 0 |
| H | OBSERVATIONS | AJOUT, Location mensuelle(CONTRACT), ... |

### Champs EXTRAITS par parsePlanning()
```
- periods[i] (par feuille):
  - ops: [ { client, section, nature, pos, neg }, ... ]
  - pos, neg, total: comptage opérations
  - shift: 'JOUR' / 'NUIT'
  - wk: bool (weekend/férié)
  - ferie: bool
```

### Champs ORPHELINS
| Colonne | Libellé | Exemple | Risque/Note |
|---------|---------|---------|------------|
| D | DEMANDE EMISE | 4t, 12t+01Sr40, Matrix1+08spc | **ORPHELIN** — Matériel/ressources. Non extrait. Critique pour capacité logistique. |
| H | OBSERVATIONS | AJOUT, Location mensuelle | **ORPHELIN (partiellement)** — Texte libre. Parser ne capture pas. Utile pour flags urgence/contrats. |
| - | Feuil1 (ORPHELINE) | - | **ONGLET ORPHELIN** — Pas listé, probablement template/notes. |
| - | Nuits (structure) | Réduite à 9-10 col | Nuits = **format différent** : CLIENT, DEMANDE, POSITIF, NEGA, OBSERVATIONS seult. Moins de détails. |

**Observations**
- **DEMANDE EMISE** (col D) = équivalent "charges/ressources" pour planning — **NON EXPLOITÉ**
- Split JOUR/NUIT implique **2 structures différentes** : parser normalise via `isNightSheet()`
- Parser agrège **pos/neg seulement** — matériel/observations textuelle = pure **orpheline**

---

## 04 — POINTAGE.xlsx

### Onglets (55)
- 54 feuilles `DD-MM-YYYY-{JOUR|NUIT|WK}` (01/05 à 28/05/2026)
- Pas de `Feuil1` observée

### Colonnes PRÉSENTES (type jour)
| Colonne | Libellé | Exemple |
|---------|---------|---------|
| C | Matricule | 11580, 13143, 16997, ... |
| D | Nom et Prémons | [VLOOKUP] → =VLOOKUP(C8,[2]SOURCE!...) |
| E | Fonction | [VLOOKUP] → =VLOOKUP(C8,[2]SOURCE!...) |
| F | Code Fonction | [VLOOKUP] |
| G | Chantiers/Activites | [VLOOKUP] |
| H | Produits | [VLOOKUP] |
| (autres colonnes) | Shift, Date | Vendredi 01/05/2026, JOUR / NUIT |

### Champs EXTRAITS par parsePointage()
```
- periods[i] (par feuille):
  - workers: [ { mat, nom, fonction, code, chantier }, ... ]
  - shift: 'JOUR' / 'NUIT'
  - wk: bool
  - date, sortKey
```

### Champs ORPHELINS
| Colonne | Libellé | Exemple | Risque/Note |
|---------|---------|---------|------------|
| H | Produits | [VLOOKUP] | **ORPHELIN** — Affectation produit par worker. Formule VLOOKUP non résolue dans extract → contenu perdu. |
| (N-Q) | Colonnes dynamiques | JOHNSON Kouakouvi Nelson (L7) | **ORPHELIN** — Colonnes L-Q semblent être validations/confirmations (dropdown). Non exploitées. |

**Observations**
- Parser cherche **matricule + fonction** comme clé header
- **Produits (col H)** = VLOOKUP non résolue = **orpheline totale**
- **Source sheet** (`[2]SOURCE!$B$6:$G$105`) référencée mais non accessible dans extraction → dépendance fonctionnelle perdue
- Colonnes au-delà de H = **colonnes de validation** (dropdown list, signatures) → orphelines

---

## 05 — CA.xlsx

### Onglets (21)
- `PROGRAMME 2014`, `SOA Envoyés`, `Prévisions Mois N`, `Occupation (2)`, `Représentations`, `Gestion base logistique`, `Evolution clients`, `Récap Global`, `Cumul annuel`, `Janvier`–`Décembre` (12 feuilles)

### Colonnes PRÉSENTES (exemple Prévisions Mois N)
| Colonne | Libellé | Exemple |
|---------|---------|---------|
| A | Client | Afritramp, Antrak Ghana, Anadarko, ... |
| B | BASE LOYER | (vide ou formule) |
| C | Plog & Divers | (vide ou formule) |
| D | Total BASE | =D4-B4 |
| E-M | CA GENERE (IMAR, HINT, IAER, EAER, AÉRIEN, EXPORT, SIRT, Afritramp, Meet&Greet) | (vide) |
| N | CA GENERE Autres | (vide) |
| O | CA TOTAL GENERE PAR CLIENT | =N4+D4 |
| P | Sources revenus | Client name |
| Q | REEL BASE 2016 | Montant (ex: 10416666) |
| R | Budget 2016 | 0 |
| S | VAR (FCFA) | =Q4-R4 |
| T | Var (%) | =Q4/R4-1 |

### Champs EXTRAITS par parseCA()
```
Vue 1 (Récap Global):
  - monthsCA[i]: { ym, label, byDep, total }
  - deps: liste
  - totalCA

Vue 2 (Cumul annuel):
  - clients: [ { client, total, loyers }, ... ]
  - totalClients, totalLoyersCA

Vue 3 (feuilles mensuelles):
  - budget: [ { ym, reel, budget, reel25 }, ... ]
  - totReel, totBudget, totReel25
```

### Champs ORPHELINS
| Colonne/Onglet | Libellé | Exemple | Risque/Note |
|---------|---------|---------|------------|
| E-M | CA GENERE (IMAR, HINT, ...) | (vide) | **ORPHELIN** — Détail par département/activité. Présent mais jamais rempli dans sample. Structure de segmentation ignorée. |
| B | BASE LOYER | (formule) | **ORPHELIN** — Segmentation loyer vs autres revenus. Pas extraite. |
| P-T | Sources/Budget détail | Sources revenus, REEL BASE, Budget, VAR | **ORPHELIN** — Colonnes de comparaison réel/budget/N-1. Parser lit SEULEMENT vue 3 (budget[]). Vue 1 (Récap) ignore détail départemental. |
| `PROGRAMME 2014` | Matrice historique | (vide) | **ONGLET ORPHELIN** — Programme 2014, cellules fusionnées massives. Probablement archive non entretenue. |
| `SOA Envoyés` | Matrix client × mois | x (marquage) | **ONGLET ORPHELIN** — Matrice clients SOA 2017 (historique). 0 cellule numérique = suivi administratif non opérationnel. |
| `Occupation (2)` | Occupation espaces client | Superficie/Price/Amount | **ONGLET PARTIELLEMENT ORPHELIN** — Parser ignore complètement. Contient calcul CA par client/espace (utile). |
| `Représentations` | Taux rétrocession % | Referencing formulas | **ONGLET ORPHELIN** — Agrégation par département, références '#REF!' detectable. Pas lu. |
| `Evolution clients` | (suite...) | - | **ONGLET ORPHELIN** — Probablement dynamique clients (entrées/sorties). Non exploité. |
| `Gestion base logistique` | Compte d'exploitation | CAP NET, Salaires, Frais | **ONGLET ORPHELIN** — Très détaillé (25 lignes × 40 colonnes), paramétré mois par mois (Prévu/Réel/Variation). Aucune extraction. Contient **KPI financier manquant**. |

**Observations**
- **parseCA() = 3 "vues" disjointes** :
  - Vue 1 (Récap Global) : mois × dép (CA GENERE non exploité)
  - Vue 2 (Cumul) : clients + loyers
  - Vue 3 (Budget) : reel/budget vs N-1
  - **Aucune consolidation** entre vues
- **Nombreux onglets orphelins** (7+) : archive, comparatif, occupation, exploitation = **gisement majeur de données perdues**
- **Cas critique** : `Gestion base logistique` = flux P&L complet, **0 extraction**

---

## 06 — ACCOSTAGE.xlsx

### Onglets (4)
- `JANVIER 26`, `FEVRIER 26`, `MARS 26`, `AVRIL 26` (70–99 lignes × 21–22 colonnes)

### Colonnes PRÉSENTES
| Colonne | Libellé | Exemple |
|---------|---------|---------|
| A | NUMEROS D'ORDRE | 1, 2, 3, ... |
| B | DATES | 2026-01-01, ... |
| C | NAVIRES | BLUE MOMMY, MV KLARA, ... |
| D | CLIENTS | MURPHY, NAVITRANS/BOSKALIS, ... |
| E | ETA | 11H45, 03H30, ... |
| F | ETD | 15H45, 10h30, ... |
| G | NBRE HEURES | 04H 00, 07h00, ... |
| H | QUAI 25 | 0/1 (occupation) |
| I | QUAI 21 | 0/1 (occupation) |
| J | MONTANT | 1500000, 1200000, ... |
| **L** | **GOTTWALD DATE** | 2026-02-06 |
| **M** | **MOB /H** | 0H30 |
| **N** | **DEMOB/H** | 05H30 |
| **O** | **Nbre H** | 5, 13, 6 |
| **P** | **BL** | 260, 261-262, 263 |
| **R** | **MAFI DATE** | 2026-03-05 |
| **S** | **SHIFT JOUR** | 1 |
| **T** | **SHIFT NUIT** | (vide) |
| **U** | **TOTAL** | =Q+R |

### Champs EXTRAITS par parseAccostage()
```
- lignes: [ { navire, client, montant, free, quai, turnH }, ... ]
- byClient: { client: montant, ... }
- byNavire: { navire: montant, ... }
- escales, nbNavires, nbClients, freeCount
- turnaroundAvgH, occupancyPct, ym
```

### Champs ORPHELINS
| Colonne | Libellé | Exemple | Risque/Note |
|---------|---------|---------|------------|
| G | NBRE HEURES | 04H 00, 08h30 | **ORPHELIN** — Durée escale. Extractible mais ignoré (parser calcule via ETA/ETD). Pourrait servir validation. |
| **BLOC GOTTWALD** | **L-P** (DATE, MOB/H, DEMOB/H, Nbre H, BL) | 2026-02-06, 0H30, 05H30, 5, 260 | **ORPHELIN MAJEUR** — Grue GOTTWALD : mobilisation, durée, nombre de mouvements, BLs. **NON EXPLOITÉ**. Critique pour KPI équipement. |
| **BLOC MAFI** | **R-U** (DATE, SHIFT JOUR, SHIFT NUIT, TOTAL) | 2026-03-05, 1, 0, 1 | **ORPHELIN MAJEUR** — Grue MAFI : shifts (jour/nuit), totals. **ZÉRO EXTRACTION**. Opérabilité équipement perdue. |
| H, I | QUAI 25, QUAI 21 | 0/1 | **ORPHELIN (PARTIELLEMENT)** — Occupation quais extraite (`quai`), mais **distinction 25 vs 21 ignorée** (normalisé en `quaiSet`). Perte info poste spécifique. |

**Observations**
- **Parser = ultra-simplifié** : lit SEULEMENT navire/client/montant
- **2 blocs opérateurs majeurs** (GOTTWALD L-P, MAFI R-U) = 10 colonnes = **100 % orphelines**
- **ETA/ETD + NBRE HEURES** = triplet redondant, parser ignore NBRE HEURES (déduit de ETA/ETD) mais colonne existe
- **Distinction QUAI 25 vs QUAI 21** = **PERDUE** (normalisée en ensemble `quaiSet`, pas segmentation)
- **Cas limité** : seul accostage quai exploité, aucune KPI équipement grue

---

## 07 — QHSE.xlsx

### Onglets (4)
- `SEMAINE DU 27AV AU 03 MAI202(2)`
- `SEMAINE DU 04 AU 10 MAI2026 (2)`
- `SEMAINE DU 11 AU 17 MAI2026 (2)`
- `SEMAINE DU 18 AU 24 MAI2026`
(545 lignes × 25 colonnes, très fusionnées)

### Colonnes PRÉSENTES (sample 25 premières lignes)
| Colonne | Libellé | Exemple |
|---------|---------|---------|
| B | PAYS | Côte d'ivoire |
| D | ENTITÉ | AFRICA GLOBAL LOGISTICS OIL&GAS |
| (header ligne 5) | SEMAINE | 18, 19, 20, 21 |
| (PERSONNEL) | EFFECTIFS | AGL, Sous-traitant, Journaliers |
| D-F | CADRES, NON CADRES, TOTAL | 5, 15, =D16+E16 |
| (ÉQUIPEMENTS) | Status | Certifié, Etalonné, En cours, A réaliser |
| (OPÉRATIONS) | (divers) | (non visible dans sample L25) |
| (texte libre) | Observations sécurité | (fusion massive) |

### Champs EXTRAITS par parseQHSE()
```
- semaine: semaine number
- effectifs: { 'AGL': X, 'Sous-traitant': Y, 'Journaliers': Z }
- totalEff: sum
- kpi: { 'C:AGL': value, 'C:Sous-traitant': value, ... ; 'J:label': value, ... }
- jsa: nombre de jours sans accident (recherche par label)
```

### Champs ORPHELINS
| Colonne/Section | Libellé | Exemple | Risque/Note |
|---------|---------|---------|------------|
| (MÉTÉO) | Ensoleillé, Nuageux, Venteux, Pluie | OUI/NON | **ORPHELIN** — Conditions météo. Colonne logique (OUI/NON). Ignorée. Utile pour risk assessment. |
| (ÉQUIPEMENTS) | Status (Certifié/Etalonné/En cours/Réaliser) | (valeur dans col F) | **ORPHELIN** — Colonne avec validation list. Non extraite (parser cherche juste `jsa`). |
| (après col 25 sample) | JSA détails, TRIF/LTIF IOGP, heures travaillées, near-miss | (non visible) | **ORPHELIN (PROBABLE)** — Sample limité à 25 lignes sur 545. Structure massive de reporting QHSE (26/27 colonnes). **Blindspot probable**. |
| (texte libre) | Observations sécurité (fusionnées) | (multi-ligne) | **ORPHELIN** — Observations libres non structurées. Parser ignore complètement. |

**Observations**
- **Sample insuffisant** : 25 lignes sur 545, **88 % des données = invisible**
- Parser = **ultra-basique** : lit effectifs + recherche texte `jsa`
- **Fallback robustesse** : si positional parsing échoue, scanne texte "jours sans acc"
- **Vrai contenu QHSE** (JSA détails, TRIF, LTIF IOGP, heures, météo, équipements) = **non visible dans sample, probablement orphelin**

---

## 08 — ACTIONS.xlsx

### Onglets (3)
- `11 05 26 (2)`
- `18 05 26 (2)`
- `25 05 26`
(33 lignes × 2 colonnes seulement)

### Colonnes PRÉSENTES
| Colonne | Libellé | Exemple |
|---------|---------|---------|
| A | Section/Catégorie | I. Actions menées semaine S-1, II. Actions à mener Semaine en cours, III. Difficultés/Dysfonctionnements/Incidents, IV. (non visible), V. (non visible) |
| B | Détail actions | · Act 1 : entrtien des yards FOXTROT, ... ; · Act 2 : sensibilisation des chauffeurs, ... |

### Champs EXTRAITS par parseActions()
```
- sections: { 'I. Actions menées ...': [ action_text1, action_text2, ... ], 'II. Actions à mener ...': [...], ... }
- total: count distinct actions
```

### Champs ORPHELINS
| Colonne/Aspect | Libellé | Exemple | Risque/Note |
|---------|---------|---------|------------|
| (structure interne) | Sections IV, V (au-delà ligne 25) | (non visible) | **ORPHELIN (PROBABLE)** — Sample 25/33 lignes. Sections IV/V coupées. Structure complète inconnue. |
| (métadonnées) | État action (En cours, Fermée, Risquée, etc) | Implicite dans texte | **ORPHELIN** — Pas de colonne status. Actions = texte libre. Utile : date fermeture, responsable, KPI réalisation. |
| (liaison) | Traçabilité → incidents/accidents | Référence incident? | **ORPHELIN** — Actions non liées à QHSE/incidents (fichiers séparés). Corrélation perdue. |

**Observations**
- **Format ultra-simplifié** : 2 colonnes seulement = **minimal actionnable**
- **Parser = extraction texte brut** : aucune structure métadonnées
- **Sections IV/V cachées** : structure réelle inconnue (>33 lignes ou hors sample)
- **Cas critique** : zéro métadonnées (date, responsable, état, priorité, KPI)

---

## 🎯 PRIORITÉS DE RÉCUPÉRATION (TOP 10)

### Rang 1 : **ACCOSTAGE — Blocs GOTTWALD & MAFI**
- **Colonnes** : L-P (GOTTWALD), R-U (MAFI)
- **Champs** : mobilisation/démobilisation heures, nombre mouvements, shifts jour/nuit
- **Impact** : Opérabilité grues (KPI critique O&G)
- **Effort** : +2 colonnes index, +2 agrégations (byGottwald, byMafi)
- **Value** : 🔴 HAUTE — Suivi équipement essentiellement manquant

### Rang 2 : **CA — Onglet "Gestion base logistique"** 
- **Colonnes** : 40 colonnes × 35 lignes (P&L complet)
- **Champs** : Prévisions/Réels mensuels par ligne (Salaires, Frais, Amortissements, Assurances, Impôts)
- **Impact** : Reporting financier opérationnel (base=0 extraction)
- **Effort** : +1 parser dédié (~parseFinance) pour scans lignes TOTAL
- **Value** : 🔴 HAUTE — Contrôle budgétaire annuel opérationnel

### Rang 3 : **PLANNING — Colonne DEMANDE EMISE (D)**
- **Colonnes** : D
- **Champs** : Matériel/ressources demandées (4t, Sr40, grues, équipes)
- **Impact** : Capacité logistique par opération
- **Effort** : +1 findIndex() pour colonne, stockage dans ops[i].demande
- **Value** : 🔴 HAUTE — Planning = allocution ressources

### Rang 4 : **CHARGES — Colonnes Mois/Statut/Mode Paiement (F/G/H)**
- **Colonnes** : F, G, H (à partir Mars)
- **Champs** : Mois déd., Statut justif., Mode paiement
- **Impact** : Traçabilité charges / profiling par statut
- **Effort** : +3 index, +3 stockage lignes[]
- **Value** : 🟠 MOYENNE — Complément structure (new dans Mar-Mai)

### Rang 5 : **ACCOSTAGE — Distinction QUAI 25 vs QUAI 21**
- **Colonnes** : H, I (actuellement fusionnées)
- **Champs** : quai spécifique (25 vs 21)
- **Impact** : KPI berth-specific (capacity planning quais)
- **Effort** : +1 champ quai25/quai21 instead of quaiSet
- **Value** : 🟠 MOYENNE — Planning portuaire

### Rang 6 : **CHARGES — N° UM (B)**
- **Colonnes** : B
- **Champs** : Code équipement/service (402301, etc.)
- **Impact** : Joignabilité avec master équipement
- **Effort** : +1 index, stockage lignes[].numUM
- **Value** : 🟠 MOYENNE — Traçabilité audit

### Rang 7 : **POINTAGE — Produits (H) & colonnes validation (N-Q)**
- **Colonnes** : H, N-Q
- **Champs** : Produit assigné (VLOOKUP), validations
- **Impact** : Affectation worker × produit/service
- **Effort** : **BLOCAGE** — VLOOKUP non résolue dans extract; besoin dépendance SOURCE sheet
- **Value** : 🟠 MOYENNE — Allocation workforce (bloquée)

### Rang 8 : **QHSE — Pleine couverture (au-delà L25)**
- **Colonnes** : Tout (25 colonnes, 545 lignes)
- **Champs** : Météo, équipements status, JSA détails, TRIF/LTIF, heures, near-miss
- **Impact** : Reportage QHSE complet
- **Effort** : **CRITIQUE** — Sample insuffisant (25/545). Nécessite relancer audit avec limit_rows=545
- **Value** : 🔴 HAUTE — Safety reporting (compliance)

### Rang 9 : **CA — Onglets orphelins (Occupation, Représentations, Evolution clients)**
- **Onglets** : Occupation (2), Représentations, Evolution clients, SOA Envoyés, PROGRAMME 2014
- **Champs** : Occupation espace/client (Superficie/Price/Amount), Taux rétrocession %, Dynamique clients, Historique
- **Impact** : Vue métier complémentaire (rentabilité espace, profil clients, archéologie)
- **Effort** : +2-3 parsers spécialisés
- **Value** : 🟡 MOYENNE-BASSE — Contexte métier (historique/analyses)

### Rang 10 : **LOYERS — Onglet "Feuil1" (Espaces dédiés)**
- **Onglets** : Feuil1
- **Champs** : Référentiel clients × espaces (2 colonnes, 46 lignes)
- **Impact** : Dictionnaire espaces (jointure avec loyers)
- **Effort** : +1 parser simple (flatten A,B)
- **Value** : 🟡 MOYENNE-BASSE — Référentiel

---

## ⚠️ RISQUES & PROBLÈMES CRITIQUES

### 1. **POINTAGE — Dépendance SOURCE sheet**
- **Problème** : Colonnes D-H = VLOOKUP vers `[2]SOURCE!$B$6:$G$105`
- **Impact** : Données non accessibles (SOURCE = sheet non distribuée?)
- **Recommandation** : **Intégrer SOURCE sheet dans template** ou **remplacer VLOOKUP par colonnes libres**

### 2. **QHSE — Sample de 25 lignes sur 545 (95 % invisible)**
- **Problème** : Rapport audit tronqué ; vraie structure QHSE inconnue
- **Recommandation** : **Relancer audit avec `limit_rows=545`** ou **accès au fichier complet**

### 3. **PLANNING — 2 formats (JOUR vs NUIT)**
- **Problème** : Nuits = structure réduite vs JOUR (9 col vs 12 col)
- **Impact** : Perte détails sur nuits (DEMANDE EMISE, OBSERVATIONS)
- **Recommandation** : **Uniformiser templates** ou **parser multi-format**

### 4. **CA — Nombreux onglets archive/unused**
- **Problème** : 7+ onglets (PROGRAMME 2014, SOA 2017, Occupation, Représentations, etc.) = clutter
- **Recommandation** : **Nettoyer fichier** (archiver séparément) ou **documenter intention**

### 5. **En-têtes non normalisés**
- **LOYERS** : "Libéllé" (faute E accent)
- **ACCOSTAGE** : "ETA" vs "Arriv" vs "Accost" (variations)
- **CHARGES** : "Libellé" vs "Libéllé" (casse/accent)
- **Recommandation** : **Standardiser labels** + regex robustes

### 6. **Cellules fusionnées massives (QHSE, CA)**
- **Impact** : Extraction difficile, samples biaisés
- **Recommandation** : **Éviter fusionnées** ; utiliser répétition ou "merged header row"

---

## 📋 TEMPLATE EXCEL "STRICT" — Recommandations

### Structure générique
```
FEUILLE 1 : DONNÉES
  ├─ En-têtes (ligne 1) : labels NORMALISÉS
  ├─ Données (lignes 2+)
  └─ Pas de fusionnées, pas de vides, pas de sous-totaux mélangés

FEUILLE 2 : RÉFÉRENTIELS (si applicable)
  ├─ Master clients, Master équipements, Master espaces, etc.
  └─ Format : ID | Libellé | Attributs

FEUILLE 3 : PARAMÈTRES (optionnel)
  └─ Métadonnées : mois, année, version, responsable
```

### Par type Excel

#### 01 LOYERS
- **Ajouter feuille SOURCE** (ou intégrer ref. depuis Feuil1)
- **Colonnes fixes** : Client | Libellé | Montant | M1 | M2 | ... | M12
- **Éliminer Feuil1** ou **fusionner dans onglet principal**

#### 02 CHARGES
- **Colonnes universelles** (tous mois) : Libellé | N°UM | Fournisseur | Affectation | Montant H.T | **Mois** | **Statut** | **Mode Paiement**
- **Valider Mois/Statut/Mode Paiement** dès janvier (rétro-compatibilité)

#### 03 PLANNING
- **2 formats harmonisés** : Jour complet = Nuit complet
- **Colonnes nuit** : Ajouter DEMANDE EMISE, OBSERVATIONS (non juste opérations basiques)

#### 04 POINTAGE
- **Intégrer SOURCE sheet** ou **remplacer VLOOKUP par colonnes libres**
- **Colonnes à ajouter** : Produit (texte libre), Validation chef (dropdown), Signature (texte)

#### 05 CA
- **Fusionner Gestion base logistique** dans onglet "Récap Global" (ou créer "Finances")
- **Éliminer onglets archive** (PROGRAMME 2014, SOA 2017)
- **Standardiser** : Cumul annuel = ref. "Vérite" pour totaux

#### 06 ACCOSTAGE
- **Ajouter blocs complets** : GOTTWALD (DATE, MOB/H, DEMOB/H, Nbre H, BL), MAFI (DATE, SHIFT JOUR, SHIFT NUIT)
- **Distinction quai** : Colonne separate "QUAI_POSTE" (25, 21, ou texte libre)

#### 07 QHSE
- **Standard : 545 lignes = 4 semaines**
- **Clarifier structure** : Sections (Météo, Personnel, Équipements, Opérations, Incidents, Actions) = lignes de début claires
- **Colonnes explicites** : EffAGL, EffSousTrait, EffJournaliers, JSA, TRIF, LTIF, Heures, NearMiss, EPI Status, etc.

#### 08 ACTIONS
- **Ajouter métadonnées** : Colonne status (Ouvert, En cours, Fermé), Responsable, Date prévue, Date réalisée, Priorité
- **Clarifier sections** : I–V = catégories fixes (Actions menées, À mener, Incidents, Difficultés, Apprentissages)

---

## 📊 SYNTHÈSE QUANTITATIVE

| Fichier | Onglets | Colonnes présentes | Colonnes exploitées | % Orphelin | Onglets orphelins |
|---------|---------|-----------|-----------|-----------|----------|
| LOYERS | 2 | 15 | 3 (Client, Libellé, Montants) | 80% | 1 (Feuil1) |
| CHARGES | 5 | 8 (moy) | 5 | 37% | 0 |
| PLANNING | 51 | 12 | 6 | 50% | 1 (Feuil1) |
| POINTAGE | 55 | 17 | 5 | 71% | 0 |
| CA | 21 | 40+ | 12 | 70% | 7+ (archive/unused) |
| ACCOSTAGE | 4 | 21-22 | 6 | 71% | 0 |
| QHSE | 4 | 25 (545 L) | 5-10 (blind) | 60%+ | 0 |
| ACTIONS | 3 | 2 | 2 | 0% | 0 |

**Moyenne générale : ~60 % orphelin**

---

## 🎯 ACTIONS IMMÉDIATES

1. ✅ **Relancer audit QHSE** avec `limit_rows=545` (voir sections IV-V)
2. ✅ **Template PLANNING** : harmoniser JOUR/NUIT, ajouter DEMANDE EMISE
3. ✅ **Template ACCOSTAGE** : inclure blocs GOTTWALD & MAFI complets
4. ✅ **Template CA** : absorber Gestion base logistique ou créer onglet Finances
5. ✅ **Template CHARGES** : valider Mois/Statut/Mode Paiement sur tous mois
6. ✅ **Template POINTAGE** : résoudre dépendance SOURCE ou refondre VLOOKUP
7. ✅ **Nettoyage CA** : archiver PROGRAMME 2014, SOA Envoyés, etc.
8. ✅ **Normalisation** : Label unifiés (accent, casse) + regex robustes

---

**Rapport généré** : 03 Juin 2026  
**Audit scope** : 8 fichiers Excel × 8 parsers JS  
**Méthodologie** : Diff sample × code source  
**Prochaine étape** : Conception templates "100 % complet" avec opérateurs

