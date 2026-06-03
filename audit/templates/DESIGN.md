# Templates Excel "stricts" — Design détaillé
**Version** : 1.0 — Phase 2
**Cible** : opérateurs base O&G AGL Côte d'Ivoire
**Mode** : Strict (validations obligatoires)
**Référence audit** : `audit/reports/SYNTHESE_ORPHELINS.md` (60 % de données perdues en moyenne)
**Cible technique** : `/home/user/segmentation-client/O_G_AGL_Analyseur_OFFLINE_v23.html` (parsers L1518-1854)

---

## 0. Convention générale

### Identité visuelle AGL
- **Couleurs** : Navy `#0A2540` (en-têtes), Gold `#C9A14A` (accents), Blanc texte sur navy
- **Police** : Calibri 11 (corps), Calibri Bold 12 (en-têtes ligne 1), Calibri Bold 14 (titres mode d'emploi)
- **Logo** : zone réservée en haut du Mode d'emploi (cellule A1:D4)

### Structure systématique de chaque fichier
1. Onglet `📖 Mode d'emploi` — **toujours en première position**
2. Onglet(s) `Données` — saisie opérateur (lignes 2+, en-têtes verrouillées ligne 1)
3. Onglet(s) `Référentiels` locaux si besoin (espaces, codes UM…)
4. Onglet `🔒 _SYS` (caché, listes nommées source des dropdowns)

### Règles strictes communes
- **Cellules figées** : Ligne 1 (en-têtes) + Colonne A
- **Verrouillage en-têtes** : ligne 1 protégée (mot de passe `AGL-OG-2026`)
- **Aucune fusion** dans la zone de données (ligne 2+)
- **Format dates** : `YYYY-MM-DD` strict (validation Excel `Date >= 2020-01-01`)
- **Format heures** : `HH:MM` 24 h (validation `Heure`)
- **Format montants** : nombre brut, sans `XOF`, sans espace, point décimal (l'app formate à l'affichage)
- **Champs obligatoires** : surlignés en gold + validation "Aucun vide"
- **Listes déroulantes** : pointent vers plages nommées de `00_MASTER_REFERENTIELS.xlsx` ou onglet `_SYS` local
- **Booléens** : valeurs autorisées `OUI` / `NON` uniquement (jamais `1/0`, sauf colonnes comptage)
- **Encodage** : UTF-8, accents conservés (FR), noms propres en MAJUSCULES
- **Nom de fichier** : `NN_NOM_vX.xlsx` (numéro de fichier + nom court + version)
- **Versioning** : cellule `_SYS!A1` = numéro de version + date génération

### Conventions parser
- Les en-têtes utilisent les libellés **exacts** que les parsers JS recherchent (`client`, `montant`, `matricule`, `fonction`, `navire`, `nature`, `positif`, `néga`, `libellé`, etc.).
- Le code court (col "Code") sert au futur mapping JSON (export structuré post-parsing).

---

## 1. Master Référentiels (fichier `00_MASTER_REFERENTIELS.xlsx`)

Fichier unique chargé en optionnel par l'app au démarrage (lecture seule), maintenu par le KAM.

### Onglet `Clients`
| # | Code | En-tête | Type | Validation | Obligatoire | Exemple |
|---|------|---------|------|------------|-------------|---------|
| 1 | id_client | ID Client | TEXTE | Code unique (3-15 car.) | OUI | FOXTROT |
| 2 | nom_client | Nom Client | TEXTE | Unique | OUI | FOXTROT |
| 3 | nom_legal | Raison sociale | TEXTE | — | NON | FOXTROT CI SA |
| 4 | segment | Segment | LISTE | Forage / Production / Service / Sismique / Construction / Marine | OUI | Forage |
| 5 | actif | Actif | LISTE | OUI/NON | OUI | OUI |
| 6 | alias | Alias (parsing) | TEXTE | Séparateur `;` | NON | FOX;FOXTROT FORAGE |

**Valeurs initiales** (pré-remplies, issues audit) :
FOXTROT, MURPHY, SLB, SAIPEM SEI, SAIPEM OM, EXPRO, IVOS, DRIL QUIP, HALLIBURTON, GEV, BAKER HUGHES, BOSKALIS, NAVITRANS, AFRITRAMP, CNR, PAA, FRENCH ARMY, ANTRAK GHANA, ANADARKO.

### Onglet `Espaces`
| # | Code | En-tête | Type | Validation | Obligatoire | Exemple |
|---|------|---------|------|------------|-------------|---------|
| 1 | id_espace | ID Espace | TEXTE | Unique | OUI | MAG-T10 |
| 2 | libelle_espace | Libellé | TEXTE | — | OUI | Magasin T10 |
| 3 | type_espace | Type | LISTE | Magasin / Yard / Bureau / Atelier / Protocole / Mixte | OUI | Magasin |
| 4 | superficie_m2 | Superficie (m²) | NOMBRE | >0 | OUI | 4500 |
| 5 | client_dedie | Client dédié | LISTE | Master.Clients ou vide | NON | FOXTROT |
| 6 | tarif_m2 | Tarif m²/mois (XOF) | NOMBRE | >=0 | NON | 6000 |

**Valeurs initiales** : extraites de LOYERS.Feuil1 (46 lignes — protocoles MURPHY/FOXTROT, Mag T10/T20, Yards 3262 m², etc.).

### Onglet `Equipements`
| # | Code | En-tête | Type | Validation | Obligatoire | Exemple |
|---|------|---------|------|------------|-------------|---------|
| 1 | id_equip | ID Équipement | TEXTE | Unique | OUI | CE-16T-01 |
| 2 | type_equip | Type | LISTE | CHARIOT ELEVATEUR / GOTTWALD / MAFI / TRACTEUR / GRUE MOBILE / CAMION / Autre | OUI | CHARIOT ELEVATEUR |
| 3 | tonnage | Tonnage | LISTE | 3T/7T/12T/13T/16T/25T/45T/Autre | OUI | 16T |
| 4 | nb_disponible | Nb disponible | NOMBRE | >=0 | OUI | 3 |
| 5 | status_cert | Status certification | LISTE | Certifié / Etalonné / En cours / À réaliser | OUI | Certifié |
| 6 | date_prochain_ctrl | Prochain contrôle | DATE | >= aujourd'hui | NON | 2026-09-30 |
| 7 | num_um | N° UM (immat.) | TEXTE | — | NON | 402301 |

**Valeurs initiales** : issues de QHSE.section Équipements (CHARIOT ELEVATEUR 16T/13T/12T, GOTTWALD, MAFI) + CHARGES.B (N° UM).

### Onglet `Sections_Planning`
| # | Code | En-tête | Type | Obligatoire | Exemple |
|---|------|---------|------|-------------|---------|
| 1 | code_section | Code | TEXTE (unique) | OUI | CWI |
| 2 | libelle | Libellé long | TEXTE | OUI | Container Wash & Inspection |
| 3 | actif | Actif | LISTE OUI/NON | OUI | OUI |

**Valeurs initiales** : CWI (SHAN), OIC, TRS, OFE, MAG, YARD, QUAI.

### Onglet `Natures_Operations`
| # | Code | En-tête | Type | Obligatoire | Exemple |
|---|------|---------|------|-------------|---------|
| 1 | id_nature | ID | TEXTE | OUI | NAT-01 |
| 2 | libelle_nature | Libellé | TEXTE | OUI | Activités en magasin et sur yard |
| 3 | famille | Famille | LISTE | OUI | Manutention |

### Onglet `Affectations`
Pour CHARGES — codes comptables analytiques.
| # | Code | En-tête | Obligatoire | Exemple |
|---|------|---------|-------------|---------|
| 1 | code_aff | Code | OUI | 40B00 |
| 2 | libelle_aff | Libellé | OUI | Base logistique commune |

### Onglet `Fournisseurs`
| # | Code | En-tête | Obligatoire | Exemple |
|---|------|---------|-------------|---------|
| 1 | id_four | ID | OUI | F-AYTHAM |
| 2 | nom_four | Nom | OUI | AYTHAM SERVICES |
| 3 | mode_paiement_defaut | Mode paiement par défaut | NON | Virement |

### Onglet `Departements_CA`
Pour le CA — segmentation revenus.
- Libellés (issus de Récap Global) : `BASE LOYER`, `Plog & Divers`, `IMAR`, `HINT`, `IAER`, `EAER`, `AÉRIEN`, `EXPORT`, `SIRT`, `Afritramp`, `Meet&Greet`, `Autres`, `TOTAL`.

### Onglet `Postes_Charges_PNL`
Pour `Gestion base logistique` (P&L) — lignes de compte d'exploitation.
- Libellés (issus de l'onglet original) : `CAP NET`, `Salaires`, `Charges sociales`, `Frais généraux`, `Amortissements`, `Assurances`, `Impôts & taxes`, `Sous-traitance`, `Carburant`, `Maintenance`, `Loyers`, `Autres charges`, `RÉSULTAT D'EXPLOITATION`.

### Onglet `Quais`
| Code | Libellé |
|------|---------|
| Q25 | Quai 25 |
| Q21 | Quai 21 |

### Onglet `📖 Mode d'emploi`
Texte : "Ce fichier est le référentiel partagé. Le KAM le met à jour 1×/mois. Ne pas modifier la structure des onglets. Pour ajouter une valeur : aller dans l'onglet concerné, ajouter une ligne en bas, sauvegarder. Les autres fichiers se synchronisent à l'ouverture."

### Impact parser
- L'app (HTML) chargera ce fichier optionnellement au démarrage via un nouvel input `<input id="masterFile">`.
- Fonction `loadMaster(wb)` → enrichit les variables globales `CONTEXTE_AGL.masters = {clients, espaces, ...}`.
- Utilisé pour : canonisation noms (cf. `canonClient()` existant L1822), pré-population dropdowns dans Mode d'emploi futur.

---

## 2. Template LOYERS (`01_LOYERS_v1.xlsx`)

### Identification
- **Rythme** : annuel (mise à jour mensuelle des colonnes M1-M12 par le KAM)
- **Responsable** : KAM
- **Justification choix** : structure conservée (1 ligne client × espace × N mois), Feuil1 archivée vers Master Espaces.

### Onglets
1. `📖 Mode d'emploi`
2. `Loyers` (données)
3. `_SYS` (caché)

### Onglet `Loyers` — Colonnes

| # | Code | Libellé en-tête | Type | Validation | Obligatoire | Source/Origine | Exemple | Note parser |
|---|------|-----------------|------|------------|-------------|----------------|---------|-------------|
| A | client | Client | LISTE | Master.Clients | OUI | Référentiel Clients | FOXTROT | `byClient` clé (L1672) |
| B | id_espace | ID Espace | LISTE | Master.Espaces | OUI | Référentiel Espaces | MAG-T10 | 🆕 jointure orpheline |
| C | libelle_facturation | Libellé facturation | TEXTE | longueur 5-200 | OUI | Saisie | Protocole Mag T10 4500m² & Yard 3262m² | Conservé existant |
| D | montant_ht_annuel | Montant H.T annuel (XOF) | NOMBRE | >0 | OUI | Saisie / =SUM(E:P) | 27469000 | Existant |
| E-P | m1 à m12 | janv. 26 / févr. 26 / … / déc. 26 | NOMBRE | >=0 | NON | Saisie ou =D/12 | 2289083 | `periods[i].byClient` (L1674) — **les en-têtes M1-M12 doivent être des dates Excel valides** (1er du mois), exigence du parser L1667 `isNum(h[c])` |
| Q | commentaire | Commentaire | TEXTE | — | NON | Saisie | Ajout avenant 2026-03 | Nouveau, ignoré parser |

### Mode d'emploi
> Le KAM saisit 1 ligne par contrat client × espace. Le montant annuel HT se ventile mensuellement en M1-M12 (le 1er de chaque mois). Si avenant en cours d'année, créer une 2e ligne avec libellé "Avenant" en commentaire. Les en-têtes M1-M12 sont des dates Excel (pas du texte), ne pas modifier.

### Champs récupérés vs existant
- ✅ Client, Libellé facturation, Montant HT, M1-M12 (existants conservés)
- 🆕 **ID Espace** (jointure avec ex-Feuil1 désormais dans Master Espaces) — **orphelin Rang 10 traité**
- 🆕 Commentaire (libre)
- ❌ Feuil1 supprimée du fichier loyers, déplacée vers `00_MASTER_REFERENTIELS.xlsx > Espaces`

### Impact parser
- `parseLoyers()` (L1650) : aucun changement requis (lit premier onglet, en-têtes inchangées).
- Enrichissement futur : si ID Espace présent, joindre avec Master.Espaces pour ajouter `byEspace`, `bySegmentEspace`.

---

## 3. Template CHARGES (`02_CHARGES_v1.xlsx`)

### Identification
- **Rythme** : mensuel (1 onglet par mois `MM-YY`)
- **Responsable** : Compta / Ops
- **Justification choix** : on harmonise la structure dès janvier (les colonnes Mois/Statut/Mode Paiement n'existaient qu'à partir de Mars → traite **orphelin Rang 4**).

### Onglets
1. `📖 Mode d'emploi`
2. `01-26`, `02-26`, …, `12-26` (12 onglets, schéma identique)
3. `_SYS`

### Onglet mensuel — Colonnes

| # | Code | Libellé en-tête | Type | Validation | Obligatoire | Source | Exemple | Note parser |
|---|------|-----------------|------|------------|-------------|--------|---------|-------------|
| A | libelle | Libellé | TEXTE | 5-200 car. | OUI | Saisie | Location Tracteur routier 4x2 | `lib` (L1642) |
| B | num_um | N° UM | TEXTE | Pattern `\d{6}` ou vide | NON | Master.Equipements | 402301 | 🆕 **orphelin Rang 6** traité |
| C | fournisseur | Fournisseur | LISTE | Master.Fournisseurs (autre = saisie libre autorisée) | OUI | Référentiel | AYTHAM SERVICES | `byFour` (L1641) |
| D | affectation | Affectation | LISTE | Master.Affectations | OUI | Référentiel | 40B00 | `byAff` (L1640) |
| E | montant_ht | Montant H.T (XOF) | NOMBRE | >0 | OUI | Saisie | 3000000 | `montant` (L1639) |
| F | mois | Mois | LISTE | JANVIER / FÉVRIER / … / DÉCEMBRE | OUI dès janvier | Saisie (auto = nom onglet) | JANVIER | 🆕 **orphelin Rang 4** |
| G | statut | Statut | LISTE | À justifier / Justifiée / Validée / Litigieuse / Annulée | OUI | Saisie | Validée | 🆕 **orphelin Rang 4** |
| H | mode_paiement | Mode Paiement | LISTE | Virement / Chèque / Espèces / Carte / Compensation / Autre | OUI | Saisie | Virement | 🆕 **orphelin Rang 4** |
| I | date_facture | Date facture | DATE | YYYY-MM-DD | NON | Saisie | 2026-01-15 | Nouveau |
| J | num_facture | N° Facture | TEXTE | — | NON | Saisie | F-2026-0123 | Nouveau |
| K | commentaire | Commentaire | TEXTE | — | NON | Saisie | — | — |

### Mode d'emploi
> Une ligne = une charge. Saisir d'abord Libellé/Montant/Fournisseur/Affectation. Statut "Validée" requis avant paiement. Le N° UM permet de joindre avec l'équipement (Master.Equipements). Reproduire la même structure tous les mois — **ne pas supprimer les colonnes F/G/H même si vides**.

### Champs récupérés vs existant
- ✅ Libellé, Fournisseur, Affectation, Montant HT
- 🆕 N° UM (B), Mois (F), Statut (G), Mode Paiement (H) — **4 orphelins traités**, harmonisés tous mois
- 🆕 Date facture, N° Facture (traçabilité)

### Impact parser
- `parseCharges()` (L1628) : ajouter `iUM`, `iStatut`, `iPaiement` (findIndex), enrichir `lignes[].numUM/statut/paiement` ; ajouter agrégats `byStatut`, `byPaiement`.
- Render : nouvel onglet "Charges par statut" + "Charges par mode paiement".

---

## 4. Template PLANNING (`03_PLANNING_v1.xlsx`)

### Identification
- **Rythme** : quotidien — 2 onglets par jour (JOUR + NUIT), un onglet WK (weekend) et FERIE si applicable
- **Responsable** : Ops (chef de quart)
- **Justification choix** : harmoniser JOUR et NUIT (mêmes colonnes) pour éliminer la perte d'info nuit (cf. observation parser L1605 `shift`). Suppression de Feuil1 (orphelin résolu).

### Onglets
1. `📖 Mode d'emploi`
2. `DD-MM-YYYY-JOUR`, `DD-MM-YYYY-NUIT` (générés à la demande par script ou modèles)
3. `_SYS`

### Onglet quotidien — Colonnes (identiques JOUR et NUIT)

| # | Code | Libellé en-tête | Type | Validation | Obligatoire | Source | Exemple | Note parser |
|---|------|-----------------|------|------------|-------------|--------|---------|-------------|
| A | num_ordre | N° | NOMBRE | >0 | OUI | Saisie / auto | 1 | Nouveau |
| B | client | CLIENT | LISTE | Master.Clients | OUI | Référentiel | BAKER HUGHES | `iCli` (L1590) |
| C | section | SECTION | LISTE | Master.Sections_Planning | OUI | Référentiel | CWI | `iSec` (L1590) |
| D | demande_emise | DEMANDE EMISE | TEXTE | 1-100 car. | OUI | Saisie | 4t, 01Sr20+01chall, 7t+01ELG | 🆕 **orphelin Rang 3** — `iDemande` à ajouter |
| E | nature | NATURE DE L'OPÉRATION | LISTE/TEXTE | Master.Natures (autre OK) | OUI | Référentiel + saisie | Activités en magasin et sur yard | `iNat` (L1591) |
| F | positif | POSITIF | NOMBRE | 0 ou 1 | OUI | Saisie | 1 | `iPos` (L1591) |
| G | nega | NEGA | NOMBRE | 0 ou 1 | OUI | Saisie | 0 | `iNeg` (L1592) |
| H | observations | OBSERVATIONS | TEXTE | — | NON | Saisie | AJOUT, Location mensuelle | 🆕 **orphelin** — capture à ajouter |
| I | id_equip | Équipement principal | LISTE | Master.Equipements | NON | Référentiel | CE-16T-01 | 🆕 nouveau |
| J | shift | Shift | LISTE | JOUR / NUIT | OUI | Auto (depuis nom onglet) | JOUR | `isNightSheet()` |

### Mode d'emploi
> Chef de quart : 1 ligne = 1 opération. **JOUR et NUIT ont la même structure**. Saisir Client/Section/Demande émise/Nature obligatoirement. POSITIF=1 si l'op a abouti, NEGA=1 si refusée/abandonnée. Observations = flags urgence/contrats. Onglet par date `DD-MM-YYYY-{JOUR|NUIT}` impératif (parser).

### Champs récupérés vs existant
- ✅ CLIENT, SECTION, NATURE, POSITIF, NEGA
- 🆕 **DEMANDE EMISE (D)** — Rang 3 priorité haute
- 🆕 **OBSERVATIONS (H)** structurée
- 🆕 N° ordre, Équipement principal
- ❌ Feuil1 supprimée (était orpheline)
- 🔧 **Harmonisation JOUR=NUIT** : la nuit était à 9-10 colonnes, on impose les 10 colonnes complètes

### Impact parser
- `parsePlanning()` (L1582) : ajouter `iDemande`, `iObs`, `iEquip` (findIndex) ; étendre `ops[i]` avec `demande`, `obs`, `equip`.
- Render : nouvelle vue "Demandes par type de matériel" (regex sur `demande_emise` pour extraire `Nt`, `SrXX`, etc.).

---

## 5. Template POINTAGE (`04_POINTAGE_v1.xlsx`)

### Identification
- **Rythme** : quotidien (JOUR + NUIT)
- **Responsable** : RH / chef de chantier
- **Justification choix** : **intégrer SOURCE en tant qu'onglet local** (problème critique 1 du SYNTHESE). VLOOKUP remplacé par formule autonome ciblant l'onglet `SOURCE_Personnel` du même fichier. Si l'opérateur préfère, mode dégradé "colonnes libres" possible (instruction Mode d'emploi).

### Onglets
1. `📖 Mode d'emploi`
2. `SOURCE_Personnel` (master matricules — maintenu par RH 1×/mois)
3. `DD-MM-YYYY-JOUR`, `DD-MM-YYYY-NUIT`, `…-WK`
4. `_SYS`

### Onglet `SOURCE_Personnel` — Colonnes

| # | Code | En-tête | Type | Obligatoire | Exemple |
|---|------|---------|------|-------------|---------|
| A | matricule | Matricule | TEXTE (unique) | OUI | 11580 |
| B | nom | Nom et Prénoms | TEXTE | OUI | JOHNSON Kouakouvi Nelson |
| C | fonction | Fonction | LISTE | OUI | CARISTE |
| D | code_fonction | Code Fonction | TEXTE | OUI | C-CAR |
| E | chantier_defaut | Chantier/Activité par défaut | LISTE | NON | FOXTROT-MAG |
| F | produit_defaut | Produit par défaut | TEXTE | NON | Manutention |
| G | actif | Actif | LISTE OUI/NON | OUI | OUI |

### Onglet quotidien — Colonnes

| # | Code | En-tête | Type | Validation | Obligatoire | Source | Exemple | Note parser |
|---|------|---------|------|------------|-------------|--------|---------|-------------|
| A | num_ordre | N° | NOMBRE | >0 | OUI | Saisie | 1 | Nouveau |
| B | date | Date | DATE | YYYY-MM-DD | OUI | Auto entête | 2026-05-01 | — |
| C | matricule | Matricule | LISTE | SOURCE_Personnel.A | OUI | Référentiel local | 11580 | `iMat` (L1563) |
| D | nom | Nom et Prénoms | TEXTE | `=VLOOKUP(C2,SOURCE_Personnel!A:G,2,FALSE)` | OUI (calc) | Auto | JOHNSON Kouakouvi Nelson | `iNom` (L1563) |
| E | fonction | Fonction | TEXTE | `=VLOOKUP(...,3,FALSE)` | OUI (calc) | Auto | CARISTE | `iFonc` (L1564) |
| F | code_fonction | Code Fonction | TEXTE | `=VLOOKUP(...,4,FALSE)` | OUI (calc) | Auto | C-CAR | `iCode` (L1564) |
| G | chantier | Chantiers/Activités | LISTE | Master + libre | OUI | Saisie ou défaut | FOXTROT-MAG | `iChant` (L1565) |
| H | produit | Produits | TEXTE | — | NON | Saisie | Manutention conteneurs | 🆕 **orphelin Rang 7** (col H libérée du VLOOKUP cassé) |
| I | shift | Shift | LISTE | JOUR / NUIT | OUI | Auto onglet | JOUR | `isNightSheet()` |
| J | heures_travaillees | Heures travaillées | NOMBRE | 0-24 | NON | Saisie | 8 | Nouveau (input QHSE) |
| K | validation_chef | Validation chef | LISTE | Validé / Refusé / À revoir | OUI | Dropdown | Validé | 🆕 (orphelin N-Q) |
| L | signature | Signature chef | TEXTE | — | NON | Saisie nom | KOUASSI A. | 🆕 (orphelin N-Q) |

### Mode d'emploi
> Le RH maintient `SOURCE_Personnel`. Le chef de chantier ne saisit que **Matricule** (col C), Chantier (G), Produit (H), Heures (J), Validation (K) — les colonnes D/E/F se remplissent automatiquement par VLOOKUP **interne au fichier** (pas de fichier externe). Si VLOOKUP ne fonctionne pas (matricule absent), saisir manuellement D/E/F et avertir le RH.

### Champs récupérés vs existant
- ✅ Matricule, Nom, Fonction, Code Fonction, Chantier
- 🆕 **Produit (H)** sans VLOOKUP externe — Rang 7 partiellement débloqué
- 🆕 Heures travaillées, Validation chef, Signature — colonnes N-Q ex-orphelines
- 🔧 **VLOOKUP refondu** : pointe vers `SOURCE_Personnel` local (problème critique 1 résolu)

### Impact parser
- `parsePointage()` (L1554) : ajouter `iProduit`, `iHeures`, `iValidation` ; étendre `workers[]` avec `produit`, `heures`, `validation`.
- Heures travaillées remontent dans la vue QHSE (cross-link "Heures du mois" pour TRIF/LTIF).

---

## 6. Templates CA — éclatement en 3 fichiers

**Justification choix (consigne 4)** : l'onglet CA original mélangeait 21 vues hétérogènes (matrice mois×dép, détail client, P&L, occupation, historique archive). On éclate en 3 livrables séparés :
- `05A_CA_FLASH_v1.xlsx` — pilotage opérationnel mensuel (Récap Global + Cumul Annuel)
- `05B_CA_PNL_v1.xlsx` — Gestion base logistique (orphelin Rang 2)
- `05C_CA_DETAIL_v1.xlsx` — feuilles mensuelles détail client (budget vs réel)

Les onglets archive (`PROGRAMME 2014`, `SOA Envoyés`, `Représentations`, `Evolution clients`, `Occupation (2)`) sont **archivés séparément** dans `archives/` et non régénérés.

---

### 6A. `05A_CA_FLASH_v1.xlsx`

#### Identification
- **Rythme** : mensuel (mise à jour T+5)
- **Responsable** : KAM / Contrôle de gestion

#### Onglets
1. `📖 Mode d'emploi`
2. `Récap Global`
3. `Cumul Annuel`
4. `_SYS`

#### Onglet `Récap Global` — Colonnes (matrice mois × départements)

| # | Code | En-tête | Type | Obligatoire | Note parser |
|---|------|---------|------|-------------|-------------|
| A | ligne_type | Type | LISTE (Mois / Cumul / Part %) | OUI | Filtrage L1713 |
| B | mois | Mois | LISTE (JANVIER…DÉCEMBRE) | OUI | `MOIS_FR_FULL` (L1684) |
| C | base_loyer | BASE LOYER | NOMBRE >=0 | NON | 🆕 orphelin extrait |
| D | plog_divers | Plog & Divers | NOMBRE >=0 | NON | 🆕 |
| E | total_base | Total BASE | NOMBRE (=C+D) | OUI | — |
| F-N | ca_imar … ca_meet_greet | IMAR / HINT / IAER / EAER / AÉRIEN / EXPORT / SIRT / Afritramp / Meet&Greet | NOMBRE >=0 | NON | 🆕 **orphelins** Rang 9 |
| O | ca_autres | CA GENERE Autres | NOMBRE | NON | — |
| P | ca_total_client | CA TOTAL GENERE | NOMBRE (=SUM(C:O)) | OUI | `byDep` (L1717), `total` (L1720) |

#### Onglet `Cumul Annuel` — Colonnes

| # | Code | En-tête | Type | Obligatoire | Note parser |
|---|------|---------|------|-------------|-------------|
| A | client | Client | LISTE Master.Clients | OUI | `colName` (L1731) |
| B | loyers | Loyers | NOMBRE >=0 | OUI | `colLoyers` (L1735) |
| C | ca_genere | CA Généré (hors loyers) | NOMBRE >=0 | OUI | — |
| D | ca_total | CA TOTAL GENERE | NOMBRE (=B+C) | OUI | `colTot` (L1734) |
| E | reel_n_1 | Réel N-1 | NOMBRE | NON | 🆕 |
| F | budget_annuel | Budget annuel | NOMBRE | NON | 🆕 |
| G | var_fcfa | VAR (FCFA) | NOMBRE (=D-F) | NON | 🆕 |
| H | var_pct | VAR (%) | POURCENTAGE | NON | 🆕 |

### 6B. `05B_CA_PNL_v1.xlsx` — Gestion base logistique (**Orphelin Rang 2**)

#### Identification
- **Rythme** : mensuel
- **Responsable** : Contrôle de gestion
- **Justification** : 0 % extraction actuellement, 25 lignes × 40 colonnes de KPI P&L = gisement critique.

#### Onglets
1. `📖 Mode d'emploi`
2. `PnL_Base` (compte d'exploitation)
3. `_SYS`

#### Onglet `PnL_Base` — Colonnes (format long, 1 ligne = 1 poste × mois × scénario)

| # | Code | En-tête | Type | Validation | Obligatoire | Exemple |
|---|------|---------|------|------------|-------------|---------|
| A | mois | Mois | LISTE (JANVIER…DÉCEMBRE) | OUI | OUI | JANVIER |
| B | code_poste | Code poste | LISTE Master.Postes_Charges_PNL | OUI | OUI | SALAIRES |
| C | libelle_poste | Libellé poste | TEXTE | Auto VLOOKUP | OUI | Salaires |
| D | categorie | Catégorie | LISTE (Produits / Charges / Résultat) | OUI | OUI | Charges |
| E | prevu | Prévu (XOF) | NOMBRE | — | OUI | 45000000 |
| F | reel | Réel (XOF) | NOMBRE | — | OUI | 43200000 |
| G | variation | Variation (XOF) | NOMBRE | =E-F | OUI | 1800000 |
| H | variation_pct | Variation (%) | POURCENTAGE | =G/E | NON | 4% |
| I | commentaire | Commentaire | TEXTE | — | NON | Sous-effectif mai |

**Choix format long vs large** : format long → 1 nouveau parser simple (`parseFinance(wb)`) — préférable à matrice 25×40 (cellules fusionnées initiales étaient ingérables).

### 6C. `05C_CA_DETAIL_v1.xlsx`

#### Identification
- **Rythme** : mensuel (12 onglets)
- **Onglets** : `Janvier`, …, `Décembre`

#### Onglet mensuel — Colonnes (compatible parseCA() Vue 3)
| # | Code | En-tête | Type | Note parser |
|---|------|---------|------|-------------|
| A | client | Client | LISTE Master.Clients | — |
| B | sources_revenus | Sources revenus | TEXTE | — |
| … (cols 17/18/21 ie. Q/R/U) | reel_base / budget / reel_25 | REEL BASE 2026 / Budget 2026 / REEL BASE 2025 | NOMBRE | **Position critique** : `cellNumOrNull(rows[r][16])`, `[17]`, `[20]` (L1764) — **conserver l'ordre exact des colonnes**. |
| Q | reel_base_2026 | REEL BASE 2026 (XOF) | NOMBRE | idx 16 |
| R | budget_2026 | Budget 2026 (XOF) | NOMBRE | idx 17 |
| S | var_fcfa | VAR (FCFA) | NOMBRE | — |
| T | var_pct | Var (%) | POURCENTAGE | — |
| U | reel_base_2025 | REEL BASE 2025 (XOF) | NOMBRE | idx 20 |

#### Mode d'emploi (commun aux 3)
> 05A = vue pilotage rapide (mois × département + client). 05B = compte d'exploitation détaillé (1 ligne par poste × mois × prévu/réel). 05C = backup historique des détails clients (12 onglets). Respecter les positions de colonnes en 05C (parser hardcodé idx 16/17/20).

### Champs récupérés (cumul des 3 fichiers) vs existant
- ✅ Toutes les vues parseCA actuelles préservées (Récap, Cumul, Budget mensuel)
- 🆕 **Gestion base logistique** (PnL_Base) — Rang 2
- 🆕 BASE LOYER, Plog & Divers, départements IMAR/HINT/IAER/EAER/AÉRIEN/EXPORT/SIRT — Rang 9
- 🆕 REEL N-1, Budget annuel, VAR explicites
- ❌ Onglets archivés (PROGRAMME 2014, SOA Envoyés, Représentations, Evolution clients, Occupation (2)) → fichier `archives/CA_ARCHIVE_pre2026.xlsx` documenté

### Impact parser
- `parseCA()` (L1693) : élargir Vue 1 pour lire toutes colonnes `byDep` (déjà supporté) ; ajouter `parseFinance(wb)` lisant `05B.PnL_Base`.
- Render : nouvel onglet "Compte d'exploitation" avec graphe Prévu vs Réel par poste.

---

## 7. Template ACCOSTAGE (`06_ACCOSTAGE_v1.xlsx`)

### Identification
- **Rythme** : mensuel (1 onglet par mois)
- **Responsable** : Ops portuaires
- **Justification choix** : intégrer **blocs GOTTWALD + MAFI complets** (orphelin Rang 1 — 10 colonnes), distinguer **Quai 25 vs Quai 21** (Rang 5).

### Onglets
1. `📖 Mode d'emploi`
2. `JANVIER 26`, …, `DÉCEMBRE 26`
3. `_SYS`

### Onglet mensuel — Colonnes

| # | Code | En-tête | Type | Validation | Obligatoire | Source | Exemple | Note parser |
|---|------|---------|------|------------|-------------|--------|---------|-------------|
| A | num_ordre | N° | NOMBRE | >0 | OUI | Saisie | 1 | — |
| B | date | DATE | DATE | YYYY-MM-DD | OUI | Saisie | 2026-01-01 | `iDate` (L1804) |
| C | navire | NAVIRE | TEXTE | 3-50 car. | OUI | Saisie | BLUE MOMMY | `iNav` (L1802) |
| D | client | CLIENT | LISTE Master.Clients | — | OUI | Référentiel | MURPHY | `iCli` (L1803) |
| E | eta | ETA | HEURE | HH:MM | OUI | Saisie | 11:45 | `iEta` (L1810) |
| F | etd | ETD | HEURE | HH:MM | OUI | Saisie | 15:45 | `iEtd` (L1811) |
| G | nbre_heures | NBRE HEURES | NOMBRE | =F-E heures | OUI | Calcul | 4.0 | Conservé (validation) |
| H | quai_poste | QUAI POSTE | LISTE | Master.Quais (Q25/Q21) | OUI | Référentiel | Q25 | 🆕 **Rang 5** — remplace H/I fusionnées |
| I | montant_xof | MONTANT (XOF) | NOMBRE | >=0 ou `FREE` | OUI | Saisie | 1500000 | `iMont` (L1805) |
| **— Bloc GOTTWALD (orphelin Rang 1) —** | | | | | | | | |
| J | gottwald_date | GOTTWALD DATE | DATE | YYYY-MM-DD | NON | Saisie | 2026-02-06 | 🆕 |
| K | gottwald_mob_h | GOTTWALD MOB/H | HEURE | HH:MM | NON | Saisie | 00:30 | 🆕 |
| L | gottwald_demob_h | GOTTWALD DEMOB/H | HEURE | HH:MM | NON | Saisie | 05:30 | 🆕 |
| M | gottwald_nbre_h | GOTTWALD Nbre H | NOMBRE | >=0 | NON | Saisie | 5 | 🆕 |
| N | gottwald_bl | GOTTWALD BL | TEXTE | — | NON | Saisie | 260 ou 261-262 | 🆕 |
| **— Bloc MAFI (orphelin Rang 1) —** | | | | | | | | |
| O | mafi_date | MAFI DATE | DATE | YYYY-MM-DD | NON | Saisie | 2026-03-05 | 🆕 |
| P | mafi_shift_jour | MAFI SHIFT JOUR | NOMBRE | 0 ou 1 | NON | Saisie | 1 | 🆕 |
| Q | mafi_shift_nuit | MAFI SHIFT NUIT | NOMBRE | 0 ou 1 | NON | Saisie | 0 | 🆕 |
| R | mafi_total | MAFI TOTAL | NOMBRE | =P+Q | NON | Calcul | 1 | 🆕 |
| S | observations | Observations | TEXTE | — | NON | Saisie | — | Nouveau |

### Mode d'emploi
> Une ligne = une escale. Renseigner Date/Navire/Client/ETA/ETD/Quai/Montant. **QUAI POSTE** est obligatoire (Q25 ou Q21). Les blocs **GOTTWALD** (J-N) et **MAFI** (O-R) sont remplis uniquement si la grue a été utilisée pour cette escale. `MONTANT = FREE` autorisé si gratuit.

### Champs récupérés vs existant
- ✅ N° ordre, Date, Navire, Client, ETA, ETD, Nbre heures, Montant
- 🆕 **Bloc GOTTWALD complet** (5 colonnes) — Rang 1
- 🆕 **Bloc MAFI complet** (4 colonnes) — Rang 1
- 🆕 **Quai poste unique Q25/Q21** — Rang 5 (remplace H/I fusionnées)
- 🆕 Observations

### Impact parser
- `parseAccostage()` (L1793) : ajouter findIndex pour blocs GOTTWALD (`iGottwaldDate`, `iGottwaldMob`, `iGottwaldDemob`, `iGottwaldNbreH`, `iGottwaldBL`) et MAFI (`iMafiDate`, `iMafiJour`, `iMafiNuit`).
- Étendre `lignes[]` avec sous-objets `gottwald: {date, mobH, demobH, nbreH, bl}`, `mafi: {date, jour, nuit, total}`.
- Nouveaux agrégats : `byGottwald: {totalNbreH, totalBLs}`, `byMafi: {totalShifts}`, `byQuai: {Q25: count, Q21: count}` (remplace `quaiSet`).
- Render : 2 KPI cards "Activité GOTTWALD" et "Activité MAFI", graphe occupation par quai.

---

## 8. Templates QHSE — éclatement en sous-onglets (`07_QHSE_v1.xlsx`)

### Identification
- **Rythme** : hebdomadaire (1 fichier = 1 semaine, ou 1 fichier = 4 semaines avec onglets séparés)
- **Responsable** : QHSE Manager
- **Justification choix (consigne 2)** : les 27 sections du fichier original sont éclatées en **6 sous-onglets thématiques** pour la lisibilité opérateur, mais **toute l'information est préservée** (cf. 07_QHSE_DEEP.md). Le parser HTML scanne tous les onglets via `wb.SheetNames.forEach` (L1523) donc peut lire chaque sous-onglet séparément.

### Onglets
1. `📖 Mode d'emploi`
2. `Entete` (méta-données semaine)
3. `QHSE_Meteo`
4. `QHSE_Personnel`
5. `QHSE_Equipements`
6. `QHSE_Incidents` (accidents/incidents/presque-accidents/médical)
7. `QHSE_Securite` (toolbox / permis / JSA / formations / audits)
8. `QHSE_Risques` (matrice causes)
9. `QHSE_Actions` (actions à valoriser)
10. `_SYS`

### Onglet `Entete`
| Code | En-tête | Type | Obligatoire | Exemple |
|------|---------|------|-------------|---------|
| pays | Pays | LISTE | OUI | Côte d'Ivoire |
| entite | Entité | TEXTE | OUI | AFRICA GLOBAL LOGISTICS OIL&GAS |
| semaine | Semaine N° | NOMBRE 1-53 | OUI | 20 |
| date_debut | Date début | DATE | OUI | 2026-05-11 |
| date_fin | Date fin | DATE | OUI | 2026-05-17 |
| nom_redacteur | Rédacteur QHSE | TEXTE | OUI | BLE OLIVIER |
| ref_doc | Référence document | TEXTE | NON | AGL-CIV-SMI-F-0531 Rev 02 |

**Note parser** : la **ligne semaine** doit rester détectable — le parser cherche `semaine` en col L (idx 11) et entier en col M (idx 12) (L1527). On préserve cette position dans `Entete`.

### Onglet `QHSE_Meteo` — **orphelin traité**
| # | Code | En-tête | Type | Obligatoire | Exemple |
|---|------|---------|------|-------------|---------|
| A | jour | Jour | LISTE (Lun…Dim) | OUI | Lun |
| B | ensoleille | Ensoleillé | LISTE OUI/NON | OUI | OUI |
| C | nuageux | Nuageux | LISTE OUI/NON | OUI | NON |
| D | venteux | Venteux | LISTE OUI/NON | OUI | NON |
| E | pluie | Pluie | LISTE OUI/NON | OUI | OUI |
| F | commentaire | Commentaire | TEXTE | NON | Pluie soirée |

### Onglet `QHSE_Personnel`
| # | Code | En-tête | Type | Obligatoire | Note parser |
|---|------|---------|------|-------------|-------------|
| A | categorie | EFFECTIFS | LISTE (AGL / Sous-traitant / Journaliers) | OUI | `c` (L1528-29) — **valeurs strictes** |
| B | cadres | CADRES | NOMBRE >=0 | OUI | — |
| C | non_cadres | NON CADRES | NOMBRE >=0 | OUI | — |
| D | total | TOTAL | NOMBRE =B+C | OUI | **idx 5 = col F** lu par parser : conserver position |
| E | deploiement_qhse | Déploiement Personnel QHSE | TEXTE | NON | 07 agents QHSE |

**Contrainte parser stricte** : `effectifs[c]` lu en `r[5]` (idx 5 = col F). La colonne TOTAL **doit rester en colonne F** (donc structure ABCDEF avec colonnes A vide ou B vide pour décaler). Solution : insérer colonnes spacers ou ajuster header en accord avec parser. **Recommandation : modifier le parser** (plus simple — voir Impact global).

### Onglet `QHSE_Equipements`
| # | Code | En-tête | Type | Obligatoire | Exemple |
|---|------|---------|------|-------------|---------|
| A | id_equip | ID Équipement | LISTE Master.Equipements | OUI | CE-16T-01 |
| B | type_equip | Type | TEXTE (auto) | OUI | CHARIOT ELEVATEUR |
| C | tonnage | Tonnage | LISTE | OUI | 16T |
| D | nb_disponible | Nb Disponible | NOMBRE >=0 | OUI | 3 |
| E | status_cert | Status Certif./Étalonn. | LISTE (Certifié / Étalonné / En cours / À réaliser) | OUI | Certifié |
| F | date_prochain_ctrl | Prochain contrôle | DATE | NON | 2026-09-30 |
| G | affectation_client | Affectation client | TEXTE | NON | FOXTROT FORAGE: activités |
| H | observations | Observations | TEXTE | NON | — |

### Onglet `QHSE_Incidents`
| # | Code | En-tête | Type | Obligatoire | Exemple |
|---|------|---------|------|-------------|---------|
| A | categorie | Catégorie | LISTE | OUI | Accident avec arrêt / Accident sans arrêt / Incident / Presque accident / Premiers soins / Maladie / Déces / Postes aménagés / Situation dangereuse / Déversement / Départ de feu / Evasan jour / Evasan nuit / Malaria/Paludisme / Crise d'ulcère / Pansement / Divers |
| B | nombre | Nombre | NOMBRE >=0 | OUI | 0 |
| C | description | Description / Contexte | TEXTE | NON | — |
| D | action_correctrice | Action correctrice | TEXTE | NON | — |
| E | responsable | Responsable suivi | TEXTE | NON | — |
| F | jours_arret | Nombre jours d'arrêt | NOMBRE | NON | 0 |
| G | jours_sans_accident | Nombre jours sans accident (cumul) | NOMBRE | OUI | 55 |
| H | record_jsa | Record N jours sans accident | NOMBRE | NON | 120 |

**Note parser critique** : le parser cherche `jours sans acc` (L1535) puis lit la valeur dans la même ligne col +1 ou ligne +1. On garde un libellé contenant "jours sans acc" en col A, valeur en col B → compatible. Le JSA sera lu comme `kpi['C:Nombre jours sans accident (cumul)']`.

### Onglet `QHSE_Securite` (toolbox / permis / formations)
| # | Code | En-tête | Type | Obligatoire | Exemple |
|---|------|---------|------|-------------|---------|
| A | activite | Activité Q-HSE | LISTE | OUI | Toolbox Meeting / Briefing HSE / Induction HSE / Permis de travail / SET CARD / JSA / Audit CIAPOL / Check-list Camion / Check-list Inspection Camion & Engins / Check-list Opérations Navire / Test d'alcoolémie / Reporting / Coaching Permis à Points / Prévention Risques Routiers / Formation Grutier / Formation Cariste |
| B | nombre | Nombre | NOMBRE >=0 | OUI | 14 |
| C | duree_heures | Durée (heures) | NOMBRE | NON | 2 |
| D | participants | Participants | NOMBRE | NON | 35 |
| E | animateur | Animateur | TEXTE | NON | BLE OLIVIER |
| F | observations | Observations | TEXTE | NON | — |

**Note parser** : la colonne `j` en col J (idx 9) avec valeur numérique en col L (idx 11) est lue par L1531-32 → `kpi['J:'+activité]=valeur`. On préserve cette position en plaçant en réalité `activite` (col A) ET un miroir col J via formule `=A2`, valeur en col L via `=B2`. **Recommandation : modifier le parser** plutôt que dupliquer (voir Impact global).

### Onglet `QHSE_Risques` (matrice causes par catégorie)
| # | Code | En-tête | Type | Obligatoire | Exemple |
|---|------|---------|------|-------------|---------|
| A | famille_risque | Famille de risque | LISTE (Mauvaise performance / Chute ou projection / Équipement sous pression / Travail en hauteur / Opération de levage / Communication / EPI / Accessoires de levage / Housekeeping / Outils/Équipement / Matières dangereuses / Procédures & Documents / Équipement électrique / Zone de travail / Instruments de mesure / Manutention manuelle/Mécanisée) | OUI | Travail en hauteur |
| B | nombre | Nombre occurrences | NOMBRE >=0 | OUI | 0 |
| C | gravite | Gravité | LISTE (Faible / Moyenne / Grave / Critique) | NON | Faible |
| D | action_immediate | Action immédiate | TEXTE | NON | — |

### Onglet `QHSE_Actions`
| # | Code | En-tête | Type | Obligatoire | Exemple |
|---|------|---------|------|-------------|---------|
| A | action | Action à valoriser | TEXTE | OUI | Réaliser un TBM avant début de chantier |
| B | responsable | Responsable | TEXTE | OUI | Chef de chantier |
| C | date_prevue | Date prévue | DATE | OUI | 2026-05-25 |
| D | statut | Statut | LISTE (Planifiée / En cours / Réalisée / En retard) | OUI | Planifiée |
| E | priorite | Priorité | LISTE (Basse / Moyenne / Haute / Critique) | OUI | Moyenne |

### Mode d'emploi
> Le QHSE Manager remplit en début de semaine `Entete` puis tous les sous-onglets selon les observations. **Garder les libellés exacts** dans les colonnes "Catégorie" (listes déroulantes) — ils servent aux KPI mensuels. Le `Nombre jours sans accident` cumul est obligatoire (J0 si reset).

### Champs récupérés vs existant
- ✅ Effectifs AGL/Sous-traitant/Journaliers (cadres/non cadres/total)
- ✅ Jours sans accident
- 🆕 **MÉTÉO complète** (Ensoleillé/Nuageux/Venteux/Pluie OUI/NON) — orphelin
- 🆕 **Équipements détaillés** (Type/Tonnage/Nb/Status/Prochain contrôle/Client affectation)
- 🆕 **Incidents typés** (16 catégories, action correctrice, responsable, jours arrêt, record JSA)
- 🆕 **Activités sécurité** (16 activités structurées : toolbox, permis, JSA, audits, formations) — colonne J/L exploitée
- 🆕 **Familles de risques** (16 familles + gravité) — sections L146-163 du fichier original
- 🆕 **Actions à valoriser** avec métadonnées
- 🆕 Rédacteur, période, référence document

### Impact parser
- `parseQHSE()` (L1521) : refactor en **multi-onglet** :
  - Détecter onglet par nom (`/qhse_meteo/i`, `/qhse_personnel/i`, etc.)
  - `parseQHSEMeteo`, `parseQHSEPersonnel`, `parseQHSEEquipements`, `parseQHSEIncidents`, `parseQHSESecurite`, `parseQHSERisques`, `parseQHSEActions` (7 sous-fonctions).
  - Sortie consolidée : `weeks[i] = {semaine, meteo:[], effectifs:{}, equipements:[], incidents:[], securite:[], risques:[], actions:[], jsa, totalEff, kpi}`.
- Compat ascendante : si fichier ancien chargé, fallback sur algo actuel.
- Render : nouvel onglet "QHSE détaillé" avec tableaux par catégorie + courbe JSA.

---

## 9. Template ACTIONS (`08_ACTIONS_v1.xlsx`)

### Identification
- **Rythme** : hebdomadaire (1 onglet par semaine `DD-MM-YY`)
- **Responsable** : Ops + QHSE
- **Justification choix (consigne 8)** : enrichir les métadonnées (statut, responsable, dates, priorité), décomposer explicitement les sections I-V.

### Onglets
1. `📖 Mode d'emploi`
2. `2026-S20`, `2026-S21`, … (1 par semaine, nom = ISO week)
3. `_SYS`

### Onglet hebdo — Colonnes

| # | Code | En-tête | Type | Validation | Obligatoire | Source | Exemple | Note parser |
|---|------|---------|------|------------|-------------|--------|---------|-------------|
| A | num | N° | NOMBRE | >0 | OUI | Saisie | 1 | — |
| B | section | Section | LISTE | OUI | OUI | Référentiel | I. Actions menées Semaine S-1 | `cur` (L1619) — **doit commencer par `I.` / `II.` / `III.` / `IV.` / `V.`** |
| C | action | Action | TEXTE | longueur 5-500 | OUI | Saisie (formatée `Act N : <description>`) | Act 1 : entretien des yards FOXTROT | `b.split(':')` (L1620) — **doit contenir `:` séparateur** |
| D | statut | Statut | LISTE | Ouverte / En cours / Réalisée / Fermée / Annulée / Bloquée | OUI | Saisie | En cours | 🆕 **orphelin** |
| E | responsable | Responsable | TEXTE | — | OUI | Saisie | KOUASSI A. | 🆕 |
| F | date_prevue | Date prévue | DATE | YYYY-MM-DD | OUI | Saisie | 2026-05-25 | 🆕 |
| G | date_realisee | Date réalisée | DATE | YYYY-MM-DD | NON | Saisie | 2026-05-23 | 🆕 |
| H | priorite | Priorité | LISTE | Basse / Moyenne / Haute / Critique | OUI | Saisie | Haute | 🆕 |
| I | lien_incident | Lien Incident QHSE | TEXTE | Référence sheet/ligne QHSE | NON | Saisie | QHSE-S20-INC-03 | 🆕 traçabilité QHSE |
| J | progres_pct | Progrès (%) | NOMBRE | 0-100 | NON | Saisie | 80 | 🆕 |
| K | commentaire | Commentaire | TEXTE | — | NON | Saisie | Reporté faute matériel | 🆕 |

### Sections normalisées (col B — liste déroulante stricte)
- `I. Actions menées Semaine S-1`
- `II. Actions à mener Semaine en cours`
- `III. Difficultés / Dysfonctionnements / Incidents`
- `IV. Apprentissages / Bonnes pratiques`
- `V. Points d'attention / Risques émergents`

### Mode d'emploi
> Une ligne = une action. Sélectionner la **Section** (I à V), saisir l'action sous le format `Act N : description`. Renseigner Statut/Responsable/Date prévue obligatoirement. À la clôture, remplir Date réalisée + passer Statut à "Réalisée" ou "Fermée". Lier aux incidents QHSE via la référence `QHSE-S<n>-INC-<m>` quand applicable.

### Champs récupérés vs existant
- ✅ Section (I-V), Action (texte)
- 🆕 **Statut, Responsable, Date prévue, Date réalisée, Priorité** (5 colonnes de métadonnées)
- 🆕 **Lien Incident QHSE** (cross-reference)
- 🆕 Progrès %, commentaire
- 🆕 **Sections IV et V** explicitées (étaient potentiellement absentes du sample audit)

### Impact parser
- `parseActions()` (L1612) : conserver lecture col A/B existante (rétro-compat) ; ajouter `iStatut`, `iResp`, `iDatePrev`, `iDateReal`, `iPriorite`, `iIncident`.
- Étendre `sections[cur]` en `sections[cur] = [{action, statut, resp, datePrev, dateReal, priorite, lienIncident}]`.
- Render : Kanban (Ouverte/En cours/Réalisée), graphe priorité × statut, alertes "En retard" (date_prevue < today et statut != Réalisée).

---

## 10. Impact global sur l'app `O_G_AGL_Analyseur_OFFLINE_v23.html`

### Parsers à modifier
| Fonction (ligne) | Modifications |
|------------------|---------------|
| `parseLoyers` (L1650) | Aucun changement obligatoire ; ajouter optionnellement `byEspace` si col B `id_espace` présente. |
| `parseCharges` (L1628) | Ajouter `iUM`, `iStatut`, `iPaiement` ; agrégats `byStatut`, `byPaiement`, `byUM`. |
| `parsePlanning` (L1582) | Ajouter `iDemande`, `iObs`, `iEquip` ; `ops[].demande/obs/equip`. |
| `parsePointage` (L1554) | Ajouter `iProduit`, `iHeures`, `iValidation` ; `workers[].produit/heures/validation`. Détecter onglet `SOURCE_Personnel` (skip de l'agrégation). |
| `parseCA` (L1693) | Ouvert aux nouvelles colonnes BASE LOYER / Plog & Divers / IMAR / HINT / etc. (déjà supporté via `depCols`). Vue 3 conservée (positions hardcodées 16/17/20). |
| **NEW** `parseFinance` | Nouveau parser pour `05B.PnL_Base` : agréger par poste × mois prévu/réel. |
| `parseAccostage` (L1793) | Ajouter findIndex pour blocs GOTTWALD (`/gottwald.*date`, `/gottwald.*mob/`, etc.) et MAFI ; ajouter `byQuai` (remplace `quaiSet`). |
| `parseQHSE` (L1521) | **Refactor majeur** : dispatcher par nom d'onglet (`/qhse_meteo|qhse_personnel|...`). 7 sous-parsers. Fallback ancien algo si onglet historique détecté. |
| `parseActions` (L1612) | Lire colonnes D-K supplémentaires ; transformer `sections[cur]` en tableau d'objets. |

### Render functions à enrichir
- Ajouter cartes KPI : Activité GOTTWALD, Activité MAFI, Occupation par quai (Q25/Q21)
- Ajouter onglet "Compte d'exploitation" (P&L Prévu vs Réel)
- Ajouter Kanban Actions
- Ajouter table Charges par statut/mode paiement
- Ajouter graphes QHSE détaillés (météo, équipements en attente certif, top familles de risques)

### Référence aux Master Référentiels
- Ajouter `<input type="file" id="masterFile">` dans la zone d'upload
- Fonction `loadMaster(wb)` au chargement → `CONTEXTE_AGL.masters = {clients, espaces, equipements, sections, natures, affectations, fournisseurs, departementsCA, postesPnL, quais}`
- Utilisation : canonisation noms (étendre `canonClient` L1822), dropdowns futurs, validation.

### Compat ascendante
- Tous les nouveaux champs sont **optionnels côté parser** : si absents (fichier ancien), le rendu fallback sur l'existant.
- Détection version via cellule `_SYS!A1` du fichier source.

---

## 11. Checklist de génération (script Python)

Pour chaque fichier ci-dessous : générer `.xlsx` avec `openpyxl`, en appliquant : police Calibri, en-têtes navy/gold ligne 1 verrouillée, validations (Data Validation), plages nommées, listes déroulantes pointant vers `_SYS`/Master, format dates/heures/nombres, cellules figées (ligne 1 + col A), protection feuille avec mot de passe, onglet `📖 Mode d'emploi` en position 1.

- [ ] `00_MASTER_REFERENTIELS.xlsx` (9 onglets + Mode d'emploi + _SYS)
- [ ] `01_LOYERS_v1.xlsx` (1 onglet données + Mode d'emploi)
- [ ] `02_CHARGES_v1.xlsx` (12 onglets mensuels harmonisés + Mode d'emploi)
- [ ] `03_PLANNING_v1.xlsx` (modèles JOUR + NUIT identiques + Mode d'emploi)
- [ ] `04_POINTAGE_v1.xlsx` (SOURCE_Personnel intégré + modèles JOUR/NUIT/WK + Mode d'emploi)
- [ ] `05A_CA_FLASH_v1.xlsx` (Récap Global + Cumul Annuel + Mode d'emploi)
- [ ] `05B_CA_PNL_v1.xlsx` (PnL_Base format long + Mode d'emploi)
- [ ] `05C_CA_DETAIL_v1.xlsx` (12 onglets mensuels avec positions colonnes Q/R/U respectées + Mode d'emploi)
- [ ] `06_ACCOSTAGE_v1.xlsx` (12 onglets mensuels avec blocs GOTTWALD+MAFI + Quai unique + Mode d'emploi)
- [ ] `07_QHSE_v1.xlsx` (Entête + 6 sous-onglets thématiques + Mode d'emploi)
- [ ] `08_ACTIONS_v1.xlsx` (modèles hebdo avec métadonnées + Mode d'emploi)

**Total** : 10 fichiers (+ archives séparées CA pré-2026 non régénérées).

### Synthèse orphelins traités (100 % du top 10 + reste)
| Rang | Sujet | Traité dans | Statut |
|------|-------|-------------|--------|
| 1 | ACCOSTAGE blocs GOTTWALD + MAFI | §7 cols J-R | ✅ |
| 2 | CA Gestion base logistique | §6B `05B_CA_PNL` | ✅ |
| 3 | PLANNING DEMANDE EMISE | §4 col D | ✅ |
| 4 | CHARGES Mois/Statut/Mode Paiement | §3 cols F/G/H | ✅ |
| 5 | ACCOSTAGE QUAI 25 vs 21 | §7 col H + Master.Quais | ✅ |
| 6 | CHARGES N° UM | §3 col B | ✅ |
| 7 | POINTAGE Produits + validation N-Q | §5 cols H/K/L (+ SOURCE_Personnel) | ✅ |
| 8 | QHSE pleine couverture | §8 (6 sous-onglets, 27 sections préservées) | ✅ |
| 9 | CA onglets orphelins | §6 (Récap Global enrichi `byDep`, Occupation archivée) | ✅ (partiel : archive volontaire) |
| 10 | LOYERS Feuil1 espaces | §1 Master.Espaces + §2 col B `id_espace` | ✅ |
| + | CHARGES Date facture/N° Facture | §3 cols I/J | ✅ bonus |
| + | PLANNING Observations + Équipement | §4 cols H/I | ✅ bonus |
| + | POINTAGE Heures travaillées | §5 col J | ✅ bonus (cross QHSE) |
| + | ACCOSTAGE NBRE HEURES | §7 col G conservé (validation ETA/ETD) | ✅ |
| + | QHSE Météo (4 indicateurs) | §8 `QHSE_Meteo` | ✅ |
| + | QHSE Équipements (status certif) | §8 `QHSE_Equipements` | ✅ |
| + | QHSE Activités sécurité (col J/L) | §8 `QHSE_Securite` | ✅ |
| + | ACTIONS métadonnées (statut/resp/date/prio) | §9 cols D-H | ✅ |
| + | ACTIONS sections IV-V | §9 liste B normalisée | ✅ |
| + | ACTIONS lien incident QHSE | §9 col I | ✅ traçabilité retrouvée |

---

**Fin DESIGN.md** — prêt pour validation puis passage à la génération `.xlsx` (Phase 3).
