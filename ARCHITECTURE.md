# Segmentation client - Document d'architecture & procedure

> Version 1 - Mode manuel
> Cible : utilisateur final fonctionnel (non-developpeur)

---

## 1. Contexte et objectif

### 1.1 Probleme metier

L'entreprise dispose d'une base CRM (Microsoft Dynamics ou outil
d'extraction tiers) contenant l'historique des operations logistiques
client : ~450 000 lignes a date, en croissance continue, sur la periode
2021 -> 2026.

Aujourd'hui, il n'existe **pas de classification automatique** des
clients par segment commercial. La consolidation et le tri se font
manuellement dans Excel, ce qui est :

- chronophage (plusieurs heures par exercice),
- source d'erreurs (formules figees, copier-coller),
- non reproductible (chaque analyse repart de zero),
- non partageable (chaque analyste a sa version).

### 1.2 Objectif

Mettre en place une **application web locale** qui :

1. Recoit en entree l'export Excel brut du CRM
2. Permet a l'utilisateur de **definir des regles de segmentation**
   (ex : VIP = CA > 500 k EUR)
3. **Classe automatiquement** chaque client dans son segment
4. Restitue un tableau Excel pret a l'emploi

### 1.3 Perimetre version 1 (la "v1 manuelle")

| Inclus | Hors perimetre v1 |
|---|---|
| Upload de fichiers Excel | Connexion directe a Dynamics |
| Definition des regles dans l'UI | Planification automatique |
| Classification + export | Historisation des segments |
| Multi-fichiers + dedup | Gestion multi-utilisateurs |
| Filtres periode | Authentification |

---

## 2. Architecture technique

### 2.1 Stack

| Composant | Choix | Pourquoi |
|---|---|---|
| Langage | **Python 3.10+** | Standard data, riche en bibliotheques |
| Framework UI | **Streamlit** | Transforme un script Python en site web sans HTML/CSS |
| Traitement donnees | **pandas** | Standard du metier, gere 5M lignes en RAM sans souci |
| Lecture Excel | **openpyxl** | Le plus fiable pour les .xlsx Siebel/Dynamics |
| Ecriture Excel | **xlsxwriter** | Rapide, mise en forme propre |
| Stockage | **Aucun** (in-memory) | Pas besoin de base de donnees pour la v1 |
| Hebergement | **Local** (poste utilisateur) | Pas d'IT, pas de cloud, controle total |

### 2.2 Composants

```
+------------------------------------------------------------------+
|                    Poste utilisateur (Windows / Mac)             |
|                                                                  |
|  +----------------+                                              |
|  |  Navigateur    |  <-- localhost:8501                          |
|  |  (Chrome,...)  |                                              |
|  +-------+--------+                                              |
|          |                                                       |
|          | HTTP local                                            |
|          v                                                       |
|  +----------------------------------------------------+          |
|  |          Processus Streamlit (Python)              |          |
|  |  +----------------+   +-----------------------+    |          |
|  |  |  Module UI     |   |  Moteur de calcul     |    |          |
|  |  |  (app.py)      |-->|  (pandas)             |    |          |
|  |  +----------------+   +-----------------------+    |          |
|  |  Lit/ecrit les fichiers du poste                   |          |
|  +----------------------+-----------------------------+          |
|                         |                                        |
|                         v                                        |
|  +----------------------------------------------------+          |
|  |  Dossier de travail (ex: Bureau/CRM-segmentation/) |          |
|  |  +-------------+  +-------------+  +------------+  |          |
|  |  | 2021.xlsx   |  | 2022.xlsx   |  | segments.  |  |          |
|  |  |             |  |             |  | json       |  |          |
|  |  +-------------+  +-------------+  +------------+  |          |
|  +----------------------------------------------------+          |
+------------------------------------------------------------------+
```

L'application **ne contacte aucun service externe**. Tout reste sur le
poste de l'utilisateur. Les donnees CRM ne quittent jamais le PC.

### 2.3 Schema des donnees

#### Donnees en entree (Excel brut Siebel/Dynamics)

Granularite : **1 ligne = 1 operation logistique** (JobfileNumber).

Colonnes cles utilisees par l'app :

| Colonne | Type | Usage dans l'app |
|---|---|---|
| `Year`, `Month` | entier | Filtres periode |
| `Customer` | texte | Identifiant client (cle d'agregation) |
| `HQ CTO Customer Name` | texte | Niveau groupe (alternative comme cle) |
| `Reporting Entity Label` | texte | Filtre dimension |
| `Zone 3 Label` | texte | Filtre dimension |
| `Market Segment Name` | texte | Filtre dimension |
| `Activity Name (EN)` | texte | Filtre dimension |
| `Product Name` | texte | Filtre dimension |
| `JobfileNumber` | texte | Identifiant operation (dedup) |
| `Turnover in EUR` | numerique | Metrique : CA |
| `Direct GM in EUR` | numerique | Metrique : marge brute |
| `TEU` | numerique | Metrique : volume conteneurs |
| `Freight Ton` | numerique | Metrique : tonnage |

#### Donnees agregees (calculees par l'app)

Granularite : **1 ligne = 1 client**.

| Colonne | Calcul |
|---|---|
| `Customer` | cle |
| `CA_EUR` | somme(`Turnover in EUR`) |
| `GM_EUR` | somme(`Direct GM in EUR`) |
| `TEU` | somme(`TEU`) |
| `Freight_Ton` | somme(`Freight Ton`) |
| `Nb_operations` | nombre distinct de `JobfileNumber` |
| `Premiere_activite` | min(`Period`) |
| `Derniere_activite` | max(`Period`) |
| `Nb_mois_actifs` | nombre de mois distincts d'activite |
| `segment` | resultat de la classification |

#### Configuration des segments (JSON)

```json
[
  {
    "name": "VIP",
    "logic": "AND",
    "conditions": [
      { "column": "CA_EUR", "op": ">=", "value": "500000" }
    ]
  },
  {
    "name": "Gold",
    "logic": "AND",
    "conditions": [
      { "column": "CA_EUR", "op": ">=", "value": "100000" },
      { "column": "Nb_operations", "op": ">=", "value": "10" }
    ]
  }
]
```

Operateurs disponibles : `==`, `!=`, `>`, `>=`, `<`, `<=`, `between`,
`contient`, `commence par`, `finit par`, `dans la liste`, `regex`,
`est vide`, `n'est pas vide`.

Logique entre conditions d'un meme segment : `AND` ou `OR`.

Logique entre segments : **premier segment qui matche gagne** (ordre
important).

### 2.4 Flow de bout en bout

```
[Export Siebel/Dynamics .xlsx] (jusqu'a 3-6 fichiers a cause de la limite 100-150k lignes)
              |
              v
[Upload dans l'app Streamlit]
              |
              v
[Detection auto de la ligne d'en-tete + concat + dedup par JobfileNumber]
              |
              v
[Filtres : annees, mois, entites, zones, marche, activite, ...]
              |
              v
[Agregation : groupby(Customer) -> CA, GM, TEU, nb operations, ...]
              |
              v
[Application des regles de segments dans l'ordre]
              |
              v
[Tableau final avec colonne "segment" + repartition]
              |
              v
[Export Excel (2 onglets : clients + repartition) ou CSV]
```

---

## 3. Procedure utilisateur

### 3.1 Mise en place initiale (1 seule fois, ~30 min)

**Pour la personne qui installe** :

1. Installer **Python 3.10 ou +** (https://www.python.org/downloads/)
2. Recuperer le repo
   ```
   git clone <url-du-repo>
   cd segmentation-client
   ```
3. Installer les dependances
   ```
   pip install -r requirements.txt
   ```
4. Verifier que ca demarre
   ```
   streamlit run app.py
   ```
   Le navigateur s'ouvre sur http://localhost:8501.
5. Creer un dossier de travail sur le bureau, par exemple
   `Bureau/CRM-segmentation/`
6. (Optionnel) Creer un raccourci de lancement :
   - **Windows** : un fichier `lancer.bat` sur le bureau :
     ```bat
     cd C:\chemin\vers\segmentation-client
     streamlit run app.py
     ```
   - **Mac** : un fichier `lancer.command` :
     ```bash
     cd /chemin/vers/segmentation-client
     streamlit run app.py
     ```
   Double-clic = l'app demarre.

### 3.2 Routine de travail (a chaque segmentation)

**Etape 1 - Extraire les donnees du CRM** (5-15 min)

1. Se connecter au CRM
2. Lancer la requete / vue qui contient les operations
3. Exporter en Excel
4. Si on depasse la limite d'export (100k Dynamics / 150k Siebel) :
   - Decouper par periode (ex : 1 export par annee)
   - Verifier que les fichiers se chevauchent le moins possible
5. Deposer les fichiers `.xlsx` dans `Bureau/CRM-segmentation/`
6. **Convention de nommage recommandee** :
   `extract_AAAAMMJJ_AnneeDebut-AnneeFin.xlsx`
   exemple : `extract_20260511_2021-2026.xlsx`

**Etape 2 - Charger les donnees dans l'app** (1 min)

1. Lancer l'app (double-clic sur le raccourci)
2. Etape 1 de l'app : glisser-deposer **tous** les fichiers Excel
3. Verifier les chiffres affiches :
   - "Lignes totales" doit etre coherent (~450 000 actuellement)
   - "Doublons supprimes" : normal si plusieurs exports se chevauchent

**Etape 3 - Filtrer la periode** (1 min)

1. Etape 2 de l'app : selectionner les annees voulues
2. Ajuster le slider mois si besoin
3. (Optionnel) Filtres dimensions : entite, zone, segment marche...

**Etape 4 - Agreger** (1 min)

1. Etape 3 de l'app : choisir la cle client (`Customer` par defaut)
2. Choisir les metriques (CA, GM, TEU, etc.)
3. Cliquer "Lancer l'agregation"
4. L'app affiche le nombre de clients uniques et un apercu

**Etape 5 - Definir / charger les regles de segments** (5 min la 1ere fois, 30 sec ensuite)

**Premiere fois :**
1. Etape 4 de l'app : ajouter chaque segment un par un
2. Pour chaque segment : conditions + logique AND/OR
3. **Verifier l'ordre** (premier match gagne)
4. Telecharger le JSON pour reutiliser plus tard

**Fois suivantes :**
1. Charger le `segments.json` sauvegarde precedemment

**Etape 6 - Classifier et exporter** (1 min)

1. Etape 5 de l'app : cliquer "Lancer la classification"
2. Verifier la repartition affichee (tableau + graphique)
3. Telecharger le fichier Excel final
4. Stocker le resultat dans `Bureau/CRM-segmentation/resultats/`

**Duree totale d'une segmentation : ~15 min** (vs plusieurs heures
en Excel manuel).

### 3.3 Cas particuliers

#### Comparer deux periodes

1. Faire la procedure complete une 1ere fois avec annees 2024-2025
   -> exporter `clients_2024_2025.xlsx`
2. Refaire avec annees 2025-2026 -> `clients_2025_2026.xlsx`
3. Joindre les deux fichiers dans Excel (RECHERCHEV ou Power Query)

#### Tester plusieurs jeux de regles

1. Sauvegarder plusieurs JSON : `segments_v1.json`, `segments_v2.json`
2. Charger chacun a tour de role et comparer les exports

---

## 4. Preconisations

### 4.1 Organisation des fichiers

```
Bureau/
└── CRM-segmentation/
    ├── extracts/                    <-- les Excel bruts du CRM
    │   ├── extract_20260111_2021-2023.xlsx
    │   ├── extract_20260111_2024-2026.xlsx
    │   └── ...
    ├── configs/                     <-- les regles de segmentation
    │   ├── segments_commercial_v1.json
    │   ├── segments_marketing_v1.json
    │   └── ...
    └── resultats/                   <-- les exports de l'app
        ├── 2026-05-11_clients_segmentes.xlsx
        └── ...
```

### 4.2 Bonnes pratiques sur les regles de segments

| Regle | Pourquoi |
|---|---|
| Toujours mettre les segments **les plus restrictifs en haut** | Premier match gagne |
| Toujours avoir un segment "fourre-tout" en dernier | Eviter trop de "Non classe" |
| Garder **3 a 7 segments** maximum | Au-dela, devient ingerable |
| Documenter le sens metier de chaque segment | A noter dans le nom : "VIP_CA>500k" |
| Conserver les versions de regles (versioning JSON) | Tracabilite |
| Tester les regles sur un mois recent | Verifier la repartition avant generaliser |

### 4.3 Sauvegardes

- Le **dossier de travail** sur le bureau doit etre synchronise
  (OneDrive / Dropbox / sauvegarde reseau).
- Les **fichiers de regles JSON** doivent etre versionnes (idealement
  dans Git, sinon dans un dossier "configs" date).
- Les **exports Excel** doivent etre dates dans le nom (`2026-05-11_...`).

### 4.4 Performances

| Volume | Action | Temps observe |
|---|---|---|
| 450k lignes | Chargement multi-fichiers | 30-60 s |
| 450k lignes | Agregation | 2-5 s |
| 450k lignes | Classification | 1-2 s |
| 2M lignes | Chargement | 2-4 min |
| 2M lignes | Agregation | 10-20 s |

Au-dela de 2M lignes, prevoir le passage en mode "data store + parquet"
(voir section 5).

---

## 5. Evolutions possibles (roadmap)

### 5.1 Court terme (1-3 mois) - "Confort utilisateur"

| Evolution | Effort | Benefice |
|---|---|---|
| Dossier surveille (l'app lit auto le dossier, plus d'upload) | 1 j | Gain de temps |
| Cache Parquet (chargement 10x plus rapide a partir de la 2eme fois) | 1 j | Confort |
| Templates de segments (commercial, marketing, gestion...) | 1 j | Standardisation |
| Comparaison de 2 periodes cote a cote dans l'app | 2 j | Analyses temporelles |
| Graphiques de distribution par segment (CA, anciennete...) | 1 j | Communication |

### 5.2 Moyen terme (3-12 mois) - "Industrialisation"

| Evolution | Effort | Prerequis |
|---|---|---|
| Deploiement sur serveur interne (multi-utilisateurs) | 1 sem | IT |
| Authentification SSO | 3 j | IT, AD/Azure AD |
| Historisation des classifications (qui etait VIP en jan 2026 ?) | 1 sem | BDD SQLite ou Postgres |
| Connexion **Power Automate -> OneDrive** : extract auto quotidien | 2 j | Licence M365 |
| Alertes : "le client X est descendu de VIP a Gold ce mois" | 3 j | Historisation |

### 5.3 Long terme (>1 an) - "Plateforme"

| Evolution | Effort | Prerequis |
|---|---|---|
| Connexion directe Dynamics via Dataverse Web API | 2 sem | Acces API + token OAuth |
| Bascule moteur sur DuckDB (gere >50M lignes) | 1 sem | Croissance volume |
| Scoring predictif (probabilite de devenir VIP) | 1 mois | Data scientist |
| Integration dans Power BI / dashboard executif | 2 sem | Power BI deploye |
| Publication d'API (autres outils consomment les segments) | 2 sem | IT |

---

## 6. Freins et risques

### 6.1 Freins techniques

| Frein | Impact | Mitigation |
|---|---|---|
| **Limite d'export CRM** (100-150k lignes) | Moyen | Multi-fichiers + dedup (deja gere v1) |
| **Tailles de fichier > 100 Mo** dans Streamlit upload | Moyen | Passage en mode "dossier surveille" |
| **Lenteur Excel sur > 2M lignes** | Faible aujourd'hui, fort dans 5 ans | Cache Parquet + DuckDB |
| Python installe sur le poste | Faible | Tutoriel pas a pas, ou empaqueter en .exe (PyInstaller) |
| Mise a jour de l'app | Faible | Git pull, ou packager les releases |

### 6.2 Freins organisationnels

| Frein | Impact | Mitigation |
|---|---|---|
| Changement de format de l'export CRM | **Fort** | Detection auto colonnes + alerte si manquantes |
| Plusieurs utilisateurs avec des regles differentes | Moyen | Versioning des JSON + revue mensuelle |
| Pas de standard "qu'est-ce qu'un segment VIP ?" | **Fort** | Atelier metier prealable, definition documentee |
| Resistance au changement (habitude Excel) | Moyen | Formation 30 min + double-run en parallele 1 mois |
| Continuite si la personne installatrice part | Moyen | Procedure documentee (ce document) + sauvegarde repo |

### 6.3 Risques securite / RGPD

| Risque | Niveau | Mitigation |
|---|---|---|
| Donnees client en clair sur le poste | Moyen | Disque chiffre (BitLocker / FileVault) |
| Fuite via OneDrive partage par erreur | Moyen | Dossier non partage par defaut |
| Pas de log des consultations | Faible (v1 mono-utilisateur) | Sera ajoute en v2 multi-utilisateurs |
| Donnees commerciales sensibles | Fort | Pas d'envoi cloud, traitement 100% local en v1 |
| Conformite RGPD (donnees clients) | A verifier | Pas de donnees nominatives sensibles aujourd'hui ; revue DPO recommandee |

### 6.4 Limites fonctionnelles assumees v1

- **Pas d'historique** : on ne sait pas qu'un client etait Gold le mois
  dernier. Solution : sauvegarder les exports dates dans `resultats/`.
- **Pas de collaboration** : une personne a la fois. Solution : aligner
  les regles via JSON partage.
- **Pas de validation automatique** : si l'export CRM change de format,
  l'app peut echouer. Solution : tests sur extract de reference.
- **Pas de regles "compose"** : segment B = segment A + critere
  supplementaire. Solution v2.
- **Pas de prediction** : c'est de la regle, pas du machine learning.

---

## 7. Resume executif

**Aujourd'hui (v1 manuelle)** : une app Streamlit locale, gratuite,
qui charge des Excel CRM, applique des regles de segmentation
configurables, et exporte le resultat. Installation 30 min, prise en
main 15 min. Pas de dependance IT.

**Demain (v2 confort)** : suppression de l'upload manuel grace a un
dossier surveille + cache Parquet. Gain de temps notable.

**Apres-demain (v3 industrialisee)** : connexion directe au CRM via
Power Automate ou API Dataverse, historisation, multi-utilisateurs.

L'architecture choisie permet **d'avancer par etapes sans jamais tout
refondre** : chaque evolution s'ajoute sur la base existante.
