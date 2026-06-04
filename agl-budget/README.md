# AGL BUDGET — Strategic Intelligence Platform

> Système d'intelligence stratégique **offline** — standard cabinet de conseil —
> basé sur données 100 % propriétaires AGL + sources institutionnelles mondiales
> à date **juin 2026**.
> Direction Stratégie — AGL Côte d'Ivoire · v2.0 · **CONFIDENTIEL**

Ce dépôt-dossier (`agl-budget/`) implémente le cahier des charges *AGL Strategic
Intelligence Platform v2.0*. Il est **indépendant** de l'app de segmentation à la
racine du repo.

---

## 1. Ce que le projet produit

- Classifier 7 000+ clients par secteur automatiquement (Référentiel Maître Client).
- Calculer la part de marché AGL par métier STATCOM (Maritime, Aérien, Hinterland).
- Analyser la concurrence par métier (positionnement dynamique 3 ans).
- Mesurer l'écart budget PFA 2026 vs réalisé (RUBRIKS vs IRIS).
- Cartographier les projets PND 2026-2030.
- Produire des projections 2026-2030 (STATCOM 3 ans + macro à date).
- Exporter des rapports Word + PPT COMEX-ready par thématique.

### Les 14 onglets

| # | Onglet | Fonction |
|---|--------|----------|
| 01 | Matching & RMC | Construction du référentiel maître |
| 02 | Marché Global STATCOM | Vision marché total CIV — 3 ans |
| 03 | Position AGL | PDM AGL vs marché — 7 métiers |
| 04 | Budget PFA vs Réel | Suivi objectifs RUBRIKS vs IRIS |
| 05 | Pipeline CRM | Analyse pipeline commercial |
| 06 | Segmentation BCG | Matrice 9 cases + scoring client |
| 07 | White Spaces | Marché non capturé — priorités |
| 08 | Carte PND 2026-2030 | Cartographie interactive projets |
| 09 | Top 20 Projets PND | Fiches détaillées + massification |
| 10 | Projections 2026-2030 | Forecasting hybride — 3 scénarios |
| 11 | Préconisations Client | Fiches client stratégiques |
| 12 | Synthèse Executive | Note COMEX — format SCR |
| 13 ★ | Analyse Concurrentielle | 7 métiers — bubble chart |
| 14 ★ | Cohortes & Rétention | Clients 3 ans — LTV — churn |

---

## 2. Les 4 bases sources

| | CRM | IRIS | STATCOM | RUBRIKS |
|---|---|---|---|---|
| Rôle | Référentiel commercial | Performance opérat. AGL | Marché total CIV | Budget PFA 2026 |
| Clé | `ID_CRM` (principal) | `ID_IRIS` | Nom client | Nom client |
| Secteur | ✅ **LOCKED** | ❌ | ❌ | Partiel |
| Priorité RMC | **PRIORITÉ 1 — LOCKED** | PRIORITÉ 2 | Fallback marchandise | Enrichissement |

**7 métiers STATCOM** : Import Maritime, Export Maritime, Import Aérien,
Export Aérien, Hinterland Import, Hinterland Export, Vue Consolidée.

---

## 3. Architecture technique

Contrainte clé (§5) : **pas d'installation serveur**, tout tourne sur un PC
standard AGL ; volumes de 300 000 à 400 000+ lignes Excel.

- **Frontend offline** (`index.html` + `js/`) : upload séquentiel →
  `FileReader` → **Web Worker** (parsing SheetJS hors thread UI) → agrégation →
  **IndexedDB** (snapshot, reprise sans re-upload) → chunked loading (10 000 lignes).
- **Pipeline Python** (`python/`) : 9 scripts indépendants. Matching probabiliste
  **Splink** (Fellegi-Sunter) + **Cleanco** (formes juridiques OHADA).
- **Règle d'or** : *monolithe interdit* — 1 module = 1 fichier, débugable isolément.

```
agl-budget/
├── index.html              coquille 14 onglets
├── js/
│   ├── utils.js            formatage FCFA/TEU/Kg, normalisation, DOM
│   ├── macro_data.js       sources macro mondiales à date juin 2026 (§7)
│   ├── store.js            persistance IndexedDB (snapshots)
│   ├── charts.js           bar / line / bubble / donut (sans dépendance)
│   ├── worker.js           Web Worker — parsing Excel en arrière-plan
│   ├── loader.js           upload séquentiel + snapshot + détection mémoire
│   └── app.js              contrôleur UI (navigation, upload, aperçu, macro)
└── python/
    ├── 01_normalise.py     Cleanco + strip géo/OHADA custom CIV
    └── requirements.txt    pandas, openpyxl, cleanco, splink, rapidfuzz
```

Les 9 scripts du pipeline (cible) :

| Script | Fonction |
|--------|----------|
| `01_normalise.py` | Cleanco + strip géo/stopwords/OHADA custom CIV ✅ |
| `02_match_crm_iris.py` | Splink FLUX 1 : CRM × IRIS |
| `03_match_crm_stat.py` | Splink FLUX 2 : CRM × STATCOM (+ marchandises) |
| `04_match_iris_stat.py` | Splink FLUX 3 pont : IRIS × STATCOM |
| `05_build_rmc.py` | Construction RMC — fusion 3 flux — propagation secteur |
| `06_agreg_statcom.py` | Agrégats STATCOM 3 ans par métier × année |
| `07_calcul_pdm.py` | PDM AGL par métier × année — CAGR réels |
| `08_projections.py` | Forecasting 2026-2030 — STATCOM base primaire |
| `09_whitespaces.py` | Opportunités non capturées |

---

## 4. Lancer (Sprint 1)

### Frontend
L'app étant offline et chargeant un Web Worker, servez le dossier en local
(le `file://` direct bloque les workers sur certains navigateurs) :

```bash
cd agl-budget
python3 -m http.server 8000
# puis ouvrir http://localhost:8000
```

Onglet **01 · Matching & RMC** : chargez un Excel de test (≈ 500 lignes) dans
n'importe quelle carte source → parsing en arrière-plan, aperçu, snapshot
IndexedDB. Onglet **📈 Macro** : sources institutionnelles à date juin 2026.

### Pipeline Python
```bash
cd agl-budget/python
pip install -r requirements.txt
python 01_normalise.py --input CRM.xlsx --col NOM_CLIENT --preview
```

---

## 5. Plan d'exécution — 6 sprints (§8)

| Sprint | Livrables | État |
|--------|-----------|------|
| **1 — Fondations** | `loader/utils/charts/macro_data/store/worker/app.js`, coquille 14 onglets, `01_normalise.py` | ✅ **fait** |
| 2 — Matching | Scripts `02`-`05` (Splink 3 flux + RMC) + onglet 01 matching | à venir |
| 3 — Marché | Scripts `06`-`07` + onglets 02/03/04 | à venir |
| 4 — Concurrence | `08_projections.py` + onglets 13/10/07/05/06 | à venir |
| 5 — PND & Stratégie | onglets 08/09/11/12/14 | à venir |
| 6 — Exports | `export_word.js`, `export_ppt.js`, `prompt.js` | à venir |

---

## 6. Fichiers à fournir (bloquant pour Sprint 2+)

CRM.xlsx (3 ans) · IRIS.xlsx (3 ans) · STATCOM_2023/2024/2025.xlsx ·
RUBRIKS.xlsx · PND_2026_2030.pdf.

**À confirmer avant de coder le matching** — noms de colonnes exacts :
- CRM : `ID_CRM ? NOM_CLIENT ? SECTEUR ? COMMERCIAL ? PIPELINE ?`
- IRIS : `ID_IRIS ? NOM_CLIENT ? CAP ? PERIODE ? METIER ?`
- STATCOM : `NOM_CLIENT ? MARCHANDISE ? VOLUME_TEU ? VOLUME_BULK ? METIER ? ANNEE ?`
- RUBRIKS : `NOM_CLIENT ? CAP_PFA ? SECTEUR ? ANNEE ?`
