# AGL BUDGET — Guide utilisateur

> Pour la Direction Stratégie AGL Côte d'Ivoire — usage mensuel.
> Pas besoin d'être développeur. ~30 min de prise en main.

---

## 1. Installation (1 seule fois — ~20 min)

### 1.1 Pré-requis
- Python 3.10 ou +  →  https://www.python.org/downloads/
- `poppler-utils` (pour lire les PDF du PND) :
  - **macOS** : `brew install poppler`
  - **Windows** : télécharger depuis https://github.com/oschwartz10612/poppler-windows
  - **Linux** : `sudo apt install poppler-utils`
- Navigateur récent (Chrome, Edge, Firefox).

### 1.2 Récupérer le projet
```bash
git clone <url-du-repo>
cd segmentation-client/agl-budget
```

### 1.3 Installer les dépendances Python
```bash
pip install -r python/requirements.txt
```

### 1.4 Vérifier que tout fonctionne
```bash
cd python
python _make_fixtures.py        # crée des données de test
python run_all.py --data-dir fixtures --out fixtures/out
```
La dernière ligne doit afficher `[ok] PIPELINE COMPLET TERMINÉ`.

---

## 2. Routine mensuelle (~15 min)

### Étape 1 — Préparer les 4 bases (5-10 min)

Récupère les exports **du mois écoulé** depuis les SI AGL :

| Source | Format attendu | Nom de fichier |
|---|---|---|
| **CRM** (référentiel client) | Excel 3 ans avec colonnes `nom_du_compte`, `crm_id_compte`, `secteur` | `CRM.xlsx` |
| **IRIS** (réalisé AGL) | Excel 3 ans avec colonnes `client`, `id`, `cap`, `periode` | `IRIS.xlsx` |
| **STATCOM** (marché CIV) | Excel 3 ans avec colonnes `metier`, `annee`, `destinataires`, `chargeurs`, `marchandise`, `volume_teu` ou `volume_kg` | `STATCOM.xlsx` (consolidé) ou `STATCOM_2023.xlsx`, etc. |
| **RUBRIKS** (budget PFA) | Excel avec colonnes `client`, `secteur`, `cap_pfa`, `annee` | `RUBRIKS.xlsx` |
| **Pipeline CRM** | Export Power BI Dynamics — fichier `opportunites.xlsx` tel quel | `opportunites.xlsx` |
| **PND** (optionnel, 1× par an) | PDF Tome 3 "Matrices d'actions" | `PND_Tome3.pdf` |

Pose-les dans un dossier dédié, par exemple :
```
Bureau/AGL-BUDGET/data-2026-06/
  CRM.xlsx
  IRIS.xlsx
  STATCOM_2023.xlsx  STATCOM_2024.xlsx  STATCOM_2025.xlsx
  RUBRIKS.xlsx
  opportunites.xlsx
  PND_Tome3.pdf
```

> **⚠ Confidentiel** — ces fichiers ne quittent jamais ton poste. Toute l'app
> tourne en local. Pense à activer BitLocker/FileVault sur ton disque.

### Étape 2 — Lancer le pipeline Python (~30 s)

```bash
cd segmentation-client/agl-budget/python
python run_all.py --data-dir "C:\Users\toi\Bureau\AGL-BUDGET\data-2026-06" \
                  --out      "C:\Users\toi\Bureau\AGL-BUDGET\out-2026-06"
```

Le script enchaîne **11 étapes** :
1. Normalisation des noms (Cleanco + OHADA + aliases)
2-4. Matching probabiliste CRM × IRIS × STATCOM
5. Construction du **Référentiel Maître Client** (RMC) 15 champs
6. Agrégats STATCOM 3 ans (volumes, nb clients, CAGR)
7. Calcul de la **PDM AGL** par métier × année
8. Projections 2026-2030 (3 scénarios)
9. Détection des white spaces (clients non capturés)
11. Budget PFA vs Réel (RUBRIKS × IRIS via RMC)

Pour extraire le PND (1× par an au maximum) :
```bash
python 10_extract_pnd.py --pdf "...chemin/PND_Tome3.pdf" --outdir out-2026-06/agreg
```

### Étape 3 — Ouvrir l'application HTML (1 min)

L'app est offline mais doit être servie via un mini-serveur local (le Web
Worker ne fonctionne pas en `file://`).

```bash
cd segmentation-client/agl-budget
python -m http.server 8000
```
Puis ouvrir **http://localhost:8000** dans le navigateur.

### Étape 4 — Charger les artefacts dans l'app (1 min)

Onglet **01 Matching & RMC** → zone *"Charger les résultats du pipeline Python"* :

Sélectionne **tous les fichiers** du dossier `out-2026-06/agreg/` + le `RMC.xlsx` +
le `opportunites.xlsx` en une seule fois. L'app détecte chaque fichier par son
nom et l'envoie au bon onglet.

Détection automatique :
- `RMC.xlsx` → onglets 11, 12
- `marche_total.csv` → onglet 02
- `volume_client.csv` → onglets 03, 06, 14
- `pdm.csv` → onglets 03, 13
- `projections.csv` → onglet 10
- `whitespaces.csv` → onglet 07
- `opportunites.xlsx` → onglet 05
- `budget_vs_real.csv` + `budget_secteur.csv` → onglet 04
- `projets_pnd.csv` + `axes_pnd.csv` → onglets 08, 09

---

## 3. Les 14 onglets — quoi faire de quoi

| # | Onglet | Pour qui | À faire dessus |
|---|---|---|---|
| 01 | **Matching & RMC** | Data Steward | Charger les bases sources + artefacts pipeline. Voir le RMC consolidé. |
| 02 | **Marché STATCOM** | DG / Stratégie | Comprendre la taille du marché CIV par métier × année. |
| 03 | **Position AGL** | DG / COMEX | Mesurer la PDM AGL vs concurrence par métier. CAGR AGL vs CAGR marché. |
| 04 | **Budget PFA vs Réel** | CFO / Contrôle de gestion | Suivre l'écart RUBRIKS vs IRIS par secteur et client. Identifier les sous-réalisations. |
| 05 | **Pipeline CRM** | Directeurs commerciaux | Funnel des opportunités, CAP pondéré, top deals. |
| 06 | **BCG** | Stratégie | Classer les clients en ★STAR / VACHE / DILEMME / POIDS MORT. |
| 07 | **White Spaces** | Sales / Business Development | Cibler les clients STATCOM absents du CRM (potentiel non capturé). |
| 08 | **PND vue d'ensemble** | Direction Stratégie | Répartition budget PND par axe, identifier les axes pertinents AGL. |
| 09 | **Top Projets PND** | BD / Sales | Filtrer les 1 480 actions du PND par axe + pertinence AGL. |
| 10 | **Projections 2026-2030** | Direction Stratégie | Trajectoires 3 scénarios par métier. Fonde la PFA et le budget. |
| 11 | **Préconisations Client** | KAM / Account Managers | Fiche stratégique par client : BCG + reco. |
| 12 | **Synthèse Executive** | COMEX | Note SCR auto-générée. Boutons export Word/PPT pour COMEX. |
| 13 | **Analyse Concurrentielle** | Stratégie | Bubble chart positionnement AGL : quadrants protéger/gagner. |
| 14 | **Cohortes & Rétention** | Sales Ops | Rétention par cohorte, churn, LTV top 20. |

Sur **chaque onglet**, en bas, une toolbar standard :
- 📋 **Copier prompt Claude** : génère un prompt enrichi avec les chiffres, prêt à coller dans claude.ai (gratuit, 0 coût additionnel).
- 📄 **Word** / 📊 **PPT** : exports COMEX-ready en 1 clic.
- 🖨️ **HTML** : version imprimable (`Ctrl+P` → Enregistrer en PDF). Fallback offline si pas de réseau.

---

## 4. Configuration

### 4.1 Changer les noms de colonnes d'une source
Édite **`python/config.py`**. Tous les scripts liront la nouvelle config.

### 4.2 Ajouter un alias manuel (pour rattraper un nom mal matché)
Édite **`python/aliases.json`** :
```json
{
  "RUBRIKS": {
    "SIR": "Société Ivoirienne de Raffinage",
    "CIE": "Compagnie Ivoirienne d'Électricité"
  }
}
```
Puis relance `python run_all.py …`.

### 4.3 Ajuster les hypothèses macro des projections
Édite **`python/08_projections.py`** → dictionnaire `MACRO_MULTIPLIERS`. Les
sources sont tracées dans chaque ligne CSV (`SOURCES_MACRO`).

Le poids relatif CAGR réel vs CAGR macro est paramétrable :
```bash
python 08_projections.py --metiers ... --w-real 0.7 --w-macro 0.3
```

---

## 5. FAQ

### Q : Un client n'apparaît pas dans le RMC, pourquoi ?
Soit il n'est pas dans le CRM (le RMC démarre depuis CRM, priorité 1 LOCKED).
Soit son nom est trop différent entre les bases → ajoute-le dans `aliases.json`
ou demande à corriger la source CRM.

### Q : Le coût total PND est colossal (~115 000 Mds FCFA), normal ?
Oui. C'est l'agrégat des Actions sur 5 ans, incluant les besoins de financement
public + privé indicatifs. Le PND est un document de planification ambitieux,
pas une enveloppe budgétaire ferme.

### Q : Pourquoi l'onglet 08 n'est pas une vraie carte géographique ?
Le PND Tome 3 ne fournit pas de coordonnées GPS pour les projets. On affiche
donc une vue par axe + résultat sectoriel à la place. Si tu as accès à une
base GPS des projets, on peut brancher une carte Leaflet en 1 sprint.

### Q : Comment historiser les analyses (« qui était ★STAR en janvier ? ») ?
Dupliquer le dossier de sortie chaque mois :
```
out-2026-01/  out-2026-02/  out-2026-03/  ...
```
Une fonction de comparaison N vs N-1 est dans la roadmap.

### Q : Faut-il connaître Python pour utiliser l'app ?
**Non** pour les onglets. Tu lances 2 commandes par mois (`run_all.py` et
`http.server`) puis tout se fait dans le navigateur. Un raccourci `.bat`/`.command`
peut être créé pour éviter de retaper les commandes.

### Q : Et la conformité RGPD ?
- Pas de remontée cloud : tout reste sur ton poste.
- Pas de données personnelles : seuls des noms d'entreprises et des CAP business.
- DPO recommandé pour valider l'usage si extension à d'autres entités.

---

## 6. Tester son installation

Pour s'assurer que rien n'est cassé après une mise à jour du code :
```bash
cd python
pip install pytest
python -m pytest tests/ -v
```
Doit afficher `21 passed`.

---

## 7. Architecture (rappel)

```
agl-budget/
├── index.html           coquille 14 onglets
├── js/                  12 modules (1 fichier = 1 module)
├── python/
│   ├── 01_normalise.py  → 11_budget_real.py   (11 scripts atomiques)
│   ├── run_all.py       (orchestrateur)
│   ├── config.py        (noms de colonnes)
│   ├── aliases.json     (corrections manuelles de matching)
│   └── tests/           (21 tests pytest)
├── GUIDE.md             (ce document)
└── README.md            (vue technique pour développeurs)
```

**Règle d'or** : 1 module = 1 fichier = débugable isolément. Aucun module
n'est en monolithe — chaque étape peut être réexécutée seule.

---

## 8. Support

Le code source est versionné dans le repo Git. Tout commit a un message
explicite. Pour signaler un bug ou demander une évolution, ouvrir une
issue dans le repo ou contacter l'équipe Stratégie.

**v1.0 — juin 2026 — CONFIDENTIEL**
