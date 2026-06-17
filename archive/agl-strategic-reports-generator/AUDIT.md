# Audit — AGL Strategic Reports Generator v9

> **Note d'archivage** : ce projet est différent de `segmentation-client`.
> Il est archivé ici temporairement car la session Claude Code on the web qui
> a réalisé l'audit était restreinte à ce dépôt. À déplacer dès que possible
> vers un dépôt dédié (`agl-strategic-reports-generator`) ou vers `APP-STRAT`.

**Fichier audité** : `AGL_Strategic_Reports_Generator_9.html` (4561 lignes)
**Date** : 2026-06-17
**Auditeur** : Claude (Opus 4.7)

---

## 1. Vue d'ensemble

Application React monofichier (Tailwind + Babel standalone) pour la Direction
Stratégie d'Africa Global Logistics — génère des rapports portuaires et
logistiques pour le COMEX.

- **4 métiers** : Import/Export Maritime (TEU/Tonnes), Import/Export Aérien (Tonnes)
- **Onglets UI** : un par métier + Multi-métiers (clients ≥2 métiers) + Hinterland Import + Hinterland Export
- **Calculs** : KPI, Top clients (Pareto), pays, transitaires, armateurs, YoY, PDM AGL, segments conteneurisé/conventionnel, Hinterland
- **Exports** : DOCX, PDF, PPTX, HTML autoporteur (tous fonctionnels au niveau des métiers individuels)

---

## 2. État du projet — fait

- Architecture `METIER_SCHEMAS` propre (l. 70-126)
- `buildAnalysis()` complet (l. 287+) : KPI, YoY, PDM AGL, Pareto, segments, tops
- **Hinterland — calcul** : `buildHinterlandAnalysis()` (l. 693-737) fonctionnel pour import et export
- **Hinterland — UI preview** : `HinterlandPreviewContent` (l. 2719-2852) complet — insights McKinsey, cartographie pays (tous, avec PDM AGL + Δ YoY), lieux de livraison granularité ville
- **Hinterland — dans Multi-métiers** : synthèses Top 10 pays Import (l. 3022+) et Export (l. 3065+)
- Exports DOCX/PDF/PPTX/HTML métier par métier OK

---

## 3. État du projet — partiel / inachevé

### 3.1. Hinterland MANQUANT des exports

Hinterland est visible en UI mais **absent des 4 exports finaux** — c'est le gros chantier resté en plan.

| Export | Ligne fin boucle métiers | Section Hinterland |
|---|---|---|
| DOCX (`generateDOCX`) | l. 1203 | absente |
| PDF (`generatePDF`) | l. 1618 | absente |
| PPTX (`generatePPTX`) | l. 2037 | absente |
| HTML autoporteur (`generateHTML`) | l. 2357-2370 | absente |

**Cause technique** : l'objet `hinterland` est calculé dans `AppRoot` (l. 2641) mais n'est jamais transmis aux fonctions `generateDOCX/PDF/PPTX/HTML` (handlers l. 4144-4180 ne passent que `analysesForReport`).

**Indices que c'était bien en cours** : commentaire l. 2616 « *L'onglet multi-métiers occupe l'index validAnalyses.length, et les onglets Hinterland après* » — l'intention d'aller jusqu'aux exports était claire.

### 3.2. Analyses Hinterland sans rows brutes

`buildHinterlandAnalysis()` construit les objets via `buildAnalysis()` sur données filtrées
mais ne stocke pas les rows brutes — donc les onglets Hinterland ne peuvent pas se
re-filtrer dynamiquement comme les métiers.

---

## 4. Marqueurs d'inachevé

- Aucun `TODO` / `FIXME` explicite dans le code
- Les `console.error` présents (l. 3940, 3943, 3967, 3987, 4048, 4082, 4116, 4162, 4180, 4239) sont du logging IndexedDB légitime, pas des oublis

---

## 5. Vérification visuelle (capture utilisateur)

Le rapport généré observé chez l'utilisateur (`AGL_Rapport_STATCOM_2026-05-29T16-43.html`)
montre 4 onglets uniquement : Import Maritime / Export Maritime / Import Aérien /
Export Aérien. Aucun onglet Multi-métiers ni Hinterland — **confirmation visuelle
du gap identifié**.

---

## 6. Prochain chantier recommandé

**Finir Hinterland dans les exports** (~200-300 lignes) :

1. Modifier les signatures : `generateDOCX(analyses, hinterland, options)` (idem PDF/PPTX/HTML) — handlers l. 4144-4180
2. Après la boucle métiers, ajouter :
   - `if (hinterland?.importHinterland) { /* section Hinterland Import */ }`
   - `if (hinterland?.exportHinterland) { /* section Hinterland Export */ }`
3. Réutiliser la structure de `HinterlandPreviewContent` (l. 2733-2851) en l'adaptant au markup de chaque format

**Estimation : Hinterland est à ~70 % fini** — logique métier OK, UI OK, exports KO.
