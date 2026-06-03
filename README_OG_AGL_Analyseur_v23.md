# O&G AGL Analyseur OFFLINE — README Claude Code

## Fichier source
`O_G_AGL_Analyseur_OFFLINE_v23.html` — application HTML mono-fichier (~1.7 Mo)

---

## 1. ARCHITECTURE

### Structure des `<script>` tags (CRITIQUE)

Le fichier contient **7 tags `<script>`** dans cet ordre :

| # | Contenu | Taille | Notes |
|---|---------|--------|-------|
| 0 | `_scanFolderMini()` | ~1.8 KB | Mini-scanner dossier autonome (fallback) |
| 1 | xlsx.js (SheetJS) | ~639 KB | Lib parsing Excel |
| 2 | Chart.js | ~205 KB | Lib graphiques |
| 3 | jsPDF + jsPDF-AutoTable | ~364 KB + ~39 KB | Lib export PDF |
| 4 | **fflate.js** | ~32 KB | Lib compression/décompression ZIP |
| 5 | **Code applicatif** | ~342 KB | 206 fonctions, parsers, renderers |

### ⚠️ BUG CRITIQUE HISTORIQUE — Frontière fflate ↔ app

**Le bug le plus grave de l'historique du projet** : les tags `<script>` de fflate et de l'app DOIVENT être **séparés** :

```html
<!-- ✅ CORRECT (v15, v23 corrigé) -->
...return _e});
</script>
<script>
"use strict";
/* ===================== ANALYSEUR AGL — MOTEUR ... */
var COLORS={...
```

```html
<!-- ❌ CASSÉ (v16→v22, certaines versions v23) -->
...setTimeout?se
"use strict";
var COLORS={...
```

**Quand fflate est fusionné/tronqué** : `var COLORS` se retrouve à `depth=1` (dans la IIFE fflate) → toutes les fonctions app (runAnalysis, renderCA...) sont **piégées dans le scope fflate** → invisibles en global → **RIEN NE MARCHE**.

**Test de vérification** :
```python
scripts = re.findall(r'<script(?:\s[^>]*)?>([\s\S]*?)</script>', html)
app = next(s for s in scripts if 'function renderCA' in s)
colors_idx = app.find('var COLORS=')
before = app[:colors_idx]
depth = before.count('{') - before.count('}')
assert depth == 0, f"COLORS à depth={depth} — fflate fusionné !"
```

**Lors de tout patch** : toujours vérifier que `depth == 0` à `var COLORS=`.

---

## 2. BUG OUVERT — Fichier volumineux 23 Mo qui bloque

### Contexte
Le fichier `FICHE DE PLANNING JOURNALIER 28052026-JOUR-NUIT.xlsx` fait **23.1 Mo** (images/dessins embarqués, potentiellement des milliers de lignes vides).

### Flux de lecture actuel
```
readWBSmart(file)
  → readWBWorker(file)          # tente le Web Worker d'abord
      → FileReader.readAsArrayBuffer
      → envoie buf brut au Worker
      → Worker fait XLSX.read(u8, {type:'array', cellStyles:false, ...})
      → timeout dynamique : max(45s, 8s × Mo)
  → si Worker échoue → readWB(file)   # fallback thread principal
      → needsClean(u8) → si > 2 Mo → cleanXlsxBytes(u8)
      → XLSX.read(cleaned, {...})
```

### Problème
1. **Worker path** : XLSX.read sur 23 Mo brut dans le Worker → très lent (parsing de tout le contenu XML y compris images inline). Le Worker timeout probablement.
2. **Fallback readWB** : `cleanXlsxBytes` fait `fflate.unzipSync(23Mo)` + `fflate.zipSync()` **sur le thread principal** → **UI gelée pendant 30+ secondes**.
3. En v15 ça marchait parce que fflate était correctement chargé → le nettoyage réduisait le fichier à ~2 Mo → le Worker parsait rapidement. Mais depuis les patches v16+, fflate était cassé → cleanXlsxBytes retournait `null` → le fichier brut de 23 Mo bloquait tout.

### Solution recommandée
**Déplacer `cleanXlsxBytes` dans le Web Worker** : le Worker a accès à fflate (l'inclure dans le blob Worker). Ainsi le nettoyage se fait off-thread → pas de gel UI.

**Implémentation suggérée** :
1. Dans `_ensureWorker()`, inclure le code source de fflate.js dans le blob Worker (comme xlsx.js y est déjà) :
```javascript
var fflSrc = _libSrc('offline: fflate.js');
var blob = new Blob([xlsxSrc, '\n', fflSrc || '', '\n', glue], {type:'text/javascript'});
```
2. Dans le `glue` du Worker, ajouter le nettoyage AVANT XLSX.read :
```javascript
self.onmessage = function(e) {
  var u8 = new Uint8Array(e.data.buf);
  // Nettoyer si gros fichier
  if (u8.length > 2*1024*1024 && typeof fflate !== 'undefined') {
    try { u8 = cleanXlsxBytesWorker(u8) || u8; } catch(e) {}
  }
  var wb = XLSX.read(u8, {...});
  // ...
};
```
3. Inclure `cleanXlsxBytes` (ou une version simplifiée) dans le blob Worker.
4. Augmenter le Worker timeout à `max(60s, 10s × Mo)`.

### `cleanXlsxBytes` — code actuel (v23)
```javascript
function cleanXlsxBytes(u8) {
  if (typeof fflate === 'undefined') return null;
  try {
    var keep = fflate.unzipSync(u8, {filter: function(f) {
      var n = f.name;
      if (/xl\/(drawings|ink|media|charts|embeddings|printerSettings|activeX)\//i.test(n)) return false;
      if (/\.(png|jpe?g|emf|wmf|gif|bmp|vml|bin)$/i.test(n)) return false;
      if (/vbaProject/i.test(n)) return false;
      if (/docProps\/thumbnail/i.test(n)) return false;
      return true;
    }});
    // Nettoyer les XML de sheets
    Object.keys(keep).forEach(function(n) {
      if (/xl\/worksheets\/sheet[0-9]+\.xml$/i.test(n)) {
        var txt = fflate.strFromU8(keep[n]);
        // Limiter à 200 000 lignes
        var rowCount = (txt.match(/<row /g) || []).length;
        if (rowCount > 200000) { /* tronquer */ }
        txt = txt.replace(/<drawing[^>]*\/>/gi, '')
                 .replace(/<legacyDrawing[^>]*\/>/gi, '')
                 .replace(/<controls>[\s\S]*?<\/controls>/gi, '');
        keep[n] = fflate.strToU8(txt);
      }
    });
    return fflate.zipSync(keep, {level: 0}); // level 0 = store only
  } catch(e) { return null; }
}
```

### `_libSrc()` — extraction du source d'une lib
```javascript
function _libSrc(marker) {
  var scripts = document.querySelectorAll('script');
  for (var i = 0; i < scripts.length; i++) {
    if (scripts[i].textContent.indexOf(marker) !== -1)
      return scripts[i].textContent;
  }
  return null;
}
```
Cherche le commentaire `/* offline: fflate.js */` ou `/* offline: xlsx.js */` dans les tags `<script>`.

---

## 3. BOUTON UPLOAD — État actuel

### HTML
```html
<button id="btnPickFolder" onclick="(function(){
  var fi = document.getElementById('fileInput');
  if (typeof window.showDirectoryPicker === 'function') {
    window.showDirectoryPicker({mode:'read'}).then(function(h) {
      window._folderHandle = h;
      if (typeof scanFolder === 'function') scanFolder(h);
      else if (typeof _scanFolderMini === 'function') _scanFolderMini(h);
    }).catch(function(e) {
      if (e && e.name === 'AbortError') return;
      if (fi) fi.click();
    });
  } else {
    if (fi) fi.click();
  }
})()">📂 Choisir le dossier RECEPTION</button>
<input type="file" id="fileInput" multiple accept=".xlsx,.xls,.csv" style="display:none"
       onchange="if(typeof handleFiles==='function'){handleFiles(Array.prototype.slice.call(this.files));this.value='';}">
```

### Flux
1. **Chrome/Edge HTTPS** : `showDirectoryPicker` → dossier sélectionné → `scanFolder(handle)` lit tous les .xlsx/.xls/.csv
2. **Fallback** (Firefox, Safari, file://) : `fileInput.click()` → sélection multiple manuelle
3. **`_scanFolderMini`** (script[0]) : scanner de secours si `scanFolder` n'est pas disponible (script principal crashé). Stocke les fichiers dans `window.allFiles` et active `btnAnalyze`.

### Problème potentiel
Si le script principal ne charge pas (erreur fflate/depth), `scanFolder` et `handleFiles` ne sont pas disponibles. Le `_scanFolderMini` prend le relais mais `runAnalysis` n'est pas non plus disponible → le bouton "Lancer l'analyse" ne fait rien.

---

## 4. DONNÉES PARSÉES — Structure

| Type | Parser | Champs clés par période |
|------|--------|------------------------|
| QHSE | `parseQHSE` | accidents, jsa, totalEff, trifIOGP, ltifIOGP, hours |
| POINTAGE | `parsePointage` | count, jour, nuit, workers[] |
| PLANNING | `parsePlanning` | ops[], pos, neg, nbOps |
| ACCOSTAGE | `parseAccostage` | navires[], total, byClient{} |
| LOYERS/CAP | `parseLoyers` | byClient{}, totalAnnuel, globalByClient |
| CA Flash | `parseCA` | deps[], monthsCA[], byClient{}, budget[]{reel,budget,reel25} |
| CHARGES | `parseCharges` | byAff{}, byFour{}, lignes[] |
| ACTIONS | `parseActions` | actions[] (ref, desc, statut, resp, echeance) |
| GENERIQUE | `parseGenerique` | sheets{} (headers, rows, numericCols) |

### Variables globales clés
```javascript
var allFiles = [];           // File objects chargés
var datasets = {};           // {TYPE: [{fileName, data, confidence, warnings}]}
var activePanel = 'SYNTHESE';
var _activeClientFilter = null;
var _filterApplying = false;
var GLOBAL_SEL = null;       // Sélection de période globale
var periodSel = {};          // Sélection par type
var charts = [];             // Chart.js instances (pour destroy)
```

---

## 5. FONCTIONS CLÉS

### Utilitaires
| Fonction | Rôle |
|----------|------|
| `el(id)` | `document.getElementById(id)` |
| `esc(s)` | Échappe HTML |
| `fmtXOF(n)` | Formate en XOF (ex: "1 234 567") |
| `canonClient(s)` | Normalise un nom de client pour matching cross-tab |
| `getPeriods(type)` | Retourne les périodes fusionnées (multi-fichiers) + filtre client actif |
| `selectedIndices(type)` / `synIndicesFor(type)` | Indices des périodes sélectionnées |
| `mkChart(id, config)` | Wrapper Chart.js (destroy + create) |
| `kpiCard(title, value, sub, trend, big, insight)` | Génère une carte KPI HTML |
| `expertBlock(type)` | Bloc d'analyse experte par type |
| `annotationBlock(type)` | Zone de notes persistante (localStorage) |

### Render functions (24)
Chaque `renderXXX(panel)` reçoit un élément DOM et écrit dans `panel.innerHTML`.

Principales : `renderSynthese`, `renderQHSE`, `renderPointage`, `renderPlanning`, `renderActions`, `renderAccostage`, `renderCA`, `renderLoyers`, `renderCharges`, `renderMarge`, `renderCorrelation`, `renderGenerique`.

### Persistance
| Mécanisme | Usage |
|-----------|-------|
| localStorage `oeg_annot_*` | Annotations par onglet |
| localStorage `oeg_hist_v20_*` | Historique analyses (snapshots) |
| localStorage `oeg_bench` | Benchmarks IOGP personnalisés |
| IDB `analyseEnc` | Dernière analyse (chiffré AES-GCM) |

---

## 6. FEATURES IMPLÉMENTÉES (38/38)

- Inter font, palette AGL navy/gold
- Escape unifié (tous overlays), dark mode, responsive mobile hamburger
- Tri colonnes (makeSortable), export Excel tableaux
- KPI flash CODIR, battle card client, rating A/B/C
- Panneau alertes, benchmark IOGP QHSE
- Bannière erreurs parsing, annotations premium, historique analyses
- Filtre croisé client×période, recherche globale cross-onglets
- N-1 insight KPI CA + Marge analytique N-1 (ca25, colonne, courbe)
- Notifications navigateur, score qualité widget topbar
- Multi-fichiers fusion YM, applyClientFilter anti-boucle
- Corrélation croisée, drill-down période, sauvegarde config
- Export PDF panel, export CODIR PDF, print CSS

---

## 7. MÉTHODE DE PATCH

### Règles absolues
1. **Python string replacement** sur le HTML — jamais de reconstruction regex des tags `<script>`
2. **`node --check`** sur le JS extrait après chaque patch
3. **Diff fonctions** v_prev → v_new : vérifier 0 fonction perdue
4. **Runtime test** : mock data → appeler chaque render function → vérifier output > 500 chars
5. **Vérifier depth=0** à `var COLORS=` après tout patch touchant les scripts

### Pattern de patch
```python
with open('file.html', 'r') as f:
    html = f.read()

# Remplacement direct dans le HTML (JAMAIS de regex reconstitution)
html = html.replace(OLD_EXACT_STRING, NEW_STRING, 1)

with open('file.html', 'w') as f:
    f.write(html)
```

### ⚠️ DANGER : reconstitution par regex
**NE JAMAIS FAIRE** :
```python
# ❌ CASSE LES TAGS </script> — INTERDIT
html = html[:app_match.start(1)] + new_app_js + html[app_match.end(1):]
```
La regex `<script(?:\s[^>]*)?>([\s\S]*?)</script>` utilise le lazy matching `*?` qui ne capture pas correctement les scripts contenant des faux `<script>` dans les strings JS (jsPDF notamment). La reconstitution mange des `</script>` → le script suivant fusionne avec le précédent → tout casse.

**TOUJOURS** utiliser `html.replace(OLD_EXACT, NEW_EXACT, 1)` pour les modifications.

---

## 8. FICHIERS DE RÉFÉRENCE

| Fichier | État |
|---------|------|
| `O_G_AGL_Analyseur_OFFLINE_v15.html` | ✅ Dernière version où fflate marchait (référence pour la frontière `<script>`) |
| `O_G_AGL_Analyseur_OFFLINE_v22.html` | ✅ Dernière version stable avant les patches upload (features complètes, fflate cassé) |
| `O_G_AGL_Analyseur_OFFLINE_v23.html` | ⚠️ Version courante — fflate réparé, upload fonctionnel, **bug fichier 23 Mo ouvert** |

---

## 9. RÉSUMÉ DU BUG À CORRIGER

**Fichier** : `FICHE DE PLANNING JOURNALIER 28052026-JOUR-NUIT.xlsx` (23.1 Mo)

**Symptôme** : l'analyse se bloque à "Lecture 3/8… nettoyage parasites" et ne progresse plus.

**Cause racine** : `cleanXlsxBytes()` (fflate unzipSync + zipSync) s'exécute sur le **thread principal** → gel UI de 30+ secondes. Le Worker timeout avant la fin.

**Fix requis** : déplacer le nettoyage dans le Web Worker (off-thread). Inclure fflate.js dans le blob Worker via `_libSrc('offline: fflate.js')`, puis appeler `cleanXlsxBytes` dans le Worker **avant** `XLSX.read`.

**Résultat attendu** : le fichier de 23 Mo est nettoyé en arrière-plan (Worker), l'UI reste fluide, le fichier nettoyé (~2 Mo) est parsé rapidement par XLSX.read.
