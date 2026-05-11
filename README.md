# Segmentation client

Application Streamlit pour classer automatiquement les clients d'un export CRM
(Excel) selon des regles de segmentation definies dans l'interface.

## Workflow

1. **Importer** un ou plusieurs fichiers Excel d'extraction CRM Siebel
   (la limite d'export est de 150 000 lignes — uploadez plusieurs extracts,
   ils seront concatenes et dedupliques sur `JobfileNumber`).
   La ligne d'en-tete reelle est detectee automatiquement (les lignes
   "Filtres appliques" et "Exported data limited to 150000 rows" sont sautees).
2. **Filtrer** la periode (annees, mois) et les dimensions
   (entite, zone, segment marche, activite, etc.).
3. **Agreger par client** : choisir la cle client (`Customer`,
   `HQ CTO Customer Name`, ...) et les metriques (somme `Turnover in EUR`,
   `Direct GM in EUR`, `TEU`, `Freight Ton`). Sont aussi calcules :
   nombre d'operations, nombre de mois actifs, premiere / derniere activite.
4. **Definir les regles de segmentation** (section centrale de l'app) :
   - **Ajouter** une regle : saisir le nom (ex: `VIP`) -> elle apparait dans la liste.
   - **Modifier** une regle : changer le nom, la logique `AND`/`OR`, ajouter /
     retirer des conditions, changer la colonne, l'operateur ou la valeur.
   - **Supprimer** une regle : bouton "Supprimer" sur chaque carte.
   - **Reordonner** : boutons "monter" / "descendre". Le **premier** segment qui
     matche est attribue au client (les autres sont ignores) - donc mettre les
     regles les plus restrictives en haut. Les clients qui ne matchent aucune
     regle recoivent `Non classe`.
   - **Persistance automatique** : les regles sont sauvegardees dans
     `configs/segments_saved.json` a chaque modification, et rechargees au
     prochain demarrage. Plus besoin de re-saisir.
   - **Import / export JSON** : pour partager une config entre postes ou
     versionner ses jeux de regles.
   - **Preset "Quartiles CA"** : genere automatiquement 4 segments
     (VIP / Gold / Silver / Bronze) bases sur les quartiles du CA.
5. **Classifier** et **exporter** en CSV ou Excel (avec onglet repartition).
   L'export peut etre filtre par segment selectionne.

### Exemples de regles

| Segment | Logique | Conditions |
|---|---|---|
| VIP | AND | `CA_EUR >= 500000` |
| Gold | AND | `CA_EUR >= 100000` ET `Nb_operations >= 10` |
| Inactif | AND | `Derniere_activite <= 2024-12-31` |
| Strategique | OR | `Market Segment Name contient MINING` OU `TEU >= 1000` |

Operateurs disponibles : `==`, `!=`, `>`, `>=`, `<`, `<=`, `between`,
`contient`, `commence par`, `finit par`, `dans la liste`, `pas dans la liste`,
`regex`, `est vide`, `n'est pas vide`.

## Lancer en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

L'app s'ouvre sur http://localhost:8501.

## Structure

- `app.py` — application Streamlit (upload, UI de regles, moteur de classification, export).
- `requirements.txt` — dependances (`streamlit`, `pandas`, `openpyxl`, `xlsxwriter`).
