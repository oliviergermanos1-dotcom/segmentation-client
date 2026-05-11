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
4. **Definir les segments** : conditions sur les colonnes agregees
   (`CA_EUR >= 100000`, `Nb_operations > 50`, ...), logique `AND` / `OR`,
   reordonnancement. Le premier segment qui matche gagne ; sinon `Non classe`.
   Preset disponible : 4 segments par quartiles de CA (VIP / Gold / Silver / Bronze).
5. **Classifier** et **exporter** en CSV ou Excel (avec onglet repartition).

La config des segments peut etre sauvegardee / rechargee en JSON.

## Lancer en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

L'app s'ouvre sur http://localhost:8501.

## Structure

- `app.py` — application Streamlit (upload, UI de regles, moteur de classification, export).
- `requirements.txt` — dependances (`streamlit`, `pandas`, `openpyxl`, `xlsxwriter`).
