# Segmentation client

Application Streamlit pour classer automatiquement les clients d'un export CRM
(Excel) selon des regles de segmentation definies dans l'interface.

## Workflow

1. **Importer** le fichier Excel brut (`.xlsx` ou `.xls`).
2. **Definir les segments** : pour chaque segment, ajouter une ou plusieurs
   conditions sur les colonnes (operateurs : `==`, `>`, `contient`, `between`,
   `dans la liste`, `regex`, etc.) et choisir la logique `AND` / `OR`.
3. **Ordonner** les segments : le premier qui matche est attribue au client.
   Les non-matches recoivent le segment `Non classe`.
4. **Lancer la classification** puis **exporter** le resultat en CSV ou Excel.

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
