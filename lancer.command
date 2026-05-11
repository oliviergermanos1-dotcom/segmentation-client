#!/bin/bash
# Script de lancement de l'app Segmentation client (Mac / Linux)
# Double-clic dans Finder pour demarrer l'application

cd "$(dirname "$0")"

# Verifie que Python est installe
if ! command -v python3 &> /dev/null; then
    echo ""
    echo "[ERREUR] Python 3 n'est pas installe."
    echo "Telecharge-le sur https://www.python.org/downloads/"
    echo ""
    read -p "Appuie sur Entree pour fermer..."
    exit 1
fi

echo "Verification des dependances..."
python3 -m pip install -q -r requirements.txt

echo ""
echo "Lancement de l'application... le navigateur va s'ouvrir."
echo "Pour arreter l'app : Ctrl+C ou ferme cette fenetre."
echo ""
python3 -m streamlit run app.py
