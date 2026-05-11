@echo off
REM Script de lancement de l'app Segmentation client (Windows)
REM Place-toi dans le dossier de l'app et lance Streamlit

cd /d "%~dp0"

REM Verifie que Python est installe
where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo [ERREUR] Python n'est pas installe ou pas dans le PATH.
    echo Telecharge-le sur https://www.python.org/downloads/
    echo Pendant l'installation, coche "Add Python to PATH".
    echo.
    pause
    exit /b 1
)

REM Installe les dependances si besoin (ne fait rien si deja installees)
echo Verification des dependances...
python -m pip install -q -r requirements.txt

REM Lance l'application
echo.
echo Lancement de l'application... le navigateur va s'ouvrir.
echo Pour arreter l'app : ferme cette fenetre.
echo.
python -m streamlit run app.py
pause
