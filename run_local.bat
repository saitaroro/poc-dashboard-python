@echo off
REM ------------------------------------------------------------
REM run_local.bat
REM - Lance le helper Python `run_local.py` qui gère l'installation
REM   des dépendances et le démarrage de l'application avec logs.
REM - Préfère un launcher Python (logging) plutôt que des echo.
REM ------------------------------------------------------------

cd /d "%~dp0"

REM Appel du script Python de démarrage (utilise logging)
python run_local.py

REM Le script Python affichera les logs et retournera le code d'erreur
exit /b %ERRORLEVEL%