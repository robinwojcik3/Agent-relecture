@echo off
setlocal

REM Chemin vers Python. Modifiez si nécessaire.
set PYTHON_EXE=python

REM Nom du dossier de l'environnement virtuel
set VENV_DIR=venv

echo #############################################
echo #    Assistant de Relecture IA - Lanceur    #
echo #############################################
echo.

REM Vérifie si Python est disponible
%PYTHON_EXE% --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERREUR] Python n'est pas trouvé.
    echo Veuillez l'installer (depuis python.org) et vous assurer qu'il est dans votre PATH.
    pause
    exit /b 1
)

REM Vérifie si l'environnement virtuel existe
if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo [INFO] Création de l'environnement virtuel dans le dossier '%VENV_DIR%'...
    %PYTHON_EXE% -m venv %VENV_DIR%
    if %errorlevel% neq 0 (
        echo [ERREUR] La création de l'environnement virtuel a échoué.
        pause
        exit /b 1
    )
    echo [INFO] Environnement créé avec succès.
)

REM Active l'environnement virtuel
call "%VENV_DIR%\Scripts\activate.bat"

echo [INFO] Installation/Mise à jour des dépendances depuis requirements.txt...
pip install -r requirements.txt --log pip_install.log
if %errorlevel% neq 0 (
    echo [ERREUR] L'installation des dépendances a échoué. Consultez pip_install.log pour les détails.
    pause
    exit /b 1
)
echo [INFO] Dépendances installées avec succès.
echo.

echo [INFO] Lancement de l'application...
%PYTHON_EXE% Start.py

echo.
echo [INFO] L'application s'est terminée.
endlocal
pause
