#!/bin/bash

# Nom du dossier de l'environnement virtuel
VENV_DIR="venv"

# Commande Python à utiliser
PYTHON_CMD="python3"

# Vérifie si python3 est disponible
if ! command -v $PYTHON_CMD &> /dev/null
then
    echo "[ERREUR] python3 n'a pas été trouvé. Veuillez installer Python 3."
    exit 1
fi

echo "#############################################"
echo "#    Assistant de Relecture IA - Lanceur    #"
echo "#############################################"
echo

# Crée l'environnement virtuel s'il n'existe pas
if [ ! -d "$VENV_DIR" ]; then
    echo "[INFO] Création de l'environnement virtuel dans le dossier '$VENV_DIR'..."
    $PYTHON_CMD -m venv $VENV_DIR
    if [ $? -ne 0 ]; then
        echo "[ERREUR] La création de l'environnement virtuel a échoué."
        exit 1
    fi
fi

# Active l'environnement virtuel
source "$VENV_DIR/bin/activate"

echo "[INFO] Installation/Mise à jour des dépendances depuis requirements.txt..."
pip install -r requirements.txt --log pip_install.log
if [ $? -ne 0 ]; then
    echo "[ERREUR] L'installation des dépendances a échoué. Consultez pip_install.log pour les détails."
    exit 1
fi
echo "[INFO] Dépendances installées avec succès."
echo

# Lance l'application principale
echo "[INFO] Lancement de l'application..."
$PYTHON_CMD Start.py

# Désactive l'environnement à la fin
deactivate

echo
echo "[INFO] L'application s'est terminée."
