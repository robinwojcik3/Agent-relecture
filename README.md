# Agent de Relecture Assistée par IA

Ce projet fournit un environnement structuré pour la relecture de documents d'étude (rapports, offres, etc.) assistée par une IA. L'application est conçue pour être autonome, multi-plateforme et reproductible.

## Fonctionnalités

- **Interface Graphique Simple** : Pour sélectionner le document, le mode de relecture et les sections à analyser.
- **Modes de Relecture Spécialisés** : Quatre modes pré-configurés (Offre, Diagnostic, Impacts, Mesures) avec leurs propres checklists.
- **Analyse Ciblée** : Permet de relire le document entier ou des sections spécifiques (détectées via les styles de titres Word).
- **Pipeline Automatisé** : Le processus est entièrement automatisé après le clic sur "Lancer l'analyse".
- **Livrable Unique** : Génère un unique document Word (`.docx`) dans le dossier `output/`, contenant le texte révisé et les commentaires de l'IA.
- **Autonomie et Reproductibilité** : Le projet gère son propre environnement virtuel et ses dépendances, sans impacter le système global de l'utilisateur.

## Prérequis

- [Python 3.8+](https://www.python.org/downloads/) doit être installé sur votre système et accessible depuis le terminal (ajouté au PATH).

## Démarrage Rapide

1.  **Clonez ou téléchargez ce dépôt** sur votre machine locale.
2.  Placez le document Word (`.docx`) que vous souhaitez analyser dans le dossier `input/`.
3.  Exécutez le script de lancement correspondant à votre système d'exploitation :
    -   Sur **Windows** : double-cliquez sur `start.bat`.
    -   Sur **macOS** ou **Linux** : ouvrez un terminal et exécutez `sh start.sh`.

    > La première fois, le script créera automatiquement un environnement virtuel (`venv/`) et installera les dépendances nécessaires. Cela peut prendre une minute. Les lancements suivants seront beaucoup plus rapides.

4.  L'application se lancera. Suivez les instructions à l'écran pour choisir votre fichier, un mode de relecture et lancer l'analyse.
5.  Le document final, contenant les commentaires de l'IA, sera enregistré dans le dossier `output/`.

## Structure du Dépôt

- `input/` : Dossier pour placer les documents source à analyser.
- `output/` : Dossier où les rapports finaux commentés sont générés.
- `work/` : Dossier de travail temporaire utilisé par l'application (peut être ignoré).
- `modes/` : Contient les configurations pour chaque mode de relecture (checklists, etc.).
- `tools/` : Contient les modules techniques internes en Python.
- `Start.py` : Le script principal de l'application.
- `requirements.txt` : Fichier listant les dépendances Python figées.
- `start.bat` / `start.sh` : Scripts de lancement pour une initialisation facile.