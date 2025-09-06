# Relecture IA — 4 modes (offre, diagnostic, impacts, mesures)

## Principe
1) Exécuter `Start.py` pour sélectionner le DOCX, le mode, et l’éventuel intervalle de pages.  
2) Copier-coller dans le chat Windsurf le **prompt généré** (affiché et écrit dans `work/agent_prompt_<mode>.txt`).  
   - L’agent exécute l’ensemble du pipeline local (extraction de références si présentes, génération de la version révisée + commentaires, conversion, comparaison Word, insertion des commentaires) et ne laisse en sortie que:  
     - `output/<nom>_AI_suivi+commentaires.docx` (copie du DOCX original avec suivi des modifications et commentaires intégrés).  
   - Les fichiers intermédiaires dans `work/` sont temporaires et nettoyés automatiquement.

## Dossiers
- `input/` : déposer le DOCX original.
- `modes/<mode>/instructions/checklist.md` : checklist ciblée par mode.
- `modes/<mode>/refs/` : PDF de références à considérer.
- `work/` : zone de travail (session.json et temporaires uniquement).
- `output/` : DOCX final avec suivi des modifications + commentaires.

## Dépendances (installation automatique si nécessaire)
- `tools/setup_tools.ps1` installe Pandoc et Poppler via winget.

