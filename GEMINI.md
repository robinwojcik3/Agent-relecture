L’utilisateur souhaite mettre en place, dans un dépôt dédié, une véritable « infrastructure de relecture assistée par IA » qui soit structurée, claire et ergonomique. L’idée n’est pas simplement d’avoir un super-prompt ponctuel à lancer dans ChatGPT, mais plutôt de concevoir un environnement complet qui encadre et normalise le processus de relecture des documents d’étude. L’objectif est de pouvoir déposer un rapport (souvent un document Word volumineux, parfois de plusieurs centaines de pages) dans un dossier unique, puis de déclencher une séquence de travail guidée, où l’agent IA applique une méthode standardisée pour analyser le rapport et restituer les résultats de manière directement exploitable.

Au cœur de cette organisation, l’utilisateur imagine quatre grandes catégories de relecture, correspondant aux différents types de livrables ou de sections critiques qu’il a à traiter dans son métier : la relecture des offres, la relecture du diagnostic (VNEI, état initial), la relecture des impacts, et la relecture des mesures (évitement, réduction, compensation, suivi). Ces quatre modes sont conçus comme des modules distincts, chacun possédant sa propre logique, ses propres checklists méthodologiques et ses propres documents de référence. L’utilisateur veut que le dépôt contienne donc, pour chaque mode, un dossier regroupant une checklist détaillée (un fichier texte clair et normatif qui guide l’IA sur ce qu’il faut vérifier) et un espace pour stocker les documents de référence (guides méthodologiques, réglementations, normes, etc.). Cette structuration permet de garantir que la relecture ne repose pas seulement sur le texte du rapport, mais qu’elle est aussi ancrée dans les bonnes pratiques et exigences réglementaires, en fonction du contexte de relecture choisi.

Le point central de l’expérience utilisateur est un fichier Python, nommé Start.py, qui sert d’interface de lancement. Lorsque ce script est exécuté, il doit guider l’utilisateur pas à pas, d’abord en détectant automatiquement le fichier Word à relire dans le dossier d’entrée, puis en lui proposant, via une interaction dans le terminal, de choisir le mode de relecture parmi les quatre disponibles. L’utilisateur veut que cette sélection soit fluide et intuitive, avec la possibilité d’utiliser les flèches du clavier pour se déplacer dans le menu et valider son choix par la touche Entrée. Une fois le mode sélectionné, le script doit demander si la relecture doit porter sur l’ensemble du document ou uniquement sur une plage de pages définie. Cela répond à un besoin très concret : dans les études d’impact volumineuses, seule une partie précise du rapport est parfois pertinente à analyser (par exemple, le diagnostic écologique qui peut s’étendre de la page 40 à la page 180). L’utilisateur veut ainsi que l’IA soit orientée directement vers la partie concernée, ce qui rendra la relecture plus pertinente, plus rapide et plus ciblée.

Après ces étapes de sélection, le script génère automatiquement un prompt standardisé qui contient toutes les instructions nécessaires pour l’IA : quel fichier relire, quelles pages analyser, quel mode appliquer, quelle checklist utiliser, et quels documents de référence consulter. Ce prompt est affiché dans le terminal, prêt à être copié-collé par l’utilisateur dans l’interface de Windsurf, afin de lancer la relecture proprement dite avec l’agent IA. L’utilisateur insiste sur le fait que ce prompt doit être uniformisé et rigoureux, afin de limiter les oublis, les angles morts et les formulations trop vagues. De cette manière, la logique de relecture reste constante d’un projet à l’autre et ne dépend pas de la qualité ou de la précision d’un prompt rédigé à la volée.

Là où l’utilisateur apporte une exigence supplémentaire cruciale, c’est sur la finalité du processus. Il ne veut pas d’un système qui produit plusieurs fichiers intermédiaires qu’il faudrait ensuite retraiter ou combiner à la main. Ce qu’il souhaite, c’est qu’à partir du moment où il copie le prompt généré dans Windsurf et qu’il lance la relecture avec l’IA, l’intégralité du processus aille jusqu’au bout automatiquement et ne nécessite plus aucune action complémentaire de sa part. Le résultat attendu doit être unique et parfaitement clair : une copie du rapport original au format Word, sauvegardée dans le dossier output/, qui contient l’ensemble des révisions proposées par l’IA en mode suivi des modifications ainsi que tous les commentaires ajoutés. Autrement dit, le seul livrable qui doit exister à l’issue du processus est ce fichier Word annoté et suivi, prêt à être ouvert et examiné. L’utilisateur n’a pas à fusionner des fichiers, à exécuter des scripts complémentaires, ni à convertir manuellement des versions intermédiaires. Le but est d’avoir une expérience « clé en main » où la seule tâche de l’utilisateur, après avoir lancé Start.py et copié le prompt, est de vérifier directement dans output/ le fichier final Word complet.

En résumé, l’objectif est de disposer d’un dépôt prêt à l’emploi qui industrialise et fiabilise le processus de relecture des rapports. L’utilisateur veut une expérience simple (un seul script de démarrage, un seul dépôt avec une structure fixe), mais derrière cette simplicité se cache une architecture rigoureuse : quatre modes de relecture spécialisés, chacun appuyé sur des checklists et des références propres, un système de sélection interactif pour cibler exactement le type et la partie du rapport à analyser, un prompt standardisé pour assurer une relecture cohérente, et surtout un pipeline entièrement intégré dont le seul et unique résultat final est un fichier Word annoté et révisé, généré automatiquement dans output/. L’ensemble vise à transformer une pratique artisanale de relecture en un processus normé, reproductible et traçable, où l’IA n’est pas seulement un outil de suggestion mais un véritable copilote qui délivre un livrable directement exploitable et finalisé.

Précisions complémentaires à intégrer dans l’interface graphique et le prompt standardisé : une fois Start.py lancé et le fichier Word sélectionné, l’interface graphique doit proposer un bouton « Afficher les titres/sections du document » qui déclenche une analyse et affiche la liste hiérarchisée des titres détectés ; l’utilisateur peut alors cocher précisément les sections à analyser. Après cette sélection, au clic sur « Lancer l’analyse », l’application effectue automatiquement trois opérations : (1) copier le document original vers un fichier de travail (jamais modifié directement), (2) découper cette copie en ne conservant que les sections choisies afin de produire un fichier DOCX « copié et découpé », et (3) générer le prompt standardisé correspondant et l’afficher dans l’interface. Le prompt doit indiquer explicitement que la relecture porte sur cette copie découpée et non sur l’original complet ni sur une simple duplication non filtrée, le chemin du fichier devant y être référencé et transmis à l’agent IA. Le reste du pipeline (application des checklists et références, production du rapport_revise.md et du commentaires.csv, conversion et comparaison, puis génération du DOCX final dans output/) demeure inchangé, mais s’applique uniquement à la version découpée du document.


Spécification GUI et Orchestration — V2
- But: refondre l’application pour que l’interface graphique (lancée via Start.py) permette de sélectionner les SECTIONS du rapport à partir de la table des matières et d’orchestrer toute la relecture de bout en bout. Aucune étape manuelle ultérieure ne doit être requise. Le seul livrable attendu est une COPIE Word finale, enregistrée dans `output/`, contenant toutes les révisions en suivi des modifications et tous les commentaires.

- Contraintes globales
  - Ne jamais modifier le fichier original. Toujours travailler sur une copie.
  - Pas d’API externe. Tout s’exécute en local.
  - Traçabilité minimale : journal d’exécution, session.json, nommage horodaté.
  - Robustesse : si la table des matières est absente, détecter les titres Word (Styles “Titre 1/2/3”) et reconstituer la hiérarchie.

- Exigences d’interface (GUI lancée par Start.py)
  1) Section “Fichier source”
     - Bouton « Sélectionner le fichier Word… » : ouvre l’explorateur et renseigne le chemin du DOCX à relire.
     - Bouton « Ouvrir le dossier » : ouvre l’emplacement du fichier sélectionné.
     - Règle immuable affichée en clair : « Le fichier original ne sera jamais modifié. Le traitement s’effectue sur une copie. »

  2) Section “Sections du document”
     - Bouton « Afficher les sections » : analyse la table des matières (ou, à défaut, les styles de titre) et affiche la liste hiérarchisée des titres avec numéros (ex. “4. Impacts”, “5. Mesures”, “6.1 Méthodologie”, etc.).
     - Liste multi-sélection avec cases à cocher. L’utilisateur peut cocher une ou plusieurs sections. Afficher un compteur (n sections sélectionnées).
     - Si aucune section n’est cochée, interpréter comme « document entier » (après confirmation).

  3) Section “Mode de relecture”
     - Quatre boutons exclusifs, visuellement mutuellement exclusifs : « Offre », « Diagnostic », « Impacts », « Mesures ». Mettre en évidence le mode actif.

  4) Section “Dossier de sortie”
     - Bouton « Choisir le dossier de sortie… » : sélection du répertoire où sera enregistré le livrable final.
     - Bouton « Ouvrir le dossier de sortie » : accès direct à ce répertoire.

  5) Section “Lancer l’analyse”
     - Bouton unique « Lancer l’analyse ». Au clic, exécuter la séquence automatique complète décrite ci-dessous.
     - Zone de texte en lecture seule affichant le prompt opérationnel généré (pour audit), mais aucune intervention manuelle n’est requise.

- Séquence automatique au clic « Lancer l’analyse »
  Étape A — Préparation
  - Valider les prérequis : fichier source choisi, mode de relecture sélectionné, dossier de sortie défini.
  - Créer un répertoire de travail `work/` si absent. Écrire `work/session.json` avec : chemin source, sections sélectionnées, mode, dossier de sortie, horodatage.
  - Copier le DOCX original vers `work/<nom>_copie_<YYYYMMDD_HHMMSS>.docx`.

  Étape B — Découpage par sections
  - À partir de la table des matières (ou de la hiérarchie de titres détectée), extraire uniquement les sections cochées par l’utilisateur et construire une « COPIE DÉCOUPÉE ».
  - Enregistrer cette copie découpée sous : `work/<nom>_SECTIONS_<hash_ou_timestamp>.docx`.
  - Cette copie découpée est le SEUL document soumis à la relecture IA. Ne jamais utiliser l’original dans les étapes suivantes.

  Étape C — Relecture IA locale (exécutée par l’agent, en local)
  - Convertir la copie découpée en Markdown de travail.
  - Appliquer la checklist et les ressources du mode sélectionné :
    - Mode « Offre » → `modes/offre/instructions/checklist.md` + `modes/offre/refs/`
    - Mode « Diagnostic » → `modes/diagnostic/instructions/checklist.md` + `modes/diagnostic/refs/`
    - Mode « Impacts » → `modes/impacts/instructions/checklist.md` + `modes/impacts/refs/`
    - Mode « Mesures » → `modes/mesures/instructions/checklist.md` + `modes/mesures/refs/`
  - Produire ces deux fichiers dans `work/` :
    1) `work/rapport_revise.md` : version révisée intégrale, structure conservée, corrections intégrées, limitée aux sections sélectionnées.
    2) `work/commentaires.csv` : colonnes exactes `ancre_textuelle,commentaire,gravite,categorie` (gravite∈{P1,P2,P3} ; categorie∈{coherence,methodologie,reglementaire,carto,redaction}). Ancres courtes et uniques par paragraphe.

  Étape D — Génération du livrable final Word (suivi des modifs + commentaires)
  - Convertir `work/rapport_revise.md` en DOCX (révisé).
  - Comparer le DOCX révisé à la COPIE DÉCOUPÉE (et non à l’original) pour générer un document Word avec suivi des modifications.
  - Injecter les commentaires de `work/commentaires.csv` à l’emplacement des ancres.
  - Enregistrer le seul livrable attendu dans le dossier de sortie choisi :
    `output/<nom>_AI_suivi+commentaires_<YYYYMMDD_HHMMSS>.docx`
  - Ne produire aucun autre fichier final dans `output/`. Tout le reste reste en `work/`.

- Règles de qualité
  - Afficher une barre d’état ou un journal pas-à-pas (Préparation → Découpage → Relecture IA → Conversion → Compare → Commentaires → Livrable prêt).
  - En cas d’échec de détection de la table des matières, consigner la bascule vers la détection des styles de titres et poursuivre.
  - Si des sections sélectionnées n’existent pas (changement de structure), notifier et proposer « document entier » ou sélection corrigée.
  - À la fin, afficher en clair le chemin exact du fichier final Word dans `output/` et le nombre d’insertions/suppressions/Commentaires appliqués (si accessible).

- Sécurité et invariants
  - L’original dans `input/` n’est jamais ouvert en écriture.
  - Toutes les écritures se font dans `work/` et `output/`.
  - Aucune étape utilisateur supplémentaire n’est requise après « Lancer l’analyse ». Le processus est intégralement automatisé. Le seul contrôle humain est de vérifier le fichier final dans `output/`.

## Gestion de l'encodage

Tous les fichiers du repository doivent être strictement encodés en UTF-8. Aucune intervention d'un agent IA ne doit jamais produire de corruption de type "mojibake" (séquences illisibles comme Ã©, Ã¨, â€™, â€¦).

**Règle impérative :** Toute lecture ou écriture de fichier doit être effectuée explicitement en UTF-8. Avant validation, il est impératif de s'assurer que le contenu conserve un affichage en français correct avec accents, caractères typographiques (é, è, à, ç, œ, °, …) et ponctuation. Toute apparition de mojibake est considérée comme une erreur critique bloquante et doit être corrigée immédiatement.

- Livrable attendu de cette tâche
  - Une interface Start.py conforme aux exigences ci-dessus, opérationnelle de bout en bout.
  - Un pipeline interne qui, à partir des sections choisies, produit automatiquement le DOCX final unique en suivi des modifications et commentaires dans `output/`, sans action manuelle intermédiaire.

