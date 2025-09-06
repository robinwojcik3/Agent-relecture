L�?Tutilisateur souhaite mettre en place, dans un dǸp��t dǸdiǸ, une vǸritable �?oinfrastructure de relecture assistǸe par IA�?? qui soit structurǸe, claire et ergonomique. L�?TidǸe n�?Test pas simplement d�?Tavoir un super-prompt ponctuel �� lancer dans ChatGPT, mais plut��t de concevoir un environnement complet qui encadre et normalise le processus de relecture des documents d�?TǸtude. L�?Tobjectif est de pouvoir dǸposer un rapport (souvent un document Word volumineux, parfois de plusieurs centaines de pages) dans un dossier unique, puis de dǸclencher une sǸquence de travail guidǸe, o�� l�?Tagent IA applique une mǸthode standardisǸe pour analyser le rapport et restituer les rǸsultats de mani��re directement exploitable.

Au c�"ur de cette organisation, l�?Tutilisateur imagine quatre grandes catǸgories de relecture, correspondant aux diffǸrents types de livrables ou de sections critiques qu�?Til a �� traiter dans son mǸtier : la relecture des offres, la relecture du diagnostic (VNEI, Ǹtat initial), la relecture des impacts, et la relecture des mesures (Ǹvitement, rǸduction, compensation, suivi). Ces quatre modes sont con��us comme des modules distincts, chacun possǸdant sa propre logique, ses propres checklists mǸthodologiques et ses propres documents de rǸfǸrence. L�?Tutilisateur veut que le dǸp��t contienne donc, pour chaque mode, un dossier regroupant une checklist dǸtaillǸe (un fichier texte clair et normatif qui guide l�?TIA sur ce qu�?Til faut vǸrifier) et un espace pour stocker les documents de rǸfǸrence (guides mǸthodologiques, rǸglementations, normes, etc.). Cette structuration permet de garantir que la relecture ne repose pas seulement sur le texte du rapport, mais qu�?Telle est aussi ancrǸe dans les bonnes pratiques et exigences rǸglementaires, en fonction du contexte de relecture choisi.

Le point central de l�?TexpǸrience utilisateur est un fichier Python, nommǸ Start.py, qui sert d�?Tinterface de lancement. Lorsque ce script est exǸcutǸ, il doit guider l�?Tutilisateur pas �� pas, d�?Tabord en dǸtectant automatiquement le fichier Word �� relire dans le dossier d�?TentrǸe, puis en lui proposant, via une interaction dans le terminal, de choisir le mode de relecture parmi les quatre disponibles. L�?Tutilisateur veut que cette sǸlection soit fluide et intuitive, avec la possibilitǸ d�?Tutiliser les fl��ches du clavier pour se dǸplacer dans le menu et valider son choix par la touche EntrǸe. Une fois le mode sǸlectionnǸ, le script doit demander si la relecture doit porter sur l�?Tensemble du document ou uniquement sur une plage de pages dǸfinie. Cela rǸpond �� un besoin tr��s concret : dans les Ǹtudes d�?Timpact volumineuses, seule une partie prǸcise du rapport est parfois pertinente �� analyser (par exemple, le diagnostic Ǹcologique qui peut s�?TǸtendre de la page 40 �� la page 180). L�?Tutilisateur veut ainsi que l�?TIA soit orientǸe directement vers la partie concernǸe, ce qui rendra la relecture plus pertinente, plus rapide et plus ciblǸe.

Apr��s ces Ǹtapes de sǸlection, le script gǸn��re automatiquement un prompt standardisǸ qui contient toutes les instructions nǸcessaires pour l�?TIA : quel fichier relire, quelles pages analyser, quel mode appliquer, quelle checklist utiliser, et quels documents de rǸfǸrence consulter. Ce prompt est affichǸ dans le terminal, prǦt �� Ǧtre copiǸ-collǸ par l�?Tutilisateur dans l�?Tinterface de Windsurf, afin de lancer la relecture proprement dite avec l�?Tagent IA. L�?Tutilisateur insiste sur le fait que ce prompt doit Ǧtre uniformisǸ et rigoureux, afin de limiter les oublis, les angles morts et les formulations trop vagues. De cette mani��re, la logique de relecture reste constante d�?Tun projet �� l�?Tautre et ne dǸpend pas de la qualitǸ ou de la prǸcision d�?Tun prompt rǸdigǸ �� la volǸe.

L�� o�� l�?Tutilisateur apporte une exigence supplǸmentaire cruciale, c�?Test sur la finalitǸ du processus. Il ne veut pas d�?Tun syst��me qui produit plusieurs fichiers intermǸdiaires qu�?Til faudrait ensuite retraiter ou combiner �� la main. Ce qu�?Til souhaite, c�?Test qu�?T�� partir du moment o�� il copie le prompt gǸnǸrǸ dans Windsurf et qu�?Til lance la relecture avec l�?TIA, l�?TintǸgralitǸ du processus aille jusqu�?Tau bout automatiquement et ne nǸcessite plus aucune action complǸmentaire de sa part. Le rǸsultat attendu doit Ǧtre unique et parfaitement clair : une copie du rapport original au format Word, sauvegardǸe dans le dossier output/, qui contient l�?Tensemble des rǸvisions proposǸes par l�?TIA en mode suivi des modifications ainsi que tous les commentaires ajoutǸs. Autrement dit, le seul livrable qui doit exister �� l�?Tissue du processus est ce fichier Word annotǸ et suivi, prǦt �� Ǧtre ouvert et examinǸ. L�?Tutilisateur n�?Ta pas �� fusionner des fichiers, �� exǸcuter des scripts complǸmentaires, ni �� convertir manuellement des versions intermǸdiaires. Le but est d�?Tavoir une expǸrience �?oclǸ en main�?? o�� la seule tǽche de l�?Tutilisateur, apr��s avoir lancǸ Start.py et copiǸ le prompt, est de vǸrifier directement dans output/ le fichier final Word complet.

En rǸsumǸ, l�?Tobjectif est de disposer d�?Tun dǸp��t prǦt �� l�?Temploi qui industrialise et fiabilise le processus de relecture des rapports. L�?Tutilisateur veut une expǸrience simple (un seul script de dǸmarrage, un seul dǸp��t avec une structure fixe), mais derri��re cette simplicitǸ se cache une architecture rigoureuse : quatre modes de relecture spǸcialisǸs, chacun appuyǸ sur des checklists et des rǸfǸrences propres, un syst��me de sǸlection interactif pour cibler exactement le type et la partie du rapport �� analyser, un prompt standardisǸ pour assurer une relecture cohǸrente, et surtout un pipeline enti��rement intǸgrǸ dont le seul et unique rǸsultat final est un fichier Word annotǸ et rǸvisǸ, gǸnǸrǸ automatiquement dans output/. L�?Tensemble vise �� transformer une pratique artisanale de relecture en un processus normǸ, reproductible et tra��able, o�� l�?TIA n�?Test pas seulement un outil de suggestion mais un vǸritable copilote qui dǸlivre un livrable directement exploitable et finalisǸ.

PrǸcisions complǸmentaires �� intǸgrer dans l�?Tinterface graphique et le prompt standardisǸ
- Affichage de la table des mati��res: une fois `Start.py` lancǸ et le fichier Word sǸlectionnǸ, l�?Tinterface graphique doit proposer un bouton �� Afficher les titres/sections du document ��. Ce bouton dǸclenche une analyse du fichier et affiche, dans l�?Tinterface, la liste des titres/sections dǸtectǸs (ex: niveaux de titres Word). L�?Tutilisateur peut alors cocher/sǸlectionner prǸcisǸment les sections �� analyser.
- SǸlection des sections �� analyser: l�?Tutilisateur choisit manuellement, dans l�?Tinterface, les sections qui feront l�?Tobjet de la relecture. Ce choix conditionne la suite du pipeline.
- DǸclenchement de l�?Tanalyse: au clic sur �� Lancer l�?Tanalyse ��, l�?Tapplication effectue, dans cet ordre, des opǸrations locales et automatisǸes:
  1) Copie d�?Tabord le document original (sǸcuritǸ: jamais modifiǸ) vers un fichier de travail.
  2) DǸcoupe ensuite cette copie en ne conservant que les sections sǸlectionnǸes par l�?Tutilisateur (suppression des autres parties), de sorte qu�?Tun nouveau fichier DOCX �� copiǸ et dǸcoupǸ �� soit produit pour la relecture.
  3) GǸn��re alors le prompt standardisǸ et l�?Taffiche dans l�?Tinterface.
- Cible du prompt: le prompt doit indiquer explicitement que la relecture porte sur la �� copie du document original dǸcoupǸe selon les sections sǸlectionnǸes ��, et non sur l�?Toriginal complet ni sur une simple copie non filtrǸe. Le chemin de ce fichier �� copiǸ et dǸcoupǸ �� doit Ǧtre rǸfǸrencǸ dans le prompt et transmis �� l�?Tagent IA.
- Chaǩnage inchangǸ des Ǹtapes IA: le reste du pipeline demeure identique (checklist/refs ��' rǸdaction `rapport_revise.md` + `commentaires.csv` ��' conversion/compare ��' DOCX final dans `output/`), mais il s�?Tapplique �� la version �� copiǸ et dǸcoupǸ �� du document.


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

