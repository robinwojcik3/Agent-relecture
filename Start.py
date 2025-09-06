#!/usr/bin/env python3
import os, sys, json, glob, msvcrt
from datetime import datetime

MODES = [
    ("offre", "Relecture des offres"),
    ("diagnostic", "Relecture du diagnostic"),
    ("impacts", "Relecture des impacts"),
    ("mesures", "Relecture des mesures"),
]

ROOT = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(ROOT, "input")
WORK_DIR  = os.path.join(ROOT, "work")
OUTPUT_DIR= os.path.join(ROOT, "output")
MODES_DIR = os.path.join(ROOT, "modes")

os.makedirs(WORK_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

def list_docx():
    paths = [p for p in glob.glob(os.path.join(INPUT_DIR, "*.docx")) if not os.path.basename(p).startswith("~$")]
    return sorted(paths, key=lambda p: os.path.getmtime(p), reverse=True)

def arrow_menu(title, options):
    idx = 0
    def redraw():
        os.system('cls' if os.name=='nt' else 'clear')
        print(title+"\n")
        for i,(key,label) in enumerate(options):
            prefix = "→ " if i==idx else "  "
            print(f"{prefix}{label} [{key}]")
        print("\nUtilisez ↑/↓ puis Entrée. Taper 1-9 fonctionne aussi.")
    redraw()
    while True:
        if os.name == 'nt':
            ch = msvcrt.getch()
            if ch in (b'\x00', b'\xe0'):
                ch2 = msvcrt.getch()
                if ch2 == b'H' and idx>0: idx -= 1; redraw()
                elif ch2 == b'P' and idx<len(options)-1: idx += 1; redraw()
            elif ch in (b'\r', b'\n'):
                return options[idx][0]
            elif ch.isdigit():
                dig = int(ch.decode())
                if 1 <= dig <= len(options): return options[dig-1][0]
        else:
            sel = input("Choix (1-{}): ".format(len(options))).strip()
            if sel.isdigit() and 1 <= int(sel) <= len(options):
                return options[int(sel)-1][0]

def pick_docx():
    docs = list_docx()
    if not docs:
        print(f"Aucun DOCX trouvé dans {INPUT_DIR}. Placez votre rapport puis relancez.")
        sys.exit(1)
    if len(docs) == 1:
        return docs[0]
    # menu sélection
    opts = [(str(i+1), os.path.basename(p)) for i,p in enumerate(docs)]
    key = arrow_menu("DOCX détectés (ordre récent → ancien)", opts)
    return docs[int(key)-1]

def ask_pages():
    ans = input("\nRelire tout le document ? [O/n] ").strip().lower()
    if ans in ("", "o", "oui", "y", "yes"):
        return None
    while True:
        try:
            deb = int(input("Page de début: ").strip())
            fin = int(input("Page de fin: ").strip())
            if deb <= 0 or fin < deb:
                print("Intervalle invalide. Recommencer.")
                continue
            return [deb, fin]
        except Exception:
            print("Entrées invalides. Recommencer.")

def build_agent_prompt(mode, docx_path, page_range):
    mode_label = dict(MODES)[mode]
    rel_docx = os.path.relpath(docx_path, ROOT).replace("\\", "/")
    pages_txt = "toutes les pages" if page_range is None else f"les pages {page_range[0]} à {page_range[1]}"
    checklist = f"modes/{mode}/instructions/checklist.md"
    refs_dir  = f"modes/{mode}/refs"
    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    out_docx  = f"output/{base_name}_AI_suivi+commentaires.docx"
    prompt = f"""Objectif: Réaliser une relecture **{mode_label}** du document `{rel_docx}`, en considérant {pages_txt}.

Contrainte: Ne pas utiliser d’API. Travailler dans le dépôt local. Produire **uniquement**:
- `{out_docx}` — copie du DOCX original avec révisions suivies et commentaires intégrés.

Référentiels à mobiliser:
- Checklist: `{checklist}`
- PDF de référence: dossiers `{refs_dir}/` (le cas échéant convertis en texte via `tools/extract_refs.ps1 -Mode {mode}` qui écrit dans `work/refs_txt/{mode}/`).

Règles d’analyse:
- Appliquer strictement la checklist du mode.
- Citer preuves et pages quand possible; tolérance zéro copier-coller résiduel.
- Poser **une seule** question fermée **uniquement si** un point critique empêche une conclusion fiable; sinon formuler une hypothèse explicite.

Exécution locale attendue (par l’agent, de bout en bout):
- Lire `{rel_docx}` (focus: {pages_txt}).
- Lire `{checklist}` + les textes `work/refs_txt/{mode}/*.txt` s’ils existent. Si absent, exécuter `pwsh -File tools/extract_refs.ps1 -Mode {mode}`.
- Écrire TEMPORAIREMENT: `work/rapport_revise.md` (version révisée intégrale, structure conservée) et `work/commentaires.csv` (schéma exact: `ancre_textuelle,commentaire,gravite,categorie`; gravite∈{{P1,P2,P3}}; categorie∈{{coherence,methodologie,reglementaire,carto,redaction}}). L’`ancre_textuelle` doit être un court extrait LITTÉRAL du paragraphe révisé ciblé (pas un ID HTML), unique par paragraphe.
- Générer le DOCX final en lançant `pwsh -File tools/run_pipeline.ps1 -Cleanup` (Word COM: Compare + insertion des commentaires), ce qui crée `{out_docx}` et supprime les fichiers temporaires dans `work/`.
- Ne laisser en fin de tâche QUE le fichier `{out_docx}`.

Contraintes de rédaction:
- Style sobre, technique, factuel. Pas de superlatifs.
- Tables en Markdown simple. Titres conservés.
- Ancre courte et unique par paragraphe ciblé.
"""
    return prompt

def main():
    os.system('cls' if os.name=='nt' else 'clear')
    print("=== Relecture IA — sélecteur ===\n")
    docx = pick_docx()
    print(f"Sélectionné: {os.path.basename(docx)}")
    mode = arrow_menu("Choisir le mode de relecture", [(m[0], f"{i+1}. {m[1]}") for i,m in enumerate(MODES)])
    print(f"\nMode: {mode}")
    pages = ask_pages()
    session = {
        "timestamp": datetime.now().isoformat(),
        "original_docx": os.path.relpath(docx, ROOT).replace("\\", "/"),
        "mode": mode,
        "page_range": pages
    }
    with open(os.path.join(WORK_DIR, "session.json"), "w", encoding="utf-8") as f:
        json.dump(session, f, ensure_ascii=False, indent=2)
    prompt = build_agent_prompt(mode, docx, pages)
    outp = os.path.join(WORK_DIR, f"agent_prompt_{mode}.txt")
    with open(outp, "w", encoding="utf-8") as f:
        f.write(prompt)
    print("\n--- PROMPT À COPIER DANS LE CHAT WINDSURF ---\n")
    print(prompt)
    print(f"\n[Enregistré dans] {outp}")
    print("\nProchaine étape:")
    print("  1) Coller ce prompt dans le chat Windsurf.")
    print("  2) À la fin, vérifier le fichier dans output/. Aucune autre action n’est requise.")
    print("\nTerminé.")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrompu.")

