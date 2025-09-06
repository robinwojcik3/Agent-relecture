#!/usr/bin/env python3
import os, sys, json, shutil
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

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

os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(WORK_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

def build_agent_prompt(mode, rel_docx, out_docx):
    mode_label = dict(MODES)[mode]
    checklist = f"modes/{mode}/instructions/checklist.md"
    refs_dir  = f"modes/{mode}/refs"
    pages_txt = "toutes les pages"
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
- Lire `{checklist}` + les textes `work/refs_txt/{mode}/*.txt` s’ils existent. Si absent, exécuter `powershell -ExecutionPolicy Bypass -File tools/extract_refs.ps1 -Mode {mode}`.
- Écrire TEMPORAIREMENT: `work/rapport_revise.md` (version révisée intégrale, structure conservée) et `work/commentaires.csv` (schéma exact: `ancre_textuelle,commentaire,gravite,categorie`; gravite∈{{P1,P2,P3}}; categorie∈{{coherence,methodologie,reglementaire,carto,redaction}}). L’`ancre_textuelle` doit être un court extrait LITTÉRAL du paragraphe révisé ciblé (pas un ID HTML), unique par paragraphe.
- Générer le DOCX final en lançant `powershell -ExecutionPolicy Bypass -File tools/run_pipeline.ps1 -Cleanup` (Word COM: Compare + insertion des commentaires), ce qui crée `{out_docx}` et supprime les fichiers temporaires dans `work/`.
- Ne laisser en fin de tâche QUE le fichier `{out_docx}`.

Contraintes de rédaction:
- Style sobre, technique, factuel. Pas de superlatifs.
- Tables en Markdown simple. Titres conservés.
- Ancre courte et unique par paragraphe ciblé.
"""
    return prompt

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Relecture IA — Assistant graphique")
        self.geometry("900x700")
        self.resizable(True, True)

        self.source_path = None      # chemin vers le fichier source original
        self.copy_relpath = None     # chemin relatif vers la copie dans input/
        self.output_dir = OUTPUT_DIR # dossier de sortie choisi
        self.mode = None             # 'offre'|'diagnostic'|'impacts'|'mesures'

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 8}

        # Section 1 — Fichier Word source
        frame1 = tk.LabelFrame(self, text="1) Fichier Word à relire")
        frame1.pack(fill="x", **pad)
        self.src_var = tk.StringVar(value="Aucun fichier sélectionné")
        tk.Label(frame1, textvariable=self.src_var, anchor="w").pack(fill="x", padx=10, pady=5)
        btns1 = tk.Frame(frame1)
        btns1.pack(anchor="w", padx=10, pady=5)
        tk.Button(btns1, text="Choisir un fichier…", command=self.choose_source).pack(side="left", padx=5)
        tk.Button(btns1, text="Ouvrir le dossier du fichier", command=self.open_source_folder).pack(side="left", padx=5)
        tk.Label(frame1, text="Règle: le fichier original n’est jamais modifié; travail sur une copie en input/.", fg="#555").pack(anchor="w", padx=10, pady=5)

        # Section 2 — Mode de relecture (4 boutons exclusifs)
        frame2 = tk.LabelFrame(self, text="2) Mode de relecture")
        frame2.pack(fill="x", **pad)
        self.mode_buttons = {}
        btnrow = tk.Frame(frame2)
        btnrow.pack(anchor="w", padx=10, pady=5)
        for key,label in MODES:
            b = tk.Button(btnrow, text=label, width=28, command=lambda k=key: self.set_mode(k))
            b.pack(side="left", padx=6, pady=3)
            self.mode_buttons[key] = b

        # Section 3 — Dossier de sortie
        frame3 = tk.LabelFrame(self, text="3) Dossier de sortie")
        frame3.pack(fill="x", **pad)
        self.out_var = tk.StringVar(value=self.output_dir)
        tk.Label(frame3, textvariable=self.out_var, anchor="w").pack(fill="x", padx=10, pady=5)
        btns3 = tk.Frame(frame3)
        btns3.pack(anchor="w", padx=10, pady=5)
        tk.Button(btns3, text="Choisir un dossier…", command=self.choose_output_dir).pack(side="left", padx=5)
        tk.Button(btns3, text="Ouvrir le dossier de sortie", command=self.open_output_dir).pack(side="left", padx=5)

        # Section 4 — Lancer l’analyse et afficher le prompt
        frame4 = tk.LabelFrame(self, text="4) Lancer l’analyse")
        frame4.pack(fill="both", expand=True, **pad)
        tk.Button(frame4, text="Lancer l’analyse", command=self.launch_analysis).pack(anchor="w", padx=10, pady=5)
        self.prompt_txt = tk.Text(frame4, height=20, wrap="word")
        self.prompt_txt.pack(fill="both", expand=True, padx=10, pady=5)

    # Actions
    def choose_source(self):
        path = filedialog.askopenfilename(filetypes=[("Documents Word", "*.docx")])
        if not path:
            return
        self.source_path = os.path.abspath(path)
        self.src_var.set(self.source_path)
        # Créer une copie immédiatement dans input/
        try:
            self.copy_relpath = self._copy_to_input(self.source_path)
        except Exception as e:
            messagebox.showerror("Erreur de copie", str(e))
            self.copy_relpath = None

    def open_source_folder(self):
        path = self.source_path
        if not path:
            messagebox.showinfo("Info", "Aucun fichier sélectionné.")
            return
        folder = os.path.dirname(path)
        try:
            os.startfile(folder)
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    def set_mode(self, key):
        self.mode = key
        # mise en évidence visuelle
        for k, b in self.mode_buttons.items():
            if k == key:
                b.configure(bg="#2e86de", fg="white", activebackground="#1e5aa6")
            else:
                b.configure(bg=self.cget("bg"), fg="black", activebackground=self.cget("bg"))

    def choose_output_dir(self):
        d = filedialog.askdirectory()
        if not d:
            return
        self.output_dir = os.path.abspath(d)
        os.makedirs(self.output_dir, exist_ok=True)
        self.out_var.set(self.output_dir)

    def open_output_dir(self):
        d = self.output_dir or OUTPUT_DIR
        os.makedirs(d, exist_ok=True)
        try:
            os.startfile(d)
        except Exception as e:
            messagebox.showerror("Erreur", str(e))

    # Copie protégée vers input/
    def _copy_to_input(self, src_path):
        os.makedirs(INPUT_DIR, exist_ok=True)
        base = os.path.basename(src_path)
        name, ext = os.path.splitext(base)
        candidate = os.path.join(INPUT_DIR, base)
        n = 1
        while os.path.exists(candidate):
            candidate = os.path.join(INPUT_DIR, f"{name} (copie {n}){ext}")
            n += 1
        shutil.copy2(src_path, candidate)
        rel = os.path.relpath(candidate, ROOT).replace("\\", "/")
        return rel

    def launch_analysis(self):
        if not self.copy_relpath:
            messagebox.showwarning("Manque fichier", "Veuillez sélectionner un fichier Word.")
            return
        if not self.mode:
            messagebox.showwarning("Manque mode", "Veuillez sélectionner un mode de relecture.")
            return
        # Construire chemin DOCX final
        base_name = os.path.splitext(os.path.basename(self.copy_relpath))[0]
        out_docx = os.path.join(self.output_dir, f"{base_name}_AI_suivi+commentaires.docx")
        # Ecrire session.json
        session = {
            "timestamp": datetime.now().isoformat(),
            "original_docx": self.copy_relpath,  # la copie dans input/
            "mode": self.mode,
            "page_range": None,
            "output_dir": self.output_dir,
        }
        with open(os.path.join(WORK_DIR, "session.json"), "w", encoding="utf-8") as f:
            json.dump(session, f, ensure_ascii=False, indent=2)
        # Construire prompt
        prompt = build_agent_prompt(self.mode, self.copy_relpath, out_docx)
        outp = os.path.join(WORK_DIR, f"agent_prompt_{self.mode}.txt")
        with open(outp, "w", encoding="utf-8") as f:
            f.write(prompt)
        # Afficher dans l'interface
        self.prompt_txt.delete("1.0", tk.END)
        self.prompt_txt.insert(tk.END, prompt)
        messagebox.showinfo("Prompt prêt", f"Prompt généré et enregistré dans:\n{outp}")


if __name__ == "__main__":
    App().mainloop()
