#!/usr/bin/env python3
import os
import sys
import json
import csv
import shutil
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Tenter d'importer les dépendances et guider l'utilisateur si elles sont manquantes.
try:
    from docx import Document
except ImportError:
    # Ce bloc ne devrait être atteint que si l'utilisateur exécute le script directement
    # sans passer par les scripts de lancement (start.bat/start.sh).
    # On affiche une erreur claire et on quitte.
    root = tk.Tk()
    root.withdraw() # Cacher la fenêtre principale de tkinter
    messagebox.showerror(
        "Dépendance Manquante",
        "La bibliothèque 'python-docx' n'est pas installée.\n\n"
        "Veuillez fermer cette fenêtre et exécuter le script 'start.bat' (ou 'start.sh') "
        "pour installer automatiquement les dépendances."
    )
    sys.exit(1)

# Importer les nouveaux outils Python locaux
from tools.python_tools import add_comments_to_docx


MODES = [
    ("offre", "Offre"),
    ("diagnostic", "Diagnostic"),
    ("impacts", "Impacts"),
    ("mesures", "Mesures"),
]

ROOT = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(ROOT, "input")
WORK_DIR = os.path.join(ROOT, "work")
OUTPUT_DIR = os.path.join(ROOT, "output")

os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(WORK_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


# ------------------------- utilitaires -------------------------

def ts_now():
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def norm_style_name(name: str) -> str:
    if not name:
        return ""
    return name.strip().lower()

def detect_head_level(style_name: str):
    n = norm_style_name(style_name)
    for prefix in ("heading ", "titre "):
        if n.startswith(prefix):
            try:
                return int(n.split(prefix, 1)[1].split()[0])
            except (ValueError, IndexError):
                return None
    return None

class Section:
    def __init__(self, title: str, level: int, start_index: int):
        self.title = title.strip()
        self.level = level
        self.start_index = start_index
        self.end_index = None
        self.number = None

    def label(self) -> str:
        num = (self.number or "?")
        return f"{num} {self.title}".strip()

def analyze_sections(docx_path: str):
    doc = Document(docx_path)
    secs = []
    for i, p in enumerate(doc.paragraphs):
        style_name = getattr(p.style, "name", "")
        lvl = detect_head_level(style_name)
        if lvl is not None and lvl >= 1:
            txt = p.text.strip()
            if txt:
                secs.append(Section(txt, lvl, i))
    for idx, s in enumerate(secs):
        s.end_index = (secs[idx + 1].start_index if idx + 1 < len(secs) else len(doc.paragraphs))
    counters = {}
    for s in secs:
        for k in list(counters.keys()):
            if k > s.level:
                counters.pop(k, None)
        counters[s.level] = counters.get(s.level, 0) + 1
        parts = [str(counters.get(l, 0)) for l in range(1, s.level + 1)]
        s.number = ".".join(parts)
    return secs

def filter_paragraphs_by_sections(docx_path: str, chosen: list, sections: list) -> str:
    doc = Document(docx_path)
    new_doc = Document()
    # Transférer les styles pour conserver la mise en forme
    for style in doc.styles:
        if style.type:
            try:
                new_doc.styles.add_style(style.name, style.type)
            except Exception:
                pass

    included_ranges = []
    chosen_set = set(chosen)
    for i, s in enumerate(sections):
        if i in chosen_set:
            included_ranges.append((s.start_index, s.end_index))
    if not included_ranges:
        included_ranges = [(0, len(doc.paragraphs))]

    for start, end in included_ranges:
        for i in range(start, end):
            p_src = doc.paragraphs[i]
            p_dest = new_doc.add_paragraph()
            for r_src in p_src.runs:
                r_dest = p_dest.add_run(r_src.text)
                r_dest.bold = r_src.bold
                r_dest.italic = r_src.italic
                r_dest.underline = r_src.underline
                if r_src.font.name: r_dest.font.name = r_src.font.name
                if r_src.font.size: r_dest.font.size = r_src.font.size
            if p_src.style.name:
                try:
                    p_dest.style = p_src.style.name
                except Exception:
                    pass

    base = os.path.splitext(os.path.basename(docx_path))[0]
    out_path = os.path.join(WORK_DIR, f"{base}_SECTIONS_{ts_now()}.docx")
    new_doc.save(out_path)
    return out_path

def docx_to_markdown(docx_path: str) -> str:
    doc = Document(docx_path)
    lines = []
    for p in doc.paragraphs:
        style_name = getattr(p.style, "name", "")
        lvl = detect_head_level(style_name)
        t = (p.text or "").strip()
        if not t:
            lines.append("")
            continue
        if lvl is not None and 1 <= lvl <= 6:
            lines.append("#" * lvl + " " + t)
        else:
            lines.append(t)
    return "\n".join(lines) + "\n"

def sanitize_text(s: str) -> str:
    import re
    s = s.replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"\s+([,;:!?])", r"\1", s)
    s = re.sub(r"\(\s+", "(", s)
    s = re.sub(r"\s+\)", ")", s)
    return s.strip()

ALLOWED_CATEGORIES = ["coherence", "methodologie", "reglementaire", "carto", "redaction"]

def load_checklist(mode: str):
    path = os.path.join(ROOT, "modes", mode, "instructions", "checklist.md")
    try:
        with open(path, "r", encoding="utf-8") as f:
            lines = [ln.strip(" \t-•") for ln in f.read().splitlines()]
        return [l for l in lines if l and not l.startswith("#")]
    except Exception:
        return []

def classify_comment_from_text(txt: str, mode: str):
    t = txt.lower()
    if any(k in t for k in ["méthod", "methodo", "protocole"]): return ("P2", "methodologie")
    if any(k in t for k in ["réglement", "natura 2000", "loi"]): return ("P2", "reglementaire")
    if any(k in t for k in ["carte", "carto", "figure", "légende"]): return ("P3", "carto")
    if any(k in t for k in ["cohér", "cohérence"]): return ("P2", "coherence")
    return ("P3", "redaction")

def generate_review(md_in: str, mode: str):
    import re
    lines = md_in.splitlines()
    revised_lines = []
    comments = []
    checklist = load_checklist(mode)
    used_anchors = {}

    def anchor_for(text: str) -> str:
        base = sanitize_text(text)[:25]
        base = re.sub(r"[^A-Za-z0-9À-ÿ\-\s]", "", base).strip()
        if not base: base = "ancre"
        k, i = base, 1
        while k.lower() in used_anchors:
            i += 1
            k = f"{base} {i}"
        used_anchors[k.lower()] = True
        return k

    full_text = "\n".join(lines).lower()
    first_heading_text = next((sanitize_text(ln.lstrip("#").strip()) for ln in lines if ln.startswith("#")), None)

    for item in checklist:
        key = item.split(":")[0]
        if key and key.lower() not in full_text:
            grav, cat = classify_comment_from_text(item, mode)
            comments.append({
                "ancre_textuelle": (first_heading_text or "Introduction"),
                "commentaire": f"Vérifier couverture de la checklist: '{item}'.",
                "gravite": grav, "categorie": cat,
            })

    for ln in lines:
        raw = ln.rstrip()
        if not raw.strip():
            revised_lines.append("")
            continue
        if raw.startswith("#"):
            parts = raw.split(" ", 1)
            title = sanitize_text(parts[1] if len(parts) > 1 else raw.lstrip("#"))
            revised_lines.append(parts[0] + " " + title)
            continue
        cleaned = sanitize_text(raw)
        revised_lines.append(cleaned)

        txt_low = cleaned.lower()
        make_comment = None
        if len(cleaned) > 600: make_comment = ("P3", "redaction", "Paragraphe très long: envisager de le scinder.")
        if any(x in txt_low for x in ["tbd", "???", "à définir"]): make_comment = ("P1", "coherence", "Marqueur d'incertain repéré: préciser/retirer.")
        if ("carte" in txt_low or "figure" in txt_low) and not re.search(r"\d", cleaned): make_comment = ("P3", "carto", "Référence à une carte/figure sans numéro.")

        if make_comment:
            grav, cat, note = make_comment
            comments.append({
                "ancre_textuelle": anchor_for(cleaned), "commentaire": note,
                "gravite": grav, "categorie": cat if cat in ALLOWED_CATEGORIES else "redaction",
            })

    return "\n".join(revised_lines) + "\n", comments

def markdown_to_docx(md_text: str, out_path: str, reference_docx: Document):
    doc = Document()
    for style in reference_docx.styles: # Conserver les styles
        if style.type: 
            try: doc.styles.add_style(style.name, style.type)
            except: pass

    for raw in md_text.splitlines():
        if not raw.strip():
            doc.add_paragraph("")
            continue
        if raw.startswith("#"):
            level = len(raw) - len(raw.lstrip("#"))
            p = doc.add_paragraph(raw[level:].strip())
            try: p.style = f"Heading {min(level,6)}"
            except: pass
        else:
            doc.add_paragraph(raw)
    doc.save(out_path)

def write_comments_csv(rows, path: str):
    fieldnames = ["ancre_textuelle", "commentaire", "gravite", "categorie"]
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for r in rows:
            if r.get("ancre_textuelle"):
                w.writerow({k: str(r.get(k, "")) for k in fieldnames})

# ---------------------- Fenêtre de sélection ----------------------

class SectionsDialog(tk.Toplevel):
    def __init__(self, master, sections, preselected_idx=None):
        super().__init__(master)
        self.title("Sélection des sections")
        self.geometry("560x420")
        self.sections, self.result = sections, None
        frm = tk.Frame(self); frm.pack(fill="both", expand=True, padx=10, pady=10)
        self.lb = tk.Listbox(frm, selectmode=tk.EXTENDED, activestyle="none")
        sb = ttk.Scrollbar(frm, orient="vertical", command=self.lb.yview)
        self.lb.config(yscrollcommand=sb.set); self.lb.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")
        for i, s in enumerate(sections):
            self.lb.insert(tk.END, f"{ '    ' * max(0, s.level - 1)}{s.label()}")
            if i in (preselected_idx or []): self.lb.selection_set(i)
        self.lb.bind("<ButtonRelease-1>", self.on_click_block)
        btns = tk.Frame(self); btns.pack(fill="x", padx=10, pady=(0, 10))
        tk.Button(btns, text="Tout cocher", command=self.sel_all).pack(side="left")
        tk.Button(btns, text="Tout décocher", command=self.sel_none).pack(side="left", padx=6)
        tk.Button(btns, text="Valider", command=self.on_ok).pack(side="right")
        tk.Button(btns, text="Annuler", command=self.on_cancel).pack(side="right", padx=6)
        self.bind("<Return>", lambda e: self.on_ok()); self.bind("<Escape>", lambda e: self.on_cancel())
        self.transient(master); self.grab_set(); self.lb.focus_set()

    def sel_all(self): self.lb.select_set(0, tk.END)
    def sel_none(self): self.lb.select_clear(0, tk.END)
    def on_ok(self): self.result = list(self.lb.curselection()); self.destroy()
    def on_cancel(self): self.result = None; self.destroy()
    def on_click_block(self, event):
        if (event.state & 0x0001) or (event.state & 0x0004): return
        idx = self.lb.nearest(event.y)
        if idx < 0: return
        end = idx + 1
        while end < len(self.sections) and self.sections[end].level > self.sections[idx].level: end += 1
        block = range(idx, end)
        if all(i in set(self.lb.curselection()) for i in block): [self.lb.selection_clear(i) for i in block]
        else: self.lb.select_clear(0, tk.END); [self.lb.selection_set(i) for i in block]

# ------------------------- GUI -------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Relecture IA — Assistant graphique")
        self.geometry("1000x760")
        self.source_path, self.copy_relpath, self.mode = None, None, None
        self.output_dir = OUTPUT_DIR
        self.sections, self.section_vars = [], []
        self._build_ui()

    def log(self, msg: str):
        self.log_txt.config(state="normal")
        self.log_txt.insert(tk.END, f"{datetime.now().strftime('%H:%M:%S')} — {msg}\n")
        self.log_txt.see(tk.END); self.log_txt.config(state="disabled"); self.update_idletasks()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 8}
        panes = tk.PanedWindow(self, orient=tk.HORIZONTAL); panes.pack(fill="both", expand=True)
        left = tk.Frame(panes); right = tk.Frame(panes)
        panes.add(left, minsize=520); panes.add(right, minsize=360)

        f1 = tk.LabelFrame(left, text="1) Fichier source"); f1.pack(fill="x", **pad)
        self.src_var = tk.StringVar(value="Aucun fichier sélectionné")
        tk.Label(f1, textvariable=self.src_var, anchor="w").pack(fill="x", padx=10, pady=5)
        b1 = tk.Frame(f1); b1.pack(anchor="w", padx=10, pady=5)
        tk.Button(b1, text="Sélectionner le fichier Word…", command=self.choose_source).pack(side="left", padx=5)
        tk.Button(b1, text="Ouvrir le dossier", command=self.open_source_folder).pack(side="left", padx=5)
        tk.Label(f1, text="Le fichier original ne sera jamais modifié.", fg="#555").pack(anchor="w", padx=10, pady=5)

        f2 = tk.LabelFrame(left, text="2) Sections du document"); f2.pack(fill="x", **pad)
        b2 = tk.Frame(f2); b2.pack(anchor="w", padx=10, pady=5)
        tk.Button(b2, text="Afficher les sections", command=self.show_sections).pack(side="left", padx=5)
        self.sections_count_var = tk.StringVar(value="0 section sélectionnée")
        tk.Label(b2, textvariable=self.sections_count_var).pack(side="left", padx=10)
        self.sections_canvas = tk.Canvas(f2, height=100); self.sections_canvas.pack(fill="x", padx=10, pady=5)
        self.sections_frame = tk.Frame(self.sections_canvas)
        self.sections_canvas.create_window((0, 0), window=self.sections_frame, anchor="nw")

        f3 = tk.LabelFrame(left, text="3) Mode de relecture"); f3.pack(fill="x", **pad)
        self.mode_buttons = {}
        btnrow = tk.Frame(f3); btnrow.pack(anchor="w", padx=10, pady=5)
        for key, label in MODES:
            b = tk.Button(btnrow, text=label, width=20, command=lambda k=key: self.set_mode(k))
            b.pack(side="left", padx=6, pady=3); self.mode_buttons[key] = b

        f4 = tk.LabelFrame(left, text="4) Dossier de sortie"); f4.pack(fill="x", **pad)
        self.out_var = tk.StringVar(value=self.output_dir)
        tk.Label(f4, textvariable=self.out_var, anchor="w").pack(fill="x", padx=10, pady=5)
        b4 = tk.Frame(f4); b4.pack(anchor="w", padx=10, pady=5)
        tk.Button(b4, text="Choisir le dossier…", command=self.choose_output_dir).pack(side="left", padx=5)
        tk.Button(b4, text="Ouvrir le dossier", command=self.open_output_dir).pack(side="left", padx=5)

        f5 = tk.LabelFrame(left, text="5) Lancer l’analyse"); f5.pack(fill="both", expand=True, **pad)
        tk.Button(f5, text="Lancer l’analyse", command=self.launch_analysis).pack(anchor="w", padx=10, pady=5)
        tk.Label(f5, text="Journal d’exécution").pack(anchor="w", padx=10)
        self.log_txt = tk.Text(f5, height=10, wrap="word", state="disabled")
        self.log_txt.pack(fill="both", expand=True, padx=10, pady=5)

        pr_frame = tk.LabelFrame(right, text="Prompt opérationnel (audit)"); pr_frame.pack(fill="both", expand=True, **pad)
        self.prompt_txt = tk.Text(pr_frame, wrap="word")
        sb = ttk.Scrollbar(pr_frame, orient="vertical", command=self.prompt_txt.yview)
        self.prompt_txt.config(yscrollcommand=sb.set, state="disabled"); self.prompt_txt.pack(side="left", fill="both", expand=True); sb.pack(side="right", fill="y")

    def choose_source(self):
        path = filedialog.askopenfilename(filetypes=[("Documents Word", "*.docx")])
        if not path: return
        self.source_path = os.path.abspath(path)
        self.src_var.set(self.source_path)
        try: self.copy_relpath = self._copy_to_input(self.source_path); self.log(f"Copie créée: {self.copy_relpath}")
        except Exception as e: messagebox.showerror("Erreur", str(e)); self.copy_relpath = None

    def open_source_folder(self):
        if not self.source_path: messagebox.showinfo("Info", "Aucun fichier."); return
        try: os.startfile(os.path.dirname(self.source_path))
        except Exception as e: messagebox.showerror("Erreur", str(e))

    def set_mode(self, key):
        self.mode = key
        for k, b in self.mode_buttons.items(): b.config(bg=("#2e86de" if k == key else "SystemButtonFace"), fg=("white" if k == key else "SystemButtonText"))

    def choose_output_dir(self):
        d = filedialog.askdirectory()
        if d: self.output_dir = os.path.abspath(d); os.makedirs(d, exist_ok=True); self.out_var.set(d)

    def open_output_dir(self):
        d = self.output_dir or OUTPUT_DIR; os.makedirs(d, exist_ok=True)
        try: os.startfile(d)
        except Exception as e: messagebox.showerror("Erreur", str(e))

    def _copy_to_input(self, src_path):
        name, ext = os.path.splitext(os.path.basename(src_path))
        dest = os.path.join(INPUT_DIR, f"{name}_copie_{ts_now()}{ext}")
        shutil.copy2(src_path, dest)
        rel = os.path.relpath(dest, ROOT)
        return rel.replace("\\", "/")

    def show_sections(self):
        if not self.copy_relpath: messagebox.showwarning("Info", "Sélectionnez un fichier."); return
        try:
            self.sections = analyze_sections(os.path.join(ROOT, self.copy_relpath))
            pre = [i for i, v in enumerate(self.section_vars) if v.get()] if self.section_vars else []
            dlg = SectionsDialog(self, self.sections, pre)
            self.wait_window(dlg)
            if dlg.result is not None:
                self.section_vars = [tk.BooleanVar(value=(i in set(dlg.result))) for i in range(len(self.sections))]
                for w in list(self.sections_frame.children.values()): w.destroy()
                summary = [self.sections[i].label() for i in sorted(dlg.result)][:6]
                if summary:
                    for lab in summary: tk.Label(self.sections_frame, text=f"☑ {lab}", anchor="w").pack(anchor="w")
                    if len(dlg.result) > 6: tk.Label(self.sections_frame, text="…", anchor="w").pack(anchor="w")
                else: tk.Label(self.sections_frame, text="(document entier)", anchor="w").pack(anchor="w")
            self.update_sections_count()
        except Exception as e: messagebox.showerror("Erreur", f"Analyse des sections impossible:\n{e}")

    def update_sections_count(self):
        n = sum(1 for v in self.section_vars if v.get())
        self.sections_count_var.set(f"{n} section(s) sélectionnée(s)")

    def build_prompt_preview(self, rel_docx: str, selected_labels: list, mode: str) -> str:
        return "\n".join([
            f"Objectif: Relecture ‘{dict(MODES)[mode]}’ du document {rel_docx}.",
            f"Sections: {', '.join(selected_labels) if selected_labels else 'document entier'}",
            f"Checklist: modes/{mode}/instructions/checklist.md",
            f"Références: modes/{mode}/refs",
        ])

    def launch_analysis(self):
        if not self.copy_relpath or not self.mode: messagebox.showwarning("Info", "Sélectionnez un fichier ET un mode."); return
        selected_idx = [i for i, v in enumerate(self.section_vars) if v.get()]
        if not selected_idx and not messagebox.askyesno("Document entier", "Traiter le document entier ?"): return

        try:
            self.log("A) Préparation…")
            ts = ts_now()
            abs_copy = os.path.join(ROOT, self.copy_relpath)
            base_name = os.path.splitext(os.path.basename(abs_copy))[0]
            work_copy = os.path.join(WORK_DIR, f"{base_name}_copie_{ts}.docx")
            shutil.copy2(abs_copy, work_copy)

            session = { "timestamp": datetime.now().isoformat(), "source": self.copy_relpath, "mode": self.mode, "sections": [self.sections[i].label() for i in selected_idx], "output_dir": self.output_dir }
            with open(os.path.join(WORK_DIR, "session.json"), "w", encoding="utf-8") as f: json.dump(session, f, ensure_ascii=False, indent=2)

            prompt = self.build_prompt_preview(self.copy_relpath, session["sections"], self.mode)
            self.prompt_txt.config(state="normal"); self.prompt_txt.delete("1.0", tk.END); self.prompt_txt.insert(tk.END, prompt); self.prompt_txt.config(state="disabled")

            self.log("B) Découpage par sections…")
            cut_docx_path = filter_paragraphs_by_sections(work_copy, selected_idx, self.sections)

            self.log("C) Relecture (simulation IA)…")
            md_src = docx_to_markdown(cut_docx_path)
            revised_md, comments = generate_review(md_src, self.mode)
            csv_path = os.path.join(WORK_DIR, "commentaires.csv"); write_comments_csv(comments, csv_path)

            self.log("D) Génération du livrable final…")
            revised_docx_path = os.path.join(WORK_DIR, "rapport_revise.docx")
            markdown_to_docx(revised_md, revised_docx_path, Document(cut_docx_path))

            self.log("Insertion des commentaires…")
            base_short = (base_name[:50]).rstrip(" _-.")
            out_name = f"{base_short}_AI_commentaires_{ts_now()}.docx"
            out_docx_path = os.path.join(self.output_dir, out_name)
            
            # Étape finale : utiliser l'outil Python pour ajouter les commentaires
            add_comments_to_docx(revised_docx_path, csv_path, out_docx_path)

            self.log("Livrable prêt !")
            self.log(f"Fichier final: {out_docx_path}")
            messagebox.showinfo("Terminé", f"Livrable généré:\n{out_docx_path}")

        except Exception as e:
            self.log(f"ERREUR: {e}")
            import traceback; traceback.print_exc()
            messagebox.showerror("Erreur Critique", f"Le processus a échoué:\n{e}")

if __name__ == "__main__":
    App().mainloop()
