import os
import re
from pathlib import Path
from datetime import datetime

from docx import Document

# Importer les fonctions utilitaires depuis Start.py (sans lancer la GUI)
import sys
sys.path.append(str(Path(__file__).resolve().parents[1]))
import Start as start_mod


ROOT = Path(__file__).resolve().parents[1]
INPUT_REL = Path("input/prédiag_DECOUPE.docx")
OUTPUT_DIR = ROOT / "output"
OUTPUT_NAME = "prédiag_AI_suivi+commentaires.docx"


def load_md_from_docx(docx_path: Path) -> str:
    return start_mod.docx_to_markdown(str(docx_path))


def rewrite_line(text: str) -> str:
    """Applique une normalisation rédactionnelle prudente (non intrusive)."""
    s = start_mod.sanitize_text(text)
    # Retouches légères de style qui ne changent pas le fond
    s = s.replace("Cette analyse repose sur un passage de terrain en date du", "Cette analyse repose sur un passage de terrain réalisé le")
    s = s.replace("Le présent chapitre", "Ce pré-diagnostic")
    s = s.replace("prè-diagnostic", "pré-diagnostic")
    # Harmoniser certaines graphies usuelles
    s = re.sub(r"\b11\s*/\s*06\s*/\s*2025\b", "11/06/2025", s)
    return s


def build_revised_md_and_comments(md_text: str):
    lines = md_text.splitlines()
    revised_lines = []
    comments = []

    def add_comment(anchor: str, text: str, gravite: str = "P3", categorie: str = "redaction"):
        if not anchor:
            return
        comments.append({
            "ancre_textuelle": anchor[:120],
            "commentaire": text,
            "gravite": gravite,
            "categorie": categorie,
        })

    for raw in lines:
        if not raw.strip():
            revised_lines.append("")
            continue
        if raw.startswith("#"):
            # Conserver les titres en nettoyant légèrement
            parts = raw.split(" ", 1)
            title = start_mod.sanitize_text(parts[1] if len(parts) > 1 else raw.lstrip("#"))
            lvl = len(parts[0])
            revised_lines.append("#" * lvl + " " + title)
            continue

        original = raw
        revised = rewrite_line(original)
        revised_lines.append(revised)

        # Créer un commentaire explicite pour toute réécriture
        if revised != original:
            add_comment(
                anchor=start_mod.sanitize_text(original)[:80],
                text=f"Proposition de reformulation: {revised}",
                gravite="P3",
                categorie="redaction",
            )

        o_low = original.lower()

        # Heuristiques de checklist Diagnostic
        if "données bibliograph" in o_low or "bibliograph" in o_low:
            add_comment(
                anchor=start_mod.sanitize_text(original)[:80],
                text=("Vérifier l'actualisation et la citation des sources (auteur, millésime). "
                      "Proposition de reformulation: préciser la source (ex.: INPN/SINP, CLC millésime, IGN)."),
                gravite="P2", categorie="coherence",
            )

        if re.search(r"\b1971\s*-\s*2000\b|\b1991\s*-\s*2020\b", original):
            add_comment(
                anchor=start_mod.sanitize_text(original)[:80],
                text=("Préciser la source climat (ex.: Météo-France/Drias), la station/résolution et la méthode de calcul. "
                      "Proposition de reformulation: ajouter la référence et le millésime des données."),
                gravite="P2", categorie="methodologie",
            )

        if "figure" in o_low and not re.search(r"figure\s*\d+", o_low):
            add_comment(
                anchor=start_mod.sanitize_text(original)[:80],
                text=("Figure non numérotée ou référence incomplète. "
                      "Proposition de reformulation: numéroter systématiquement les figures et harmoniser les appels dans le texte."),
                gravite="P3", categorie="carto",
            )

        if "corine land cover" in o_low or re.search(r"\bclc\b", o_low):
            add_comment(
                anchor=start_mod.sanitize_text(original)[:80],
                text=("Préciser le millésime CLC (ex.: CLC2018/CLC2012), la méthode (photo-interprétation), et les limites d'usage à l'échelle du site. "
                      "Proposition de reformulation: compléter la référence (auteur, année, URL)."),
                gravite="P2", categorie="coherence",
            )

        if "enjeu" in o_low and ("habitat" in o_low or "espèce" in o_low or "espèces" in o_low or "avifaune" in o_low):
            add_comment(
                anchor=start_mod.sanitize_text(original)[:80],
                text=("Justifier le niveau d'enjeu (critères: statut légal/Liste rouge, effectifs/aires, fonctionnalité, rareté locale). "
                      "Proposition de reformulation: expliciter le critère au premier appel et renvoyer au tableau de synthèse."),
                gravite="P2", categorie="coherence",
            )

        if re.search(r"\([A-Z][a-z]+\s+[a-z]{2,}\)", original):
            add_comment(
                anchor=start_mod.sanitize_text(original)[:80],
                text=("Vérifier l'italique des noms scientifiques et l'orthographe. "
                      "Proposition de reformulation: italique pour le binôme latin à chaque première occurrence."),
                gravite="P3", categorie="redaction",
            )

        if any(k in o_low for k in ["zone d'étude", "périmètre", "fenêtre phéno", "phénolog"]):
            add_comment(
                anchor=start_mod.sanitize_text(original)[:80],
                text=("Cadrage spatial/temporal: confirmer limites (ZI/ZR/ZE), calendrier par groupe (fenêtres phénologiques) et exceptions justifiées. "
                      "Proposition de reformulation: préciser les dates et périmètres d'analyse."),
                gravite="P2", categorie="methodologie",
            )

    # Ajouter des commentaires de couverture de checklist (appel aux fonctions Start existantes)
    md_full = "\n".join(lines)
    revised_md_full = "\n".join(revised_lines)
    _, checklist_comments = start_mod.generate_review(md_full, "diagnostic")
    comments.extend(checklist_comments)

    return "\n".join(revised_lines) + "\n", comments


def ensure_word_compare_with_highlight(original_docx: Path, revised_docx: Path, comments_csv: Path, output_docx: Path):
    # Utilise le script PowerShell fourni (compare + insert comments + highlight vert)
    start_mod.run_compare_and_comment_ps(str(original_docx), str(revised_docx), str(comments_csv), str(output_docx))


def main():
    # 1) Validation explicite du chemin de travail
    src = ROOT / INPUT_REL
    if not src.exists():
        raise SystemExit(f"ERREUR: Fichier de travail introuvable: {INPUT_REL}")

    print(f"OK: utilisation EXCLUSIVE du fichier: {INPUT_REL}")

    # 2) Conversion DOCX -> Markdown
    md_src = load_md_from_docx(src)

    # 3) Réécriture prudente + génération des commentaires (mode diagnostic)
    revised_md, comments = build_revised_md_and_comments(md_src)

    # 4) MD -> DOCX (conserver les styles du document d'entrée autant que possible)
    ref_doc = Document(str(src))
    work_dir = ROOT / "work"
    work_dir.mkdir(exist_ok=True)
    revised_docx = work_dir / "_revised.docx"
    start_mod.markdown_to_docx(revised_md, str(revised_docx), ref_doc)

    # 5) CSV de commentaires
    comments_csv = work_dir / "_comments.csv"
    start_mod.write_comments_csv(comments, str(comments_csv))

    # 6) Comparaison sous Word + surlignage vert des insertions + insertion des commentaires
    output_path = OUTPUT_DIR / OUTPUT_NAME
    OUTPUT_DIR.mkdir(exist_ok=True)
    ensure_word_compare_with_highlight(src, revised_docx, comments_csv, output_path)

    print(f"Livrable généré: {output_path}")


if __name__ == "__main__":
    main()
