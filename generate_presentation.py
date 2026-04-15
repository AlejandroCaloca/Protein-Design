"""
Generate a Protein Design presentation (.pptx) using python-pptx.
Run:  python generate_presentation.py
Output: Protein_Design.pptx
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── colour palette ──────────────────────────────────────────────────────────
DARK_BLUE  = RGBColor(0x1A, 0x37, 0x5E)   # slide backgrounds / headings
MID_BLUE   = RGBColor(0x27, 0x5D, 0x9C)   # accent bars
TEAL       = RGBColor(0x00, 0xB0, 0xA0)   # highlights
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GREY = RGBColor(0xF4, 0xF6, 0xF9)
DARK_GREY  = RGBColor(0x33, 0x33, 0x33)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


# ── helpers ─────────────────────────────────────────────────────────────────

def add_filled_rect(slide, left, top, width, height, fill_rgb):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    shape.line.fill.background()
    return shape


def add_text_box(slide, text, left, top, width, height,
                 font_size=24, bold=False, color=WHITE,
                 align=PP_ALIGN.LEFT, wrap=True):
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "Calibri"
    return txBox


def add_bullet_slide(prs, title_text, bullets, bg=LIGHT_GREY):
    slide_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(slide_layout)

    # background
    add_filled_rect(slide, 0, 0, 13.33, 7.5, bg)

    # left accent bar
    add_filled_rect(slide, 0, 0, 0.18, 7.5, TEAL)

    # title band
    add_filled_rect(slide, 0.18, 0, 13.15, 1.3, DARK_BLUE)
    add_text_box(slide, title_text, 0.4, 0.15, 12.5, 1.0,
                 font_size=32, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

    # bullets
    txBox = slide.shapes.add_textbox(
        Inches(0.55), Inches(1.55), Inches(12.3), Inches(5.6)
    )
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, (level, text) in enumerate(bullets):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.level = level
        p.space_before = Pt(4)
        indent = "    " * level
        bullet_char = "▸ " if level == 0 else "◦ "
        run = p.add_run()
        run.text = indent + bullet_char + text
        run.font.size = Pt(20 if level == 0 else 17)
        run.font.bold = (level == 0)
        run.font.color.rgb = DARK_GREY if bg == LIGHT_GREY else WHITE
        run.font.name = "Calibri"

    return slide


# ── slide definitions ────────────────────────────────────────────────────────

def build_title_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # full-bleed dark background
    add_filled_rect(slide, 0, 0, 13.33, 7.5, DARK_BLUE)

    # decorative teal rectangle (bottom)
    add_filled_rect(slide, 0, 6.5, 13.33, 1.0, TEAL)

    # diagonal accent
    add_filled_rect(slide, 0, 0, 0.4, 7.5, MID_BLUE)

    # main title
    add_text_box(slide, "Protein Design", 0.7, 1.5, 11.5, 1.8,
                 font_size=60, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # subtitle
    add_text_box(slide,
                 "From Sequence to Structure to Function",
                 0.7, 3.4, 11.5, 0.9,
                 font_size=28, bold=False, color=TEAL, align=PP_ALIGN.CENTER)

    # footer
    add_text_box(slide,
                 "AlejandroCaloca / Protein-Design  •  2026",
                 0.7, 6.6, 11.5, 0.6,
                 font_size=16, bold=False, color=WHITE, align=PP_ALIGN.CENTER)


def build_outline_slide(prs):
    add_bullet_slide(prs, "Outline", [
        (0, "What is Protein Design?"),
        (0, "Key Approaches"),
        (0, "Computational Tools"),
        (0, "Deep-Learning Methods"),
        (0, "Applications"),
        (0, "Challenges & Future Directions"),
    ], bg=DARK_BLUE)


def build_what_is_slide(prs):
    add_bullet_slide(prs, "What is Protein Design?", [
        (0, "Goal: engineer proteins with desired structure, stability, or activity"),
        (1, "Inverse of the protein-folding problem"),
        (1, "Find a sequence that folds into a target 3-D shape"),
        (0, "Two broad paradigms"),
        (1, "De novo design  –  create entirely new folds not found in nature"),
        (1, "Redesign  –  modify existing proteins to alter function or stability"),
        (0, "Why it matters"),
        (1, "Therapeutic proteins, enzymes, biosensors, biomaterials"),
        (1, "Accelerates drug discovery and synthetic biology"),
    ])


def build_approaches_slide(prs):
    add_bullet_slide(prs, "Key Approaches", [
        (0, "Physics-based energy minimisation"),
        (1, "Rosetta, FoldX  –  model van-der-Waals, H-bonds, electrostatics"),
        (0, "Fragment-based assembly"),
        (1, "Combine structural fragments from the PDB to build new backbones"),
        (0, "Directed evolution (in-lab)"),
        (1, "Iterative mutagenesis + selection to improve function"),
        (0, "Machine-learning sequence design"),
        (1, "Train on known structure–sequence pairs; sample new sequences"),
    ])


def build_tools_slide(prs):
    add_bullet_slide(prs, "Computational Tools", [
        (0, "Rosetta Suite"),
        (1, "RosettaDesign, RosettaFold, FastRelax, Enzyme Design"),
        (0, "AlphaFold2 / ColabFold"),
        (1, "High-accuracy structure prediction used to validate designs"),
        (0, "PyMOL / ChimeraX"),
        (1, "Visualisation and manual editing of protein structures"),
        (0, "Modeller, GROMACS, OpenMM"),
        (1, "Homology modelling and molecular-dynamics refinement"),
    ], bg=DARK_BLUE)


def build_dl_slide(prs):
    add_bullet_slide(prs, "Deep-Learning Methods", [
        (0, "ProteinMPNN"),
        (1, "Graph neural network for fixed-backbone sequence design"),
        (0, "ESM (Evolutionary Scale Modelling)  –  Meta AI"),
        (1, "Protein language model; ESM-IF1 for inverse folding"),
        (0, "RFdiffusion"),
        (1, "Diffusion model for de novo backbone generation"),
        (0, "ProteinGenerator / Chroma"),
        (1, "End-to-end diffusion over sequence + structure"),
        (0, "LigandMPNN"),
        (1, "Extends ProteinMPNN to account for small-molecule context"),
    ])


def build_applications_slide(prs):
    add_bullet_slide(prs, "Applications", [
        (0, "Therapeutics"),
        (1, "Designed binders for viral antigens (e.g., SARS-CoV-2 RBD)"),
        (1, "Miniproteins as potential antivirals"),
        (0, "Enzymes & Biocatalysis"),
        (1, "Retro-aldolase, Kemp eliminase  –  reactions absent in nature"),
        (0, "Biosensors"),
        (1, "Fluorescent biosensors by embedding chromophore-binding sites"),
        (0, "Biomaterials"),
        (1, "Self-assembling protein cages, hydrogels, scaffolds"),
        (0, "Synthetic Biology"),
        (1, "Orthogonal translation systems, logic-gate proteins"),
    ], bg=DARK_BLUE)


def build_challenges_slide(prs):
    add_bullet_slide(prs, "Challenges & Future Directions", [
        (0, "Experimental validation remains expensive and slow"),
        (1, "High-throughput assays (PACE, cell-surface display) are helping"),
        (0, "Accurately designing dynamics & allostery"),
        (1, "Most methods optimise static structure, not conformational change"),
        (0, "Incorporating non-standard amino acids"),
        (1, "Expanding the chemical alphabet for new function"),
        (0, "Interpretability of deep-learning models"),
        (1, "Understanding why a sequence folds as predicted"),
        (0, "Future outlook"),
        (1, "Fully automated design-build-test cycles"),
        (1, "Foundation models fine-tuned on wet-lab feedback loops"),
    ])


def build_summary_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_filled_rect(slide, 0, 0, 13.33, 7.5, DARK_BLUE)
    add_filled_rect(slide, 0, 0, 0.4, 7.5, TEAL)
    add_filled_rect(slide, 0, 6.5, 13.33, 1.0, MID_BLUE)

    add_text_box(slide, "Summary", 0.7, 0.8, 11.5, 1.0,
                 font_size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    points = [
        "Protein design bridges sequence, structure, and function",
        "Physics-based and ML approaches are increasingly complementary",
        "Deep-learning tools (RFdiffusion, ProteinMPNN) are transforming the field",
        "Experimental validation and interpretability remain open challenges",
    ]
    for i, pt in enumerate(points):
        add_text_box(slide, f"✔  {pt}",
                     0.8, 2.0 + i * 1.1, 11.5, 0.9,
                     font_size=22, bold=False, color=WHITE, align=PP_ALIGN.LEFT)

    add_text_box(slide, "Thank you  |  github.com/AlejandroCaloca/Protein-Design",
                 0.7, 6.6, 11.5, 0.6,
                 font_size=16, bold=False, color=WHITE, align=PP_ALIGN.CENTER)


# ── main ────────────────────────────────────────────────────────────────────

def main():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    build_title_slide(prs)
    build_outline_slide(prs)
    build_what_is_slide(prs)
    build_approaches_slide(prs)
    build_tools_slide(prs)
    build_dl_slide(prs)
    build_applications_slide(prs)
    build_challenges_slide(prs)
    build_summary_slide(prs)

    out = "Protein_Design.pptx"
    prs.save(out)
    print(f"Saved → {out}  ({prs.slides.__len__()} slides)")


if __name__ == "__main__":
    main()
