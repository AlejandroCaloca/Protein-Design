"""
Generate a Protein Design presentation (.pptx) using python-pptx.
Styling follows the Georgetown University corporate template.

Run:    python generate_presentation.py
Output: Protein_Design.pptx  (17 slides)
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── Georgetown University brand colours ────────────────────────────────────
GU_BLUE   = RGBColor(0x04, 0x1E, 0x42)   # Georgetown Navy Blue  (#041E42)
GU_GOLD   = RGBColor(0xC8, 0xA9, 0x51)   # Georgetown Gold       (#C8A951)
GU_GREY   = RGBColor(0x53, 0x56, 0x5A)   # Georgetown Grey       (#53565A)
GU_LGREY  = RGBColor(0xF2, 0xF2, 0xF2)   # Light background grey
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
DARK_TEXT = RGBColor(0x1A, 0x1A, 0x1A)

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)


# ── low-level helpers ────────────────────────────────────────────────────────

def blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def rect(slide, left, top, width, height, fill_rgb, line=False):
    shape = slide.shapes.add_shape(
        1, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_rgb
    if line:
        shape.line.color.rgb = fill_rgb
    else:
        shape.line.fill.background()
    return shape


def tbox(slide, text, left, top, width, height,
         size=20, bold=False, color=WHITE, align=PP_ALIGN.LEFT):
    tb = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = "Georgia" if bold else "Calibri"
    return tb


# ── Georgetown slide chrome ──────────────────────────────────────────────────

def gu_chrome(slide, light_bg=False):
    """Apply Georgetown template chrome: background, top bar, gold rule."""
    bg = GU_LGREY if light_bg else GU_BLUE
    rect(slide, 0, 0, 13.33, 7.5, bg)
    # top navy band (only needed on light slides)
    if light_bg:
        rect(slide, 0, 0, 13.33, 1.25, GU_BLUE)
    # bottom gold rule
    rect(slide, 0, 7.25, 13.33, 0.25, GU_GOLD)


def gu_slide_number(slide, num):
    tbox(slide, str(num), 12.8, 7.1, 0.4, 0.3,
         size=11, bold=False, color=GU_GREY, align=PP_ALIGN.RIGHT)


def gu_title_bar(slide, title_text, light_bg=False):
    """Place a slide title inside the top bar."""
    fg = WHITE
    tbox(slide, title_text, 0.35, 0.15, 12.5, 1.0,
         size=30, bold=True, color=fg, align=PP_ALIGN.LEFT)


def gu_bullet_slide(prs, title_text, bullets, slide_num, light_bg=True):
    """Create a content slide with Georgetown chrome and bullet list."""
    slide = blank_slide(prs)
    gu_chrome(slide, light_bg=light_bg)
    gu_title_bar(slide, title_text, light_bg)
    gu_slide_number(slide, slide_num)

    # content area
    tb = slide.shapes.add_textbox(
        Inches(0.45), Inches(1.45), Inches(12.4), Inches(5.65)
    )
    tf = tb.text_frame
    tf.word_wrap = True

    text_color = DARK_TEXT if light_bg else WHITE

    for i, (level, text) in enumerate(bullets):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.level = level
        p.space_before = Pt(5 if level == 0 else 3)
        bullet_char = "▸  " if level == 0 else "◦  "
        indent = "       " * level
        r = p.add_run()
        r.text = indent + bullet_char + text
        r.font.size = Pt(19 if level == 0 else 16)
        r.font.bold = (level == 0)
        r.font.color.rgb = text_color
        r.font.name = "Georgia" if (level == 0 and light_bg) else "Calibri"

    return slide


# ── individual slide builders ────────────────────────────────────────────────

# Slide 1 – Title
def s01_title(prs):
    slide = blank_slide(prs)
    rect(slide, 0, 0, 13.33, 7.5, GU_BLUE)
    rect(slide, 0, 5.8, 13.33, 0.08, GU_GOLD)   # gold divider
    rect(slide, 0, 7.25, 13.33, 0.25, GU_GOLD)   # bottom rule

    tbox(slide, "GEORGETOWN UNIVERSITY", 0.6, 0.4, 11.5, 0.5,
         size=14, bold=False, color=GU_GOLD, align=PP_ALIGN.LEFT)
    tbox(slide, "Protein Design", 0.6, 1.4, 11.5, 2.0,
         size=62, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    tbox(slide, "From Sequence to Structure to Function",
         0.6, 3.6, 11.5, 0.8,
         size=26, bold=False, color=GU_GOLD, align=PP_ALIGN.LEFT)
    tbox(slide,
         "Department of Biochemistry & Molecular Biology  |  2026",
         0.6, 6.85, 11.5, 0.45,
         size=13, bold=False, color=WHITE, align=PP_ALIGN.LEFT)


# Slide 2 – Agenda / Outline
def s02_outline(prs):
    gu_bullet_slide(prs, "Agenda", [
        (0, "What is Protein Design?"),
        (0, "The Protein Folding Problem"),
        (0, "Key Approaches — Overview"),
        (0, "Physics-Based Energy Minimisation"),
        (0, "Fragment-Based Assembly & Directed Evolution"),
        (0, "Computational Tools"),
        (0, "Deep-Learning Methods Overview"),
        (0, "ProteinMPNN & RFdiffusion"),
        (0, "AlphaFold2, ESMFold & LigandMPNN"),
        (0, "Applications — Therapeutics"),
        (0, "Applications — Enzymes & Biocatalysis"),
        (0, "Applications — Biosensors, Biomaterials & Synbio"),
        (0, "Challenges"),
        (0, "Future Directions"),
        (0, "Sources"),
    ], slide_num=2, light_bg=False)


# Slide 3 – What is Protein Design?
def s03_what_is(prs):
    gu_bullet_slide(prs, "What is Protein Design?", [
        (0, "Goal: engineer proteins with a desired structure, stability, or activity"),
        (1, "Inverse of the protein-folding problem"),
        (1, "Find a sequence that folds reliably into a target 3-D shape"),
        (0, "Two broad paradigms"),
        (1, "De novo design — create entirely new folds not found in nature"),
        (1, "Redesign — modify existing proteins to alter function or stability"),
        (0, "Why it matters"),
        (1, "Therapeutic proteins, industrial enzymes, biosensors, biomaterials"),
        (1, "Accelerates drug discovery and synthetic biology pipelines"),
    ], slide_num=3)


# Slide 4 – The Protein Folding Problem
def s04_folding(prs):
    gu_bullet_slide(prs, "The Protein Folding Problem", [
        (0, "Anfinsen's dogma (1973): sequence determines structure"),
        (1, "Nobel Prize in Chemistry 1972 for establishing this principle"),
        (0, "Levinthal's paradox: random search of conformational space is impossible"),
        (1, "A 100-aa protein would take longer than the age of the universe to fold randomly"),
        (1, "Proteins must follow guided folding pathways"),
        (0, "Structure → Function relationship"),
        (1, "α-helices, β-sheets, and loops create binding pockets, active sites"),
        (0, "Design inverts this: target function → desired structure → sequence"),
    ], slide_num=4)


# Slide 5 – Key Approaches Overview
def s05_approaches_overview(prs):
    gu_bullet_slide(prs, "Key Approaches — Overview", [
        (0, "Physics-based methods"),
        (1, "Model atomic interactions via energy functions to score/design sequences"),
        (0, "Fragment-based assembly"),
        (1, "Assemble new backbones from experimentally observed structural fragments"),
        (0, "Directed evolution"),
        (1, "Iterative in-lab mutagenesis and selection — no structural model needed"),
        (0, "Machine-learning / deep-learning methods"),
        (1, "Train on structure–sequence pairs from the PDB; sample new sequences"),
        (0, "Hybrid approaches"),
        (1, "Combine energy-based validation with ML-generated candidates"),
    ], slide_num=5, light_bg=False)


# Slide 6 – Physics-Based Energy Minimisation
def s06_physics(prs):
    gu_bullet_slide(prs, "Physics-Based Energy Minimisation", [
        (0, "Rosetta energy function (REF15)"),
        (1, "Van der Waals, electrostatics, solvation, H-bonds, torsion terms"),
        (1, "Used in RosettaDesign, FastRelax, Enzyme Design"),
        (0, "FoldX"),
        (1, "Rapid empirical ΔΔG estimates for point mutations"),
        (0, "Design workflow"),
        (1, "Fix backbone → optimise side-chain rotamers → minimise energy"),
        (1, "Monte Carlo sampling or gradient-based minimisation"),
        (0, "Limitations"),
        (1, "Computationally expensive; energy function approximations can fail"),
    ], slide_num=6)


# Slide 7 – Fragment-Based Assembly & Directed Evolution
def s07_fragment_direvo(prs):
    gu_bullet_slide(prs, "Fragment-Based Assembly & Directed Evolution", [
        (0, "Fragment-based assembly"),
        (1, "Mine the PDB for 3-9 residue structural fragments"),
        (1, "Assemble fragments to build novel topologies (Rosetta ab initio)"),
        (1, "Enabled design of non-natural TIM-barrel and beta-barrel folds"),
        (0, "Directed evolution"),
        (1, "Error-prone PCR or DNA shuffling to introduce mutations"),
        (1, "Screen or select for improved activity/stability"),
        (1, "Frances Arnold: Nobel Prize 2018 for directed evolution of enzymes"),
        (0, "Combining both"),
        (1, "Computationally propose scaffolds; wet-lab evolution to fine-tune"),
    ], slide_num=7, light_bg=False)


# Slide 8 – Computational Tools
def s08_tools(prs):
    gu_bullet_slide(prs, "Computational Tools", [
        (0, "Rosetta Suite  (rosettacommons.org)"),
        (1, "RosettaDesign, RosettaFold2, FastRelax, RosettaScript pipelines"),
        (0, "AlphaFold2 / ColabFold"),
        (1, "High-accuracy structure prediction used to validate designs in silico"),
        (0, "ESMFold (Meta AI)"),
        (1, "Language-model-based folding; ~1000× faster than AlphaFold2"),
        (0, "PyMOL / UCSF ChimeraX"),
        (1, "Visualisation, manual mutagenesis, figure generation"),
        (0, "GROMACS / OpenMM / AMBER"),
        (1, "Molecular dynamics — assess stability and dynamics of designs"),
        (0, "Modeller"),
        (1, "Homology modelling to build initial structural templates"),
    ], slide_num=8)


# Slide 9 – Deep-Learning Methods Overview
def s09_dl_overview(prs):
    gu_bullet_slide(prs, "Deep-Learning Methods — Overview", [
        (0, "Why deep learning?"),
        (1, "PDB contains >220 000 structures — rich training data"),
        (1, "Models learn implicit rules of protein stability and interaction"),
        (0, "Protein language models (PLMs)"),
        (1, "Trained on millions of sequences; capture evolutionary constraints"),
        (1, "ESM-2, ProtTrans, Ankh"),
        (0, "Structure-conditioned models"),
        (1, "Take 3-D coordinates as input; predict optimal sequences"),
        (0, "Generative models"),
        (1, "Hallucinate or diffuse entirely new backbones and sequences"),
        (0, "Key advantage over physics"),
        (1, "Orders of magnitude faster; capture patterns hard to encode analytically"),
    ], slide_num=9, light_bg=False)


# Slide 10 – ProteinMPNN & RFdiffusion
def s10_mpnn_rfdiff(prs):
    gu_bullet_slide(prs, "ProteinMPNN & RFdiffusion", [
        (0, "ProteinMPNN  (Dauparas et al., Science 2022)"),
        (1, "Message-passing graph neural network over protein backbone atoms"),
        (1, "Given fixed backbone coordinates → outputs sequence probabilities"),
        (1, "Outperforms Rosetta in experimental success rates"),
        (1, "Variants: soluble-MPNN, LigandMPNN, ProteinMPNN-SSM"),
        (0, "RFdiffusion  (Watson et al., Nature 2023)"),
        (1, "Diffusion model built on RoseTTAFold architecture"),
        (1, "Generates de novo backbone conditioned on user-specified motifs"),
        (1, "Applications: binders, symmetric assemblies, enzyme active sites"),
        (1, "Combined with ProteinMPNN for end-to-end design pipelines"),
    ], slide_num=10)


# Slide 11 – AlphaFold2, ESMFold & LigandMPNN
def s11_af2_esm_ligmpnn(prs):
    gu_bullet_slide(prs, "AlphaFold2, ESMFold & LigandMPNN", [
        (0, "AlphaFold2  (Jumper et al., Nature 2021)"),
        (1, "Transformer + Evoformer; predicts structure with near-experimental accuracy"),
        (1, "Used as oracle to validate designed sequences in silico"),
        (1, "AlphaFold3 extends to protein–DNA/RNA–ligand complexes (2024)"),
        (0, "ESMFold  (Lin et al., Science 2023)"),
        (1, "Single-sequence folding from ESM-2 language model"),
        (1, "Enables rapid screening of millions of designed sequences"),
        (0, "LigandMPNN  (Dauparas et al., Nature Methods 2024)"),
        (1, "Extends ProteinMPNN to include small-molecule and nucleic-acid context"),
        (1, "Enables enzyme and drug-binding site design"),
    ], slide_num=11, light_bg=False)


# Slide 12 – Applications: Therapeutics
def s12_therapeutics(prs):
    gu_bullet_slide(prs, "Applications — Therapeutics", [
        (0, "Designed miniprotein binders"),
        (1, "de novo binders to SARS-CoV-2 spike RBD (Cao et al., Science 2020)"),
        (1, "Computationally designed influenza inhibitors"),
        (0, "Antibody engineering"),
        (1, "Stability engineering of clinical antibodies via Rosetta"),
        (1, "Bispecific antibodies with redesigned interfaces"),
        (0, "Protein-based vaccines"),
        (1, "Self-assembling nanoparticle immunogens displaying multiple antigens"),
        (0, "Therapeutic enzymes"),
        (1, "PEGylated asparaginase for ALL; redesigned for reduced immunogenicity"),
        (0, "Cell-penetrating proteins"),
        (1, "Designed to cross membranes for intracellular delivery"),
    ], slide_num=12)


# Slide 13 – Applications: Enzymes & Biocatalysis
def s13_enzymes(prs):
    gu_bullet_slide(prs, "Applications — Enzymes & Biocatalysis", [
        (0, "De novo enzyme design"),
        (1, "Retro-aldolase: computationally designed active site on TIM-barrel scaffold"),
        (1, "Kemp eliminase: model C–H acid chemistry not found in natural enzymes"),
        (0, "Industrial biocatalysis"),
        (1, "Directed evolution of P450 enzymes for pharmaceutical synthesis"),
        (1, "Thermostable cellulases for biomass degradation"),
        (0, "Computational enzyme redesign"),
        (1, "Switch substrate specificity or stereoselectivity via active-site mutations"),
        (0, "Metalloenzyme design"),
        (1, "Introduce non-native metal cofactors (Mn, Fe, Cu) for new reactions"),
    ], slide_num=13, light_bg=False)


# Slide 14 – Applications: Biosensors, Biomaterials & Synbio
def s14_biosensors_mats(prs):
    gu_bullet_slide(prs, "Applications — Biosensors, Biomaterials & Synbio", [
        (0, "Fluorescent biosensors"),
        (1, "Designed binding pockets that shift chromophore fluorescence on ligand binding"),
        (1, "dLight, GRAB sensors for in vivo neurotransmitter imaging"),
        (0, "Self-assembling biomaterials"),
        (1, "Symmetric protein cages (T33-09) for drug delivery"),
        (1, "Hydrogels, fibres, and lattices from designed coiled-coil proteins"),
        (0, "Synthetic biology"),
        (1, "Orthogonal translation systems with non-canonical amino acids"),
        (1, "Logic-gate proteins that respond to multiple input signals"),
        (1, "Designed protein–protein interaction networks as synthetic circuits"),
    ], slide_num=14)


# Slide 15 – Challenges
def s15_challenges(prs):
    gu_bullet_slide(prs, "Challenges", [
        (0, "Experimental validation remains the bottleneck"),
        (1, "Synthesis, expression, purification, and biophysical characterisation are slow"),
        (1, "High-throughput display methods (yeast, phage, PACE) are improving throughput"),
        (0, "Designing protein dynamics and allostery"),
        (1, "Most methods optimise the ground-state structure, not conformational change"),
        (0, "Non-standard amino acids and post-translational modifications"),
        (1, "Expanding chemical space beyond the canonical 20 AAs"),
        (0, "Model interpretability"),
        (1, "Deep-learning models are largely black boxes — hard to extract design rules"),
        (0, "Off-target and immunogenicity risks"),
        (1, "Designed sequences may trigger immune responses in therapeutic contexts"),
    ], slide_num=15, light_bg=False)


# Slide 16 – Future Directions
def s16_future(prs):
    gu_bullet_slide(prs, "Future Directions", [
        (0, "Fully automated design–build–test (DBT) cycles"),
        (1, "Robotic labs + ML feedback loops to iterate in days, not months"),
        (0, "Foundation models for proteins"),
        (1, "Large-scale pre-training on sequence, structure, and function data"),
        (1, "Fine-tuning on wet-lab readouts (affinity, stability, activity)"),
        (0, "Multi-state and dynamic design"),
        (1, "Explicitly optimise ensembles of conformations, not a single structure"),
        (0, "Multi-modal design"),
        (1, "Co-design of protein–small molecule, protein–RNA, protein–lipid systems"),
        (0, "Expanding the genetic code"),
        (1, "200+ non-canonical AAs available via amber suppression and genetic code expansion"),
        (0, "Democratisation"),
        (1, "Cloud-based tools (ColabDesign, EvolutionaryScale API) lower barriers"),
    ], slide_num=16)


# Slide 17 – Sources
def s17_sources(prs):
    slide = blank_slide(prs)
    gu_chrome(slide, light_bg=True)
    gu_title_bar(slide, "Sources")
    gu_slide_number(slide, 17)

    sources = [
        "1.  Anfinsen, C.B. (1973). Principles that govern the folding of protein chains. Science, 181(4096), 223–230.",
        "2.  Jumper, J. et al. (2021). Highly accurate protein structure prediction with AlphaFold. Nature, 596, 583–589.",
        "3.  Dauparas, J. et al. (2022). Robust deep learning-based protein sequence design using ProteinMPNN. Science, 378(6615), 49–56.",
        "4.  Watson, J.L. et al. (2023). De novo design of protein structure and function with RFdiffusion. Nature, 620, 1089–1100.",
        "5.  Lin, Z. et al. (2023). Evolutionary-scale prediction of atomic-level protein structure with a language model. Science, 379(6637), 1123–1130.",
        "6.  Dauparas, J. et al. (2024). Atomic context-conditioned protein sequence design using LigandMPNN. Nature Methods, 21, 1121–1131.",
        "7.  Cao, L. et al. (2020). De novo design of picomolar SARS-CoV-2 miniprotein inhibitors. Science, 370(6515), 426–431.",
        "8.  Arnold, F.H. (2018). Directed Evolution: Bringing New Chemistry to Life. Nobel Lecture. Angew. Chem. Int. Ed., 58, 14420–14426.",
        "9.  Jiang, L. et al. (2008). De novo computational design of retro-aldol enzymes. Science, 319(5868), 1387–1391.",
        "10. Röthlisberger, D. et al. (2008). Kemp elimination catalysts by computational enzyme design. Nature, 453, 190–195.",
    ]

    tb = slide.shapes.add_textbox(
        Inches(0.45), Inches(1.45), Inches(12.4), Inches(5.8)
    )
    tf = tb.text_frame
    tf.word_wrap = True

    for i, src in enumerate(sources):
        p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
        p.space_before = Pt(4)
        r = p.add_run()
        r.text = src
        r.font.size = Pt(13)
        r.font.bold = False
        r.font.color.rgb = DARK_TEXT
        r.font.name = "Calibri"


# ── main ────────────────────────────────────────────────────────────────────

def main():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    s01_title(prs)
    s02_outline(prs)
    s03_what_is(prs)
    s04_folding(prs)
    s05_approaches_overview(prs)
    s06_physics(prs)
    s07_fragment_direvo(prs)
    s08_tools(prs)
    s09_dl_overview(prs)
    s10_mpnn_rfdiff(prs)
    s11_af2_esm_ligmpnn(prs)
    s12_therapeutics(prs)
    s13_enzymes(prs)
    s14_biosensors_mats(prs)
    s15_challenges(prs)
    s16_future(prs)
    s17_sources(prs)

    out = "Protein_Design.pptx"
    prs.save(out)
    print(f"Saved → {out}  ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
