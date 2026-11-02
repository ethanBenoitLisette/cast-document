"""Génère un document de suivi de tournoi vierge (.docx), sur le modèle du
document « PIVOT'GRAB #1 ».

Usage : python vierge.py
Dépendance : pip install python-docx

À l'ouverture dans Word, accepter la mise à jour des champs pour générer
le sommaire et la numérotation des pages.
"""

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# --- Configuration ----------------------------------------------------------

TOURNOI = "NOM DU TOURNOI #X"
FICHIER_SORTIE = "Document de suivi - VIERGE.docx"
NB_POOLS = 4
LIGNES_PAR_POOL = 26
LIGNES_ABREVIATIONS = 15
NB_HEAD_TO_HEADS = 4
LIGNES_PAR_HEAD_TO_HEAD = 4

POLICE = "Aptos"
BLEU_TITRE = RGBColor(0x0F, 0x47, 0x61)
BLEU_PAGE = RGBColor(0x4A, 0x90, 0xD9)
ORANGE = "ED7D31"
GRIS = "767171"
JAUNE = "FFFF00"

RANKINGS = [
    ("LUMI", "LumiRank", "FRW", "West France Power Rankings – Ouest de la France"),
    ("EUPR", "LumiRank Europe", "FR CVL", "Centre-Val de Loire Power Rankings"),
    ("FR", "France Power Rankings", "FR NAQ", "Nouvelle Aquitaine Power Rankings"),
    ("FR TLS", "Toulouse Power Rankings", "FR ARA", "Auvergne-Rhône-Alpes Power Rankings"),
    ("FR OCC", "Occitanie Power Rankings", "HM", "Honorable Mention"),
]

LIENS_UTILES = [
    ("ULTRANK TOURNAMENT TIERER par @KENNIKY1", "https://kenniky.github.io/web_ultrank_scorer.html"),
    ("SUPERMAJOR", "https://www.supermajor.gg/"),
    ("ULTIMATE FRAME DATA", "https://ultimateframedata.com/smash"),
    ("SMASH GUIDE – GUIDE COMPLET POUR SUPER SMASH BROS. ULTIMATE par EYDOU & ERAMTH GIN", "https://smash-guide.com/"),
    ("THE FIGHTING GAME GLOSSARY", "https://glossary.infil.net/"),
    ("TERMINOLOGY | SMASHPEDIA | FANDOM", "https://supersmashbros.fandom.com/wiki/Terminology"),
    ("SMASHBROTHERS – STAGES (AT)", "https://smashbrothers.at/en/rules/stages/"),
    ("SUPER SMASH BROS. ULTIMATE STAGE COMPARISON TOOL – TOURNAMETA", "https://tournameta.com/stage-comparison/"),
]


# --- Helpers XML ------------------------------------------------------------

def ajouter_champ(paragraphe, instruction, texte_defaut=""):
    """Insère un champ Word (PAGE, NUMPAGES, TOC...) dans un paragraphe."""
    def element(tag, **attrs):
        el = OxmlElement(tag)
        for cle, valeur in attrs.items():
            el.set(qn(cle), valeur)
        return el

    run = paragraphe.add_run()
    run._r.append(element("w:fldChar", **{"w:fldCharType": "begin", "w:dirty": "true"}))
    instr = element("w:instrText", **{"xml:space": "preserve"})
    instr.text = f" {instruction} "
    run._r.append(instr)
    run._r.append(element("w:fldChar", **{"w:fldCharType": "separate"}))
    run = paragraphe.add_run(texte_defaut)
    paragraphe.add_run()._r.append(element("w:fldChar", **{"w:fldCharType": "end"}))
    return run


def fond_cellule(cellule, couleur):
    tc_pr = cellule._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), couleur)
    tc_pr.append(shd)


def surlignage(run, couleur=JAUNE):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), couleur)
    run._r.get_or_add_rPr().append(shd)


def espacement_lettres(run, twips):
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:val"), str(twips))
    run._r.get_or_add_rPr().append(spacing)


def bordure_gauche(paragraphe, couleur="156082"):
    p_bdr = OxmlElement("w:pBdr")
    gauche = OxmlElement("w:left")
    for cle, valeur in (("val", "single"), ("sz", "18"), ("space", "6"), ("color", couleur)):
        gauche.set(qn(f"w:{cle}"), valeur)
    p_bdr.append(gauche)
    paragraphe._p.get_or_add_pPr().append(p_bdr)


def bordure_page(section):
    pg_borders = OxmlElement("w:pgBorders")
    pg_borders.set(qn("w:offsetFrom"), "page")
    for cote in ("top", "left", "bottom", "right"):
        bord = OxmlElement(f"w:{cote}")
        for cle, valeur in (("val", "single"), ("sz", "4"), ("space", "24"), ("color", "auto")):
            bord.set(qn(f"w:{cle}"), valeur)
        pg_borders.append(bord)
    section._sectPr.append(pg_borders)


def lien_hypertexte(paragraphe, url, texte):
    r_id = paragraphe.part.relate_to(
        url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True
    )
    lien = OxmlElement("w:hyperlink")
    lien.set(qn("r:id"), r_id)
    run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    style = OxmlElement("w:rStyle")
    style.set(qn("w:val"), "Hyperlink")
    r_pr.append(style)
    run.append(r_pr)
    t = OxmlElement("w:t")
    t.text = texte
    run.append(t)
    lien.append(run)
    paragraphe._p.append(lien)


# --- Helpers de mise en page ------------------------------------------------

def configurer_styles(doc):
    normal = doc.styles["Normal"]
    normal.font.name = POLICE
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(6)

    for nom, taille in (("Heading 1", 18), ("Heading 2", 15), ("Heading 3", 13)):
        style = doc.styles[nom]
        style.font.name = POLICE
        # Les polices du thème sont prioritaires sur font.name : on les retire
        r_fonts = style.element.rPr.rFonts
        for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
            r_fonts.attrib.pop(qn(attr), None)
        style.font.size = Pt(taille)
        style.font.bold = False
        style.font.italic = False
        style.font.all_caps = True
        style.font.color.rgb = BLEU_TITRE
        style.paragraph_format.space_before = Pt(12)
        style.paragraph_format.space_after = Pt(6)

    lien = doc.styles.add_style("Hyperlink", 2)
    lien.font.color.rgb = RGBColor(0x46, 0x78, 0x86)
    lien.font.underline = True

    # Demande à Word de mettre à jour le sommaire et les numéros à l'ouverture
    maj = OxmlElement("w:updateFields")
    maj.set(qn("w:val"), "true")
    doc.settings.element.append(maj)


def configurer_section(section):
    section.page_height = Cm(29.7)
    section.page_width = Cm(21.0)
    section.top_margin = section.bottom_margin = Cm(2.5)
    section.left_margin = section.right_margin = Cm(2.5)
    bordure_page(section)

    entete = section.header.paragraphs[0]
    bordure_gauche(entete)
    run = entete.add_run(f"Document de suivi de tournoi – {TOURNOI}")
    run.font.color.rgb = BLEU_TITRE
    run.font.size = Pt(10)

    pied = section.footer.paragraphs[0]
    pied.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = pied.add_run("Page ")
    run.font.color.rgb = BLEU_PAGE
    espacement_lettres(run, 40)
    ajouter_champ(pied, "PAGE", "1")
    pied.add_run(" | ")
    ajouter_champ(pied, "NUMPAGES", "1")


def titre(doc, texte, niveau=1, italique=False):
    paragraphe = doc.add_heading(texte, level=niveau)
    for run in paragraphe.runs:
        run.font.italic = italique
    return paragraphe


def saut_de_page(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def ligne_champ(doc, libelle, valeur="", gras_italique=False):
    paragraphe = doc.add_paragraph()
    paragraphe.add_run(f"{libelle} : ")
    run = paragraphe.add_run(valeur or "……………………………………………………")
    run.bold = run.italic = gras_italique
    return paragraphe


def source(doc, texte="……………………"):
    run = doc.add_paragraph().add_run(f"Source : {texte}")
    run.italic = True


def nota_bene(doc, texte, surligne=False):
    paragraphe = doc.add_paragraph()
    run = paragraphe.add_run("Nota bene")
    run.bold = run.italic = True
    paragraphe.add_run(" : ")
    run = paragraphe.add_run(texte)
    if surligne:
        surlignage(run)
    return paragraphe


def emplacement_image(doc, texte, hauteur_cm=8):
    """Cadre vide où coller une capture d'écran."""
    table = doc.add_table(rows=1, cols=1)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cellule = table.cell(0, 0)
    cellule.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    table.rows[0].height = Cm(hauteur_cm)
    paragraphe = cellule.paragraphs[0]
    paragraphe.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraphe.add_run(f"[ {texte} ]")
    run.italic = True
    run.font.color.rgb = RGBColor(0x76, 0x71, 0x71)
    # Paragraphe séparateur : évite que deux cadres consécutifs fusionnent
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def tableau(doc, entetes, nb_lignes, largeurs_cm, colonnes_grises=()):
    """Tableau à en-tête orange ; les colonnes_grises sont pré-remplies en gris
    (comme la colonne PR/Rank* du modèle, à surligner en jaune si classé)."""
    table = doc.add_table(rows=nb_lignes + 1, cols=len(entetes))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, texte in enumerate(entetes):
        cellule = table.cell(0, i)
        fond_cellule(cellule, ORANGE)
        paragraphe = cellule.paragraphs[0]
        paragraphe.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraphe.add_run(texte)
        run.bold = True
        run.font.size = Pt(10)
    for ligne in table.rows[1:]:
        for i in colonnes_grises:
            fond_cellule(ligne.cells[i], GRIS)
    for ligne in table.rows:
        for i, largeur in enumerate(largeurs_cm):
            ligne.cells[i].width = Cm(largeur)
        for cellule in ligne.cells:
            cellule.paragraphs[0].paragraph_format.space_after = Pt(0)
    return table


# --- Sections du document ---------------------------------------------------

def page_de_garde(doc, organisation="", participation="", auteur=""):
    for _ in range(8):
        doc.add_paragraph()
    paragraphe = doc.add_paragraph()
    paragraphe.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraphe.add_run(TOURNOI)
    run.bold = run.italic = run.underline = True
    run.font.size = Pt(28)
    paragraphe.add_run(" :").font.size = Pt(28)

    paragraphe = doc.add_paragraph()
    paragraphe.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraphe.add_run("Document de suivi du tournoi")
    run.bold = True
    run.font.size = Pt(28)

    for _ in range(10):
        doc.add_paragraph()
    doc.add_paragraph("Organisation de l’évènement")
    run = doc.add_paragraph().add_run(organisation or "……………………………………")
    run.bold = run.italic = True
    paragraphe = doc.add_paragraph("Avec la participation de ")
    run = paragraphe.add_run(participation or "……………………")
    run.bold = run.italic = True
    doc.add_paragraph()
    doc.add_paragraph("Création & conception du document")
    run = doc.add_paragraph().add_run(auteur or "Prénom « pseudo » NOM")
    run.bold = run.italic = True
    saut_de_page(doc)


def sommaire(doc):
    paragraphe = doc.add_paragraph()
    run = paragraphe.add_run("SOMMAIRE")
    run.font.size = Pt(20)
    run = doc.add_paragraph().add_run("TABLE DES MATIÈRES")
    run.font.size = Pt(15)
    run.font.color.rgb = BLEU_TITRE
    ajouter_champ(
        doc.add_paragraph(),
        'TOC \\o "1-3" \\h \\z \\u',
        "Clic droit > Mettre à jour les champs pour générer le sommaire.",
    )
    saut_de_page(doc)


def informations_generales(doc):
    titre(doc, "Informations générales")
    ligne_champ(doc, "Organisation", gras_italique=True)
    doc.add_paragraph("avec la participation de ……………………", style="List Bullet")
    ligne_champ(doc, "Lieu de l’évènement")
    ligne_champ(doc, "Date(s)")
    ligne_champ(doc, "Nombre de joueurs/euses", "…… entrants")
    ligne_champ(doc, "Diffusions")
    ligne_champ(doc, "Bracket - start.gg")
    ligne_champ(doc, "Profil Supermajor")
    ligne_champ(doc, "UltRank Tournament Tierer - @kenniky1", "kenniky.github.io/web_ultrank_scorer.html")
    doc.add_paragraph()
    emplacement_image(doc, "Capture UltRank Tournament Tierer", 7)
    source(doc, f"UltRank Tournament Tierer – {TOURNOI}")
    saut_de_page(doc)


def formats_et_ruleset(doc):
    titre(doc, "Formats des brackets et ruleset")
    paragraphe = doc.add_paragraph()
    paragraphe.add_run("Format du tournoi et des sets").bold = True
    paragraphe.add_run(" : Bracket à ……………… élimination - ")
    run = paragraphe.add_run("BO…")
    run.bold = run.underline = True
    run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
    emplacement_image(doc, "Capture des brackets / pools (start.gg)", 7)
    emplacement_image(doc, "Image du ruleset (stages, règles de ban)", 10)
    source(doc, f"start.gg – {TOURNOI}")
    ligne_champ(doc, "Liens utiles")
    saut_de_page(doc)


def profils_joueurs(doc):
    titre(doc, "Profil des joueurs/euses")
    entetes = ["Tag/Struc.*", "Pseudo - Supermajor", "Seed", "Région/Ville", "Perso(s)", "PR/Rank*"]
    largeurs = [2.6, 4.4, 1.3, 3.0, 2.2, 2.8]
    for numero in range(1, NB_POOLS + 1):
        titre(doc, f"Pool {numero}", niveau=2, italique=True)
        tableau(doc, entetes, LIGNES_PAR_POOL, largeurs, colonnes_grises=(5,))
        saut_de_page(doc)

    titre(doc, "Abréviations et annotations", niveau=3)
    tableau(
        doc,
        ["Abré.°/Tricode", "Nom complet", "Abré.°/Tricode", "Nom complet"],
        LIGNES_ABREVIATIONS,
        [3.0, 5.2, 3.0, 5.2],
    )
    doc.add_paragraph()
    nota_bene(
        doc,
        "les classements inscrits dans les tableaux (colonne « PR/Rank »), ci-dessus, "
        "sont issus des derniers Power Rankings en date.",
        surligne=True,
    ).add_run(" (……………………)")

    table = tableau(doc, ["Abréviations", "Nom complet", "Abréviations", "Nom complet"], len(RANKINGS), [2.6, 5.6, 2.6, 5.6])
    for ligne, valeurs in zip(table.rows[1:], RANKINGS):
        for i, valeur in enumerate(valeurs):
            run = ligne.cells[i].paragraphs[0].add_run(valeur)
            run.bold = i % 2 == 0
            run.italic = i % 2 == 1
    saut_de_page(doc)

    titre(doc, "Représentation des personnages et des régions", niveau=3)
    nota_bene(
        doc,
        "seuls les mains characters (personnages principaux) des joueurs/euses ont été "
        "pris en compte pour la création de ce tableau.",
    )
    emplacement_image(doc, "Tier list des personnages joués", 10)
    emplacement_image(doc, "Tier list des régions représentées", 8)
    source(doc)
    saut_de_page(doc)


def deroule_bracket(doc):
    titre(doc, "Déroulé du bracket")
    nota_bene(
        doc,
        "le bracket Singles est définitif, néanmoins, des upsets/DQs peuvent être vecteurs "
        "de changements durant les phases de pools.",
    )
    doc.add_paragraph(
        "Le bracket présenté, ci-dessous, est une projection théorique (dit le Projected), "
        "à l’aide du seeding."
    )
    titre(doc, "Top 32", niveau=2)
    emplacement_image(doc, "Bracket Top 32 – Winners", 8)
    emplacement_image(doc, "Bracket Top 32 – Losers", 6)
    source(doc, f"start.gg – Top 32 – {TOURNOI}")
    saut_de_page(doc)

    emplacement_image(doc, "Visuel du seeding", 8)
    source(doc)
    titre(doc, "Représentation des personnages", niveau=3)
    emplacement_image(doc, "Représentation des personnages (Supermajor)", 7)
    source(doc, f"Supermajor – Top 32 – {TOURNOI}")
    saut_de_page(doc)

    titre(doc, "Top 8", niveau=2)
    emplacement_image(doc, "Bracket Top 8", 6)
    source(doc, f"start.gg – Top 8 – {TOURNOI}")
    titre(doc, "Head-to-Heads et setcounts", niveau=3)
    for _ in range(NB_HEAD_TO_HEADS):
        head_to_head(doc)
    saut_de_page(doc)


def head_to_head(doc):
    paragraphe = doc.add_paragraph()
    paragraphe.paragraph_format.space_before = Pt(12)
    rouge = RGBColor(0xFF, 0x00, 0x00)
    for texte, couleur, italique in (
        ("Joueur 1 ", None, False), ("(seed)", rouge, False), (" vs Joueur 2 ", None, False),
        ("(seed)", rouge, False), (" – Phase", None, True),
    ):
        run = paragraphe.add_run(texte)
        run.bold = True
        run.italic = italique
        if couleur:
            run.font.color.rgb = couleur

    table = tableau(doc, ["", "", "", "", ""], LIGNES_PAR_HEAD_TO_HEAD, [2.6, 1.4, 7.0, 1.4, 2.6])
    entete = table.rows[0].cells
    entete[0].merge(entete[1]).paragraphs[0].runs[0].text = "JOUEUR 1"
    entete[2].paragraphs[0].runs[0].text = "0 - 0"
    entete[3].merge(entete[4]).paragraphs[0].runs[0].text = "JOUEUR 2"
    # Nettoie les paragraphes vides laissés par la fusion
    for cellule in (table.cell(0, 0), table.cell(0, 3)):
        for p in cellule.paragraphs[1:]:
            p._p.getparent().remove(p._p)

    for ligne in table.rows[1:]:
        for i, texte in enumerate(("Joueur 1", "", "Tournoi (phase)", "", "Joueur 2")):
            paragraphe = ligne.cells[i].paragraphs[0]
            paragraphe.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = paragraphe.add_run(texte)
            run.bold = True
            run.italic = i == 2


def liens_utiles(doc):
    titre(doc, "Liens utiles")
    for nom, url in LIENS_UTILES:
        paragraphe = doc.add_paragraph()
        paragraphe.paragraph_format.space_before = Pt(12)
        run = paragraphe.add_run(nom)
        run.bold = run.italic = True
        lien_hypertexte(doc.add_paragraph(), url, url.removeprefix("https://"))
    saut_de_page(doc)


def notes_annexes(doc):
    titre(doc, "Notes annexes")


def main():
    doc = Document()
    configurer_styles(doc)
    configurer_section(doc.sections[0])

    page_de_garde(doc)
    sommaire(doc)
    informations_generales(doc)
    formats_et_ruleset(doc)
    profils_joueurs(doc)
    deroule_bracket(doc)
    liens_utiles(doc)
    notes_annexes(doc)

    doc.save(FICHIER_SORTIE)
    print(f"Document généré : {FICHIER_SORTIE}")


if __name__ == "__main__":
    main()
