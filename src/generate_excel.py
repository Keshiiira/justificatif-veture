"""
generate_excel.py
=================
Génère le classeur Excel "Justificatifs de Vêture – Allocation Mensuelle".
Reproduit fidèlement la mise en page du document papier officiel (Ariège).
Feuilles : Formulaire + Configuration.
Format : .xlsx (sans macros).

Disposition Formulaire :
  Ligne 1  – Titre fusionné A1:G1
  Lignes 3-6  – Identité assistant(e) / adresse
  Lignes 8-9  – Téléphone/Fax, Email/@/domaine
  Ligne 11 – Note italique soulignée
  Lignes 13-14 – Enfant accueilli / Référent social
  Ligne 16 – En-têtes tableau factures (A-E)
  Lignes 17-24 – 8 lignes de saisie
  Ligne 25 – TOTAL
  Ligne 26 – Attestation
  Ligne 28 – Fait à / Le
  Ligne 30 – M./Mme + signature
  Lignes 33-34 – Pied de page (9 pt)
"""
from __future__ import annotations
import argparse
import os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import range_boundaries
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins

# ── Palette ──────────────────────────────────────────────────────────────────
BLANC        = "FFFFFF"
GRIS_ENTETE  = "D9D9D9"   # fond gris clair pour les en-têtes du tableau
GRIS_FONCE   = "595959"   # fond sombre pour les en-têtes de configuration
GRIS_ALT     = "F2F2F2"   # lignes alternées configuration
FONT_NAME    = "Calibri"

# ── Helpers ───────────────────────────────────────────────────────────────────
def _s(style: str = "thin") -> Side:
    return Side(style=style)

def bdr_bottom() -> Border:
    """Bordure inférieure seule – style champ texte papier."""
    return Border(bottom=_s("thin"))

def bdr_all(style: str = "thin") -> Border:
    s = _s(style)
    return Border(top=s, bottom=s, left=s, right=s)

def fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)

def fnt(bold: bool = False, size: int = 11, color: str = "000000",
        italic: bool = False, underline: bool = False) -> Font:
    return Font(
        name=FONT_NAME, size=size, bold=bold, color=color,
        italic=italic, underline="single" if underline else None,
    )

def aln(h: str = "left", v: str = "center",
        wrap: bool = False, indent: int = 0) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=indent)

def apply_bottom_border(ws, rng: str) -> None:
    """Bordure inférieure sur toutes les cellules de la dernière ligne de la plage."""
    min_col, min_row, max_col, max_row = range_boundaries(rng)
    bdr = bdr_bottom()
    for c in range(min_col, max_col + 1):
        ws.cell(row=max_row, column=c).border = bdr

def apply_outer_border(ws, rng: str, style: str = "thin") -> None:
    """Bordure périphérique sur une plage (fusionnée ou non)."""
    min_col, min_row, max_col, max_row = range_boundaries(rng)
    s = _s(style)
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(row=r, column=c).border = Border(
                top    = s if r == min_row else None,
                bottom = s if r == max_row else None,
                left   = s if c == min_col else None,
                right  = s if c == max_col else None,
            )

def merged(ws, rng: str, value=None, font_=None, aln_=None, fill_=None) -> None:
    """Fusionne une plage et applique valeur + styles à la cellule maître."""
    ws.merge_cells(rng)
    min_col, min_row, *_ = range_boundaries(rng)
    c = ws.cell(row=min_row, column=min_col)
    if value  is not None: c.value     = value
    if font_  is not None: c.font      = font_
    if aln_   is not None: c.alignment = aln_
    if fill_  is not None: c.fill      = fill_

# ── Données de configuration (exemples modifiables dans l'onglet Configuration)
CONFIG_AF = {
    "nom_prenom": "Dupont Marie",
    "adresse":    "12 Rue des Lilas, 09000 Foix",
    "telephone":  "06 12 34 56 78",
    "fax":        "",
    "email":      "marie.dupont@email.fr",
    "fait_a":     "Foix",
}
ENFANTS_EXEMPLE = [
    ("Martin Lucas",  "Mme Bernard Sophie"),
    ("Petit Emma",    "M. Leclerc Paul"),
]

# ── Feuille Configuration ─────────────────────────────────────────────────────
def build_configuration_sheet(wb: Workbook) -> None:
    ws = wb.create_sheet("Configuration")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 42

    # Ligne 1 : titre
    merged(ws, "A1:B1", value="⚙  CONFIGURATION",
           font_=fnt(bold=True, size=13, color=BLANC),
           aln_=aln("center"), fill_=fill(GRIS_FONCE))
    apply_outer_border(ws, "A1:B1", "medium")
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 6

    # Ligne 3 : section en-tête AF
    merged(ws, "A3:B3",
           value="Informations fixes – Assistant(e) Familial(e)",
           font_=fnt(bold=True, size=11, color=BLANC),
           aln_=aln("left", indent=1), fill_=fill("404040"))
    apply_outer_border(ws, "A3:B3")
    ws.row_dimensions[3].height = 22

    # Lignes 4-9 : champs AF
    # B4=nom_prenom  B5=adresse  B6=telephone  B7=fax  B8=email  B9=fait_a
    fields = [
        ("Nom / Prénom AF",  CONFIG_AF["nom_prenom"]),
        ("Adresse",          CONFIG_AF["adresse"]),
        ("Téléphone",        CONFIG_AF["telephone"]),
        ("Fax",              CONFIG_AF["fax"]),
        ("Email",            CONFIG_AF["email"]),
        ("Fait à (ville)",   CONFIG_AF["fait_a"]),
    ]
    for i, (label, val) in enumerate(fields):
        row = 4 + i
        lc = ws.cell(row=row, column=1, value=label)
        lc.font = fnt(bold=True, size=10); lc.fill = fill("F2F2F2")
        lc.border = bdr_all(); lc.alignment = aln("left", indent=1)
        vc = ws.cell(row=row, column=2, value=val)
        vc.font = fnt(size=10); vc.border = bdr_all()
        vc.alignment = aln("left", indent=1)
        ws.row_dimensions[row].height = 18

    ws.row_dimensions[10].height = 8

    # Ligne 11 : section en-tête Enfants
    merged(ws, "A11:B11",
           value="Enfants accueillis et référents sociaux",
           font_=fnt(bold=True, size=11, color=BLANC),
           aln_=aln("left", indent=1), fill_=fill("404040"))
    apply_outer_border(ws, "A11:B11")
    ws.row_dimensions[11].height = 22

    # Ligne 12 : en-têtes colonnes enfants
    for ci, h in enumerate(["Enfant (Nom Prénom)", "Référent social"], start=1):
        c = ws.cell(row=12, column=ci, value=h)
        c.font = fnt(bold=True, size=10, color=BLANC)
        c.fill = fill(GRIS_FONCE); c.border = bdr_all(); c.alignment = aln("center")
    ws.row_dimensions[12].height = 20

    # Lignes 13-30 : données enfants (18 slots)
    for i, (enfant, ref) in enumerate(ENFANTS_EXEMPLE):
        row = 13 + i
        bg = BLANC if i % 2 == 0 else GRIS_ALT
        for ci, val in ((1, enfant), (2, ref)):
            c = ws.cell(row=row, column=ci, value=val)
            c.font = fnt(size=10); c.fill = fill(bg)
            c.border = bdr_all(); c.alignment = aln("left", indent=1)
        ws.row_dimensions[row].height = 18

    for row in range(13 + len(ENFANTS_EXEMPLE), 31):
        bg = BLANC if row % 2 == 0 else GRIS_ALT
        for ci in (1, 2):
            c = ws.cell(row=row, column=ci, value="")
            c.fill = fill(bg); c.border = bdr_all()
            c.alignment = aln("left", indent=1)
        ws.row_dimensions[row].height = 18

    ws.row_dimensions[31].height = 8
    merged(ws, "A32:B32",
           value="ℹ  Ajoutez autant de lignes enfant/référent que nécessaire (lignes 13 à 30).",
           font_=fnt(italic=True, size=9, color="666666"),
           aln_=aln("left", indent=1))

# ── Feuille Formulaire ────────────────────────────────────────────────────────
def build_formulaire_sheet(wb: Workbook) -> None:
    ws = wb.create_sheet("Formulaire", 0)

    # Largeurs : A:5 | B:14 | C:24 | D:12 | E:10 | F:10 | G:10
    for col, w in [("A", 5), ("B", 14), ("C", 24), ("D", 12),
                   ("E", 10), ("F", 10), ("G", 10)]:
        ws.column_dimensions[col].width = w

    # ── Ligne 1 : Titre centré ──────────────────────────────────────────────
    merged(ws, "A1:G1",
           value="JUSTIFICATIFS DE VETURE – ALLOCATION MENSUELLE",
           font_=fnt(bold=True, size=12),
           aln_=aln("center"))
    apply_outer_border(ws, "A1:G1", "medium")
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 8  # espace

    # ── Ligne 3 : Assistant(e) Familial(e) ─────────────────────────────────
    c = ws.cell(row=3, column=1, value="Assistant(e) Familial(e) :")
    c.font = fnt(bold=True, size=11); c.alignment = aln("left")
    merged(ws, "B3:G3", value="=Configuration!B4",
           font_=fnt(size=11), aln_=aln("left", indent=1))
    apply_bottom_border(ws, "B3:G3")
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 6  # espace

    # ── Ligne 5 : Employeur ────────────────────────────────────────────────
    merged(ws, "A5:G5",
           value="Employé(e) par le Conseil Départemental de l'Ariège",
           font_=fnt(size=11), aln_=aln("left", indent=1))
    ws.row_dimensions[5].height = 16

    # ── Ligne 6 : Adresse ─────────────────────────────────────────────────
    c = ws.cell(row=6, column=1, value="Adresse :")
    c.font = fnt(bold=True, size=11); c.alignment = aln("left")
    merged(ws, "B6:G6", value="=Configuration!B5",
           font_=fnt(size=11), aln_=aln("left", wrap=True, indent=1))
    apply_bottom_border(ws, "B6:G6")
    ws.row_dimensions[6].height = 18
    ws.row_dimensions[7].height = 6  # espace

    # ── Ligne 8 : Téléphone / Fax ─────────────────────────────────────────
    c = ws.cell(row=8, column=1, value="Téléphone :")
    c.font = fnt(bold=True, size=11); c.alignment = aln("left")
    merged(ws, "B8:D8", value="=Configuration!B6",
           font_=fnt(size=11), aln_=aln("left", indent=1))
    apply_bottom_border(ws, "B8:D8")

    c = ws.cell(row=8, column=5, value="Fax :")
    c.font = fnt(bold=True, size=11); c.alignment = aln("left")
    merged(ws, "F8:G8", value="=Configuration!B7",
           font_=fnt(size=11), aln_=aln("left", indent=1))
    apply_bottom_border(ws, "F8:G8")
    ws.row_dimensions[8].height = 18

    # ── Ligne 9 : Email / @ / domaine ─────────────────────────────────────
    c = ws.cell(row=9, column=1, value="Email :")
    c.font = fnt(bold=True, size=11); c.alignment = aln("left")
    merged(ws, "B9:D9", value="=Configuration!B8",
           font_=fnt(size=11), aln_=aln("left", indent=1))
    apply_bottom_border(ws, "B9:D9")

    c = ws.cell(row=9, column=5, value="@")
    c.font = fnt(size=11); c.alignment = aln("center")

    merged(ws, "F9:G9", font_=fnt(size=11), aln_=aln("left", indent=1))
    apply_bottom_border(ws, "F9:G9")
    ws.row_dimensions[9].height = 18
    ws.row_dimensions[10].height = 8  # espace

    # ── Ligne 11 : Note italique soulignée ────────────────────────────────
    merged(ws, "A11:G11",
           value=("Au présent document, merci d'agrafer EXCLUSIVEMENT les factures "
                  "correspondantes aux achats de l'enfant accueilli"),
           font_=fnt(italic=True, underline=True, size=10),
           aln_=aln("left", wrap=True, indent=1))
    ws.row_dimensions[11].height = 28
    ws.row_dimensions[12].height = 8  # espace

    # ── Ligne 13 : Identité enfant accueilli ──────────────────────────────
    merged(ws, "A13:D13",
           value="Identité de l'enfant accueilli et concerné par l'allocation :",
           font_=fnt(bold=True, size=11), aln_=aln("left"))
    merged(ws, "E13:G13",
           font_=fnt(size=11, color="000060"), aln_=aln("left", indent=1))
    apply_bottom_border(ws, "E13:G13")

    dv_enfant = DataValidation(
        type="list",
        formula1="Configuration!$A$13:$A$30",
        allow_blank=True, showDropDown=False,
        showErrorMessage=True,
        errorTitle="Valeur invalide",
        error="Veuillez sélectionner un enfant depuis la liste.",
    )
    ws.add_data_validation(dv_enfant)
    dv_enfant.add(ws["E13"])
    ws.row_dimensions[13].height = 20

    # ── Ligne 14 : Référent social ────────────────────────────────────────
    merged(ws, "A14:B14", value="Référent social :",
           font_=fnt(bold=True, size=11), aln_=aln("left"))
    merged(ws, "C14:G14",
           value=('=IFERROR(INDEX(Configuration!$B$13:$B$30,'
                  'MATCH(E13,Configuration!$A$13:$A$30,0)),"")'),
           font_=fnt(italic=True, size=11, color="444444"),
           aln_=aln("left", indent=1))
    apply_bottom_border(ws, "C14:G14")
    ws.row_dimensions[14].height = 20
    ws.row_dimensions[15].height = 8  # espace

    # ── Ligne 16 : En-têtes tableau factures ──────────────────────────────
    for ci, h in [(1, "N° de la facture"), (2, "Type d'Achat"),
                  (3, "Fournisseurs"), (4, "Date"), (5, "Coût")]:
        c = ws.cell(row=16, column=ci, value=h)
        c.font = fnt(bold=True, size=11)
        c.fill = fill(GRIS_ENTETE); c.border = bdr_all(); c.alignment = aln("center")
    ws.row_dimensions[16].height = 22

    # ── Lignes 17-24 : 8 lignes de saisie ─────────────────────────────────
    for i in range(8):
        row = 17 + i
        n = ws.cell(row=row, column=1, value=i + 1)
        n.font = fnt(size=10, color="666666"); n.border = bdr_all(); n.alignment = aln("center")
        for ci in (2, 3):
            c = ws.cell(row=row, column=ci, value="")
            c.font = fnt(size=11); c.border = bdr_all(); c.alignment = aln("left", indent=1)
        d = ws.cell(row=row, column=4, value="")
        d.font = fnt(size=11); d.border = bdr_all()
        d.alignment = aln("center"); d.number_format = "DD/MM/YYYY"
        e = ws.cell(row=row, column=5, value="")
        e.font = fnt(size=11); e.border = bdr_all()
        e.alignment = aln("right"); e.number_format = '#,##0.00 "€"'
        ws.row_dimensions[row].height = 18

    # ── Ligne 25 : TOTAL ───────────────────────────────────────────────────
    merged(ws, "A25:D25", value="TOTAL",
           font_=fnt(bold=True, size=11),
           fill_=fill(GRIS_ENTETE), aln_=aln("right", indent=1))
    apply_outer_border(ws, "A25:D25")
    tv = ws.cell(row=25, column=5, value="=SUM(E17:E24)")
    tv.font = fnt(bold=True, size=11); tv.fill = fill(GRIS_ENTETE)
    tv.border = bdr_all(); tv.number_format = '#,##0.00 "€"'; tv.alignment = aln("right")
    ws.row_dimensions[25].height = 20

    # ── Ligne 26 : Attestation ─────────────────────────────────────────────
    merged(ws, "A26:G26",
           value="Je soussigné(e), auteur du présent état, en certifie l'exactitude.",
           font_=fnt(size=11), aln_=aln("left", indent=1))
    ws.row_dimensions[26].height = 18
    ws.row_dimensions[27].height = 8  # espace

    # ── Ligne 28 : Fait à / Le ─────────────────────────────────────────────
    c = ws.cell(row=28, column=1, value="Fait à")
    c.font = fnt(bold=True, size=11); c.alignment = aln("left")
    merged(ws, "B28:D28", value="=Configuration!B9",
           font_=fnt(size=11), aln_=aln("left", indent=1))
    apply_bottom_border(ws, "B28:D28")

    c = ws.cell(row=28, column=5, value="Le")
    c.font = fnt(bold=True, size=11); c.alignment = aln("center")
    merged(ws, "F28:G28", value="=TODAY()",
           font_=fnt(size=11), aln_=aln("center"))
    ws["F28"].number_format = "DD/MM/YYYY"
    apply_bottom_border(ws, "F28:G28")
    ws.row_dimensions[28].height = 20
    ws.row_dimensions[29].height = 8  # espace

    # ── Ligne 30 : M./Mme + zone signature ────────────────────────────────
    c = ws.cell(row=30, column=1, value="M./Mme")
    c.font = fnt(bold=True, size=11); c.alignment = aln("left")
    merged(ws, "B30:G30", font_=fnt(size=11), aln_=aln("left", indent=1))
    apply_bottom_border(ws, "B30:G30")
    ws.row_dimensions[30].height = 24
    ws.row_dimensions[31].height = 6
    ws.row_dimensions[32].height = 6

    # ── Lignes 33-34 : Pied de page ───────────────────────────────────────
    merged(ws, "A33:G33",
           value=("Ce document doit être complété MENSUELLEMENT ET PAR ENFANT "
                  "et adressé à : D.S.D. - DAEFP"),
           font_=fnt(size=9), aln_=aln("left", indent=1))
    ws.row_dimensions[33].height = 14

    merged(ws, "A34:G34",
           value="ASS FAM – mars 2017/ justificatifs vêture (allocation mensuelle)",
           font_=fnt(size=9), aln_=aln("left", indent=1))
    ws.row_dimensions[34].height = 14

    # ── Validations de données ─────────────────────────────────────────────
    # Col D (lignes 17-24) : date
    dv_date = DataValidation(
        type="date", allow_blank=True,
        showErrorMessage=True,
        errorTitle="Date invalide",
        error="Veuillez entrer une date valide (JJ/MM/AAAA).",
    )
    ws.add_data_validation(dv_date)
    for r in range(17, 25):
        dv_date.add(ws.cell(row=r, column=4))

    # Col E (lignes 17-24) : nombre décimal ≥ 0
    dv_num = DataValidation(
        type="decimal", operator="greaterThanOrEqual", formula1="0",
        allow_blank=True,
        showErrorMessage=True,
        errorTitle="Montant invalide",
        error="Veuillez entrer un montant positif.",
    )
    ws.add_data_validation(dv_num)
    for r in range(17, 25):
        dv_num.add(ws.cell(row=r, column=5))

    # ── Mise en page ───────────────────────────────────────────────────────
    ws.freeze_panes = "A17"  # A17 = 1ère cellule non gelée → lignes 1-16 figées (incl. en-têtes)
    ws.page_setup.paperSize   = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToHeight = 1
    ws.page_setup.fitToWidth  = 1
    ws.print_area = "A1:G34"
    ws.page_margins = PageMargins(
        left=0.79, right=0.79, top=0.79, bottom=0.79,
        header=0.3, footer=0.3,
    )


# ── SAVE & MAIN ───────────────────────────────────────────────────────────────
def save_as_xlsx(wb: Workbook, output_path: str) -> None:
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    wb.save(output_path)
    print(f"✅  Classeur Excel généré : {output_path}")

def generate(output_path: str) -> None:
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]
    build_formulaire_sheet(wb)
    build_configuration_sheet(wb)
    save_as_xlsx(wb, output_path)

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Génère le classeur Excel Justificatifs de Vêture (.xlsx)."
    )
    parser.add_argument(
        "--output", default="output/justificatif_veture.xlsx",
        help="Chemin du fichier Excel généré",
    )
    args = parser.parse_args()
    generate(args.output)

if __name__ == "__main__":
    main()

