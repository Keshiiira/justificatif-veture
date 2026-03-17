"""
generate_excel.py
=================
Génère le classeur Excel "Justificatifs de Vêture".
Feuilles : Formulaire + Configuration.
Esthétique document administratif officiel : fond blanc, gris neutre,
bordures noires fines, sans couleurs criantes.
"""
from __future__ import annotations
import argparse, os
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import range_boundaries
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.page import PageMargins

# ── Palette ──────────────────────────────────────────────────────────────────
NOIR          = "000000"
BLANC         = "FFFFFF"
GRIS_TITRE    = "1A1A2E"   # quasi-noir bleuté pour le titre principal
GRIS_SECTION  = "404040"   # texte des bandeaux de section
GRIS_BG_SEC   = "D9D9D9"   # fond des bandeaux de section
GRIS_BG_COL   = "595959"   # fond des en-têtes de colonnes du tableau
GRIS_ALT      = "F2F2F2"   # lignes alternées du tableau
GRIS_LABEL    = "F7F7F7"   # fond des cellules libellé (très léger)
GRIS_TOTAL    = "D9D9D9"   # fond ligne TOTAL
GRIS_SAISIE   = "FFFFFF"   # cellules de saisie : blanc pur

# ── Helpers ───────────────────────────────────────────────────────────────────
def _side(style="thin"):  return Side(style=style)

def border_all(style="thin") -> Border:
    s = _side(style)
    return Border(top=s, bottom=s, left=s, right=s)

def border_outer(style="thin") -> Border:
    s = _side(style)
    return Border(top=s, bottom=s, left=s, right=s)

def fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)

def font(bold=False, size=10, color=NOIR, italic=False) -> Font:
    return Font(bold=bold, size=size, color=color, italic=italic, name="Calibri")

def align(h="left", v="center", wrap=False, indent=0) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=indent)

def apply_border_merged(ws, rng: str, style="thin") -> None:
    """Bordure périphérique fine sur toutes les cellules d'une plage fusionnée."""
    min_col, min_row, max_col, max_row = range_boundaries(rng)
    s = _side(style)
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(row=r, column=c).border = Border(
                top    = s if r == min_row else None,
                bottom = s if r == max_row else None,
                left   = s if c == min_col else None,
                right  = s if c == max_col else None,
            )

def section_header(ws, rng: str, text: str, row_h: int = 22) -> None:
    """Bandeau de section gris avec texte noir gras."""
    ws.merge_cells(rng)
    min_col, min_row, *_ = range_boundaries(rng)
    c = ws.cell(row=min_row, column=min_col)
    c.value = text
    c.font  = font(bold=True, size=10, color=GRIS_SECTION)
    c.fill  = fill(GRIS_BG_SEC)
    c.alignment = align("left", indent=1)
    apply_border_merged(ws, rng, "medium")
    ws.row_dimensions[min_row].height = row_h

def label_cell(ws, row, col, text, width_hint=None) -> None:
    c = ws.cell(row=row, column=col, value=text)
    c.font      = font(bold=True, size=10)
    c.fill      = fill(GRIS_LABEL)
    c.border    = border_all("thin")
    c.alignment = align("left", indent=1)
    ws.row_dimensions[row].height = 18

def value_cell(ws, row, col, value="", fmt=None, saisie=False) -> None:
    c = ws.cell(row=row, column=col, value=value)
    c.font      = font(size=10)
    c.fill      = fill(GRIS_SAISIE)
    c.border    = border_all("thin")
    c.alignment = align("left", indent=1)
    if fmt: c.number_format = fmt
    ws.row_dimensions[row].height = 18

def value_merged(ws, rng: str, value="", fmt=None, italic=False) -> None:
    ws.merge_cells(rng)
    min_col, min_row, *_ = range_boundaries(rng)
    c = ws.cell(row=min_row, column=min_col, value=value)
    c.font      = font(size=10, italic=italic)
    c.fill      = fill(GRIS_SAISIE)
    c.alignment = align("left", indent=1)
    if fmt: c.number_format = fmt
    apply_border_merged(ws, rng, "thin")
    ws.row_dimensions[min_row].height = 18

# ── CONFIG ────────────────────────────────────────────────────────────────────
CONFIG_AF = {
    "nom_prenom": "Dupont Marie",
    "adresse":    "12 Rue des Lilas, 09000 Foix",
    "telephone":  "06 12 34 56 78",
    "email":      "marie.dupont@email.fr",
    "fait_a":     "Foix",
}
ENFANTS_EXEMPLE = [
    ("Martin Lucas",  "Mme Bernard Sophie"),
    ("Petit Emma",    "M. Leclerc Paul"),
]

# ── FEUILLE CONFIGURATION ─────────────────────────────────────────────────────
def build_configuration_sheet(wb: Workbook) -> None:
    ws = wb.create_sheet("Configuration")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 5

    # Titre
    ws.merge_cells("A1:B1")
    c = ws["A1"]
    c.value     = "⚙  CONFIGURATION"
    c.font      = font(bold=True, size=13, color=BLANC)
    c.fill      = fill(GRIS_BG_COL)
    c.alignment = align("center")
    apply_border_merged(ws, "A1:B1", "medium")
    ws.row_dimensions[1].height = 28

    ws.row_dimensions[2].height = 6

    # Section AF
    section_header(ws, "A3:B3", "Informations fixes – Assistant(e) Familial(e)")
    fields = [
        ("Nom / Prénom AF",  CONFIG_AF["nom_prenom"]),
        ("Adresse",          CONFIG_AF["adresse"]),
        ("Téléphone",        CONFIG_AF["telephone"]),
        ("Email",            CONFIG_AF["email"]),
        ("Fait à (ville)",   CONFIG_AF["fait_a"]),
    ]
    for i, (lbl, val) in enumerate(fields):
        row = 4 + i
        label_cell(ws, row, 1, lbl)
        value_cell(ws, row, 2, val)

    ws.row_dimensions[9].height = 8

    # Section enfants
    section_header(ws, "A10:B10", "Enfants accueillis et référents sociaux")
    for col, h in enumerate(["Enfant (Nom Prénom)", "Référent social"], start=1):
        c = ws.cell(row=11, column=col, value=h)
        c.font      = font(bold=True, size=10, color=BLANC)
        c.fill      = fill(GRIS_BG_COL)
        c.border    = border_all()
        c.alignment = align("center")
    ws.row_dimensions[11].height = 20

    for i, (enfant, referent) in enumerate(ENFANTS_EXEMPLE):
        row  = 12 + i
        bg   = BLANC if i % 2 == 0 else GRIS_ALT
        for col, val in ((1, enfant), (2, referent)):
            c = ws.cell(row=row, column=col, value=val)
            c.fill      = fill(bg)
            c.border    = border_all()
            c.alignment = align("left", indent=1)
            c.font      = font(size=10)
        ws.row_dimensions[row].height = 18

    for row in range(12 + len(ENFANTS_EXEMPLE), 30):
        bg = BLANC if row % 2 == 0 else GRIS_ALT
        for col in (1, 2):
            c = ws.cell(row=row, column=col, value="")
            c.fill      = fill(bg)
            c.border    = border_all()
            c.alignment = align("left", indent=1)
        ws.row_dimensions[row].height = 18

    ws.merge_cells("A31:B31")
    note = ws["A31"]
    note.value     = "ℹ  Ajoutez autant de lignes enfant/référent que nécessaire (lignes 12 à 29)."
    note.font      = font(size=9, italic=True, color="666666")
    note.alignment = align("left", indent=1)

# ── FEUILLE FORMULAIRE ────────────────────────────────────────────────────────
def build_formulaire_sheet(wb: Workbook) -> None:
    ws = wb.create_sheet("Formulaire", 0)

    # Colonnes : A(étroit) B(libellé) C(valeur large) D(date) E(coût)
    for col, w in {"A": 6, "B": 32, "C": 34, "D": 14, "E": 14}.items():
        ws.column_dimensions[col].width = w

    # ── TITRE ──────────────────────────────────────────────────────────────
    ws.merge_cells("A1:E1")
    t = ws["A1"]
    t.value     = "JUSTIFICATIFS DE VÊTURE – ALLOCATION MENSUELLE"
    t.font      = font(bold=True, size=15, color=BLANC)
    t.fill      = fill(GRIS_TITRE)
    t.alignment = align("center")
    apply_border_merged(ws, "A1:E1", "medium")
    ws.row_dimensions[1].height = 34

    # ── SOUS-TITRE ─────────────────────────────────────────────────────────
    ws.merge_cells("A2:E2")
    st = ws["A2"]
    st.value     = ("Pour être pris en charge, les justificatifs de vêture "
                    "doivent être présentés mensuellement et par enfant.")
    st.font      = font(italic=True, size=9, color="444444")
    st.fill      = fill(GRIS_ALT)
    st.alignment = align("center", wrap=True)
    apply_border_merged(ws, "A2:E2", "thin")
    ws.row_dimensions[2].height = 28

    ws.row_dimensions[3].height = 6

    # ── MOIS ET ANNÉE ──────────────────────────────────────────────────────
    ws.merge_cells("A4:B4")
    lm = ws["A4"]
    lm.value     = "Mois et Année :"
    lm.font      = font(bold=True, size=10)
    lm.fill      = fill(GRIS_LABEL)
    lm.alignment = align("right", indent=1)
    apply_border_merged(ws, "A4:B4", "thin")

    ws.merge_cells("C4:E4")
    vm = ws["C4"]
    vm.font        = font(size=10)
    vm.fill        = fill(GRIS_SAISIE)
    vm.number_format = "MMMM YYYY"
    vm.alignment   = align("left", indent=1)
    apply_border_merged(ws, "C4:E4", "thin")
    ws.row_dimensions[4].height = 20

    ws.row_dimensions[5].height = 8

    # ── SECTION ASSISTANT FAMILIAL ─────────────────────────────────────────
    section_header(ws, "A6:E6", "IDENTITÉ DE L'ASSISTANT(E) FAMILIAL(E)")

    af_rows = [
        (7,  "Nom / Prénom :",  "=Configuration!B4"),
        (8,  "Adresse :",       "=Configuration!B5"),
        (9,  "Téléphone :",     "=Configuration!B6"),
        (10, "Email :",         "=Configuration!B7"),
    ]
    for row, lbl, formula in af_rows:
        label_cell(ws, row, 2, lbl)
        value_merged(ws, f"C{row}:E{row}", formula)
        ws.row_dimensions[row].height = 18

    ws.row_dimensions[11].height = 8

    # ── SECTION ENFANT ─────────────────────────────────────────────────────
    section_header(ws, "A12:E12",
                   "IDENTITÉ DE L'ENFANT ACCUEILLI ET CONCERNÉ PAR L'ALLOCATION")

    # Enfant sélection (liste déroulante)
    label_cell(ws, 13, 2, "Enfant accueilli :")
    ws.merge_cells("C13:E13")
    ce = ws["C13"]
    ce.font      = font(size=10, color="000060")
    ce.fill      = fill(GRIS_SAISIE)
    ce.alignment = align("left", indent=1)
    apply_border_merged(ws, "C13:E13", "thin")
    ws.row_dimensions[13].height = 20

    dv = DataValidation(
        type="list",
        formula1="Configuration!$A$12:$A$29",
        allow_blank=True, showDropDown=False,
        showErrorMessage=True,
        errorTitle="Valeur invalide",
        error="Veuillez sélectionner un enfant depuis la liste.",
    )
    ws.add_data_validation(dv)
    dv.add(ws["C13"])

    # Référent (formule auto)
    label_cell(ws, 14, 2, "Référent social :")
    ws.merge_cells("C14:E14")
    cr = ws["C14"]
    cr.value     = '=IFERROR(INDEX(Configuration!$B$12:$B$29,MATCH(C13,Configuration!$A$12:$A$29,0)),"")'
    cr.font      = font(size=10, italic=True, color="444444")
    cr.fill      = fill(GRIS_ALT)
    cr.alignment = align("left", indent=1)
    apply_border_merged(ws, "C14:E14", "thin")
    ws.row_dimensions[14].height = 18

    ws.row_dimensions[15].height = 8

    # ── SECTION FACTURES ───────────────────────────────────────────────────
    section_header(ws, "A16:E16", "DÉTAIL DES FACTURES")

    # En-têtes colonnes
    col_headers = ["N°", "Nature de l'achat", "Fournisseur / Magasin", "Date", "Montant (€)"]
    for ci, h in enumerate(col_headers, start=1):
        c = ws.cell(row=17, column=ci, value=h)
        c.font      = font(bold=True, size=10, color=BLANC)
        c.fill      = fill(GRIS_BG_COL)
        c.border    = border_all("thin")
        c.alignment = align("center")
    ws.row_dimensions[17].height = 22

    # 10 lignes de factures
    for i in range(10):
        row = 18 + i
        bg  = BLANC if i % 2 == 0 else GRIS_ALT

        # N°
        n = ws.cell(row=row, column=1, value=i + 1)
        n.font      = font(size=9, color="666666")
        n.fill      = fill(GRIS_ALT)
        n.border    = border_all("thin")
        n.alignment = align("center")

        # Colonnes saisie
        for col in (2, 3, 4, 5):
            c = ws.cell(row=row, column=col, value="")
            c.fill      = fill(bg)
            c.border    = border_all("thin")
            c.alignment = align("left", indent=1)
            if col == 4: c.number_format = "DD/MM/YYYY"
            if col == 5:
                c.number_format = '#,##0.00 "€"'
                c.alignment     = align("right")
        ws.row_dimensions[row].height = 18

    # Ligne TOTAL
    rT = 28
    ws.merge_cells(f"A{rT}:D{rT}")
    tl = ws[f"A{rT}"]
    tl.value     = "TOTAL"
    tl.font      = font(bold=True, size=10)
    tl.fill      = fill(GRIS_TOTAL)
    tl.alignment = align("right", indent=1)
    apply_border_merged(ws, f"A{rT}:D{rT}", "medium")

    tv = ws.cell(row=rT, column=5)
    tv.value        = "=SUM(E18:E27)"
    tv.font         = font(bold=True, size=10)
    tv.fill         = fill(GRIS_TOTAL)
    tv.border       = border_all("medium")
    tv.number_format = '#,##0.00 "€"'
    tv.alignment    = align("right")
    ws.row_dimensions[rT].height = 22

    ws.row_dimensions[29].height = 8

    # ── SECTION SIGNATURE ──────────────────────────────────────────────────
    section_header(ws, "A30:E30", "CERTIFICATION ET SIGNATURE")

    # Texte de certification
    ws.merge_cells("A31:E31")
    cert = ws["A31"]
    cert.value = ('=CONCATENATE("Je soussigné(e) ",Configuration!B4,'
                  '" certifie avoir réglé les achats de vêtements ci-dessus listés.")')
    cert.font      = font(size=10, italic=True)
    cert.fill      = fill(BLANC)
    cert.alignment = align("left", wrap=True, indent=1)
    apply_border_merged(ws, "A31:E31", "thin")
    ws.row_dimensions[31].height = 24

    ws.row_dimensions[32].height = 8

    # Fait à / Le :
    lf = ws.cell(row=33, column=2, value="Fait à :")
    lf.font = font(bold=True, size=10); lf.fill = fill(GRIS_LABEL)
    lf.border = border_all(); lf.alignment = align("left", indent=1)

    vf = ws.cell(row=33, column=3, value="=Configuration!B8")
    vf.font = font(size=10); vf.fill = fill(GRIS_SAISIE)
    vf.border = border_all(); vf.alignment = align("left", indent=1)

    ll = ws.cell(row=33, column=4, value="Le :")
    ll.font = font(bold=True, size=10); ll.fill = fill(GRIS_LABEL)
    ll.border = border_all(); ll.alignment = align("center")

    vl = ws.cell(row=33, column=5, value="=TODAY()")
    vl.font = font(size=10); vl.fill = fill(GRIS_SAISIE)
    vl.border = border_all(); vl.number_format = "DD/MM/YYYY"
    vl.alignment = align("center")
    ws.row_dimensions[33].height = 20

    # M. / Mme :
    lmm = ws.cell(row=34, column=2, value="M. / Mme :")
    lmm.font = font(bold=True, size=10); lmm.fill = fill(GRIS_LABEL)
    lmm.border = border_all(); lmm.alignment = align("left", indent=1)

    ws.merge_cells("C34:E34")
    vmm = ws["C34"]
    vmm.fill = fill(GRIS_SAISIE)
    vmm.alignment = align("left", indent=1)
    apply_border_merged(ws, "C34:E34", "thin")
    ws.row_dimensions[34].height = 24

    ws.row_dimensions[35].height = 8

    # Signature (zone vide)
    ls = ws.cell(row=36, column=2, value="Signature :")
    ls.font = font(bold=True, size=10); ls.fill = fill(GRIS_LABEL)
    ls.border = border_all(); ls.alignment = align("left", "top", indent=1)

    ws.merge_cells("C36:E38")
    sz = ws["C36"]
    sz.fill      = fill(BLANC)
    sz.alignment = align("center")
    apply_border_merged(ws, "C36:E38", "thin")
    for row in (36, 37, 38):
        ws.row_dimensions[row].height = 22

    ws.row_dimensions[39].height = 10

    # ── MISE EN PAGE ───────────────────────────────────────────────────────
    ws.freeze_panes = "A3"
    ws.page_setup.paperSize    = ws.PAPERSIZE_A4
    ws.page_setup.orientation  = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToPage    = True
    ws.page_setup.fitToHeight  = 1
    ws.page_setup.fitToWidth   = 1
    ws.print_area = "A1:E38"
    ws.page_margins = PageMargins(
        left=0.55, right=0.55, top=0.75, bottom=0.75,
        header=0.3, footer=0.3
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

