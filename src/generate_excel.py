"""
generate_excel.py
=================
Génère le classeur Excel macro-activé "Justificatifs de Vêture".
Feuilles : Formulaire + Configuration.
Fonctionnalités : bordures corrigées sur cellules fusionnées,
deux boutons VBA (Ajouter factures / Générer PDF), format .xlsm.
"""
from __future__ import annotations
import argparse, io, os, zipfile
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import range_boundaries
from openpyxl.worksheet.datavalidation import DataValidation
from vba_builder import VBA_SOURCE, build_vba_project

BLEU_EN_TETE = "1F4E79"
BLEU_CLAIR   = "D6E4F0"
GRIS_CLAIR   = "F2F2F2"
BLANC        = "FFFFFF"
JAUNE_SAISIE = "FFFACD"

def thin_border(top=True, bottom=True, left=True, right=True) -> Border:
    t = Side(style="thin")
    return Border(top=t if top else None, bottom=t if bottom else None,
                  left=t if left else None, right=t if right else None)

def header_fill(color=BLEU_EN_TETE): return PatternFill("solid", fgColor=color)
def cell_fill(color):                return PatternFill("solid", fgColor=color)

def apply_merged_border(ws, cell_range: str) -> None:
    """Applique une bordure fine sur toutes les cellules périphériques d'une plage fusionnée."""
    min_col, min_row, max_col, max_row = range_boundaries(cell_range)
    t = Side(style="thin")
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            ws.cell(row=row, column=col).border = Border(
                top=t    if row == min_row else None,
                bottom=t if row == max_row else None,
                left=t   if col == min_col else None,
                right=t  if col == max_col else None,
            )

CONFIG_AF = {"nom_prenom":"Dupont Marie","adresse":"12 Rue des Lilas, 09000 Foix",
             "telephone":"06 12 34 56 78","email":"marie.dupont@email.fr","fait_a":"Foix"}
ENFANTS_EXEMPLE = [("Martin Lucas","Mme Bernard Sophie"),("Petit Emma","M. Leclerc Paul")]

def build_configuration_sheet(wb):
    ws = wb.create_sheet("Configuration")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 5
    ws.merge_cells("A1:B1"); c=ws["A1"]; c.value="⚙  CONFIGURATION"
    c.font=Font(bold=True,size=14,color=BLANC); c.fill=header_fill(BLEU_EN_TETE)
    c.alignment=Alignment(horizontal="center",vertical="center"); apply_merged_border(ws,"A1:B1"); ws.row_dimensions[1].height=28
    ws.row_dimensions[2].height=6
    ws.merge_cells("A3:B3"); c=ws["A3"]; c.value="Informations fixes – Assistant(e) Familial(e)"
    c.font=Font(bold=True,size=11,color=BLANC); c.fill=header_fill("2E75B6")
    c.alignment=Alignment(horizontal="left",vertical="center",indent=1); apply_merged_border(ws,"A3:B3"); ws.row_dimensions[3].height=22
    for i,(label,value) in enumerate([("Nom / Prénom AF",CONFIG_AF["nom_prenom"]),("Adresse",CONFIG_AF["adresse"]),
            ("Téléphone",CONFIG_AF["telephone"]),("Email",CONFIG_AF["email"]),("Fait à (ville)",CONFIG_AF["fait_a"])]):
        row=4+i
        lbl=ws.cell(row=row,column=1,value=label); lbl.font=Font(bold=True,size=10); lbl.fill=cell_fill(BLEU_CLAIR); lbl.border=thin_border(); lbl.alignment=Alignment(vertical="center",indent=1)
        val=ws.cell(row=row,column=2,value=value); val.font=Font(size=10); val.fill=cell_fill(JAUNE_SAISIE); val.border=thin_border(); val.alignment=Alignment(vertical="center",indent=1)
        ws.row_dimensions[row].height=18
    ws.row_dimensions[9].height=8
    ws.merge_cells("A10:B10"); c=ws["A10"]; c.value="Enfants accueillis et référents sociaux"
    c.font=Font(bold=True,size=11,color=BLANC); c.fill=header_fill("2E75B6")
    c.alignment=Alignment(horizontal="left",vertical="center",indent=1); apply_merged_border(ws,"A10:B10"); ws.row_dimensions[10].height=22
    for col,h in enumerate(["Enfant (Nom Prénom)","Référent social"],start=1):
        cell=ws.cell(row=11,column=col,value=h); cell.font=Font(bold=True,size=10,color=BLANC)
        cell.fill=header_fill(BLEU_EN_TETE); cell.border=thin_border(); cell.alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[11].height=20
    for i,(enfant,referent) in enumerate(ENFANTS_EXEMPLE):
        row=12+i; fill=cell_fill(BLANC if i%2==0 else GRIS_CLAIR)
        for col,val in ((1,enfant),(2,referent)):
            c=ws.cell(row=row,column=col,value=val); c.fill=fill; c.border=thin_border(); c.alignment=Alignment(vertical="center",indent=1); c.font=Font(size=10)
        ws.row_dimensions[row].height=18
    for row in range(12+len(ENFANTS_EXEMPLE),30):
        fill=cell_fill(BLANC if row%2==0 else GRIS_CLAIR)
        for col in (1,2):
            c=ws.cell(row=row,column=col,value=""); c.fill=fill; c.border=thin_border(); c.alignment=Alignment(vertical="center",indent=1)
        ws.row_dimensions[row].height=18
    ws.merge_cells("A31:B31"); note=ws["A31"]
    note.value="ℹ  Ajoutez autant de lignes enfant/référent que nécessaire (lignes 12 à 29)."
    note.font=Font(italic=True,size=9,color="666666"); note.alignment=Alignment(horizontal="left",indent=1)

def build_formulaire_sheet(wb):
    ws=wb.create_sheet("Formulaire",0)
    for col,w in {"A":8,"B":30,"C":35,"D":15,"E":14}.items(): ws.column_dimensions[col].width=w
    # Titre
    ws.merge_cells("A1:E1"); t=ws["A1"]; t.value="JUSTIFICATIFS DE VÊTURE – ALLOCATION MENSUELLE"
    t.font=Font(bold=True,size=14,color=BLANC); t.fill=header_fill(BLEU_EN_TETE)
    t.alignment=Alignment(horizontal="center",vertical="center"); apply_merged_border(ws,"A1:E1"); ws.row_dimensions[1].height=30
    # Sous-titre
    ws.merge_cells("A2:E2"); st=ws["A2"]
    st.value="Pour être pris en charge, les justificatifs de vêture doivent être présentés mensuellement et par enfant."
    st.font=Font(italic=True,size=9,color="444444"); st.fill=cell_fill(BLEU_CLAIR)
    st.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); apply_merged_border(ws,"A2:E2"); ws.row_dimensions[2].height=30
    ws.row_dimensions[3].height=8
    # Section AF
    ws.merge_cells("A4:E4"); s1=ws["A4"]; s1.value="IDENTITÉ DE L'ASSISTANT(E) FAMILIAL(E)"
    s1.font=Font(bold=True,size=11,color=BLANC); s1.fill=header_fill("2E75B6")
    s1.alignment=Alignment(horizontal="left",vertical="center",indent=1); apply_merged_border(ws,"A4:E4"); ws.row_dimensions[4].height=22
    for row,label,formula in [(5,"Nom / Prénom :","=Configuration!B4"),(6,"Adresse :","=Configuration!B5"),
                               (7,"Téléphone :","=Configuration!B6"),(8,"Email :","=Configuration!B7")]:
        lbl=ws.cell(row=row,column=2,value=label); lbl.font=Font(bold=True,size=10); lbl.fill=cell_fill(BLEU_CLAIR); lbl.border=thin_border(); lbl.alignment=Alignment(vertical="center",indent=1)
        rng=f"C{row}:E{row}"; ws.merge_cells(rng); val=ws.cell(row=row,column=3,value=formula)
        val.font=Font(size=10); val.fill=cell_fill(BLANC); val.alignment=Alignment(vertical="center",indent=1); apply_merged_border(ws,rng); ws.row_dimensions[row].height=18
    ws.row_dimensions[9].height=8
    # Section Enfant
    ws.merge_cells("A10:E10"); s2=ws["A10"]; s2.value="IDENTITÉ DE L'ENFANT ACCUEILLI ET CONCERNÉ PAR L'ALLOCATION"
    s2.font=Font(bold=True,size=11,color=BLANC); s2.fill=header_fill("2E75B6")
    s2.alignment=Alignment(horizontal="left",vertical="center",indent=1); apply_merged_border(ws,"A10:E10"); ws.row_dimensions[10].height=22
    lbl_e=ws.cell(row=11,column=2,value="Enfant accueilli :"); lbl_e.font=Font(bold=True,size=10); lbl_e.fill=cell_fill(BLEU_CLAIR); lbl_e.border=thin_border(); lbl_e.alignment=Alignment(vertical="center",indent=1)
    ws.merge_cells("C11:E11"); ve=ws["C11"]; ve.font=Font(size=10,color="000080"); ve.fill=cell_fill(JAUNE_SAISIE); ve.alignment=Alignment(vertical="center",indent=1); apply_merged_border(ws,"C11:E11"); ws.row_dimensions[11].height=20
    dv=DataValidation(type="list",formula1="Configuration!$A$12:$A$29",allow_blank=True,showDropDown=False,
                      showErrorMessage=True,errorTitle="Valeur invalide",error="Veuillez sélectionner un enfant depuis la liste.")
    ws.add_data_validation(dv); dv.add(ws["C11"])
    lbl_r=ws.cell(row=12,column=2,value="Référent social :"); lbl_r.font=Font(bold=True,size=10); lbl_r.fill=cell_fill(BLEU_CLAIR); lbl_r.border=thin_border(); lbl_r.alignment=Alignment(vertical="center",indent=1)
    ws.merge_cells("C12:E12"); vr=ws["C12"]; vr.value='=IFERROR(INDEX(Configuration!$B$12:$B$29,MATCH(C11,Configuration!$A$12:$A$29,0)),"")'; vr.font=Font(size=10,italic=True,color="444444"); vr.fill=cell_fill(BLANC); vr.alignment=Alignment(vertical="center",indent=1); apply_merged_border(ws,"C12:E12"); ws.row_dimensions[12].height=20
    ws.row_dimensions[13].height=8
    # Section Factures
    ws.merge_cells("A14:E14"); s3=ws["A14"]; s3.value="DÉTAIL DES FACTURES"
    s3.font=Font(bold=True,size=11,color=BLANC); s3.fill=header_fill("2E75B6")
    s3.alignment=Alignment(horizontal="left",vertical="center",indent=1); apply_merged_border(ws,"A14:E14"); ws.row_dimensions[14].height=22
    for col_idx,header in enumerate(["N°","Type d'Achat","Fournisseurs","Date","Coût (€)"],start=1):
        cell=ws.cell(row=15,column=col_idx,value=header); cell.font=Font(bold=True,size=10,color=BLANC)
        cell.fill=header_fill(BLEU_EN_TETE); cell.border=thin_border(); cell.alignment=Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[15].height=20
    for i in range(10):
        row=16+i; n=ws.cell(row=row,column=1,value=i+1); n.font=Font(size=9,color="888888"); n.fill=cell_fill(GRIS_CLAIR); n.border=thin_border(); n.alignment=Alignment(horizontal="center",vertical="center")
        fill=cell_fill(BLANC if i%2==0 else GRIS_CLAIR)
        for col in (2,3,4,5):
            c=ws.cell(row=row,column=col,value=""); c.fill=fill; c.border=thin_border(); c.alignment=Alignment(vertical="center",indent=1)
            if col==4: c.number_format="DD/MM/YYYY"
            if col==5: c.number_format='#,##0.00 "\u20ac"'
        ws.row_dimensions[row].height=18
    rT=26; ws.merge_cells(f"A{rT}:D{rT}"); tl=ws[f"A{rT}"]; tl.value="TOTAL"; tl.font=Font(bold=True,size=10,color=BLANC); tl.fill=header_fill(BLEU_EN_TETE); tl.alignment=Alignment(horizontal="right",vertical="center",indent=1); apply_merged_border(ws,f"A{rT}:D{rT}")
    tv=ws.cell(row=rT,column=5); tv.value="=SUM(E16:E25)"; tv.font=Font(bold=True,size=10); tv.fill=cell_fill(BLEU_CLAIR); tv.border=thin_border(); tv.number_format='#,##0.00 "\u20ac"'; tv.alignment=Alignment(horizontal="center",vertical="center"); ws.row_dimensions[rT].height=20
    ws.row_dimensions[27].height=10
    # Section Signature
    ws.merge_cells("A28:E28"); s4=ws["A28"]; s4.value="SIGNATURE"; s4.font=Font(bold=True,size=11,color=BLANC); s4.fill=header_fill("2E75B6"); s4.alignment=Alignment(horizontal="left",vertical="center",indent=1); apply_merged_border(ws,"A28:E28"); ws.row_dimensions[28].height=22
    ws.merge_cells("A29:E29"); ss=ws["A29"]; ss.value='=CONCATENATE("Je soussigné(e) ",Configuration!B4," certifie avoir réglé les achats ci-dessus listés.")'; ss.font=Font(size=10,italic=True); ss.fill=cell_fill(BLANC); ss.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True,indent=1); apply_merged_border(ws,"A29:E29"); ws.row_dimensions[29].height=22
    ws.row_dimensions[30].height=8
    lf=ws.cell(row=31,column=2,value="Fait à :"); lf.font=Font(bold=True,size=10); lf.fill=cell_fill(BLEU_CLAIR); lf.border=thin_border(); lf.alignment=Alignment(vertical="center",indent=1)
    vf=ws.cell(row=31,column=3,value="=Configuration!B8"); vf.font=Font(size=10); vf.fill=cell_fill(BLANC); vf.border=thin_border(); vf.alignment=Alignment(vertical="center",indent=1)
    ll=ws.cell(row=31,column=4,value="Le :"); ll.font=Font(bold=True,size=10); ll.fill=cell_fill(BLEU_CLAIR); ll.border=thin_border(); ll.alignment=Alignment(vertical="center",indent=1)
    vl=ws.cell(row=31,column=5,value="=TODAY()"); vl.font=Font(size=10); vl.fill=cell_fill(BLANC); vl.border=thin_border(); vl.number_format="DD/MM/YYYY"; vl.alignment=Alignment(horizontal="center",vertical="center"); ws.row_dimensions[31].height=20
    lm=ws.cell(row=32,column=2,value="M. / Mme :"); lm.font=Font(bold=True,size=10); lm.fill=cell_fill(BLEU_CLAIR); lm.border=thin_border(); lm.alignment=Alignment(vertical="center",indent=1)
    ws.merge_cells("C32:E32"); vm=ws["C32"]; vm.fill=cell_fill(JAUNE_SAISIE); vm.alignment=Alignment(vertical="center",indent=1); apply_merged_border(ws,"C32:E32"); ws.row_dimensions[32].height=28
    ws.row_dimensions[33].height=8
    ls=ws.cell(row=34,column=2,value="Signature :"); ls.font=Font(bold=True,size=10); ls.fill=cell_fill(BLEU_CLAIR); ls.border=thin_border(); ls.alignment=Alignment(vertical="center",indent=1)
    ws.merge_cells("C34:E36"); sz=ws["C34"]; sz.fill=cell_fill(BLANC); apply_merged_border(ws,"C34:E36")
    for row in (34,35,36): ws.row_dimensions[row].height=20
    ws.row_dimensions[37].height=8
    ws.row_dimensions[38].height=36
    ws.row_dimensions[39].height=8
    ws.freeze_panes="A3"
    from openpyxl.worksheet.page import PageMargins
    ws.page_setup.paperSize=ws.PAPERSIZE_A4; ws.page_setup.orientation=ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToPage=True; ws.page_setup.fitToHeight=1; ws.page_setup.fitToWidth=1
    ws.print_area="A1:E36"
    ws.page_margins=PageMargins(left=0.5,right=0.5,top=0.75,bottom=0.75,header=0.3,footer=0.3)

DRAWING_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>1</xdr:col><xdr:colOff>114300</xdr:colOff><xdr:row>37</xdr:row><xdr:rowOff>57150</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>38</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp macro="AjouterFactures" textlink="">
      <xdr:nvSpPr>
        <xdr:cNvPr id="1" name="BtnAjouterFactures"/>
        <xdr:cNvSpPr><a:spLocks noGrp="1"/></xdr:cNvSpPr>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>
        <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="375623"/></a:solidFill>
        <a:ln w="19050"><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></a:ln>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:lstStyle/>
        <a:p><a:pPr algn="ctr"/><a:r>
          <a:rPr lang="fr-FR" b="1" sz="1100" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr>
          <a:t>Ajouter des factures</a:t>
        </a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>3</xdr:col><xdr:colOff>114300</xdr:colOff><xdr:row>37</xdr:row><xdr:rowOff>57150</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>38</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:sp macro="GenererPDF" textlink="">
      <xdr:nvSpPr>
        <xdr:cNvPr id="2" name="BtnGenererPDF"/>
        <xdr:cNvSpPr><a:spLocks noGrp="1"/></xdr:cNvSpPr>
      </xdr:nvSpPr>
      <xdr:spPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>
        <a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>
        <a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill>
        <a:ln w="19050"><a:solidFill><a:srgbClr val="2E75B6"/></a:solidFill></a:ln>
      </xdr:spPr>
      <xdr:txBody>
        <a:bodyPr anchor="ctr"/>
        <a:lstStyle/>
        <a:p><a:pPr algn="ctr"/><a:r>
          <a:rPr lang="fr-FR" b="1" sz="1100" dirty="0"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr>
          <a:t>Generer le PDF</a:t>
        </a:r></a:p>
      </xdr:txBody>
    </xdr:sp>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>
"""

def save_as_xlsm(wb: Workbook, output_path: str) -> None:
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    # Trouver le fichier de la feuille Formulaire
    import re
    with zipfile.ZipFile(buf, "r") as zf:
        names = zf.namelist()
        members = {n: zf.read(n) for n in names}
    wb_xml = members["xl/workbook.xml"].decode("utf-8")
    wb_rels_xml = members["xl/_rels/workbook.xml.rels"].decode("utf-8")
    m = re.search(r'<sheet[^>]+name="Formulaire"[^>]+r:id="(rId\d+)"', wb_xml)
    sheet_filename = "sheet1.xml"
    if m:
        rid = m.group(1)
        m2 = re.search(rf'Id="{rid}"[^>]+Target="([^"]+)"', wb_rels_xml)
        if m2: sheet_filename = os.path.basename(m2.group(1))
    vba_bytes = build_vba_project({"GestionVeture": VBA_SOURCE})
    # Patcher Content_Types
    ct = members["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace('ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"',
                    'ContentType="application/vnd.ms-excel.sheet.macroEnabled.12+xml"')
    if "vbaProject" not in ct:
        ct = ct.replace("</Types>",'<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')
    members["[Content_Types].xml"] = ct.encode("utf-8")
    # Patcher workbook rels
    wr = wb_rels_xml
    if "vbaProject" not in wr:
        wr = wr.replace("</Relationships>",'<Relationship Id="rIdVBA" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')
    members["xl/_rels/workbook.xml.rels"] = wr.encode("utf-8")
    members["xl/vbaProject.bin"] = vba_bytes
    members["xl/drawings/drawing1.xml"] = DRAWING_XML.encode("utf-8")
    members["xl/drawings/_rels/drawing1.xml.rels"] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
    # Relation feuille → drawing
    rels_key = f"xl/worksheets/_rels/{sheet_filename}.rels"
    draw_rel = '<Relationship Id="rIdDraw1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
    if rels_key in members:
        sr = members[rels_key].decode("utf-8")
        if "drawing" not in sr: sr = sr.replace("</Relationships>", draw_rel + "</Relationships>")
    else:
        sr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + draw_rel + "</Relationships>"
    members[rels_key] = sr.encode("utf-8")
    # Ajouter <drawing> dans le worksheet
    ws_key = f"xl/worksheets/{sheet_filename}"
    if ws_key in members:
        wx = members[ws_key].decode("utf-8")
        # Assurer namespace r:
        if 'xmlns:r=' not in wx:
            wx = wx.replace("<worksheet ", '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
        if "<drawing " not in wx:
            wx = wx.replace("</worksheet>", '<drawing r:id="rIdDraw1"/></worksheet>')
        members[ws_key] = wx.encode("utf-8")
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
        for name, data in members.items():
            zf_out.writestr(name, data)
    print(f"✅  Classeur Excel généré : {output_path}")

def generate(output_path: str) -> None:
    wb = Workbook()
    if "Sheet" in wb.sheetnames: del wb["Sheet"]
    build_formulaire_sheet(wb)
    build_configuration_sheet(wb)
    save_as_xlsm(wb, output_path)

def main() -> None:
    parser = argparse.ArgumentParser(description="Génère le classeur Excel Justificatifs de Vêture (.xlsm).")
    parser.add_argument("--output", default="output/justificatif_veture.xlsm", help="Chemin du fichier Excel généré")
    args = parser.parse_args()
    generate(args.output)

if __name__ == "__main__":
    main()
