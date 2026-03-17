"""
vba_builder.py
==============
Construit un fichier vbaProject.bin (format OLE2 / CFB) contenant
un module VBA standard avec les trois macros du classeur de vêture :
  - AjouterFactures       → copie les images choisies dans invoices/
  - GenererPDF            → exporte le formulaire et fusionne les factures
  - ResetPourProchainEnfant → efface enfant, factures, images

Référence des formats :
  MS-CFB  (Compound File Binary)
  MS-OVBA (Office VBA File Format)
"""

from __future__ import annotations

import struct
from typing import Dict

# ──────────────────────────────────────────────────────────────────
# Constantes CFB
# ──────────────────────────────────────────────────────────────────
FREESECT   = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT    = 0xFFFFFFFD
SECTOR_SZ  = 512          # version 3 : secteurs de 512 octets

DE_EMPTY   = 0
DE_STORAGE = 1
DE_STREAM  = 2
DE_ROOT    = 5

DE_BLACK = 1
DE_RED   = 0

# ──────────────────────────────────────────────────────────────────
# Code source VBA
# ──────────────────────────────────────────────────────────────────
VBA_SOURCE = (
    "Option Explicit\r\n"
    "\r\n"
    "' ================================================================\r\n"
    "'  GESTION VETURE - Justificatifs de Veture - Macros\r\n"
    "' ================================================================\r\n"
    "\r\n"
    "Private Function WorkbookFolder() As String\r\n"
    "    WorkbookFolder = ThisWorkbook.Path & Application.PathSeparator\r\n"
    "End Function\r\n"
    "\r\n"
    "' ----------------------------------------------------------------\r\n"
    "'  Bouton 1 : Ajouter des factures (copie dans invoices/)\r\n"
    "' ----------------------------------------------------------------\r\n"
    "Sub AjouterFactures()\r\n"
    "    Dim fd As FileDialog\r\n"
    "    Set fd = Application.FileDialog(msoFileDialogFilePicker)\r\n"
    "    fd.Title = \"Selectionner les photos de factures\"\r\n"
    "    fd.AllowMultiSelect = True\r\n"
    "    fd.Filters.Clear\r\n"
    "    fd.Filters.Add \"Images\", \"*.jpg;*.jpeg;*.png;*.heic;*.heif;*.bmp\"\r\n"
    "    fd.Filters.Add \"Tous les fichiers\", \"*.*\"\r\n"
    "    If fd.Show = False Then Exit Sub\r\n"
    "\r\n"
    "    Dim destFolder As String\r\n"
    "    destFolder = WorkbookFolder() & \"invoices\" & Application.PathSeparator\r\n"
    "    If Dir(destFolder, vbDirectory) = \"\" Then MkDir destFolder\r\n"
    "\r\n"
    "    Dim i As Integer, count As Integer\r\n"
    "    count = 0\r\n"
    "    For i = 1 To fd.SelectedItems.Count\r\n"
    "        Dim src As String, dst As String\r\n"
    "        src = fd.SelectedItems(i)\r\n"
    "        dst = destFolder & Dir(src)\r\n"
    "        If Dir(dst) <> \"\" Then\r\n"
    "            dst = destFolder & Format(count, \"000\") & \"_\" & Dir(src)\r\n"
    "        End If\r\n"
    "        FileCopy src, dst\r\n"
    "        count = count + 1\r\n"
    "    Next i\r\n"
    "    MsgBox count & \" facture(s) ajoutee(s) dans le dossier invoices/.\", _\r\n"
    "           vbInformation, \"Factures ajoutees\"\r\n"
    "End Sub\r\n"
    "\r\n"
    "' ----------------------------------------------------------------\r\n"
    "'  Bouton 2 : Exporter en PDF (mode impression Excel natif)\r\n"
    "' ----------------------------------------------------------------\r\n"
    "Sub GenererPDF()\r\n"
    "    ' Proposer un emplacement de sauvegarde\r\n"
    "    Dim enfant As String\r\n"
    "    enfant = Trim(CStr(ThisWorkbook.Sheets(\"Formulaire\").Range(\"C11\").Value))\r\n"
    "    Dim defaultName As String\r\n"
    "    If enfant <> \"\" Then\r\n"
    "        defaultName = \"justificatif_\" & enfant & \".pdf\"\r\n"
    "    Else\r\n"
    "        defaultName = \"justificatif_veture.pdf\"\r\n"
    "    End If\r\n"
    "\r\n"
    "    Dim savePath As Variant\r\n"
    "    savePath = Application.GetSaveAsFilename( _\r\n"
    "        InitialFileName:=WorkbookFolder() & defaultName, _\r\n"
    "        FileFilter:=\"Fichiers PDF (*.pdf), *.pdf\", _\r\n"
    "        Title:=\"Enregistrer le PDF\")\r\n"
    "    If savePath = False Then Exit Sub\r\n"
    "\r\n"
    "    ' Export de la feuille Formulaire en PDF\r\n"
    "    On Error GoTo ErrExport\r\n"
    "    ThisWorkbook.Sheets(\"Formulaire\").ExportAsFixedFormat _\r\n"
    "        Type:=xlTypePDF, _\r\n"
    "        Filename:=CStr(savePath), _\r\n"
    "        Quality:=xlQualityStandard, _\r\n"
    "        IncludeDocProperties:=False, _\r\n"
    "        IgnorePrintAreas:=False, _\r\n"
    "        OpenAfterPublish:=False\r\n"
    "    On Error GoTo 0\r\n"
    "\r\n"
    "    ' Succes : ouvrir le PDF et proposer la reinitialisation\r\n"
    "    Dim q As String : q = Chr(34)\r\n"
    "    Shell \"explorer.exe \" & q & CStr(savePath) & q, vbNormalFocus\r\n"
    "\r\n"
    "    Dim rep As Integer\r\n"
    "    rep = MsgBox(\"PDF genere avec succes !\" & vbCrLf & CStr(savePath) & vbCrLf & vbCrLf & _\r\n"
    "                 \"Voulez-vous reinitialiser le formulaire pour le prochain enfant ?\" & vbCrLf & _\r\n"
    "                 \"(Efface : identite de l'enfant, detail des factures, images dans invoices/)\", _\r\n"
    "                 vbYesNo + vbQuestion, \"PDF cree\")\r\n"
    "    If rep = vbYes Then Call ResetPourProchainEnfant\r\n"
    "    Exit Sub\r\n"
    "\r\n"
    "ErrExport:\r\n"
    "    MsgBox \"Erreur lors de la creation du PDF :\" & vbCrLf & Err.Description, _\r\n"
    "           vbCritical, \"Erreur PDF\"\r\n"
    "    On Error GoTo 0\r\n"
    "End Sub\r\n"
    "\r\n"
    "' ----------------------------------------------------------------\r\n"
    "'  Reinitialiser le formulaire pour le prochain enfant\r\n"
    "' ----------------------------------------------------------------\r\n"
    "Sub ResetPourProchainEnfant()\r\n"
    "    Dim ws As Worksheet\r\n"
    "    Set ws = ThisWorkbook.Sheets(\"Formulaire\")\r\n"
    "    ' Identite de l'enfant\r\n"
    "    ws.Range(\"C11\").ClearContents\r\n"
    "    ' Detail des factures\r\n"
    "    ws.Range(\"B16:E25\").ClearContents\r\n"
    "    ' Signature M./Mme\r\n"
    "    ws.Range(\"C32\").ClearContents\r\n"
    "\r\n"
    "    ' Supprimer les images du dossier invoices/\r\n"
    "    Dim invFolder As String\r\n"
    "    invFolder = WorkbookFolder() & \"invoices\" & Application.PathSeparator\r\n"
    "    Dim del As Integer : del = 0\r\n"
    "    Dim exts As Variant\r\n"
    "    exts = Array(\"jpg\", \"jpeg\", \"png\", \"heic\", \"heif\", \"bmp\", _\r\n"
    "                 \"JPG\", \"JPEG\", \"PNG\", \"HEIC\", \"HEIF\", \"BMP\")\r\n"
    "    Dim e As Variant\r\n"
    "    For Each e In exts\r\n"
    "        Dim f As String\r\n"
    "        f = Dir(invFolder & \"*.\" & e)\r\n"
    "        Do While f <> \"\"\r\n"
    "            Kill invFolder & f\r\n"
    "            del = del + 1\r\n"
    "            f = Dir()\r\n"
    "        Loop\r\n"
    "    Next e\r\n"
    "\r\n"
    "    MsgBox \"Formulaire reinitialise !\" & vbCrLf & _\r\n"
    "           del & \" image(s) supprimee(s).\" & vbCrLf & _\r\n"
    "           \"Vous pouvez saisir les informations du prochain enfant.\", _\r\n"
    "           vbInformation, \"Reinitialise\"\r\n"
    "End Sub\r\n"
)


# ──────────────────────────────────────────────────────────────────
# VBA compression (MS-OVBA §2.4.1) – chunk non compressé
# ──────────────────────────────────────────────────────────────────
def _vba_compress(data: bytes) -> bytes:
    """
    Compresse des octets avec l'algorithme VBA (chunks bruts de 4096 octets).
    Le byte de signature 0x01 précède les chunks.
    """
    out = bytearray(b"\x01")  # SignatureByte
    pos = 0
    length = max(len(data), 1)
    while pos < length:
        chunk = data[pos: pos + 4096]
        # Compléter à exactement 4096 octets
        chunk = chunk + b"\x00" * (4096 - len(chunk))
        out += struct.pack("<H", 0x3FFD)  # en-tête chunk brut
        out += chunk
        pos += 4096
    return bytes(out)


# ──────────────────────────────────────────────────────────────────
# Enregistrements VBA dir stream (MS-OVBA §2.3)
# ──────────────────────────────────────────────────────────────────
def _rec(rec_id: int, data: bytes) -> bytes:
    return struct.pack("<HI", rec_id, len(data)) + data


def _build_dir_stream(module_names: list[str]) -> bytes:
    out = bytearray()
    w = lambda i, d: out.extend(_rec(i, d))

    # Propriétés du projet
    w(0x0001, struct.pack("<I", 1))          # SYSKIND: Win32
    w(0x0002, struct.pack("<I", 0x0409))     # LCID
    w(0x0014, struct.pack("<I", 0x0409))     # LCIDINVOKE
    w(0x0003, struct.pack("<H", 1252))       # CODEPAGE Win-1252
    w(0x0004, b"VBAProject")                 # NAME
    w(0x0005, b"")                           # DOCSTRING
    w(0x0040, b"")                           # DOCSTRINGUNICODE
    w(0x0006, b"")                           # HELPFILEPATH1
    w(0x003D, b"")                           # HELPFILEPATH2
    w(0x0007, struct.pack("<I", 0))          # HELPCONTEXT
    w(0x0008, struct.pack("<I", 0))          # LIBFLAGS
    # PROJECTVERSION : enregistrement spécial (taille fixe = 4)
    out.extend(struct.pack("<HI", 0x0009, 4))
    out.extend(struct.pack("<IH", 0x61CC15C4, 0x000E))
    w(0x000C, b"")                           # CONSTANTS
    w(0x003C, b"")                           # CONSTANTSUNICODE

    # En-tête MODULES
    out.extend(struct.pack("<HI", 0x000F, 4))
    out.extend(struct.pack("<HH", len(module_names), 0))
    w(0x0013, struct.pack("<H", 0xFFFF))     # COOKIE

    # Enregistrements par module
    for name in module_names:
        nb = name.encode("latin-1")
        nu = name.encode("utf-16-le")
        w(0x0019, nb)                        # MODULENAME
        w(0x0047, nu)                        # MODULENAMEUNICODE
        w(0x001A, nb)                        # MODULESTREAMNAME
        w(0x0032, nu)                        # MODULESTREAMNAMERECORDUNICODE
        w(0x001C, b"")                       # MODULEDOCSTRING
        w(0x0048, b"")                       # MODULEDOCSTRINGUNICODE
        w(0x0031, struct.pack("<I", 0))      # MODULEOFFSET
        out.extend(struct.pack("<HI", 0x0021, 0))  # MODULETYPE: standard
        out.extend(struct.pack("<HI", 0x002B, 0))  # MODULE terminator

    out.extend(struct.pack("<HI", 0x0010, 0))  # MODULES terminator
    return bytes(out)


def _build_project_stream(module_names: list[str]) -> bytes:
    lines = [
        'ID="{00000000-0000-0000-0000-000000000000}"',
        "Document=ThisWorkbook/&H00000000",
    ]
    for name in module_names:
        lines.append(f"Module={name}")
    lines += [
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        'CMG=""',
        'DPB=""',
        'GC=""',
    ]
    return ("\r\n".join(lines) + "\r\n").encode("latin-1")


# ──────────────────────────────────────────────────────────────────
# Entrée de répertoire CFB (128 octets)
# ──────────────────────────────────────────────────────────────────
def _dir_entry(
    name: str,
    obj_type: int,
    color: int = DE_BLACK,
    left: int = FREESECT,
    right: int = FREESECT,
    child: int = FREESECT,
    clsid: bytes = b"\x00" * 16,
    start: int = ENDOFCHAIN,
    size: int = 0,
) -> bytes:
    name_utf16 = name.encode("utf-16-le") if name else b""
    name_len   = len(name_utf16) + 2 if name else 0
    name_buf   = (name_utf16 + b"\x00" * 64)[:64]
    return (
        name_buf
        + struct.pack("<H", min(name_len, 64))
        + struct.pack("<B", obj_type)
        + struct.pack("<B", color)
        + struct.pack("<I", left)
        + struct.pack("<I", right)
        + struct.pack("<I", child)
        + clsid
        + struct.pack("<I", 0)       # StateBits
        + struct.pack("<Q", 0)       # CreatedTime
        + struct.pack("<Q", 0)       # ModifiedTime
        + struct.pack("<I", start)
        + struct.pack("<Q", size)    # SizeLow + SizeHigh
    )


# ──────────────────────────────────────────────────────────────────
# Utilitaires
# ──────────────────────────────────────────────────────────────────
def _pad(data: bytes) -> bytes:
    """Complète à un multiple de SECTOR_SZ."""
    r = len(data) % SECTOR_SZ
    if r:
        data += b"\x00" * (SECTOR_SZ - r)
    return data


def _nsectors(size: int) -> int:
    if size == 0:
        return 0
    return (size + SECTOR_SZ - 1) // SECTOR_SZ


# ──────────────────────────────────────────────────────────────────
# Point d'entrée public : construit vbaProject.bin
# ──────────────────────────────────────────────────────────────────
def build_vba_project(modules: Dict[str, str]) -> bytes:
    """
    Construit un vbaProject.bin OLE2 complet.

    modules : dict {nom_module -> source_vba}
    Retourne les octets du fichier .bin.
    """
    module_names = list(modules.keys())

    # ── Calcul des flux ───────────────────────────────────────────
    vba_proj_raw = b"\xCC\x61\x00\x00"   # stub minimal _VBA_PROJECT
    dir_raw      = _vba_compress(_build_dir_stream(module_names))
    module_raws  = {
        name: _vba_compress(src.encode("latin-1"))
        for name, src in modules.items()
    }
    project_raw   = _build_project_stream(module_names)
    projectwm_raw = b""  # toujours vide

    # ── Allocation des secteurs ───────────────────────────────────
    # Secteur 0 : FAT
    # Secteurs 1-2 : Répertoire (4 entrées × 128 = 512 octets → 1 secteur minimum
    #                mais nous avons >= 6 entrées → 2 secteurs)
    fat: list[int] = [FATSECT, 2, ENDOFCHAIN]   # s0=FAT, s1→s2, s2=END
    next_sector = 3

    stream_info: dict[str, tuple[int, int]] = {}  # name → (start, size)

    def alloc(key: str, data: bytes) -> None:
        nonlocal next_sector
        if not data:
            stream_info[key] = (ENDOFCHAIN, 0)
            return
        n = _nsectors(len(data))
        start = next_sector
        stream_info[key] = (start, len(data))
        for i in range(n):
            fat.append(next_sector + 1 if i < n - 1 else ENDOFCHAIN)
            next_sector += 1

    alloc("PROJECT",    project_raw)
    alloc("PROJECTwm",  projectwm_raw)
    alloc("_VBA_PROJECT", vba_proj_raw)
    alloc("dir",        dir_raw)
    for name, data in module_raws.items():
        alloc(name, data)

    # Compléter la FAT à 128 entrées (1 secteur)
    while len(fat) < SECTOR_SZ // 4:
        fat.append(FREESECT)

    # ── Entrées de répertoire ─────────────────────────────────────
    mod_names   = list(module_raws.keys())
    mod_indices = list(range(6, 6 + len(mod_names)))

    entries = [
        # 0 : Root Entry  (child=1=VBA)
        _dir_entry("Root Entry", DE_ROOT, child=1, start=ENDOFCHAIN, size=0),
        # 1 : VBA storage  (child=4=_VBA_PROJECT, right=2=PROJECT)
        _dir_entry("VBA", DE_STORAGE, child=4, right=2),
        # 2 : PROJECT stream  (right=3=PROJECTwm)
        _dir_entry(
            "PROJECT", DE_STREAM, right=3,
            start=stream_info["PROJECT"][0],
            size=stream_info["PROJECT"][1],
        ),
        # 3 : PROJECTwm stream
        _dir_entry(
            "PROJECTwm", DE_STREAM,
            start=stream_info["PROJECTwm"][0],
            size=stream_info["PROJECTwm"][1],
        ),
        # 4 : _VBA_PROJECT stream  (right=5=dir)
        _dir_entry(
            "_VBA_PROJECT", DE_STREAM, right=5,
            start=stream_info["_VBA_PROJECT"][0],
            size=stream_info["_VBA_PROJECT"][1],
        ),
        # 5 : dir stream  (right=premier module si présent)
        _dir_entry(
            "dir", DE_STREAM,
            right=mod_indices[0] if mod_indices else FREESECT,
            start=stream_info["dir"][0],
            size=stream_info["dir"][1],
        ),
    ]

    for i, name in enumerate(mod_names):
        right = mod_indices[i + 1] if i + 1 < len(mod_indices) else FREESECT
        entries.append(
            _dir_entry(
                name, DE_STREAM, right=right,
                start=stream_info[name][0],
                size=stream_info[name][1],
            )
        )

    # Compléter à un multiple de 4 entrées (1 secteur = 4 entrées × 128 octets)
    while len(entries) % 4:
        entries.append(_dir_entry("", DE_EMPTY))

    # ── En-tête CFB ───────────────────────────────────────────────
    header  = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"  # magic
    header += b"\x00" * 16                            # CLSID
    header += struct.pack("<H", 0x003E)               # version mineure
    header += struct.pack("<H", 0x0003)               # version majeure 3
    header += struct.pack("<H", 0xFFFE)               # ordre des octets LE
    header += struct.pack("<H", 9)                    # secteur = 2^9 = 512
    header += struct.pack("<H", 6)                    # mini-secteur = 2^6
    header += b"\x00" * 6                             # réservé
    header += struct.pack("<I", 0)                    # secteurs répertoire (v3)
    header += struct.pack("<I", 1)                    # nb secteurs FAT
    header += struct.pack("<I", 1)                    # premier secteur répertoire
    header += struct.pack("<I", 0)                    # signature transaction
    header += struct.pack("<I", 4096)                 # seuil mini-flux
    header += struct.pack("<I", ENDOFCHAIN)           # début mini-FAT (aucun)
    header += struct.pack("<I", 0)                    # nb secteurs mini-FAT
    header += struct.pack("<I", ENDOFCHAIN)           # début DIFAT (aucun)
    header += struct.pack("<I", 0)                    # nb secteurs DIFAT
    header += struct.pack("<I", 0)                    # FAT dans secteur 0
    header += struct.pack("<I", FREESECT) * 108       # DIFAT vide
    assert len(header) == 512

    # ── Assemblage ────────────────────────────────────────────────
    fat_sector = struct.pack("<" + "I" * len(fat), *fat)
    dir_data   = b"".join(entries)
    # Répertoire sur 2 secteurs (slots 1 et 2)
    dir_sec1   = dir_data[:512]
    dir_sec2   = (dir_data[512:] + b"\x00" * 512)[:512]

    raw_map = {
        "PROJECT":       project_raw,
        "PROJECTwm":     projectwm_raw,
        "_VBA_PROJECT":  vba_proj_raw,
        "dir":           dir_raw,
        **module_raws,
    }
    stream_data = bytearray()
    for key in ["PROJECT", "PROJECTwm", "_VBA_PROJECT", "dir"] + mod_names:
        stream_data += _pad(raw_map[key])

    return header + fat_sector + dir_sec1 + dir_sec2 + bytes(stream_data)


# ──────────────────────────────────────────────────────────────────
# Auto-test rapide
# ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys

    data = build_vba_project({"GestionVeture": VBA_SOURCE})
    # Vérification du magic OLE2
    assert data[:8] == b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1", "Magic OLE2 incorrect"
    # Vérification de la taille (multiple de 512)
    assert len(data) % 512 == 0, f"Taille non multiple de 512 : {len(data)}"
    out = "/tmp/test_vba.bin"
    with open(out, "wb") as f:
        f.write(data)
    print(f"✅  vbaProject.bin : {len(data)} octets → {out}")
    sys.exit(0)
