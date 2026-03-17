"""
generate_pdf.py
===============
Pipeline CLI pour :
  1. Exporter la feuille 'Formulaire' du classeur Excel en PDF.
  2. Convertir les images de factures (jpg, jpeg, png, heic) en pages PDF.
  3. Fusionner formulaire.pdf + factures → output/final.pdf.

Méthodes d'export Excel → PDF (par ordre de préférence) :
  - LibreOffice (cross-platform) : `libreoffice --headless --convert-to pdf`
  - Fallback manuel : message d'erreur avec instructions

Usage :
    python src/generate_pdf.py [--excel output/justificatif_veture.xlsx]
                               [--invoices invoices/]
                               [--output output/final.pdf]
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

# ---------------------------------------------------------------------------
# Support HEIC optionnel
# ---------------------------------------------------------------------------
try:
    from pillow_heif import register_heif_opener  # type: ignore

    register_heif_opener()
    HEIC_SUPPORTED = True
except ImportError:
    HEIC_SUPPORTED = False

from PIL import Image
from pypdf import PdfReader, PdfWriter

# Formats d'image supportés
IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png"}
if HEIC_SUPPORTED:
    IMAGE_EXTENSIONS.update({".heic", ".heif"})

# Taille A4 en points (72 dpi)
A4_WIDTH_PT  = 595
A4_HEIGHT_PT = 842


# ---------------------------------------------------------------------------
# Export Excel → PDF
# ---------------------------------------------------------------------------

def excel_to_pdf_libreoffice(excel_path: str, output_dir: str) -> str | None:
    """
    Convertit le fichier Excel en PDF via LibreOffice.
    Retourne le chemin du PDF généré, ou None en cas d'échec.
    """
    lo = shutil.which("libreoffice") or shutil.which("soffice")
    if lo is None:
        return None

    try:
        result = subprocess.run(
            [lo, "--headless", "--convert-to", "pdf",
             "--outdir", output_dir, excel_path],
            capture_output=True,
            text=True,
            timeout=120,
        )
        if result.returncode != 0:
            print(f"⚠  LibreOffice erreur : {result.stderr.strip()}", file=sys.stderr)
            return None

        # LibreOffice crée un fichier <nom>.pdf dans output_dir
        stem = Path(excel_path).stem
        pdf_path = os.path.join(output_dir, f"{stem}.pdf")
        return pdf_path if os.path.isfile(pdf_path) else None
    except (subprocess.TimeoutExpired, FileNotFoundError) as exc:
        print(f"⚠  Erreur LibreOffice : {exc}", file=sys.stderr)
        return None


def excel_to_pdf(excel_path: str, pdf_path: str) -> bool:
    """
    Tente d'exporter excel_path vers pdf_path.
    Retourne True si réussi, False sinon.
    """
    output_dir = str(Path(pdf_path).parent)
    os.makedirs(output_dir, exist_ok=True)

    # Tentative LibreOffice
    result = excel_to_pdf_libreoffice(excel_path, output_dir)
    if result and os.path.isfile(result):
        # Renommer si nécessaire
        if os.path.abspath(result) != os.path.abspath(pdf_path):
            os.replace(result, pdf_path)
        print(f"✅  Formulaire PDF généré : {pdf_path}")
        return True

    # Aucune méthode disponible
    print(
        "\n❌  Impossible de convertir l'Excel en PDF automatiquement.\n"
        "   → Installez LibreOffice puis relancez, OU\n"
        "   → Ouvrez le fichier Excel et exportez manuellement en PDF.\n"
        "   → Sous Windows : Excel COM est disponible (voir README).\n",
        file=sys.stderr,
    )
    return False


# ---------------------------------------------------------------------------
# Conversion image → PDF
# ---------------------------------------------------------------------------

def image_to_pdf_bytes(image_path: str) -> bytes:
    """
    Convertit une image en page PDF A4.
    Retourne les octets du PDF.
    """
    with Image.open(image_path) as img:
        # Convertir en RGB (HEIC, RGBA, etc.)
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")

        # Calculer la taille pour tenir en A4 en gardant le ratio
        img_w, img_h = img.size
        ratio = min(A4_WIDTH_PT / img_w, A4_HEIGHT_PT / img_h)
        new_w = int(img_w * ratio)
        new_h = int(img_h * ratio)
        img = img.resize((new_w, new_h), Image.LANCZOS)

        # Créer une page A4 blanche et coller l'image centrée
        page = Image.new("RGB", (A4_WIDTH_PT, A4_HEIGHT_PT), (255, 255, 255))
        x_offset = (A4_WIDTH_PT  - new_w) // 2
        y_offset = (A4_HEIGHT_PT - new_h) // 2
        page.paste(img, (x_offset, y_offset))

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp_path = tmp.name

        page.save(tmp_path, "PDF", resolution=72)
        with open(tmp_path, "rb") as f:
            data = f.read()
        os.unlink(tmp_path)
        return data


# ---------------------------------------------------------------------------
# Fusion PDF
# ---------------------------------------------------------------------------

def merge_pdfs(pdf_paths: list[str], output_path: str) -> None:
    """Fusionne plusieurs fichiers PDF en un seul."""
    writer = PdfWriter()
    for path in pdf_paths:
        writer.append(path)
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    with open(output_path, "wb") as out:
        writer.write(out)
    writer.close()


# ---------------------------------------------------------------------------
# Pipeline principal
# ---------------------------------------------------------------------------

def run_pipeline(
    excel_path: str,
    invoices_dir: str,
    output_pdf: str,
) -> None:
    output_dir = str(Path(output_pdf).parent)
    os.makedirs(output_dir, exist_ok=True)

    pdfs_to_merge: list[str] = []

    # ---- 1. Export Excel → PDF -----------------------------------------
    formulaire_pdf = os.path.join(output_dir, "formulaire.pdf")
    ok = excel_to_pdf(excel_path, formulaire_pdf)
    if ok and os.path.isfile(formulaire_pdf):
        pdfs_to_merge.append(formulaire_pdf)
    else:
        print(
            "⚠  Le formulaire PDF n'a pas pu être généré. "
            "Le PDF final ne contiendra que les factures.",
            file=sys.stderr,
        )

    # ---- 2. Conversion images → PDF ------------------------------------
    invoices_path = Path(invoices_dir)
    if not invoices_path.is_dir():
        print(f"⚠  Dossier factures introuvable : {invoices_dir}", file=sys.stderr)
    else:
        image_files = sorted(
            p for p in invoices_path.iterdir()
            if p.suffix.lower() in IMAGE_EXTENSIONS
        )
        if not image_files:
            print(f"ℹ  Aucune image trouvée dans {invoices_dir}.")
        else:
            for img_path in image_files:
                tmp_pdf = os.path.join(
                    output_dir, img_path.stem + "_facture.pdf"
                )
                try:
                    pdf_bytes = image_to_pdf_bytes(str(img_path))
                    with open(tmp_pdf, "wb") as f:
                        f.write(pdf_bytes)
                    pdfs_to_merge.append(tmp_pdf)
                    print(f"   📄  {img_path.name} → {tmp_pdf}")
                except Exception as exc:
                    print(
                        f"⚠  Impossible de convertir {img_path.name} : {exc}",
                        file=sys.stderr,
                    )

    # ---- 3. Fusion ----------------------------------------------------- 
    if not pdfs_to_merge:
        print("❌  Aucun PDF à fusionner. Arrêt.", file=sys.stderr)
        sys.exit(1)

    merge_pdfs(pdfs_to_merge, output_pdf)
    print(f"\n✅  PDF final généré : {output_pdf}  ({len(pdfs_to_merge)} pages/source(s))")


# ---------------------------------------------------------------------------
# Point d'entrée
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Génère le PDF final : formulaire + factures."
    )
    parser.add_argument(
        "--excel",
        default="output/justificatif_veture.xlsx",
        help="Chemin vers le classeur Excel (défaut : output/justificatif_veture.xlsx)",
    )
    parser.add_argument(
        "--invoices",
        default="invoices/",
        help="Dossier contenant les images de factures (défaut : invoices/)",
    )
    parser.add_argument(
        "--output",
        default="output/final.pdf",
        help="Chemin du PDF final (défaut : output/final.pdf)",
    )
    args = parser.parse_args()
    run_pipeline(args.excel, args.invoices, args.output)


if __name__ == "__main__":
    main()
