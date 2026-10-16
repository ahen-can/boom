"""Render a deck to PNGs (one per slide) plus a contact sheet, for visual checks.

    python render.py deck.pptx [--out DIR] [--width 1280]

Needs LibreOffice (soffice). PNGs come from PyMuPDF (installed with pip if missing),
falling back to pdftoppm. Prints the paths as JSON.
"""
import argparse
import importlib
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw


def to_pdf(deck: Path, out: Path) -> Path:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise SystemExit("LibreOffice (soffice) not found: render skipped. Check the deck in PowerPoint.")
    profile = Path(tempfile.mkdtemp(prefix="lo_profile_"))
    subprocess.run([soffice, f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "pdf",
                    "--outdir", str(out), str(deck)], check=True, capture_output=True, timeout=240)
    pdf = out / (deck.stem + ".pdf")
    if not pdf.is_file():
        raise SystemExit("LibreOffice did not produce a PDF")
    return pdf


def load_fitz():
    try:
        return importlib.import_module("pymupdf")
    except ImportError:
        r = subprocess.run([sys.executable, "-m", "pip", "install", "--quiet", "pymupdf"],
                           capture_output=True, text=True, timeout=180)
        if r.returncode != 0:
            return None
        importlib.invalidate_caches()
        try:
            return importlib.import_module("pymupdf")
        except ImportError:
            return None


def pdf_to_pngs(pdf: Path, out: Path, width: int) -> list[Path]:
    fitz = load_fitz()
    stem = pdf.stem
    if fitz is not None:
        paths = []
        with fitz.open(str(pdf)) as doc:
            for i, page in enumerate(doc, 1):
                zoom = width / page.rect.width
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                p = out / f"{stem}_slide{i:02d}.png"
                pix.save(str(p))
                paths.append(p)
        return paths
    if shutil.which("pdftoppm"):
        subprocess.run(["pdftoppm", "-png", "-scale-to-x", str(width), "-scale-to-y", "-1", str(pdf),
                        str(out / f"{stem}_slide")], check=True, timeout=240)
        return sorted(out.glob(f"{stem}_slide*.png"))
    raise SystemExit("Neither PyMuPDF nor pdftoppm is available; only the PDF was made.")


def contact_sheet(pngs: list[Path], dest: Path, cols: int = 3, thumb_w: int = 480) -> Path:
    thumbs = []
    for p in pngs:
        im = Image.open(p).convert("RGB")
        im.thumbnail((thumb_w, thumb_w))
        thumbs.append(im)
    th = max(t.height for t in thumbs)
    rows = (len(thumbs) + cols - 1) // cols
    pad, label = 16, 24
    sheet = Image.new("RGB", (cols * (thumb_w + pad) + pad, rows * (th + label + pad) + pad), "white")
    d = ImageDraw.Draw(sheet)
    for i, t in enumerate(thumbs):
        x = pad + (i % cols) * (thumb_w + pad)
        y = pad + (i // cols) * (th + label + pad)
        d.text((x, y), f"Slide {i + 1}", fill=(60, 60, 60))
        sheet.paste(t, (x, y + label))
        d.rectangle([x - 1, y + label - 1, x + t.width, y + label + t.height], outline=(200, 200, 200))
    sheet.save(dest)
    return dest


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("--out", help="output folder (default: <deck folder>/<deck name>_render)")
    ap.add_argument("--width", type=int, default=1280)
    a = ap.parse_args()
    deck = Path(a.deck).resolve()
    out = Path(a.out) if a.out else deck.parent / f"{deck.stem}_render"
    out.mkdir(parents=True, exist_ok=True)
    pdf = to_pdf(deck, out)
    pngs = pdf_to_pngs(pdf, out, a.width)
    sheet = contact_sheet(pngs, out / f"{deck.stem}_contact_sheet.png") if pngs else None
    print(json.dumps({"pdf": str(pdf), "slides": [str(p) for p in pngs],
                      "contact_sheet": str(sheet) if sheet else None}, indent=2))


if __name__ == "__main__":
    main()
