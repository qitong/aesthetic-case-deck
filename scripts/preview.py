#!/usr/bin/env python3
"""Render a PPTX to PNGs + a numbered contact sheet for visual QA (needs LibreOffice).

Usage:
    python preview.py output/case.pptx [--dpi 60] [--slides 3-8]

Writes <name>_preview/slide_NN.png and <name>_preview/contact.jpg.
LibreOffice substitutes missing fonts (e.g. 思源宋体 → Songti), so judge layout/overflow,
not exact typography.
"""
import argparse, glob, shutil, subprocess, sys
from pathlib import Path

from PIL import Image, ImageDraw


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pptx")
    ap.add_argument("--dpi", type=int, default=60)
    ap.add_argument("--slides", default=None, help="e.g. 3-8")
    ap.add_argument("--cols", type=int, default=4)
    a = ap.parse_args()
    src = Path(a.pptx).resolve()
    out = src.parent / f"{src.stem}_preview"
    out.mkdir(exist_ok=True)
    soffice = shutil.which("soffice") or shutil.which("libreoffice") or "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if not Path(soffice).exists() and not shutil.which(soffice):
        sys.exit("LibreOffice not found: brew install --cask libreoffice")
    subprocess.run([soffice, "--headless", "--norestore", "--convert-to", "pdf", "--outdir", str(out), str(src)],
                   check=True, capture_output=True)
    pdf = out / f"{src.stem}.pdf"
    for f in glob.glob(str(out / "slide-*.png")):
        Path(f).unlink()
    cmd = ["pdftoppm", "-r", str(a.dpi), "-png"]
    if a.slides:
        f, l = a.slides.split("-")
        cmd += ["-f", f, "-l", l]
    subprocess.run(cmd + [str(pdf), str(out / "slide")], check=True)
    pngs = sorted(glob.glob(str(out / "slide-*.png")))
    ims = [Image.open(p) for p in pngs]
    w, h = ims[0].size
    cols = a.cols
    rows = (len(ims) + cols - 1) // cols
    S = Image.new("RGB", (cols * w, rows * h), "white"); d = ImageDraw.Draw(S)
    for i, (p, im) in enumerate(zip(pngs, ims)):
        x, y = (i % cols) * w, (i // cols) * h
        S.paste(im, (x, y)); d.rectangle([x, y, x + 28, y + 16], fill="red")
        d.text((x + 4, y + 2), Path(p).stem.split("-")[-1].lstrip("0"), fill="white")
    S.save(out / "contact.jpg", quality=85)
    print(f"{len(pngs)} slides → {out/'contact.jpg'}  (single slides: {out}/slide-NN.png)")


if __name__ == "__main__":
    main()
