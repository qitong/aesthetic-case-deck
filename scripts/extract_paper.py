#!/usr/bin/env python3
"""Extract text, metadata and figures from a reference paper PDF.

Usage:
    python extract_paper.py paper.pdf --out work/paper
    python extract_paper.py paper.pdf --out work/paper --crop 7 0.08 0.10 0.92 0.55 --name fig5
        (crop page 7, region given as fractions x0 y0 x1 y1 of the page → figs/fig5.png)

Writes:
    paper.md          full text with page markers (read this, not the PDF, to save context)
    meta.json         title / DOI / journal guess / page count
    figs/             embedded raster figures (p<page>_<n>.png, small icons skipped)
    pages/            each page rendered at 110 dpi (use for vector figures + --crop)
    figs_index.jpg    contact sheet of extracted figures
"""
import argparse, json, re
from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image, ImageDraw


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--out", default="work/paper")
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--min-px", type=int, default=250, help="skip images smaller than this (logos/icons)")
    ap.add_argument("--crop", nargs=5, metavar=("PAGE", "X0", "Y0", "X1", "Y1"))
    ap.add_argument("--name", default=None)
    a = ap.parse_args()
    out = Path(a.out); (out / "figs").mkdir(parents=True, exist_ok=True); (out / "pages").mkdir(exist_ok=True)
    doc = fitz.open(a.pdf)

    if a.crop:
        pg = int(a.crop[0]) - 1
        x0, y0, x1, y1 = map(float, a.crop[1:])
        r = doc[pg].rect
        clip = fitz.Rect(r.x0 + x0 * r.width, r.y0 + y0 * r.height, r.x0 + x1 * r.width, r.y0 + y1 * r.height)
        pix = doc[pg].get_pixmap(dpi=300, clip=clip)
        dst = out / "figs" / f"{a.name or f'crop_p{pg+1}'}.png"
        pix.save(dst)
        print(f"→ {dst} ({pix.width}×{pix.height})")
        return

    text = []
    for i, page in enumerate(doc, 1):
        text.append(f"\n\n<!-- page {i} -->\n" + page.get_text("text"))
        page.get_pixmap(dpi=a.dpi).save(out / "pages" / f"p{i:02d}.png")
    full = "".join(text)
    (out / "paper.md").write_text(full, encoding="utf-8")

    doi = re.search(r"10\.\d{4,9}/[^\s\"<>]+", full)
    meta = {"file": str(Path(a.pdf).resolve()), "pages": len(doc),
            "title": (doc.metadata or {}).get("title") or "",
            "author": (doc.metadata or {}).get("author") or "",
            "doi": doi.group(0).rstrip(".") if doi else None,
            "first_lines": [l for l in full.splitlines() if l.strip()][:12]}
    (out / "meta.json").write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")

    figs = []
    for pno, page in enumerate(doc, 1):
        for n, info in enumerate(page.get_images(full=True), 1):
            try:
                pix = fitz.Pixmap(doc, info[0])
                if pix.width < a.min_px or pix.height < a.min_px:
                    continue
                if pix.n - pix.alpha >= 4:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                dst = out / "figs" / f"p{pno:02d}_{n}.png"
                pix.save(dst)
                figs.append(dst)
            except Exception:
                continue
    if figs:
        tw, th, cols = 280, 220, 5
        rows = (len(figs) + cols - 1) // cols
        S = Image.new("RGB", (cols * tw, rows * (th + 22)), "white"); d = ImageDraw.Draw(S)
        for i, f in enumerate(figs):
            im = Image.open(f).convert("RGB"); im.thumbnail((tw - 8, th - 8))
            x, y = (i % cols) * tw, (i // cols) * (th + 22)
            S.paste(im, (x + 4, y + 4)); d.text((x + 6, y + th + 3), f.name, fill="black")
        S.save(out / "figs_index.jpg", quality=85)
    print(f"pages {len(doc)} · figures {len(figs)} · doi {meta['doi']}")
    print(f"→ {out/'paper.md'}\n→ {out/'figs_index.jpg' if figs else '(no raster figures – crop from pages/)'}")


if __name__ == "__main__":
    main()
