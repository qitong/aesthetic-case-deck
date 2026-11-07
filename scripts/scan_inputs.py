#!/usr/bin/env python3
"""Inventory a case-material folder: photos, papers, notes, videos.

Usage:
    python scan_inputs.py <folder> [--out work/inputs]

Writes:
    inventory.json      every file with type, size, EXIF date, guessed tags
    inventory.md        human-readable table (show this to the user to confirm mapping)
    contact_sheet_*.jpg numbered photo thumbnails – READ these to identify views/timepoints
Tags are only guesses from filenames/EXIF; always confirm with the images and the user.
"""
import argparse, json, re
from datetime import datetime
from pathlib import Path

from PIL import Image, ImageDraw, ImageOps, ExifTags

IMG = {".jpg", ".jpeg", ".png", ".heic", ".webp", ".tif", ".tiff", ".bmp"}
VID = {".mp4", ".mov", ".m4v"}
DOC = {".pdf", ".docx", ".doc", ".txt", ".md", ".xlsx", ".xls", ".csv", ".pptx"}

VIEW = [("front", r"正面|正位|front|frontal|\bF\b|0deg|0°"),
        ("oblique", r"45|斜|oblique|three.?quarter|3-4"),
        ("profile", r"侧面|侧位|profile|lateral|90|side"),
        ("smile", r"笑|smile|dynamic|动态"),
        ("upgaze", r"上视|向上看|upgaze|look.?up"),
        ("closeup", r"特写|局部|close|detail")]
PHASE = [("before", r"术前|治疗前|before|pre\b|pre[-_]|baseline|基线"),
         ("after", r"术后|治疗后|after|post\b|post[-_]|即刻"),
         ("marking", r"标记|设计|marking|design|plan")]
SIDE = [("left", r"左|\bL\b|left"), ("right", r"右|\bR\b|right")]
DATE_RE = [r"(20\d{2})[-_.年]?([01]\d)[-_.月]?([0-3]\d)", r"(20\d{2})[-_.]([1-9])[-_.]([1-9]|[12]\d|3[01])"]
REL_RE = r"(D\d+|W\d+|M\d+|\d+\s?(?:d|w|m|mo|天|周|个?月|年))"


def guess(name, table):
    return [k for k, pat in table if re.search(pat, name, re.I)]


def exif_date(im):
    try:
        ex = im.getexif()
        for k in (36867, 306):                          # DateTimeOriginal, DateTime
            v = ex.get(k) or ex.get_ifd(0x8769).get(k)
            if v:
                return datetime.strptime(str(v)[:19], "%Y:%m:%d %H:%M:%S").date().isoformat()
    except Exception:
        pass
    return None


def file_date(name):
    for pat in DATE_RE:
        m = re.search(pat, name)
        if m:
            y, mo, d = map(int, m.groups())
            try:
                return datetime(y, mo, d).date().isoformat()
            except ValueError:
                pass
    return None


def contact_sheets(photos, out, per=24, cols=6, tw=300):
    sheets = []
    for s in range(0, len(photos), per):
        chunk = photos[s:s + per]
        th = int(tw * 4 / 3)
        rows = (len(chunk) + cols - 1) // cols
        S = Image.new("RGB", (cols * tw, rows * (th + 34)), "white")
        d = ImageDraw.Draw(S)
        for i, ph in enumerate(chunk):
            try:
                im = ImageOps.exif_transpose(Image.open(ph["path"])).convert("RGB")
                im.thumbnail((tw - 8, th - 8))
            except Exception:
                continue
            x, y = (i % cols) * tw, (i // cols) * (th + 34)
            S.paste(im, (x + (tw - im.size[0]) // 2, y + (th - im.size[1]) // 2))
            d.text((x + 6, y + th + 4), f"#{ph['id']} {Path(ph['path']).name[:34]}", fill="black")
        p = out / f"contact_sheet_{s // per + 1}.jpg"
        S.save(p, quality=85)
        sheets.append(str(p))
    return sheets


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder", nargs="+")
    ap.add_argument("--out", default="work/inputs")
    a = ap.parse_args()
    out = Path(a.out); out.mkdir(parents=True, exist_ok=True)
    files = []
    for f in a.folder:
        p = Path(f).expanduser()
        files += [p] if p.is_file() else sorted(q for q in p.rglob("*") if q.is_file() and not q.name.startswith("."))
    inv = {"photos": [], "videos": [], "documents": [], "other": []}
    for q in files:
        ext = q.suffix.lower()
        rel = str(q.resolve())
        name = str(q.relative_to(q.anchor))[-120:]
        if ext in IMG:
            rec = {"id": len(inv["photos"]) + 1, "path": rel}
            try:
                im = Image.open(q)
                w, h = ImageOps.exif_transpose(im).size
                rec.update(w=w, h=h, orient="portrait" if h > w else "landscape", exif_date=exif_date(im))
            except Exception as e:
                rec["error"] = f"unreadable ({e.__class__.__name__}); HEIC needs pillow-heif"
            rec.update(name_date=file_date(name), rel_time=(re.search(REL_RE, q.stem, re.I) or [None])[0],
                       view=guess(name, VIEW), phase=guess(name, PHASE), side=guess(q.stem, SIDE),
                       folder=q.parent.name)
            inv["photos"].append(rec)
        elif ext in VID:
            inv["videos"].append({"path": rel, "mb": round(q.stat().st_size / 1e6, 1)})
        elif ext in DOC:
            kind = "paper?" if ext == ".pdf" else "notes?"
            inv["documents"].append({"path": rel, "ext": ext, "kb": q.stat().st_size // 1024, "guess": kind})
        else:
            inv["other"].append(rel)
    inv["contact_sheets"] = contact_sheets([p for p in inv["photos"] if "error" not in p], out)
    (out / "inventory.json").write_text(json.dumps(inv, ensure_ascii=False, indent=2), encoding="utf-8")
    md = ["| # | 文件 | 尺寸 | 日期(EXIF/文件名) | 视角 | 阶段 | 文件夹 |", "|---|---|---|---|---|---|---|"]
    for p in inv["photos"]:
        md.append(f"| {p['id']} | {Path(p['path']).name} | {p.get('w','?')}×{p.get('h','?')} | "
                  f"{p.get('exif_date') or '-'} / {p.get('name_date') or p.get('rel_time') or '-'} | "
                  f"{','.join(p['view']) or '?'} | {','.join(p['phase']) or '?'} | {p['folder']} |")
    md += ["", "**Documents**"] + [f"- {d['path']} ({d['kb']} KB)" for d in inv["documents"]]
    md += ["", "**Videos**"] + [f"- {v['path']} ({v['mb']} MB)" for v in inv["videos"]]
    (out / "inventory.md").write_text("\n".join(md), encoding="utf-8")
    print(f"photos {len(inv['photos'])} · documents {len(inv['documents'])} · videos {len(inv['videos'])}")
    print(f"→ {out/'inventory.md'}")
    for s in inv["contact_sheets"]:
        print(f"→ {s}")


if __name__ == "__main__":
    main()
