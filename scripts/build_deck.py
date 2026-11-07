#!/usr/bin/env python3
"""Build an editable medical-aesthetics case-report PPTX from a deck spec (JSON/YAML).

Usage:
    python build_deck.py deck_spec.json --out output/case.pptx

Every element is a native PowerPoint object (text boxes, shapes, charts, pictures),
so the clinician can edit wording, move markers and swap photos afterwards.
Layout catalogue and spec fields: references/layouts.md
"""
import argparse, hashlib, json, math, os, re, sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageOps
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Emu, Inches, Pt

# ----------------------------------------------------------------------------- themes
# Distilled from the two reference decks:
#   gallery  – literature deep-read style: charcoal/black gradient + silver light slides, CJK serif
#   nocturne – patient-journey case deck: flat black photo gallery, Latin serif, lavender accent
#   clinical – white, for print / journals / sponsors that dislike dark decks
THEMES = {
    "gallery": {
        "dark":  {"bg": ["2A2A2D", "060607"], "bg_angle": 315, "text": "F2F2F2", "muted": "9C9CA2",
                  "faint": "5A5A60", "card": "FFFFFF", "card_alpha": 0.06, "card_line": "4A4A50"},
        "light": {"bg": ["FAFAFA", "B9B9BD"], "bg_angle": 45, "text": "1C1C1E", "muted": "6A6A70",
                  "faint": "A8A8AE", "card": "FFFFFF", "card_alpha": 0.72, "card_line": "D4D4D8"},
        "accent": "C9B58E", "accent2": "8FA3C8", "alert": "D9534F",
        "default_tone": "dark",
        "font_cn": "Source Han Serif CN", "font_cn_bold": "Source Han Serif CN SemiBold",
        "font_en": "Source Han Serif CN",
    },
    "nocturne": {
        "dark":  {"bg": ["000000", "000000"], "bg_angle": 90, "text": "F5F5F5", "muted": "A0A0A8",
                  "faint": "505058", "card": "FFFFFF", "card_alpha": 0.05, "card_line": "3A3A40"},
        "light": {"bg": ["F1F0FA", "F1F0FA"], "bg_angle": 90, "text": "141418", "muted": "5E5E6A",
                  "faint": "B0B0BC", "card": "FFFFFF", "card_alpha": 0.85, "card_line": "D8D6EC"},
        "accent": "8C8CE0", "accent2": "F2D16B", "alert": "E04848",
        "default_tone": "dark",
        "font_cn": "Source Han Serif CN", "font_cn_bold": "Source Han Serif CN SemiBold",
        "font_en": "Constantia",
    },
    "clinical": {
        "dark":  {"bg": ["1D2733", "0F151C"], "bg_angle": 315, "text": "F4F6F8", "muted": "9AA6B2",
                  "faint": "4A5562", "card": "FFFFFF", "card_alpha": 0.06, "card_line": "3A4654"},
        "light": {"bg": ["FFFFFF", "F3F5F7"], "bg_angle": 90, "text": "1A2330", "muted": "5F6B78",
                  "faint": "C3CAD2", "card": "F5F7F9", "card_alpha": 1.0, "card_line": "DDE2E7"},
        "accent": "2F6F9F", "accent2": "C08A4A", "alert": "C0392B",
        "default_tone": "light",
        "font_cn": "Source Han Sans CN", "font_cn_bold": "Source Han Sans CN Bold",
        "font_en": "Source Han Sans CN",
    },
}

SW, SH = 13.333, 7.5          # slide size (in)
ML, MR = 0.8, 0.8             # side margins
CW = SW - ML - MR             # content width
TOP = 1.75                    # content top (below title/subtitle)
BOTTOM = 6.75                 # content bottom (above cite/footer)


# ----------------------------------------------------------------------------- helpers
def rgb(h):
    return RGBColor.from_string(h.lstrip("#").upper())


def set_alpha(shape, alpha):
    """Apply transparency to a shape's solid fill (alpha 0..1 = opacity)."""
    if alpha >= 0.999:
        return
    sf = shape._element.spPr.find(qn("a:solidFill"))
    if sf is None:
        return
    clr = sf[0]
    a = OxmlElement("a:alpha")
    a.set("val", str(int(alpha * 100000)))
    clr.append(a)


def no_line(shape):
    shape.line.fill.background()


def char_w(ch, size):
    """Approximate rendered width (in) of one character."""
    o = ord(ch)
    if o > 0x2E80:                       # CJK & full-width punctuation
        return size / 72 * 1.0
    if ch in "ilI.,:;'|!·":
        return size / 72 * 0.28
    if ch.isupper() or ch in "mwMW%":
        return size / 72 * 0.68
    return size / 72 * 0.52


def est_lines(text, width, size):
    lines = 0
    for para in str(text).split("\n"):
        para = re.sub(r"\*\*", "", para)
        w = sum(char_w(c, size) for c in para)
        lines += max(1, math.ceil(w / (width * 0.96)))
    return lines


def est_height(text, width, size, spacing=1.32):
    return est_lines(text, width, size) * size * spacing / 72


def dots_height(items, w, size, gap):
    h = 0
    for it in items:
        txt = it.get("text", "") if isinstance(it, dict) else it
        h += est_height(txt, w - 0.32, size) + gap
        if isinstance(it, dict) and it.get("sub"):
            h += est_height(it["sub"], w - 0.32, size * 0.82) + 0.03
    return h


def fit_size(candidates, measure, avail):
    """Largest font size whose measured height fits the available height."""
    for sz in candidates:
        if measure(sz) <= avail:
            return sz
    return candidates[-1]


class Deck:
    def __init__(self, spec, base_dir, out_path):
        self.spec = spec
        self.base = Path(base_dir)
        self.out = Path(out_path)
        self.cache = self.out.parent / ".deck_cache"
        self.cache.mkdir(parents=True, exist_ok=True)
        meta = spec.get("meta", {})
        self.meta = meta
        theme_name = spec.get("theme", "gallery")
        if theme_name not in THEMES:
            sys.exit(f"unknown theme {theme_name!r}; choose from {list(THEMES)}")
        self.t = json.loads(json.dumps(THEMES[theme_name]))
        self._merge_theme(spec.get("theme_overrides", {}))
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Inches(SW), Inches(SH)
        self.blank = self.prs.slide_layouts[6]
        self.page = 0
        self.warnings = []

    def _merge_theme(self, ov):
        for k, v in ov.items():
            if isinstance(v, dict) and isinstance(self.t.get(k), dict):
                self.t[k].update(v)
            else:
                self.t[k] = v

    # ---------------------------------------------------------------- paths & images
    def path(self, p):
        if not p:
            return None
        q = Path(os.path.expanduser(p))
        if not q.is_absolute():
            q = self.base / q
        if not q.exists():
            self.warnings.append(f"missing file: {q}")
            return None
        return q

    def prep_image(self, src, box_w, box_h, fit="cover", focus=(0.5, 0.45), eyebar=None,
                   max_px=2200):
        """Crop/resize an image for a box; burn in an eye bar if requested.
        Returns (path, w_in, h_in) where w/h are the placed size inside the box."""
        p = self.path(src)
        if p is None:
            return None, box_w, box_h
        key = hashlib.md5(f"{p}{p.stat().st_mtime}{box_w:.3f}{box_h:.3f}{fit}{focus}{eyebar}".encode()).hexdigest()[:16]
        out = self.cache / f"{key}.jpg"
        im = ImageOps.exif_transpose(Image.open(p)).convert("RGB")
        W, H = im.size
        if fit == "cover":
            target = box_w / box_h
            if W / H > target:                      # too wide → crop sides
                nw = int(H * target)
                x0 = int(min(max(focus[0] * W - nw / 2, 0), W - nw))
                im = im.crop((x0, 0, x0 + nw, H))
            else:                                   # too tall → crop top/bottom
                nh = int(W / target)
                y0 = int(min(max(focus[1] * H - nh / 2, 0), H - nh))
                im = im.crop((0, y0, W, y0 + nh))
            pw, ph = box_w, box_h
        else:                                       # contain
            r = min(box_w / W, box_h / H)
            pw, ph = W * r, H * r
        if eyebar:                                  # [y0, y1] relative, burned in permanently
            d = ImageDraw.Draw(im)
            y0, y1 = eyebar[0] * im.size[1], eyebar[1] * im.size[1]
            x0, x1 = (eyebar[2] * im.size[0], eyebar[3] * im.size[0]) if len(eyebar) == 4 else (0, im.size[0])
            d.rectangle([x0, y0, x1, y1], fill=(0, 0, 0))
        if max(im.size) > max_px:
            im.thumbnail((max_px, max_px), Image.LANCZOS)
        if not out.exists():
            im.save(out, "JPEG", quality=88, optimize=True)
        return out, pw, ph

    # ---------------------------------------------------------------- primitives
    def tone(self, s):
        return self.t[s.get("tone") or self.t["default_tone"]]

    def new_slide(self, s):
        slide = self.prs.slides.add_slide(self.blank)
        self.page += 1
        tn = self.tone(s)
        bg_img = s.get("bg_image")
        if bg_img and self.path(bg_img):
            p, _, _ = self.prep_image(bg_img, SW, SH, focus=tuple(s.get("bg_focus", (0.5, 0.5))))
            slide.shapes.add_picture(str(p), 0, 0, Inches(SW), Inches(SH))
            if s.get("bg_dim", 0.45):
                r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(SW), Inches(SH))
                r.fill.solid(); r.fill.fore_color.rgb = rgb("000000"); no_line(r)
                set_alpha(r, s.get("bg_dim", 0.45))
        else:
            r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(SW), Inches(SH))
            no_line(r)
            if tn["bg"][0] == tn["bg"][1]:
                r.fill.solid(); r.fill.fore_color.rgb = rgb(tn["bg"][0])
            else:
                r.fill.gradient(); r.fill.gradient_angle = tn.get("bg_angle", 315)
                st = r.fill.gradient_stops
                st[0].color.rgb = rgb(tn["bg"][0]); st[0].position = 0
                st[1].color.rgb = rgb(tn["bg"][1]); st[1].position = 1.0
            r.name = "!!bg"
        if s.get("notes"):
            slide.notes_slide.notes_text_frame.text = s["notes"]
        return slide

    def text(self, slide, x, y, w, h, text, size=16, color=None, bold=False, align="left",
             anchor="top", font=None, spacing=1.18, italic=False, name=None, tn=None):
        tn = tn or self.t["dark"]
        tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        if name:
            tb.name = name
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                              "bottom": MSO_ANCHOR.BOTTOM}[anchor]
        paras = str(text).split("\n")
        for i, para in enumerate(paras):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                           "right": PP_ALIGN.RIGHT}[align]
            p.line_spacing = spacing
            # **bold** inline emphasis
            for j, seg in enumerate(re.split(r"\*\*", para)):
                if not seg:
                    continue
                r = p.add_run()
                r.text = seg
                self.style_run(r, size, color or tn["text"], bold or (j % 2 == 1), font, italic)
        return tb

    def style_run(self, r, size, color, bold, font=None, italic=False):
        f = r.font
        f.size = Pt(size)
        f.bold = False
        f.italic = italic
        cn = font or (self.t["font_cn_bold"] if bold else self.t["font_cn"])
        en = font or self.t["font_en"]
        if bold and not font and self.t["font_en"] == self.t["font_cn"]:
            en = cn
        f.name = en
        if bold and en == self.t.get("font_en") and en != self.t["font_cn"]:
            f.bold = True
        f.color.rgb = rgb(color)
        rPr = r._r.get_or_add_rPr()
        latin = rPr.find(qn("a:latin"))
        for tag in ("a:cs", "a:ea"):
            el = rPr.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                latin.addnext(el)
            el.set("typeface", cn)

    def rect(self, slide, x, y, w, h, fill, alpha=1.0, line=None, radius=None, line_w=0.75,
             shape=None, dash=False):
        kind = shape or (MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE)
        r = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
        if radius and kind == MSO_SHAPE.ROUNDED_RECTANGLE:
            r.adjustments[0] = min(0.5, radius / min(w, h))
        if fill:
            r.fill.solid(); r.fill.fore_color.rgb = rgb(fill); set_alpha(r, alpha)
        else:
            r.fill.background()
        if line:
            r.line.color.rgb = rgb(line); r.line.width = Pt(line_w)
            if dash:
                r.line.dash_style = MSO_LINE.DASH
        else:
            no_line(r)
        r.shadow.inherit = False
        r.text_frame.text = ""
        return r

    def card(self, slide, tn, x, y, w, h):
        return self.rect(slide, x, y, w, h, tn["card"], tn["card_alpha"], tn["card_line"], radius=0.14)

    def hline(self, slide, x, y, w, color, width=0.6):
        c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y), Inches(x + w), Inches(y))
        c.line.color.rgb = rgb(color); c.line.width = Pt(width)
        return c

    def arrow(self, slide, x1, y1, x2, y2, color, width=2.0):
        c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
        c.line.color.rgb = rgb(color); c.line.width = Pt(width)
        ln = c.line._get_or_add_ln()
        tail = OxmlElement("a:tailEnd"); tail.set("type", "triangle"); tail.set("w", "med"); tail.set("len", "med")
        ln.append(tail)
        return c

    def picture(self, slide, img, x, y, w, h, fit="cover", credit=None, tn=None):
        """img: str or {path, focus, eyebar, fit, annotations, credit}. Returns placed box."""
        if isinstance(img, str):
            img = {"path": img}
        fit = img.get("fit", fit)
        p, pw, ph = self.prep_image(img.get("path"), w, h, fit, tuple(img.get("focus", (0.5, 0.45))),
                                    img.get("eyebar"))
        px, py = x + (w - pw) / 2, y + (h - ph) / 2
        tn = tn or self.t["dark"]
        if p is None:
            r = self.rect(slide, x, y, w, h, tn["card"], max(tn["card_alpha"], 0.08), tn["faint"], dash=True)
            self.text(slide, x, y + h / 2 - 0.2, w, 0.4, f"[缺图] {img.get('path', '')}", 10,
                      tn["muted"], align="center", tn=tn)
            return (x, y, w, h)
        pic = slide.shapes.add_picture(str(p), Inches(px), Inches(py), Inches(pw), Inches(ph))
        if img.get("name"):
            pic.name = img["name"]
        self.annotate(slide, (px, py, pw, ph), img.get("annotations", []))
        cr = img.get("credit", credit)
        if cr is True:
            cr = self.meta.get("photo_credit")
        if cr:
            self.text(slide, px, py + ph - 0.32, pw - 0.12, 0.25, cr, 8, "E6E6E6", align="right",
                      font=self.t["font_en"])
        return (px, py, pw, ph)

    def annotate(self, slide, box, anns):
        """Native, editable overlays positioned in image-relative coords (0..1)."""
        bx, by, bw, bh = box
        for a in anns or []:
            t = a.get("type", "dot")
            col = a.get("color", "FFFFFF")
            X = lambda v: bx + v * bw
            Y = lambda v: by + v * bh
            if t in ("dot", "point"):
                d = a.get("size", 0.11)
                o = self.rect(slide, X(a["x"]) - d / 2, Y(a["y"]) - d / 2, d, d, col, a.get("alpha", 0.95),
                              shape=MSO_SHAPE.OVAL)
            elif t in ("zone", "ellipse"):
                w, h = a.get("w", 0.12) * bw, a.get("h", 0.06) * bh
                filled = t == "zone"
                o = self.rect(slide, X(a["x"]) - w / 2, Y(a["y"]) - h / 2, w, h,
                              col if filled else None, a.get("alpha", 0.75),
                              line=None if filled else col, line_w=a.get("line_w", 1.5),
                              shape=MSO_SHAPE.OVAL, dash=a.get("dashed", not filled))
                if a.get("rotate"):
                    o.rotation = a["rotate"]
            elif t == "arrow":
                self.arrow(slide, X(a["x"]), Y(a["y"]), X(a["x2"]), Y(a["y2"]), a.get("color", self.t["alert"]),
                           a.get("width", 2.0))
            elif t == "line":
                c = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(X(a["x"])), Inches(Y(a["y"])),
                                               Inches(X(a["x2"])), Inches(Y(a["y2"])))
                c.line.color.rgb = rgb(col); c.line.width = Pt(a.get("width", 1.5))
                if a.get("dashed"):
                    c.line.dash_style = MSO_LINE.DASH
            if a.get("label"):
                lx = X(a.get("lx", a["x"] + 0.04)); ly = Y(a.get("ly", a["y"] - 0.02))
                self.text(slide, lx, ly, 2.2, 0.3, a["label"], a.get("label_size", 10), col)

    def dots(self, slide, tn, x, y, w, items, size=15, gap=0.22, max_y=BOTTOM, dot_color=None):
        """Dot bullet list. Returns final y."""
        cy = y
        for it in items:
            sub = None
            if isinstance(it, dict):
                txt, sub = it.get("text", ""), it.get("sub")
            else:
                txt = it
            h = est_height(txt, w - 0.32, size)
            d = 0.075
            self.rect(slide, x, cy + size / 72 * 0.62 - d / 2, d, d, dot_color or tn["muted"],
                      shape=MSO_SHAPE.OVAL)
            self.text(slide, x + 0.3, cy, w - 0.32, h, txt, size, tn["text"], tn=tn)
            cy += h
            if sub:
                sh = est_height(sub, w - 0.32, size * 0.82)
                self.text(slide, x + 0.3, cy + 0.03, w - 0.32, sh, sub, size * 0.82, tn["muted"], tn=tn)
                cy += sh + 0.03
            cy += gap
        if cy > max_y + 0.05:
            self.warnings.append(f"slide {self.page}: bullet text may overflow (ends at {cy:.2f}in)")
        return cy

    def floor(self, s):
        """Bottom edge for photo/figure content; leaves room for a note line when present."""
        return 6.2 if s.get("note") else BOTTOM

    def header(self, slide, s, tn):
        if s.get("title"):
            self.text(slide, ML, 0.52, CW, 0.62, s["title"], s.get("title_size", 28), tn["text"],
                      bold=True, name="!!title", tn=tn, anchor="bottom")
        if s.get("subtitle"):
            self.text(slide, ML, 1.16, CW, 0.4, s["subtitle"], 14, tn["muted"], name="!!subtitle", tn=tn)

    def chrome(self, slide, s, tn):
        """Citation line, footer disclaimer, logo, page number."""
        if s.get("cite"):
            cite = s["cite"] if isinstance(s["cite"], str) else " · ".join(s["cite"])
            self.text(slide, ML, 6.86, CW - 0.6, 0.3, cite, 7.5, tn["faint"] if tn is self.t["dark"] else tn["muted"],
                      name="zz_cite", tn=tn)
        if s.get("layout") in ("cover", "thanks", "chapter", "statement") and not s.get("force_chrome"):
            pass
        else:
            footer = s.get("footer", self.meta.get("footer"))
            if footer:
                self.text(slide, ML, 7.16, 8.5, 0.2, footer, 7, tn["faint"], tn=tn, name="zz_footer")
            if self.meta.get("page_numbers", True):
                self.text(slide, SW - MR - 0.6, 7.16, 0.6, 0.2, str(self.page), 7, tn["faint"],
                          align="right", tn=tn)
        logo = self.meta.get("logo")
        if logo and s.get("logo", True) and self.path(logo):
            lw = self.meta.get("logo_width", 1.1)
            im = Image.open(self.path(logo))
            lh = lw * im.size[1] / im.size[0]
            pos = self.meta.get("logo_pos", "top-right")
            lx = SW - MR - lw if "right" in pos else ML
            ly = 0.42 if "top" in pos else 7.05 - lh
            slide.shapes.add_picture(str(self.path(logo)), Inches(lx), Inches(ly), Inches(lw), Inches(lh))

    # ---------------------------------------------------------------- build
    def build(self):
        for i, s in enumerate(self.spec.get("slides", [])):
            if s.get("skip"):
                continue
            fn = getattr(self, "L_" + s.get("layout", "bullets"), None)
            if fn is None:
                self.warnings.append(f"slide {i+1}: unknown layout {s.get('layout')!r} → bullets")
                fn = self.L_bullets
            tn = self.tone(s)
            slide = self.new_slide(s)
            fn(slide, s, tn)
            self.chrome(slide, s, tn)
        self.out.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(self.out)
        return self.warnings

    # ================================================================= LAYOUTS
    def L_cover(self, slide, s, tn):
        if s.get("kicker"):
            self.text(slide, 1.0, 2.0, 11, 0.4, s["kicker"], 14, tn["muted"], tn=tn)
        self.text(slide, 1.0, 2.35, 11.3, 1.5, s.get("title", ""), s.get("title_size", 66), tn["text"],
                  bold=True, anchor="bottom", name="!!hero", tn=tn)
        y = 3.95
        if s.get("title_en"):
            self.text(slide, 1.0, y, 11.3, 0.45, s["title_en"], 16, tn["muted"], tn=tn,
                      font=self.t["font_en"], name="!!heroen")
            y += 0.55
        if s.get("subtitle"):
            self.text(slide, 1.0, y, 11.3, 0.45, s["subtitle"], 18, tn["text"], tn=tn)
        by = 6.1
        for line in [s.get("presenter"), s.get("affiliation"), s.get("date")]:
            if line:
                self.text(slide, 1.0, by, 11, 0.3, line, 12, tn["muted"], tn=tn)
                by += 0.32

    def L_chapter(self, slide, s, tn):
        self.text(slide, 0, 2.55, SW, 0.5, s.get("no", ""), 22, tn["muted"], align="center",
                  font=self.t["font_en"], name="!!chno", tn=tn)
        self.text(slide, 0.8, 3.05, SW - 1.6, 1.1, s.get("title", ""), 52, tn["text"], bold=True,
                  align="center", anchor="middle", name="!!chtitle", tn=tn)
        if s.get("title_en"):
            self.text(slide, 0, 4.25, SW, 0.4, s["title_en"], 13, tn["muted"], align="center",
                      font=self.t["font_en"], name="!!chen", tn=tn)

    def L_statement(self, slide, s, tn):
        self.text(slide, 1.4, 2.2, SW - 2.8, 2.4, s.get("text", s.get("title", "")), s.get("size", 32),
                  tn["text"], bold=s.get("bold", False), align="center", anchor="middle", tn=tn)
        if s.get("sub"):
            self.text(slide, 1.4, 4.75, SW - 2.8, 0.8, s["sub"], 15, tn["muted"], align="center", tn=tn)

    def L_bullets(self, slide, s, tn):
        self.header(slide, s, tn)
        top = TOP if s.get("subtitle") else 1.55
        items = s.get("bullets", [])
        avail = (5.8 if s.get("note") else BOTTOM) - top - 0.25
        size = s.get("size") or fit_size([22, 20, 18, 16, 15, 14, 13], lambda z: dots_height(items, CW - 0.4, z, 0.3), avail)
        y = self.dots(slide, tn, ML + 0.2, top + 0.25, CW - 0.4, items, size, gap=0.3)
        if s.get("note"):
            self.note(slide, tn, s["note"], max(y + 0.15, 5.9))

    def note(self, slide, tn, text, y=6.15):
        self.text(slide, ML, y, CW, 0.5, text, 12.5, tn["muted"], tn=tn, name="a9_c")

    def L_image_text(self, slide, s, tn):
        self.header(slide, s, tn)
        ratio = s.get("image_ratio", 0.56)
        side = s.get("image_side", "left")
        top = TOP if s.get("subtitle") else 1.55
        cap_h = 0.32 if s.get("caption") else 0
        iw, ih = CW * ratio, self.floor(s) - top - cap_h - 0.05
        ix = ML if side == "left" else ML + CW - iw
        img = s.get("image") or {}
        box = self.picture(slide, img, ix, top, iw, ih, fit=(img.get("fit", "contain") if isinstance(img, dict) else "contain"),
                           tn=tn)
        if s.get("caption"):
            self.text(slide, box[0], box[1] + box[3] + 0.08, box[2] if box[2] > 3 else iw, 0.3,
                      s["caption"], 9.5, tn["muted"], tn=tn, name="a1_cap")
        tx = ML + iw + 0.5 if side == "left" else ML
        tw = CW - iw - 0.5
        bullets = s.get("bullets", [])
        ty = top + 0.25
        avail = BOTTOM - ty - (0.55 if s.get("heading") else 0)
        size = s.get("size") or fit_size([20, 18, 17, 16, 15, 14, 13, 12], lambda z: dots_height(bullets, tw, z, 0.26), avail)
        if s.get("heading"):
            self.text(slide, tx, ty, tw, 0.4, s["heading"], 18, tn["text"], bold=True, tn=tn)
            ty += 0.55
        self.dots(slide, tn, tx, ty, tw, bullets, size, gap=0.26)
        if s.get("note"):
            self.note(slide, tn, s["note"], 6.35)

    def _grid(self, n, cols, x, y, w, h, gap=0.25, label_h=0.3, aspect=None):
        cols = cols or (n if n <= 4 else math.ceil(n / 2))
        rows = math.ceil(n / cols)
        cw = (w - gap * (cols - 1)) / cols
        chh = (h - gap * (rows - 1)) / rows - label_h
        if aspect:                                  # keep w/h aspect, centre the grid
            ih = min(chh, cw / aspect)
            iw = ih * aspect
        else:
            iw, ih = cw, chh
        tot_w = cols * iw + (cols - 1) * gap
        ox = x + (w - tot_w) / 2
        cells = []
        for i in range(n):
            r, c = divmod(i, cols)
            cells.append((ox + c * (iw + gap), y + r * (ih + label_h + gap), iw, ih))
        return cells

    def L_gallery(self, slide, s, tn):
        self.header(slide, s, tn)
        imgs = s.get("images", [])
        top = TOP + 0.1 if s.get("subtitle") else 1.6
        aspect = s.get("aspect")
        cells = self._grid(len(imgs), s.get("cols"), ML, top, CW, self.floor(s) - top, 0.3, 0.38, aspect)
        for im, (x, y, w, h) in zip(imgs, cells):
            box = self.picture(slide, im, x, y, w, h, fit=im.get("fit", "cover") if isinstance(im, dict) else "cover",
                               tn=tn)
            lab = im.get("label") if isinstance(im, dict) else None
            if lab:
                lx = x if s.get("label_align") == "center" else box[0]
                lw = w if s.get("label_align") == "center" else max(box[2], 2.5)
                self.text(slide, lx, box[1] + box[3] + 0.08, lw, 0.3, lab, 10.5, tn["muted"], tn=tn,
                          align=s.get("label_align", "left"))
        if s.get("note"):
            self.note(slide, tn, s["note"], 6.35)

    def L_compare(self, slide, s, tn):
        """Before/after. pairs: [{view, before, after, before_label, after_label}]"""
        self.header(slide, s, tn)
        pairs = s.get("pairs") or [{"before": s.get("before"), "after": s.get("after"), "view": s.get("view")}]
        bullets = s.get("bullets")
        aspect = s.get("aspect", 3 / 4)
        top = TOP if s.get("subtitle") else 1.55
        area_w = CW * (0.64 if bullets else 1.0)
        n = len(pairs)
        has_view = any(p.get("view") for p in pairs) and n > 1
        view_h = 0.38 if has_view else 0
        gap_pair, gap_in = 0.55, 0.12
        label_h = 0.42
        pw = (area_w - gap_pair * (n - 1)) / n
        iw = (pw - gap_in) / 2
        ih = iw / aspect
        avail = self.floor(s) - top - view_h - label_h
        if ih > avail:
            ih = avail; iw = ih * aspect; pw = iw * 2 + gap_in
        tot = n * pw + (n - 1) * gap_pair
        ox = ML + (area_w - tot) / 2
        y = top + view_h + max(0, (avail - ih) / 2) * (0 if bullets else 0.5)
        lb_b, lb_a = s.get("labels", [self.meta.get("before_label", "治疗前"), self.meta.get("after_label", "治疗后")])
        for i, p in enumerate(pairs):
            x = ox + i * (pw + gap_pair)
            if has_view and p.get("view"):
                self.text(slide, x, y - view_h, pw, 0.3, p["view"], 12, tn["muted"], align="center", tn=tn)
            for j, key in enumerate(("before", "after")):
                img = p.get(key)
                if isinstance(img, str):
                    img = {"path": img}
                img = dict(img or {})
                img.setdefault("credit", True)
                xx = x + j * (iw + gap_in)
                self.picture(slide, img, xx, y, iw, ih, tn=tn)
                lab = p.get(f"{key}_label", lb_b if j == 0 else lb_a)
                self.text(slide, xx, y + ih + 0.1, iw, 0.3, lab, 12, tn["text"] if j else tn["muted"],
                          align="center", tn=tn)
        if bullets:
            bx = ML + area_w + 0.45
            self.dots(slide, tn, bx, top + 0.35, CW - area_w - 0.45, bullets, s.get("size", 15))
        if s.get("note"):
            self.note(slide, tn, s["note"], 6.35)

    def L_journey(self, slide, s, tn):
        """Patient journey photo strip with timeline. steps: [{image, label, note}]"""
        self.header(slide, s, tn)
        steps = s.get("steps", [])
        n = max(1, len(steps))
        top = TOP + 0.1 if s.get("subtitle") else 1.65
        aspect = s.get("aspect", 3 / 4)
        has_notes = any(st.get("note") for st in steps)
        note_h = 1.0 if has_notes else 0
        gap = 0.22
        iw = (CW - gap * (n - 1)) / n
        ih = iw / aspect
        avail = self.floor(s) - top - 0.75 - note_h
        if ih > avail:
            ih = avail; iw = ih * aspect
        tot = n * iw + (n - 1) * gap
        ox = ML + (CW - tot) / 2
        for i, st in enumerate(steps):
            x = ox + i * (iw + gap)
            img = st.get("image")
            if img:
                img = {"path": img} if isinstance(img, str) else dict(img)
                img.setdefault("credit", True)
                self.picture(slide, img, x, top, iw, ih, tn=tn)
        ly = top + ih + 0.28
        self.hline(slide, ox, ly, tot, tn["faint"], 0.8)
        for i, st in enumerate(steps):
            x = ox + i * (iw + gap)
            cx = x + iw / 2
            d = 0.13
            self.rect(slide, cx - d / 2, ly - d / 2, d, d, self.t["accent"] if st.get("highlight") else tn["muted"],
                      shape=MSO_SHAPE.OVAL)
            self.text(slide, x, ly + 0.14, iw, 0.3, st.get("label", ""), 12.5, tn["text"], bold=True,
                      align="center", tn=tn)
            if st.get("note"):
                self.text(slide, x + 0.02, ly + 0.5, iw - 0.04, note_h, st["note"], 9.5, tn["muted"],
                          align="center", tn=tn)
        if s.get("note"):
            self.note(slide, tn, s["note"], 6.35)

    def L_timeline(self, slide, s, tn):
        """Treatment course without photos. events: [{date, title, text}]"""
        self.header(slide, s, tn)
        ev = s.get("events", [])
        n = max(1, len(ev))
        top = TOP + 0.2 if s.get("subtitle") else 1.8
        ly = top + 0.55
        self.hline(slide, ML, ly, CW, tn["faint"], 1.0)
        cw = (CW - 0.2 * (n - 1)) / n
        need_all = 0
        for e in ev:
            its = e.get("text") or []
            its = its if isinstance(its, list) else [its]
            need_all = max(need_all, 0.7 + (est_height(e.get("title", ""), cw - 0.4, 14) if e.get("title") else 0)
                           + dots_height(its, cw - 0.4, e.get("size", 11.5 if n >= 4 else 12.5), 0.1))
        for i, e in enumerate(ev):
            x = ML + i * (cw + 0.2)
            self.text(slide, x, top, cw, 0.35, e.get("date", ""), 13, self.t["accent"], bold=True, tn=tn)
            d = 0.16
            self.rect(slide, x, ly - d / 2, d, d, self.t["accent"] if e.get("highlight") else tn["muted"],
                      shape=MSO_SHAPE.OVAL)
            items = e.get("text") or []
            items = items if isinstance(items, list) else [items]
            ch_max = self.floor(s) - ly - 0.35
            need = 0.6 + (est_height(e.get("title", ""), cw - 0.4, 14) if e.get("title") else 0) \
                + dots_height(items, cw - 0.4, e.get("size", 11.5 if n >= 4 else 12.5), 0.1)
            ch = ch_max if s.get("fill") else min(ch_max, max(need_all, 1.6))
            self.card(slide, tn, x, ly + 0.3, cw, ch)
            yy = ly + 0.5
            if e.get("title"):
                th = est_height(e["title"], cw - 0.4, 14)
                self.text(slide, x + 0.2, yy, cw - 0.4, th, e["title"], 14, tn["text"], bold=True, tn=tn)
                yy += th + 0.12
            if e.get("text"):
                items = e["text"] if isinstance(e["text"], list) else [e["text"]]
                sz = e.get("size", 11.5 if n >= 4 else 12.5)
                self.dots(slide, tn, x + 0.2, yy, cw - 0.4, items, sz, gap=0.1, max_y=ly + 0.3 + ch)
        if s.get("note"):
            self.note(slide, tn, s["note"], 6.35)

    def L_table(self, slide, s, tn):
        self.header(slide, s, tn)
        cols, rows = s.get("columns", []), s.get("rows", [])
        top = TOP if s.get("subtitle") else 1.6
        ratios = s.get("col_widths") or [1] * len(cols)
        tw = CW - 0.6
        widths = [tw * r / sum(ratios) for r in ratios]
        bottom = 6.0 if s.get("note") else BOTTOM

        def measure(z):
            hs = [max(est_height(str(c), widths[k] - 0.2, z) for k, c in enumerate(row)) + 0.16 + z / 72 * 0.6
                  for row in rows]
            return hs, 0.45 + sum(hs) + 0.2
        size = s.get("size") or fit_size([18, 17, 16, 15, 14, 13, 12, 11, 10], lambda z: measure(z)[1], bottom - top)
        heights, total = measure(size)
        if top + total > bottom:
            self.warnings.append(f"slide {self.page}: table taller than area ({total:.2f}in); consider size/split")
        self.card(slide, tn, ML, top, CW, min(total, bottom - top + 0.4))
        x0, y = ML + 0.3, top + 0.18
        for k, c in enumerate(cols):
            self.text(slide, x0 + sum(widths[:k]), y, widths[k] - 0.2, 0.3, c, 10, tn["muted"], tn=tn)
        y += 0.42
        for ri, row in enumerate(rows):
            for k, c in enumerate(row):
                bold = k == 0 and s.get("first_col_bold", True)
                col = self.t["accent"] if (s.get("highlight_col") == k) else tn["text"]
                self.text(slide, x0 + sum(widths[:k]), y + 0.1, widths[k] - 0.2, heights[ri] - 0.1, str(c),
                          size, col, bold=bold, tn=tn)
            y += heights[ri]
            if ri < len(rows) - 1:
                self.hline(slide, x0, y, tw, tn["card_line"], 0.5)
        if s.get("note"):
            self.note(slide, tn, s["note"], max(6.15, min(top + total + 0.15, 6.3)))

    def L_kpi(self, slide, s, tn):
        self.header(slide, s, tn)
        y = TOP if s.get("subtitle") else 1.55
        if s.get("intro"):
            intro = s["intro"] if isinstance(s["intro"], str) else "\n".join(s["intro"])
            h = est_height(intro, CW, 14)
            self.text(slide, ML, y, CW, h, intro, 14, tn["text"], tn=tn, name="a1_t")
            y += h + 0.3
        stats = s.get("stats", [])
        n = max(1, len(stats))
        gap = 0.3
        cw = (CW - gap * (n - 1)) / n
        ch = 2.0
        for i, st in enumerate(stats):
            x = ML + i * (cw + gap)
            self.card(slide, tn, x, y, cw, ch)
            self.text(slide, x + 0.3, y + 0.25, cw - 0.5, 0.85, str(st.get("value", "")), s.get("value_size", 40),
                      self.t["accent"] if st.get("accent") else tn["text"], bold=True, tn=tn, font=self.t["font_en"])
            self.text(slide, x + 0.3, y + 1.15, cw - 0.5, 0.75, st.get("label", ""), 11.5, tn["muted"], tn=tn)
        if s.get("note"):
            self.note(slide, tn, s["note"], y + ch + 0.35)

    def L_cards(self, slide, s, tn):
        """Numbered / titled cards: conclusions, key claims, summary."""
        self.header(slide, s, tn)
        items = s.get("items", [])
        n = max(1, len(items))
        top = TOP + 0.1 if s.get("subtitle") else 1.65
        gap = 0.3
        cols = s.get("cols", n if n <= 4 else math.ceil(n / 2))
        rows = math.ceil(n / cols)
        cw = (CW - gap * (cols - 1)) / cols
        bottom = 6.0 if s.get("note") else BOTTOM
        ch_max = (bottom - top - gap * (rows - 1)) / rows

        def need(z):
            hs = []
            for it in items:
                h = 0.28 + (0.72 if s.get("numbered", True) or it.get("num") is not None else 0)
                if it.get("title"):
                    h += est_height(it["title"], cw - 0.6, 17) + 0.18
                b = it.get("body") or it.get("bullets")
                if isinstance(b, list):
                    h += dots_height(b, cw - 0.6, z, 0.12)
                elif b:
                    h += est_height(b, cw - 0.6, z)
                hs.append(h + 0.35)
            return max(hs) if hs else 1
        bsize = s.get("size") or fit_size([17, 16, 15, 14, 13, 12, 11], need, ch_max)
        ch = ch_max if s.get("fill") else min(ch_max, max(need(bsize), 2.2))
        for i, it in enumerate(items):
            r, c = divmod(i, cols)
            x, y = ML + c * (cw + gap), top + r * (ch + gap)
            self.card(slide, tn, x, y, cw, ch)
            yy = y + 0.28
            num = it.get("num", (i + 1) if s.get("numbered", True) else None)
            if num is not None:
                self.text(slide, x + 0.32, yy, 1.2, 0.6, str(num), 30, tn["text"], bold=True, tn=tn,
                          font=self.t["font_en"])
                yy += 0.72
            if it.get("title"):
                th = est_height(it["title"], cw - 0.6, 17)
                self.text(slide, x + 0.32, yy, cw - 0.6, th, it["title"], 17, tn["text"], bold=True, tn=tn)
                yy += th + 0.18
            body = it.get("body") or it.get("bullets")
            if isinstance(body, list):
                self.dots(slide, tn, x + 0.32, yy, cw - 0.6, body, bsize, gap=0.12, max_y=y + ch)
            elif body:
                bh = est_height(body, cw - 0.6, bsize)
                self.text(slide, x + 0.32, yy, cw - 0.6, bh, body, bsize, tn["muted"], tn=tn)
                if yy + bh > y + ch:
                    self.warnings.append(f"slide {self.page}: card {i+1} text may overflow")
        if s.get("note"):
            self.note(slide, tn, s["note"], 6.2)

    def L_two_col(self, slide, s, tn):
        self.header(slide, s, tn)
        top = TOP + 0.1 if s.get("subtitle") else 1.65
        gap = 0.35
        cw = (CW - gap) / 2
        bottom = 6.0 if s.get("note") else BOTTOM
        cols = [s.get("left", {}), s.get("right", {})]
        need = lambda z: max(dots_height(c.get("bullets", []), cw - 0.7, z, 0.2) for c in cols) + 1.3
        size = s.get("size") or fit_size([19, 18, 17, 16, 15, 14, 13, 12], need, bottom - top)
        ch = bottom - top if s.get("fill") else min(bottom - top, max(need(size), 2.6))
        for i, col in enumerate(cols):
            x = ML + i * (cw + gap)
            self.card(slide, tn, x, top, cw, ch)
            self.text(slide, x + 0.35, top + 0.3, cw - 0.7, 0.45, col.get("heading", ""), 18, tn["text"],
                      bold=True, tn=tn)
            self.dots(slide, tn, x + 0.35, top + 0.95, cw - 0.7, col.get("bullets", []), size,
                      gap=0.2, max_y=bottom)
        if s.get("note"):
            self.note(slide, tn, s["note"], min(6.2, top + ch + 0.3))

    def L_letter_rows(self, slide, s, tn):
        """PICO / ABCD-style rows: rows: [{letter, name, text}]"""
        self.header(slide, s, tn)
        rows = s.get("rows", [])
        top = TOP if s.get("subtitle") else 1.6
        bottom = 6.0 if s.get("note") else BOTTOM
        rh = (bottom - top) / max(1, len(rows))
        for i, r in enumerate(rows):
            y = top + i * rh
            self.text(slide, ML + 0.1, y, 0.9, rh, r.get("letter", ""), 38, tn["text"], bold=True,
                      anchor="middle", tn=tn, font=self.t["font_en"])
            self.text(slide, ML + 1.1, y, 1.3, rh, r.get("name", ""), 13, tn["muted"], anchor="middle", tn=tn)
            self.text(slide, ML + 2.5, y, CW - 2.6, rh, r.get("text", ""), 15, tn["text"], anchor="middle", tn=tn,
                      bold=r.get("bold", False))
            if i < len(rows) - 1:
                self.hline(slide, ML + 0.1, y + rh, CW - 0.2, tn["card_line"], 0.5)
        if s.get("note"):
            self.note(slide, tn, s["note"], 6.2)

    def L_patient(self, slide, s, tn):
        """Patient profile: 1–3 photos + facts list. facts: [[k, v], …] or [{k, v}]"""
        self.header(slide, s, tn)
        imgs = s.get("images", [])
        top = TOP if s.get("subtitle") else 1.5
        img_w = CW * s.get("image_ratio", 0.58 if len(imgs) >= 2 else 0.42)
        cells = self._grid(len(imgs), len(imgs), ML, top, img_w, self.floor(s) - top, 0.18, 0.05, s.get("aspect", 3 / 4))
        for im, (x, y, w, h) in zip(imgs, cells):
            im = {"path": im} if isinstance(im, str) else dict(im)
            im.setdefault("credit", True)
            self.picture(slide, im, x, y, w, h, tn=tn)
        fx = ML + img_w + 0.55
        fw = CW - img_w - 0.55
        y = top + 0.15
        size = s.get("size", 13.5)
        for f in s.get("facts", []):
            k, v = (f["k"], f["v"]) if isinstance(f, dict) else f
            vh = est_height(v, fw - 1.35, size)
            self.text(slide, fx, y, 1.25, vh, k, size - 1.5, tn["muted"], tn=tn)
            self.text(slide, fx + 1.35, y, fw - 1.35, vh, v, size, tn["text"], tn=tn)
            y += vh + 0.14
            self.hline(slide, fx, y - 0.07, fw, tn["card_line"], 0.4)
            y += 0.04
        if y > BOTTOM:
            self.warnings.append(f"slide {self.page}: patient facts overflow ({y:.2f}in)")
        if s.get("note"):
            self.note(slide, tn, s["note"], 6.35)

    def _chart(self, slide, tn, ch_spec, x, y, w, h):
        ctype = ch_spec.get("type", "bar")
        cd = CategoryChartData()
        cd.categories = ch_spec.get("categories", [])
        for se in ch_spec.get("series", []):
            cd.add_series(se["name"], se["values"])
        xl = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED, "hbar": XL_CHART_TYPE.BAR_CLUSTERED,
              "line": XL_CHART_TYPE.LINE_MARKERS, "radar": XL_CHART_TYPE.RADAR_MARKERS,
              "radar_filled": XL_CHART_TYPE.RADAR_FILLED}[ctype]
        gf = slide.shapes.add_chart(xl, Inches(x), Inches(y), Inches(w), Inches(h), cd)
        ch = gf.chart
        palette = ch_spec.get("colors") or [tn["muted"], self.t["accent"], self.t["accent2"], tn["text"]]
        ch.font.size = Pt(10); ch.font.color.rgb = rgb(tn["muted"]); ch.font.name = self.t["font_cn"]
        ch.has_legend = len(ch_spec.get("series", [])) > 1
        if ch.has_legend:
            ch.legend.position = XL_LEGEND_POSITION.BOTTOM
            ch.legend.include_in_layout = False
            ch.legend.font.size = Pt(11); ch.legend.font.color.rgb = rgb(tn["text"])
        plot = ch.plots[0]
        for i, ser in enumerate(plot.series):
            col = rgb(palette[i % len(palette)])
            if ctype in ("bar", "hbar"):
                ser.format.fill.solid(); ser.format.fill.fore_color.rgb = col
            else:
                ser.format.line.color.rgb = col; ser.format.line.width = Pt(2.25)
                try:
                    ser.marker.format.fill.solid(); ser.marker.format.fill.fore_color.rgb = col
                    ser.marker.format.line.color.rgb = col; ser.marker.size = 7
                except Exception:
                    pass
            if ch_spec.get("data_labels") and ctype in ("bar", "hbar", "line"):
                ser.data_labels.show_value = True
                ser.data_labels.font.size = Pt(10); ser.data_labels.font.color.rgb = rgb(tn["text"])
        if ctype in ("bar", "hbar"):
            plot.gap_width = 80
        try:
            va = ch.value_axis
            if "max" in ch_spec: va.maximum_scale = ch_spec["max"]
            if "min" in ch_spec: va.minimum_scale = ch_spec["min"]
            if "step" in ch_spec: va.major_unit = ch_spec["step"]
            va.has_major_gridlines = ctype.startswith("radar")
            if va.has_major_gridlines:
                va.major_gridlines.format.line.color.rgb = rgb(tn["faint"])
            va.format.line.color.rgb = rgb(tn["faint"])
            va.tick_labels.font.size = Pt(9); va.tick_labels.font.color.rgb = rgb(tn["muted"])
            ca = ch.category_axis
            ca.tick_labels.font.size = Pt(11 if ctype.startswith("radar") else 10.5)
            ca.tick_labels.font.color.rgb = rgb(tn["text"])
            ca.format.line.color.rgb = rgb(tn["faint"])
        except Exception:
            pass
        if ch_spec.get("title"):
            ch.has_title = True
            ch.chart_title.text_frame.text = ch_spec["title"]
            r = ch.chart_title.text_frame.paragraphs[0].runs[0]
            r.font.size = Pt(12); r.font.color.rgb = rgb(tn["text"])
        return gf

    def L_chart(self, slide, s, tn):
        """Native editable chart (bar/line/radar) + optional photos or bullets.
        Use radar for facial-assessment scales (pre vs post), line for scores over visits."""
        self.header(slide, s, tn)
        top = TOP if s.get("subtitle") else 1.55
        imgs = s.get("images", [])
        bullets = s.get("bullets")
        if imgs:
            iw = CW * s.get("image_ratio", 0.5)
            cells = self._grid(len(imgs), len(imgs), ML, top, iw, self.floor(s) - top, 0.18, 0.05, s.get("aspect", 3 / 4))
            for im, (x, y, w, h) in zip(imgs, cells):
                im = {"path": im} if isinstance(im, str) else dict(im)
                im.setdefault("credit", True)
                self.picture(slide, im, x, y, w, h, tn=tn)
            cx, cw = ML + iw + 0.3, CW - iw - 0.3
        elif bullets:
            cx, cw = ML, CW * 0.6
        else:
            cx, cw = ML + 0.5, CW - 1.0
        ch_h = self.floor(s) - top - (0.5 if s.get("scale_legend") else 0)
        self._chart(slide, tn, s.get("chart", {}), cx, top, cw, ch_h)
        if s.get("scale_legend"):
            self.text(slide, cx, self.floor(s) - 0.4, cw, 0.35, s["scale_legend"], 10, tn["muted"], align="center", tn=tn)
        if bullets and not imgs:
            self.dots(slide, tn, ML + CW * 0.6 + 0.4, top + 0.3, CW * 0.4 - 0.4, bullets, s.get("size", 15))
        if s.get("note"):
            self.note(slide, tn, s["note"], 6.35)

    def L_plan(self, slide, s, tn):
        """Treatment plan: annotated photos (injection map) + products with site doses.
        products: [{name, image?, technique?, items: [{site, dose}] or ["text"]}]"""
        self.header(slide, s, tn)
        top = TOP if s.get("subtitle") else 1.5
        imgs = s.get("images", [])
        iw = CW * s.get("image_ratio", 0.55)
        cells = self._grid(len(imgs), len(imgs), ML, top, iw, self.floor(s) - top, 0.18, 0.05, s.get("aspect", 3 / 4))
        for im, (x, y, w, h) in zip(imgs, cells):
            im = {"path": im} if isinstance(im, str) else dict(im)
            im.setdefault("credit", True)
            self.picture(slide, im, x, y, w, h, tn=tn)
        px, pw = ML + iw + 0.5, CW - iw - 0.5
        y = top + 0.1
        prods = s.get("products", [])
        size = s.get("size", 13 if len(prods) >= 3 else 14)
        for pr in prods:
            name_x = px
            if pr.get("image") and self.path(pr["image"]):
                p, ww, hh = self.prep_image(pr["image"], 1.3, 0.55, "contain")
                slide.shapes.add_picture(str(p), Inches(px), Inches(y), Inches(ww), Inches(hh))
                name_x = px + 1.45
            self.text(slide, name_x, y + 0.05, pw - (name_x - px), 0.35, pr.get("name", ""), 15,
                      pr.get("color", self.t["accent"]), bold=True, tn=tn)
            y += 0.58 if pr.get("image") else 0.45
            if pr.get("technique"):
                th = est_height(pr["technique"], pw, 11)
                self.text(slide, px, y - 0.05, pw, th, pr["technique"], 11, tn["muted"], tn=tn)
                y += th + 0.05
            for it in pr.get("items", []):
                if isinstance(it, dict):
                    mk = it.get("marker", pr.get("marker_color"))
                    d = 0.1
                    if mk:
                        self.rect(slide, px, y + 0.08, d, d, mk, shape=MSO_SHAPE.OVAL)
                    self.text(slide, px + (0.22 if mk else 0), y, pw * 0.42, 0.3, it.get("site", ""), size,
                              tn["text"], tn=tn)
                    self.text(slide, px + pw * 0.44, y, pw * 0.56, 0.3, it.get("dose", ""), size, tn["text"],
                              tn=tn, font=self.t["font_en"] if self.t["font_en"] != self.t["font_cn"] else None)
                    y += size / 72 * 1.5
                else:
                    h = est_height(it, pw - 0.3, size)
                    self.dots(slide, tn, px, y, pw, [it], size, gap=0.05)
                    y += h + 0.05
            y += 0.28
        if y > BOTTOM:
            self.warnings.append(f"slide {self.page}: plan column overflows ({y:.2f}in)")
        if s.get("note"):
            self.note(slide, tn, s["note"], 6.35)

    def L_full_image(self, slide, s, tn):
        img = s.get("image") or {}
        img = {"path": img} if isinstance(img, str) else img
        top = 1.45 if s.get("title") else 0.5
        self.header(slide, s, tn)
        self.picture(slide, img, ML, top, CW, BOTTOM - top + 0.05, fit=img.get("fit", "contain"), tn=tn)
        if s.get("caption"):
            self.text(slide, ML, 6.62, CW, 0.25, s["caption"], 9.5, tn["muted"], tn=tn)

    def L_video(self, slide, s, tn):
        self.header(slide, s, tn)
        top = TOP if s.get("subtitle") else 1.55
        vw = CW * 0.6
        vh = vw * 9 / 16
        vp = self.path(s.get("video"))
        poster = self.path(s.get("poster")) if s.get("poster") else None
        if vp:
            kw = {"poster_frame_image": str(poster)} if poster else {}
            slide.shapes.add_movie(str(vp), Inches(ML), Inches(top), Inches(vw), Inches(vh),
                                   mime_type=s.get("mime", "video/mp4"), **kw)
        else:
            self.picture(slide, s.get("poster") or "", ML, top, vw, vh, tn=tn)
        if s.get("caption"):
            self.text(slide, ML, top + vh + 0.1, vw, 0.3, s["caption"], 9.5, tn["muted"], tn=tn)
        self.dots(slide, tn, ML + vw + 0.45, top + 0.3, CW - vw - 0.45, s.get("bullets", []), s.get("size", 14))

    def L_questions(self, slide, s, tn):
        self.header(slide, s, tn)
        self.dots(slide, tn, ML + 0.4, 1.9, CW - 0.8, s.get("bullets", s.get("questions", [])),
                  s.get("size", 19), gap=0.42)

    def L_references(self, slide, s, tn):
        self.header(slide, s, tn)
        refs = s.get("refs", [])
        size = s.get("size", 9.5)
        cols = s.get("cols", 2 if len(refs) > 8 else 1)
        cw = (CW - 0.4 * (cols - 1)) / cols
        per = math.ceil(len(refs) / cols)
        for c in range(cols):
            y = 1.55
            for i in range(c * per, min(len(refs), (c + 1) * per)):
                txt = f"{i + 1}. {refs[i]}"
                h = est_height(txt, cw, size, 1.25)
                self.text(slide, ML + c * (cw + 0.4), y, cw, h, txt, size, tn["muted"], tn=tn, spacing=1.1)
                y += h + 0.06
            if y > BOTTOM + 0.2:
                self.warnings.append(f"slide {self.page}: references overflow; split into 2 slides")

    def L_thanks(self, slide, s, tn):
        self.text(slide, 0, 2.7, SW, 1.2, s.get("title", "谢谢"), 54, tn["text"], bold=True, align="center",
                  anchor="middle", name="!!hero", tn=tn)
        if s.get("subtitle"):
            self.text(slide, 0, 3.95, SW, 0.4, s["subtitle"], 15, tn["muted"], align="center", tn=tn)
        if s.get("contact"):
            self.text(slide, 0, 5.8, SW, 0.3, s["contact"], 11, tn["muted"], align="center", tn=tn)


def load_spec(p):
    txt = Path(p).read_text(encoding="utf-8")
    if p.endswith((".yaml", ".yml")):
        import yaml
        return yaml.safe_load(txt)
    return json.loads(txt)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec", nargs="?")
    ap.add_argument("--out")
    ap.add_argument("--list-layouts", action="store_true")
    a = ap.parse_args()
    if a.list_layouts:
        for n in sorted(k[2:] for k in dir(Deck) if k.startswith("L_")):
            doc = (getattr(Deck, "L_" + n).__doc__ or "").strip().splitlines()
            print(f"{n:12s} {doc[0] if doc else ''}")
        return
    if not (a.spec and a.out):
        ap.error("spec and --out are required")
    spec = load_spec(a.spec)
    deck = Deck(spec, Path(a.spec).resolve().parent, a.out)
    warns = deck.build()
    try:
        import subprocess
        fams = subprocess.run(["fc-list", ":", "family"], capture_output=True, text=True).stdout
        for f in {deck.t["font_cn"], deck.t["font_en"]}:
            if f not in fams:
                warns.append(f"font '{f}' not installed here – PowerPoint/LibreOffice will substitute "
                             f"(install it, or set theme_overrides.font_cn/font_en)")
    except Exception:
        pass
    size = Path(a.out).stat().st_size / 1e6
    print(f"✓ {a.out}  ({deck.page} slides, {size:.1f} MB)")
    for w in warns:
        print("⚠", w)


if __name__ == "__main__":
    main()
