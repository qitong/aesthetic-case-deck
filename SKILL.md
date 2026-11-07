---
name: aesthetic-case-deck
description: >-
  Turn a medical-aesthetics (医美) patient's course of treatment — 病程记录 / patient journey,
  treatment records (products, sites, doses, sessions) and before/after photos — into an
  editable, conference-grade case-report PPTX, optionally anchored on a reference paper
  (文献精读 + 病例). Asks the user what inputs exist (paper PDF, photo folder, notes, videos),
  then uses either the built-in default template (distilled from real conference case
  decks: dark photo gallery, patient profile, assessment radar, injection-map treatment plan,
  before/after by view, multi-year patient journey, literature appraisal) or a user-supplied
  design.md. Use when the user says 病例分享PPT / case report PPT / 案例汇报 / 病例幻灯片 /
  patient journey deck / 把病程和前后对比照做成PPT / 结合这篇文献做病例分享 / 文献+病例精读.
  For split-face (半脸对照) instrument-data trial decks use medical-aesthetics-report; for a
  journal manuscript use aesthetic-case-writer.
---

# 医美病例汇报 PPT（Case-Report Deck）

Input → editable PPTX. Pipeline: **intake interview → inventory inputs → (extract paper) →
storyline → deck_spec.json → build_deck.py → preview & fix**.

Every slide element is native PowerPoint (text boxes, shapes, charts, pictures), so the doctor
can edit wording, move injection markers, and swap photos afterwards.

`<skill>` below = this skill's directory.

## Iron rules (read first)

1. **Never invent clinical data.** Doses, dates, products, scores, ages, follow-up times and
   outcomes come only from the user's materials. Missing → ask, or omit the slide. Never fill a
   radar/GAIS with "plausible" numbers.
2. **Separate sources.** On literature slides, distinguish 原文所述 / 我们的解读 / 原文未报告,
   and put a citation line (`cite`) on every slide that uses a paper's data or figures.
3. **Privacy & consent.** Confirm photo consent before building. Use `eyebar` for any photo that
   must be de-identified (it is burned into the pixels, not an overlay). Strip names to a
   pseudonym (e.g. "患者 A，36 岁") unless the user says the patient consented to real name.
4. **Compliance footer.** Default footer on every content slide (editable in `meta.footer`) — see
   `references/content-rules.md` for China (NMPA / 医疗广告) vs. EU/US wording, off-label notes,
   and sponsor (MLR) review.

## Step 0 — Intake interview (ask before doing anything)

Use **AskUserQuestion** (one call, up to 4 questions). Default wording:

1. **你手上有哪些输入？** (multiSelect) — 参考论文 PDF ｜ 照片文件夹 ｜ 病程记录/治疗记录（文本、Word、Excel）｜ 视频（操作/随访）
2. **用哪套版式？** — 默认模板（Recommended，总结自文献精读 + patient-journey 病例 deck）｜ 我提供 design.md ｜ 先看默认模板样例再决定
3. **汇报场景？** — 学术会议病例分享（默认）｜ 厂商/品牌活动（需 MLR 合规审阅）｜ 院内培训/读书会 ｜ 患者沟通（不推荐展示他人照片）
4. **语言？** — 中文（默认，术语首次出现附英文缩写）｜ English ｜ 中英双语标题

Then ask in plain text for the **paths** of whatever they selected (paper, photo folder,
notes, design.md, logo, product pack shots) plus: patient pseudonym/age/sex, presenter name &
affiliation, photo credit text (e.g. "Dr XXX"), and whether consent covers conference
use. If the user already gave some of this, don't re-ask it.

Full question bank, what to do with each answer, and a 病程记录 template to send the doctor:
`references/intake.md` and `assets/case-notes-template.md`.

## Step 1 — Inventory inputs

```bash
python <skill>/scripts/scan_inputs.py <photo_folder> [<notes files>…] --out work/inputs
```
Read `work/inputs/inventory.md` and **look at** `contact_sheet_*.jpg`. Map each photo to
(timepoint, view: front / oblique / profile / smile / upgaze, phase: before / marking / after).
Filenames and EXIF are only hints. Show the user a short mapping table if anything is
ambiguous (two candidate "after" photos, unknown dates). Prefer photos with matched
lighting/pose/background for every before/after pair; flag mismatches rather than hiding them.

Notes (docx/pdf/xlsx) → read with the `docx` / `pdf` / `xlsx` skills, then normalise the
course of treatment into the table in `assets/case-notes-template.md` (session × date ×
product × site × dose × technique × assessment × AE). Ask about gaps.

## Step 2 — Paper (only if provided)

```bash
python <skill>/scripts/extract_paper.py paper.pdf --out work/paper
# vector figure? crop it from a page render:
python <skill>/scripts/extract_paper.py paper.pdf --out work/paper --crop 10 0.05 0.1 0.95 0.6 --name fig16
```
Read `work/paper/paper.md` (not the PDF) and `figs_index.jpg`. Pull: citation, study design
(PICO), sample, technique steps, products/doses, outcomes with exact numbers, complications
monitored, level of evidence, limitations. For appraisal depth you may also run the
`appraise-paper` skill. Note where the user's case agrees with or departs from the paper —
that comparison is the heart of a 文献+病例 deck.

## Step 3 — Choose the template

- **Default template** → `references/default-template.md`: storyline **A (文献 + 病例)** or
  **B (纯病例)**, slide-by-slide recipes, which layout to use, and how many slides per section.
  Drop sections with no data; never pad.
- **User design.md** → `references/design-md.md`: how to read it (front-matter tokens → theme
  overrides; body → slide order/sections/rules), precedence, and what to do when it asks for
  something the renderer can't do. Give users `assets/design-template.md` as a starting point.

Themes: `gallery` (default; charcoal + silver), `nocturne` (pure-black photo gallery,
Latin serif, lavender accent), `clinical` (white, for print/journals). Any slide can set
`"tone": "light"` for a silver/light break (KPI, appraisal tables).

## Step 4 — Write the storyline, then the spec

1. Write `work/storyline.md` first: one-sentence take-home message, section list, and for each
   slide a **conclusion-style title** + the evidence on it (source file / paper page). Show it
   to the user for a quick OK when the deck is > 15 slides or the scenario is sponsor-facing.
2. Write `work/deck_spec.json` (paths relative to the spec file). Layout catalogue & every
   field: `references/layouts.md`. Copy `assets/example_spec.json` as a skeleton.
   - Injection maps: `annotations` on a photo (`zone`, `dot`, `ellipse`, `arrow`, `line`, with
     relative x/y). Place them from the doctor's marking photo or description; if the doctor
     supplied an already-marked photo, use it as-is instead of redrawing.
   - Assessment scales (e.g. facial-assessment wheel, GAIS, Merz, WSRS) → `chart` layout with
     radar/bar using the real scores; if only an image of the scale exists, place the image.
   - Speaker notes go in `notes` — write them for anything the doctor needs to *say* but
     shouldn't be on the slide.

## Step 5 — Build, preview, fix

```bash
python <skill>/scripts/build_deck.py work/deck_spec.json --out output/<name>.pptx
python <skill>/scripts/preview.py output/<name>.pptx          # → <name>_preview/contact.jpg
```
Read the warnings (overflow, missing files, missing fonts) and **look at** `contact.jpg`, then
single slides at higher dpi (`--dpi 110 --slides 5-9`) where needed. Check: text overflow,
faces cropped at chin/forehead (fix with `focus`), before/after pairs at equal scale, markers
on the right anatomy, no slide with only 1–2 lines of text on a large empty card, citation on
every literature slide. Iterate on the spec — not by hand-editing the PPTX.

Fonts: decks specify 思源宋体 (`Source Han Serif CN`), as in the reference decks. If it's not
installed on the presenting machine, either install it or set
`theme_overrides: {"font_cn": "...", "font_en": "..."}`. LibreOffice previews substitute fonts —
judge layout, not glyph shapes.

## Step 6 — Deliver

Give the user: the PPTX path, slide count, a list of anything omitted for lack of data, every
`[缺图]` placeholder, and assumptions to verify (e.g. "marker positions drawn from your
description of the injection sites — please check slide 11"). Offer a compliance pass
(`references/content-rules.md` checklist) if sponsor-facing.

## Files

| Path | Purpose |
|---|---|
| `scripts/scan_inputs.py` | inventory folder → inventory.md/json + numbered photo contact sheets |
| `scripts/extract_paper.py` | paper PDF → paper.md, meta.json, figs/, pages/, `--crop` vector figures |
| `scripts/build_deck.py` | deck_spec.json/yaml → editable PPTX (3 themes, 22 layouts) |
| `scripts/preview.py` | PPTX → PNGs + contact sheet (LibreOffice) |
| `references/intake.md` | interview question bank, input handling, mapping rules |
| `references/default-template.md` | default storylines A/B, slide recipes, distilled design language |
| `references/layouts.md` | every layout and its spec fields |
| `references/design-md.md` | interpreting a user's design.md |
| `references/content-rules.md` | evidence wording, photo standards, compliance (NMPA/FDA/EMA) |
| `assets/design-template.md` | fill-in design.md for users |
| `assets/case-notes-template.md` | 病程记录 template to send to the doctor |
| `assets/example_spec.json` | full example spec using every layout |

Related skills: `aesthetic-anatomy-illustrator` (anatomy/injection-plane figures for the
background section), `appraise-paper` (critical appraisal), `aesthetic-case-writer`
(manuscript), `medical-aesthetics-report` (split-face trial decks).
