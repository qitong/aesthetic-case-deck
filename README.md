# aesthetic-case-deck

A Claude Code skill that turns a medical-aesthetics (医美) patient's course of treatment —
病程记录 / patient journey, treatment records and before/after photos — into an **editable,
conference-grade case-report PPTX**, optionally anchored on a reference paper (文献精读 + 病例).

- **Intake interview**: asks which inputs exist (paper PDF, photo folder, notes, videos), which
  template to use (built-in default or your own `design.md`), audience and language.
- **Default template** distilled from real conference decks: literature deep-read structure
  (chapters, appraisal, PICO, KPI, conclusion cards) + patient-journey case structure (profile →
  assessment radar → injection-map treatment plan → before/after by view → multi-year journey).
- **Native PowerPoint output**: text, injection markers, arrows and charts are all editable objects.
- **Guardrails**: never invents clinical data; citations on literature slides; burned-in eye bars
  for de-identification; compliance footers by audience (China NMPA / advertising rules, EU, US).

## Install

```bash
git clone https://github.com/qitong/aesthetic-case-deck ~/.claude/skills/aesthetic-case-deck
pip install python-pptx Pillow PyMuPDF pyyaml
# preview rendering (optional): LibreOffice + poppler
brew install --cask libreoffice && brew install poppler
```

Then in Claude Code: *“用 aesthetic-case-deck 把这个病例做成 PPT，照片在 ./photos，论文是 ./paper.pdf”*.

## Scripts

| Script | Purpose |
|---|---|
| `scripts/scan_inputs.py` | inventory a folder → `inventory.md` + numbered photo contact sheets |
| `scripts/extract_paper.py` | paper PDF → text, metadata, figures; `--crop` for vector figures |
| `scripts/build_deck.py` | `deck_spec.json/yaml` → PPTX (3 themes, 22 layouts; `--list-layouts`) |
| `scripts/preview.py` | PPTX → PNGs + contact sheet for visual QA |

See `SKILL.md` for the workflow, `references/layouts.md` for the spec format,
`assets/design-template.md` for a fill-in `design.md`, and `assets/example_spec.json` for a
full example (placeholder paths render as `[缺图]` boxes).

Fonts: decks default to 思源宋体 (Source Han Serif CN); override via `theme_overrides`.
