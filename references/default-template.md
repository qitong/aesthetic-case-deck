# Default template

Distilled from two non-public conference decks: a long literature deep-read with representative
cases, and a patient-journey case deck. Only the abstract design rules and section structure are
kept here; no wording, data or images from those decks.

## Design language

| Element | Rule (from references) |
|---|---|
| Canvas | 16:9. Dark charcoal→black gradient for most slides; silver/light gradient for data-heavy breaks (KPI, appraisal tables, decision trees). Patient photos always on black. |
| Type | One serif family (思源宋体 SemiBold for titles, Regular for body). Titles 28 pt, body 14–20 pt, captions 9–11 pt, citation 7.5 pt. Big numbers 40 pt. |
| Titles | **Conclusion-style** (e.g. "先支撑、后融合：泪沟用量可以更少" rather than "技术"), with a muted subtitle giving scope/source ("原文图 12 · 内侧段"). Patient slides may use the patient's pseudonym ("患者 A，48 岁"). |
| Bullets | Small round dots, 3–4 bullets max, each a full sentence that says something specific. |
| Containers | Translucent rounded cards with hairline border; tables are text in a card with hairline row separators, header row muted and small. |
| Chapters | Centered number "01" + big title + English subtitle ("Anatomy"). 6–9 chapters in a long deck. |
| Photos | Matched pairs, labelled 治疗前/治疗后 (or before/after) under each; photo credit "Dr XXX" in the lower right of each patient photo; journey as a horizontal strip with dates. |
| Annotations | White dashed ellipses for problem areas, red arrows for vectors (sagging), filled translucent zones / dots for injection sites colour-coded per product. |
| Chrome | Citation line bottom-left on every literature slide; compliance footer (see content-rules.md) tiny at the bottom; sponsor/brand logo top-right only if requested. |
| Evidence voice | Mark what the paper says vs. what it doesn't report vs. your interpretation; grade claims; flag atypical cases — never overclaim (details in content-rules.md). |

## Storyline A — 文献 + 病例 (paper provided)

Typical 25–45 slides. Section sizes are guides; drop anything without content.

| # | Section | Slides | Layouts |
|---|---|---|---|
| 0 | Cover: technique/paper short title (CN) + full English title | 1 | `cover` |
| 1 | **Hook**: before/after — the user's case (or published cases) | 1 | `compare` / `gallery` |
| 2 | Clinical problem the paper addresses (paper figure + 3 points) | 1–2 | `image_text` |
| 3 | Status quo: consensus/standard practice, complication rates | 1–3 | `table`, `kpi`, `image_text`, `video` |
| 4 | The paper's core claims (claim · basis) | 1 | `table` or `cards` |
| — | Chapter 01 文献与作者 | 1 | `chapter` |
| 5 | Paper info (journal, type, n, level of evidence) and author background | 1–3 | `cards`, `gallery` |
| — | Chapter 02 解剖/机制 Background | 1 | `chapter` |
| 6 | Key terms → anatomy layers → mechanism (one idea per slide, figure-led) | 3–10 | `image_text`, `full_image` |
| 7 | Anatomy summary | 1 | `cards` |
| — | Chapter 03 评估/分型 | 1 | `chapter` |
| 8 | Patient classification / checklist / decision flow | 1–4 | `image_text`, `table`, `full_image` |
| — | Chapter 04 技术 Technique | 1 | `chapter` |
| 9 | Overview → step 1 → step 2 … → product rheology choice | 2–8 | `two_col`, `image_text`, `video` |
| — | Chapter 05 结果与证据评价 | 1 | `chapter` (tone light ok) |
| 10 | Results (exact numbers) · representative case · PICO · outcome limitations · claims vs evidence | 3–6 | `kpi`, `image_text`, `letter_rows`, `cards`, `table` |
| — | Chapter 06 我们的病例 Our Case | 1 | `chapter` |
| 11 | Patient profile | 1 | `patient` |
| 12 | Pre-treatment assessment (scale) | 1 | `chart` (radar) |
| 13 | Treatment plan (injection map + products + L/R doses) — one per session if different | 1–3 | `plan` |
| 14 | Before/after per view (front, oblique, profile; smile if available) | 1–4 | `compare` |
| 15 | Patient journey (multi-timepoint strip) + course table | 1–3 | `journey`, `timeline` |
| 16 | Post-treatment assessment (pre vs post radar / GAIS) | 0–1 | `chart` |
| 17 | Case vs paper: what matched, what differed, why | 1 | `two_col` / `table` |
| — | Chapter 07 讨论与结论 | 1 | `chapter` |
| 18 | Strengths & limitations of the paper | 1 | `two_col` |
| 19 | Clinical takeaways (each with evidence strength) | 1 | `table` / `cards` |
| 20 | Conclusions (3 numbered cards) | 1 | `cards` |
| 21 | Discussion questions | 1 | `questions` |
| 22 | Thanks | 1 | `thanks` |
| 23 | Appendix: abbreviations, figure index, references, image/video sources | 2–4 | `table`, `gallery`, `references` |

## Storyline B — 纯病例 / patient journey (no paper)

Typical 12–22 slides (the reference patient-journey deck is this shape).

| # | Section | Layouts |
|---|---|---|
| 0 | Cover — can be a one-line philosophy statement as the title | `cover` |
| 1 | Concept behind the approach (why this patient, why this sequence, why long-term follow-up) | `statement`, `image_text`, `full_image` |
| 2 | Optional background: relevant anatomy or the assessment framework used | `image_text`, `gallery` |
| 3 | Chapter: the case | `chapter` or `statement` |
| 4 | Patient profile (pseudonym, age, complaint, goals, history, assessment) with problem-area annotations | `patient` |
| 5 | Pre-treatment assessment scale | `chart` |
| 6 | Treatment plan (per session) | `plan` |
| 7 | Post-treatment assessment | `chart` |
| 8 | Before/after: front, oblique, profile (one slide per view, or 2–3 views on one) | `compare` |
| 9 | Patient journey strips per view (front / oblique / profile) | `journey` |
| 10 | Course timeline with all sessions & AEs | `timeline` / `table` |
| 11 | Key learnings (3 cards) | `cards` |
| 12 | Discussion question / provocative closing question | `statement` / `questions` |
| 13 | Thanks | `thanks` |

## Slide recipes

- **Hook**: 2 views × before/after, subtitle = who/what/how long. No bullets.
- **Clinical problem**: paper figure (contain, ≤ 56 % width), caption with original figure
  number, 3 bullets: what it is → why it's missed → what this deck answers.
- **Status quo table**: 3 columns (aspect · current practice · rationale/evidence), 4–6 rows;
  subtitle states how strong the underlying evidence is.
- **Results (KPI)**: 1–2 lines of sample description, 3–4 big numbers, footnote defining the
  outcome and exactly what "0 complications" covers. `tone: light`.
- **Representative case**: paper's case figure + a caveat when it is atypical (e.g. doses
  at the top of the reported range).
- **Patient profile**: oblique + profile photo with dashed ellipses on concern areas and red
  arrows for descent vectors; facts list (主诉 / 诉求 / 既往治疗 / 评估 / 禁忌筛查).
- **Treatment plan**: front + oblique with zones/dots coloured per product; right column =
  product (pack shot optional) → site → "L x ml / R y ml"; footer note on approved labelling.
- **Before/after**: same view, same scale; labels under photos; optional 2–3 bullets on what
  changed (observable only — no unmeasured claims like "collagen increased").
- **Patient journey**: 4–6 timepoints, date under each, one-line note (what was done).
  One strip per view if photos exist.
- **Case vs paper**: two columns 一致 / 不同, each bullet specific (dose, plane, cannula, follow-up).
- **Conclusions**: exactly 3 numbered cards, title = the claim, body = the qualifier.
