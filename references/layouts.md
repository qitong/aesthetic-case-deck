# deck_spec reference

```json
{
  "theme": "gallery",                       // gallery | nocturne | clinical
  "theme_overrides": {                      // optional, usually from design.md front-matter
    "accent": "C9B58E", "font_cn": "Source Han Serif CN", "font_en": "Constantia",
    "dark": {"bg": ["2A2A2D", "060607"]}, "light": {"text": "1C1C1E"}, "default_tone": "dark"
  },
  "meta": {
    "footer": "仅供医疗专业人士学术交流",
    "photo_credit": "Dr XXX",         // shown on photos with credit:true (default for patient layouts)
    "before_label": "治疗前", "after_label": "治疗后",
    "logo": "assets/logo.png", "logo_width": 1.1, "logo_pos": "top-right",
    "page_numbers": true
  },
  "slides": [ { "layout": "...", ... } ]
}
```
YAML (`.yaml`) works too. Paths are relative to the spec file. Canvas 13.333 × 7.5 in.

## Fields every slide accepts

| Field | Meaning |
|---|---|
| `layout` | one of the layouts below |
| `title`, `subtitle` | conclusion-style title; muted scope/source line |
| `tone` | `dark` / `light` (overrides theme default) |
| `cite` | string or list → citation line bottom-left |
| `note` | summary/caveat line under the content |
| `notes` | speaker notes |
| `footer` | override meta.footer for this slide (`""` to hide) |
| `bg_image`, `bg_dim` | full-bleed background photo with dark overlay (0–1) |
| `size` | force body font size (otherwise auto-fit) |
| `logo` | `false` to hide the logo here |
| `skip` | `true` to leave the slide out |

## Image objects

Anywhere an image is accepted: a path string, or
```json
{"path": "photos/a.jpg", "fit": "cover|contain", "focus": [0.5, 0.4],
 "eyebar": [0.36, 0.46], "credit": true, "label": "2024-03",
 "annotations": [ ... ]}
```
- `focus` = which point to keep when cropping (x, y as 0–1). Faces: y≈0.4.
- `eyebar` = [y0, y1] or [y0, y1, x0, x1] relative; burned into the pixels.
- `fit: contain` for paper figures/diagrams (no cropping); `cover` for photos.

### Annotations (native shapes, positions relative to the placed image)

| type | fields | use |
|---|---|---|
| `dot` | x, y, size(in, 0.11), color, alpha | bolus / injection point |
| `zone` | x, y, w, h (rel), rotate, color, alpha | filled translucent injection area |
| `ellipse` | x, y, w, h, rotate, color, dashed(true) | concern area outline |
| `arrow` | x, y, x2, y2, color (alert red), width | descent vector / cannula path |
| `line` | x, y, x2, y2, dashed | ligament line, reference line |
| any + `label` | lx, ly, label_size | text label next to the mark |

Colour-code annotations per product and repeat that colour as `marker_color` in the `plan`
product list.

## Layouts

| layout | key fields | notes |
|---|---|---|
| `cover` | title, title_en, subtitle, kicker, presenter, affiliation, date, bg_image | |
| `chapter` | no ("01"), title, title_en | |
| `statement` | text, sub, size, bold | big centred sentence |
| `bullets` | bullets[], note | items may be `{text, sub}`; `**bold**` inline |
| `image_text` | image, caption, bullets[], heading, image_side (left/right), image_ratio (0.56) | figure + dots |
| `full_image` | image, caption | one big figure / anatomy render |
| `gallery` | images[{path,label,fit}], cols, aspect (w/h), label_align | 2–8 images |
| `compare` | pairs[{view, before, after, before_label, after_label}], labels[], bullets[], aspect (0.75) | before/after; 1 pair + bullets, or 2–3 pairs |
| `journey` | steps[{image, label, note, highlight}], aspect | photo strip on a timeline |
| `timeline` | events[{date, title, text (str/list), highlight}] | sessions without photos |
| `patient` | images[] (with annotations), facts[[k, v]], image_ratio, aspect | profile |
| `plan` | images[] (with annotations), products[{name, image, technique, marker_color, color, items[{site, dose, marker}] or [str]}], image_ratio | treatment plan |
| `chart` | chart{type: bar/hbar/line/radar/radar_filled, categories[], series[{name, values}], max, min, step, colors[], data_labels, title}, images[], bullets[], scale_legend | native editable chart |
| `table` | columns[], rows[[]], col_widths[], first_col_bold, highlight_col | auto font size |
| `kpi` | intro (str/list), stats[{value, label, accent}], value_size | big numbers |
| `cards` | items[{num, title, body (str/list)}], cols, numbered (true), fill | conclusions, claims |
| `two_col` | left{heading, bullets}, right{heading, bullets}, fill | pros/cons, case vs paper |
| `letter_rows` | rows[{letter, name, text, bold}] | PICO, A.R.T., etc. |
| `video` | video (mp4), poster, caption, bullets[] | embedded movie |
| `questions` | bullets[] | discussion |
| `references` | refs[], cols | auto 2 columns when > 8 |
| `thanks` | title, subtitle, contact | |

## Capacity guide (to avoid overflow)

- `bullets`: ≤ 5 items, ≤ 60 CJK chars each.
- `image_text`: ≤ 4 bullets, ≤ 50 chars each.
- `table`: ≤ 7 rows × 4 cols; cells ≤ 40 chars. Longer → split into two slides.
- `cards`: 3 cards (max 4 per row), body ≤ 70 chars.
- `plan`: ≤ 3 products × 3 sites.
- `journey`: 3–6 steps; notes ≤ 12 chars.
- `references`: ≤ 16 per slide.

The builder auto-shrinks text within sensible limits and prints `⚠` warnings when it still
overflows — fix those by shortening or splitting, not by forcing tiny `size`.
