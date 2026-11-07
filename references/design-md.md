# Using a user-supplied design.md

A design.md tells you how *this* deck should look and flow. It overrides the default
template where it speaks, and the default fills every gap. Template: `assets/design-template.md`.

## Reading it

1. **YAML front-matter (optional)** → machine tokens. Map directly:

| design.md key | spec |
|---|---|
| `theme` | `theme` (gallery / nocturne / clinical) — base to start from |
| `colors.background_dark: ["#xxxxxx", "#yyyyyy"]` | `theme_overrides.dark.bg` |
| `colors.background_light` | `theme_overrides.light.bg` |
| `colors.text_dark_bg` / `colors.text_light_bg` | `theme_overrides.dark.text` / `.light.text` |
| `colors.muted` | `theme_overrides.dark.muted` (and light, if given separately) |
| `colors.accent`, `colors.accent2`, `colors.alert` | `theme_overrides.accent` / `accent2` / `alert` |
| `default_tone` | `theme_overrides.default_tone` |
| `fonts.cn`, `fonts.cn_bold`, `fonts.en` | `theme_overrides.font_cn` / `font_cn_bold` / `font_en` |
| `footer`, `photo_credit`, `logo`, `logo_pos`, `page_numbers`, `before_label`, `after_label` | `meta.*` |

Strip `#` from hex colours.

2. **Markdown body** → structure and rules, interpreted by you:
   - *Section list / slide list* → slide order. Map each described slide to the closest layout
     (table in `layouts.md`). If they list sections only, fill the slides within each section
     using the default recipes.
   - *Rules* ("每页不超过 3 条", "标题用问句", "所有前后对比用黑底", "不出现品牌名") → apply
     to every slide; re-check during preview.
   - *Reference images* (screenshots of slides they like) → read them and match spacing,
     tone, and composition with the closest layout + overrides.

## Precedence

user's explicit instruction in chat > design.md > default-template.md > layout defaults.
Iron rules in SKILL.md (no invented data, privacy, citations) are never overridden.

## When design.md asks for something the renderer can't do

- Slide size other than 16:9, slide masters, animations, or custom fonts embedded → tell
  the user; offer the nearest option (e.g. animations can be added in PowerPoint afterwards).
- A layout that doesn't exist (e.g. a "split-slider" half-before/half-after face) → use the
  nearest (`compare`) and say so; if they want it repeatedly, add a new `L_<name>` method to
  `build_deck.py` following the existing pattern (header → geometry → primitives → warnings).
- Conflicting rules (e.g. "最多 15 页" but 6 timepoints × 3 views) → ask which wins.

Record how you interpreted the design.md in `work/storyline.md` (a short "design decisions"
list) so the user can correct it.
