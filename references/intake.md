# Intake: questions, inputs, mapping

## 1. Question bank

Ask with AskUserQuestion first (≤4 questions, see SKILL.md Step 0), then collect paths and
facts in plain text. Only ask what you can't read from the materials.

| Topic | Ask | Why it matters |
|---|---|---|
| Inputs | 论文 PDF? 照片文件夹? 病程记录? 视频? 产品包装图/logo? | decides storyline A vs B and which slides exist |
| Template | 默认模板 / design.md / 先看样例 | Step 3 path |
| Scenario | 学术会议 / 厂商活动 / 院内培训 / 患者沟通 | footer, compliance depth, tone |
| Language | 中文 / English / 双语 | titles, term glossing |
| Patient | 化名、年龄、性别、Fitzpatrick 分型（如相关）、主诉、诉求、既往治疗、禁忌筛查 | patient profile slide |
| Course | 每次治疗的日期、产品、部位、层次、剂量（左/右）、针具（锐/钝针、规格）、进针点 | plan + timeline slides |
| Assessment | 用了哪种量表？分值？谁评的（医生/患者/盲评）？哪个时间点？ | radar/bar charts; never invent |
| Outcome | 随访时间点、GAIS、满意度、不良事件（含"无"） | results + safety |
| Photos | 拍摄条件是否一致？哪些可公开？是否需要遮眼？ | pairing & privacy |
| Consent | 患者是否书面同意会议/学术展示？是否允许出现真名？ | rule 3 |
| Presenter | 讲者、单位、日期、照片署名（如 "Dr XXX"）、logo | cover, credit, chrome |
| Message | 这次分享想让听众带走的一句话是什么？ | storyline anchor |

If the user doesn't know the take-home message, propose 2–3 options from the materials
(e.g. "长期分层管理优于单次大剂量" / "先深层支撑再浅层融合") and let them choose.

## 2. Handling each input type

**Photos** — `scan_inputs.py` then *look* at contact sheets. Build a mapping:

```
timepoint   view     phase    file                 note
2024-03     front    before   IMG_0012.jpg
2024-03     oblique  marking  IMG_0015.jpg         doctor's own markings → use as plan image
2024-04     front    after    IMG_0230.jpg         4w
```
- Pair before/after by the **same view**; prefer same distance, lighting, expression.
- If lighting/pose differ noticeably, say so in the speaker notes (and to the user) — never
  "correct" it with colour edits that could be read as manipulation.
- HEIC: needs `pip install pillow-heif` (scan reports unreadable files).
- Video: embed with the `video` layout (poster = a frame or the first photo). Large videos make
  huge PPTX; suggest trimming if > 100 MB.

**病程记录 (notes)** — any format. Normalise into the session table in
`assets/case-notes-template.md`. Flag: missing left/right split, units (ml vs syringes),
ambiguous product names (brand ↔ registered Chinese name), dates written as relative ("1个月后").

**Paper** — `extract_paper.py`. Use `paper.md` for numbers, `figs/` for raster figures,
`pages/` + `--crop` for vector diagrams. Keep the figure numbers from the paper in captions
("原文图 17").

**Excel (scores, instrument data)** — read with the `xlsx` skill; long tables → `chart`
layout series.

**design.md** — see `design-md.md`.

## 3. Decision: storyline A or B

- Paper provided and the case uses (or contrasts with) its technique → **A 文献 + 病例**.
- Paper provided only as background (1–2 facts) → **B**, with 1–2 literature slides and cites.
- No paper → **B 纯病例**.
- Several patients → one case section per patient (profile → plan → before/after → journey),
  then a combined discussion. More than 3 patients: consider a summary table slide.
