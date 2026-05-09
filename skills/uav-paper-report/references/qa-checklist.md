# QA Checklist

Use this checklist before delivering a deck.

## Build And Export

- PPTX generation finishes without exceptions.
- PDF export with LibreOffice succeeds.
- PNG render creates one image per slide and a preview grid.
- Final outputs include PPTX, PDF, preview grid, and slide PNGs.
- Public or uploaded PPTX assets have been sanitized with `scripts/sanitize_pptx_privacy.py`.
- Privacy scan reports no personal names, editor metadata, notes slides, comments, custom properties, or visible presenter placeholders containing real names.
- Multi-template public examples pass `scripts/run_template_matrix.py`; local template generalization cases pass `scripts/run_template_smoke.py` when testing new template families.

## Visual Review

Check the preview grid first, then individual risk slides.

- Cover and thanks slide use consistent accepted blue.
- When adapting a template, visible master placeholders such as `单击此处...`, `Click to...`, dates, footers, and slide-number placeholders have been removed or intentionally covered before adding content.
- Rendered body slides do not expose template sample text such as lab names, `author`, `title`, dates, footer bars, or page-number widgets unless the user explicitly wants that template element.
- Slide regions are computed from the actual deck size; no layout assumes a fixed 13.33 x 7.5 canvas.
- No text overlaps other text, images, tables, formulas, or rules.
- No text visually touches a rule, metric row, image, table, or formula after LibreOffice PDF export.
- No image/text collision.
- No page has obvious meaningless blank lines.
- No body bullet, method explanation, table cell, or figure interpretation uses manual line breaks to fake wrapping.
- No continuous bullet block has a gap larger than about one normal text line between paragraphs.
- No slide uses empty paragraphs or standalone bullet markers to create vertical spacing.
- Content slides do not leave more than about 20% empty area unless intentionally sparse.
- No content slide has a large empty middle band between top content and bottom metrics/summary.
- No content slide has a wide internal horizontal whitespace band that makes the page feel half-empty.
- Figures are not stretched; tables are not squeezed.
- Figure boxes match the actual crop aspect ratio; no image appears tiny inside an oversized region.
- Figure crops contain only the intended figure content; they do not include paper body text, full captions, page margins, or unrelated neighboring figures.
- Architecture or pipeline crops that are truncated or caption-heavy have been redrawn as native PPT diagrams instead of inserted as compromised images.
- Images are not visually tiny inside oversized boxes. The displayed image should occupy the intended evidence region and match its source aspect ratio.
- Image boxes are chosen from the actual crop aspect ratio, not copied from unrelated slides.
- Image-heavy slides have enough adjacent interpretation text; if the interpretation cannot fit, the evidence must be split across slides.
- Multi-image slides have separate, balanced image boxes; no generic image helper has pushed an image into text or left a large blank band.
- Every shown figure has a nearby interpretation bullet explaining what it proves.
- Figure captions are omitted unless there is a stable caption lane; interpretation should usually be in nearby body bullets.
- Formula rows are aligned and evenly spaced.
- Multi-line formulas have intentional continuation alignment and consistent font size.
- Method section font size is consistent with other sections.
- Paragraphs have bullets or clear markers.
- First sentences in body areas are also formatted as bullets or labeled claims; no loose unformatted opening sentence.
- Red/bold emphasis is sparse and meaningful.
- Red emphasis is applied only to the decisive word, phrase, or number, not to a full explanatory paragraph.
- Three-column layouts have balanced column heights; no single column wraps into the next horizontal region.
- Slides do not mix vertical stacking and side-by-side reading paths without a visible separator.
- Text regions were given rendered slack; no final text line touches a rule, image, table, formula, or bottom margin.

## Risk Slide Review

Open individual PNGs, not only the preview grid, for:

- three-column layouts, metric rows, or middle-band filler layouts;
- method summary slides with more than one diagram;
- architecture, FSM, pipeline, and hardware-result slides;
- slides with square/tall figures placed beside text;
- any slide where the crop was changed;
- any slide where a shape-overlap checker previously failed.
- any slide changed in this pass to fix blank area, text overlap, image overlap, or rule collision.
- any slide flagged by `scan_rendered_slides.py` for high blank fraction, top/bottom blank, internal whitespace, or an empty body quadrant.

For each risk slide, verify the rendered PNG directly:

- top bullets, figures, rules, and bottom bullets are in separate regions;
- the figure is not clipped and has no unintended paper text;
- text does not wrap into meaningless one-word or empty lines;
- empty area is intentional and not caused by a bad crop or mismatched figure box.
- dense text boxes have visible slack below the final rendered line.
- metric rows sit in a clean band and do not collide with preceding text or slide bottom.
- there are no accidental paragraph gaps inside a bullet block.

## Text Review

Search for and remove:

- `报告口径`
- `讲解重点`
- `应该强调`
- `这页`
- `这一页`
- `这些图说明`
- `可以看到`
- `问题是`
- `代价是`
- `下一步应该`
- `如果结合`

Replace them with direct conclusions or evidence interpretation.

## Privacy Review

Before committing or uploading a skill package:

- Run `python scripts/sanitize_pptx_privacy.py --in-place assets/examples`.
- Run `python scripts/sanitize_pptx_privacy.py --check-only --fail-on-warning assets/examples`.
- Pass site-specific `--replace "旧姓名=报告人"` and `--forbid "旧姓名"` values when private inputs contain real names or machine/user identifiers. Do not hard-code those values into the public skill.
- Inspect cover and final pages of public example PPTX files. Real presenter names must be replaced with generic placeholders such as `报告人` or `Presenter`.
- Do not publish PPTX files that still contain notes slides, comment authors, custom properties, document creators, company names, or machine/user identifiers.

## Formula Review

- Formula text is editable where feasible.
- Formula font is Times New Roman.
- Formula sizes are visually consistent across method slides.
- Formula text does not show code-style braces/underscores such as `p_{k+1}` or `L_clearance` in the rendered PNG.
- Long formulas are split across slides or shortened with explanation.
- Explanatory bullets say what each term means and why it matters.

## Output Response

Final response should be concise and include links to:

- PPTX
- PDF
- preview grid

Mention any unresolved limitation, especially if a formula had to be rasterized or if a source figure is low resolution.
