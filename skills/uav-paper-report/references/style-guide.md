# Style Guide

## Style Match

- Keep the accepted header rhythm and gray header band from the generated examples.
- Cover and final thanks slide must use the same blue band color as the accepted examples (`RGB(0,121,192)` in the accepted scripts).
- Do not add section divider slides by default.
- Use full-width rules to separate body regions. Avoid nested cards or decorative boxes.
- If an example deck or generated output will be committed to a public repository, first remove personal visible names, document authors, editor metadata, notes, comments, and custom properties with `scripts/sanitize_pptx_privacy.py`.

## Typography

- Force Times New Roman on every run, including East Asian and complex script settings.
- Normal hierarchy:
  - Section number/title: accepted style, compact.
  - Main bullet: 19-20 pt, bold only for first-level claim.
  - Secondary bullet: 17.5-19 pt.
  - Dense notes: no smaller than 16.8 pt unless unavoidable.
  - Tables: 12.5-14 pt.
  - Formulas: 19-24 pt.
- Avoid large size jumps between pages in the same section.
- Do not leave raw empty paragraphs or blank bullets.
- Do not use manual line breaks inside body bullets, method descriptions, table cells, or figure explanations. A line break in these regions usually means the text box or wording is wrong. Split into another bullet or redesign the region instead.
- Do not use empty paragraphs as spacers. Vertical rhythm must come from box placement, fixed bands, and small paragraph spacing.

## Bullets And Indents

Every body paragraph should have a visual marker. Use:

- Level 0: `●`, hanging indent about 210000/-140000 EMU.
- Level 1: `•`, hanging indent about 350000/-120000 EMU.
- Level 2: `–`, hanging indent about 560000/-105000 EMU.

Do not make the first sentence of a slide a loose unbulleted sentence.

## Layout Density

- Target less than 20% unused visual area on most content slides.
- Split method/formula slides when content crowds or formula spacing becomes uneven.
- Do not mix top-to-bottom and left-to-right layouts randomly on the same slide.
- Use one main read per slide: claim + evidence object + interpretation.
- If there is a figure, include interpretation; if there is a table, highlight the key result.
- Choose a layout family per slide:
  - Top-to-bottom: title, bullets, figure/table, rule, interpretation.
  - Left-right: bullets/interpretation on one side and figure/table on the other.
  - Two-figure comparison: two balanced figure boxes with shared top claim and bottom interpretation.
- Avoid mixing layout families on the same slide unless the regions are clearly separated.
- For content slides, empty area above or below a small figure is a defect. Re-crop the figure, move to a left-right layout, or add technical interpretation.
- Do not leave a large empty middle band between top bullets and bottom metrics. Add a middle logic/comparison region or move to a left-right layout.
- Three-column layouts need conservative text lengths. Keep each column to 2-3 short bullets, allow wider columns for English-heavy terms, and inspect the rendered PNG because export wrapping is often taller than PPTX bounds.
- Rules, metric rows, figure boxes, and bottom interpretation bands should have visible rendered gaps from text. A line that visually touches a rule or crosses into the next region is a layout failure even if the shape-overlap checker passes.
- If a slide is repaired for blank area, rerender that slide and inspect it full-size; the repair often creates new wrapping or rule-collision defects.
- If a content slide has a large empty band, change the layout family. Typical fixes are a left-right evidence layout, a larger natural-ratio figure, a native metric/table band, or splitting dense method content across two slides. Do not add an unformatted sentence just to occupy space.
- If one body quadrant is empty while another quadrant is dense, rebalance the slide with a small native table, a term explanation band, a larger natural-ratio visual, or a split slide. Do not leave a lower-right or lower-left region blank just because the top text fits.
- Body text needs a wrapping budget. Keep English-heavy bullets shorter, use nonbreaking spaces for units and method names, and reserve extra text-box height so LibreOffice export does not create orphan words or apparent blank lines.
- If two paragraphs look separated by a blank line after rendering, reduce `space_after`, increase the text-box height, or split the content into two visually separated regions with a rule/label. Raw blank gaps inside a continuous bullet block are not acceptable.
- Wide empty areas are layout failures, not copywriting problems. Prefer a larger crop, a native diagram, a compact table, or a left-right evidence layout over adding a loose filler sentence.
- Define fixed non-overlapping regions before placing elements. Use one of these slide plans and keep it through the slide:
  - top-to-bottom: title/claim -> evidence -> interpretation;
  - left-right: text/interpretation -> figure/table/formula;
  - evidence-first: large figure/table -> compact nearby interpretation.
- A slide with both text and an image must start from the image crop ratio. After the image region is fixed, write text to the remaining space; do not shrink or stretch the image to fit already-written text.
- Leave a bottom safety margin for rendered text. If the planned content leaves no slack below the final line, split the slide or shorten the wording before rendering.
- Do not place a small figure in a large box just to fill the slide. Either crop/enlarge it at its natural ratio, redraw it as PPT shapes, or replace the slide with a table/diagram explanation.

## Figures

- Preserve aspect ratio. Crop with intent; do not stretch.
- Keep figures away from text boxes. Add at least 0.12-0.18 in visual gap where possible.
- Use figure text in the body, not fragile captions below images.
- Avoid placing tiny original paper figures without explanation.
- Before inserting a crop, inspect it directly. It should contain the actual figure and only necessary labels; it should not include paper body text, full captions, page margins, or neighboring unrelated figures.
- Match the figure box to the source aspect ratio:
  - wide grids/results: use a wide, low-height box;
  - square/near-square process images: use a balanced box;
  - tall pipeline screenshots: use a narrow/tall box or convert to text explanation.
- Do not use one generic image box for different figure shapes. Preserved aspect ratio inside a poorly chosen box still creates bad proportions, tiny figures, or large blank bands.
- The image box should be selected from the crop's real aspect ratio before text is placed. If the crop occupies only a small strip inside the box, change the slide layout or crop again.
- If the paper caption is needed for meaning, paraphrase it in bullets instead of placing a separate caption under the figure.
- For hardware/result grids, keep the image large enough to inspect the sequence and place interpretation below a rule. Do not crop off subfigure labels unless the sequence remains clear.
- For architecture/FSM/method diagrams, use dedicated slide functions or fixed regions so the diagram never crosses into top bullets or bottom interpretation.
- If the paper's architecture/FSM/method figure is cropped, truncated, or caption-heavy, redraw the important structure with native PPT shapes. A clean editable diagram is better than a compromised paper crop.
- Match crop and layout together. If a preserved-aspect image appears tiny inside a large box, the box is wrong for that figure; resize/reposition the region or choose a different crop.
- Decide the image box from the real crop aspect ratio before writing surrounding text. Wide result strips, tall hardware composites, and square diagrams need different slide functions.
- Image interpretation belongs near the image in bullets. Avoid fragile standalone captions below figures unless the layout leaves a stable caption lane.
- If a figure cannot be made large enough for inspection and still leave room for interpretation, split it into a figure slide and a result-interpretation slide.
- Never allow a figure to share vertical or horizontal space with text. If the image and text need the same region, the slide is overloaded and must be split or redesigned.
- Prefer no separate caption over a badly placed caption. Put the figure meaning in nearby bullets.

## Formulas

- Prefer editable PowerPoint text/shape formulas; avoid raster formula images unless the original formula is too complex and must be quoted visually.
- Use Times New Roman formula runs and stable row positions.
- Keep labels `(1)`, `(3)`, etc. aligned and gray; formula body black.
- Do not place formulas too close together; do not leave a formula page half empty.
- Split long method derivations into: representation, objective, constraints, post-processing.
- Formula slides should not mix several layout families. Use a clear formula band plus term interpretation, or split into two slides if formula rows and explanation bullets compete for space.
- Use fixed row positions for formula bands. Keep equation labels aligned, equation body in one consistent size, and explanatory notes in a visibly separate but nearby bullet region.
- If a formula spans more than one line, make the continuation line intentionally aligned with the formula body. Do not let it look like an accidental body-text wrap.
- Prefer editable run-based formulas with smaller superscript/subscript text using DrawingML baseline. Avoid rendering code notation such as `x_{k+1}`, `L_clearance`, or raw braces in the final PNG.
- Formula interpretation should be phrased as technical meaning, not only variable names. Use labels such as `目标速度`, `速度边界`, `航向对齐`, `安全裕度`, and `控制连续`.
- Formula rows should sit in a dedicated band with consistent row height. If row spacing alternates between cramped and loose, split the derivation into two method slides.
- Formula slides should use fewer wider rows, aligned equation numbers, and a nearby term table or bullet interpretation. Do not mix formula rows, normal paragraphs, and tables in the same unseparated region.
- If a formula block takes more than about 55% of the body height, split the derivation across two method slides instead of reducing body font or tightening row spacing.

## Tables And Metrics

- Tables should be native PPT tables or editable shape/table systems.
- Header row light gray, body white, compact margins.
- Red only for the winning method/name/value, decisive metric, or one key term inside a paragraph. Do not turn an entire paragraph red because it contains a key term.
- Use metric rows for 2-3 key numbers, not for decorative filler.

## Wording

Use direct slide claims:

- Good: `消融实验显示，去除 ESDF 后总耗时从 5.55 ms 降至 0.37 ms。`
- Good: `该方法将连续轨迹约束转化为控制点上的可导代价。`
- Bad: `这页主要讲解消融实验。`
- Bad: `这里应该强调 EGO 的优势。`
