---
name: uav-paper-report
description: Use when creating, revising, auditing, or regenerating Chinese academic PPT/PPTX paper-report decks for UAV, robotics, planning, SLAM, control, or autonomous systems papers from a paper/PDF and a user-provided PPT style. Especially use for thesis/group-meeting reports that need style-matched layout, Times New Roman text, dense but readable method slides, formulas, figures, tables, restrained red emphasis, and PDF/PNG visual QA.
---

# UAV Paper Report

## Overview

This skill preserves the accepted quality level for Chinese UAV/robotics paper report PPTs. It combines accepted AI-generated example decks, reusable `python-pptx` scaffolds, and mandatory visual/text QA. The public skill package does not include the original private report template.

When this skill triggers, also use the `presentations` skill/plugin for PPTX creation, rendering, and visual verification.

## Workflow

Use this sequence for every full deck or substantial revision:

1. Inspect the user's private template when supplied; otherwise inspect the accepted example decks before writing slides.
2. Read the paper first: problem, assumptions, technical points, method equations, algorithm flow, experiments, limitations, and connection to UAV research.
3. Build a slide plan before coding. For a full paper, target 18-24 slides unless the user asks otherwise.
4. Prepare clean paper assets before slide layout: crop figures from rendered PDF pages, inspect each crop, and reject crops that include unrelated body text, captions, or truncated subfigures.
5. Before placing content on each slide, choose the layout family, reserve fixed regions, calculate image aspect-ratio fit, and estimate text line count. Do not start from a generic image box or loose text box.
6. Generate editable PPTX with native PowerPoint content: text boxes, shapes, native diagrams, tables, editable formula text, and clean cropped paper figures.
7. Export to PDF, render slide PNGs, inspect the preview grid and individual risk slides, then iterate until the QA gates pass. Run render and scan steps sequentially; do not scan while PNGs are still being written.
8. If the deck, examples, or scaffolds will be published or uploaded, run `scripts/sanitize_pptx_privacy.py` on every PPTX asset and verify no personal metadata, notes, comments, or visible presenter names remain.
9. For multi-template work, register publishable examples in `assets/template-profiles/matrix.json` and run `scripts/run_template_matrix.py`. The current public matrix covers 8 paper/template combinations, including FoV-CBF, PRIMER, and SPOT. New generated decks are blocking failures if PPTX audit, export, render, or rendered-slide scan fails. Legacy reference PDFs may stay non-blocking only when explicitly marked as `legacy_pdf_baseline`, and `--baseline-dir` is only an explicit fallback when LibreOffice export is unavailable.
10. For local template generalization work, run `scripts/run_template_smoke.py` with a local template config. The smoke test generates small stress decks from recent UAV paper content and checks 4:3/16:9 sizing, master placeholder cleanup, wide/tall image regions, tables, formula rows, English-heavy wrapping, restrained emphasis, estimated rendered text height, narrow-body wrapping, and rendered blank-area scanning. Also run it with `assets/template-profiles/requirement-smoke.json` so `balanced`, `brief`, `method-detailed`, and `experiment-detailed` requests are all exercised. Its rendered scan must use strict thresholds equivalent to `--blank-warn 0.74 --min-band-fraction 0.10 --body-blank-warn 0.89`. Do not publish private template PPTX files used by smoke tests.
11. Sync final PPTX, PDF, preview grid, slide PNGs, and any reproducible build/crop scripts to the user's requested output directory.

## Deck Structure

Default full-paper structure:

- Cover, using the accepted blue cover style and consistent final thanks slide.
- 3-4 background/problem/motivation slides.
- 1 technical summary slide with metrics or a method chain.
- 6-8 method slides. Split dense formula content across pages; do not crowd one method slide.
- 4-5 experiment/result slides with tables and figure explanations.
- 2-3 conclusion, limitation, and research启发 slides.
- No decorative section divider slides unless the user-supplied style clearly requires them.

The deck is a live paper report, not a product brochure. Write enough content for oral reporting: each normal slide should have meaningful bullets, figures/tables/formulas where useful, and less than about 20% visually wasted blank space.

## Style Contract

Use the accepted header rhythm, blue cover/ending band, and horizontal rules. Keep a calm academic look. If a private template provides a logo, use it only in the generated local output, not in public assets.

- Font: Times New Roman for all runs, including Chinese text when the user asks for 新罗马.
- Body sizes must be consistent within a deck. Pick a small hierarchy and reuse it: main bullets about 17.8-18.5 pt, secondary bullets about 16.8-17.4 pt, tertiary bullets about 16.2-16.6 pt, tables about 11.8-13.0 pt, metric labels about 11.8-12.2 pt, metric notes about 14.5-15.0 pt, formulas about 18.5-21 pt, and native diagram labels at least about 14 pt. Do not tune body font size slide by slide just to make content fit.
- Template profiles may use different fixed hierarchies. Use `audit_pptx_text.py --profile compact`, `--profile dense-visual`, or `--profile classic-large` instead of loosening thresholds manually. The profile must match the template rhythm, and font drift inside that profile still fails.
- Bullets: every paragraph in body areas needs a marker and proper hanging indent. Use `●` for level 0, `•` for level 1, `–` for level 2.
- Emphasis: bold/red only for key terms, core numbers, algorithm names, or decisive comparison results. Apply red to the specific word/run, not the whole paragraph. Do not make most of a slide red.
- Layout: keep one reading direction per slide. Use either top-to-bottom regions or a clear left-right comparison, not mixed random blocks.
- Stress layouts must avoid mixing a left text block, a right metric/note rail, and a full-width bottom table on one slide. That pattern repeatedly creates meaningless wraps and weak region boundaries; use a full-width top-to-bottom stack or a clean left visual/right interpretation split.
- Images: preserve aspect ratio; crop intentionally; do not let text overlap images. The image box must match the source image's real aspect ratio closely enough that the rendered image does not look tiny, stretched, or clipped. Do not embed pipeline, architecture, or flow labels as tiny raster text; rebuild those labels as native PPT shapes so the audit can enforce font size and fit.
- Text regions: use fewer, wider text boxes instead of many narrow boxes. If an English-heavy bullet wraps while visible horizontal space remains elsewhere on the slide, redesign the region rather than accepting the wrap.
- Long English method names should be translated, shortened, or moved into a term table. Do not let a long English phrase create a one-word or two-line orphan in a body bullet when the Chinese explanation is the real content.
- Figure crops: crop to the actual figure content. Do not include paper body text, full figure captions, page margins, or unrelated neighboring figures. Subfigure labels are acceptable when they are part of the figure.
- Captions: avoid separate figure captions if placement becomes awkward. Put the interpretation in body bullets instead.
- Tables: use compact native tables, consistent font, enough row height, and red emphasis only for key winning values. Do not use equal-width columns for semantic two-column tables such as `符号/含义`, `变量/含义`, `对象/进入方式`, or `对象/约束`; the explanation column must be visibly wider to avoid orphan characters and unnatural wrapping.
- Formulas: prefer editable PowerPoint text/shape formulas. Align formula rows, keep formula font consistent, and split complex derivations across slides.
- Manual line breaks: never insert `\n` inside body bullets, method explanations, table cells, or figure interpretations to force wrapping. Use separate bullet paragraphs or resize the region. Manual breaks are allowed only for cover titles and intentionally split formula rows.
- Generation helpers must reject `\n` in body text by default. If a title or diagram node truly needs multiple lines, pass an explicit `allow_newlines=True` or a list of node lines. Do not silently replace `\n` with spaces.
- Generation helpers must also reject empty body paragraphs, empty bullet blocks, empty table cells, and empty formula runs. Blank vertical spacing must come from coordinates and region sizing, not hidden paragraphs.
- Bold level-0 bullets only when the whole paragraph is a short claim. For normal body paragraphs, keep the paragraph black/regular and emphasize only key terms or numbers. Whole-paragraph bold makes adjacent slides look like different font sizes.

Read `references/style-guide.md` when doing a visual refresh or when layout quality is the main issue.

## Layout Hardening Rules

Treat the LibreOffice-exported PNG as the layout source of truth. `python-pptx` shape bounds are only a first pass; Chinese/English mixed text often wraps taller after export.

- Text boxes must have render slack. Leave at least 0.18-0.25 in below dense text before a rule, image, table, or metric row. If a rendered line touches a rule, it is an overlap defect even when the XML bounds do not overlap.
- Keep paragraph spacing small and deliberate. A gap larger than about one normal text line between two body bullets is a defect unless a rule, label, or visual object separates regions.
- Three-column slides are risk slides. Give each column enough width, keep each column to 2-3 bullets, use 16.8-17.2 pt secondary text, and inspect the PNG at full size. If one column wraps much deeper than the others, widen it, shorten the wording, or split the slide.
- Metric rows need a clean band. Do not put metric rows directly after a dense paragraph block; separate them with a rule and a visible gap, and verify the notes do not collide with slide bottom or other text.
- Avoid "top block + bottom block" layouts that leave an empty middle band. If a slide has a wide blank band through the center, add a middle comparison/logic region or change to a left-right layout.
- For content slides, a central empty band taller than about 0.9 in or a visually empty region over about one fifth of the slide is a blocking defect unless the slide is intentionally a cover/thanks/section page.
- Text/image overlap includes visual touch. A figure, table, formula, or rule should not visually touch text; keep a white gap after rendering, not just in code coordinates.
- If fixing one slide moves content into another region, rerender immediately and inspect that individual PNG before doing broad edits.
- Unexplained line breaks are blocking defects. If a rendered bullet wraps while there is visible horizontal room, shorten/no-wrap the English term, widen the text box, or move to a left-right layout; do not keep a hand-inserted blank line or one-word orphan line.
- Do not create empty paragraph objects for vertical spacing. Use shape coordinates and paragraph `space_after` only; the XML must not contain empty body paragraphs or standalone bullet markers.
- Keep paragraph spacing small. Body paragraphs should not use large `space_before` or `space_after` values to imitate blank lines; the audit script should fail abnormal paragraph spacing. Fix sparse pages by moving/resizing regions, adding a native table/diagram/metric row, or rewriting content into real bullets, not by adding paragraph spacing.
- Do not leave empty text bodies on decorative shapes. If a shape is only a line, band, card background, or image shadow, remove its `p:txBody` after creation or have the audit script fail it.
- Do not accept text-box driven "blank-line spacing". If a slide appears to contain a blank line or an orphan wrapped word after rendering, widen the text region, shorten the sentence, or split the content; never add an empty paragraph.
- Do not accept formula-row driven wrapping. If an editable formula/text formula row is longer than the reserved row width, split it into shorter aligned rows or widen the formula band before rendering.
- No shape may extend beyond the slide bounds. This includes process boxes, cards, hidden shadows, tables, pictures, and decorative rectangles; run the PPTX audit before export.
- Method slides must not alternate between vertical stacking and side-by-side blocks without a visible separator. Pick one layout family per slide; if formulas plus interpretation no longer fit cleanly, split into two method slides.
- Use a region budget before coding a slide: header/title area, top claim area, evidence area, interpretation area, and bottom safety margin. A content region cannot be used by both text and image/table/formula even partially.
- Use rendered text height, not only shape coordinates, as the acceptance standard. When a text box follows an image/table/formula/rule, reserve at least one rendered line of slack or split the slide.
- PPTX audits must estimate rendered text height for plain body and metric-note text boxes and include that expanded text footprint in overlap checks. A text box that is technically within its XML bounds still fails if the estimated rendered footprint collides with a rule, table, figure, formula, or bottom margin.
- PPTX audits must warn on narrow body text boxes that create avoidable 4-line wrapping or long-token orphan lines. Fix the source layout by widening, shortening, or splitting the region; do not suppress the warning.
- If a slide contains both a figure and bullets, decide the figure box from the crop aspect ratio first, then write bullets to the remaining width. Do not write text first and squeeze the figure afterward.
- Figure-first slides must allocate the image region from the crop aspect ratio before placing text. If the displayed image occupies less than about 75% of its intended evidence region, the region is mismatched and must be resized or redesigned.
- On image slides, the image should either be large enough to inspect or be omitted. A small figure with a large surrounding blank band is a failed slide even if no overlap is present.
- Blank-area repairs should change the layout, not merely add loose text. Use one of: enlarge a clean figure to its natural aspect ratio, add a native table/metric row, add a compact method diagram, move section regions closer together, or split/recombine slides so each page has a clear read.
- If `scan_rendered_slides.py` warns about high blank fraction, high body blank fraction, large top/bottom blank, large internal horizontal whitespace, or an empty body quadrant, the slide must be rerendered after a layout change. Do not silence the warning by loosening thresholds unless the slide is a cover/thanks page.
- If a paper architecture/pipeline figure is truncated, caption-heavy, visually weak after cropping, or contains small embedded labels, redraw it as editable PPT shapes instead of forcing the crop into the deck. Use the crop only as a reference.

## Wording Contract

Write like a normal academic presentation. Avoid process commentary, prompting language, or advice to the speaker.
Slide body text should describe what the paper does, how the method is constructed, what the experiment reports, and what limitation remains. Do not write meta-commentary about how to present the slide.

Banned or suspicious wording includes:

- `报告口径`, `讲解重点`, `应该强调`, `这页解释了`, `这一页说明`, `可以看到`, `问题是`, `代价是`
- `这篇论文适合借鉴到`, `下一步应该`, `如果结合`, `不是只在仿真中有效`
- `最值得借鉴`, `最有价值`, `一句话概括`, `更像是`, `可以把该框架作为`, `需要监控`, `需要与...配合`
- Any sentence that tells the presenter what to say instead of stating the slide's claim.

Preferred forms:

- Use direct forms such as `本文提出...`, `本文将...写入...`, `实验结果表明...`, `消融实验显示...`, `该方法完成...验证`.
- Replace vague `问题是...` with the actual technical conflict.
- Replace `这页/这些图说明...` with direct evidence interpretation.
- Replace advice like `如果继续做...可以...` with a concrete applicability statement: `对...任务，该框架提供...建模路径`.
- Replace value judgments like `最值得借鉴的是...` with method claims: `本文的建模方式是...`.

Run `scripts/audit_pptx_text.py --strict-body-hierarchy --fail-on-warning` before final delivery for newly generated decks.

## Implementation Pattern

For Python decks, start from `assets/scaffolds/build_ego_deck.py` and `assets/scaffolds/build_v60_deck.py` as examples. Keep reusable helpers:

- `set_run_font`: force Times New Roman for latin, east Asian, and complex scripts.
- `set_textbox`: bullet hierarchy with fixed hanging indents and no empty paragraphs.
- `compact_text`: nonbreaking spaces and no-wrap terms to avoid meaningless line breaks.
- `equation_block`: fixed row positions for formula blocks.
- When adapting user templates, choose a true blank layout when available and remove visible master placeholders such as `单击此处...`, `Click to...`, dates, footers, and slide-number placeholders before adding generated content. If the template keeps sample lab names, author/title/date/footer fields, or slide numbers on the master, cover the generated content canvas with an intentional background layer before placing your own header and body regions.
- Never assume a fixed 13.33 x 7.5 canvas. Read `prs.slide_width` and `prs.slide_height`, then compute all regions from the current template size.
- For editable formulas, prefer a `math_run` or equivalent helper that creates normal runs plus smaller superscript/subscript runs via DrawingML baseline. Avoid code-style formula text such as `p_{k+1}`, `L_clearance`, `J_path`, `GSD_0`, or braces showing in the rendered PNG.
- Formula tables should use display-style symbols (`π`, `λ`, `Pₛ`, `uₙ`, `uₛ`, `Ω`) instead of code-style labels such as `pi_theta`, `u_safe`, `L_smooth`, `J_path`, `GSD_0`, or `P_safe`. If a long English subscript is semantically needed, explain it in the adjacent Chinese text, not in the symbol column.
- `add_table`, `metric_row`, `visual_slide`, and split/table slide helpers.
- `add_pic`: preserve aspect ratio within an explicit box. Choose the box from the image's real width/height; do not reuse one generic figure box for square, wide, and tall figures.
- `assert_layout` or an equivalent shape-overlap checker: run before saving and fix geometry failures instead of bypassing them. This does not replace rendered PNG review.
- A render-budget helper or equivalent manual calculation: estimate mixed Chinese/English line count, use conservative text box heights, and never place a rule/image immediately below the expected final line.
- A slide lint pass before save: reject empty paragraphs, body-text `\n`, standalone bullet markers, oversized paragraph spacing, font sizes outside the section hierarchy, picture regions with poor aspect-ratio fit, unmasked master-sample text, and content slides whose planned occupied area is visibly under target.
- The slide lint pass must also reject empty `p:txBody` nodes, mixed empty/text paragraphs, manual `a:br` line-break elements, code-style math labels with underscores, semantic explanation tables whose meaning column is too narrow, table/metric/plain-body fonts below the profile threshold, long formula rows in narrow formula boxes, narrow body boxes with avoidable wrapping, estimated rendered text overflow, content-shape overlap, fixed-hierarchy font drift, and any shape that exceeds the slide bounds. Run `scripts/audit_pptx_text.py --strict-body-hierarchy --fail-on-warning` on the generated PPTX before rendering.
- For template-specific audits, use the matching profile: `--profile compact` for compact blue decks, `--profile dense-visual` for figure/formula-heavy or dark-cyan decks, and `--profile classic-large` for older large-font blue decks. Do not mark a new deck as legacy to bypass rendered scan warnings.
- Before publishing generated examples, run `scripts/repair_pptx_layout.py` only as a mechanical cleanup pass for empty decorative text bodies, profile-based bullet/table/metric/plain-body font drift, explicit newline replacement, and visible code-style math labels. It reserializes the PPTX through `python-pptx` so LibreOffice can load the repaired package. It does not replace source layout fixes, rendered scan fixes, or manual slide inspection.
- For public regression examples, keep layout-density repairs reproducible. Use `scripts/densify_regression_examples.py` when strict rendered scans expose example-specific blank bands, undersized figures, broken metric regions, or idempotency issues. Do not hand-edit those PPTX assets without also updating the repair script.
- LibreOffice export on Windows can fail on otherwise valid PPTX files when the source/profile path is deep or the package has not been reserialized. Use the export helper in `scripts/run_template_matrix.py`, which stages a `python-pptx`-canonicalized copy under a short workspace-root path before converting to PDF. Treat direct `soffice` failures from deep paths as an environment/export problem to reproduce with short-path staging before changing layout thresholds.
- A privacy lint pass before upload or publication: use `scripts/sanitize_pptx_privacy.py --in-place` for example/generated PPTX files, then run it again with `--check-only --fail-on-warning`. Do not publish PPTX files containing author names, editor names, comments, notes slides, custom properties, or visible personal presenter text.

For visual slides:

- Use a generic `visual_slide` only when the slide has the same top-text / one-image / bottom-text rhythm as the helper expects.
- Write dedicated layout functions for system diagrams, multi-figure comparisons, dense method-loop summaries, hardware-result grids, and any slide with two or more figures.
- Reserve fixed non-overlapping regions for top bullets, figures, separators, and interpretation bullets. Leave at least about 0.12-0.18 in between text and images.
- If a figure makes the page sparse, add concise technical interpretation or choose a better left-right layout; do not enlarge the figure into text, stretch it, or keep a half-empty page.
- If a source figure crop remains awkward after two layout attempts, either re-crop it from the PDF page or omit the fragile caption and explain the evidence in bullets.
- For result figures, every image must have adjacent explanation of what the viewer should learn from it. Do not add a separate caption below the image unless it is already stable in the template rhythm.
- If an image and its explanation cannot both fit cleanly, split the figure evidence and interpretation across two slides. Do not squeeze the image, shrink text below the hierarchy, or place interpretation over the figure.
- For formula slides, use a fixed formula band with row positions, aligned equation labels, and nearby term explanations. Avoid formula rows that look like normal bullets, and avoid mixing formula text sizes on the same page unless a row is explicitly a note.
- Formula explanation should use normal technical wording, not raw variable-code labels. For example use `目标速度 / 推进方向`, not `L_v / L_vmax / L_yaw` as the only visible explanation.

Do not paste a whole paper into slides. Extract the thesis, equations, figures, tables, and results that matter for a report.

## Required QA Gates

Before final response:

1. Build the PPTX without exceptions.
2. Export PPTX to PDF with LibreOffice.
3. Render PDF pages to PNG and make a preview grid using `scripts/render_pptx_previews.py`.
4. Run `scripts/audit_pptx_text.py --strict-body-hierarchy --fail-on-warning` on the PPTX. This must pass with no warnings for suspicious wording, empty text bodies, mixed empty paragraphs, manual newlines, abnormal bullet/table/metric/plain-body font sizes, estimated rendered text overflow, avoidable narrow-column wrapping, body hierarchy drift, or out-of-bounds shapes.
5. After rendering has fully completed, run `scripts/scan_rendered_slides.py` on the rendered PNGs. For template smoke or final multi-template QA, use strict thresholds: `--fail-on-warning --ignore-edge-slides --blank-warn 0.74 --min-band-fraction 0.10 --body-blank-warn 0.89`. Do not launch this in parallel with rendering.
6. Visually inspect the preview grid and every risk slide individually. Risk slides include formulas, tables, multi-image layouts, large figures, hardware/result grids, or any slide changed in this pass.
7. Open the actual figure crops when a rendered slide looks wrong; fix the crop first if it contains captions/body text or cuts off the visual.
8. Fix all blocking issues: overlap, text/image collision, bad image crops, wrong image proportions, meaningless blank lines, excessive blank area, inconsistent cover/ending color, wrong font scale, formula crowding, malformed table proportions, and AI-sounding wording.
9. For repository uploads or shared skill packages, run `scripts/sanitize_pptx_privacy.py` over `assets/examples`, then run `--check-only --fail-on-warning`.
10. After any final edit, inspect at least the cover, thanks slide, all formula slides, all image-heavy slides, and any slide that previously had excessive blank area. Passing automated scripts alone is not enough.
11. For public multi-template regression, run `python scripts/run_template_matrix.py --out-dir out/template-matrix --keep-going`. Treat any non-legacy scan/export/audit failure as a blocking defect.
12. For local template smoke testing, run `python scripts/run_template_smoke.py --templates assets/template-profiles/template-smoke.local.example.json --template-root <template-root> --requirements assets/template-profiles/requirement-smoke.json --keep-going`. Treat any audit/export/render/scan failure as a layout-generalization defect. Use `--template-id`, `--paper-id`, and `--requirement-id` to reproduce a failing combination directly. Do not rerun audits without the matching template profile; `classic-large`, `dense-visual`, and `compact` use different fixed font hierarchies.

The preview grid is not enough for final QA. Open individual slide PNGs when a slide contains formulas, dense tables, multiple figures, any small text, a three-column layout, a metric row, or a slide changed in this pass for blank-area/overlap issues.

## Resources

- `assets/examples/`: AI-generated and sanitized example decks.
- `assets/examples/ego-planner-report.pptx`: accepted EGO-Planner report deck.
- `assets/examples/uav-paper-report-v60.pptx`: accepted earlier UAV paper report deck.
- `assets/examples/uav-rl-privileged-report.pptx`: accepted RL privileged-information quadrotor report deck with editable run-based formulas, native network diagram, and validated density.
- `assets/examples/uav-multi-quad-cbf-report.pptx`: accepted multi-quadrotor cooperative-manipulation CBF report deck with formula/table-heavy method slides, native tables, paper figure crops, and validated blank-area scan.
- `assets/examples/quad-lcd-dark-template-report.pptx`: dark-cyan template adaptation example that must pass PPTX audit, LibreOffice export, PNG render, and rendered-slide scan.
- `assets/examples/uav-fov-cbf-regression-report.pptx`: FoV-CBF certification report deck used as a recent-paper classic-blue regression case.
- `assets/examples/uav-primer-regression-report.pptx`: PRIMER perception-aware multiagent planning report deck used as a recent-paper classic-blue regression case.
- `assets/examples/uav-spot-regression-report.pptx`: SPOT spatio-temporal obstacle-free planning report deck used as a recent-paper classic-blue regression case.
- `assets/template-profiles/matrix.json`: 8-case multi-template regression matrix covering blue large, dense visual, compact formula/table, dark-cyan, and repaired classic-blue styles.
- `assets/template-profiles/public-template-smoke.json`: publishable smoke-test template list built from sanitized example decks.
- `assets/template-profiles/requirement-smoke.json`: smoke-test density/focus variants for balanced, brief, method-detailed, and experiment-detailed requests.
- `assets/scaffolds/*.py`: build scripts from accepted decks. Use as patterns, not as fixed content.
- `references/dependencies.md`: packages and environment checks.
- `references/style-guide.md`: detailed layout and typography rules.
- `references/qa-checklist.md`: final review checklist.
- `scripts/sanitize_pptx_privacy.py`: remove PPTX personal metadata, notes, comments, custom properties, and visible presenter names before publishing.
- `scripts/run_template_matrix.py`: run template-aware audit, render, and scan across the registered examples.
- `scripts/run_template_smoke.py`: generate and QA local stress decks across multiple PPTX templates and requirement variants without publishing private templates.
- `scripts/repair_pptx_layout.py`: mechanically clean generated PPTX examples before publication; use after source layout fixes, not instead of them.
- `scripts/densify_regression_examples.py`: reproducibly apply public regression example density fixes after strict rendered-scan failures.
