---
name: dimer-ppt-skills
description: Convert images/screenshots into high-fidelity editable PowerPoint using CJK-safe shape reconstruction, VBA/native PPT generation, layout guards, and visual review.
triggers:
  - dimer ppt
  - dimer-ppt
  - dimer-ppt-skills
  - image to editable ppt
  - screenshot to editable powerpoint
  - PNG to editable PPT
  - fix xiaobei ppt
  - improve ppt-master output
  - Chinese text overflow in PPT
  - PPT text overlap
---



# Dimer PPT Skills

This skill reconstructs an image as editable PowerPoint shapes with an explicit
layout-safety loop. It is an end-to-end image-to-editable-PPT skill, not only a
post-checker. It exists because image-to-PPT tools often succeed at editability
but fail at Chinese text fitting, dense labels, and icon/text collisions.

Geometric overlap detection is necessary but not sufficient. Some text
intentionally sits inside shapes in the source image, while other overlaps are
layout failures. This skill must compare the rendered PPT against the original
image before deciding whether a remaining overlap is acceptable.

## Scope

Use this skill to produce the final `.pptx` from a source image. The skill must
cover the same core job as `xiaobei-skill-image-to-vba`:

- Inspect a source image.
- Produce an element manifest.
- Generate editable PowerPoint objects through VBA/native Office shapes.
- Materialize a `.pptx` when local PowerPoint automation is available.
- Deliver the `.bas` macro source when VBA is generated.

Then it adds a required safety layer:

- Expand CJK/mixed text boxes with safe slack.
- Detect text overflow using PowerPoint's own text engine.
- Detect text/text and text/icon overlaps.
- Repair overflow where possible and regenerate for structural overlap.
- Render and visually compare against the source image so intentional overlaps
  are separated from accidental collisions.

## Strategy

Prefer this order:

1. Use `xiaobei-skill-image-to-vba` as the preferred generation basis when it is
   installed. Read its `SKILL.md` only as needed, then add this skill's safe text
   rules and guard loop on top.
2. If `xiaobei-skill-image-to-vba` is not installed, implement the same pattern
   directly: manifest -> VBA/native shapes -> PowerPoint materialization.
3. Use `ppt-master` SVG-to-PPTX only when it gives better structure for the page.
4. Do not use a full-slide screenshot as the final object when the user asks for a truly editable deck.

Use preserved raster crops only for complex artwork that would be bad as shapes.
Never preserve the whole source image unless the user explicitly accepts a
background-image-assisted result.

## Mandatory Artifacts and Hard Gates

Do not jump straight from source image to final PPT. A valid reconstruction run
must produce these artifacts in order:

1. `visual_spec.md` or `visual_spec.json`
   - Canvas dimensions and pixel-to-slide scale.
   - Source-derived palette with sampled coordinates.
   - Font hierarchy: title, subtitle, panel title, card label, bilingual
     micro-label, number, equation.
   - Stroke rules: border weights, connector weights, dash rhythm, arrowhead
     style.
   - Icon language: cartoon outline, filled flat, line-art, raster-preserved,
     etc.
   - Region density targets: dense / medium / sparse for each major panel.

2. `manifest.json` or `manifest.md`
   - Every major element has `id`, `type`, `bucket`, `bbox_px`,
     `safe_bbox_px`, `style`, `text`, and `z_order`.
   - Every connector/arrow has source pixel endpoints and semantic anchors.
   - Every raster-preserved crop has `crop_path`, `bbox_px`, `target_bbox`,
     and `preserve_reason`.

3. `skeleton.pptx` and `skeleton_preview.png`
   - Major bboxes, text boxes, connectors, and preserved-crop slots are drawn
     as simple labeled objects.
   - No styling pass may begin until skeleton layout is visually close to the
     source.

4. `styled.pptx` and `styled_preview.png`
   - Main editable/hybrid reconstruction after skeleton approval.

5. `source_render_review_sheet.png` plus tiled region crops
   - Source, render, and diff side by side.
   - Dense figures must have region tiles even when layout guard is quiet.

6. `layout_guard.json` or equivalent report
   - Reports text overflow, text/text overlap, text/icon overlap, negative
     extents, out-of-bounds objects, picture count, shape count, and text count.

Hard gates:

- **Spec gate**: no `visual_spec` = stop.
- **Manifest gate**: no manifest row for a visible region = stop and parse the
  missing region.
- **Skeleton gate**: if major panel/card bbox drift is visually obvious or
  exceeds about 2% of slide width/height, stop and fix coordinates before
  styling.
- **Text gate**: any unintended overflow, clipping, auto-shrink, or unexpected
  wrapping = stop and repair `safe_bbox_px` / manual line breaks.
- **Region gate**: if any major region has obvious source/render drift, create a
  local crop and fix that region before final delivery.
- **Density gate**: if the source is dense but the render looks sparse or
  simplified, stop and add missing details or switch selected complex icons to
  hybrid crops.
- **Icon-language gate**: if source icons are cartoon outline but render uses
  generic Office symbols/emoji/flat placeholders, stop and rebuild the icon
  family or preserve minimal crops.
- **Disclosure gate**: if hybrid crops are used, final report must list picture
  count and which elements are not editable.

Passing XML stats alone is not enough. A deck can be technically editable and
still fail fidelity.

## Reconstruction Mode Decision

Choose and record one mode before generation:

| Mode | Use when | Required disclosure |
|---|---|---|
| `editable-strict` | User requires every visible object to be editable, and accepts lower icon/texture fidelity. | State that complex icons may be approximate. Picture count should be 0 except unavoidable embedded media. |
| `hybrid-fidelity` | User wants close visual match and complex icons/illustrations would look poor as Office shapes. | List all preserved crops and picture count. Text, panels, arrows, charts, and labels remain editable. |
| `pixel-reference-only` | User explicitly asks for a visual reference slide or QA page. | Must not be described as editable reconstruction. |

Default for dense academic infographic screenshots: `hybrid-fidelity` unless the
user explicitly says all objects must be editable. Default for simple diagrams:
`editable-strict`.

Mode cannot change mid-run without updating the manifest and final disclosure.

## Required Workflow

1. Inspect the image and record canvas size, aspect ratio, regions, text, icons,
   panels, arrows, and repeated grids.
2. Create an element manifest before generating the deck. Every text element must
   include `bbox_px`, text content, font estimate, and a `safe_bbox_px` that is
   wider/taller than the visual text by default.
3. Generate an editable deck using native PowerPoint objects:
   - Text: real text boxes.
   - Panels and simple icons: shapes, freeforms, lines, connectors.
   - Arrows/connectors: editable lines/connectors.
   - Complex artwork: smallest useful raster crop only, with explicit disclosure.
4. Materialize the deck:
   - Windows + PowerPoint: try VBA import/run first when using `.bas`.
   - If VBA import is blocked, use PowerPoint COM direct materialization from the
     same manifest/shape plan rather than falling back to a picture.
   - WPS-only: provide `.bas` plus `.pptx` fallback when possible; do not assume
     automatic macro import.
5. Use a CJK-safe text policy:
   - Prefer `Microsoft YaHei` for Chinese UI/diagram text on Windows.
   - Make each Chinese or mixed CJK/Latin text box 15-25% wider than the apparent SVG/image text.
   - Reduce dense label font sizes by 1-2 pt compared with the visual estimate.
   - For short badges and labels, set `WordWrap = False` only when the box is wide enough.
   - For long bilingual labels, manually split lines instead of relying on PowerPoint wrapping.
6. Run automatic text repair before judging the output:
   - `scripts/repair_text_layout_windows.ps1 -Pptx <deck.pptx> -Output <fixed.pptx>`
   - Use this especially after `xiaobei`-style generation, where text boxes often
     fit too tightly.
7. Run layout guard:
   - `scripts/pptx_xml_stats.py <deck.pptx>`
   - `scripts/pptx_layout_guard_windows.ps1 -Pptx <deck.pptx> -Report <report.json>`
8. Export a PowerPoint PNG render at the source image size.
9. Create a visual review sheet:
   - `scripts/make_visual_review_sheet.py --source <source.png> --render <render.png> --out <sheet.png>`
10. Use source-aware visual review on remaining warnings:
   - Mark an overlap as acceptable only if the source image shows the same
     semantic relationship, such as a label intentionally centered inside a badge.
   - Mark it as a defect when text covers an icon, chart, or another label in
     the generated render but not in the source.
   - Prefer local region comparison around each warning; whole-slide metrics are
     only a rough signal.
   - For guard warnings, create per-issue local crops:
     `scripts/make_guard_issue_crops.py --source <source.png> --render <render.png> --report <guard.json> --outdir <issue_crops_dir>`
   - Even when guard is quiet, create tiled source/render crops for dense pages:
     `scripts/make_region_review_tiles.py --source <source.png> --render <render.png> --outdir <tiles_dir>`
   - Review high-risk regions explicitly: bottom KPI/cards, formula panels,
     icon+label cards, candidate ranking blocks, and any area with text inside a
     small card.
11. If the report contains real overflow or severe overlaps:
   - First fix the source manifest/coordinates and regenerate.
   - If only text is slightly too large, rerun with `-FixOverflow`.
   - Render again and repeat until remaining warnings are intentional.

## Naming Contract

Generated shapes should use stable prefixes:

- `AITVBA_TXT_*` for normal text boxes.
- `AITVBA_ICON_*` for icon shapes.
- `AITVBA_PANEL_*` for panel/card backgrounds.
- `AITVBA_LINE_*` for lines/arrows/connectors.
- `AITVBA_DECOR_*` for decoration.
- `AITVBA_RASTER_*` for preserved crops.

This makes collision reports easier to read. If using an existing generator that
does not follow this contract, still run the guard; it will fall back to generic
shape names.

## End-To-End Deliverable Gate

Do not stop after writing a macro or guard report if local PowerPoint can create
the deck. A completed run should include:

1. Editable `.pptx` created from the source image.
2. `.bas` or generator source when applicable.
3. Element manifest.
4. XML stats proving the file is not a full-slide picture shortcut.
5. Layout guard report after repair.
6. Rendered PNG and/or visual review sheet.
7. A concise note of remaining visual differences and which overlaps were kept
   because they match the source.

If automation is blocked, state the blocker and still deliver the `.bas` plus
manual run steps.

## Guard Interpretation

- `overflow`: text's PowerPoint bounds exceed its text box. Must fix unless the
  text is decorative and intentionally clipped.
- `text_text_overlap`: two non-decorative text boxes overlap. Usually fix by
  moving, splitting lines, shrinking, or widening safe boxes.
- `text_shape_overlap`: text overlaps an icon/shape. Fix when overlap is large
  or the shape is not the intended background panel.

Small decorative text such as `+`, `-`, `!`, single digits, and formula glyphs may
legitimately sit inside shapes. The guard suppresses most of these.

Do not mark a text/icon overlap as intentional only because it is geometrically
inside a nearby shape. Check the source render. Examples:

- Acceptable: `VS` centered in the original purple circle.
- Defect: `鏇村ソ / A-Prefer / 56.2` covering a thumbs-up illustration when the
  original places the text beside or below the icon.
- Defect: candidate labels crossing bars or heatmap cells when the source has
  clear separation.

## Visual Review Contract

For each remaining guard warning, inspect a local region in both source and
render. Record:

- `same_as_source`: yes/no.
- `source_relation`: inside badge, beside icon, below chart, on top of icon, etc.
- `render_relation`: same labels for the generated PPT.
- `decision`: keep / fix.
- `fix`: move, widen, shrink, split, send behind, or simplify nearby icon.

The model should use vision here. A script can surface suspicious geometry, but
the final decision about whether overlap is faithful must be source-aware.

## Fidelity Drift Diagnosis

Before making another local fix, classify the source/render difference. This
prevents "one more font tweak" loops when the real problem is generation method
or visual language.

- `global_layout_drift`: panel sizes, gutters, margins, vertical compression, or
  bottom-strip height differ from the source.
- `visual_density_drift`: the source is dense with small marks/details, but the
  reconstruction looks sparse, simplified, or posterized.
- `typography_drift`: title/header/label hierarchy differs; Chinese and English
  lines have different relative scale; labels wrap differently.
- `icon_language_drift`: source uses outlined cartoon/academic infographic
  icons, but render uses generic flat shapes, emoji, or simplified symbols.
- `palette_stroke_drift`: fills, borders, arrows, dash rhythms, shadows,
  gradients, or stroke widths differ.
- `semantic_local_drift`: text is inside/beside/under an icon differently than
  in the source, even when no geometric overflow is reported.

Fix in this order:

1. global layout,
2. typography scale,
3. icon scale and icon language,
4. palette/strokes/effects,
5. micro-label containment.

If a review sheet shows strong differences across all quadrants, treat it as a
manifest/spec failure rather than a one-off local bug.

## Spec Lock + Skeleton Requirement

Borrow the `ppt-master` discipline before drawing:

- Create a compact source-derived spec lock: canvas, coordinate scale, sampled
  palette, font hierarchy, stroke widths, dash rhythm, icon language, repeated
  icon scales, and object density by major region.
- Re-read or reference this spec before generating each major region. Do not
  invent new colors, sizes, or icon styles from memory mid-slide.

Borrow the `xiaobei` discipline before styling:

- Create an element manifest with `bbox_px`, `safe_bbox_px`, style, text,
  z-order, and connector endpoints.
- Validate a skeleton pass first: draw major bboxes/connectors as simple gray
  objects with ids. If the skeleton is off, fix the manifest/coordinates before
  styling.
- A styled reconstruction with a bad skeleton is not acceptable, even if all
  objects are editable.

## Local Repair Loop

When visual review marks an issue as `fix`, do not globally rescale the slide.
Patch only the affected local region:

1. Identify the smallest region that contains the wrong relationship.
2. Compare the source crop and render crop.
3. Decide the semantic repair:
   - Move label to match source relative position.
   - Widen text box without covering neighboring icon.
   - Shrink or simplify icon if it crowds text.
   - Split text into two lines if source does so.
   - Reorder z-index only when source proves the element should be behind.
4. Regenerate the deck from the manifest/generator, not by manually dragging in
   PowerPoint unless the user explicitly asks for manual edits.
5. Re-export, rerun guard, regenerate issue crops, and repeat.

Acceptance gate: a deliverable should have zero unresolved `fix` decisions. Any
remaining guard warning must be listed with `same_as_source=yes` and a reason.
For dense diagrams, also require tiled region review; silent layout drift that is
not detected by geometry still counts as a defect if a tile visibly differs in
text/icon relationship from the source.

## Deliverables

Always report:

- Final editable `.pptx`.
- Manifest path.
- Macro/VBA source path if generated.
- Layout guard report path.
- Whether full-slide image background was used.
- How many pictures, shapes, and text nodes are in the PPTX.
- Remaining layout warnings and whether they are intentional.
- Visual review sheet or render comparison path when fidelity was evaluated.
- Per-issue crop directory and issue review decisions when overlap warnings were present.
- Region tile review directory for dense diagrams, even when no guard issue remains.

Use this final report shape:

```markdown
## Output
- PPTX:
- Preview:
- Mode: editable-strict / hybrid-fidelity / pixel-reference-only

## Editability
- Shapes:
- Text runs:
- Pictures:
- Non-editable preserved crops:

## Gate Status
- visual_spec:
- manifest:
- skeleton:
- styled render:
- layout guard:
- review sheet:
- region tiles:

## Remaining Differences
- Global layout:
- Typography:
- Icon language:
- Palette/strokes:
- Region-specific issues:
```
