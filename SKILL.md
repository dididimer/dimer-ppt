---
name: dimer-ppt
description: Convert screenshots, diagrams, and infographic images into high-fidelity editable PowerPoint decks with manifest-first reconstruction, native PPT shapes, render-diff QA, and strict text-containment checks. Use when the user asks for image-to-editable-PPT, screenshot-to-PowerPoint, editable PPT recreation, layout repair, Chinese/CJK text fitting, or fixing text/icon overlap in generated slides.
---

# Dimer PPT

Reconstruct a source image as an editable PowerPoint slide or deck. The goal is
visual fidelity plus editability: panels, labels, connectors, arrows, tables,
badges, and repeated diagram structure should be native PowerPoint objects.
Use minimal raster crops only when a small complex illustration would be worse
as shape art. Never satisfy an editable-PPT request by placing the full source
image on a slide.

This skill is intentionally self-contained for public use. It may integrate with
local Office automation and bundled QA scripts, but its instructions must not
copy private or third-party skill text, examples, tables, or code unless that
material is explicitly licensed for reuse.

## Operating Modes

Choose one mode before generation and record it in the manifest.

| Mode | Use When | Disclosure |
|---|---|---|
| `editable-strict` | All visible objects should be editable and approximate icons are acceptable. | State that complex icon artwork may be simplified. |
| `hybrid-fidelity` | Complex logos, cartoons, photos, screenshots, or textured artwork need closer visual fidelity. | List each raster crop and picture count. |
| `reference-overlay` | The user explicitly accepts a source image as a reference/background. | Do not describe the result as a fully editable reconstruction. |

Default to `hybrid-fidelity` for dense academic or UI screenshots with many
distinctive small illustrations. Default to `editable-strict` for clean diagrams,
flowcharts, tables, and simple infographic layouts.

## Required Artifacts

Every serious reconstruction run must create these artifacts before final
delivery:

1. `visual_spec.md` or `visual_spec.json`
   - Source size, target slide size, scale, margins, major regions.
   - Source-derived colors, font hierarchy, stroke widths, dash rhythm, arrow
     style, and icon language.
   - Region density targets: sparse, medium, or dense.

2. `manifest.json` or `manifest.md`
   - Each visible region has `id`, `type`, `bucket`, `bbox_px`,
     `safe_bbox_px`, `style`, `text`, and `z_order`.
   - Each connector has pixel endpoints and semantic anchors.
   - Each raster crop has `crop_path`, source bbox, target bbox, and reason.

3. `skeleton.pptx` and `skeleton_preview.png`
   - Draw major bboxes and connectors as simple labeled objects.
   - Do not start styling until the skeleton visibly matches the source layout.

4. `styled.pptx` and `styled_preview.png`
   - Final styled reconstruction exported from PowerPoint at source-image size.

5. `source_render_review_sheet.png`
   - Side-by-side source, PowerPoint render, and amplified diff.

6. Region review tiles
   - Required for dense diagrams even if the geometry guard is quiet.

7. `layout_guard.json`
   - Reports overflow, overlaps, out-of-bounds objects, shape count, text count,
     and picture count.

## Hard Gates

A deck is not done until all gates pass:

- **Spec gate**: no visual spec means stop.
- **Manifest gate**: if a visible region is missing from the manifest, stop and
  parse it.
- **Skeleton gate**: if major panels, cards, or connector endpoints drift by
  more than about 2% of slide width/height, fix the manifest before styling.
- **Text gate**: any unintended overflow, clipping, auto-shrink, lost character,
  unexpected wrap, or text outside its intended container is a failure.
- **Render gate**: actual PowerPoint render must be inspected. XML stats and
  layout-guard output are not enough.
- **No-render fallback gate**: if PowerPoint/WPS render export is blocked, do
  not label the deck final. Deliver only a clearly marked draft, or pause and
  ask for render access/time to retry.
- **Region gate**: if any review tile shows text/icon/container relations that
  differ from the source, fix the local region.
- **Editability gate**: report picture count and explain every non-editable
  raster crop. A full-slide source image is not acceptable unless the chosen
  mode is `reference-overlay`.

## Render Comparison Rules

The final judge is the PowerPoint-rendered image, not the generator's internal
geometry. Export the slide at the same pixel size as the source image and inspect
`styled_preview.png` plus `source_render_review_sheet.png`.

Rendered preview hard failures:

- Text appears outside a card, box, badge, process node, score cell, or label
  area when the source keeps it inside.
- A text baseline crosses a border that it should not cross.
- Captions such as `Support?`, `Complete?`, `Extractor`, `Retriever`, or
  `Helpful/Logical?` attach to the outside of a node instead of sitting where
  they appear in the source.
- Group headers such as `Context_Chunks` or `Related Chunks` collide with the
  first item in their group.
- Scores or short labels wrap, truncate, or lose characters.
- Icon and text relationships differ semantically from the source, even if the
  geometry guard does not complain.

If the user provides a screenshot of the PPT canvas showing a defect, treat it
as stronger evidence than prior self-report. Reopen the generated deck, reproduce
the issue, and repair the manifest or generator.

If render export is unavailable because of sandboxing, Office automation,
quota, or a missing presentation app, the work may continue only as a draft
construction pass. In that case:

- Say `render gate: blocked` in the QA report.
- Do not claim the result is visually verified.
- Prefer waiting/retrying Office render export over inventing a new lower-quality
  generation path.
- If a non-rendered OOXML fallback is produced, name it `*_draft.pptx` unless the
  user explicitly accepts it as final after opening it themselves.

## Text Containment Policy

PowerPoint text metrics differ from screenshots and SVG text. Build text with
slack from the start.

- Give every text element its own `safe_bbox_px`, normally 20-35% wider than the
  apparent source text and at least 1.3x its expected rendered height.
- Prefer separate text boxes over parent-shape text for small labels, badges,
  process captions, and mixed icon/text nodes.
- For one-line labels, disable wrapping only when the safe box is wide enough.
- For two-line labels, split into two text boxes when PowerPoint inflates the
  combined text bounds.
- For CJK or mixed CJK/Latin text, use a CJK-safe font on Windows, reduce dense
  labels by 1-2 pt, and add extra width.
- Tiny decorative micro-text inside icons can be omitted in `editable-strict`;
  do not keep it if it creates overflow.
- If an icon crowds a label, shrink the icon, move the label, or switch that icon
  to a small crop in `hybrid-fidelity`.

## Workflow

1. Inspect the source image. Record canvas size, aspect ratio, major regions,
   repeated elements, text hierarchy, colors, strokes, arrows, and icon style.
2. Choose reconstruction mode and target slide size. Use ratio-safe coordinate
   mapping unless the source ratio exactly matches the target canvas.
3. Write the visual spec and manifest. Do not generate the deck before the
   manifest covers all visible regions.
4. Generate the skeleton deck from the manifest. Export a skeleton preview and
   compare it with the source.
5. Generate the styled deck using native PowerPoint shapes where possible:
   - Text: editable text boxes.
   - Panels/cards: editable rectangles, rounded rectangles, freeforms.
   - Connectors/arrows: editable lines or connectors.
   - Repeated icons: editable shape families or minimal raster crops.
6. Export the styled deck to PNG from PowerPoint at the source image size.
   If this step is blocked, stop final delivery or mark the output as draft.
7. Run bundled QA scripts:
   - `scripts/pptx_xml_stats.py <deck.pptx>`
   - `scripts/pptx_layout_guard_windows.ps1 -Pptx <deck.pptx> -Report <report.json>`
   - `scripts/make_visual_review_sheet.py --source <source> --render <render> --out <sheet>`
   - `scripts/make_region_review_tiles.py --source <source> --render <render> --outdir <dir>`
   - `scripts/make_guard_issue_crops.py --source <source> --render <render> --report <guard> --outdir <dir>` when guard warnings need local review.
8. Repair locally. Fix the manifest/generator, regenerate, re-export, and rerun
   checks. Do not hand-edit the final PPT unless the user asks for manual edits.

## Guard Interpretation

The guard is a triage tool, not the final judge.

- `overflow`: fix unless the source clearly clips the same decorative text.
- `text_text_overlap`: fix when two independent labels collide in the render.
- `text_shape_overlap`: fix when text covers an icon, connector, or unrelated
  shape differently from the source.
- Warnings involving text intentionally inside its own card or badge may be kept
  only after source/render crop review confirms `same_as_source=yes`.

For every remaining warning, record:

- `same_as_source`: yes/no.
- `source_relation`: inside card, above box, below icon, centered in badge, etc.
- `render_relation`: the same relation in the generated preview.
- `decision`: keep/fix.
- `reason_or_fix`: concise explanation or repair action.

No final response may say "warnings are intentional" without this source-aware
review.

## Repair Priorities

When the render differs from the source, repair in this order:

1. Global layout: canvas, panel sizes, margins, gutters, and row heights.
2. Typography: title sizes, label sizes, line breaks, and safe text boxes.
3. Connectors: endpoints, arrowheads, elbow routes, and z-order.
4. Icon language: scale, placement, density, and crop-vs-shape choice.
5. Palette and strokes: fills, borders, dash style, line weight, shadows.
6. Micro details: badge text, tiny decorative marks, and repeated motifs.

Do not globally rescale the slide to fix a local text failure. Patch the local
region and regenerate from the same source of truth.

## Shape Naming

Use stable prefixes so reports are readable:

- `DIMER_TXT_*` for text boxes.
- `DIMER_PANEL_*` for panels, cards, and boxes.
- `DIMER_LINE_*` for connectors, arrows, and guide lines.
- `DIMER_ICON_*` for editable icon parts.
- `DIMER_RASTER_*` for preserved image crops.
- `DIMER_DECOR_*` for nonsemantic decoration.

## Public Release Hygiene

For a GitHub release:

- Keep this skill in original wording. Do not include private/local skill text,
  copied examples, copied tables, or garbled upstream content.
- Bundle only scripts you wrote or scripts whose license permits redistribution.
- Add license headers to scripts if they are intended for public release.
- If the skill optionally works with external tools, describe them as optional
  integrations in README/NOTICE, not as copied dependencies.
- If provenance is uncertain, rewrite or remove the content before publishing.

## Final Report Template

Use this structure when delivering a result:

```markdown
## Output
- PPTX:
- Preview:
- Mode:

## Editability
- Shapes:
- Text nodes:
- Pictures:
- Preserved crops:
- Full-slide source image used: yes/no

## QA Artifacts
- visual_spec:
- manifest:
- skeleton:
- styled render:
- XML stats:
- layout guard:
- review sheet:
- region tiles:
- issue crops:

## Remaining Differences
- Global layout:
- Typography:
- Icons:
- Palette/strokes:
- Accepted warnings:
```
