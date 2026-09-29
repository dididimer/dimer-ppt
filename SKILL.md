---
name: dimer-ppt
description: Reconstruct screenshots, diagrams, and dense infographics as editable PowerPoint slides. Use for image-to-editable-PPT requests and for repairing visual or text-layout drift in those slides.
---

# Dimer PPT

Rebuild the source image with editable PowerPoint text, panels, connectors, and repeated structures. Preserve a complex icon or illustration as a small crop when redrawing it would hurt fidelity. Never present a full-slide source image as an editable reconstruction.

## Choose fidelity and effort

Record both choices in the manifest. Honor a user's explicit editability or fidelity requirement.

| Fidelity | Use when | Report |
| --- | --- | --- |
| `editable-strict` | Clean diagrams or every visible object must be editable. | Any simplified artwork. |
| `hybrid-fidelity` | Dense figures with distinctive illustrations. | Each preserved crop and the picture count. |
| `reference-overlay` | The user explicitly accepts an image reference. | That the slide is not an editable reconstruction. |

| Effort | Use when | QA scope |
| --- | --- | --- |
| `fast` | Simple, sparse diagram with clear text and few unique objects. | One build and render, local checks, then repair actual failures. |
| `balanced` | Default for ordinary screenshots and diagrams. | Use a skeleton only when layout is uncertain; inspect the full render once, then inspect ranked problem regions. |
| `strict` | Dense academic figures or explicit close-match requests. | Use a skeleton and the full review process in [strict QA](references/strict-qa.md). |

Effort controls inspection work, not the editability promise. If a fast or balanced run reveals complex layout or repeated failures, raise its effort level and record why.

## Build

1. Inspect the source once for canvas ratio, major regions, text, colors, strokes, connectors, and repeated objects. Use source crops only where details cannot be read at overview size.
2. Save a compact `manifest.json` with the chosen modes, source and slide sizes, semantic regions, editable objects, connector anchors, and any preserved crops. Reuse shared styles and repeated components. Keep the full manifest in a file; return only a short summary to the model after local tools consume it.
3. Generate the PPTX from that source of truth. Use native shapes and editable text for semantic content. Before preserving an icon as a picture, check whether its parts can be drawn as editable shapes or reused from a native component. Read [icon editability](references/icon-editability.md) when the source contains picture icons. For `balanced` or `strict`, create and inspect a skeleton first when major geometry is uncertain; correct geometry before styling.
4. Export a PowerPoint/WPS render at the source aspect ratio. If rendering is unavailable, label the result `*_draft.pptx` and report that visual QA is unverified.

## Check and repair

- Run `scripts/pptx_xml_stats.py` and, where PowerPoint automation is available, `scripts/pptx_layout_guard_windows.ps1`. Treat their reports as triage, not proof of visual fidelity.
- Compare the source and render at the same aspect ratio. For `balanced` and `strict`, run `scripts/make_region_review_tiles.py --source <source> --render <render> --outdir <dir> --top-k 3`. Read `review_summary.json` first; show the model only the highest-priority crops needed to diagnose a real discrepancy. Use `--all` when strict review or unresolved semantic drift requires all regions.
- Fix clipping, missing text, wrong object relationships, and material layout drift. Apply deterministic text or geometry fixes locally when possible. Change only affected manifest objects or generator code, then rerender.
- After two broad repair passes, switch to targeted regions. If a pass makes no measurable progress, identify the remaining cause instead of repeating the same full-slide review. An unresolved hard failure makes the output a draft, not a verified final slide.

For dense or difficult figures, read [strict QA](references/strict-qa.md). For CJK or mixed-script fitting, read [layout rules](references/layout-rules.md).

## Deliver

Provide the PPTX and a rendered preview. Report the fidelity and effort modes; editable shape/text and picture counts; preserved crops; QA result; and any remaining visible differences. Do not claim that XML or geometry checks alone verify the appearance.
