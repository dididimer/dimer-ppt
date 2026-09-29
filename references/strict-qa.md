# Strict reconstruction QA

Use this reference for dense figures or when the user asks for a close visual match. Keep each artifact on disk. Do not return a complete manifest, tool log, or all crops to the model after every pass.

## Artifacts

- `visual_spec.md` or a compact equivalent in the manifest: canvas, major regions, font hierarchy, source colors, strokes, icon style, and density.
- `manifest.json`: stable IDs, geometry, text, style or shared style reference, z-order, connector anchors, and each crop's source/target bounds and reason.
- `skeleton.pptx` and preview: major boxes and connectors before detailed styling.
- `styled.pptx` and a PowerPoint/WPS render at the source aspect ratio.
- Layout and editability counts, a source/render review sheet, and ranked region crops. Generate all tiles for a strict audit if needed; inspect the ranked index and review suspicious regions first.

## Acceptance gates

- Every major visible region appears in the manifest. Correct panel or connector drift greater than roughly 2% of slide width or height before styling.
- No unintended text clipping, unexpected wrap, missing character, or text outside its intended container. PowerPoint text metrics differ from image metrics; give a text box enough slack and verify the render.
- The rendered slide preserves meaningful icon/text, connector, and container relationships. Numeric pixel differences are triage signals; antialiasing, font substitution, and small registration offsets can raise them without a semantic error.
- Count pictures and explain preserved crops. A full-slide source image is valid only in `reference-overlay` mode.
- If Office render export is blocked, report `render gate: blocked`, deliver a clearly named draft, and do not describe it as visually verified.

The user's screenshot of a defect is stronger evidence than a previous self-report. Reopen the generated deck and repair the source manifest or generator.

## Review and repair

1. Inspect the full source and render for global layout and visual density. Make a review sheet locally when it helps diagnose a global mismatch.
2. Run the geometry guard. Warnings are candidates for review, not automatic failures. For any accepted overlap, record its intended source relationship and confirm the rendered relationship matches.
3. Read `review_summary.json` from the region-tile script. Inspect high-scoring or semantically sensitive regions first. Use local issue crops from `scripts/make_guard_issue_crops.py` when a guard warning needs visual judgment.
4. Repair in this order: global geometry; typography; connectors; icon language; palette and strokes; micro details. Fix the local region instead of rescaling the whole slide to address a single text problem.
5. Regenerate and rerender after a change. Stop repeating a pass that does not reduce the remaining issues; report unresolved differences as a draft.

Stable shape names (`DIMER_TXT_*`, `DIMER_PANEL_*`, `DIMER_LINE_*`, `DIMER_ICON_*`, `DIMER_RASTER_*`) make reports and targeted repairs easier to interpret.
