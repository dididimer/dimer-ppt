# dimer-ppt

Codex skill for converting dense screenshots and figures into high-fidelity editable PowerPoint decks.

The skill reconstructs an image as editable PowerPoint objects where practical, then runs a PowerPoint-based render and layout review loop. It is designed for dense academic diagrams, mixed English/CJK labels, and figures where ordinary image-to-PPT conversion tends to create text overflow or icon collisions.

## What It Does

- Builds a source-derived `visual_spec.md`.
- Creates an element `manifest.json`.
- Generates editable PowerPoint shapes, text boxes, arrows, charts, and callouts.
- Preserves only small complex icon crops when full editable redrawing would reduce fidelity.
- Exports PowerPoint previews for source/render comparison.
- Runs layout guards for text overflow, overlap, out-of-bounds objects, and picture/shape counts.

## Install

Clone or copy this repository into your Codex skills directory:

```powershell
git clone https://github.com/dididimer/dimer-ppt.git "$env:USERPROFILE\.codex\skills\dimer-ppt"
```

Then restart or refresh Codex so the skill list reloads.

## Usage

Ask Codex to use `dimer-ppt`:

```text
请使用 dimer-ppt skill，把这张图转为可编辑 PPT，要求排版和图片一致。
```

## Notes

- On Windows, the full workflow uses local PowerPoint COM automation to create and render `.pptx` files.
- PowerPoint may briefly open and close during generation or validation; that is expected.
- Hybrid-fidelity mode may keep small raster crops for complex cartoon icons, while text, panels, arrows, charts, tokens, and layout containers remain editable.

