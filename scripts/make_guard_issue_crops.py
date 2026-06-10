#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw


def load_box(issue: dict) -> tuple[float, float, float, float] | None:
    boxes = []
    for key in ("box", "a_box", "b_box", "text_box", "other_box"):
        b = issue.get(key)
        if isinstance(b, dict):
            boxes.append(b)
    if not boxes:
        return None
    left = min(float(b["left"]) for b in boxes)
    top = min(float(b["top"]) for b in boxes)
    right = max(float(b.get("right", b["left"] + b["width"])) for b in boxes)
    bottom = max(float(b.get("bottom", b["top"] + b["height"])) for b in boxes)
    return left, top, right, bottom


def scale_box(box: tuple[float, float, float, float], sx: float, sy: float, pad: int, width: int, height: int) -> tuple[int, int, int, int]:
    l, t, r, b = box
    x1 = max(0, int(l * sx) - pad)
    y1 = max(0, int(t * sy) - pad)
    x2 = min(width, int(r * sx) + pad)
    y2 = min(height, int(b * sy) + pad)
    if x2 <= x1:
        x2 = min(width, x1 + 8)
    if y2 <= y1:
        y2 = min(height, y1 + 8)
    return x1, y1, x2, y2


def annotate(im: Image.Image, title: str) -> Image.Image:
    h = 34
    out = Image.new("RGB", (im.width, im.height + h), "white")
    out.paste(im, (0, h))
    draw = ImageDraw.Draw(out)
    draw.text((8, 9), title, fill=(0, 0, 0))
    draw.rectangle([0, h, out.width - 1, out.height - 1], outline=(80, 80, 80), width=2)
    return out


def main() -> int:
    parser = argparse.ArgumentParser(description="Create local source/render/diff crops for each layout guard issue.")
    parser.add_argument("--source", required=True)
    parser.add_argument("--render", required=True)
    parser.add_argument("--report", required=True)
    parser.add_argument("--outdir", required=True)
    parser.add_argument("--slide-width-pt", type=float, default=960)
    parser.add_argument("--slide-height-pt", type=float, default=640)
    parser.add_argument("--pad", type=int, default=36)
    parser.add_argument("--max-issues", type=int, default=50)
    args = parser.parse_args()

    source = Image.open(args.source).convert("RGB")
    render = Image.open(args.render).convert("RGB").resize(source.size, Image.Resampling.LANCZOS)
    diff = ImageChops.difference(source, render).point(lambda p: min(255, p * 4))
    sx = source.width / args.slide_width_pt
    sy = source.height / args.slide_height_pt

    report = json.loads(Path(args.report).read_text(encoding="utf-8-sig"))
    issues = report.get("issues", [])[: args.max_issues]
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    index = []
    for idx, issue in enumerate(issues, 1):
        box = load_box(issue)
        if box is None:
            continue
        crop_box = scale_box(box, sx, sy, args.pad, source.width, source.height)
        src_crop = source.crop(crop_box)
        ren_crop = render.crop(crop_box)
        dif_crop = diff.crop(crop_box)
        title = f"{idx:02d} {issue.get('kind','issue')} | {issue.get('text') or issue.get('a_text','')}"
        src_panel = annotate(src_crop, "SOURCE")
        ren_panel = annotate(ren_crop, "RENDER")
        dif_panel = annotate(dif_crop, "DIFF x4")
        sheet = Image.new("RGB", (src_panel.width * 3 + 24, src_panel.height), "white")
        sheet.paste(src_panel, (0, 0))
        sheet.paste(ren_panel, (src_panel.width + 12, 0))
        sheet.paste(dif_panel, (src_panel.width * 2 + 24, 0))
        name = f"issue_{idx:02d}_{issue.get('kind','issue')}.png"
        path = outdir / name
        sheet.save(path)
        issue_summary = {
            "index": idx,
            "kind": issue.get("kind"),
            "path": str(path),
            "crop_px": crop_box,
            "text": issue.get("text") or issue.get("a_text"),
            "raw_issue": issue,
        }
        index.append(issue_summary)

    (outdir / "issue_index.json").write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"outdir": str(outdir), "issue_crops": len(index)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
