#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageStat


def scaled_size(size: tuple[int, int], max_size: tuple[int, int]) -> tuple[int, int]:
    """Return an aspect-ratio-preserving size that fits within max_size."""
    width, height = size
    max_width, max_height = max_size
    scale = min(max_width / width, max_height / height)
    return max(1, round(width * scale)), max(1, round(height * scale))


def scale_and_center(im: Image.Image, canvas_size: tuple[int, int]) -> Image.Image:
    """Scale an image proportionally and center it on a common review canvas."""
    converted = im.convert("RGB")
    target_size = scaled_size(converted.size, canvas_size)
    scaled = converted.resize(target_size, Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", canvas_size, "white")
    x = (canvas.width - scaled.width) // 2
    y = (canvas.height - scaled.height) // 2
    canvas.paste(scaled, (x, y))
    return canvas


def main() -> int:
    parser = argparse.ArgumentParser(description="Create a source/render/diff review sheet for PPT reconstruction QA.")
    parser.add_argument("--source", required=True, help="Original source image.")
    parser.add_argument("--render", required=True, help="Rendered PPT slide PNG.")
    parser.add_argument("--out", required=True, help="Output review sheet PNG.")
    parser.add_argument("--width", type=int, default=1536, help="Maximum review-panel width.")
    parser.add_argument("--height", type=int, default=1024, help="Maximum review-panel height.")
    args = parser.parse_args()

    if args.width <= 0 or args.height <= 0:
        parser.error("--width and --height must be positive")

    with Image.open(args.source) as source_image, Image.open(args.render) as render_image:
        max_input_size = (
            max(source_image.width, render_image.width),
            max(source_image.height, render_image.height),
        )
        canvas_size = scaled_size(max_input_size, (args.width, args.height))
        src = scale_and_center(source_image, canvas_size)
        ren = scale_and_center(render_image, canvas_size)
    diff = ImageChops.difference(src, ren)
    stat = ImageStat.Stat(diff)
    mean = sum(stat.mean) / 3
    rms = (sum(v * v for v in stat.rms) / 3) ** 0.5
    diff_vis = diff.point(lambda p: min(255, p * 4))

    label_h = 44
    gutter = 16
    panel_width, panel_height = canvas_size
    sheet_w = panel_width * 3 + gutter * 4
    sheet_h = panel_height + label_h + gutter * 2
    sheet = Image.new("RGB", (sheet_w, sheet_h), "white")
    draw = ImageDraw.Draw(sheet)

    panels = [
        ("SOURCE", src),
        ("RENDER", ren),
        (f"DIFF x4 mean={mean:.2f} rms={rms:.2f}", diff_vis),
    ]
    x = gutter
    for label, im in panels:
        draw.rectangle([x - 1, gutter - 1, x + panel_width + 1, gutter + panel_height + 1], outline=(80, 80, 80), width=2)
        sheet.paste(im, (x, gutter))
        draw.text((x, gutter + panel_height + 10), label, fill=(0, 0, 0))
        x += panel_width + gutter

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(out)
    print({"out": str(out), "panel_size": canvas_size, "mean_abs_diff": round(mean, 3), "rms": round(rms, 3)})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
