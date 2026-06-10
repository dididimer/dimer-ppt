#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageStat


def fit(im: Image.Image, size: tuple[int, int]) -> Image.Image:
    return im.convert("RGB").resize(size, Image.Resampling.LANCZOS)


def main() -> int:
    parser = argparse.ArgumentParser(description="Create a source/render/diff review sheet for PPT reconstruction QA.")
    parser.add_argument("--source", required=True, help="Original source image.")
    parser.add_argument("--render", required=True, help="Rendered PPT slide PNG.")
    parser.add_argument("--out", required=True, help="Output review sheet PNG.")
    parser.add_argument("--width", type=int, default=1536)
    parser.add_argument("--height", type=int, default=1024)
    args = parser.parse_args()

    src = fit(Image.open(args.source), (args.width, args.height))
    ren = fit(Image.open(args.render), (args.width, args.height))
    diff = ImageChops.difference(src, ren)
    stat = ImageStat.Stat(diff)
    mean = sum(stat.mean) / 3
    rms = (sum(v * v for v in stat.rms) / 3) ** 0.5
    diff_vis = diff.point(lambda p: min(255, p * 4))

    label_h = 44
    gutter = 16
    sheet_w = args.width * 3 + gutter * 4
    sheet_h = args.height + label_h + gutter * 2
    sheet = Image.new("RGB", (sheet_w, sheet_h), "white")
    draw = ImageDraw.Draw(sheet)

    panels = [
        ("SOURCE", src),
        ("RENDER", ren),
        (f"DIFF x4 mean={mean:.2f} rms={rms:.2f}", diff_vis),
    ]
    x = gutter
    for label, im in panels:
        draw.rectangle([x - 1, gutter - 1, x + args.width + 1, gutter + args.height + 1], outline=(80, 80, 80), width=2)
        sheet.paste(im, (x, gutter))
        draw.text((x, gutter + args.height + 10), label, fill=(0, 0, 0))
        x += args.width + gutter

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(out)
    print({"out": str(out), "mean_abs_diff": round(mean, 3), "rms": round(rms, 3)})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
