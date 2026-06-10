#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw


def annotate(im: Image.Image, title: str) -> Image.Image:
    h = 30
    out = Image.new("RGB", (im.width, im.height + h), "white")
    out.paste(im, (0, h))
    draw = ImageDraw.Draw(out)
    draw.text((6, 8), title, fill=(0, 0, 0))
    draw.rectangle([0, h, out.width - 1, out.height - 1], outline=(90, 90, 90), width=2)
    return out


def main() -> int:
    parser = argparse.ArgumentParser(description="Split source/render/diff into review tiles for semantic layout QA.")
    parser.add_argument("--source", required=True)
    parser.add_argument("--render", required=True)
    parser.add_argument("--outdir", required=True)
    parser.add_argument("--cols", type=int, default=4)
    parser.add_argument("--rows", type=int, default=3)
    parser.add_argument("--width", type=int, default=1536)
    parser.add_argument("--height", type=int, default=1024)
    args = parser.parse_args()

    source = Image.open(args.source).convert("RGB").resize((args.width, args.height), Image.Resampling.LANCZOS)
    render = Image.open(args.render).convert("RGB").resize((args.width, args.height), Image.Resampling.LANCZOS)
    diff = ImageChops.difference(source, render).point(lambda p: min(255, p * 4))
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    tile_w = args.width // args.cols
    tile_h = args.height // args.rows
    index = []
    for row in range(args.rows):
        for col in range(args.cols):
            x1 = col * tile_w
            y1 = row * tile_h
            x2 = args.width if col == args.cols - 1 else (col + 1) * tile_w
            y2 = args.height if row == args.rows - 1 else (row + 1) * tile_h
            box = (x1, y1, x2, y2)
            src = annotate(source.crop(box), f"SOURCE r{row+1}c{col+1}")
            ren = annotate(render.crop(box), f"RENDER r{row+1}c{col+1}")
            dif = annotate(diff.crop(box), "DIFF x4")
            sheet = Image.new("RGB", (src.width * 3 + 24, src.height), "white")
            sheet.paste(src, (0, 0))
            sheet.paste(ren, (src.width + 12, 0))
            sheet.paste(dif, (src.width * 2 + 24, 0))
            name = f"tile_r{row+1}_c{col+1}.png"
            path = outdir / name
            sheet.save(path)
            index.append({"row": row + 1, "col": col + 1, "bbox_px": box, "path": str(path)})

    (outdir / "tile_index.json").write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"outdir": str(outdir), "tiles": len(index)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
