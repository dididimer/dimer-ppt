#!/usr/bin/env python3
"""Create a small, ranked set of source/render review tiles.

The score is deliberately based on significant residual pixels after allowing a
one-pixel local match. This makes the script useful for triage: font
anti-aliasing and tiny raster offsets do not by themselves dominate the queue.
It is not a visual-approval gate.
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

from PIL import Image, ImageChops, ImageDraw, ImageFilter


DEFAULT_TOP_K = 3
RESIDUAL_TOLERANCE = 64
STRONG_RESIDUAL = 128
MIN_MATERIAL_PIXELS = 9
TILE_FILENAME = re.compile(r"tile_r\d+_c\d+\.png")


def annotate(im: Image.Image, title: str) -> Image.Image:
    h = 30
    out = Image.new("RGB", (im.width, im.height + h), "white")
    out.paste(im, (0, h))
    draw = ImageDraw.Draw(out)
    draw.text((6, 8), title, fill=(0, 0, 0))
    draw.rectangle([0, h, out.width - 1, out.height - 1], outline=(90, 90, 90), width=2)
    return out


def fit_to_canvas(im: Image.Image, max_width: int, max_height: int) -> tuple[Image.Image, dict[str, Any]]:
    """Resize without distortion and center the image in an aspect-safe canvas."""
    scale = min(max_width / im.width, max_height / im.height)
    resized_size = (max(1, round(im.width * scale)), max(1, round(im.height * scale)))
    resized = im.resize(resized_size, Image.Resampling.LANCZOS)
    canvas = Image.new("RGB", (max_width, max_height), "white")
    offset = ((max_width - resized.width) // 2, (max_height - resized.height) // 2)
    canvas.paste(resized, offset)
    return canvas, {
        "original_size": [im.width, im.height],
        "resized_size": list(resized_size),
        "offset_px": list(offset),
        "scale": scale,
    }


def resize_pair(source: Image.Image, render: Image.Image, max_width: int, max_height: int) -> tuple[Image.Image, Image.Image, dict[str, Any]]:
    """Fit both images to an aspect-safe canvas based on the source aspect ratio."""
    scale = min(max_width / source.width, max_height / source.height)
    canvas_size = (max(1, round(source.width * scale)), max(1, round(source.height * scale)))
    fitted_source, source_info = fit_to_canvas(source, *canvas_size)
    fitted_render, render_info = fit_to_canvas(render, *canvas_size)
    return fitted_source, fitted_render, {
        "canvas_size": list(canvas_size),
        "source": source_info,
        "render": render_info,
    }


def shifted(im: Image.Image, dx: int, dy: int) -> Image.Image:
    """Return a translated copy without ImageChops.offset's wraparound."""
    out = Image.new("RGB", im.size, "white")
    out.paste(im, (dx, dy))
    return out


def max_channel(im: Image.Image) -> Image.Image:
    red, green, blue = im.split()
    return ImageChops.lighter(ImageChops.lighter(red, green), blue)


def local_residual(source: Image.Image, render: Image.Image) -> Image.Image:
    """Find residual colour error after tolerating a one-pixel translation.

    Images are lightly blurred first so a single anti-aliased edge pixel is not
    treated as a meaningful regional change.
    """
    source_blur = source.filter(ImageFilter.GaussianBlur(radius=1))
    render_blur = render.filter(ImageFilter.GaussianBlur(radius=1))
    best = Image.new("L", source.size, 255)
    for dy in (-1, 0, 1):
        for dx in (-1, 0, 1):
            candidate = max_channel(ImageChops.difference(source_blur, shifted(render_blur, dx, dy)))
            best = ImageChops.darker(best, candidate)
    return best


def count_above(im: Image.Image, threshold: int) -> int:
    return sum(im.histogram()[threshold + 1 :])


def score_region(source: Image.Image, render: Image.Image) -> dict[str, Any]:
    """Return sortable regional metrics without using MAE as an approval rule."""
    residual = local_residual(source, render)
    pixels = source.width * source.height
    changed_pixels = count_above(residual, RESIDUAL_TOLERANCE)
    strong_pixels = count_above(residual, STRONG_RESIDUAL)
    changed_coverage = changed_pixels / pixels
    strong_coverage = strong_pixels / pixels
    # Coverage rewards broad layout/content errors; strong coverage separates
    # obvious changes from a faint residual with equal area.
    score = changed_coverage + strong_coverage
    material = changed_pixels >= MIN_MATERIAL_PIXELS
    return {
        "score": round(score, 8),
        "changed_pixels": changed_pixels,
        "strong_pixels": strong_pixels,
        "changed_coverage": round(changed_coverage, 8),
        "strong_coverage": round(strong_coverage, 8),
        "review_recommended": material,
    }


def tile_box(row: int, col: int, cols: int, rows: int, width: int, height: int) -> tuple[int, int, int, int]:
    tile_w = width // cols
    tile_h = height // rows
    return (
        col * tile_w,
        row * tile_h,
        width if col == cols - 1 else (col + 1) * tile_w,
        height if row == rows - 1 else (row + 1) * tile_h,
    )


def save_tile(source: Image.Image, render: Image.Image, diff: Image.Image, box: tuple[int, int, int, int], path: Path, row: int, col: int) -> None:
    src = annotate(source.crop(box), f"SOURCE r{row}c{col}")
    ren = annotate(render.crop(box), f"RENDER r{row}c{col}")
    dif = annotate(diff.crop(box), "DIFF x4")
    sheet = Image.new("RGB", (src.width * 3 + 24, src.height), "white")
    sheet.paste(src, (0, 0))
    sheet.paste(ren, (src.width + 12, 0))
    sheet.paste(dif, (src.width * 2 + 24, 0))
    sheet.save(path)


def remove_previous_tiles(outdir: Path) -> None:
    """Remove only review tiles this script may have created in a prior run."""
    for path in outdir.iterdir():
        if path.is_file() and TILE_FILENAME.fullmatch(path.name):
            path.unlink()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Ranked source/render regional review tiles for semantic layout QA.")
    parser.add_argument("--source", required=True)
    parser.add_argument("--render", required=True)
    parser.add_argument("--outdir", required=True)
    parser.add_argument("--cols", type=int, default=4)
    parser.add_argument("--rows", type=int, default=3)
    parser.add_argument("--width", type=int, default=1536, help="Maximum review canvas width; source aspect ratio is preserved.")
    parser.add_argument("--height", type=int, default=1024, help="Maximum review canvas height; source aspect ratio is preserved.")
    parser.add_argument("--top-k", type=int, default=DEFAULT_TOP_K, help="Number of material-difference regions to write (default: 3).")
    parser.add_argument("--all", action="store_true", help="Write every region, including regions with no material difference.")
    args = parser.parse_args()
    if args.cols < 1 or args.rows < 1 or args.width < 1 or args.height < 1:
        parser.error("--cols, --rows, --width, and --height must be positive")
    if args.top_k < 0:
        parser.error("--top-k must be non-negative")
    return args


def main() -> int:
    args = parse_args()
    original_source = Image.open(args.source).convert("RGB")
    original_render = Image.open(args.render).convert("RGB")
    source, render, geometry = resize_pair(original_source, original_render, args.width, args.height)
    if args.cols > source.width or args.rows > source.height:
        raise SystemExit(
            "--cols and --rows must not exceed the scaled review canvas "
            f"({source.width}x{source.height})"
        )
    diff = ImageChops.difference(source, render).point(lambda p: min(255, p * 4))
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    ranked: list[dict[str, Any]] = []
    for row in range(args.rows):
        for col in range(args.cols):
            box = tile_box(row, col, args.cols, args.rows, source.width, source.height)
            metrics = score_region(source.crop(box), render.crop(box))
            ranked.append({"row": row + 1, "col": col + 1, "bbox_px": list(box), **metrics})

    ranked.sort(key=lambda item: (-item["score"], -item["changed_pixels"], item["row"], item["col"]))
    for rank, item in enumerate(ranked, start=1):
        item["rank"] = rank

    recommended = [item for item in ranked if item["review_recommended"]]
    selected = ranked if args.all else recommended[: args.top_k]
    remove_previous_tiles(outdir)
    tile_index: list[dict[str, Any]] = []
    selected_keys = {(item["row"], item["col"]) for item in selected}
    for item in ranked:
        if (item["row"], item["col"]) not in selected_keys:
            continue
        name = f"tile_r{item['row']}_c{item['col']}.png"
        path = outdir / name
        save_tile(source, render, diff, tuple(item["bbox_px"]), path, item["row"], item["col"])
        tile_index.append({**item, "path": str(path)})

    # Retain the original file and list shape for existing consumers; it now
    # lists the tiles actually emitted. The separate summary is exhaustive.
    (outdir / "tile_index.json").write_text(json.dumps(tile_index, ensure_ascii=False, indent=2), encoding="utf-8")
    summary = {
        "schema_version": 2,
        "source": str(args.source),
        "render": str(args.render),
        "geometry": geometry,
        "config": {
            "cols": args.cols,
            "rows": args.rows,
            "max_width": args.width,
            "max_height": args.height,
            "top_k": args.top_k,
            "all": args.all,
            "residual_tolerance": RESIDUAL_TOLERANCE,
            "small_shift_tolerance_px": 1,
            "minimum_material_pixels": MIN_MATERIAL_PIXELS,
        },
        "has_material_differences": bool(recommended),
        "recommended_regions": len(recommended),
        "selected_regions": len(tile_index),
        "ranked_regions": ranked,
    }
    summary_path = outdir / "review_summary.json"
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"outdir": str(outdir), "tiles": len(tile_index), "total_regions": len(ranked), "has_material_differences": bool(recommended), "summary": str(summary_path)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
