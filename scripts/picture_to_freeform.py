#!/usr/bin/env python3
"""Opt-in conversion of simple PPTX picture objects into native freeforms.

The converter deliberately accepts only flat, low-colour artwork. It leaves an
image untouched whenever its pixels or geometry cannot be represented within
the requested budgets. This makes it useful for small icon screenshots, while
avoiding an automatic low-fidelity conversion of photos and illustrations.
"""

from __future__ import annotations

import argparse
import io
import json
import sys
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import numpy as np
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from scipy import ndimage


EXIT_USAGE = 2


@dataclass(frozen=True)
class Contour:
    """One closed pixel-edge contour in clockwise/counter-clockwise order."""

    points: tuple[tuple[int, int], ...]

    @property
    def signed_area(self) -> float:
        return sum(
            x1 * y2 - x2 * y1
            for (x1, y1), (x2, y2) in zip(self.points, self.points[1:] + self.points[:1])
        ) / 2.0


class ConversionSkipped(Exception):
    """An expected, reportable reason to leave the source picture untouched."""


def parse_csv(values: Iterable[str] | None) -> set[str]:
    result: set[str] = set()
    for value in values or []:
        result.update(part.strip() for part in value.split(",") if part.strip())
    return result


def parse_indices(values: Iterable[str] | None) -> set[int]:
    result: set[int] = set()
    for token in parse_csv(values):
        try:
            index = int(token)
        except ValueError as exc:
            raise argparse.ArgumentTypeError(f"picture index is not an integer: {token!r}") from exc
        if index < 1:
            raise argparse.ArgumentTypeError("picture indices are 1-based positive integers")
        result.add(index)
    return result


def crop_visible_image(shape) -> Image.Image:
    """Return the source image after the same rectangular crop as *shape*."""

    with Image.open(io.BytesIO(shape.image.blob)) as opened:
        image = opened.convert("RGBA")

    crop = (shape.crop_left, shape.crop_top, shape.crop_right, shape.crop_bottom)
    if any(value < 0 for value in crop) or crop[0] + crop[2] >= 1 or crop[1] + crop[3] >= 1:
        raise ConversionSkipped("unsupported_crop_values")

    width, height = image.size
    left = int(round(width * crop[0]))
    top = int(round(height * crop[1]))
    right = width - int(round(width * crop[2]))
    bottom = height - int(round(height * crop[3]))
    if right - left < 2 or bottom - top < 2:
        raise ConversionSkipped("crop_has_no_usable_pixels")
    return image.crop((left, top, right, bottom))


def has_unsupported_picture_transform(shape) -> bool:
    """Reject flip/effect transforms rather than guessing at their rendering."""

    xfrm = shape._element.spPr.xfrm
    if xfrm is not None and (xfrm.get("flipH") or xfrm.get("flipV")):
        return True
    blip = shape._element.blipFill.blip
    # blip effects alter the source pixels but are not exposed consistently by python-pptx.
    return any(child.tag.rsplit("}", 1)[-1] != "extLst" for child in blip)


def color_layers(image: Image.Image, alpha_threshold: int, max_colors: int) -> list[tuple[tuple[int, int, int], np.ndarray]]:
    """Return exact flat RGB layers, or skip artwork that needs quantization."""

    pixels = np.asarray(image, dtype=np.uint8)
    rgb = pixels[:, :, :3]
    alpha = pixels[:, :, 3]
    visible = alpha >= alpha_threshold
    if not np.any(visible):
        raise ConversionSkipped("no_pixels_above_alpha_threshold")

    # A flat PNG usually has a fully opaque (or near-opaque) core. Sampling
    # only that core avoids treating anti-aliased transparent edge pixels as a
    # new colour. We still require every visible pixel to use one of the core
    # RGB values, so gradients and baked-background anti-aliasing are skipped.
    max_alpha = int(alpha.max())
    core = alpha >= max(alpha_threshold, max_alpha - 8)
    colors, counts = np.unique(rgb[core].reshape(-1, 3), axis=0, return_counts=True)
    if len(colors) == 0:
        raise ConversionSkipped("no_flat_color_core")
    if len(colors) > max_colors:
        raise ConversionSkipped(f"color_budget_exceeded:{len(colors)}>{max_colors}")

    palette = [tuple(int(channel) for channel in color) for color in colors]
    matches = np.zeros(visible.shape, dtype=bool)
    layers: list[tuple[tuple[int, int, int], np.ndarray]] = []
    for color in palette:
        mask = visible & np.all(rgb == color, axis=2)
        matches |= mask
        if np.any(mask):
            layers.append((color, mask))
    unknown = int(np.count_nonzero(visible & ~matches))
    if unknown:
        total = int(np.count_nonzero(visible))
        raise ConversionSkipped(f"non_flat_visible_pixels:{unknown}/{total}")

    # Preserve painter order deterministically: colours with more opaque pixels
    # are drawn first. Their masks are disjoint, so this cannot hide artwork.
    return sorted(layers, key=lambda item: (-int(np.count_nonzero(item[1])), item[0]))


def trace_contours(mask: np.ndarray) -> list[Contour]:
    """Trace the directed exposed pixel edges of a binary mask.

    A vertex with multiple outgoing edges denotes diagonal-only contact. It is
    intentionally rejected because choosing a connection there changes holes
    and can turn distinct icon details into an incorrect shape.
    """

    height, width = mask.shape
    outgoing: dict[tuple[int, int], list[tuple[int, int]]] = {}

    def add(start: tuple[int, int], end: tuple[int, int]) -> None:
        outgoing.setdefault(start, []).append(end)

    for y, x in np.argwhere(mask):
        if y == 0 or not mask[y - 1, x]:
            add((int(x), int(y)), (int(x + 1), int(y)))
        if x == width - 1 or not mask[y, x + 1]:
            add((int(x + 1), int(y)), (int(x + 1), int(y + 1)))
        if y == height - 1 or not mask[y + 1, x]:
            add((int(x + 1), int(y + 1)), (int(x), int(y + 1)))
        if x == 0 or not mask[y, x - 1]:
            add((int(x), int(y + 1)), (int(x), int(y)))

    if any(len(ends) != 1 for ends in outgoing.values()):
        raise ConversionSkipped("ambiguous_diagonal_contour_topology")

    contours: list[Contour] = []
    while outgoing:
        start = next(iter(outgoing))
        current = start
        points = [start]
        while True:
            end = outgoing.pop(current)[0]
            current = end
            if current == start:
                break
            points.append(current)
        contours.append(Contour(tuple(remove_collinear(points))))
    return contours


def remove_collinear(points: list[tuple[int, int]]) -> list[tuple[int, int]]:
    result: list[tuple[int, int]] = []
    for index, point in enumerate(points):
        previous = points[index - 1]
        following = points[(index + 1) % len(points)]
        if (point[0] - previous[0]) * (following[1] - point[1]) != (
            point[1] - previous[1]
        ) * (following[0] - point[0]):
            result.append(point)
    if len(result) < 3:
        raise ConversionSkipped("degenerate_contour")
    return result


def point_segment_distance(point: np.ndarray, start: np.ndarray, end: np.ndarray) -> float:
    segment = end - start
    length_sq = float(np.dot(segment, segment))
    if length_sq == 0:
        return float(np.linalg.norm(point - start))
    projection = float(np.dot(point - start, segment)) / length_sq
    nearest = start + min(1.0, max(0.0, projection)) * segment
    return float(np.linalg.norm(point - nearest))


def rdp_open(points: list[tuple[int, int]], epsilon: float) -> list[tuple[int, int]]:
    if len(points) <= 2:
        return points
    start = np.asarray(points[0], dtype=float)
    end = np.asarray(points[-1], dtype=float)
    distances = [point_segment_distance(np.asarray(point, dtype=float), start, end) for point in points[1:-1]]
    if not distances or max(distances) <= epsilon:
        return [points[0], points[-1]]
    pivot = 1 + int(np.argmax(distances))
    return rdp_open(points[: pivot + 1], epsilon)[:-1] + rdp_open(points[pivot:], epsilon)


def simplify_closed(points: tuple[tuple[int, int], ...], epsilon: float) -> tuple[tuple[int, int], ...]:
    """Simplify a ring by splitting it at its farthest pair of points."""

    if len(points) <= 3:
        return points
    first = np.asarray(points[0], dtype=float)
    pivot = max(range(1, len(points)), key=lambda index: float(np.linalg.norm(np.asarray(points[index]) - first)))
    first_half = rdp_open(list(points[: pivot + 1]), epsilon)
    second_half = rdp_open(list(points[pivot:]) + [points[0]], epsilon)
    simplified = first_half[:-1] + second_half[:-1]
    if len(simplified) < 3:
        raise ConversionSkipped("simplification_collapsed_contour")
    return tuple(simplified)


def contains_point(polygon: tuple[tuple[int, int], ...], point: tuple[int, int]) -> bool:
    """Even-odd containment test for associating a hole with one outer ring."""

    x, y = point
    inside = False
    for (x1, y1), (x2, y2) in zip(polygon, polygon[1:] + polygon[:1]):
        if (y1 > y) != (y2 > y) and x < (x2 - x1) * (y - y1) / (y2 - y1) + x1:
            inside = not inside
    return inside


def plan_freeforms(
    image: Image.Image,
    alpha_threshold: int,
    max_colors: int,
    epsilon: float,
    max_shapes: int,
) -> list[tuple[tuple[int, int, int], tuple[tuple[tuple[int, int], ...], ...]]]:
    """Create colour-layer paths: one outer contour plus any contained holes."""

    plans = []
    component_lower_bound = 0
    for color, mask in color_layers(image, alpha_threshold, max_colors):
        # Count 4-connected components for reporting and sanity checking. The
        # actual contour tracer preserves holes as paths within each component.
        _, components = ndimage.label(mask, structure=np.array([[0, 1, 0], [1, 1, 1], [0, 1, 0]]))
        component_lower_bound += int(components)
        if component_lower_bound > max_shapes:
            raise ConversionSkipped(f"shape_budget_exceeded:{component_lower_bound}>{max_shapes}")
        contours = [Contour(simplify_closed(contour.points, epsilon)) for contour in trace_contours(mask)]
        outers = [contour for contour in contours if contour.signed_area > 0]
        holes = [contour for contour in contours if contour.signed_area < 0]
        if not outers:
            raise ConversionSkipped("no_outer_contours")
        for outer in outers:
            contained_holes = tuple(hole.points for hole in holes if contains_point(outer.points, hole.points[0]))
            plans.append((color, (outer.points,) + contained_holes))
    return plans


def picture_record(global_index: int, slide_index: int, shape) -> dict:
    return {
        "picture_index": global_index,
        "slide_index": slide_index,
        "picture_name": shape.name,
        "source_shape_type": str(shape.shape_type),
        "source_bounds_emu": {
            "left": int(shape.left),
            "top": int(shape.top),
            "width": int(shape.width),
            "height": int(shape.height),
        },
        "status": "skipped",
    }


def selected(global_index: int, shape, indices: set[int], names: set[str], all_pictures: bool) -> bool:
    return all_pictures or global_index in indices or shape.name in names


def restore_picture_coordinate_box(native, source_shape, image: Image.Image, paths) -> None:
    """Keep the source picture's bounding box, including transparent margins.

    FreeformBuilder normally tightens `a:path` and `a:xfrm` to the opaque
    points. A picture can have meaningful transparent padding, so restore the
    cropped image's full local coordinate box and translate its path points
    back into that box before applying the original object transform.
    """

    min_x = min(x for path in paths for x, _ in path)
    min_y = min(y for path in paths for _, y in path)
    custom_path = list(native._element.spPr.custGeom.pathLst)[0]
    custom_path.set("w", str(image.width))
    custom_path.set("h", str(image.height))
    for element in custom_path.iter():
        if element.tag.rsplit("}", 1)[-1] == "pt":
            element.set("x", str(int(element.get("x")) + min_x))
            element.set("y", str(int(element.get("y")) + min_y))
    native.left = source_shape.left
    native.top = source_shape.top
    native.width = source_shape.width
    native.height = source_shape.height
    native.rotation = source_shape.rotation


def convert_picture(slide, shape, global_index: int, args) -> dict:
    image = crop_visible_image(shape)
    plans = plan_freeforms(
        image,
        args.alpha_threshold,
        args.max_colors,
        args.simplify_epsilon,
        args.max_shapes,
    )
    shape_count = len(plans)
    node_count = sum(len(path) for _, paths in plans for path in paths)
    if shape_count > args.max_shapes:
        raise ConversionSkipped(f"shape_budget_exceeded:{shape_count}>{args.max_shapes}")
    if node_count > args.max_nodes:
        raise ConversionSkipped(f"node_budget_exceeded:{node_count}>{args.max_nodes}")

    parent = shape._element.getparent()
    insertion_index = parent.index(shape._element)
    created = []
    x_scale = shape.width / image.width
    y_scale = shape.height / image.height
    try:
        for plan_index, (color, paths) in enumerate(plans, start=1):
            outer, *holes = paths
            builder = slide.shapes.build_freeform(*outer[0], scale=(x_scale, y_scale))
            builder.add_line_segments(outer[1:], close=True)
            for hole in holes:
                builder.move_to(*hole[0]).add_line_segments(hole[1:], close=True)
            native = builder.convert_to_shape(shape.left, shape.top)
            restore_picture_coordinate_box(native, shape, image, paths)
            native.name = f"DIMER_ICON_{global_index}_C{plan_index}"
            native.fill.solid()
            native.fill.fore_color.rgb = RGBColor(*color)
            native.line.fill.background()
            # python-pptx's default freeform style references a theme effect
            # (often a drop shadow); the source bitmap has no such effect.
            for effect_ref in native._element.xpath("./p:style/a:effectRef"):
                effect_ref.set("idx", "0")
            created.append(native)

        # build_freeform appends objects. Move each native shape into the exact
        # slot formerly occupied by the picture to retain z-order.
        for offset, native in enumerate(created):
            parent.remove(native._element)
            parent.insert(insertion_index + offset, native._element)
        parent.remove(shape._element)
    except Exception:
        for native in created:
            if native._element.getparent() is not None:
                native._element.getparent().remove(native._element)
        raise

    return {
        "status": "converted",
        "cropped_source_size_px": {"width": image.width, "height": image.height},
        "native_shape_count": shape_count,
        "native_node_count": node_count,
        "native_shape_names": [native.name for native in created],
        "colors": ["#%02X%02X%02X" % color for color, _ in plans],
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_pptx", type=Path)
    parser.add_argument("output_pptx", type=Path)
    selection = parser.add_argument_group("explicit picture selection")
    selection.add_argument("--picture-indices", action="append", metavar="N[,N...]", help="1-based indices across all slides")
    selection.add_argument("--picture-names", action="append", metavar="NAME[,NAME...]", help="exact PowerPoint picture names")
    selection.add_argument("--all-pictures", action="store_true", help="explicitly process every top-level picture")
    parser.add_argument("--max-colors", type=int, default=4, help="maximum exact flat RGB layers per image (default: 4)")
    parser.add_argument("--max-shapes", type=int, default=24, help="maximum native freeforms replacing one picture (default: 24)")
    parser.add_argument("--max-nodes", type=int, default=1200, help="maximum path vertices replacing one picture (default: 1200)")
    parser.add_argument("--alpha-threshold", type=int, default=16, help="minimum alpha treated as visible (default: 16)")
    parser.add_argument("--simplify-epsilon", type=float, default=0.8, help="contour simplification tolerance in source pixels (default: 0.8)")
    parser.add_argument("--report", type=Path, help="JSON report path (default: OUTPUT.pptx.json)")
    return parser


def validate_args(parser: argparse.ArgumentParser, args) -> tuple[set[int], set[str]]:
    try:
        indices = parse_indices(args.picture_indices)
    except argparse.ArgumentTypeError as exc:
        parser.error(str(exc))
    names = parse_csv(args.picture_names)
    if not (args.all_pictures or indices or names):
        parser.error("choose --all-pictures, --picture-indices, or --picture-names; conversion is opt-in")
    if args.max_colors < 1 or args.max_shapes < 1 or args.max_nodes < 3:
        parser.error("all budgets must be positive, and --max-nodes must be at least 3")
    if not 1 <= args.alpha_threshold <= 255:
        parser.error("--alpha-threshold must be between 1 and 255")
    if args.simplify_epsilon < 0:
        parser.error("--simplify-epsilon must be non-negative")
    if not args.input_pptx.is_file():
        parser.error(f"input PPTX not found: {args.input_pptx}")
    if args.input_pptx.resolve() == args.output_pptx.resolve():
        parser.error("output PPTX must differ from input PPTX")
    return indices, names


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    indices, names = validate_args(parser, args)
    report_path = args.report or args.output_pptx.with_suffix(args.output_pptx.suffix + ".json")
    presentation = Presentation(args.input_pptx)
    report = {
        "ok": True,
        "input_pptx": str(args.input_pptx),
        "output_pptx": str(args.output_pptx),
        "budgets": {
            "max_colors": args.max_colors,
            "max_shapes": args.max_shapes,
            "max_nodes": args.max_nodes,
            "alpha_threshold": args.alpha_threshold,
            "simplify_epsilon": args.simplify_epsilon,
        },
        "pictures": [],
        "summary": Counter(),
    }

    global_index = 0
    for slide_index, slide in enumerate(presentation.slides, start=1):
        # Snapshot avoids iterating over the native shapes appended during a conversion.
        for shape in list(slide.shapes):
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue
            global_index += 1
            record = picture_record(global_index, slide_index, shape)
            if not selected(global_index, shape, indices, names, args.all_pictures):
                record["reason"] = "not_selected"
            elif has_unsupported_picture_transform(shape):
                record["reason"] = "unsupported_picture_transform"
            else:
                try:
                    record.update(convert_picture(slide, shape, global_index, args))
                except ConversionSkipped as exc:
                    record["reason"] = str(exc)
                except Exception as exc:  # Preserve the source object and make failures visible.
                    record["reason"] = f"unexpected_error:{type(exc).__name__}"
                    record["detail"] = str(exc)
            report["pictures"].append(record)
            report["summary"][record["status"]] += 1
            if record["status"] == "converted":
                report["summary"]["native_shape_count"] += record["native_shape_count"]
                report["summary"]["native_node_count"] += record["native_node_count"]

    report["summary"] = dict(report["summary"])
    report["summary"]["pictures_before"] = len(report["pictures"])
    report["summary"]["pictures_after"] = sum(
        shape.shape_type == MSO_SHAPE_TYPE.PICTURE
        for slide in presentation.slides
        for shape in slide.shapes
    )
    args.output_pptx.parent.mkdir(parents=True, exist_ok=True)
    presentation.save(args.output_pptx)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
