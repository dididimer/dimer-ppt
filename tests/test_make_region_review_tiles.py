import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image, ImageDraw


REPO = Path(__file__).resolve().parents[1]
SCRIPT = REPO / "scripts" / "make_region_review_tiles.py"


def save_pair(directory: Path, size=(120, 80), changes=()):
    source = Image.new("RGB", size, "white")
    render = source.copy()
    draw = ImageDraw.Draw(render)
    for box in changes:
        draw.rectangle(box, fill="black")
    source_path = directory / "source.png"
    render_path = directory / "render.png"
    source.save(source_path)
    render.save(render_path)
    return source_path, render_path


def run_tiles(directory: Path, source: Path, render: Path, *extra: str):
    outdir = directory / "tiles"
    command = [
        sys.executable,
        str(SCRIPT),
        "--source",
        str(source),
        "--render",
        str(render),
        "--outdir",
        str(outdir),
        "--cols",
        "2",
        "--rows",
        "2",
        "--width",
        "120",
        "--height",
        "80",
        *extra,
    ]
    completed = subprocess.run(command, check=True, capture_output=True, text=True)
    return outdir, json.loads(completed.stdout), json.loads((outdir / "review_summary.json").read_text(encoding="utf-8"))


class RegionReviewTilesTests(unittest.TestCase):
    def test_ranks_regions_and_limits_output_to_top_k(self):
        with tempfile.TemporaryDirectory() as temp:
            directory = Path(temp)
            source, render = save_pair(
                directory,
                changes=((5, 5, 34, 24), (65, 5, 74, 14), (5, 48, 24, 62), (65, 48, 72, 55)),
            )
            outdir, stdout, summary = run_tiles(directory, source, render, "--top-k", "2")

            self.assertEqual([(item["row"], item["col"]) for item in summary["ranked_regions"][:3]], [(1, 1), (2, 1), (1, 2)])
            index = json.loads((outdir / "tile_index.json").read_text(encoding="utf-8"))
            self.assertEqual([(item["row"], item["col"]) for item in index], [(1, 1), (2, 1)])
            self.assertEqual(stdout["tiles"], 2)
            self.assertEqual(summary["selected_regions"], 2)

    def test_default_top_k_is_three(self):
        with tempfile.TemporaryDirectory() as temp:
            directory = Path(temp)
            source, render = save_pair(
                directory,
                changes=((5, 5, 34, 24), (65, 5, 84, 19), (5, 48, 21, 62), (65, 48, 78, 59)),
            )
            outdir, _, summary = run_tiles(directory, source, render)

            index = json.loads((outdir / "tile_index.json").read_text(encoding="utf-8"))
            self.assertEqual(len(index), 3)
            self.assertEqual(summary["selected_regions"], 3)

    def test_no_difference_emits_no_default_tiles_but_all_emits_every_region(self):
        with tempfile.TemporaryDirectory() as temp:
            directory = Path(temp)
            source, render = save_pair(directory)
            outdir, stdout, summary = run_tiles(directory, source, render)

            self.assertFalse(summary["has_material_differences"])
            self.assertEqual(summary["selected_regions"], 0)
            self.assertEqual(stdout["tiles"], 0)
            self.assertEqual(json.loads((outdir / "tile_index.json").read_text(encoding="utf-8")), [])

            all_dir = directory / "all"
            completed = subprocess.run(
                [
                    sys.executable,
                    str(SCRIPT),
                    "--source", str(source), "--render", str(render), "--outdir", str(all_dir),
                    "--cols", "2", "--rows", "2", "--width", "120", "--height", "80", "--all",
                ],
                check=True,
                capture_output=True,
                text=True,
            )
            self.assertEqual(json.loads(completed.stdout)["tiles"], 4)
            self.assertEqual(len(json.loads((all_dir / "tile_index.json").read_text(encoding="utf-8"))), 4)

    def test_second_run_removes_only_stale_review_tiles(self):
        with tempfile.TemporaryDirectory() as temp:
            directory = Path(temp)
            source, render = save_pair(
                directory,
                changes=((5, 5, 34, 24), (65, 5, 84, 19), (5, 48, 21, 62), (65, 48, 78, 59)),
            )
            outdir, _, first_summary = run_tiles(directory, source, render, "--all")
            self.assertEqual(first_summary["selected_regions"], 4)
            self.assertEqual(len(list(outdir.glob("tile_r*_c*.png"))), 4)

            (outdir / "keep_me.png").write_bytes(b"unrelated")
            _, _, second_summary = run_tiles(directory, source, render)
            self.assertEqual(second_summary["selected_regions"], 3)
            self.assertEqual(len(list(outdir.glob("tile_r*_c*.png"))), 3)
            self.assertTrue((outdir / "keep_me.png").exists())

    def test_rejects_grid_larger_than_scaled_canvas(self):
        with tempfile.TemporaryDirectory() as temp:
            directory = Path(temp)
            source, render = save_pair(directory)
            completed = subprocess.run(
                [
                    sys.executable,
                    str(SCRIPT),
                    "--source", str(source), "--render", str(render), "--outdir", str(directory / "tiles"),
                    "--cols", "121", "--rows", "1", "--width", "120", "--height", "80",
                ],
                capture_output=True,
                text=True,
            )
            self.assertNotEqual(completed.returncode, 0)
            self.assertIn("must not exceed the scaled review canvas", completed.stderr)

    def test_preserves_source_aspect_ratio_and_ignores_one_pixel_shift(self):
        with tempfile.TemporaryDirectory() as temp:
            directory = Path(temp)
            source = Image.new("RGB", (200, 100), "white")
            ImageDraw.Draw(source).rectangle((40, 30, 120, 70), fill="black")
            render = Image.new("RGB", (200, 100), "white")
            ImageDraw.Draw(render).rectangle((41, 30, 121, 70), fill="black")
            source_path = directory / "source.png"
            render_path = directory / "render.png"
            source.save(source_path)
            render.save(render_path)

            outdir = directory / "tiles"
            completed = subprocess.run(
                [
                    sys.executable,
                    str(SCRIPT),
                    "--source", str(source_path), "--render", str(render_path), "--outdir", str(outdir),
                    "--cols", "2", "--rows", "1", "--width", "100", "--height", "100",
                ],
                check=True,
                capture_output=True,
                text=True,
            )
            summary = json.loads((outdir / "review_summary.json").read_text(encoding="utf-8"))
            self.assertEqual(summary["geometry"]["canvas_size"], [100, 50])
            self.assertEqual(summary["geometry"]["source"]["resized_size"], [100, 50])
            self.assertFalse(summary["has_material_differences"])
            self.assertEqual(json.loads(completed.stdout)["tiles"], 0)


if __name__ == "__main__":
    unittest.main()
