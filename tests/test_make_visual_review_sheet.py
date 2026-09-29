from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "make_visual_review_sheet.py"
SPEC = importlib.util.spec_from_file_location("make_visual_review_sheet", SCRIPT)
assert SPEC and SPEC.loader
review_sheet = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(review_sheet)


class MakeVisualReviewSheetTests(unittest.TestCase):
    def test_non_three_by_two_input_keeps_its_aspect_ratio(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            source = temp / "source.png"
            render = temp / "render.png"
            output = temp / "review.png"
            Image.new("RGB", (1000, 500), (200, 10, 10)).save(source)
            Image.new("RGB", (800, 400), (10, 10, 200)).save(render)

            with patch.object(
                sys,
                "argv",
                [
                    "make_visual_review_sheet.py",
                    "--source",
                    str(source),
                    "--render",
                    str(render),
                    "--out",
                    str(output),
                ],
            ):
                self.assertEqual(review_sheet.main(), 0)

            with Image.open(output) as sheet:
                # A 2:1 input fits the default 1536-pixel maximum width at 1536x768.
                self.assertEqual(sheet.size, (4672, 844))
                self.assertEqual(sheet.getpixel((16, 16)), (200, 10, 10))


if __name__ == "__main__":
    unittest.main()
