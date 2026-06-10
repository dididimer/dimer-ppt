#!/usr/bin/env python3
from __future__ import annotations

import json
import re
import sys
import zipfile
from pathlib import Path


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: pptx_xml_stats.py <deck.pptx>", file=sys.stderr)
        return 2
    path = Path(sys.argv[1])
    if not path.exists():
        print(json.dumps({"ok": False, "error": f"not found: {path}"}, ensure_ascii=False))
        return 2
    with zipfile.ZipFile(path) as zf:
        slides = sorted(n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n))
        media = [n for n in zf.namelist() if n.startswith("ppt/media/")]
        result = {
            "ok": True,
            "file": str(path),
            "slide_count": len(slides),
            "media_count": len(media),
            "media": media,
            "slides": [],
        }
        for slide in slides:
            xml = zf.read(slide).decode("utf-8", errors="ignore")
            result["slides"].append({
                "slide": slide,
                "pictures": xml.count("<p:pic>"),
                "shapes": xml.count("<p:sp>"),
                "groups": xml.count("<p:grpSp>"),
                "text_nodes": xml.count("<a:t"),
            })
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

