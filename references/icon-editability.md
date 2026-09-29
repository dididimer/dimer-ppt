# Editable icon strategy

Apply this when a reconstructed PPT contains small screenshot or PNG icons. Count picture objects on the slide, not only files in `ppt/media`. Keep original source images available for visual comparison.

1. **Semantic icon:** If its meaning can be expressed with a small family of shapes (document, node graph, robot, shield, trophy, warning), build or reuse native PowerPoint shapes. `scripts/native_icon_components.py` provides an explicit picture-name-to-component map for common examples. Reuse one component at every repeated occurrence. The result is deliberately editable and may simplify decorative detail.
2. **Flat artwork:** For an icon with truly few exact colors, try `scripts/picture_to_freeform.py` on explicitly selected pictures. The script preserves the original object if color, shape, or node limits are exceeded. Do not raise those limits merely to force a noisy screenshot through the converter.
3. **Complex artwork:** Keep photos, textures, gradients, antialiased screenshots, and detailed people as small raster crops in `hybrid-fidelity`. Record each crop and its reason. If the user requires every element editable, rebuild such artwork manually and disclose any visual simplification.

An SVG inserted as one picture is still one picture object. If an original SVG is available, PowerPoint's **Convert to Shape** can disassemble it into editable pieces; verify the resulting slide after conversion.

## Verify a replacement

- Generate a PPTX from the same source and render both versions through PowerPoint/WPS at identical dimensions. `scripts/render_pptx_windows.ps1` performs the PowerPoint export on Windows.
- Check picture count, native shape count, and selection of individual icon pieces. Ensure the replacement did not change neighboring text, connectors, or panel geometry.
- Inspect changed regions rather than accepting a slide-wide pixel score. A simplified editable icon can differ from the screenshot while still being visually appropriate; report that tradeoff.
- For a reproducible example, see `examples/native-icons/`: the same slide changes from 23 picture objects to 3, with before/after PPTX files, PowerPoint renders, a mapping, and a report.
