# Layout Safety Rules

## Text Box Sizing

- Give CJK and mixed CJK/Latin text 15-25% horizontal slack.
- Give dense labels 20-35% slack because WPS and PowerPoint often render wider
  than browser/SVG text.
- Avoid negative letter spacing. Avoid viewport-scaled fonts.
- Prefer manual line breaks for bilingual labels that are longer than the local
  card width.
- Short icon symbols may be text, but name them as decorative and keep them in
  small boxes.

## Collision Fix Priority

1. Move text to its own clear zone.
2. Widen the safe text box inside the same panel.
3. Reduce font size by 0.5 pt steps.
4. Split into two lines.
5. Simplify nearby icon geometry.

Do not solve collisions by sending text behind icons or by using a full-slide
image background.

## Chinese Font Policy

Windows: prefer Microsoft YaHei for UI diagrams. For very dense small labels,
SimSun can fit more characters but looks less modern.

WPS may substitute fonts differently from Microsoft PowerPoint, so always render
and inspect with the app the user is likely to use.

