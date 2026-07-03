# Themes

PPT Skill includes a theme engine that can recolor generated presentations without modifying the JSON deck spec.

## Built-in Themes

| Theme | Style |
|---|---|
| `default` | Blue accent on light background |
| `ocean` | Teal/blue ocean palette |
| `sunset` | Warm orange/red palette |
| `forest` | Green nature palette |
| `berry` | Purple/magenta palette |
| `monochrome` | Grayscale palette |
| `minimal` | Clean gray palette |
| `dark` | Dark background with light text |

## Applying a Theme

```bash
python scripts/build_pptx.py deck.json output.pptx --theme ocean
```

## Custom Themes

Create a JSON file in `references/themes/`:

```json
{
  "name": "my_theme",
  "background": "#FFFFFF",
  "title_color": "#1F2937",
  "body_color": "#4B5563",
  "accent_color": "#3B82F6",
  "secondary_color": "#10B981"
}
```

Apply it by path:

```bash
python scripts/build_pptx.py deck.json output.pptx --theme references/themes/my_theme.json
```

## Theme Engine Behavior

- Background colors are applied to slides.
- Text colors are chosen for contrast against the background.
- Accent colors are used for decorative lines, charts, and highlighted elements.
- The engine detects underlying shape fills to keep white text readable on colored cards.
