# Themes

PPT Skill includes a theme engine that can recolor generated presentations without modifying the JSON deck spec.

## Built-in Themes

| Theme | Style |
|---|---|
| `default` | Blue accent on light background |
| `dark` | Dark background with light text |
| `warm` | Warm orange/red palette |
| `forest` | Green nature palette |
| `minimal` | Clean gray palette |

## Custom Themes in `references/themes/`

| Theme | File | Style |
|---|---|---|
| `ocean_tech` | `references/themes/ocean_tech.json` | Teal/blue tech palette |
| `gold_corporate` | `references/themes/gold_corporate.json` | Gold corporate palette |
| `rose_elegant` | `references/themes/rose_elegant.json` | Purple/magenta elegant palette |
| `white_gold` | `references/themes/white_gold.json` | White/gold corporate palette |
| `cyber_neon` | `references/themes/cyber_neon.json` | Cyberpunk neon palette |

## Applying a Theme

```bash
python scripts/build_pptx.py deck.json output.pptx --theme dark
python scripts/build_pptx.py deck.json output.pptx --theme ocean_tech
```

## Creating a Custom Theme

Create a JSON file in `references/themes/`:

```json
{
  "name": "my_theme",
  "primary": "#3B82F6",
  "secondary": "#0B1F3A",
  "accent": ["#3B82F6", "#10B981", "#F59E0B"],
  "background": "#FFFFFF",
  "text": "#1F2937",
  "light_bg": "#F3F4F6",
  "font_title": "Microsoft YaHei",
  "font_body": "Microsoft YaHei"
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
