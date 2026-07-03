# 主题

PPT Skill 的主题引擎可以在不修改 JSON 描述文件的情况下，为生成的演示文稿重新着色。

## 内置主题

| 主题 | 风格 |
|---|---|
| `default` | 浅色背景 + 蓝色强调 |
| `dark` | 深色背景 + 浅色文字 |
| `warm` | 暖橙红色调 |
| `forest` | 绿色自然色调 |
| `minimal` | 简洁灰色调 |

## `references/themes/` 中的自定义主题

| 主题 | 文件 | 风格 |
|---|---|---|
| `ocean_tech` | `references/themes/ocean_tech.json` | 青蓝科技色调 |
| `gold_corporate` | `references/themes/gold_corporate.json` | 金色商务色调 |
| `rose_elegant` | `references/themes/rose_elegant.json` | 紫红优雅色调 |
| `white_gold` | `references/themes/white_gold.json` | 白金商务色调 |
| `cyber_neon` | `references/themes/cyber_neon.json` | 赛博朋克霓虹色调 |

## 应用主题

```bash
python scripts/build_pptx.py deck.json output.pptx --theme dark
python scripts/build_pptx.py deck.json output.pptx --theme ocean_tech
```

## 创建自定义主题

在 `references/themes/` 下创建 JSON 文件：

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

通过路径应用：

```bash
python scripts/build_pptx.py deck.json output.pptx --theme references/themes/my_theme.json
```

## 主题引擎行为

- 为背景应用背景色。
- 根据背景自动选择对比度合适的文字颜色。
- 强调色用于装饰线、图表和高亮元素。
- 引擎会检测文字下方的形状填充，确保彩色卡片上的白色文字仍然可读。
