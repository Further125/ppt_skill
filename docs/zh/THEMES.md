# 主题

PPT Skill 的主题引擎可以在不修改 JSON 描述文件的情况下，为生成的演示文稿重新着色。

## 内置主题

| 主题 | 风格 |
|---|---|
| `default` | 浅色背景 + 蓝色强调 |
| `ocean` | 青蓝海洋色调 |
| `sunset` | 暖橙红色调 |
| `forest` | 绿色自然色调 |
| `berry` | 紫红浆果色调 |
| `monochrome` | 灰度色调 |
| `minimal` | 简洁灰色调 |
| `dark` | 深色背景 + 浅色文字 |

## 应用主题

```bash
python scripts/build_pptx.py deck.json output.pptx --theme ocean
```

## 自定义主题

在 `references/themes/` 下创建 JSON 文件：

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

通过路径应用：

```bash
python scripts/build_pptx.py deck.json output.pptx --theme references/themes/my_theme.json
```

## 主题引擎行为

- 为背景应用背景色。
- 根据背景自动选择对比度合适的文字颜色。
- 强调色用于装饰线、图表和高亮元素。
- 引擎会检测文字下方的形状填充，确保彩色卡片上的白色文字仍然可读。
