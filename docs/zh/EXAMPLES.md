# 示例集

`examples/` 目录包含展示 PPT Skill 各种功能的 JSON 示例。

## 入门示例

| 文件 | 说明 |
|---|---|
| `demo_deck.json` | 最小示例，包含封面、内容、结尾页 |
| `full_demo.json` | 多种布局和功能的综合示例 |
| `preview_test.json` | 用于 `quick_preview.py` 的单页示例 |

## 布局示例

| 文件 | 说明 |
|---|---|
| `accent_line_demo.json` | 装饰线开关演示 |
| `title_content` 布局 | 见 `demo_deck.json`、`full_demo.json` |
| `two_column` / `three_column` | 见 `full_demo.json` |
| `image_content` | 见 `image_mask_demo.json` |

## 图表与表格示例

| 文件 | 说明 |
|---|---|
| `charts_anim_demo.json` | 带动画的原生图表 |
| `table_test.json` | 表格布局演示 |
| `advanced_demo.json` | 高级图表：瀑布图、漏斗图、甘特图 |

## 样式示例

| 文件 | 说明 |
|---|---|
| `style_dark_cyber.json` | 深色赛博主题 |
| `style_dark_cyber_built.pptx` | 深色赛博主题的预构建输出 |
| `style_minimal_corporate.json` | 极简商务主题 |
| `style_infographic_rich.json` | 信息图风格 |

## 高级功能示例

| 文件 | 说明 |
|---|---|
| `rich_text_demo.json` | 混合格式文本（粗体、颜色、字号） |
| `free_shape_demo.json` | 自由形状与定位 |
| `image_mask_demo.json` | 图片遮罩与裁剪 |
| `schema_test.json` | 声明式布局 Schema 示例 |
| `schema_comprehensive.json` | 综合 Schema 布局示例 |
| `compare_test.json` | 并排对比布局 |

## 意图路由

| 文件 | 说明 |
|---|---|
| `auto_route_test.json` | `--auto-route` 自动布局推断示例 |

## 运行示例

```bash
python scripts/build_pptx.py examples/demo_deck.json output/demo.pptx
python scripts/build_pptx.py examples/full_demo.json output/full.pptx --theme ocean_tech_tech
python scripts/build_pptx.py examples/auto_route_test.json output/auto.pptx --auto-route
```

## 创建自己的 Deck

从 `examples/demo_deck.json` 开始，修改 `slides` 数组。每张幻灯片需要 `layout` 字段（使用 `--auto-route` 时可省略）和该布局对应的内容字段。
