# PPT Skill — AI-Powered PowerPoint Generator

> 基于 Python + python-pptx 的 PPT 自动生成系统。
> 支持 13 种幻灯片布局、7 种原生图表、原生表格、时间轴、流程图、
> 团队卡片、动画效果，以及主题换肤、多媒体交互等扩展功能。

📄 文档语言 / Documentation Languages:
- [README.md](README.md) — 英文完整文档 / English documentation
- [README.zh-CN.md](README.zh-CN.md) — 中文完整文档 / Chinese documentation
- [docs/QUICK_START.md](docs/QUICK_START.md) — 英文快速开始
- [docs/USAGE.md](docs/USAGE.md) — 英文使用指南
- [docs/layouts.md](docs/layouts.md) — 英文布局速查
- [docs/THEMES.md](docs/THEMES.md) — 英文主题系统
- [docs/TEMPLATES.md](docs/TEMPLATES.md) — 英文模板系统
- [docs/FAQ.md](docs/FAQ.md) — 英文常见问题
- [docs/zh/QUICK_START.md](docs/zh/QUICK_START.md) — 中文快速开始
- [docs/zh/USAGE.md](docs/zh/USAGE.md) — 中文使用指南
- [docs/zh/layouts.md](docs/zh/layouts.md) — 中文布局速查
- [docs/zh/THEMES.md](docs/zh/THEMES.md) — 中文主题系统
- [docs/zh/TEMPLATES.md](docs/zh/TEMPLATES.md) — 中文模板系统
- [docs/zh/FAQ.md](docs/zh/FAQ.md) — 中文常见问题

---

## 快速开始

> **环境要求**：Python 3.10+，依赖 `python-pptx` / `Pillow` / `PyMuPDF`。
> 可选安装 LibreOffice 用于高保真渲染，Noto Sans CJK 用于中文显示。
> 建议先创建并激活虚拟环境，然后安装依赖：
> ```bash
> python -m venv .venv
> source .venv/bin/activate  # Windows: .venv\Scripts\activate
> pip install -r requirements.txt
> ```

```bash
# 0. 进入 skill 目录
cd ppt_skill

# 1. 写 JSON 描述文件（见下方格式）

# 2. 构建 PPT（用系统 python，已装好 python-pptx）
python3 scripts/build_pptx.py deck.json output.pptx

# 3. 渲染预览图（默认走 LibreOffice 高保真；--soft 走纯 Python 兜底）
python3 scripts/render_slides.py output.pptx preview_dir

# 4. QA 检查
python3 scripts/qa_check.py output.pptx

# 5. 导出 PDF/长图/HTML
python3 scripts/export.py output.pptx --format pdf --output output.pdf
```

---

## 项目结构

```
ppt_skill/
├── SKILL.md                    # 本文档
├── templates/
│   ├── base_template.pptx      # 基础模板（13种布局）
│   └── create_template.py      # 模板生成脚本
├── scripts/
│   ├── build_pptx.py           # 核心构建引擎
│   ├── render_slides.py        # PPTX → PNG 预览
│   ├── qa_check.py             # 质量检查
│   ├── open_pptx.py            # 读取现有 PPT
│   ├── animator.py             # 动画引擎
│   ├── theme_engine.py         # 主题换肤引擎
│   └── export.py               # 导出工具（PDF/长图/网页）
├── references/
│   └── layouts.json            # 布局配置目录
├── examples/
│   ├── demo_deck.json          # 基础功能示例
│   ├── full_demo.json          # 完整功能示例
│   └── charts_anim_demo.json   # 图表+动画示例
└── output/                     # 输出目录
```

---

## 支持的布局（13种）

| 布局类型 | 描述 | 关键字段 |
|---------|------|---------|
| `cover` | 封面页（深蓝背景+标题+副标题+日期） | title, subtitle, date |
| `toc` | 目录页 | title, items[] |
| `title_content` | 标题+单栏内容 | title, content（字符串或列表） |
| `two_column` | 双栏对比 | title, left_content, right_content |
| `chart` | 原生图表页 | title, description, chart_data |
| `table` | 原生表格页 | title, description, table_data |
| `timeline` | 水平时间轴 | title, timeline_data.milestones[] |
| `image_content` | 左图右文 | title, image_path, content |
| `quote` | 金句引用页 | quote, author |
| `team` | 团队卡片 | title, team_data.members[] |
| `data_highlight` | 大数字强调 | title, big_number, label |
| `process` | 流程步骤 | title, process_data.steps[] |
| `closing` | 结尾感谢页 | title, subtitle |

---

## 图表类型（7种原生图表）

chart_data 中设置 `type` 字段：

| type | 图表 |
|------|------|
| `column` | 柱状图 |
| `line` | 折线图（带标记） |
| `pie` | 饼图 |
| `area` | 面积图 |
| `bar` | 条形图（水平） |
| `doughnut` | 环形图 |
| `radar` | 雷达图 |

chart_data 示例：
```json
{
  "type": "column",
  "categories": ["2022", "2023", "2024"],
  "values": [1200, 1850, 2800],
  "series_name": "营收（万元）"
}
```

---

## 动画效果

构建时添加 `--animate` 参数：

```bash
python scripts/build_pptx.py deck.json out.pptx --animate fade
```

支持的效果：`fade`, `fly_in`, `wipe`, `zoom`, `bounce`, `appear`, `swivel`, `float`

添加 `--no-stagger` 取消延迟，所有元素同时进场。

---

## JSON 格式规范

```json
{
  "title": "演示标题",
  "slides": [
    {
      "layout": "cover",
      "title": "主标题",
      "subtitle": "副标题",
      "date": "2026年6月"
    },
    {
      "layout": "title_content",
      "title": "内容页",
      "content": ["要点1", "要点2", "要点3"]
    }
  ]
}
```

每个 slide 必须包含 `layout` 字段，其他字段根据布局类型变化。

---

## 主题换肤

```bash
# 应用深色主题
python scripts/build_pptx.py deck.json out.pptx --theme dark

# 应用自定义主题
python scripts/build_pptx.py deck.json out.pptx --theme /path/to/theme.json
```

主题文件格式见 `references/themes/`。

---

## 导出功能

```bash
# 导出 PDF
python scripts/export.py out.pptx --format pdf

# 导出长图（所有幻灯片拼接）
python scripts/export.py out.pptx --format long_image

# 导出网页（HTML演示文稿）
python scripts/export.py out.pptx --format html
```

---

## 高级可视化扩展

支持组合图、瀑布图、漏斗图、甘特图等复杂图表：

```json
{
  "layout": "chart",
  "title": "组合图示例",
  "chart_data": {
    "type": "combo",
    "categories": ["Q1", "Q2", "Q3", "Q4"],
    "series": [
      {"name": "营收", "type": "column", "values": [100, 150, 200, 250]},
      {"name": "增长率", "type": "line", "values": [10, 15, 20, 25]}
    ]
  }
}
```

---

## 多媒体与交互

```json
{
  "layout": "image_content",
  "title": "含视频的页面",
  "image_path": "/path/to/video.mp4",
  "content": "视频说明文字",
  "hyperlink": "https://example.com"
}
```

---

## QA 检查项

qa_check.py 自动检测：
- 未替换的占位符（__TITLE__ 等残留）
- 文本溢出（密度过高警告）
- 空幻灯片（无可见内容）

输出：`gate_result.json` 格式结果。

---

## 注意事项

1. **字体**：模板使用 "Microsoft YaHei"，Linux 环境建议安装 `fonts-noto-cjk`
2. **渲染**：软渲染（Pillow）为近似预览，安装 LibreOffice + poppler 可获得像素级精确预览
3. **图片路径**：image_path 使用绝对路径或相对于 JSON 文件的路径
4. **动画**：需在 PowerPoint / WPS 中打开才能看到效果

