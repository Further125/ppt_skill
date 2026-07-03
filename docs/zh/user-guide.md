> 本文档为中文版。

# PPT Skill -- 完整使用指南

> 基于 JSON 驱动的原生 PPTX 自动生成系统。

---

## 1. 安装与环境

### 依赖

- Python 3.10+
- 已预装: python-pptx, Pillow, PyMuPDF
- 可选: LibreOffice (高保真渲染)

### 验证安装

```bash
cd ppt_skill
python scripts/build_pptx.py examples/demo_deck.json /tmp/test.pptx
```

---

## 2. 快速上手

```json
{
  "title": "AI Agent 概述",
  "slides": [
    { "layout": "cover", "title": "封面" },
    { "layout": "title_content", "title": "内容", "content": ["A", "B"] },
    { "layout": "closing", "title": "谢谢" }
  ]
}
```

```bash
python scripts/build_pptx.py deck.json output.pptx
python scripts/build_pptx.py deck.json output_dark.pptx --theme dark
```

---

## 3. JSON 规格

### 顶层结构

| 字段 | 类型 | 必填 | 说明 |
|------|------|------|------|
| title | string | 否 | Deck 标题 |
| subtitle | string | 否 | 副标题 |
| slides | array | 是 | 幻灯片数组 |

### 通用字段

| 字段 | 类型 | 默认值 | 说明 |
|------|------|--------|------|
| layout | string | 推断 | 布局类型 |
| accent_line | bool | true | 标题下划线 |
| shapes | array | -- | 自由形状 |
| hyperlink | string | -- | 超链接 |

### 内容格式

- 纯字符串: "content": "文本"
- 项目符号: "content": ["A", "B"]
- 富文本: "content": [{"text": "粗体", "bold": true}]

---

## 4. 布局参考

| 布局 | 用途 | 关键字段 |
|------|------|---------|
| cover | 封面 | title, subtitle, date |
| toc | 目录 | title, items |
| title_content | 标题+正文 | title, content |
| two_column | 双栏 | title, left_content, right_content |
| chart | 图表 | title, description, chart_data |
| table | 表格 | title, description, table_data |
| timeline | 时间轴 | title, timeline_data |
| image_content | 图文 | title, image_path, content |
| quote | 引用 | quote, author |
| team | 团队 | title, team_data |
| data_highlight | 大数字 | title, big_number, label |
| process | 流程 | title, process_data |
| tree | 树状图 | title, tree_data |
| waterfall | 瀑布图 | title, description, chart_data |
| funnel | 漏斗图 | title, description, chart_data |
| gantt | 甘特图 | title, description, chart_data |
| wordcloud | 词云 | title, description, words |
| closing | 结尾 | title, subtitle |

---

## 5. 图表

原生图表 type: column, bar, line, pie, doughnut, area, scatter, radar, bubble

高级图表: waterfall(瀑布图), funnel(漏斗图), gantt(甘特图), wordcloud(词云)

---

## 6. 主题

内置主题：default、dark、warm、forest、minimal。自定义主题：cyber_neon、gold_corporate、ocean_tech、rose_elegant、white_gold。

```bash
python scripts/build_pptx.py deck.json out.pptx --theme dark
```

自定义主题: references/themes/ 下创建 JSON，含 primary, secondary, accent, background, text, light_bg, font_title, font_body

---

## 7. 高级功能

### 自由形状

坐标格式: 60%, 2in, 5cm, 100px, 或 EMU
形状类型: rectangle, rounded_rectangle, oval, triangle, diamond, pentagon, hexagon, star, arrow_*, chevron, parallelogram, trapezoid, donut

### 图片遮罩

mask: circle, rounded_rectangle, oval, triangle, diamond, hexagon, star, pentagon, heart, cloud, sun, moon

### 动画

```bash
python scripts/build_pptx.py deck.json out.pptx --animate fade
```

效果: fade, fly_in, wipe, zoom, bounce, appear, swivel, float

---

## 8. 导出与预览

```bash
# 渲染预览
python scripts/render_slides.py deck.pptx preview_dir --dpi 150

# PDF
python scripts/export.py deck.pptx --format pdf --output deck.pdf

# 长图
python scripts/export.py deck.pptx --format long_image --output deck.png

# HTML
python scripts/export.py deck.pptx --format html --output deck.html

# 快照
python scripts/snapshot.py deck.pptx --page 3 --output page3.png
```

---

## 9. 质量检查

```bash
# QA 检查
python scripts/qa_check.py output.pptx

# 色彩对比度审计
python scripts/color_guard.py output.pptx --theme-json references/themes/dark.json --format markdown
```

color_guard 检测项: 文本与背景对比度(WCAG)、形状填充与文本对比度、背景眩光
输出级别: CRITICAL (< 2.0) / WARN (< 4.5) / INFO

---

## 10. CLI 参考

### build_pptx.py

| 选项 | 说明 |
|------|------|
| --template PATH | 自定义模板 |
| --theme NAME | 主题或 JSON 路径 |
| --animate EFFECT | 动画 |
| --no-stagger | 取消延迟 |
| --auto-route | 自动推断布局 |
| --export FORMAT | pdf / long_image / html |
| --audit | 自动色彩审计 |

### render_slides.py

| 选项 | 说明 |
|------|------|
| --dpi N | 渲染 DPI(默认 150) |
| --soft | 纯 Python 软渲染 |
| --pages LIST | 仅渲染指定页 |

---

## 11. 常见问题

**Q: Linux 下字体显示异常?**
A: 安装 fonts-noto-cjk 或复制项目自带的 NotoSansCJKsc-Regular.otf

**Q: 动画在预览中看不到?**
A: 动画需在 PowerPoint / WPS 中打开才能看到效果

**Q: 图表颜色太浅/太深?**
A: 使用 --theme 指定主题，或运行 color_guard 检查对比度

**Q: 如何自定义模板?**
A: 使用 scripts/template_analyzer.py 分析现有 PPTX，生成 contract.json，再转换为 schema

---

## 12. Markdown → Deck 转换

将 Markdown 文件直接转换为 JSON deck：

```bash
python scripts/md_to_deck.py input.md --output deck.json --theme dark
```

### 支持的 Markdown 语法

| 语法 | 输出 |
|------|------|
| `# 标题` | cover 布局 |
| `## 标题` | title_content 布局 |
| `- 项目` | 项目符号列表 |
| `1. 项目` | 编号列表 |
| `> 引用 -- 作者` | quote 布局 |
| `\|表格\|` | table 布局 |
| `\`\`\`python` | 代码块（见下方） |
| `\`\`\`chart` | chart/tree/highlight/wordcloud/team |

### 嵌套列表

用缩进控制层级（2 空格 = 1 级）：

```markdown
- 一级项目
  - 二级项目
    - 三级项目
```

渲染为带缩进的层级 bullets。

### 代码块

Markdown 代码块自动渲染为带标签的代码框：

```markdown
    ```python
    def hello():
        return "world"
    ```
```

效果：
- 顶部灰色标签栏显示语言名（如 `python`）
- 等宽字体（Courier New）
- 浅灰背景（#F3F4F6）
- 支持的语言：`python`, `js`, `java`, `cpp`, `json`, `yaml`, `xml`, `sql`, `bash`, `sh`

### URL 自动超链接

Markdown 中的 URL 自动转换为蓝色下划线超链接：

```markdown
- 官网：https://example.com
```

---

## 14. 约束布局系统

### 14.1 概述

约束布局是 ppt_skill 的核心特性。它通过**约束求解器**自动计算每页的最大可读字号，并通过 **attach 链**自动编排元素位置。

**核心特性：**
- 二分搜索找最大字号
- font_scale 比例控制（标题大、正文小）
- attach 相对定位（避免硬编码坐标）
- 内容少时自动 expand 填充
- 内容多时自动 shrink 压缩行距
- 溢出时截断并预警

### 14.2 可用约束模板

| 模板 | 用途 | 关键字段 |
|------|------|---------|
| `cover_constraint` | 封面 | `title`, `subtitle`, `meta` |
| `section_constraint` | 章节页 | `title`, `subtitle`, `chapter_num` |
| `title_content_constraint` | 标题+正文 | `title`, `content` |
| `two_column_constraint` | 双栏对比 | `title`, `left_content`, `right_content` |
| `quote_constraint` | 引用 | `quote`, `author` |
| `data_highlight_constraint` | 大数字 | `title`, `big_number`, `description` |
| `three_column_constraint` | 三栏 | `title`, `col_1`, `col_2`, `col_3` |
| `chart_constraint` | 图表 | `title`, `description`, `chart_data` |
| `table_constraint` | 表格 | `title`, `description`, `table_data` |
| `timeline_constraint` | 时间线 | `title`, `timeline_data` |
| `team_constraint` | 团队 | `title`, `team_data` |
| `process_constraint` | 流程 | `title`, `process_data` |
| `tree_constraint` | 树形图 | `title`, `tree_data` |
| `image_text_split_constraint` | 左图右文 | `title`, `image`, `content` |

### 14.3 数据格式

所有约束模板的数据格式与固定布局一致，但多了一套 `_constraint` 后缀的模板：

```json
{
  "layout": "title_content_constraint",
  "title": "约束布局演示",
  "content": ["字体尽量大", "边框自动编排", "溢出自动截断"]
}
```

### 14.4 特殊数据格式

**timeline_data:**
```json
{
  "timeline_data": {
    "milestones": [
      {"date": "2024 Q1", "title": "立项", "description": "产品立项"},
      {"date": "2024 Q2", "title": "MVP", "description": "最小可行产品"}
    ]
  }
}
```

**team_data:**
```json
{
  "team_data": {
    "members": [
      {"name": "张三", "role": "CEO", "desc": "战略"},
      {"name": "Member A", "role": "CTO", "desc": "技术"}
    ]
  }
}
```

**process_data:**
```json
{
  "process_data": {
    "steps": [
      {"step": "01", "title": "需求", "desc": "收集需求"},
      {"step": "02", "title": "设计", "desc": "方案设计"}
    ]
  }
}
```

**tree_data:**
```json
{
  "tree_data": {
    "root": "产品架构",
    "children": [
      {"name": "前端", "children": [{"name": "Web"}, {"name": "App"}]},
      {"name": "后端", "children": [{"name": "API"}]}
    ]
  }
}
```

### 14.5 构建约束模板 deck

```bash
python scripts/build_pptx.py deck.json output.pptx --theme minimal
```

约束模板支持所有主题。

---

## 15. 更新日志

### 2026-07-01

- **代码块渲染**：支持 `[lang]...[/code]` 标记，带语言标签栏和等宽字体
- **嵌套列表**：支持 4 级缩进，通过 `indent:N|` 前缀控制
- **URL 超链接**：自动检测并转换为蓝色下划线链接
- **雷达图修复**：回退到原生 XL_CHART_TYPE.RADAR_MARKERS 渲染
- **代码块背景**：标签栏 + 代码区双矩形无缝衔接

---

*文档版本: 2026-07-01*