# 布局参考

演示文稿（deck spec）中的每一页幻灯片都必须包含一个 `layout` 字段。本文档介绍所有支持的布局、对应字段及其视觉效果。

---

## `cover`

全屏标题幻灯片，包含大标题、副标题和可选日期。

```json
{
  "layout": "cover",
  "title": "演示标题",
  "subtitle": "副标题",
  "date": "2026-06-28"
}
```

| 字段 | 类型 | 必填 | 说明 |
|-------|------|----------|-------------|
| `title` | string | 是 | 主标题文本 |
| `subtitle` | string | 否 | 副标题文本 |
| `date` | string | 否 | 日期字符串 |

---

## `toc`

带编号条目的目录页。

```json
{
  "layout": "toc",
  "title": "目录",
  "items": ["简介", "方法", "结果", "结论"]
}
```

| 字段 | 类型 | 必填 | 说明 |
|-------|------|----------|-------------|
| `title` | string | 否 | 章节标题（默认值："目录"） |
| `items` | string[] | 是 | 章节名称列表 |

---

## `title_content`

最常用的布局：上方为标题，下方为正文区域。

```json
{
  "layout": "title_content",
  "title": "核心发现",
  "content": "纯文本段落或项目符号列表",
  "accent_line": true
}
```

`content` 可以是：
- 普通字符串（渲染为段落）
- 字符串列表（渲染为项目符号）
- 富文本运行字典数组（详见 [advanced.md](advanced.md#rich-text)）

| 字段 | 类型 | 必填 | 说明 |
|-------|------|----------|-------------|
| `title` | string | 是 | 幻灯片标题 |
| `content` | string / string[] / object[] | 否 | 正文文本 |
| `accent_line` | bool | 否 | 是否显示蓝色下划线（默认值：true） |

---

## `two_column`

左右并排的文本双栏布局。

```json
{
  "layout": "two_column",
  "title": "对比",
  "left_content": ["要点 A1", "要点 A2"],
  "right_content": ["要点 B1", "要点 B2"]
}
```

| 字段 | 类型 | 必填 | 说明 |
|-------|------|----------|-------------|
| `title` | string | 是 | 幻灯片标题 |
| `left_content` | string / string[] / object[] | 否 | 左侧栏文本 |
| `right_content` | string / string[] / object[] | 否 | 右侧栏文本 |

---

## `chart`

原生 PowerPoint 图表（柱状图、折线图、饼图、环形图等）。

```json
{
  "layout": "chart",
  "title": "销售增长",
  "description": "同比对比",
  "chart_data": {
    "type": "bar",
    "categories": ["第一季度", "第二季度", "第三季度", "第四季度"],
    "values": [120, 150, 180, 210],
    "series_name": "收入"
  }
}
```

| 字段 | 类型 | 必填 | 说明 |
|-------|------|----------|-------------|
| `title` | string | 是 | 幻灯片标题 |
| `description` | string | 否 | 副标题文本 |
| `chart_data` | object | 是 | 图表配置 |

**chart_data 字段：**
- `type`: `"bar"`、`"column"`、`"line"`、`"pie"`、`"doughnut"`、`"area"`、`"scatter"`、`"radar"`、`"bubble"`
- `categories`: string[] — X 轴标签
- `values`: number[] — Y 轴数据
- `series_name`: string — 图例标签

---

## `table`

带样式表头和交替行颜色的数据表格。

```json
{
  "layout": "table",
  "title": "框架对比",
  "description": "关键指标",
  "table_data": {
    "headers": ["名称", "星标数", "语言"],
    "rows": [
      ["LlamaFactory", "72K", "Python"],
      ["Unsloth", "67K", "Python"]
    ]
  }
}
```

| 字段 | 类型 | 必填 | 说明 |
|-------|------|----------|-------------|
| `title` | string | 是 | 幻灯片标题 |
| `description` | string | 否 | 副标题文本 |
| `table_data` | object | 是 | 表格数据 |

**table_data 字段：**
- `headers`: string[] — 列标题
- `rows`: string[][] — 数据行

---

## `timeline`

节点在水平线上方/下方交替排列的时间轴。

```json
{
  "layout": "timeline",
  "title": "项目路线图",
  "timeline_data": {
    "milestones": [
      {"date": "2025 年第一季度", "title": "调研", "desc": "文献综述"},
      {"date": "2025 年第二季度", "title": "设计", "desc": "架构设计"},
      {"date": "2025 年第三季度", "title": "构建", "desc": "功能实现"},
      {"date": "2025 年第四季度", "title": "发布", "desc": "生产上线"}
    ]
  }
}
```

| 字段 | 类型 | 必填 | 说明 |
|-------|------|----------|-------------|
| `title` | string | 是 | 幻灯片标题 |
| `timeline_data` | object | 是 | 里程碑数据 |

**timeline_data 字段：**
- `milestones`: `{date, title, desc}` 数组

---

## `image_content`

左侧图片、右侧文本。

```json
{
  "layout": "image_content",
  "title": "架构图",
  "image_path": "diagram.png",
  "content": ["核心要点 1", "核心要点 2"]
}
```

| 字段 | 类型 | 必填 | 说明 |
|-------|------|----------|-------------|
| `title` | string | 是 | 幻灯片标题 |
| `image_path` | string | 是 | 图片文件路径 |
| `content` | string / string[] / object[] | 否 | 正文文本 |

---

## `quote`

居中显示的引用语及出处。

```json
{
  "layout": "quote",
  "quote": "简洁是最高级的复杂。",
  "author": "列奥纳多·达·芬奇"
}
```

| 字段 | 类型 | 必填 | 说明 |
|-------|------|----------|-------------|
| `quote` | string | 是 | 引用文本 |
| `author` | string | 否 | 出处 |

---

## `team`

团队成员卡片，包含头像、姓名、角色和描述。

```json
{
  "layout": "team",
  "title": "我们的团队",
  "team_data": {
    "members": [
      {"name": "Alice", "role": "产品经理", "desc": "产品战略"},
      {"name": "Bob", "role": "工程师", "desc": "后端系统"}
    ]
  }
}
```

| 字段 | 类型 | 必填 | 说明 |
|-------|------|----------|-------------|
| `title` | string | 是 | 幻灯片标题 |
| `team_data` | object | 是 | 成员数据 |

---

## `data_highlight`

单独显示一个大数字及标签，用于强调关键数据。

```json
{
  "layout": "data_highlight",
  "title": "市场规模",
  "big_number": "$12B",
  "label": "总可寻址市场"
}
```

| 字段 | 类型 | 必填 | 说明 |
|-------|------|----------|-------------|
| `title` | string | 是 | 幻灯片标题 |
| `big_number` | string | 是 | 大数字文本 |
| `label` | string | 是 | 数字下方标签 |

---

## `process`

带编号和标题的水平步骤卡片。

```json
{
  "layout": "process",
  "title": "工作流程",
  "process_data": {
    "steps": [
      {"title": "输入"},
      {"title": "处理"},
      {"title": "输出"}
    ]
  }
}
```

| 字段 | 类型 | 必填 | 说明 |
|-------|------|----------|-------------|
| `title` | string | 是 | 幻灯片标题 |
| `process_data` | object | 是 | 步骤数据 |

**process_data 字段：**
- `steps`: `{title}` 数组 — 每一步对应一张带编号的彩色卡片

---

## `waterfall`

展示累计变化的瀑布图。

```json
{
  "layout": "waterfall",
  "title": "利润分析",
  "description": "从收入到净利润",
  "chart_data": {
    "categories": ["收入", "销售成本", "运营费用", "净利润"],
    "values": [100, -40, -30, 30]
  }
}
```

需要 `scripts/advanced_charts.py`。

---

## `funnel`

展示逐步递减的漏斗图。

```json
{
  "layout": "funnel",
  "title": "转化漏斗",
  "chart_data": {
    "categories": ["访客", "线索", "客户"],
    "values": [1000, 300, 80]
  }
}
```

---

## `gantt`

用于项目排期的甘特图。

```json
{
  "layout": "gantt",
  "title": "项目排期",
  "chart_data": {
    "tasks": [
      {"name": "设计", "start": 1, "duration": 3},
      {"name": "构建", "start": 4, "duration": 5}
    ]
  }
}
```

---

## `wordcloud`

根据词频数据生成的词云图片。

```json
{
  "layout": "wordcloud",
  "title": "关键词云",
  "words": [
    {"text": "AI", "weight": 100},
    {"text": "ML", "weight": 80},
    {"text": "NLP", "weight": 60}
  ]
}
```

---

## `closing`

致谢 / 结束页。

```json
{
  "layout": "closing",
  "title": "感谢聆听",
  "subtitle": "有问题吗？"
}
```

| 字段 | 类型 | 必填 | 说明 |
|-------|------|----------|-------------|
| `title` | string | 否 | 结束语文本（默认值："Thank You"） |
| `subtitle` | string | 否 | 副标题文本 |

---

## 约束布局

约束布局使用求解器自动计算字体大小和位置。它们使用 `_constraint` 后缀，但数据格式与普通布局相同。

### `cover_constraint`

```json
{"layout": "cover_constraint", "title": "演示标题", "subtitle": "副标题", "meta": "2026-06-28"}
```

### `section_constraint`

```json
{"layout": "section_constraint", "title": "章节标题", "chapter_num": "01", "subtitle": "副标题"}
```

### `title_content_constraint`

```json
{"layout": "title_content_constraint", "title": "标题", "content": ["要点 A", "要点 B"]}
```

### `two_column_constraint`

```json
{"layout": "two_column_constraint", "title": "标题", "left_content": ["左侧内容"], "right_content": ["右侧内容"]}
```

### `quote_constraint`

```json
{"layout": "quote_constraint", "quote": "引用语", "author": "作者"}
```

### `data_highlight_constraint`

```json
{"layout": "data_highlight_constraint", "title": "标题", "big_number": "99", "description": ["说明 1", "说明 2"]}
```

### `three_column_constraint`

```json
{"layout": "three_column_constraint", "title": "标题", "col_1": ["栏 1"], "col_2": ["栏 2"], "col_3": ["栏 3"]}
```

### `chart_constraint`

```json
{"layout": "chart_constraint", "title": "标题", "description": ["说明"], "chart_data": {...}}
```

### `table_constraint`

```json
{"layout": "table_constraint", "title": "标题", "description": ["说明"], "table_data": {...}}
```

### `timeline_constraint`

```json
{"layout": "timeline_constraint", "title": "标题",
 "timeline_data": {"milestones": [{"date": "第一季度", "title": "里程碑", "description": "说明"}]}}
```

### `team_constraint`

```json
{"layout": "team_constraint", "title": "标题",
 "team_data": {"members": [{"name": "姓名", "role": "角色", "desc": "说明"}]}}
```

### `process_constraint`

```json
{"layout": "process_constraint", "title": "标题",
 "process_data": {"steps": [{"step": "01", "title": "步骤", "desc": "说明"}]}}
```

### `tree_constraint`

```json
{"layout": "tree_constraint", "title": "标题",
 "tree_data": {"root": "根节点", "children": [{"name": "子节点"}]}}
```

### `image_text_split_constraint`

```json
{"layout": "image_text_split_constraint", "title": "标题", "image": "path.png", "content": ["说明"]}
```

---

## 通用字段（所有布局）

每一页幻灯片还可以包含以下字段：

| 字段 | 类型 | 说明 |
|-------|------|-------------|
| `accent_line` | bool | 是否显示蓝色标题下划线（默认值：true） |
| `template_slide_name` | string | 指定要克隆的模板幻灯片 |
| `shapes` | object[] | 自由形状数组（详见 [advanced.md](advanced.md)） |
