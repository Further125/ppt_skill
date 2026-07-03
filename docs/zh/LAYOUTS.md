# 布局速查

## 通用字段

大多数幻灯片支持：

- `layout`（字符串，必填）：布局名称
- `title`（字符串）：标题
- `subtitle`（字符串）：副标题
- `accent_line`（布尔值）：是否显示标题下方的装饰线

## 经典布局

| 布局 | 描述 | 关键字段 |
|---|---|---|
| `cover` | 封面页 | `title`, `subtitle`, `date` |
| `toc` | 目录页 | `title`, `items` |
| `section` | 章节分隔页 | `title`, `number` |
| `title_content` | 标题 + 正文 | `title`, `content` |
| `two_column` | 双栏对比 | `title`, `left_content`, `right_content` |
| `three_column` | 三栏布局 | `title`, `columns` |
| `chart` | 原生图表 | `title`, `description`, `chart_data` |
| `table` | 原生表格 | `title`, `description`, `table_data` |
| `timeline` | 水平时间轴 | `title`, `timeline_data` |
| `process` | 流程步骤 | `title`, `process_data` |
| `team` | 团队卡片 | `title`, `team_data` |
| `quote` | 引用页 | `quote`, `author` |
| `data_highlight` | 大数字强调 | `title`, `big_number`, `label` |
| `image_content` | 图文页 | `title`, `image_path`, `content` |
| `image_text_split` | 左图右文 | `title`, `image_path`, `content` |
| `wordcloud` | 词云 | `title`, `words` |
| `closing` | 结尾页 | `title`, `subtitle` |

## 约束布局

基于约束自动计算字号和位置，布局名以 `_constraint` 结尾。

| 布局 | 描述 |
|---|---|
| `cover_constraint` | 自适应封面 |
| `section_constraint` | 章节分隔 |
| `title_content_constraint` | 标题 + 列表 |
| `two_column_constraint` | 双栏对比 |
| `three_column_constraint` | 三栏等分 |
| `quote_constraint` | 大引用 |
| `data_highlight_constraint` | 大数字 |
| `chart_constraint` | 约束版图表 |
| `table_constraint` | 约束版表格 |
| `timeline_constraint` | 约束版时间轴 |
| `team_constraint` | 约束版团队卡片 |
| `process_constraint` | 约束版流程 |
| `tree_constraint` | 约束版树状图 |
| `image_hero_constraint` | 全宽图片英雄区 |
| `image_text_split_constraint` | 约束版左图右文 |
| `hero_top_constraint` | 顶部标题 + 底部内容 |

## 图表类型

`chart_data.type` 可选值：

- `column` 柱状图
- `line` 折线图
- `bar` 条形图
- `pie` 饼图
- `doughnut` 环形图
- `area` 面积图
- `radar` 雷达图
- `bubble` 气泡图
- `scatter` 散点图
- `combo` 组合图

## 示例

```json
{
  "layout": "chart",
  "title": "营收增长",
  "description": "同比增长",
  "chart_data": {
    "type": "column",
    "categories": ["2022", "2023", "2024"],
    "values": [1200, 1850, 2800],
    "series_name": "营收"
  }
}
```
