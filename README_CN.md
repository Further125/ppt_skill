# PPT Skill — 基于 JSON 的 PPT 生成器

[English README](README.md) | 本文档为中文版

基于 Python + python-pptx 的 PPT 自动生成系统。编写 JSON 描述文件，生成精美的 `.pptx`。

## 快速开始

```bash
# 安装依赖
pip install -r requirements.txt

# 从 JSON 生成 PPT
python scripts/build_pptx.py examples/demo_deck.json output/my_deck.pptx

# 应用主题
python scripts/build_pptx.py examples/demo_deck.json output/my_deck.pptx --theme ocean

# 导出 PDF
python scripts/export.py output/my_deck.pptx --format pdf --output my_deck.pdf
```

## 文档

- [README.md](README.md) — 英文版
- [docs/zh/QUICK_START.md](docs/zh/QUICK_START.md) — 快速开始
- [docs/zh/USAGE.md](docs/zh/USAGE.md) — 使用指南
- [docs/zh/layouts.md](docs/zh/layouts.md) — 布局速查
- [docs/zh/user-guide.md](docs/zh/user-guide.md) — 完整使用指南
- [docs/zh/architecture.md](docs/zh/architecture.md) — 架构设计
- [docs/zh/advanced.md](docs/zh/advanced.md) — 高级功能
- [docs/zh/comparison.md](docs/zh/comparison.md) — 对比分析
- [docs/zh/THEMES.md](docs/zh/THEMES.md) — 主题系统
- [docs/zh/TEMPLATES.md](docs/zh/TEMPLATES.md) — 模板系统
- [docs/zh/EXAMPLES.md](docs/zh/EXAMPLES.md) — 示例集
- [docs/zh/FAQ.md](docs/zh/FAQ.md) — 常见问题
- [docs/zh/TESTING.md](docs/zh/TESTING.md) — 测试指南

## 主要特性

- **约束布局系统**：二分查找自动计算字号，attach-chain 自动编排元素位置
- **声明式布局 Schema**：用 JSON 定义布局，支持百分比坐标
- **意图路由**：省略 `layout` 字段，系统根据内容自动推断布局
- **20+ 种布局**：封面、目录、单栏、双栏、三栏、图表、表格、时间轴、流程图、团队卡片、树状图、词云等
- **富文本与自由形状**：支持按 run 设置粗体/颜色/字号，图片遮罩、裁剪、层级控制
- **导出工具**：PDF、长图、HTML 演示文稿

## 项目结构

```
ppt_skill/
├── README.md                   # 英文完整文档
├── README_CN.md                # 本文档（中文版）
├── LICENSE                     # MIT 许可证
├── requirements.txt            # Python 依赖
├── scripts/                    # 核心脚本
├── templates/                  # PPTX 模板
├── references/                 # 布局配置与主题
├── examples/                   # 示例 JSON
├── tests/                      # 测试用例
├── docs/                       # 英文文档
├── docs/zh/                    # 中文文档
└── output/                     # 输出目录
```

## 支持的布局

| 布局 | 描述 | 关键字段 |
|---|---|---|
| `cover` | 封面页 | `title`, `subtitle`, `date` |
| `toc` | 目录页 | `title`, `items` |
| `title_content` | 标题 + 内容 | `title`, `content` |
| `two_column` | 双栏对比 | `title`, `left_content`, `right_content` |
| `chart` | 原生图表 | `title`, `description`, `chart_data` |
| `table` | 原生表格 | `title`, `description`, `table_data` |
| `timeline` | 时间轴 | `title`, `timeline_data` |
| `team` | 团队卡片 | `title`, `team_data` |
| `quote` | 引用页 | `quote`, `author` |
| `data_highlight` | 大数字强调 | `title`, `big_number`, `label` |
| `image_content` | 图文页 | `title`, `image_path`, `content` |
| `wordcloud` | 词云 | `title`, `words` |
| `closing` | 结尾页 | `title`, `subtitle` |

详见 [docs/zh/layouts.md](docs/zh/layouts.md)。

## JSON 示例

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

## 许可

MIT
