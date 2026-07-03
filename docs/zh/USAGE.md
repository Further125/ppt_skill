# 使用指南

## 命令行入口

`scripts/build_pptx.py` 是主要的构建入口。

```bash
python scripts/build_pptx.py <input.json> <output.pptx> [选项]
```

常用选项：

| 选项 | 说明 |
|---|---|
| `--theme <name>` | 应用内置或自定义主题 |
| `--auto-route` | 根据内容自动推断布局 |
| `--animate <effect>` | 添加入场动画 |
| `--template <path>` | 使用自定义 PPTX 模板 |
| `--verbose` | 打印详细调试信息 |

## 自动路由

在 JSON 中省略 `layout` 字段，并加上 `--auto-route`：

```json
{
  "title": "自动路由示例",
  "slides": [
    {"title": "封面", "subtitle": "自动推断"},
    {"quote": "少即是多", "author": "某人"},
    {"big_number": "99.9%", "label": "可用性"}
  ]
}
```

```bash
python scripts/build_pptx.py deck.json output.pptx --auto-route
```

## 主题

内置主题：`default`、`ocean`、`sunset`、`forest`、`berry`、`monochrome`、`minimal`、`dark`。

```bash
python scripts/build_pptx.py deck.json output.pptx --theme ocean
```

自定义主题文件：

```bash
python scripts/build_pptx.py deck.json output.pptx --theme references/themes/my_theme.json
```

## QA 检查

```bash
python scripts/qa_check.py output.pptx
```

自动检测空幻灯片、文本溢出、未替换占位符等问题。

## 其他工具

| 脚本 | 用途 |
|---|---|
| `render_slides.py` | PPTX 转 PNG 预览 |
| `export.py` | PDF / 长图 / HTML 导出 |
| `snapshot.py` | 截取指定页面或区域 |
| `template_analyzer.py` | 从 PPTX 模板提取视觉契约 |
| `contract_to_schema.py` | 将模板契约转换为声明式 Schema |
| `quick_preview.py` | 单页 JSON 转 PNG，用于 LLM 可视化反馈 |
