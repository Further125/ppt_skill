# 快速开始

## 1. 安装

```bash
pip install -r requirements.txt
```

推荐 Python：3.10+

## 2. 编写 JSON 描述文件

```json
{
  "title": "我的第一个演示",
  "slides": [
    {
      "layout": "cover",
      "title": "Hello PPT Skill",
      "subtitle": "JSON 驱动的 PPT 生成"
    },
    {
      "layout": "title_content",
      "title": "核心要点",
      "content": [
        "编写一次 JSON",
        "自动生成 PPTX",
        "应用主题和布局"
      ]
    }
  ]
}
```

保存为 `my_deck.json`。

## 3. 构建 PPT

```bash
python scripts/build_pptx.py my_deck.json output/my_deck.pptx
```

## 4. 预览

```bash
# 软渲染预览（无需 LibreOffice）
python scripts/render_slides.py output/my_deck.pptx output/preview --soft
```

## 5. 导出

```bash
python scripts/export.py output/my_deck.pptx --format pdf --output my_deck.pdf
```
