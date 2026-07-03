# 常见问题

## 一般问题

### PPT Skill 是什么？

PPT Skill 是一个基于 Python 的工具，可根据 JSON 描述生成 PowerPoint 文件。它适用于自动化、LLM 集成和批量生成演示文稿。

### 需要安装 Microsoft PowerPoint 吗？

不需要。PPT Skill 使用 `python-pptx` 创建 `.pptx` 文件。你只需要 PowerPoint 或兼容软件来打开输出文件。

### 需要 LibreOffice 吗？

不需要。LibreOffice 是可选的，仅用于高保真 PDF/PNG 渲染。软渲染兜底功能无需 LibreOffice 即可工作。

## 使用问题

### 文本溢出了怎么办？

- 使用约束布局（`*_constraint`）自动计算字号。
- 缩短内容或拆分到多张幻灯片。
- 使用 `qa_check.py` 检测溢出。

### 中文字符显示为方块怎么办？

默认字体是 "Microsoft YaHei"。在 Linux 上安装 CJK 字体：

```bash
# Debian/Ubuntu
sudo apt-get install fonts-noto-cjk
```

或注册项目自带字体：

```bash
sudo mkdir -p /usr/share/fonts/opentype/noto
sudo cp fonts/NotoSansCJKsc-Regular.otf /usr/share/fonts/opentype/noto/
sudo fc-cache -fv
```

### 如何添加自己的布局？

1. 在 `references/layout_schemas/` 中定义声明式 Schema。
2. 或在 `scripts/build_pptx.py` 中添加布局函数。
3. 在 `references/layouts.json` 中注册布局。

### 可以用 Markdown 代替 JSON 吗？

可以。使用 `scripts/md_to_deck.py`：

```bash
python scripts/md_to_deck.py input.md output.json
```

## 故障排查

### 报错 `ModuleNotFoundError: No module named 'pptx'`

安装依赖：

```bash
pip install -r requirements.txt
```

### 生成的 PDF 是文本报告而不是真正的 PDF

未安装 LibreOffice。如需 PDF 导出请安装 LibreOffice，或使用长图/HTML 导出（无需 LibreOffice）。

### 预览中看不到动画

动画已嵌入 PPTX，但仅在 PowerPoint 或 WPS 中打开时可见。PNG 预览不渲染动画。
