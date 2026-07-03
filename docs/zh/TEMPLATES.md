# 模板

PPT Skill 使用基础 PPTX 模板，其中包含预定义的幻灯片布局和占位形状。

## 基础模板

`templates/base_template.pptx` 是默认模板，包含所有支持布局的占位形状。

## 模板工作原理

- 构建器加载模板。
- 每张幻灯片使用模板中的一个布局。
- 名为 `__TITLE__`、`__CONTENT__`、`__IMAGE__` 等的占位形状被生成的内容替换。
- 约束布局系统可以绕过模板占位形状，动态计算元素位置。

## 使用自定义模板

```bash
python scripts/build_pptx.py deck.json output.pptx --template templates/my_template.pptx
```

## 创建模板

1. 在 PowerPoint 或 LibreOffice 中创建 PPTX 文件。
2. 为每种需要的布局添加幻灯片。
3. 添加具有标准名称的占位形状：
   - `__TITLE__` — 标题
   - `__SUBTITLE__` — 副标题
   - `__CONTENT__` — 正文内容
   - `__IMAGE__` — 图片占位
   - `__CHART__` — 图表占位
   - `__TABLE__` — 表格占位
4. 保存模板，并通过 `--template` 引用。

## 模板分析

从任意模板提取视觉契约（颜色、字体、布局结构）：

```bash
python scripts/template_analyzer.py templates/base_template.pptx --output contract.json
```

将契约转换为声明式 Schema：

```bash
python scripts/contract_to_schema.py contract.json --layout cover --output schema.json
```

## 模板市场

`templates/market/` 包含社区或替代模板包。每个包通常包括：

- `template.pptx` — 模板文件
- `config.json` — 元数据和支持的布局
- `analysis.json` — 提取的视觉契约
