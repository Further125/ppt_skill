> 本文档为中文版。

# PPT Skill -- 架构设计文档

> 面向开发者的内部架构、模块设计与扩展指南。

---

## 1. 项目结构

```
ppt_skill/
├── scripts/
│   ├── build_pptx.py         # 核心构建引擎
│   ├── theme_engine.py       # 主题换肤引擎
│   ├── color_guard.py        # WCAG 色彩审计
│   ├── layout_schema.py      # 声明式布局 Schema
│   ├── intent_router.py      # 自动布局推断
│   ├── template_analyzer.py  # 模板视觉契约提取
│   ├── contract_to_schema.py # 契约转 Schema
│   ├── text_fitter.py        # 字体自适应
│   ├── animator.py           # 动画引擎
│   ├── advanced_charts.py    # 高级图表
│   ├── render_slides.py      # PPTX -> PNG
│   ├── export.py             # 导出工具
│   ├── snapshot.py           # 快照截取
│   └── qa_check.py           # 质量检查
├── templates/
│   └── base_template.pptx    # 主模板
├── references/
│   ├── themes/               # 主题 JSON
│   └── layout_schemas/       # 布局 Schema
├── examples/                 # 示例 Deck
└── docs/                     # 文档
```

---

## 2. 核心构建流程

build_pptx.py 三阶段构建:

### Stage 1: 解析与路由
1. 加载 deck JSON
2. --auto-route 时调用 intent_router 推断 layout
3. 加载模板 PPTX，提取可用布局

### Stage 2: 逐页构建
对每页 slide_spec:
1. 优先 schema 渲染(layout_schema.render_slide_from_schema)
2. schema 不存在则回退模板克隆(clone_slide)
3. 根据 layout 类型调用对应渲染函数
4. 同步标题装饰线(sync_accent_line)
5. 处理 hyperlink/video_path/shapes
6. 检测溢出并调整(_adjust_slide_for_overflow)

### Stage 3: 后处理
1. 应用主题(theme_engine.apply_theme)
2. 应用动画(animator)
3. 色彩审计(color_guard)
4. 导出(export)

---

## 3. 布局系统

### 3.1 传统布局(模板克隆)
- 从 base_template.pptx 找到对应模板幻灯片
- copy.deepcopy 克隆 XML 元素到新幻灯片
- 占位符文本匹配(如 __TITLE__, __CONTENT__)
- 优势: 兼容性好，继承母版样式
- 局限: 布局固定，难以精确控制

### 3.2 声明式 Schema 布局(新)
JSON Schema 定义布局，支持:
- 百分比坐标(5%, left-half, center)
- attach 相对定位(attach to title bottom + 1%)
- content_source 字段映射
- adaptive 溢出策略(shrink/expand/truncate)

---

## 4. 主题引擎

theme_engine.py 设计原理:

### 4.1 颜色映射
- 定义旧版固定颜色常量(OLD_ACCENT_BLUE, OLD_DARK_TEXT 等)
- 遍历所有形状，将旧颜色映射到主题新颜色
- 背景形状(>80% 面积)设为 theme.background
- 图表背景强制浅色，文本深色(保证可读性)
- 表格保持浅色背景 + 深色文字

### 4.2 文本颜色智能选择
根据形状填充色计算 WCAG 对比度，自动选择白字或黑字:
- 计算白色和深色文本与填充色的对比度
- 选择对比度更高的颜色
- 大面积深色背景使用 dimmed white(#E2E8F0) 避免眩光

---

## 5. 色彩对比度审计

color_guard.py 实现:

### 检测维度
1. shape_fill_text: 形状填充上的文本对比度
2. textbox: 文本框文本与背景对比度
3. bg_bg_contrast: 图表/表格背景与幻灯片背景反差(INFO)

### 对比度计算
WCAG 2.1 标准:
- 相对亮度: L = 0.2126*R + 0.7152*G + 0.0722*B
- 对比度: (L1 + 0.05) / (L2 + 0.05)
- CRITICAL: < 2.0 | WARN: < 4.5 | INFO: >= 4.5

### 背景采样
- 文本形状中心点落在另一形状内 -> 使用该形状填充作为背景
- 否则使用幻灯片级背景色

---

## 6. 文本自适应

text_fitter.py 使用 Pillow 实现:
1. 加载字体(Noto Sans CJK / DejaVu)
2. 二分搜索最大可用字号:
   low=1, high=初始字号
   mid = (low + high) // 2
   Pillow 测量文本在 mid 字号下的宽高
   若超出边界 -> high = mid - 1
   否则 -> low = mid + 1
3. 返回最大不溢出字号
4. 实际渲染用 Microsoft YaHei，宽度微调系数 1.04

---

## 7. 声明式布局 Schema

layout_schema.py 核心概念:

### Region 系统
- 百分比: 5%, left-half, top-third, center
- 绝对单位: 2in, 5cm, 100px
- 原始 EMU: 5000000

### Attach 系统
相对定位，避免硬编码:
{ attach: { to: title, edge: bottom, offset: 1% } }

### Adaptive 策略
- shrink: 缩小字号直到 fit
- expand: 扩展形状高度
- truncate: 截断文本

---

## 8. 意图路由

intent_router.py 根据内容字段推断布局:

匹配规则示例:
- {quote, author} -> quote
- {big_number, label} -> data_highlight
- {left_content, right_content} -> two_column
- {title, subtitle} -> cover
- {chart_data} -> chart
- {table_data} -> table

按优先级顺序匹配，若 slide_spec 包含某规则的所有 key，
则分配对应 layout。

---

## 9. 模板分析

template_analyzer.py 提取视觉契约:
1. 遍历模板所有幻灯片
2. 提取形状类型、位置、尺寸、填充色、字体、字号
3. 识别占位符文本模式(__TITLE__ 等)
4. 输出 contract.json(颜色、字体、布局结构、占位符映射)

contract_to_schema.py 转换为声明式 Schema:
- 计算相对坐标和百分比
- 生成 attach 关系
- 输出 JSON Schema

---

## 10. 渲染与导出

### 10.1 渲染

render_slides.py 两种方式:

LibreOffice 模式(默认):
- soffice --headless --convert-to pdf
- pdftoppm 或 PyMuPDF PDF -> PNG
- 优点: 像素级精确，字体渲染一致
- 缺点: 需 LibreOffice，首次加载慢

软渲染(--soft):
- python-pptx 读取形状文本和颜色
- Pillow 绘制到图像
- 优点: 纯 Python，速度快
- 缺点: 近似渲染，复杂图表可能失真

### 10.2 导出
- PDF: soffice 转换
- 长图: 垂直拼接所有幻灯片 PNG
- HTML: reveal.js 风格幻灯片

---

## 12. 约束求解系统

### 12.1 设计目标

"字体尽量大，边框自动编排"——约束求解系统的核心目标。

### 12.2 架构

```
constraint_layout.py
├── _pptx_wrap_text()          # Pillow 文本换行模拟
├── _measure_required_height() # 计算文本所需高度
├── _truncate_text_to_height() # 溢出截断
├── _expand_font_size()        # 填充可用空间
├── solve_constraints()        # 主求解器
│   ├── 1. 解析 schema shapes
│   ├── 2. 按 attach 拓扑排序
│   ├── 3. 收集 textboxes + images
│   ├── 4. 二分搜索 best_base
│   ├── 5. 计算实际高度 + expand/shrink
│   ├── 6. 应用 region + attach 链
│   └── 7. 返回 (region, style) 列表
└── _compute_image_height()    # 图片宽高比保持
```

### 12.3 二分搜索算法

```python
lo, hi = 8, 96
best_base = 8
while lo <= hi:
    mid = (lo + hi) // 2
    if _fits(mid):      # 所有 textbox 都能 fit?
        best_base = mid
        lo = mid + 1    # 尝试更大
    else:
        hi = mid - 1    # 太大，缩小
```

`_fits(base_size)` 检查每个 textbox：
- 实际字号 = base_size × font_scale
- 计算换行后的文本高度
- 若高度 ≤ 可用高度 → fit
- 若单行模式，还需检查宽度

### 12.4 Attach 链计算

```python
# 拓扑排序：被 attach 的目标排在前面
sorted_shapes = topological_sort(shapes, attach_edges)

# 逐个计算 region
for shape in sorted_shapes:
    if shape.attach:
        parent = shape.attach.to
        parent_region = computed[parent]
        shape.region.top = parent_region.top + parent_region.height + offset
```

支持 8 个方向：`top`, `bottom`, `left`, `right`, `top-left`, `top-right`, `bottom-left`, `bottom-right`。

### 12.5 Expand 策略

当文本高度 < 区域高度的 50% 时，尝试增大字号直到填充 60%：

```python
if h < max_height * 0.50:
    expanded = _expand_font_size(text, width, max_height, current_size)
    actual_size = expanded
    h = measure(text, expanded, width)
```

### 12.6 Shrink 策略

当 best_base < 14pt 时，自动降级：

1. **行距压缩**：1.40 → 1.20（+14% 垂直空间）
2. **字号比例压缩**：所有 font_scale × 0.85
3. **最终警告**：若仍 < 10pt，提示用户精简内容

```python
if best_base < 14:
    best_base_tight = _run_solver(line_spacing=1.20)
    if best_base_tight > best_base:
        best_base = best_base_tight
    if best_base < 10:
        # 压缩 font_scale 再试
        shrunk_textboxes = [tb.font_scale *= 0.85 for tb in textboxes]
        best_base_shrunk = _run_solver(shrunk_textboxes, line_spacing=1.20)
```

### 12.7 主题协调

theme_engine.py 在应用主题时，会检测 textbox 下方的 auto_shape 填充色：

```python
def _detect_underlying_fill(textbox, slide):
    # 检查所有与 textbox 重叠的 auto_shape
    # 返回面积最大的那个的 fill 颜色
    # 基于该 fill 的亮度选择白字或黑字
```

这解决了 colored shapes（如 process 卡片）上的白字被主题全局改深色的问题。

---

## 13. 扩展指南

### 11.1 添加新布局

1. 在 base_template.pptx 中创建模板幻灯片
2. 在 build_pptx.py 中添加 elif layout_type == xxx 分支
3. 实现渲染函数(参考现有布局)
4. 在 layouts.md 中补充文档
5. (可选) 创建 layout_schemas/xxx.json 声明式 Schema

### 11.2 添加新主题

1. 在 references/themes/ 下创建 JSON 文件
2. 包含字段: name, primary, secondary, accent[], background, text, light_bg, font_title, font_body
3. 使用 --theme /path/to/theme.json 测试

### 11.3 添加新图表类型

1. 在 advanced_charts.py 中实现绘制函数
2. 在 build_pptx.py 的 chart 布局分支中增加 type 判断
3. 使用 matplotlib + python-pptx 插入图片方式嵌入

---

*文档版本: 2026-07-01*