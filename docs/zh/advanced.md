> 本文档为中文版。

# 高级功能

## 富文本

任何接受字符串的文本字段也可以接受 **run dicts** 数组，以实现混合格式。

```json
{
  "content": [
    {"text": "Rust", "bold": true, "color": "#3B82F6", "size": 24},
    {"text": "速度极快。"}
  ]
}
```

**Run dict 字段：**

| 字段 | 类型 | 说明 |
|-------|------|-------------|
| `text` | string | 文本内容（必填） |
| `bold` | bool | 粗体 |
| `italic` | bool | 斜体 |
| `underline` | bool | 下划线 |
| `color` | string | 十六进制颜色（`#RRGGBB`） |
| `size` | number | 字号，单位为磅 |
| `font` | string | 字体名称 |

> **提示：** 当 run 的文本中出现 `\n` 时，会自动创建一个新段落。这对于在形状中垂直居中非常有用。

---

## 自由形状

任何幻灯片都可以包含 `shapes` 数组，用于精确定位自定义元素。

### 形状类型

#### 文本框

```json
{
  "type": "text",
  "left": 1000000,
  "top": 2000000,
  "width": 3000000,
  "height": 1000000,
  "content": "你好，世界",
  "font_size": 18,
  "color": "#374151",
  "alignment": "center",
  "vertical_center": true
}
```

#### 图片

```json
{
  "type": "image",
  "path": "photo.jpg",
  "left": 5000000,
  "top": 2000000,
  "width": 2000000,
  "height": 1500000,
  "mode": "fit"
}
```

`mode`：`"fit"`（保持纵横比，留黑边）或 `"stretch"`（完全填充）。

#### 形状（矩形、椭圆形等）

```json
{
  "type": "shape",
  "shape_type": "rounded_rectangle",
  "left": 1000000,
  "top": 4000000,
  "width": 2000000,
  "height": 1000000,
  "fill": "#3B82F6",
  "border_color": "#1E3A5F",
  "border_width": 2,
  "content": "点击我",
  "alignment": "center",
  "vertical_center": true
}
```

**可用的 shape_types：**

- `rectangle`
- `rounded_rectangle`
- `oval`
- `triangle`
- `diamond`
- `pentagon`
- `hexagon`
- `star`
- `arrow_right`, `arrow_left`, `arrow_up`, `arrow_down`
- `chevron`
- `parallelogram`
- `trapezoid`
- `donut`

#### 图表

```json
{
  "type": "chart",
  "left": 1000000,
  "top": 4000000,
  "width": 8000000,
  "height": 4000000,
  "chart_data": {
    "type": "bar",
    "categories": ["A", "B", "C"],
    "values": [10, 20, 30]
  }
}
```

### 通用形状字段

所有形状都支持：

| 字段 | 类型 | 说明 |
|-------|------|-------------|
| `left` | int | X 坐标，单位为 EMU |
| `top` | int | Y 坐标，单位为 EMU |
| `width` | int | 宽度，单位为 EMU |
| `height` | int | 高度，单位为 EMU |
| `z_order` | int | 堆叠顺序（数值越大越靠前） |

---

## 图片遮罩

将图片遮罩成某种形状（圆形、圆角矩形等）：

```json
{
  "type": "image",
  "path": "avatar.jpg",
  "mask": "circle",
  "left": 1000000,
  "top": 1000000,
  "width": 1500000,
  "height": 1500000
}
```

**可用的遮罩：**

- `circle` / `oval`
- `rounded_rectangle`
- `rectangle`
- `triangle`
- `diamond`
- `hexagon`
- `star`
- `pentagon`
- `heart`
- `cloud`
- `sun`
- `moon`

---

## 图片裁剪

使用比例（0.0–1.0）从任意边缘裁剪图片：

```json
{
  "type": "image",
  "path": "photo.jpg",
  "crop": {
    "left": 0.1,
    "right": 0.1,
    "top": 0,
    "bottom": 0
  }
}
```

每个值表示从对应边缘裁剪掉的图片尺寸比例。

---

## Z-Order（图层控制）

控制哪些形状显示在其他形状之上：

```json
{
  "type": "shape",
  "shape_type": "rectangle",
  "left": 0,
  "top": 0,
  "width": 10000000,
  "height": 5000000,
  "fill": "#F3F4F6",
  "z_order": 0
},
{
  "type": "image",
  "path": "overlay.png",
  "z_order": 5
}
```

- `z_order: 0` = 底层（最下方图层）
- 数值越大 = 越靠前（最上方图层）

---

## 强调线开关

标题下方的蓝色下划线装饰可以按幻灯片禁用：

```json
{
  "layout": "title_content",
  "title": "简洁标题",
  "accent_line": false,
  "content": "此幻灯片无下划线。"
}
```

默认值为 `true`（显示强调线）。

---

## 自动适配系统

构建器使用 Pillow 测量文本，并通过二分查找找到能适配形状的最大字号。该功能会自动应用于：

- 标题和副标题
- 内容文本框
- 表格单元格
- 时间轴/流程/团队文本框
- data_highlight 中的大数字

在自定义脚本中，可以通过不向 `replace_placeholder_text` 传递 `auto_fit=True` 来禁用自动适配。

---

## 主题

通过命令行切换配色方案：

```bash
python scripts/build_pptx.py deck.json output.pptx --theme ocean_tech
```

内置主题：

| 主题 | 说明 |
|-------|-------------|
| `default` | 藏青色 + 白色 |
| `ocean` | 蓝绿色 + 浅蓝色 |
| `sunset` | 橙色 + 珊瑚色 |
| `forest` | 绿色 + 鼠尾草色 |
| `berry` | 紫色 + 洋红色 |
| `monochrome` | 灰度 |

自定义主题可以以 JSON 文件形式添加到 `references/themes/`，需包含 `primary`、`secondary`、`accent`、`dark`、`light`、`background`、`text` 颜色字段。

---

## 导出格式

### PDF

```bash
python scripts/export.py deck.pptx --format pdf --output deck.pdf
```

需要 LibreOffice（`soffice`）。若不可用，则回退到文本报告。

### 长图

将所有幻灯片垂直拼接成一张 PNG：

```bash
python scripts/export.py deck.pptx --format long_image --output deck.png
```

### HTML 幻灯片

简单的 reveal.js 风格 HTML：

```bash
python scripts/export.py deck.pptx --format html --output deck.html
```

---

## 截图工具

截取特定页面或区域：

```bash
# Full page
python scripts/snapshot.py deck.pptx --page 3 --output page3.png

# Pixel crop (x, y, w, h)
python scripts/snapshot.py deck.pptx --page 3 --crop "200,300,800,600" --output clip.png

# EMU crop (same unit as shapes)
python scripts/snapshot.py deck.pptx --page 3 --crop-emu "500000,1000000,2000000,1500000" --output clip.png

# List page dimensions
python scripts/snapshot.py deck.pptx --list
```

---

## EMU 坐标系

`shapes` 中的所有位置均使用 **EMU**（English Metric Units，英制公制单位）：

- `1 inch = 914,400 EMU`
- `1 cm = 360,000 EMU`
- 标准 16:9 幻灯片：`12,192,000 × 6,858,000 EMU`

16:9 幻灯片快速参考：

| 位置 | EMU |
|----------|-----|
| 左边缘 | 0 |
| 右边缘 | 12,192,000 |
| 上边缘 | 0 |
| 下边缘 | 6,858,000 |
| 中心 X | 6,096,000 |
| 中心 Y | 3,429,000 |
