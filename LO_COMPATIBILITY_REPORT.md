# PPT Skill LibreOffice 兼容性改进报告

## 1. 环境检查结果

- **LibreOffice 版本**: 7.4.7.2（可用）
- **pdftoppm**: 未安装
- **中文字体**: 系统未安装任何中文字体

## 2. 排版错位的根因分析

### 2.1 字体缺失（主因）

`build_pptx.py` 中所有文本统一使用 `"Microsoft YaHei"` 字体。LibreOffice 在 Linux 容器中找不到该字体，回退到不含 CJK 字形的默认字体，导致：
- 所有中文显示为方块（□）
- 文本框实际占位宽度/高度与 PowerPoint 中完全不同
- 下方/右侧元素（包括图片）被错误地推动或重叠

### 2.2 图片插入行为不一致

原 `add_image_placeholder()` 使用：
```python
slide.shapes.add_picture(image_path, x, y, cx, cy)
```
同时传入宽高，python-pptx 会**拉伸**图片到精确尺寸。这与 PowerPoint 中常见的"保持比例、居中适应"行为不同，导致：
- 非 4:3 或 16:9 的图片被压扁/拉长
- 在 LibreOffice 中观察到的"图片被放大"现象

### 2.3 Soft Preview 未真实渲染图片

原 `render_slides.py` 的 fallback 模式只画灰色占位符和交叉线，无法验证图片的实际排版效果。

## 3. 已完成的代码改进

### 3.1 build_pptx.py — 图片保持比例

`add_image_placeholder()` 新增 `mode` 参数：
- `"fit"`（默认）：保持宽高比，缩放至适应 placeholder 并居中
- `"stretch"`：原行为，拉伸到精确尺寸

```python
def add_image_placeholder(slide, image_path, placeholder_shape, mode="fit"):
    ...
```

### 3.2 render_slides.py — 真实图片渲染 + LibreOffice 增强

1. **Soft Preview 真正渲染图片**：读取 `shape.image.blob`，用 Pillow 粘贴到画布上
2. **LibreOffice 路径增强**：
   - 设置 `HOME` 和 `UserInstallation` 为可写目录，避免容器内启动失败
   - 增加 `subprocess.run(timeout=60)` 防止卡死
   - `pdftoppm` 不可用时，自动 fallback 到 PyMuPDF 将 PDF 转 PNG
3. **主逻辑调整**：不再强制要求 `pdftoppm`，只要 LibreOffice 可用即尝试渲染

## 4. 对比验证

### 4.1 改进前（LibreOffice 渲染）
- 标题：□□□□□□□□（方块）
- 图片：被拉伸，位置因文本框塌陷而偏移

### 4.2 改进后（Soft Preview，不依赖 LibreOffice）
- 标题：正常显示
- 图片：保持原始比例，居中适应 placeholder

### 4.3 LibreOffice 仍受限
即使代码改进后，LibreOffice 渲染仍有方块字问题——这是**环境字体缺失**导致，非代码可完全解决。

## 5. 部署建议

### 方案 A：安装中文字体（推荐，解决 LibreOffice 错位）

在容器/服务器上执行：
```bash
# Debian/Ubuntu
apt-get update && apt-get install -y fonts-noto-cjk

# 或手动复制项目自带字体到系统目录
sudo mkdir -p /usr/share/fonts/opentype/noto
sudo cp fonts/NotoSansCJKsc-Regular.otf /usr/share/fonts/opentype/noto/
sudo fc-cache -fv
```

### 方案 B：使用 Soft Preview 替代 LibreOffice（已增强）

优点：
- 不依赖外部字体配置
- 渲染结果可控、可预测
- 速度快，无进程卡死风险

缺点：
- 图表、表格为简化渲染（非真实图表）
- 复杂形状/渐变/动画不支持

### 方案 C：统一使用无衬线西文字体（应急）

如果只需要英文/数字内容，可将 `build_pptx.py` 中的 `"Microsoft YaHei"` 替换为 `"DejaVu Sans"`（LibreOffice 自带），可完全避免错位。

## 6. 后续可优化项

1. **字体嵌入 PPTX**：通过修改 OPC/XML 将 `NotoSansCJKsc-Regular.otf` 嵌入到生成的 PPTX 中，确保任何打开该文件的环境都能正确显示。
2. **Soft Preview 图表渲染**：用 matplotlib 在 fallback 模式下生成真实图表图片，替代 `[Chart]` 占位符。
3. **文本框 auto_size 锁定**：在 `build_pptx.py` 中设置 `shape.text_frame.auto_size = None`，防止不同渲染器对文本框高度的计算差异影响布局。
