> 本文档为中文版。

# PPT Skill vs 开源项目对比分析

> 对比 ppt-skill 与相关开源项目，分析定位、能力边界与适用场景。

---

## 对比项目一览

| 项目 | 语言 | 定位 | 核心输出 |
|------|------|------|---------|
| **ppt-skill (ours)** | Python | JSON 驱动原生 PPTX 生成 | 可编辑 .pptx |
| PPTAgent / DeepPresenter | Python | AI Agent 全流程 PPT 生成 | 可编辑 .pptx |
| PPTskill | Python | SVG -> PPTX 转换 | 可编辑 .pptx |
| ChatPPT-MCP | Python/Node | 在线 PPT 生成服务 | .pptx / 在线 |
| pptx-tools | C#/.NET | MCP Server for PowerPoint | 编辑 .pptx |
| ppt-agent-skills | Python | HTML PPT Agent 工作流 | HTML / PNG |

---

## 1. PPTAgent (DeepPresenter)

**仓库**: icip-cas/PPTAgent | **论文**: EMNLP 2025 + ACL 2026
**定位**: 全自动 AI Agent PPT 生成系统

### 核心能力
- 两阶段编辑式生成: 分析参考 PPT -> 提取模式 -> 迭代编辑
- 支持 PDF/Word/URL/Markdown 多源输入
- 内置 PDF 解析(MinerU)、网页搜索(Tavily)、文生图
- 自有微调模型 DeepPresenter-9B
- Web UI + Docker Compose 部署
- MCP Server 支持

### 与 ppt-skill 对比

| 维度 | PPTAgent | ppt-skill |
|------|----------|-----------|
| 输入方式 | 自然语言/文档 | JSON 规格 |
| 自动化程度 | 全自动(Agent) | 半自动(JSON) |
| 内容生成 | 联网搜索+LLM | 用户提供 |
| 可控性 | 中等 | **高** |
| 部署复杂度 | 高(Docker+模型) | **低**(纯 Python) |
| 离线能力 | 需服务 | **完全离线** |
| 中文字体 | 一般 | **原生支持** |
| 预览渲染 | 无 | **LibreOffice+Pillow** |
| 对比度审计 | 无 | **color_guard** |
| 动画 | 无 | **8 种** |
| 图表类型 | 基础 | **9+4 种** |

### 优劣分析

**PPTAgent 优势:**
- 真正的零输入生成: 给一个主题即可自动调研生成
- 学术级评估体系 PPTEval(Content/Design/Coherence)
- 支持参考 PPT 风格迁移
- 多模态: 文生图、PDF 解析

**PPTAgent 劣势:**
- 部署复杂: Docker + 模型 + 搜索 API
- 可控性弱: Agent 决策黑盒
- 成本高: LLM API + 搜索 API
- 中文支持有限，无内置预览

**ppt-skill 优势:**
- 精确控制: JSON 精确到每个字、每个颜色
- 轻量部署: 纯 Python，无外部依赖
- 中文优化: Microsoft YaHei + Noto Sans CJK
- 质量保障: color_guard + QA 检查
- 实时预览: LibreOffice 高保真渲染
- 丰富图表: 18 种布局 + 13 种图表

**ppt-skill 劣势:**
- 不提供自动内容生成
- 无 LLM Agent 能力

### 适用场景
- **PPTAgent**: 快速原型、AI 辅助调研
- **ppt-skill**: 内容确定、精确控制、批量生成

---

## 2. PPTskill

**仓库**: AIPMAndy/PPTskill (上游 hugohe3/ppt-master)
**定位**: 基于 SVG 的咨询风格 PPT 生成器

### 核心能力
- SVG -> 原生 PPTX 转换(真实 DrawingML)
- 10+ 专业模板(麦肯锡风、学术风、杂志风)
- 支持 PDF/DOCX/URL/Markdown 输入
- OpenClaw/Cursor 集成

### 与 ppt-skill 对比

| 维度 | PPTskill | ppt-skill |
|------|----------|-----------|
| 设计方式 | SVG 绘制 | python-pptx 原生 |
| 模板数量 | 10+ 专业风格 | 1 基础 + 8 主题 |
| 输入方式 | 文档/Markdown | JSON |
| 布局灵活性 | 模板固定 | **JSON 自由定义** |
| 图表 | 无原生 | **9+4 种** |
| 动画 | 无 | **8 种** |
| 主题系统 | 模板级切换 | **运行时换肤** |
| 渲染预览 | 无 | **双模式** |
| 质量审计 | 无 | **color_guard + QA** |

### 优劣分析

**PPTskill 优势:**
- 设计模板更精美(咨询风、杂志风)
- SVG 转换保证形状精确
- 支持多种文档输入

**PPTskill 劣势:**
- 模板固定，难自定义布局
- 无原生图表、动画、预览、质量检查
- 依赖 cairo/pycairo 安装复杂

### 适用场景
- **PPTskill**: 精美咨询风格、文档快速转换
- **ppt-skill**: 数据可视化、精确控制、批量生成

---

## 3. ChatPPT-MCP

**定位**: 在线 PPT 生成服务(通过 MCP 协议接入)

### 核心能力
- MCP 协议的 PPT 生成服务
- 主题生成、文档导入、在线编辑
- 18 种智能文档处理 API
- Streamable HTTP / Python / Node.js 接入
- 商业服务(非完全开源)

### 与 ppt-skill 对比

| 维度 | ChatPPT-MCP | ppt-skill |
|------|-------------|-----------|
| 部署方式 | 云服务 | **本地/私有** |
| 开源程度 | 部分开源 | **完全开源** |
| 隐私安全 | 需上传数据 | **完全本地** |
| 定制化 | 受限于 API | **完全可控** |
| 成本 | 按调用付费 | **免费** |
| 离线能力 | 无 | **完全离线** |

### 优劣分析

**ChatPPT-MCP 优势:**
- 即开即用，无需配置
- 持续更新，多客户端支持

**ChatPPT-MCP 劣势:**
- 商业服务，数据隐私风险
- 无法本地部署，定制能力受限

### 适用场景
- **ChatPPT-MCP**: 个人快速生成、无技术团队
- **ppt-skill**: 企业内网、数据敏感、深度定制

---

## 4. pptx-tools

**仓库**: jongalloway/pptx-tools
**定位**: .NET MCP Server for PowerPoint 编辑

### 核心能力
- 22 个 MCP Tools 操作 PPTX
- 读取/更新文本、插入图片、创建表格、更新图表
- 幻灯片管理(增删改查、重排序)
- 媒体管理(去重、压缩、提取)
- 完全不生成新 PPTX，只编辑现有文件

### 与 ppt-skill 对比

| 维度 | pptx-tools | ppt-skill |
|------|------------|-----------|
| 核心能力 | 编辑现有 PPTX | **从零生成** |
| 技术栈 | C#/.NET | Python |
| 图表 | 更新数据(保留样式) | **从零创建** |
| 布局 | 基于现有模板 | **可新建** |
| 自动化 | Agent 驱动 | 脚本驱动 |

### 优劣分析

**pptx-tools 优势:**
- 完美的 PPTX 编辑能力(不破坏格式)
- MCP 协议标准化，AI Agent 友好
- 图表数据更新保留所有样式

**pptx-tools 劣势:**
- 不能从零创建 PPTX
- 不能修改母版/主题样式
- .NET 技术栈对 Python 生态不友好
- 无渲染预览能力

### 适用场景
- **pptx-tools**: 已有模板，AI 填充数据(如 QBR)
- **ppt-skill**: 从零创建、自定义布局

---

## 5. ppt-agent-skills

**定位**: HTML 格式演示文稿的 Agent 工作流

### 核心能力
- 模拟顶级 PPT 设计公司工作流
- 需求调研 -> 资料搜集 -> 大纲 -> 设计稿
- 输出 HTML 格式(非 PPTX)
- 多阶段 Subagent 协作
- 人工审计断点
- 每页并行生成

### 与 ppt-skill 对比

| 维度 | ppt-agent-skills | ppt-skill |
|------|------------------|-----------|
| 输出格式 | HTML/PNG/SVG | **原生 PPTX** |
| 可编辑性 | 不可编辑 | **完全可编辑** |
| 内容生成 | LLM 全自动 | 用户提供 |
| 工作流 | 复杂多阶段 Agent | **单步 JSON->PPTX** |
| 部署 | 需 LLM + Subagent | **纯 Python** |
| 设计自由度 | 极高(HTML/CSS) | 受限于 PPTX |
| 图表 | HTML 图表 | **原生 PowerPoint** |

### 优劣分析

**ppt-agent-skills 优势:**
- 设计自由度极高(HTML/CSS)
- 全自动内容生成
- 支持人工审计

**ppt-agent-skills 劣势:**
- 输出不是 PPTX，无法编辑
- 工作流复杂，出错恢复困难
- 依赖 LLM API，成本高
- 无原生图表交互

### 适用场景
- **ppt-agent-skills**: 网页演示、追求极致设计
- **ppt-skill**: 原生 PPTX、数据驱动、可编辑

---

## 综合对比矩阵

| 能力 | ppt-skill | PPTAgent | PPTskill | ChatPPT | pptx-tools | ppt-agent |
|------|:---------:|:--------:|:--------:|:-------:|:----------:|:---------:|
| 原生 PPTX | **是** | 是 | 是 | 是 | 编辑 | 否 |
| 零输入生成 | 否 | **是** | 部分 | **是** | 否 | **是** |
| JSON 驱动 | **是** | 否 | 否 | 否 | 否 | 否 |
| 主题换肤 | **是** | 有限 | 模板级 | 有限 | 否 | CSS |
| 原生图表 | **9+4** | 基础 | 无 | 基础 | 更新 | HTML |
| 动画 | **8种** | 无 | 无 | 有限 | 否 | 否 |
| 渲染预览 | **双模式** | 无 | 无 | 在线 | 否 | 否 |
| 对比度审计 | **是** | 无 | 无 | 无 | 否 | 否 |
| 文本自适应 | **是** | 无 | 无 | 无 | 否 | 否 |
| 自由形状 | **是** | 无 | 无 | 无 | 否 | 是 |
| 中文字体 | **原生** | 一般 | 一般 | 一般 | 一般 | 一般 |
| 离线部署 | **是** | 否 | 是 | 否 | 是 | 否 |
| 开源免费 | **是** | 是 | 是 | 部分 | 是 | 是 |

---

## 总结

### ppt-skill 的独特定位

ppt-skill 在生态中的独特价值在于 **精确可控的 JSON -> 原生 PPTX 流水线**:

1. **确定性**: JSON 规格保证输出完全可预测，适合自动化流水线
2. **原生可编辑**: 输出真实 PPTX，非图片/HTML，保留完整编辑能力
3. **质量保障**: color_guard 对比度审计 + QA 检查，避免视觉灾难
4. **中文优化**: Microsoft YaHei + Noto Sans CJK，中文显示专业
5. **轻量部署**: 纯 Python，无需 Docker/模型/API，完全离线
6. **丰富图表**: 18 种布局 + 9 种原生图表 + 4 种高级图表

### 与其他项目的协作关系

```
[PPTAgent] ----(内容生成)----> [JSON] ----(ppt-skill)----> [PPTX]
                                   |
[ChatPPT] ----(在线服务)----------|
                                   |
[pptx-tools] <---(编辑数据)-------|
```

- 可用 PPTAgent 生成内容大纲，输出为 JSON，再用 ppt-skill 精确控制视觉
- 可用 ChatPPT 快速生成初稿，导出 JSON 后用 ppt-skill 精细化调整
- 可用 ppt-skill 生成基础 PPTX，再用 pptx-tools 批量更新数据
- 可用 ppt-agent-skills 做网页演示，同时用 ppt-skill 生成可编辑版本

### 选型建议

| 场景 | 推荐方案 |
|------|---------|
| 需要 AI 自动调研生成 | PPTAgent |
| 精美咨询风格、文档转换 | PPTskill |
| 快速在线生成、无技术团队 | ChatPPT-MCP |
| 已有模板、批量更新数据 | pptx-tools |
| 网页演示、极致设计 | ppt-agent-skills |
| **精确控制、数据驱动、批量生成、中文场景** | **ppt-skill** |

---

*文档版本: 2026-07-01*