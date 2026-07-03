# PPT Skill vs. Open Source Projects: Comparative Analysis

> Compare ppt-skill with related open source projects, analyzing positioning, capability boundaries, and applicable scenarios.

---

## Comparison Overview

| Project | Language | Positioning | Core Output |
|---------|----------|-------------|-------------|
| **ppt-skill (ours)** | Python | JSON-driven native PPTX generation | Editable .pptx |
| PPTAgent / DeepPresenter | Python | AI Agent full-process PPT generation | Editable .pptx |
| PPTskill | Python | SVG -> PPTX conversion | Editable .pptx |
| ChatPPT-MCP | Python/Node | Online PPT generation service | .pptx / online |
| pptx-tools | C#/.NET | MCP Server for PowerPoint | Edit .pptx |
| ppt-agent-skills | Python | HTML PPT Agent workflow | HTML / PNG |

---

## 1. PPTAgent (DeepPresenter)

**Repository**: icip-cas/PPTAgent | **Papers**: EMNLP 2025 + ACL 2026
**Positioning**: Fully automated AI Agent PPT generation system

### Core Capabilities

- Two-stage editing-based generation: analyze reference PPTs -> extract patterns -> iterative editing
- Supports multi-source input: PDF / Word / URL / Markdown
- Built-in PDF parsing (MinerU), web search (Tavily), text-to-image
- Proprietary fine-tuned model DeepPresenter-9B
- Web UI + Docker Compose deployment
- MCP Server support

### Comparison with ppt-skill

| Dimension | PPTAgent | ppt-skill |
|-----------|----------|-----------|
| Input Method | Natural language / documents | JSON spec |
| Automation Level | Fully automated (Agent) | Semi-automated (JSON) |
| Content Generation | Web search + LLM | User-provided |
| Controllability | Medium | **High** |
| Deployment Complexity | High (Docker + model) | **Low** (pure Python) |
| Offline Capability | Requires services | **Fully offline** |
| Chinese Fonts | Average | **Native support** |
| Preview Rendering | None | **LibreOffice + Pillow** |
| Contrast Audit | None | **color_guard** |
| Animations | None | **8 types** |
| Chart Types | Basic | **9 + 4 types** |

### Pros and Cons Analysis

**PPTAgent Advantages:**

- True zero-input generation: provide a topic and it automatically researches and generates
- Academic-grade evaluation system PPTEval (Content / Design / Coherence)
- Supports reference PPT style transfer
- Multimodal: text-to-image, PDF parsing

**PPTAgent Disadvantages:**

- Complex deployment: Docker + model + search API
- Weak controllability: Agent decision-making is a black box
- High cost: LLM API + search API
- Limited Chinese support, no built-in preview

**ppt-skill Advantages:**

- Precise control: JSON precise down to every character and color
- Lightweight deployment: pure Python, no external dependencies
- Chinese optimization: Microsoft YaHei + Noto Sans CJK
- Quality assurance: color_guard + QA checks
- Real-time preview: LibreOffice high-fidelity rendering
- Rich charts: 18 layouts + 13 chart types

**ppt-skill Disadvantages:**

- Does not provide automatic content generation
- No LLM Agent capabilities

### Applicable Scenarios

- **PPTAgent**: Rapid prototyping, AI-assisted research
- **ppt-skill**: Fixed content, precise control, batch generation

---

## 2. PPTskill

**Repository**: AIPMAndy/PPTskill (upstream hugohe3/ppt-master)
**Positioning**: SVG-based consulting-style PPT generator

### Core Capabilities

- SVG -> native PPTX conversion (real DrawingML)
- 10+ professional templates (McKinsey style, academic style, magazine style)
- Supports PDF / DOCX / URL / Markdown input
- OpenClaw / Cursor integration

### Comparison with ppt-skill

| Dimension | PPTskill | ppt-skill |
|-----------|----------|-----------|
| Design Approach | SVG drawing | python-pptx native |
| Template Count | 10+ professional styles | 1 basic + 8 themes |
| Input Method | Documents / Markdown | JSON |
| Layout Flexibility | Fixed templates | **Freely defined in JSON** |
| Charts | No native charts | **9 + 4 types** |
| Animations | None | **8 types** |
| Theme System | Template-level switching | **Runtime skinning** |
| Rendered Preview | None | **Dual mode** |
| Quality Audit | None | **color_guard + QA** |

### Pros and Cons Analysis

**PPTskill Advantages:**

- More refined design templates (consulting style, magazine style)
- SVG conversion ensures precise shapes
- Supports multiple document inputs

**PPTskill Disadvantages:**

- Fixed templates, difficult to customize layouts
- No native charts, animations, preview, or quality checks
- Complex installation due to cairo / pycairo dependency

### Applicable Scenarios

- **PPTskill**: Refined consulting style, rapid document conversion
- **ppt-skill**: Data visualization, precise control, batch generation

---

## 3. ChatPPT-MCP

**Positioning**: Online PPT generation service (accessed via MCP protocol)

### Core Capabilities

- PPT generation service using MCP protocol
- Topic generation, document import, online editing
- 18 intelligent document processing APIs
- Streamable HTTP / Python / Node.js access
- Commercial service (not fully open source)

### Comparison with ppt-skill

| Dimension | ChatPPT-MCP | ppt-skill |
|-----------|-------------|-----------|
| Deployment | Cloud service | **Local / private** |
| Openness | Partially open source | **Fully open source** |
| Privacy & Security | Requires data upload | **Fully local** |
| Customization | Limited by API | **Fully controllable** |
| Cost | Pay per call | **Free** |
| Offline Capability | None | **Fully offline** |

### Pros and Cons Analysis

**ChatPPT-MCP Advantages:**

- Ready to use out of the box, no configuration needed
- Continuous updates, multi-client support

**ChatPPT-MCP Disadvantages:**

- Commercial service, data privacy risks
- Cannot be deployed locally, customization limited

### Applicable Scenarios

- **ChatPPT-MCP**: Personal rapid generation, no technical team
- **ppt-skill**: Enterprise intranet, data-sensitive, deep customization

---

## 4. pptx-tools

**Repository**: jongalloway/pptx-tools
**Positioning**: .NET MCP Server for PowerPoint editing

### Core Capabilities

- 22 MCP Tools to manipulate PPTX
- Read / update text, insert images, create tables, update charts
- Slide management (CRUD, reordering)
- Media management (deduplication, compression, extraction)
- Does not generate new PPTX at all, only edits existing files

### Comparison with ppt-skill

| Dimension | pptx-tools | ppt-skill |
|-----------|------------|-----------|
| Core Capability | Edit existing PPTX | **Generate from scratch** |
| Tech Stack | C# / .NET | Python |
| Charts | Update data (keep styles) | **Create from scratch** |
| Layout | Based on existing templates | **Can create new layouts** |
| Automation | Agent-driven | Script-driven |

### Pros and Cons Analysis

**pptx-tools Advantages:**

- Perfect PPTX editing capability (does not break formatting)
- MCP protocol standardization, AI Agent friendly
- Chart data updates preserve all styles

**pptx-tools Disadvantages:**

- Cannot create PPTX from scratch
- Cannot modify master / theme styles
- .NET tech stack is unfriendly to the Python ecosystem
- No rendered preview capability

### Applicable Scenarios

- **pptx-tools**: Existing templates, AI fills data (e.g., QBR)
- **ppt-skill**: Create from scratch, custom layouts

---

## 5. ppt-agent-skills

**Positioning**: Agent workflow for HTML-format presentations

### Core Capabilities

- Simulates top-tier PPT design company workflows
- Requirements research -> material collection -> outline -> design draft
- Outputs HTML format (not PPTX)
- Multi-stage Subagent collaboration
- Human audit breakpoints
- Per-page parallel generation

### Comparison with ppt-skill

| Dimension | ppt-agent-skills | ppt-skill |
|-----------|------------------|-----------|
| Output Format | HTML / PNG / SVG | **Native PPTX** |
| Editability | Not editable | **Fully editable** |
| Content Generation | Fully automated by LLM | User-provided |
| Workflow | Complex multi-stage Agent | **Single-step JSON -> PPTX** |
| Deployment | Requires LLM + Subagent | **Pure Python** |
| Design Freedom | Very high (HTML / CSS) | Limited by PPTX |
| Charts | HTML charts | **Native PowerPoint** |

### Pros and Cons Analysis

**ppt-agent-skills Advantages:**

- Extremely high design freedom (HTML / CSS)
- Fully automatic content generation
- Supports human audit

**ppt-agent-skills Disadvantages:**

- Output is not PPTX, cannot be edited
- Complex workflow, difficult error recovery
- Depends on LLM API, high cost
- No native chart interactivity

### Applicable Scenarios

- **ppt-agent-skills**: Web presentations, pursuit of ultimate design
- **ppt-skill**: Native PPTX, data-driven, editable

---

## Comprehensive Comparison Matrix

| Capability | ppt-skill | PPTAgent | PPTskill | ChatPPT | pptx-tools | ppt-agent |
|------------|:---------:|:--------:|:--------:|:-------:|:----------:|:---------:|
| Native PPTX | **Yes** | Yes | Yes | Yes | Edit | No |
| Zero-input Generation | No | **Yes** | Partial | **Yes** | No | **Yes** |
| JSON-driven | **Yes** | No | No | No | No | No |
| Theme Skinning | **Yes** | Limited | Template-level | Limited | No | CSS |
| Native Charts | **9 + 4** | Basic | None | Basic | Update | HTML |
| Animations | **8 types** | None | None | Limited | No | No |
| Rendered Preview | **Dual mode** | None | None | Online | No | No |
| Contrast Audit | **Yes** | None | None | None | No | No |
| Text Auto-fit | **Yes** | None | None | None | No | No |
| Freeform Shapes | **Yes** | None | None | None | No | Yes |
| Chinese Fonts | **Native** | Average | Average | Average | Average | Average |
| Offline Deployment | **Yes** | No | Yes | No | Yes | No |
| Open Source & Free | **Yes** | Yes | Yes | Partial | Yes | Yes |

---

## Summary

### ppt-skill's Unique Positioning

ppt-skill's unique value in the ecosystem lies in its **precisely controllable JSON -> native PPTX pipeline**:

1. **Determinism**: JSON specs guarantee fully predictable output, suitable for automated pipelines
2. **Natively Editable**: Produces real PPTX, not images / HTML, preserving full editability
3. **Quality Assurance**: color_guard contrast audit + QA checks avoid visual disasters
4. **Chinese Optimization**: Microsoft YaHei + Noto Sans CJK for professional Chinese display
5. **Lightweight Deployment**: Pure Python, no Docker / model / API required, fully offline
6. **Rich Charts**: 18 layouts + 9 native charts + 4 advanced charts

### Collaboration with Other Projects

```
[PPTAgent] ----(content generation)----> [JSON] ----(ppt-skill)----> [PPTX]
                                              |
[ChatPPT] ----(online service)---------------|
                                              |
[pptx-tools] <---(edit data)-----------------|
```

- Use PPTAgent to generate content outlines, output as JSON, then use ppt-skill for precise visual control
- Use ChatPPT to quickly generate a draft, export JSON, then fine-tune with ppt-skill
- Use ppt-skill to generate a base PPTX, then use pptx-tools to batch-update data
- Use ppt-agent-skills for web presentations, while using ppt-skill to generate an editable version

### Selection Recommendations

| Scenario | Recommended Solution |
|----------|----------------------|
| Need AI automatic research and generation | PPTAgent |
| Refined consulting style, document conversion | PPTskill |
| Rapid online generation, no technical team | ChatPPT-MCP |
| Existing templates, batch data updates | pptx-tools |
| Web presentations, ultimate design | ppt-agent-skills |
| **Precise control, data-driven, batch generation, Chinese scenarios** | **ppt-skill** |

---

*Document version: 2026-07-01*
