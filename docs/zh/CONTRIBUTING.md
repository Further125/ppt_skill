# 贡献指南

感谢您对 PPT Skill 的关注！本文档提供参与项目的指南。

## 如何贡献

### 报告问题

- 使用 GitHub Issues 报告 bug 或请求功能。
- 提供能复现问题的最小 JSON deck。
- 说明 Python 版本、操作系统和已安装的依赖版本。

### 提交变更

1. Fork 仓库并创建功能分支。
2. 进行清晰、聚焦的提交。
3. 根据需要添加或更新测试和文档。
4. 确保 Python 脚本通过语法检查：`python -m py_compile scripts/*.py`
5. 打开 Pull Request，并附上描述性的标题和摘要。

## 代码风格

- Python 代码遵循 PEP 8。
- 保持函数聚焦，为非显而易见的逻辑添加注释。
- 优先使用相对路径，避免硬编码绝对路径。

## 文档规范

- 英文文档放在 `docs/`。
- 中文文档放在 `docs/zh/`。
- 每个文件使用单一语言。

## 开发环境

```bash
git clone <repo-url>
cd ppt_skill
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

快速构建测试：

```bash
python scripts/build_pptx.py examples/demo_deck.json output/test.pptx
```

## 许可证

通过贡献，您同意您的贡献将在 MIT 许可证下发布。
