# 测试指南

PPT Skill 使用 `tests/` 目录下的 JSON 固定用例进行回归测试和布局验证。

## 运行测试构建

```bash
python scripts/build_pptx.py tests/test_all_constraints.json output/test.pptx --theme minimal
python scripts/qa_check.py output/test.pptx
```

## 测试分类

| 目录/文件 | 用途 |
|---|---|
| `tests/test_*.json` | 特定布局或功能的回归测试 |
| `tests/stress_*.json` | 高密度或边界条件内容的压力测试 |
| `tests/business_consulting.json` | 商务风格多布局 Deck |
| `tests/tech_report.json` | 技术报告风格 Deck |
| `tests/academic_course.json` | 学术演示风格 |
| `tests/test_image*.png/json` | 图片布局测试 |

## 添加新测试

1. 在 `tests/` 下创建 JSON deck。
2. 运行构建器，确认生成有效的 PPTX。
3. 运行 `qa_check.py`，确保输出 `RESULT: PASS`。
4. 提交 JSON 文件。

## 持续集成清单

最小 CI 流程应包括：

1. 安装依赖：`pip install -r requirements.txt`
2. 语法检查：`python -m py_compile scripts/*.py`
3. 构建测试：`python scripts/build_pptx.py tests/test_all_constraints.json output/ci.pptx`
4. QA 检查：`python scripts/qa_check.py output/ci.pptx`
