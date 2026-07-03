# Testing Guide

PPT Skill uses JSON fixtures in `tests/` for regression and layout validation.

## Running a Test Build

```bash
python scripts/build_pptx.py tests/test_all_constraints.json output/test.pptx --theme minimal
python scripts/qa_check.py output/test.pptx
```

## Test Categories

| Directory/File | Purpose |
|---|---|
| `tests/test_*.json` | Regression tests for specific layouts or features |
| `tests/stress_*.json` | Stress tests with dense or edge-case content |
| `tests/business_consulting.json` | Business-style multi-layout deck |
| `tests/tech_report.json` | Technical report style deck |
| `tests/academic_course.json` | Academic presentation style |
| `tests/test_image*.png/json` | Image layout tests |

## Adding a New Test

1. Create a JSON deck in `tests/`.
2. Run the builder and confirm it produces a valid PPTX.
3. Run `qa_check.py` and ensure `RESULT: PASS`.
4. Commit the JSON file.

## Continuous Integration Checklist

A minimal CI pipeline should:

1. Install dependencies: `pip install -r requirements.txt`
2. Syntax check: `python -m py_compile scripts/*.py`
3. Build test: `python scripts/build_pptx.py tests/test_all_constraints.json output/ci.pptx`
4. QA check: `python scripts/qa_check.py output/ci.pptx`
