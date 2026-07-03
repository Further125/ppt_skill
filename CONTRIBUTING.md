# Contributing to PPT Skill

Thank you for your interest in contributing! This document provides guidelines for participating in the project.

## How to Contribute

### Reporting Issues

- Use GitHub Issues to report bugs or request features.
- Include a minimal JSON deck that reproduces the problem.
- Mention your Python version, OS, and installed dependency versions.

### Submitting Changes

1. Fork the repository and create a feature branch.
2. Make your changes with clear, focused commits.
3. Add or update tests and documentation as needed.
4. Ensure Python scripts pass syntax checks: `python -m py_compile scripts/*.py`
5. Open a pull request with a descriptive title and summary.

### Code Style

- Follow PEP 8 for Python code.
- Keep functions focused and document non-obvious logic.
- Prefer relative paths and avoid hardcoded absolute paths.

### Documentation

- English documentation lives in `docs/`.
- Chinese documentation lives in `docs/zh/`.
- Keep each file in a single language.

## Development Setup

```bash
git clone <repo-url>
cd ppt_skill
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Run a quick build test:

```bash
python scripts/build_pptx.py examples/demo_deck.json output/test.pptx
```

## License

By contributing, you agree that your contributions will be licensed under the MIT License.
