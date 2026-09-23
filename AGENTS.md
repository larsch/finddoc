## Python tools

- Use `uv` for Python environment and dependency work.
- Use `uvx` for one-off tool execution.
- Use `pytest` for tests.
- Use `ruff check` and `ruff format` for linting and formatting.
- Avoid `pip`, `pipx`, `poetry`, `hatch`, `tox`, `nox`, `black`, `isort`, and `unittest`.
- Never use `setup.py`, `setup.cfg`, or `requirements.txt`. All dependencies and metadata belong in `pyproject.toml` (PEP 621).

## Commit preflight checks

- All unittests must pass before committing.
- Code must pass linting checks before committing.
- Code must be properly formatted before committing.
