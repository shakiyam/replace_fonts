# AGENTS.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

replace_fonts is a Python tool that replaces and unifies fonts in PowerPoint presentations. It solves the problem of mixed fonts in presentations by replacing all fonts with the theme's default fonts (heading/body fonts for Latin/East Asian text).

## Key Commands

### Initial Setup

```bash
# Install git hooks (run once after cloning)
make hooks
```

This installs a pre-commit hook that automatically updates the version number in `replace_fonts.py` to the current date (YYYY-MM-DD format) whenever `replace_fonts.py` is modified and committed.

### Building and Testing

```bash
# Run all checks (see `make help` for what it includes)
make all

# Run pytest tests
make pytest

# Run integration tests
make test
```

### Formatting, Linting, and Type Checking

```bash
# Run all formatters (rewrites files in place)
make format

# Run all linters
make lint

# Run Python type checking
make mypy

# List all targets, including individual formatters and linters
make help
```

### Building Docker Images

```bash
# Build production image
make build

# Build development image (includes dev dependencies)
make build_dev
```

When `make trivy` reports vulnerabilities whose fixed versions the distribution already ships, run `NO_CACHE=1 make trivy` to rebuild the cached layers that still hold the old packages.

### Updating Dependencies

```bash
# Update uv.lock and export requirements.txt from it
make update_uv_lock
```

The project uses `uv` to manage dependencies. Edit `pyproject.toml` to add/remove dependencies, then run the make command above to regenerate `uv.lock` and `requirements.txt` with pinned versions.

## Architecture Notes

- **Main Script**: `replace_fonts.py` - CLI, file I/O, backup, entry point
- **Logger**: `logger.py` - Log output function and LogFn type alias
- **Define Theme Fonts**: `define_theme_fonts.py` - Font policy loading, validation, and theme font updating
- **Apply Theme Fonts**: `apply_theme_fonts.py` - Font replacement logic, shape/slide processing
- **Shell Wrappers**: `replace_fonts` (production) and `replace_fonts_dev` (development) - Convenient wrapper scripts for running the tool via Docker
- **Docker Support**: Two Dockerfiles - production (`Dockerfile`) and development (`dev.Dockerfile`)
- **Module List**: When adding or removing a Python module, update the module lists in `Dockerfile`, `dev.Dockerfile` (COPY), and `replace_fonts_dev` (mounts that override the image copies with the working tree)
- **Testing**:
  - pytest-based unit tests in `test/test_replace_fonts.py`
  - Integration tests in `test/run.sh` using sample files
  - Log verification to ensure correct font replacements
- **Logging Behavior**: Log files use append mode (`'a'`) by design to preserve historical records when the same file is processed multiple times
- **Error Handling**: `FileNotFoundError` and `PackageNotFoundError` are intentionally handled together with the same error message, as `PackageNotFoundError` can also be raised when a file does not exist (not just for invalid PPTX files)

## Development Workflow

1. The project uses containerized tools for consistency - most development tools (ruff, hadolint, shellcheck, etc.) run inside Docker containers
2. When running Python code for testing or verification, always use `./replace_fonts_dev python3` instead of the host's `python3` to ensure correct dependencies and environment
3. All shell scripts and tools follow strict error handling (`set -Eeu -o pipefail`)

## File Structure

- **`tools/`**: Containerized development tool wrappers - project-independent tools using standalone Docker images or binaries
- **`replace_fonts_dev`**: Wrapper script for project-dependent tools (mypy, pytest) using the replace_fonts_dev Docker image
- **`test/`**: Test suite with sample PPTX files, test scripts, and expected log outputs
  - `test/original/`: Sample PPTX files for testing
  - `test/expected/`: Expected log outputs for verification
  - `test/test_replace_fonts.py`: pytest test cases
  - `test/run.sh`: Integration test script

## Code Style

- Use Enums for related constants (e.g., ThemeFont, FontScript)
- Use mapping dictionaries to eliminate repetitive conditional logic

## Decisions

- **Dependencies pinned in `uv.lock`**: The container images install from `uv.lock`, and local installs without containers (see README) install from `requirements.txt` exported from it, so every environment runs the tested versions. Runtime and development dependencies are resolved together, so a development tool may constrain a runtime version; this is accepted as keeping the tested and shipped versions identical
- **Exported `requirements.txt` committed**: Windows users install with pip through `setup.bat` instead of uv, because the SendTo shortcut records the interpreter path and a system Python is less likely to disappear from under it than the uv-managed Python behind a project venv; the exported file gives them the pinned versions without uv
- **Development dependencies as an extra**: `dev` stays in `[project.optional-dependencies]` rather than `[dependency-groups]`, because `uv sync` installs the `dev` group by default but not extras, so a plain `uv sync` installs only what the application needs
- **No platform restriction in `uv.lock`**: `[tool.uv] environments` is left unset because narrowing it drops only wheels, not packages, while making installs fail on unlisted platforms; revisit if the daily update check keeps reporting lock changes caused only by platforms nobody uses
- **System Python in images**: The Dockerfiles run `uv sync` with `UV_PROJECT_ENVIRONMENT=/usr/local` instead of creating a venv, so the installed commands stay on the default `PATH`
