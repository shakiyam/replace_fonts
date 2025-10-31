# CLAUDE.md

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
# Run all checks (lint, update requirements, build, test)
make all

# Run pytest tests
make pytest

# Run integration tests
make test
```

### Linting and Type Checking

```bash
# Run all linting
make lint

# Individual linters:
make ruff      # Python linting
make hadolint  # Dockerfile linting
make shellcheck # Shell script linting
make shfmt     # Shell script formatting
make mypy      # Python type checking
```

### Building Docker Images

```bash
# Build production image
make build

# Build development image (includes dev dependencies)
make build_dev
```

### Updating Dependencies

```bash
# Update production requirements
make update_requirements

# Update development requirements
make update_requirements_dev
```

The project uses `uv` to manage dependencies. Edit `pyproject.toml` to add/remove dependencies, then run the make commands above to regenerate `*.txt` files with pinned versions.

## Architecture Notes

- **Main Script**: `replace_fonts.py` - Core font replacement logic using python-pptx library
- **Shell Wrappers**: `replace_fonts` (production) and `replace_fonts_dev` (development) - Convenient wrapper scripts for running the tool via Docker
- **Docker Support**: Two Dockerfiles - production (`Dockerfile`) and development (`Dockerfile_dev`)
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
4. Type hints are used throughout the codebase for type safety
5. Before implementing changes, run tests first to understand current behavior
6. Make changes incrementally in small, verifiable steps
7. Run tests after changes to verify behavior

## File Structure

- **`tools/`**: Containerized development tool wrappers (hadolint, ruff, shellcheck, shfmt, uv) - project-independent tools using standalone Docker images
- **`replace_fonts_dev`**: Wrapper script for project-dependent tools (mypy, pytest) using the replace_fonts_dev Docker image
- **`test/`**: Test suite with sample PPTX files, test scripts, and expected log outputs
  - `test/original/`: Sample PPTX files for testing
  - `test/expected/`: Expected log outputs for verification
  - `test/test_replace_fonts.py`: pytest test cases
  - `test/run.sh`: Integration test script

## Code Style

- Use type hints for all function parameters and return values
- Follow existing naming conventions (clear, descriptive variable names)
- Eliminate code duplication by extracting common patterns
- Keep functions focused on single responsibilities
- Use Enums for related constants (e.g., ThemeFont, FontScript)
- Use mapping dictionaries to eliminate repetitive conditional logic
