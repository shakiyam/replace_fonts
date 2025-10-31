MAKEFLAGS += --no-builtin-rules
MAKEFLAGS += --warn-undefined-variables
SHELL := /bin/bash
.SHELLFLAGS := -eu -o pipefail -c
ALL_TARGETS := $(shell grep -E -o ^[0-9A-Za-z_-]+: $(MAKEFILE_LIST) | sed 's/://')
.PHONY: $(ALL_TARGETS)
.DEFAULT_GOAL := help

all: lint update_requirements_dev build_dev mypy pytest update_requirements build test ## Lint, update requirements.txt, build, and test

build: ## Build image replace_fonts from Dockerfile
	@echo -e "\033[36m$@\033[0m"
	@./tools/build.sh ghcr.io/shakiyam/replace_fonts Dockerfile

build_dev: ## Build image replace_fonts_dev from Dockerfile_dev
	@echo -e "\033[36m$@\033[0m"
	@./tools/build.sh ghcr.io/shakiyam/replace_fonts_dev Dockerfile_dev

hadolint: ## Lint Dockerfile
	@echo -e "\033[36m$@\033[0m"
	@./tools/hadolint.sh Dockerfile Dockerfile_dev

help: ## Print this help
	@echo 'Usage: make [target]'
	@echo ''
	@echo 'Targets:'
	@awk 'BEGIN {FS = ":.*?## "} /^[0-9A-Za-z_-]+:.*?## / {printf "\033[36m%-30s\033[0m %s\n", $$1, $$2}' $(MAKEFILE_LIST)

hooks: ## Install git hooks
	@echo -e "\033[36m$@\033[0m"
	@ln -sf ../../hooks/pre-commit .git/hooks/pre-commit
	@echo "Git hooks installed"

lint: ruff hadolint markdownlint shellcheck shfmt ## Lint for all dependencies

markdownlint: ## Lint Markdown files
	@echo -e "\033[36m$@\033[0m"
	@./tools/markdownlint.sh "*.md"

mypy: ## Lint Python code
	@echo -e "\033[36m$@\033[0m"
	@[[ -d .mypy_cache ]] || mkdir .mypy_cache
	@./replace_fonts_dev mypy *.py test/*.py

pytest: ## Run pytest
	@echo -e "\033[36m$@\033[0m"
	@./replace_fonts_dev pytest

ruff: ## Lint Python code
	@echo -e "\033[36m$@\033[0m"
	@./tools/ruff.sh check

shellcheck: ## Lint shell scripts
	@echo -e "\033[36m$@\033[0m"
	@./tools/shellcheck.sh replace_fonts replace_fonts_dev test/*.sh tools/*.sh hooks/*

shfmt: ## Lint shell scripts
	@echo -e "\033[36m$@\033[0m"
	@./tools/shfmt.sh -l -d -i 2 -ci -bn replace_fonts replace_fonts_dev test/*.sh tools/*.sh hooks/*

test: ## Test replace_fonts
	@echo -e "\033[36m$@\033[0m"
	@./test/run.sh

update_requirements: ## Update requirements.txt
	@echo -e "\033[36m$@\033[0m"
	@./tools/uv.sh pip compile --upgrade --strip-extras --output-file requirements.txt pyproject.toml

update_requirements_dev: ## Update requirements_dev.txt
	@echo -e "\033[36m$@\033[0m"
	@./tools/uv.sh pip compile --upgrade --strip-extras --extra dev --output-file requirements_dev.txt pyproject.toml
