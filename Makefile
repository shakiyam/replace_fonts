MAKEFLAGS += --no-builtin-rules
MAKEFLAGS += --warn-undefined-variables
SHELL := /bin/bash
.SHELLFLAGS := -eu -o pipefail -c
ALL_TARGETS := $(shell grep -E -o ^[0-9A-Za-z_-]+: $(MAKEFILE_LIST) | sed 's/://')
.PHONY: $(ALL_TARGETS)
.DEFAULT_GOAL := help

all: lint update_requirements_dev build_dev mypy test_dev update_requirements build test ## Lint, update requirements.txt, build, and test

build: ## Build image replace_fonts from Dockerfile
	@echo -e "\033[36m$@\033[0m"
	@./tools/build.sh ghcr.io/shakiyam/replace_fonts Dockerfile

build_dev: ## Build image replace_fonts_dev from Dockerfile_dev
	@echo -e "\033[36m$@\033[0m"
	@./tools/build.sh ghcr.io/shakiyam/replace_fonts_dev Dockerfile_dev

flake8: ## Lint Python code
	@echo -e "\033[36m$@\033[0m"
	@./tools/flake8.sh --max-line-length=88

hadolint: ## Lint Dockerfile
	@echo -e "\033[36m$@\033[0m"
	@./tools/hadolint.sh Dockerfile Dockerfile_dev

help: ## Print this help
	@echo 'Usage: make [target]'
	@echo ''
	@echo 'Targets:'
	@awk 'BEGIN {FS = ":.*?## "} /^[0-9A-Za-z_-]+:.*?## / {printf "\033[36m%-30s\033[0m %s\n", $$1, $$2}' $(MAKEFILE_LIST)

lint: flake8 hadolint shellcheck shfmt ## Lint for all dependencies

mypy: ## Lint Python code
	@echo -e "\033[36m$@\033[0m"
	@./tools/mypy.sh ghcr.io/shakiyam/replace_fonts_dev --ignore-missing-imports replace_fonts.py

shellcheck: ## Lint shell scripts
	@echo -e "\033[36m$@\033[0m"
	@./tools/shellcheck.sh replace_fonts replace_fonts_dev test/*.sh tools/*.sh

shfmt: ## Lint shell scripts
	@echo -e "\033[36m$@\033[0m"
	@./tools/shfmt.sh -l -d -i 2 -ci -bn replace_fonts replace_fonts_dev test/*.sh tools/*.sh

test: ## Test replace_fonts
	@echo -e "\033[36m$@\033[0m"
	@./test/run.sh

test_dev: ## Test replace_fonts
	@echo -e "\033[36m$@\033[0m"
	@./test/run_dev.sh

update_requirements: ## Update requirements.txt
	@echo -e "\033[36m$@\033[0m"
	@./tools/uv.sh pip compile --upgrade --strip-extras --output-file requirements.txt requirements.in

update_requirements_dev: ## Update requirements_dev.txt
	@echo -e "\033[36m$@\033[0m"
	@./tools/uv.sh pip compile --upgrade --strip-extras --output-file requirements_dev.txt requirements_dev.in
