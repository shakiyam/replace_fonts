#!/bin/bash
set -Eeu -o pipefail

cd "$(dirname "$0")"
# shellcheck disable=SC1091
. ../tools/colored_echo.sh
# shellcheck disable=SC1091
. ../tools/container_engine.sh

# Started without the wrapper's -u, the image itself must not run as root;
# skipped without a container engine, where the wrapper falls back to the local Python
if CONTAINER_ENGINE=$(detect_container_engine 2>/dev/null); then
  IMAGE_UID=$($CONTAINER_ENGINE container run --rm --entrypoint id ghcr.io/shakiyam/replace_fonts -u)
  if [[ $IMAGE_UID == 0 ]]; then
    echo_error 'The replace_fonts image runs as root.'
    exit 1
  fi
fi

./clean.sh
cp original/sample?.pptx .
../replace_fonts --code sample?.pptx
