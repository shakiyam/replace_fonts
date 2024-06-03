#!/bin/bash
set -eu -o pipefail

cd "$(dirname "$0")"
./clean.sh
cp original/sample?.pptx .
../replace_fonts_dev python3.12 /replace_fonts.py --code sample?.pptx
