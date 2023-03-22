#!/bin/bash
set -eu -o pipefail

cd "$(dirname "$0")"
./clean.sh
cp original/sample?.pptx .
../replace_fonts sample1.pptx sample2.pptx sample3.pptx
