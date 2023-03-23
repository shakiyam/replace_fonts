#!/bin/bash
set -eu -o pipefail

cd "$(dirname "$0")"
./clean.sh
cp original/sample?.pptx .
../replace_fonts --code sample?.pptx
