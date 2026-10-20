#!/bin/bash
set -Eeu -o pipefail

cd "$(dirname "$0")"
rm -f sample?.log sample?.pptx sample?\ -\ backup.pptx sample?\ -\ backup\ \(*\).pptx
