#!/usr/bin/env bash
# Собирает документ и презентацию, затем рендерит их в PNG для визуальной проверки (нужен LibreOffice).
set -euo pipefail
cd "$(dirname "$0")/.."
PY=.venv/bin/python
$PY scripts/make_assets.py
$PY scripts/make_docx.py
$PY scripts/make_pptx.py
if command -v soffice >/dev/null; then
  rm -rf build/render && mkdir -p build/render
  soffice --headless --convert-to pdf --outdir build/render out/*.pptx out/*.docx >/dev/null 2>&1
  for f in build/render/*.pdf; do
    case "$f" in *презентация*) name=slide ;; *) name=page ;; esac
    pdftoppm -r 60 -png "$f" "build/render/$name"
  done
  ls build/render/*.pdf
  for f in build/render/*.pdf; do pdfinfo "$f" | grep Pages; done
fi
