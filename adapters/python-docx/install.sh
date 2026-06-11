#!/usr/bin/env bash
# Installs python-docx into a local venv so the adapter is self-contained.
set -euo pipefail
cd "$(dirname "$0")"
python3 -m venv .venv
.venv/bin/pip install --quiet python-docx
.venv/bin/python -c "import docx; print('python-docx', docx.__version__)"
