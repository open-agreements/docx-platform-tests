#!/usr/bin/env bash
# Installs dolanmiu/docx locally so the adapter can report the exact package
# version and exercise public exports during unsupported-path checks.
set -euo pipefail
cd "$(dirname "$0")"
npm ci --silent
node -p "'docx ' + require('./node_modules/docx/package.json').version"
