#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

if ! command -v cargo >/dev/null 2>&1; then
  echo "cargo is required to build the docx-rs adapter" >&2
  exit 1
fi

cargo build --release
