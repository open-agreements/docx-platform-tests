#!/usr/bin/env bash
set -euo pipefail

adapter_dir="$(cd "$(dirname "$0")" && pwd)"
binary="$adapter_dir/target/release/docx-rs-adapter"

if [[ "${1:-}" == "--print-library-version" ]]; then
  echo "docx-rs 0.4.20"
  exit 0
fi

if [[ -x "$binary" ]]; then
  exec "$binary" "$@"
fi

echo "docx-rs adapter binary is missing; run adapters/docx-rs/install.sh" >&2
exit 1
