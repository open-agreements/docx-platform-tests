#!/usr/bin/env bash
# Runs the built adapter with whichever dotnet install.sh used.
set -euo pipefail
here="$(cd "$(dirname "$0")" && pwd)"
if command -v dotnet >/dev/null 2>&1; then
  DOTNET=dotnet
else
  DOTNET="$here/.dotnet/dotnet"
fi
exec "$DOTNET" "$here/build/OpenXmlSdkAdapter.dll" "$@"
