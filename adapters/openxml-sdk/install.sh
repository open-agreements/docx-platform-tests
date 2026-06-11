#!/usr/bin/env bash
# Builds the Open XML SDK adapter. Uses the dotnet on PATH when present
# (GitHub ubuntu runners preinstall one); otherwise installs a user-local
# SDK via the official install script -- no root required either way.
set -euo pipefail
cd "$(dirname "$0")"

if command -v dotnet >/dev/null 2>&1; then
  DOTNET=dotnet
else
  curl -fsSL https://dot.net/v1/dotnet-install.sh -o dotnet-install.sh
  bash dotnet-install.sh --channel 8.0 --install-dir "$PWD/.dotnet"
  rm dotnet-install.sh
  DOTNET="$PWD/.dotnet/dotnet"
fi

"$DOTNET" publish src/OpenXmlSdkAdapter.csproj -c Release -o build --nologo
bash run-adapter.sh --print-library-version
