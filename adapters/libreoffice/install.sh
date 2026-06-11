#!/usr/bin/env bash
# Installs LibreOffice Writer and pyuno on the Ubuntu CI runner.
set -euo pipefail

if [ "$(id -u)" -eq 0 ]; then
  APT_GET=(apt-get)
elif command -v sudo >/dev/null 2>&1; then
  APT_GET=(sudo apt-get)
else
  APT_GET=(apt-get)
fi

"${APT_GET[@]}" update
"${APT_GET[@]}" install -y libreoffice-writer python3-uno

soffice --version
