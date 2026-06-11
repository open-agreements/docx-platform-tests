#!/usr/bin/env bash
# Builds @usejunior/docx-core at the pinned safe-docx commit and installs the
# packed tarball locally, exposing node_modules/.bin/safe-docx-conformance-adapter.
# Temporary until the bin ships in a published npm release (issue #3); a fork
# reproduces this build from the public repository alone.
#
# Set SAFE_DOCX_SRC to an existing safe-docx checkout to pack from it instead
# of cloning (local development convenience; CI always clones the pin).
set -euo pipefail
cd "$(dirname "$0")"

if [ -n "${SAFE_DOCX_SRC:-}" ]; then
  SRC="$SAFE_DOCX_SRC"
else
  REPO_URL=$(node -p "require('./safe-docx.pin.json').repository")
  SHA=$(node -p "require('./safe-docx.pin.json').commitSha")
  rm -rf build
  mkdir -p build/src
  git -C build/src init -q
  git -C build/src fetch -q --depth 1 "$REPO_URL" "$SHA"
  git -C build/src checkout -q "$SHA"
  (cd build/src && npm ci --silent)
  SRC="$PWD/build/src"
fi

(cd "$SRC/packages/docx-core" && npm run build >/dev/null)
TARBALL=$(cd "$SRC/packages/docx-core" && npm pack --silent)
rm -f package.json package-lock.json
npm init -y >/dev/null
npm install --silent "$SRC/packages/docx-core/$TARBALL"
test -x node_modules/.bin/safe-docx-conformance-adapter
echo "installed safe-docx-conformance-adapter from $TARBALL"
