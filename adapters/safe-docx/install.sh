#!/usr/bin/env bash
# Builds @usejunior/docx-core from safe-docx source and installs the packed
# tarball locally, exposing node_modules/.bin/safe-docx-conformance-adapter.
# The source commit comes from safe-docx.pin.json: pinnedCommitSha if set
# (reproducing a past matrix), otherwise the tip of trackingBranchName —
# resolved to a concrete commit here and recorded in build-info.json so
# results stay attributable. A fork reproduces this build from the public
# repository alone.
#
# Set SAFE_DOCX_SRC to an existing safe-docx checkout to pack from it instead
# of cloning (local development convenience; CI always clones).
set -euo pipefail
cd "$(dirname "$0")"

if [ -n "${SAFE_DOCX_SRC:-}" ]; then
  SRC="$SAFE_DOCX_SRC"
  SHA=$(git -C "$SRC" rev-parse HEAD 2>/dev/null || echo "unknown")
else
  REPO_URL=$(node -p "require('./safe-docx.pin.json').repository")
  SHA=$(node -p "require('./safe-docx.pin.json').pinnedCommitSha ?? ''")
  if [ -z "$SHA" ]; then
    BRANCH=$(node -p "require('./safe-docx.pin.json').trackingBranchName")
    SHA=$(git ls-remote "$REPO_URL" "refs/heads/$BRANCH" | cut -f1)
    if [ -z "$SHA" ]; then
      echo "could not resolve branch $BRANCH on $REPO_URL" >&2
      exit 1
    fi
  fi
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
node -e "
  const { writeFileSync } = require('node:fs');
  const version = require('./node_modules/@usejunior/docx-core/package.json').version;
  writeFileSync('build-info.json', JSON.stringify({
    packageVersion: version,
    sourceCommitSha: process.argv[1],
  }, null, 2) + '\n');
" "$SHA"
echo "installed safe-docx-conformance-adapter from $TARBALL (source commit $SHA)"
