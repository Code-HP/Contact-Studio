#!/bin/bash
# Runs all tests against a fresh bundle of Contact_Studio.js.
HERE="$(cd "$(dirname "$0")" && pwd)"
BUNDLE="$(mktemp -t contact_studio_bundle).js"
bash "$HERE/extract.sh" "$BUNDLE" || exit 1
status=0
for t in "$HERE"/*.test.js; do
  echo "== $(basename "$t")"
  node "$t" "$BUNDLE" || status=1
done
rm -f "$BUNDLE"
exit $status
