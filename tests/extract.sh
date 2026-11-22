#!/bin/bash
# Builds a DOM-free bundle of the pure logic in Contact_Studio.js for the tests.
# Usage: extract.sh <output-file>
HERE="$(cd "$(dirname "$0")" && pwd)"
SRC="$HERE/../Contact_Studio.js"
OUT="$1"
range() { awk -v a="$1" -v b="$2" 'index($0,a){p=1} index($0,b){p=0} p' "$SRC"; }
{
  echo '"use strict";'
  echo 'globalThis.BroadcastChannel = undefined;'
  echo 'const window = { addEventListener() {} }; const els = {}; function renderAll() {}'
  awk 'index($0,"const state = {"){p=1} p{print} p && /^  };/{exit}' "$SRC"
  range "2. CONSTANTS" "3. DOM"
  range "4. UTILITIES" "12. FILTERS"
  awk '/^  function mergeRecords\(/{p=1} /^  function mergeGroupNow\(/{p=0} p' "$SRC"
  echo 'module.exports = { state, MATCH_TYPES, formatPhoneCellValue, parseLabeledPhone, computeDuplicatesForList, rebuildContactDerived, hydrateContact, normPhone, normEmail, normName, PHONE_LABEL_MAP, buildPayload, applyLoadedData, pushHistory, undo, redo, history, saveStateNow, snapshot, computeAllDuplicates, parseAnyFile, splitPhoneListLabeled, mergeRecords, normPhoneKey, nameMatchKey };'
} > "$OUT"
