#!/bin/bash
# Usage: recalc_copy.sh source.xlsx dest.xlsx [timeout] -> copie (si différente) puis recalcul LibreOffice in place sur la copie
set -e
SK=/root/.claude/skills/synced/613869b7-18fd-4e4f-a3ea-f504b0dc1277_6a8a304d-5172-4887-9b7a-248c0ff3e593/xlsx/scripts
if [ "$(realpath "$1")" != "$(realpath -m "$2")" ]; then cp "$1" "$2"; fi
chmod u+w "$2"
DEST=$(realpath "$2")
cd "$SK" && python3 recalc.py "$DEST" ${3:-240}
