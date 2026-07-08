#!/usr/bin/env bash
# Sync the canonical skill into the plugin distribution copy, and (optionally)
# produce an installable zip of the skill for the drop-in install method.
#
#   plugins/sync_plugin.sh          # sync canonical -> plugin copy
#   plugins/sync_plugin.sh --zip    # also write dist/three-statement-model-skill.zip
#
# The canonical source is .claude/skills/three-statement-model/ (auto-loads as a
# project skill). The plugin copy under plugins/three-statement-model/skills/ is a
# generated distribution artifact — never edit it by hand.
set -euo pipefail
REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
CANON="$REPO_ROOT/.claude/skills/three-statement-model"
PLUGIN_SKILL="$REPO_ROOT/plugins/three-statement-model/skills/three-statement-model"

if [ ! -d "$CANON" ]; then
  echo "canonical skill not found at $CANON" >&2
  exit 1
fi

rm -rf "$PLUGIN_SKILL"
mkdir -p "$(dirname "$PLUGIN_SKILL")"
cp -r "$CANON" "$PLUGIN_SKILL"
# Drop caches if any slipped in.
find "$PLUGIN_SKILL" -name '__pycache__' -type d -prune -exec rm -rf {} + 2>/dev/null || true
echo "synced: $CANON -> $PLUGIN_SKILL"

if [ "${1:-}" = "--zip" ]; then
  DIST="$REPO_ROOT/dist"
  mkdir -p "$DIST"
  ZIP="$DIST/three-statement-model-skill.zip"
  rm -f "$ZIP"
  ( cd "$(dirname "$CANON")" && zip -rq "$ZIP" "three-statement-model" -x '*/__pycache__/*' )
  echo "zip: $ZIP"
fi
