#!/usr/bin/env bash
set -euo pipefail

REPO_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO_DIR"

git add -A .
# Keep local runtime state out of deployment commits.
git restore --staged server/telegram-bridge-state.json 2>/dev/null || true

if git diff --cached --quiet; then
  echo "No changes to deploy."
  exit 0
fi

commit_msg="chore: auto deploy $(date +%Y-%m-%d_%H-%M-%S)"
git commit -m "$commit_msg"
git push origin main

echo "Pushed to main. GitHub Pages workflow will deploy automatically."
