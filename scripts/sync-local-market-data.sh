#!/bin/zsh
set -eu

cd "/Users/klland/Documents/Stock Analysis"

echo "[$(date '+%Y-%m-%d %H:%M:%S %Z')] refresh start"

git fetch origin
git merge --ff-only origin/main

node scripts/update-market-data.mjs
node --check app.js
node --check data/market-data.js

if git diff --quiet -- data/market-data.js; then
  echo "No local market data changes to commit."
  exit 0
fi

git add data/market-data.js
git commit -m "chore: update market data"
git push origin main
