#!/bin/zsh
set -eu

cd "/Users/klland/Documents/Stock Analysis"

echo "[$(date '+%Y-%m-%d %H:%M:%S %Z')] remote sync start"

git fetch origin
git merge --ff-only origin/main

node --check app.js
node --check data/market-manifest.js
node --check data/market-data.js
node --check data/market-history.js

echo "Remote market data synced. GitHub Actions is the only market-data writer."
