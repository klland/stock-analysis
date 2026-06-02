#!/bin/zsh
set -eu

cd "/Users/klland/Documents/Stock Analysis"

git fetch origin
git merge --ff-only origin/main
