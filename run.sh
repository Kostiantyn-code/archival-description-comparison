#!/usr/bin/env sh
set -eu
cd "$(dirname "$0")"

if [ ! -x ".venv/bin/python" ]; then
    echo "Creating a local Python environment..."
    python3 -m venv .venv
fi
if ! cmp -s requirements.txt .venv/installed-requirements.txt; then
    echo "Installing dependencies (first launch or changed requirements)..."
    .venv/bin/python -m pip install -r requirements.txt
    cp requirements.txt .venv/installed-requirements.txt
fi
.venv/bin/python main.py "$@"
