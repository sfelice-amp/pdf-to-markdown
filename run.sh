#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"

if ! command -v brew >/dev/null 2>&1; then
  echo "Homebrew is required to install Tesseract and Poppler."
  echo "Install Homebrew from https://brew.sh/ and run ./run.sh again."
  exit 1
fi

for pkg in tesseract poppler; do
  if ! command -v "$pkg" >/dev/null 2>&1; then
    echo "Installing $pkg..."
    brew install "$pkg"
  else
    echo "$pkg is already installed."
  fi
done

if [ ! -d .venv ]; then
  python3 -m venv .venv
fi

PYTHON_EXEC="$(pwd)/.venv/bin/python3"
PIP_EXEC="$(pwd)/.venv/bin/pip"

if [ ! -x "$PYTHON_EXEC" ]; then
  echo "Unable to find Python in .venv."
  exit 1
fi

"$PIP_EXEC" install --upgrade pip
"$PIP_EXEC" install -r requirements.txt

PORT=0
for candidate in $(seq 5000 5010); do
  if ! lsof -iTCP:"$candidate" -sTCP:LISTEN -t >/dev/null 2>&1; then
    PORT=$candidate
    break
  fi
done

if [ "$PORT" -eq 0 ]; then
  echo "No available port found between 5000 and 5010. Please stop the existing server or free a port."
  exit 1
fi

echo "Starting the local web app at http://127.0.0.1:$PORT"
"$PYTHON_EXEC" app.py "$PORT"
