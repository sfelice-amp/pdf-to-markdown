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

if [ -z "$VIRTUAL_ENV" ]; then
  if [ ! -d .venv ]; then
    python3 -m venv .venv
  fi
  source .venv/bin/activate
else
  echo "Using active Python virtual environment."
fi

pip install --upgrade pip
pip install -r requirements.txt

echo "Starting the local web app at http://127.0.0.1:5000"
python app.py
