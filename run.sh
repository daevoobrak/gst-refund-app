#!/usr/bin/env bash
# GST Refund Working Generator – Mac/Linux launcher

cd "$(dirname "$0")"

echo "========================================="
echo " GST Refund Working Generator"
echo "========================================="
echo

# Check Python
if ! command -v python3 &>/dev/null; then
    echo "ERROR: python3 not found. Install from https://python.org"
    exit 1
fi
echo "Found: $(python3 --version)"
echo

# Create venv if needed
if [ ! -d ".venv" ]; then
    echo "Setting up virtual environment (first run only)..."
    python3 -m venv .venv
fi

source .venv/bin/activate

echo "Installing dependencies..."
pip install --quiet flask "pdfplumber==0.10.4" openpyxl

echo
echo "Starting server at http://localhost:5050"
echo "Press Ctrl+C to stop."
echo

# Open browser
sleep 1 && open http://localhost:5050 2>/dev/null &

python3 app.py
