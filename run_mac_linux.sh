#!/bin/bash
echo "========================================"
echo "  Assignment Project Generator — Setup  "
echo "========================================"
echo ""

# Check Python 3
if ! command -v python3 &> /dev/null; then
    echo "[ERROR] python3 not found. Install it from https://python.org"
    exit 1
fi

echo "[1/2] Installing dependencies..."
pip3 install -r requirements.txt --quiet

echo "[2/2] Launching app..."
echo ""
python3 assignment_generator.py
