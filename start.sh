#!/bin/bash
echo "=========================================="
echo "  CMMC L2 SSP Drafting Tool"
echo "=========================================="
echo

# Check Python
if ! command -v python3 &>/dev/null; then
    echo "ERROR: Python 3 is not installed."
    echo "Install with: brew install python3  (macOS)"
    echo "Or: sudo apt install python3 python3-pip  (Linux)"
    exit 1
fi

echo "Installing/updating dependencies..."
pip3 install -r requirements.txt -q

echo
echo "Starting application..."
echo "Open your browser to: http://127.0.0.1:5000"
echo
echo "Press Ctrl+C to stop the server."
echo
python3 app.py
