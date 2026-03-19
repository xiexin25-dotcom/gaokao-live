#!/bin/bash
set -e
cd "$(dirname "$0")"

echo ""
echo "  ============================================"
echo "   Gaokao Volunteer Planner v3.4  (JiLin)"
echo "  ============================================"
echo ""

# Check Python3
if command -v python3 &>/dev/null; then
    PY=python3
elif command -v python &>/dev/null; then
    PY=python
else
    echo "  [ERROR] Python3 not found"
    echo "  Install: https://www.python.org/downloads/"
    exit 1
fi

echo "  Python: $($PY --version)"

# Check app.py
if [ ! -f app.py ]; then
    echo "  [ERROR] app.py not found. Run inside gaokao_local folder."
    exit 1
fi

# Check data
if ! ls data/*.xlsx &>/dev/null; then
    echo "  [ERROR] Data file missing: data/*.xlsx"
    exit 1
fi

echo "  Installing/checking packages..."
$PY -m pip install flask openpyxl pandas numpy -q

echo ""
echo "  Starting server:  http://localhost:5000"
echo "  Press Ctrl+C to stop"
echo ""

# Auto-open browser after 2s
(sleep 2 && (
    open "http://localhost:5000" 2>/dev/null ||
    xdg-open "http://localhost:5000" 2>/dev/null
) ) &

PYTHONIOENCODING=utf-8 $PY app.py
