#!/bin/bash
clear
echo ""
echo "  ========================"
echo "   GEOLOCATOR"
echo "  ========================"
echo ""
if ! command -v python3 &>/dev/null; then
    echo "  ❌  Python 3 not found."
    exit 1
fi
echo "  ✅  $(python3 --version)"
pip3 install streamlit requests openpyxl pandas --quiet
echo "  ✅  Ready — opening at http://localhost:8501"
echo "  Leave this window open. Press Ctrl+C to stop."
echo ""
cd "$(dirname "$0")"
python3 -m streamlit run app.py --server.headless false --browser.gatherUsageStats false
