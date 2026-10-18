#!/usr/bin/env bash
# ============================================================
#  Thesis AI Detector - launcher for Mac (and Linux)
#  How to use:
#    1. Right-click run.sh -> Open With -> Terminal, or
#       open Terminal in this folder and run:  bash run.sh
#    2. You can also pass the file directly:   bash run.sh thesis.docx
# ============================================================
set -e
cd "$(dirname "$0")"

echo ""
echo " ================================================"
echo "  Thesis AI Detector"
echo " ================================================"
echo ""

# 1. Check Python is installed
if ! command -v python3 >/dev/null 2>&1; then
  echo "[X] Python 3 was not found."
  echo ""
  echo "Please install it first:"
  echo "  1. Go to https://www.python.org/downloads/"
  echo "  2. Download Python 3.10 or newer and install it."
  echo "  3. Then run this script again:  bash run.sh"
  exit 1
fi

# 2. Create private environment on first run only
if [ ! -f ".venv/bin/python" ]; then
  echo "First-time setup: creating a private folder (.venv)..."
  echo "This takes a few minutes only once."
  python3 -m venv .venv
fi

# shellcheck disable=SC1091
source .venv/bin/activate

# 3. Install needed libraries (fast if already installed)
echo "Checking libraries (fast if already done)..."
python -m pip install --quiet --upgrade pip >/dev/null 2>&1 || true
python -m pip install --quiet -r requirements.txt

# 4. Find the thesis file: argument or ask
THESIS="$1"
if [ -z "$THESIS" ]; then
  echo ""
  echo "Tip: next time you can run:  bash run.sh thesis.docx"
  echo ""
  printf "Type the name of your file (example: thesis.docx) and press Enter: "
  read -r THESIS
fi

if [ -z "$THESIS" ]; then
  echo "[X] No file given. Put your file next to run.sh and try again."
  exit 1
fi

if [ ! -f "$THESIS" ]; then
  echo "[X] File not found: \"$THESIS\""
  echo "    Make sure the file is in the same folder as run.sh,"
  echo "    or type the full path, e.g. /Users/You/Documents/thesis.docx"
  exit 1
fi

echo ""
echo "Analysing \"$THESIS\"..."
echo "This takes 1-5 minutes. Please wait."
echo ""
python ai_detector.py "$THESIS" --report report.html --csv results.csv

echo ""
echo " ================================================"
echo "  DONE!"
echo "  Open report.html (double-click it) to see your"
echo "  colour-coded results in your web browser."
echo "  results.csv has one row per paragraph."
echo " ================================================"
