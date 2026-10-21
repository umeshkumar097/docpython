#!/bin/bash
# Double-click in Finder: sets up Python packages once, then opens the app in the browser.
cd "$(dirname "$0")" || exit 1
if [ ! -d .venv ]; then
  python3 -m venv .venv || { echo "Python 3 was not found. Install it from python.org."; read -r; exit 1; }
fi
source .venv/bin/activate
pip install -q -r requirements.txt
# localhost: only this computer can open the app; documents never leave it.
streamlit run app.py --server.address localhost
