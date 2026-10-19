#!/bin/bash
# Double-click in Finder: sets up Python packages once, then opens the app in the browser.
cd "$(dirname "$0")" || exit 1
if [ ! -d .venv ]; then
  python3 -m venv .venv || { echo "Python 3 nahi mila. python.org se install karo."; read -r; exit 1; }
fi
source .venv/bin/activate
pip install -q -r requirements.txt
streamlit run app.py
