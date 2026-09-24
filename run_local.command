#!/bin/bash
cd "$(dirname "$0")"
pip3 install -q matplotlib openpyxl 2>/dev/null
python3 local_app.py
