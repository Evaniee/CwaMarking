@echo off
if not exist .venv (
    echo Creating Virtual Environment, this might take a while!
    python -m venv .venv
    .venv\Scripts\pip3 install setuptools
    .venv\Scripts\pip3 install tk
    .venv\Scripts\pip3 install pywin32
    .venv\Scripts\pip3 install os-sys
) else (
    echo Using existing venv. Delete .venv folder an re-run to re-create venv
)
.venv\Scripts\python Excel.py
