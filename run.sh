#/bin/bash
python3 -m venv env
source env/bin/activate
pip install -r dependencies.txt
python3 main.py
deactivate
