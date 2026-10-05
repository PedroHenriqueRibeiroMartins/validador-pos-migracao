@echo off
cd /d "%~dp0"
echo Preparando o Validador Pos-Migracao...
python -m pip install -q -r requirements.txt
python -m streamlit run app.py
pause
