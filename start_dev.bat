@echo off
setlocal
cd /d "%~dp0"

if not exist .venv\Scripts\python.exe (
  echo Creando entorno virtual...
  py -3.13 -m venv .venv 2>nul || py -3 -m venv .venv
)

call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt

python main.py
pause
