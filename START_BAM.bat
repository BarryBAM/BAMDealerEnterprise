@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Creating BAM Python environment...
  python -m venv .venv
)
call ".venv\Scripts\activate.bat"
python -m pip install --disable-pip-version-check -r requirements.txt
python app\app.py
pause
