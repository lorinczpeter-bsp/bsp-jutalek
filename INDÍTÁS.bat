@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 (
  echo Python 3.11 vagy ujabb szukseges. Telepites: https://www.python.org/downloads/
  pause
  exit /b 1
)
if not exist ".venv\Scripts\python.exe" (
  py -3 -m venv .venv
  if errorlevel 1 goto :error
)
".venv\Scripts\python.exe" -c "import sys; assert sys.version_info >= (3,11), 'Python 3.11 vagy ujabb szukseges'"
if errorlevel 1 goto :error
if not exist ".venv\ready-v1.txt" (
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 goto :error
  echo ready > ".venv\ready-v1.txt"
)
".venv\Scripts\python.exe" -m streamlit run app.py --server.address 127.0.0.1 --browser.gatherUsageStats false
if errorlevel 1 goto :error
exit /b 0
:error
echo.
echo Az inditas sikertelen. A fenti hibauzenet mutatja az okat.
pause
exit /b 1
