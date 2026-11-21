@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
  set PYRUN=py -3
) else (
  set PYRUN=python
)
%PYRUN% -c "import sys; print(sys.version)"
%PYRUN% -c "import importlib.util; print('pywin32 OK' if importlib.util.find_spec('win32com') else 'pywin32 NAO instalado')"
start "" http://127.0.0.1:8765/diag
pause
