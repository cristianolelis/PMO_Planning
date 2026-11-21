@echo off
setlocal
cd /d "%~dp0"
echo ==================================================
echo HELPER LOCAL MPP - GERADOR EAP
echo ==================================================
where py >nul 2>nul
if %errorlevel%==0 (
  set PYRUN=py -3
  goto run
)
where python >nul 2>nul
if %errorlevel%==0 (
  set PYRUN=python
  goto run
)
echo [ERRO] Python nao foi encontrado no PATH.
echo Instale Python 3 e marque "Add Python to PATH".
pause
goto :eof
:run
echo Usando interpretador: %PYRUN%
%PYRUN% -m pip install --disable-pip-version-check pywin32
if %errorlevel% neq 0 (
  echo [ERRO] Falha ao instalar pywin32.
  pause
  goto :eof
)
echo Iniciando helper em http://127.0.0.1:8765 ...
%PYRUN% helper_mpp_server.py
echo.
echo O helper foi encerrado. Se houve erro acima, copie a mensagem completa.
pause
