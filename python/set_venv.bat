@echo off
setlocal

set PYTHONHOME=C:\Users\nkmanager\AppData\Local\Programs\Python\Python310
set PATH=%PYTHONHOME%;%PYTHONHOME%\DLLs;%PYTHONHOME%\Scripts;%PATH%

if not exist %PYTHONHOME% (
    echo.
    echo    ERROR !   %PYTHONHOME% not found.
    echo    Check this Scripts
    echo.
    goto :eof
)

if not exist .vscode (mkdir .vscode)
call :genJSON > .vscode\settings.json
call :genBat    > tmp.bat
powershell Start-Process tmp.bat -Verb runas -Wait
del tmp.bat
pip install -r requirements.txt -t .venv\Lib\site-packages 1>nul 2>nul
goto :eof

:genJSON
    echo {
    echo     "terminal.integrated.defaultProfile.windows": "Command Prompt",
    echo     "python.envFile": "${workspaceFolder}\\.venv",
    echo     "python.defaultInterpreterPath": "${workspaceFolder}\\.venv\\Scripts\\python.exe"
    echo }
    exit /b

:genBat
    echo @echo off
    echo cd "%~dp0"
    echo "%PYTHONHOME%\python" -m venv --system-site-packages --symlinks --without-pip --clear .venv
   exit /b
