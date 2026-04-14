@echo off
cd /d "%~dp0"

REM Try python in PATH first, then fall back to default install location
where python >nul 2>&1
if %ERRORLEVEL% equ 0 (
    python app.py -d "%USERPROFILE%\pikaraoke-songs" -L lo
) else (
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" app.py -d "%USERPROFILE%\pikaraoke-songs" -L lo
)
pause
