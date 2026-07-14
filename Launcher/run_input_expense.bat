@echo off
REM Run PAD flow: import expense rows for selected SQT values.
setlocal enableextensions enabledelayedexpansion

set "PROJECT_ROOT=%~dp0.."
for %%I in ("%PROJECT_ROOT%") do set "PROJECT_ROOT=%%~fI"

if exist "%PROJECT_ROOT%\.venv\Scripts\python.exe" (
  "%PROJECT_ROOT%\.venv\Scripts\python.exe" "%PROJECT_ROOT%\scripts\pad_launcher.py" --project-root "%PROJECT_ROOT%" --flow input_expense
  exit /b !errorlevel!
)

where py >nul 2>nul
if not errorlevel 1 (
  py -3 "%PROJECT_ROOT%\scripts\pad_launcher.py" --project-root "%PROJECT_ROOT%" --flow input_expense
  exit /b !errorlevel!
)

where python >nul 2>nul
if not errorlevel 1 (
  python "%PROJECT_ROOT%\scripts\pad_launcher.py" --project-root "%PROJECT_ROOT%" --flow input_expense
  exit /b !errorlevel!
)

echo [ERROR] Python not found. Cannot launch PAD flow.
exit /b 2
