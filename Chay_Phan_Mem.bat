@echo off
REM ============================================================
REM   Tro Ly Quyet Toan RPA - Khoi dong phan mem
REM   Chi can double-click file nay la mo duoc phan mem.
REM ============================================================
setlocal enableextensions
cd /d "%~dp0"
title Tro Ly Quyet Toan RPA

REM --- 1) Uu tien moi truong ao .venv da cai san trong du an ---
if exist "%~dp0.venv\Scripts\pythonw.exe" (
  start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0main.py"
  exit /b 0
)
if exist "%~dp0.venv\Scripts\python.exe" (
  "%~dp0.venv\Scripts\python.exe" "%~dp0main.py"
  exit /b 0
)

REM --- 2) Khong co .venv: tim Python he thong ---
set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY (
  where python >nul 2>nul && set "PY=python"
)
if not defined PY (
  echo.
  echo [LOI] Khong tim thay Python tren may tinh.
  echo Hay cai Python 3.11 tro len tai: https://www.python.org/downloads/
  echo Khi cai nho TICH CHON o "Add Python to PATH".
  echo.
  pause
  exit /b 1
)

REM --- 3) Kiem tra thu vien; thieu thi tu cai (chi lan dau) ---
%PY% -c "import PySide6, watchdog, openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Dang cai dat thu vien can thiet ^(chi lan dau, can Internet^), vui long cho...
  %PY% -m pip install -r requirements.txt
  if errorlevel 1 (
    echo.
    echo [LOI] Cai dat thu vien that bai. Kiem tra ket noi Internet roi chay lai.
    echo.
    pause
    exit /b 1
  )
)

REM --- 4) Mo phan mem (uu tien khong kem cua so den) ---
set "PYW="
where pythonw >nul 2>nul && set "PYW=pythonw"
if not defined PYW (
  where pyw >nul 2>nul && set "PYW=pyw"
)
if defined PYW (
  start "" %PYW% "%~dp0main.py"
) else (
  %PY% "%~dp0main.py"
)

endlocal
exit /b 0
