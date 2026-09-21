@echo off
chcp 65001 >nul
setlocal

cd /d "%~dp0"

echo Building BuyerExcessReports.exe...
echo.

set "PYTHON_CMD="
py -3 --version >nul 2>nul
if not errorlevel 1 set "PYTHON_CMD=py -3"
if not defined PYTHON_CMD (
    python --version >nul 2>nul
    if not errorlevel 1 set "PYTHON_CMD=python"
)

if not defined PYTHON_CMD (
    echo No Python runtime found.
    echo Please install Python 3, or make sure python.exe is available in PATH.
    echo.
    pause
    exit /b 1
)

echo Using Python command: %PYTHON_CMD%
echo.

%PYTHON_CMD% -m pip install -r requirements.txt pyinstaller
if errorlevel 1 (
    echo.
    echo Build failed while installing dependencies.
    pause
    exit /b 1
)

if exist build rmdir /s /q build
if exist release\BuyerExcessReports rmdir /s /q release\BuyerExcessReports

%PYTHON_CMD% -m PyInstaller ^
  --clean ^
  --onedir ^
  --console ^
  --hidden-import tkinter ^
  --hidden-import tkinter.ttk ^
  --hidden-import tqdm ^
  --name BuyerExcessReports ^
  --distpath release ^
  --workpath build ^
  --specpath build ^
  generate_buyer_excess_report.py
if errorlevel 1 (
    echo.
    echo Build failed while creating the exe.
    pause
    exit /b 1
)

if not exist release\BuyerExcessReports\input mkdir release\BuyerExcessReports\input
if not exist release\BuyerExcessReports\input\AVTC mkdir release\BuyerExcessReports\input\AVTC
if not exist release\BuyerExcessReports\input\RAKEN mkdir release\BuyerExcessReports\input\RAKEN
if not exist release\BuyerExcessReports\output mkdir release\BuyerExcessReports\output
if not exist release\BuyerExcessReports\output\AVTC mkdir release\BuyerExcessReports\output\AVTC
if not exist release\BuyerExcessReports\output\RAKEN mkdir release\BuyerExcessReports\output\RAKEN
copy /Y "Windows執行檔(exe)使用說明.txt" "release\BuyerExcessReports\Windows執行檔(exe)使用說明.txt" >nul

echo.
echo Done.
echo Release folder:
echo %cd%\release\BuyerExcessReports
echo.
pause
