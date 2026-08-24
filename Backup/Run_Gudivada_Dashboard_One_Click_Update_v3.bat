@echo off
setlocal EnableExtensions
title Gudivada Dashboard - One Click Update
cd /d "%~dp0"
set "LOG=%CD%\Dashboard_Update_Log.txt"

echo ============================================================ > "%LOG%"
echo GUDIVADA SUB DIVISION - DASHBOARD UPDATE >> "%LOG%"
echo Started: %date% %time% >> "%LOG%"
echo ============================================================ >> "%LOG%"

echo [1/5] Checking input files...
if not exist "Opened.xlsx" goto FAIL_OPEN
if not exist "Closed.xlsx" goto FAIL_CLOSED

echo [2/5] Running Net Accounts updater...
python "Gudivada_Net_Accounts_Updater_v4_RowColours.py" >> "%LOG%" 2>&1
if errorlevel 1 goto FAIL_PY
if not exist "Gudivada_Net_Accounts_Latest.xlsx" goto FAIL_LATEST

echo [3/5] Updating index.html from latest Excel...
python "Update_Gudivada_Live_Dashboard_From_Excel.py" >> "%LOG%" 2>&1
if errorlevel 1 goto FAIL_DASH
if not exist "index.html" goto FAIL_INDEX

echo [4/5] Deploying latest dashboard to Vercel...
vercel --prod >> "%LOG%" 2>&1
if errorlevel 1 goto FAIL_VERCEL

echo [5/5] SUCCESS
echo.
echo ============================================================
echo LIVE DASHBOARD UPDATED WITH LATEST DATA
echo ============================================================
echo https://gudivada-sub-division-dashboard.vercel.app
echo.
echo Log: %LOG%
echo.
pause
exit /b 0

:FAIL_OPEN
echo ERROR: Opened.xlsx not found.>>"%LOG%"
goto FAIL
:FAIL_CLOSED
echo ERROR: Closed.xlsx not found.>>"%LOG%"
goto FAIL
:FAIL_PY
echo ERROR: Net Accounts updater failed.>>"%LOG%"
goto FAIL
:FAIL_LATEST
echo ERROR: Gudivada_Net_Accounts_Latest.xlsx not generated.>>"%LOG%"
goto FAIL
:FAIL_DASH
echo ERROR: Dashboard data update failed.>>"%LOG%"
goto FAIL
:FAIL_INDEX
echo ERROR: index.html not found.>>"%LOG%"
goto FAIL
:FAIL_VERCEL
echo ERROR: Vercel deployment failed.>>"%LOG%"
goto FAIL
:FAIL
echo.
echo ============================================================
echo UPDATE FAILED - SEE LOG
echo ============================================================
echo %LOG%
echo ============================================================
pause
exit /b 1
