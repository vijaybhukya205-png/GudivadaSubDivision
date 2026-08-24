@echo off
setlocal EnableExtensions
title Gudivada Dashboard - One Click Update

cd /d "%~dp0"

echo ============================================================
echo GUDIVADA SUB DIVISION - ONE CLICK DASHBOARD UPDATE
echo ============================================================
echo.

if not exist "Opened.xlsx" (
    echo ERROR: Opened.xlsx not found.
    echo Put the latest Opened.xlsx in:
    echo %CD%
    pause
    exit /b 1
)

if not exist "Closed.xlsx" (
    echo ERROR: Closed.xlsx not found.
    echo Put the latest Closed.xlsx in:
    echo %CD%
    pause
    exit /b 1
)

echo [1/4] Processing latest Opened.xlsx + Closed.xlsx...
python Gudivada_Net_Accounts_Updater_v4_RowColours.py

if errorlevel 1 (
    echo.
    echo ERROR: Net Accounts updater failed.
    echo Check the error above.
    pause
    exit /b 1
)

if not exist "Gudivada_Net_Accounts_Latest.xlsx" (
    echo.
    echo ERROR: Gudivada_Net_Accounts_Latest.xlsx was not generated.
    pause
    exit /b 1
)

echo.
echo [2/4] Net Accounts Excel updated successfully.
echo.

echo [3/4] Checking dashboard file...
if not exist "index.html" (
    echo ERROR: index.html not found.
    pause
    exit /b 1
)

echo Dashboard file found.
echo.
echo IMPORTANT:
echo If index.html is generated/updated by a separate dashboard
echo builder, run that builder before deployment.
echo.

echo [4/4] Deploying to Vercel production...
vercel --prod

if errorlevel 1 (
    echo.
    echo ERROR: Vercel deployment failed.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo LIVE DASHBOARD UPDATED SUCCESSFULLY
echo ============================================================
echo.
echo Permanent URL:
echo https://gudivada-sub-division-dashboard.vercel.app
echo.
pause
