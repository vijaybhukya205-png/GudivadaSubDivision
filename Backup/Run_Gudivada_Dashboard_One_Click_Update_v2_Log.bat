@echo off
setlocal EnableExtensions
title Gudivada Dashboard - One Click Update

cd /d "%~dp0"

set "LOG=%CD%\Dashboard_Update_Log.txt"

echo ============================================================ > "%LOG%"
echo GUDIVADA SUB DIVISION - DASHBOARD UPDATE >> "%LOG%"
echo Started: %date% %time% >> "%LOG%"
echo ============================================================ >> "%LOG%"
echo.

echo ============================================================
echo GUDIVADA SUB DIVISION - ONE CLICK DASHBOARD UPDATE
echo ============================================================
echo.
echo Log file:
echo %LOG%
echo.

echo [1/4] Checking input files...
echo [1/4] Checking input files... >> "%LOG%"

if not exist "Opened.xlsx" (
    echo ERROR: Opened.xlsx not found.
    echo ERROR: Opened.xlsx not found. >> "%LOG%"
    echo.
    echo The latest Opened.xlsx must be placed in:
    echo %CD%
    echo.
    goto FAILED
)

if not exist "Closed.xlsx" (
    echo ERROR: Closed.xlsx not found.
    echo ERROR: Closed.xlsx not found. >> "%LOG%"
    echo.
    echo The latest Closed.xlsx must be placed in:
    echo %CD%
    echo.
    goto FAILED
)

echo Opened.xlsx found.
echo Closed.xlsx found.
echo Opened.xlsx found. >> "%LOG%"
echo Closed.xlsx found. >> "%LOG%"

echo.
echo [2/4] Running Net Accounts updater...
echo [2/4] Running Net Accounts updater... >> "%LOG%"
echo.

python "Gudivada_Net_Accounts_Updater_v4_RowColours.py" >> "%LOG%" 2>&1
set "PYERR=%ERRORLEVEL%"

type "%LOG%" | more +5

if not "%PYERR%"=="0" (
    echo.
    echo ERROR: Python updater failed. Error code: %PYERR%
    echo ERROR: Python updater failed. Error code: %PYERR% >> "%LOG%"
    goto FAILED
)

if not exist "Gudivada_Net_Accounts_Latest.xlsx" (
    echo.
    echo ERROR: Gudivada_Net_Accounts_Latest.xlsx was not generated.
    echo ERROR: Gudivada_Net_Accounts_Latest.xlsx was not generated. >> "%LOG%"
    goto FAILED
)

echo.
echo SUCCESS: Net Accounts Excel updated.
echo SUCCESS: Net Accounts Excel updated. >> "%LOG%"

echo.
echo [3/4] Checking dashboard...
echo [3/4] Checking dashboard... >> "%LOG%"

if not exist "index.html" (
    echo ERROR: index.html not found.
    echo ERROR: index.html not found. >> "%LOG%"
    goto FAILED
)

echo index.html found.
echo index.html found. >> "%LOG%"

echo.
echo [4/4] Deploying to Vercel production...
echo [4/4] Deploying to Vercel production... >> "%LOG%"
echo.

vercel --prod >> "%LOG%" 2>&1
set "VERCELERR=%ERRORLEVEL%"

type "%LOG%" | more +1

if not "%VERCELERR%"=="0" (
    echo.
    echo ERROR: Vercel deployment failed. Error code: %VERCELERR%
    echo ERROR: Vercel deployment failed. Error code: %VERCELERR% >> "%LOG%"
    goto FAILED
)

echo.
echo ============================================================
echo SUCCESS - LIVE DASHBOARD DEPLOYED
echo ============================================================
echo Permanent URL:
echo https://gudivada-sub-division-dashboard.vercel.app
echo.
echo Full log:
echo %LOG%
echo ============================================================

echo SUCCESS - LIVE DASHBOARD DEPLOYED >> "%LOG%"
echo Permanent URL: https://gudivada-sub-division-dashboard.vercel.app >> "%LOG%"
echo Finished: %date% %time% >> "%LOG%"

echo.
echo Press any key to close this window...
pause >nul
exit /b 0

:FAILED
echo.
echo ============================================================
echo UPDATE FAILED
echo ============================================================
echo Full error log:
echo %LOG%
echo ============================================================
echo.
echo The window will remain open.
echo Press any key to close this window...
pause >nul
exit /b 1
