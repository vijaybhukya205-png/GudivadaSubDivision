@echo off
setlocal EnableExtensions
title Gudivada Sub Division - GitHub One Click Update
cd /d "%~dp0"
set "LOG=%CD%\Dashboard_Update_Log.txt"

echo ============================================================ > "%LOG%"
echo GUDIVADA SUB DIVISION - GITHUB DASHBOARD UPDATE >> "%LOG%"
echo Started: %date% %time% >> "%LOG%"
echo ============================================================ >> "%LOG%"

echo [1/6] Checking input files...
if not exist "Opened.xlsx" goto FAIL_OPEN
if not exist "Closed.xlsx" goto FAIL_CLOSED

echo [2/6] Running Net Accounts updater...
python "Gudivada_Net_Accounts_Updater_v4_RowColours.py" >> "%LOG%" 2>&1
if errorlevel 1 goto FAIL_PY
if not exist "Gudivada_Net_Accounts_Latest.xlsx" goto FAIL_LATEST

echo [3/6] Updating Health Card index.html...
python "Update_Gudivada_Live_Dashboard_From_Excel.py" >> "%LOG%" 2>&1
if errorlevel 1 goto FAIL_DASH
if not exist "index.html" goto FAIL_INDEX

echo [4/6] Checking GitHub connection...
git --version >> "%LOG%" 2>&1
if errorlevel 1 goto FAIL_GIT
git remote get-url origin >> "%LOG%" 2>&1
if errorlevel 1 goto FAIL_REMOTE

echo [5/6] Committing dashboard update...
git add index.html Gudivada_Net_Accounts_Latest.xlsx Update_Gudivada_Live_Dashboard_From_Excel.py >> "%LOG%" 2>&1
if errorlevel 1 goto FAIL_ADD

git diff --cached --quiet
if errorlevel 0 goto PUSH

git commit -m "Gudivada Dashboard update %date% %time%" >> "%LOG%" 2>&1
if errorlevel 1 goto FAIL_COMMIT

:PUSH
echo [6/6] Publishing to GitHub...
git branch -M main >> "%LOG%" 2>&1
git push -u origin main >> "%LOG%" 2>&1
if errorlevel 1 goto FAIL_PUSH

echo.
echo ============================================================
echo SUCCESS - GUDIVADA HEALTH CARD UPDATED
echo ============================================================
echo Repository:
echo https://github.com/vijaybhukya205-png/GudivadaSubDivision
echo GitHub Pages:
echo https://vijaybhukya205-png.github.io/GudivadaSubDivision/
echo ============================================================
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
echo ERROR: Latest Excel was not generated.>>"%LOG%"
goto FAIL
:FAIL_DASH
echo ERROR: Health Card index.html update failed.>>"%LOG%"
goto FAIL
:FAIL_INDEX
echo ERROR: index.html not found.>>"%LOG%"
goto FAIL
:FAIL_GIT
echo ERROR: Git is not installed or unavailable.>>"%LOG%"
goto FAIL
:FAIL_REMOTE
echo ERROR: GitHub origin is not configured.>>"%LOG%"
goto FAIL
:FAIL_ADD
echo ERROR: Git add failed.>>"%LOG%"
goto FAIL
:FAIL_COMMIT
echo ERROR: Git commit failed.>>"%LOG%"
goto FAIL
:FAIL_PUSH
echo ERROR: GitHub push failed.>>"%LOG%"
goto FAIL

:FAIL
echo.
echo ============================================================
echo UPDATE FAILED - SEE Dashboard_Update_Log.txt
echo ============================================================
pause
exit /b 1
