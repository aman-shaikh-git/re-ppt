@echo off

setlocal
:: This line forces the terminal to look in the folder where the .bat is saved
cd /d "%~dp0"

title RePPT Launcher

echo Checking environment...

:: 1. Try standard python
python --version >nul 2>&1
if %errorlevel% == 0 (set PY_CMD=python & goto :FOUND)

:: 2. Try the 'py' launcher (Standard Windows Python install)
py --version >nul 2>&1
if %errorlevel% == 0 (set PY_CMD=py & goto :FOUND)

:: 3. Hard-check common Anaconda locations (The Failsafe)
if exist "%USERPROFILE%\anaconda3\python.exe" (set PY_CMD="%USERPROFILE%\anaconda3\python.exe" & goto :FOUND)
if exist "%LOCALAPPDATA%\anaconda3\python.exe" (set PY_CMD="%LOCALAPPDATA%\anaconda3\python.exe" & goto :FOUND)

echo ERROR: RePPT couldn't find Python.
echo --------------------------------------------------
echo Please do one of the following:
echo 1. Install Python from python.org (Check 'Add to PATH').
echo 2. Run this file from the 'Anaconda Prompt'.
echo --------------------------------------------------
pause
exit /b

:FOUND
echo Using %PY_CMD%...

echo Installing/Updating dependencies (this may take a moment)...
%PY_CMD% -m pip install --upgrade pip --quiet
%PY_CMD% -m pip install -r requirements.txt --quiet

echo Launching RePPT...
%PY_CMD% -m streamlit run ui_wrapper_v0_2.py

pause