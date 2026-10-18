@echo off
setlocal
cd /d "%~dp0"

echo.
echo  ================================================
echo   Thesis AI Detector
echo  ================================================
echo.

where python >nul 2>nul
if errorlevel 1 goto NoPython
goto HavePython

:NoPython
where py >nul 2>nul
if errorlevel 1 goto MissingPython
set PYTHON=py
goto CheckVenv

:MissingPython
echo  [X] Python was not found on your computer.
echo.
echo  Please install Python first:
echo    1. Go to https://www.python.org/downloads/
echo    2. Download Python 3.10 or newer.
echo    3. IMPORTANT: tick Add python.exe to PATH during install.
echo    4. Then double-click this file again.
echo.
pause
exit /b 1

:HavePython
set PYTHON=python

:CheckVenv
if exist ".venv\Scripts\python.exe" goto HaveVenv
echo  First-time setup: creating a private folder...
echo  This takes a few minutes only once.
%PYTHON% -m venv .venv
if errorlevel 1 goto VenvFail
goto HaveVenv

:VenvFail
echo  [X] Could not create the environment.
pause
exit /b 1

:HaveVenv
call ".venv\Scripts\activate.bat"

echo  Checking libraries - first time takes a few minutes...
%PYTHON% -m pip install --quiet --disable-pip-version-check --upgrade pip >nul 2>nul
%PYTHON% -m pip install --quiet --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto PipFail

if not "%~1"=="" goto HaveArg
goto AskFile

:PipFail
echo  [X] Could not install libraries. Check your internet connection,
echo      then double-click this file again.
pause
exit /b 1

:HaveArg
set THESIS=%~1
goto CheckFile

:AskFile
echo.
echo  Tip: next time, drag your thesis file onto Run.bat.
echo.
set /p "THESIS=Type the file name and press Enter: "
:TrimLoop
if not "%THESIS:~-1%"==" " goto TrimDone
set "THESIS=%THESIS:~0,-1%"
goto TrimLoop
:TrimDone
if "%THESIS%"=="" goto NoFile
goto CheckFile

:NoFile
echo  [X] No file given. Put your file next to Run.bat and try again.
pause
exit /b 1

:CheckFile
if exist "%THESIS%" goto RunIt
echo  [X] File not found: "%THESIS%"
echo      Check the spelling, or type the full path, then try again.
pause
exit /b 1

:RunIt
echo.
echo  Analysing your document. Please wait a few minutes.
echo.
%PYTHON% ai_detector.py "%THESIS%" --report report.html --csv results.csv
if errorlevel 1 goto RunFail

echo.
echo  ================================================
echo   DONE! Open report.html to see your results.
echo  ================================================
echo.
pause
exit /b 0

:RunFail
echo.
echo  [X] Something went wrong. See README.md section 4.
pause
exit /b 1
