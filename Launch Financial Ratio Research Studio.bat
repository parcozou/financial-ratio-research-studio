@echo off
setlocal

cd /d "%~dp0"
title Financial Ratio Research Studio

echo Starting Financial Ratio Research Studio...
echo.

where python >nul 2>nul
if %errorlevel%==0 goto run_python

where py >nul 2>nul
if %errorlevel%==0 goto run_py

if exist "%USERPROFILE%\anaconda3\python.exe" goto run_anaconda

echo Python could not be found on this computer.
echo Please install Python or Anaconda first, then try again.
echo.
pause
exit /b 1

:run_python
python -m streamlit run streamlit_app.py
goto handle_error

:run_py
py -3 -m streamlit run streamlit_app.py
goto handle_error

:run_anaconda
"%USERPROFILE%\anaconda3\python.exe" -m streamlit run streamlit_app.py
goto handle_error

:handle_error
if %errorlevel%==0 exit /b 0

echo.
echo The app could not start.
echo If Streamlit or the project dependencies are missing, run:
echo     pip install -r requirements.txt
echo.
pause
exit /b 1
