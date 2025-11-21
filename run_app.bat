@echo off
REM Quick start script for Windows users with uv

echo.
echo 🦷 Starting Dental Statement Automation...
echo.

REM Check if uv is installed
where uv >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo ❌ Error: uv is not installed
    echo Install from: https://github.com/astral-sh/uv
    pause
    exit /b 1
)

REM Ensure we're using Python 3.12 (3.14 not yet supported by Streamlit)
if not exist ".venv" (
    echo 📦 Creating virtual environment with Python 3.12...
    uv venv --python 3.12 || uv venv --python 3.11 || uv venv --python 3.10
)

REM Install dependencies
echo 📦 Checking dependencies...
uv pip install -r requirements.txt --quiet

REM Run the Streamlit app
echo.
echo 🚀 Starting Streamlit app...
echo    App will open in your browser at http://localhost:8501
echo.
streamlit run app.py

