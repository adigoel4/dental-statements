#!/bin/bash
# Quick start script for the Streamlit app

echo "🦷 Starting Dental Statement Automation..."
echo ""

# Check if uv is installed
if ! command -v uv &> /dev/null
then
    echo "❌ Error: uv is not installed"
    echo "Install it from: https://github.com/astral-sh/uv"
    exit 1
fi

# Ensure we're using Python 3.12 (3.14 not yet supported by Streamlit)
if [ ! -d ".venv" ]; then
    echo "📦 Creating virtual environment with Python 3.12..."
    uv venv --python 3.12 || uv venv --python 3.11 || uv venv --python 3.10
fi

# Install dependencies if needed
echo "📦 Checking dependencies..."
uv pip install -r requirements.txt --quiet

# Run the Streamlit app
echo "🚀 Starting Streamlit app..."
echo "   App will open in your browser at http://localhost:8501"
echo ""
streamlit run app.py

