# Dental Statement Processor

Automate processing of dental patient balance statements with PDF parsing and tracking data merging.

## Quick Start

Run the web app:

**Mac/Linux:**
```bash
./run_app.sh
```

**Windows:**
```bash
run_app.bat
```

Or manually:
```bash
uv venv --python 3.12
uv pip install -r requirements.txt
streamlit run app.py
```

App opens at `http://localhost:8501`

## Usage

1. **Upload PDF**: Patient balance report from dental software
2. **Upload Excel** (optional): Previous month's tracking sheet
3. **Configure Columns**: 
   - Enter old sheet column names
   - Select which to transfer to new sheet
   - Preview output structure
4. **Process**: Click button and wait for results
5. **Download**: Timestamped Excel file with merged data

## What It Does

- Parses PDF patient balance reports
- Filters for outstanding balances (Patient Balance > 0 AND Family Balance > 0)
- Merges with tracking data from previous month
- Exports merged Excel with your notes/status preserved

## Deploy to Cloud

1. Push to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Deploy `app.py`
4. Access from anywhere

## Command Line Alternative

Edit `config.py` and run:
```bash
python process_statements.py
```

## Requirements

- Python 3.10, 3.11, or 3.12 (not 3.14)
- uv package manager
- Dependencies in `requirements.txt`
