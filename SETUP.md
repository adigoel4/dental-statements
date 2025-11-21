# Windows Setup Guide

Complete setup for a fresh Windows laptop in 7 easy steps.

---

## Step 1: Install Python

1. Go to **https://www.python.org/downloads/**
2. Download and run the installer
3. ⚠️ **CHECK THE BOX:** "Add Python to PATH"
4. Click "Install Now"

**Verify:**
```bash
python --version
```

---

## Step 2: Install UV

Open Command Prompt and run:

```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Close and reopen Command Prompt, then verify:
```bash
uv --version
```

---

## Step 3: Install VS Code (Optional)

1. Go to **https://code.visualstudio.com/**
2. Download and install
3. ✅ Check "Add to PATH" during installation

---

## Step 4: Get the Files

Download the `dental-statements` folder and save it to your Desktop:
```
C:\Users\YourName\Desktop\dental-statements\
```

---

## Step 5: Install Dependencies

Open Command Prompt in the `dental-statements` folder:
- Navigate to folder in File Explorer
- Click address bar → type `cmd` → Enter

Run these commands:
```bash
uv venv
.venv\Scripts\activate
uv pip install -r requirements.txt
```

You should see `(.venv)` in your prompt.

---

## Step 6: Add Your Files

1. Put your PDF in: `input\PAT_BAL_20251116.pdf`
2. Put old tracking sheet in: `input\Statements.xlsx` (if you have one)
3. Edit `config.py` → update PDF filename if needed

---

## Step 7: Run It!

```bash
.venv\Scripts\activate
python process_statements.py
```

**Results:** Open `output\Merged_Statements.xlsx`

---

## Monthly Workflow

Each month:

1. **Save last month's work:**
   - Copy `output\Merged_Statements.xlsx` → `input\Statements.xlsx`

2. **Add new PDF:**
   - Save to `input\` folder
   - Update filename in `config.py`

3. **Run:**
   ```bash
   cd Desktop\dental-statements
   .venv\Scripts\activate
   python process_statements.py
   ```

4. **Review:** Open `output\Merged_Statements.xlsx`

---

## Quick Reference

| Task | Command |
|------|---------|
| Activate environment | `.venv\Scripts\activate` |
| Run script | `python process_statements.py` |
| Check setup | `python check_setup.py` |
| Edit settings | Open `config.py` in Notepad |

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| "python is not recognized" | Reinstall Python with "Add to PATH" checked |
| "uv is not recognized" | Close and reopen Command Prompt |
| "File not found" | Put files in `input\` folder |
| "No module named..." | Activate venv, then `uv pip install -r requirements.txt` |

---

## What It Does

**Input:** PDF report (10,000+ patients) + old tracking sheet

**Output:** Excel with ~825 outstanding accounts

**Process:**
- Parses PDF → Filters to outstanding balances → Merges with your tracking data → Exports Excel

**Result:** Current balances + your old NOTES/STATUS merged together ✨

---

## File Structure

```
dental-statements/
├── input/                      # Put files here
│   ├── PAT_BAL_20251116.pdf
│   └── Statements.xlsx
├── output/                     # Results here
│   └── Merged_Statements.xlsx
├── config.py                   # Settings
├── process_statements.py       # Main script
└── SETUP.md                    # This file
```

---

## Success Checklist

- [ ] Output file created with ~825 records
- [ ] Balances match PDF
- [ ] Old NOTES/STATUS carried forward
- [ ] Dates show as MM/DD/YY

Done! 🎉
