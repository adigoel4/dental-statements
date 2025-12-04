# Dental Statement Processor

A Streamlit web app that automates processing of dental patient balance statements.

---

## How to Make Changes (Windows)

This guide explains how to pull the repo, make changes, and have them appear on the live Streamlit app.

### Step 1: Install Git

1. Download Git from [git-scm.com/download/win](https://git-scm.com/download/win)
2. Run the installer. Accept all default options.
3. Restart your computer after installation.

### Step 2: Install Python

1. Download Python 3.12 from [python.org/downloads](https://python.org/downloads/)
2. Run the installer.
3. **Important:** Check the box that says "Add Python to PATH" before clicking Install.

### Step 3: Clone the Repository

1. Open **Command Prompt** (search "cmd" in Windows Start menu)
2. Navigate to where you want the project:
   ```
   cd Desktop
   ```
3. Clone the repo (replace with your actual GitHub URL):
   ```
   git clone https://github.com/YOUR_USERNAME/dental-statements.git
   ```
4. Enter the project folder:
   ```
   cd dental-statements
   ```

### Step 4: Install Dependencies

Run this command to install required packages:
```
pip install streamlit pandas pdfplumber openpyxl
```

### Step 5: Test Locally

Run the app on your computer first:
```
streamlit run app.py
```

This opens the app in your browser at `http://localhost:8501`. Test your changes here before pushing.

Press `Ctrl+C` in Command Prompt to stop the app.

### Step 6: Make Your Changes

1. Open `app.py` in any text editor (Notepad, VS Code, etc.)
2. Make your changes
3. Save the file
4. Test locally with `streamlit run app.py`

### Step 7: Push Changes to GitHub

After testing, push your changes to make them live:

```
git add .
git commit -m "Describe what you changed"
git push
```

**First time pushing?** Git will ask for your GitHub username and password. Use a Personal Access Token instead of your password:
1. Go to GitHub → Settings → Developer Settings → Personal Access Tokens
2. Generate a new token with "repo" permissions
3. Use this token as your password

### Step 8: See Changes on Production

Streamlit Cloud automatically detects changes to your GitHub repo.

1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Your app will show "Updating..." for 1-2 minutes
3. Refresh the page to see your changes live

---

## Quick Reference

| Task | Command |
|------|---------|
| Pull latest changes | `git pull` |
| Test locally | `streamlit run app.py` |
| Push changes | `git add .` then `git commit -m "message"` then `git push` |

---

## Troubleshooting

**"git is not recognized"**  
Restart Command Prompt after installing Git. If still broken, reinstall Git and check "Add to PATH".

**"python is not recognized"**  
Reinstall Python and check "Add Python to PATH" during installation.

**"streamlit is not recognized"**  
Run `pip install streamlit` again.

**Changes not showing on production?**  
1. Make sure you ran `git push`
2. Check Streamlit Cloud dashboard for errors
3. Wait 2-3 minutes for deployment

---

## What This App Does

- Parses PDF patient balance reports from dental software
- Filters for outstanding balances (Patient Balance > 0 AND Family Balance > 0)
- Merges with tracking data from previous month's Excel sheet
- Exports merged Excel file with notes/status preserved
