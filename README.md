# BS Weekly Lead Manager

Lead management tool for B+S Commercial Kitchens — generates 20 Australian
hospitality business contacts per week for sales outreach.

---

## Quick Setup

### 1. Install dependencies

```
cd C:\Users\malej\bs_lead_manager
pip install -r requirements.txt
```

### 2. Set your Anthropic API key

**PowerShell:**
```powershell
$env:ANTHROPIC_API_KEY = "sk-ant-..."
```

**Command Prompt (CMD):**
```
set ANTHROPIC_API_KEY=sk-ant-...
```

Get your key at: https://console.anthropic.com

### 3. Run the app

```
python app.py
```

Then open: **http://localhost:5000**

---

## Features

- **Generate leads** — Claude AI produces realistic Australian hospitality
  business contacts (restaurants, hotels, casinos, catering, food trucks)
- **No repeats** — already-seen businesses are never generated again
- **Priority sorting** — high-value targets (5-star hotels, casinos, corporate
  catering) shown first with a ★ HIGH badge
- **Mark sent** — track which businesses you've contacted
- **Excel export** — one sheet per category + summary sheet
- **CSV export** — backup option
- **History** — past weeks collapsible with per-week export

## File structure

```
bs_lead_manager/
  app.py          # Flask backend + Anthropic API + Excel export
  database.py     # SQLite persistence
  requirements.txt
  leads.db        # created automatically on first run
  templates/
    index.html    # single-page frontend
  exports/        # Excel files saved here
  README.md
```
