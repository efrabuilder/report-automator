# 📊 Report Automator

Automated report generator that reads CSV/Excel data, produces styled Excel + PDF reports with charts, and sends them via email on a configurable schedule.

Built by **Efraín Rojas Artavia**

---

## Features

- ✅ Reads any **CSV or Excel** file
- ✅ Generates **bar, line and pie charts** automatically
- ✅ Exports a styled **Excel report** (Summary + Data + Charts sheets)
- ✅ Exports a styled **PDF report** with charts embedded
- ✅ Sends both files via **email** (Gmail / SMTP)
- ✅ Runs on a **schedule** — once, daily, weekly, or every N hours

---

## Project Structure

```
report-automator/
├── report_automator.py   # Main script
├── config.py             # All settings (edit this)
├── requirements.txt      # Dependencies
├── data/
│   └── sample_data.csv   # Example input data
└── output/               # Generated reports (auto-created)
    └── charts/           # Chart images (auto-created)
```

---

## Quick Start

```bash
# 1. Clone the repo
git clone https://github.com/efrabuilder/report-automator.git
cd report-automator

# 2. Install dependencies
pip install -r requirements.txt

# 3. Add your data file to data/
# (or use the included sample_data.csv)

# 4. Edit config.py to match your columns and preferences

# 5. Run
python report_automator.py
```

Reports will appear in the `output/` folder.

---

## Configuration

All settings live in `config.py`:

### Data source
```python
"data_source": "data/your_file.csv",  # CSV or Excel
"report_name": "My Report",
```

### Charts
```python
"charts": [
    { "type": "bar",  "title": "Revenue by Region", "x": "Region", "y": "Revenue" },
    { "type": "line", "title": "Monthly Trend", "x": "Month", "y_cols": ["Revenue", "Expenses"] },
    { "type": "pie",  "title": "Distribution",  "label": "Category", "value": "Revenue" },
]
```

### Schedule
```python
"schedule": {
    "mode": "daily",   # "once" | "daily" | "weekly" | "interval"
    "at":   "08:00",   # time for daily/weekly
    "day":  "monday",  # for weekly
    "hours": 6         # for interval
}
```

### Email
```python
"email": {
    "enabled":    True,
    "sender":     "you@gmail.com",
    "password":   "your_app_password",  # Gmail App Password
    "recipients": ["recipient@email.com"],
}
```

> **Gmail tip:** Use an [App Password](https://myaccount.google.com/apppasswords) instead of your real password.

---

## Tech Stack

| Tool | Purpose |
|------|---------|
| `pandas` | Data loading & statistics |
| `matplotlib` | Chart generation |
| `openpyxl` | Excel report creation |
| `fpdf2` | PDF report generation |
| `schedule` | Job scheduling |
| `smtplib` | Email delivery |

---

## License

MIT — free to use and modify.
