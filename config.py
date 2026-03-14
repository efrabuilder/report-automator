"""
config.py — Edit this file to customize your report.
All settings are in the CONFIG dictionary.
"""

CONFIG = {

    # ── Input ──────────────────────────────────────────────────────────────────
    "data_source": "data/sample_data.csv",   # path to your CSV or Excel file
    "report_name": "Monthly Sales Report",

    # ── Charts ─────────────────────────────────────────────────────────────────
    # Each chart needs: type ("bar" | "line" | "pie"), title, and column names.
    "charts": [
        {
            "type":  "bar",
            "title": "Revenue by Region",
            "x":     "Region",
            "y":     "Revenue"
        },
        {
            "type":   "line",
            "title":  "Monthly Trend",
            "x":      "Month",
            "y_cols": ["Revenue", "Expenses"]
        },
        {
            "type":  "pie",
            "title": "Sales Distribution",
            "label": "Category",
            "value": "Revenue"
        }
    ],

    # ── Schedule ───────────────────────────────────────────────────────────────
    # mode: "once" | "daily" | "weekly" | "interval"
    "schedule": {
        "mode": "daily",
        "at":   "08:00",     # for daily / weekly
        "day":  "monday",    # for weekly only
        "hours": 6           # for interval only
    },

    # ── Email ──────────────────────────────────────────────────────────────────
    "email": {
        "enabled":    False,          # set True to activate
        "sender":     "you@gmail.com",
        "password":   "your_app_password",   # use an App Password, not your real password
        "smtp_host":  "smtp.gmail.com",
        "smtp_port":  587,
        "recipients": ["boss@company.com", "team@company.com"],
        "subject":    "Automated Report — Ready",
        "body":       "Hi,\n\nPlease find this period's automated report attached.\n\nBest,\nReport Automator"
    }
}
