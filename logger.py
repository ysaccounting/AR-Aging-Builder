import os
import requests
from datetime import datetime, timezone


def log_report_run(as_of_date, row_count: int, grand_total: float, elapsed_seconds: float):
    """Log a report generation event to Supabase ar_aging_logs table."""
    try:
        url = os.environ["SUPABASE_URL"] + "ar_aging_logs"
        key = os.environ["SUPABASE_KEY"]
        headers = {
            "apikey": key,
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
            "Prefer": "return=minimal",
        }
        payload = {
            "timestamp":       datetime.now(timezone.utc).isoformat(),
            "as_of_date":      str(as_of_date),
            "invoice_count":   row_count,
            "grand_total":     round(grand_total, 2),
            "elapsed_seconds": round(elapsed_seconds, 2),
        }
        response = requests.post(url, json=payload, headers=headers, timeout=5)
        response.raise_for_status()
    except Exception as e:
        # Logging failure should never break the app
        print(f"Supabase logging error: {e}")
