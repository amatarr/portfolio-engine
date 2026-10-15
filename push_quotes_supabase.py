"""
Reads live CQGXL quotes straight off the open Position_Analysis.xlsx
(via xlwings/COM, same mechanism xlwings_refresh.py uses) and pushes
them to Supabase, so the deployed Streamlit app -- which can't see this
machine's Excel -- can read live-ish prices instead.

Run this in a loop (or re-run periodically) while Excel/CQGXL is open.
Needs supabase_url and supabase_service_key in .streamlit/secrets.toml.
"""

import time
import tomllib

import xlwings as xw
from supabase import create_client

from vol_calculator.xlwings_sync import xw_read_live_quotes

WORKBOOK_PATH = "Position_Analysis.xlsm"
TABLE_NAME = "Quotes"
PUSH_INTERVAL_SECONDS = 5


def _load_secrets():
    with open(".streamlit/secrets.toml", "rb") as f:
        return tomllib.load(f)


def read_live_quotes():
    """Thin wrapper over the shared watchlist reader -- xlwings_refresh.py prices the book from the same quotes."""

    wb = xw.Book(WORKBOOK_PATH)
    return xw_read_live_quotes(wb)


def push_once(client, quotes):

    import datetime

    if not quotes:
        return 0

    now = datetime.datetime.now(datetime.timezone.utc).isoformat()

    rows = [
        {"contract": q["contract"], "price": q["price"], "updated_at": now}
        for q in quotes
    ]

    # One network round-trip for every contract instead of one per
    # contract -- at a short poll interval, 18 sequential calls (~100ms
    # each) was eating a meaningful chunk of the cycle budget on its own.
    client.table(TABLE_NAME).upsert(rows, on_conflict="contract").execute()

    return len(quotes)


def main():

    secrets = _load_secrets()
    client = create_client(secrets["supabase_url"], secrets["supabase_service_key"])

    print(f"Pushing live CQGXL quotes to Supabase every {PUSH_INTERVAL_SECONDS}s. Ctrl+C to stop.")

    while True:
        try:
            quotes = read_live_quotes()
            n = push_once(client, quotes)
            print(f"Pushed {n} quote(s)")
        except Exception as e:
            # Keep the loop alive through transient COM/network errors --
            # one bad cycle shouldn't stop future ones from refreshing.
            print(f"Cycle failed, will retry next interval: {type(e).__name__}: {e}")

        time.sleep(PUSH_INTERVAL_SECONDS)


if __name__ == "__main__":
    main()
