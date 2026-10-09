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

from vol_calculator.cqg_quotes import _convert_cqg_symbol

WORKBOOK_PATH = "Position_Analysis.xlsx"
TABLE_NAME = "Quotes"
PUSH_INTERVAL_SECONDS = 10

# (commodity, symbol_col, price_col) -- matches the Home sheet's own
# CORN/SOYBEANS/WHEAT layout (rows 3-8 under each header).
_BLOCKS = [
    ("Corn", 1, 2),
    ("Soybean", 5, 6),
    ("SRW Wheat", 9, 10),
]
_ROWS = range(3, 9)


def _load_secrets():
    with open(".streamlit/secrets.toml", "rb") as f:
        return tomllib.load(f)


def read_live_quotes():

    wb = xw.Book(WORKBOOK_PATH)
    home = wb.sheets["Home"]

    quotes = []

    for commodity, symbol_col, price_col in _BLOCKS:
        for row in _ROWS:
            symbol = home.range((row, symbol_col)).value
            price = home.range((row, price_col)).value

            if not symbol or price is None:
                continue

            try:
                price = float(price)
            except (TypeError, ValueError):
                # CQGXL cells show text like "Disconnected from server." or
                # "#N/A" during a transient connection drop instead of a
                # number -- skip just this one contract for this cycle
                # rather than crashing the whole loop over it.
                print(f"Skipping {symbol}: non-numeric price {price!r}")
                continue

            quotes.append({
                "contract": _convert_cqg_symbol(str(symbol)),
                "price": price,
                "commodity": commodity,
            })

    return quotes


def push_once(client, quotes):

    import datetime

    now = datetime.datetime.now(datetime.timezone.utc).isoformat()

    for q in quotes:
        client.table(TABLE_NAME).upsert({
            "contract": q["contract"],
            "price": q["price"],
            "updated_at": now,
        }, on_conflict="contract").execute()

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
