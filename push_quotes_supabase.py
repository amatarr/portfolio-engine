"""
Reads live CQGXL quotes off Live_Quotes.xls (via xlwings/COM) and pushes
them to Supabase, so the deployed Streamlit app -- which can't see this
machine's Excel -- can read live-ish prices instead.

The watchlist keeps itself current: once a day (and at start-up) the
next CONTRACTS_PER_COMMODITY contracts still trading are written into
the Quotes sheet as CQGXLContractData formulas, so an expired contract
drops off and the next listed month takes its place. The workbook is
created if it doesn't exist yet. Supabase keeps an expired contract's
last row; the app hides it except to price a position still held on it.

Run this in a loop (Start Live Quotes.bat) while Excel/CQGXL is open.
Needs supabase_url and supabase_service_key in .streamlit/secrets.toml.

    py push_quotes_supabase.py                         # push every 5 s
    py push_quotes_supabase.py --dry-run --once        # roll + read, no Supabase
    py push_quotes_supabase.py --dry-run --once --today 2026-12-15
"""

import argparse
import datetime
import os
import time
import tomllib

import xlwings as xw

from vol_calculator.contracts import active_contracts
from vol_calculator.cqg_quotes import _convert_cqg_symbol, to_cqg_symbol

WORKBOOK_PATH = "Live_Quotes.xls"
SHEET_NAME = "Quotes"
TABLE_NAME = "Quotes"
PUSH_INTERVAL_SECONDS = 5
CONTRACTS_PER_COMMODITY = 6
XLS_FILE_FORMAT = 56  # Excel 97-2003 workbook

# "Pushed by" cell on the Quotes sheet: the name typed next to the label
# goes out with every price (the Quotes table's source column), and the
# app shows it as "Live Quotes (<name>)".
PUSHED_BY_LABEL_CELL = "D1"
PUSHED_BY_CELL = "E1"

# One named block per commodity on the Quotes sheet: a symbol column and
# a price column, one row per contract.
WATCHLIST_NAMES = {
    "Corn": "Watchlist_Corn",
    "Soybean": "Watchlist_Soybean",
    "SRW Wheat": "Watchlist_Wheat",
}


def _load_secrets():
    with open(".streamlit/secrets.toml", "rb") as f:
        return tomllib.load(f)


def open_workbook():
    """The open Live_Quotes.xls (opening or creating it as needed)."""

    path = os.path.abspath(WORKBOOK_PATH)
    if os.path.exists(path):
        return xw.Book(path)

    app = xw.apps.active if xw.apps.count else xw.App()
    wb = app.books.add()
    wb.sheets[0].name = SHEET_NAME
    _save(wb, path)
    return wb


def _save(wb, path=None):
    alerts = wb.app.display_alerts
    wb.app.display_alerts = False  # no .xls compatibility prompt
    try:
        if path:
            wb.api.SaveAs(path, XLS_FILE_FORMAT)
        else:
            wb.api.Save()
    finally:
        wb.app.display_alerts = alerts


def sync_watchlists(wb, today):
    """
    Writes today's contracts into the Quotes sheet, rewriting a row only
    when its contract changed, and keeps each Watchlist_* name on its
    block. Returns the number of rows rewritten.
    """

    try:
        sheet = wb.sheets[SHEET_NAME]
    except Exception:
        sheet = wb.sheets.add(SHEET_NAME)

    changed = 0
    for i, (commodity, name) in enumerate(WATCHLIST_NAMES.items()):
        top = 1 + i * (CONTRACTS_PER_COMMODITY + 2)
        if sheet.range((top, 1)).value != commodity:
            sheet.range((top, 1)).value = commodity
            sheet.range((top, 2)).value = "price"

        for row, contract in enumerate(active_contracts(commodity, CONTRACTS_PER_COMMODITY, today), top + 1):
            symbol = to_cqg_symbol(contract)
            if sheet.range((row, 1)).value != symbol:
                sheet.range((row, 1)).value = symbol
                sheet.range((row, 2)).formula = f'=CQGXLContractData("{symbol}", "Close")'
                changed += 1

        block = sheet.range((top + 1, 1), (top + CONTRACTS_PER_COMMODITY, 2))
        wb.names.add(name, f"='{SHEET_NAME}'!{block.address}")

    if sheet.range(PUSHED_BY_LABEL_CELL).value != "Pushed by":
        sheet.range(PUSHED_BY_LABEL_CELL).value = "Pushed by"
        changed += 1

    if changed:
        _save(wb)
    return changed


def read_live_quotes(wb):
    """
    The live CQGXL watchlist -- [{"commodity", "contract", "price"}, ...].
    Skips a cell showing text (e.g. "Disconnected from server.") or an
    error instead of a number.
    """

    quotes = []
    for commodity, name in WATCHLIST_NAMES.items():
        try:
            rng = wb.names[name].refers_to_range
        except Exception:
            continue

        for symbol, price, *_ in rng.options(ndim=2).value:
            if not symbol or price is None:
                continue
            try:
                price = float(price)
            except (TypeError, ValueError):
                continue
            quotes.append({
                "commodity": commodity,
                "contract": _convert_cqg_symbol(str(symbol)),
                "price": price,
            })

    return quotes


def read_pushed_by(wb):
    """The name in the "Pushed by" cell, or None when it's blank."""

    try:
        value = wb.sheets[SHEET_NAME].range(PUSHED_BY_CELL).value
    except Exception:
        return None
    name = str(value).strip() if value is not None else ""
    return name[:40] or None


_source_column = True  # False once Supabase says the Quotes table has no source column


def push_once(client, quotes, pushed_by=None):

    global _source_column

    if not quotes:
        return 0

    now = datetime.datetime.now(datetime.timezone.utc).isoformat()

    rows = [
        {"contract": q["contract"], "price": q["price"], "updated_at": now}
        for q in quotes
    ]
    if _source_column:
        for row in rows:
            row["source"] = pushed_by

    # One network round-trip for every contract instead of one per
    # contract -- at a short poll interval, 18 sequential calls (~100ms
    # each) was eating a meaningful chunk of the cycle budget on its own.
    try:
        client.table(TABLE_NAME).upsert(rows, on_conflict="contract").execute()
    except Exception as e:
        if not _source_column or "source" not in str(e):
            raise
        # Prices keep flowing without the name until the column exists.
        _source_column = False
        print(
            f'The {TABLE_NAME} table has no source column, so the "Pushed by" name is not sent. Run in the '
            f'Supabase SQL Editor: alter table "{TABLE_NAME}" add column if not exists source text;'
            " -- then restart this window."
        )
        for row in rows:
            row.pop("source")
        client.table(TABLE_NAME).upsert(rows, on_conflict="contract").execute()

    return len(quotes)


def main():

    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--dry-run", action="store_true", help="roll and read the watchlist, but don't touch Supabase")
    parser.add_argument("--once", action="store_true", help="run one cycle and exit")
    parser.add_argument("--today", type=datetime.date.fromisoformat, help="pretend it's this date (YYYY-MM-DD), to test a roll")
    args = parser.parse_args()

    client = None
    if not args.dry_run:
        from supabase import create_client

        secrets = _load_secrets()
        client = create_client(secrets["supabase_url"], secrets["supabase_service_key"])

    print(f"Pushing live CQGXL quotes to Supabase every {PUSH_INTERVAL_SECONDS}s. Ctrl+C to stop.")

    synced_for = None
    while True:
        try:
            today = args.today or datetime.date.today()
            wb = open_workbook()
            if synced_for != today:
                changed = sync_watchlists(wb, today)
                print(f"Watchlist for {today}: {changed} row(s) rolled")
                synced_for = today

            quotes = read_live_quotes(wb)
            pushed_by = read_pushed_by(wb)
            if not quotes:
                print(
                    f"No prices on {WORKBOOK_PATH} yet (cells show #N/A?) -- with more than one Excel "
                    "window open, open it in the one where CQGXL is connected."
                )
            if client is None:
                for q in quotes:
                    print(f"  {q['commodity']:<10} {q['contract']}  {q['price']}")
                print(f"Read {len(quotes)} quote(s), pushed by {pushed_by or '(blank)'} (dry run, nothing pushed)")
            else:
                n = push_once(client, quotes, pushed_by)
                print(f"Pushed {n} quote(s) as {pushed_by or '(no name)'}" if _source_column else f"Pushed {n} quote(s)")
        except Exception as e:
            # Keep the loop alive through transient COM/network errors --
            # one bad cycle shouldn't stop future ones from refreshing.
            print(f"Cycle failed, will retry next interval: {type(e).__name__}: {e}")

        if args.once:
            break
        time.sleep(PUSH_INTERVAL_SECONDS)


if __name__ == "__main__":
    main()
