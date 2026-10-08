"""
Live version of refresh_excel.py -- writes straight into the open
Excel workbook via xlwings/COM instead of saving a closed file with
openpyxl. If Position_Analysis.xlsx is already open, this attaches to
that running instance (no file lock issue, changes appear instantly);
otherwise it opens it fresh.

Only writes values, never structure -- the sheet layout, charts, and
conditional formatting are the ones already built via openpyxl/
refresh_excel.py and stay exactly where they are.
"""

import xlwings as xw

from vol_calculator import (
    Market, Portfolio, latest_daily_prop_file, import_bushel_positions,
    fetch_latest_sofr, contracts_from_imported, prices_from_imported,
)
from vol_calculator.xlwings_sync import (
    xw_write_future_spot_selector, xw_read_future_spot,
    xw_write_interest_rate_cell, xw_read_interest_rate,
    xw_write_positions_table, xw_write_greeks_sheet,
)

WORKBOOK_PATH = "Position_Analysis.xlsx"
DAILY_PROP_FOLDER = "Daily prop"
DEFAULT_INTEREST_RATE = 0.039


def main():

    path = latest_daily_prop_file(DAILY_PROP_FOLDER)
    if path is None:
        raise SystemExit(f"No Daily_Prop_Bod_Detail_*.csv found in {DAILY_PROP_FOLDER!r}")

    imported = import_bushel_positions(path)
    contracts = contracts_from_imported(imported)

    portfolio = Portfolio()
    for instrument, _ in imported:
        portfolio.add(instrument)

    prices = prices_from_imported(imported)

    try:
        sofr = fetch_latest_sofr()
    except Exception:
        sofr = None

    wb = xw.Book(WORKBOOK_PATH)
    home = wb.sheets["Home"]

    xw_write_future_spot_selector(home, contracts)
    xw_write_interest_rate_cell(home, sofr=sofr, fallback_rate=DEFAULT_INTEREST_RATE)
    interest_rate = xw_read_interest_rate(home, fallback_rate=DEFAULT_INTEREST_RATE)

    market = Market(futures_prices=prices, interest_rate=interest_rate)

    xw_write_positions_table(home, imported, start_row=13)

    future_spot = xw_read_future_spot(home, contracts)
    xw_write_greeks_sheet(wb.sheets["Greeks"], portfolio, market, contracts)

    wb.save()

    print(
        f"Refreshed (live) from {path} -- {len(imported)} legs, {len(contracts)} contracts, "
        f"Future Spot={future_spot}, Interest Rate={interest_rate * 100:.2f}%"
    )


if __name__ == "__main__":
    main()
