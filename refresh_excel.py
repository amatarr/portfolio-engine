"""
Refreshes Position_Analysis.xlsx from the latest Daily Prop file --
re-run this any time the Daily Prop export or the Future Spot selector
(Home!B10) changes. No xlwings, no live recompute: this is a plain
script that reads the workbook, recalculates everything in Python, and
writes the results back, same engine as the Streamlit app.
"""

import openpyxl

from vol_calculator import (
    Market, Portfolio, latest_daily_prop_file, import_bushel_positions,
    write_positions_table, write_greeks_sheet, write_spot_vol_ladder_sheet,
    write_future_spot_selector, read_future_spot,
    contracts_from_imported, prices_from_imported,
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
    market = Market(futures_prices=prices, interest_rate=DEFAULT_INTEREST_RATE)

    wb = openpyxl.load_workbook(WORKBOOK_PATH)

    write_future_spot_selector(wb["Home"], contracts)
    write_positions_table(wb["Home"], imported, start_row=13)

    write_greeks_sheet(wb["Greeks"], portfolio, market, contracts)

    future_spot = read_future_spot(wb["Home"], contracts)
    write_spot_vol_ladder_sheet(wb["Spot&Vol Ladder"], portfolio, market, future_spot)

    wb.save(WORKBOOK_PATH)

    print(f"Refreshed from {path} -- {len(imported)} legs, {len(contracts)} contracts, Future Spot={future_spot}")


if __name__ == "__main__":
    main()
