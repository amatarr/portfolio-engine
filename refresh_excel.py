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
    fetch_latest_sofr,
    write_positions_table, write_greeks_sheet, write_spot_vol_ladder_sheet,
    write_heat_maps_sheet, write_stress_sheet, write_strike_sheet,
    write_breakevens_sheet, write_payoff_sheet,
    write_future_spot_selector, read_future_spot,
    write_interest_rate_cell, read_interest_rate,
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

    try:
        sofr = fetch_latest_sofr()
    except Exception:
        sofr = None

    wb = openpyxl.load_workbook(WORKBOOK_PATH)

    write_future_spot_selector(wb["Home"], contracts)
    write_interest_rate_cell(wb["Home"], sofr=sofr, fallback_rate=DEFAULT_INTEREST_RATE)
    interest_rate = read_interest_rate(wb["Home"], fallback_rate=DEFAULT_INTEREST_RATE)

    market = Market(futures_prices=prices, interest_rate=interest_rate)

    write_positions_table(wb["Home"], imported, start_row=13)

    write_greeks_sheet(wb["Greeks"], portfolio, market, contracts)

    future_spot = read_future_spot(wb["Home"], contracts)
    write_spot_vol_ladder_sheet(wb["Spot&Vol Ladder"], portfolio, market, future_spot)
    write_heat_maps_sheet(wb["Heat Maps"], portfolio, market, future_spot)
    write_stress_sheet(wb["Stress"], portfolio, market, future_spot)
    write_strike_sheet(wb["Strike"], portfolio, market)
    write_payoff_sheet(wb["Payoff"], portfolio, market, future_spot)
    write_breakevens_sheet(wb["Breakevens"], portfolio, market, future_spot)

    wb.save(WORKBOOK_PATH)

    print(
        f"Refreshed from {path} -- {len(imported)} legs, {len(contracts)} contracts, "
        f"Future Spot={future_spot}, Interest Rate={interest_rate * 100:.2f}%"
    )


if __name__ == "__main__":
    main()
