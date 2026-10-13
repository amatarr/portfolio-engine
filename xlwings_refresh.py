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
    commodity_from_contract, approx_option_expiry_days,
    build_manual_row, ADD_LEG_STATUS_CELL,
)
from vol_calculator.excel_sync import _row_to_instrument
from vol_calculator.xlwings_sync import (
    xw_reset_dropdown_validations,
    xw_write_future_spot_selector, xw_read_future_spot,
    xw_write_interest_rate_cell, xw_read_interest_rate,
    xw_write_positions_table, xw_read_positions_rows,
    xw_write_add_leg_block, xw_read_add_leg_inputs, xw_clear_add_leg_inputs,
    xw_write_greeks_sheet,
    xw_write_spot_vol_ladder_sheet, xw_write_heat_maps_sheet,
    xw_write_stress_sheet, xw_write_strike_sheet,
    xw_write_payoff_sheet, xw_write_breakevens_sheet,
)

WORKBOOK_PATH = "Position_Analysis.xlsm"
DAILY_PROP_FOLDER = "Daily prop"
DEFAULT_INTEREST_RATE = 0.039


def _get_workbook():
    """
    xw.Book.caller() works when this is invoked via RunPython from a
    macro inside the workbook itself (the Refresh button) -- it gets a
    reference to that exact calling book. Falls back to opening/
    attaching by path for standalone terminal runs (testing, or running
    this script directly without the button).
    """

    try:
        return xw.Book.caller()
    except Exception:
        return xw.Book(WORKBOOK_PATH)


def _manual_rows_from_sheet(home, extra_new_row=None):
    """
    Reads whatever Manual rows are currently on the sheet and optionally
    appends a brand-new one that doesn't exist there yet (used by
    add_leg() to inject its new leg before the full rebuild below sees
    it). clear_portfolio() passes nothing and gets [].
    """

    rows = [r for r in xw_read_positions_rows(home, start_row=13) if r.get("Source") == "Manual"]
    if extra_new_row is not None:
        rows.append(extra_new_row)
    return rows


def _refresh(extra_new_row=None, keep_manual=True):
    """
    Shared body behind main()/add_leg()/clear_portfolio(). extra_new_row
    (a single Manual row dict, Source="Manual") is folded in alongside
    whatever Manual rows already exist, before the Positions table and
    Portfolio are rebuilt -- so a just-added leg is priced in the same
    pass that writes it. keep_manual=False (clear_portfolio) drops every
    existing Manual row instead of preserving them.
    """

    path = latest_daily_prop_file(DAILY_PROP_FOLDER)
    if path is None:
        raise SystemExit(f"No Daily_Prop_Bod_Detail_*.csv found in {DAILY_PROP_FOLDER!r}")

    imported = import_bushel_positions(path)
    contracts = contracts_from_imported(imported)
    prices = prices_from_imported(imported)

    wb = _get_workbook()
    home = wb.sheets["Home"]

    manual_rows = _manual_rows_from_sheet(home, extra_new_row) if keep_manual else (
        [extra_new_row] if extra_new_row is not None else []
    )
    manual_instruments = [_row_to_instrument(row) for row in manual_rows]

    portfolio = Portfolio()
    for instrument, _ in imported:
        portfolio.add(instrument)
    for instrument, _ in manual_instruments:
        portfolio.add(instrument)

    sofr = None
    sofr_error = None
    try:
        sofr = fetch_latest_sofr()
    except Exception as e:
        sofr_error = f"{type(e).__name__}: {e}"

    xw_reset_dropdown_validations(home)
    xw_write_future_spot_selector(home, contracts)
    xw_write_add_leg_block(home, contracts)
    xw_write_interest_rate_cell(home, sofr=sofr, fallback_rate=DEFAULT_INTEREST_RATE, error=sofr_error)
    interest_rate = xw_read_interest_rate(home, fallback_rate=DEFAULT_INTEREST_RATE)

    market = Market(futures_prices=prices, interest_rate=interest_rate)
    for instrument, meta in manual_instruments:
        if meta.get("settlement_price") is not None:
            market.futures_prices[meta["contract"]] = meta["settlement_price"]

    xw_write_positions_table(home, imported, extra_rows=manual_rows, start_row=13)

    future_spot = xw_read_future_spot(home, contracts)
    xw_write_greeks_sheet(wb.sheets["Greeks"], portfolio, market, contracts)
    xw_write_spot_vol_ladder_sheet(wb.sheets["Spot&Vol Ladder"], portfolio, market, future_spot)
    xw_write_heat_maps_sheet(wb.sheets["Heat Maps"], portfolio, market, future_spot)
    xw_write_stress_sheet(wb.sheets["Stress"], portfolio, market, future_spot)
    xw_write_strike_sheet(wb.sheets["Strike"], portfolio, market)
    xw_write_payoff_sheet(wb.sheets["Payoff"], portfolio, market, future_spot)
    xw_write_breakevens_sheet(wb.sheets["Breakevens"], portfolio, market, future_spot)

    wb.save()

    print(
        f"Refreshed (live) from {path} -- {len(imported)} Daily Prop legs + {len(manual_rows)} Manual legs, "
        f"{len(contracts)} contracts, Future Spot={future_spot}, Interest Rate={interest_rate * 100:.2f}%"
    )


def main():
    _refresh()


def add_leg():
    """Reads the Add Leg input block, validates it, appends it as a new Manual row, clears the inputs, and runs a full refresh with it included."""

    wb = _get_workbook()
    home = wb.sheets["Home"]

    inputs = xw_read_add_leg_inputs(home)

    try:
        new_row = build_manual_row(
            inputs, commodity_lookup=commodity_from_contract, auto_expiry_fn=approx_option_expiry_days,
        )
    except ValueError as e:
        message = f"Add Leg failed: {e}"
        home[ADD_LEG_STATUS_CELL].value = message
        wb.save()
        # A validation failure is easy to miss if it only lands in a
        # cell -- pop up a real dialog too, same as a VBA MsgBox would.
        import ctypes
        ctypes.windll.user32.MessageBoxW(0, message, "Add Leg", 0x30)
        return

    xw_clear_add_leg_inputs(home)
    home[ADD_LEG_STATUS_CELL].value = f"Added {new_row['Quantity']} {new_row['Contract']} ({new_row['Type']})"
    _refresh(extra_new_row=new_row)


def clear_portfolio():
    """Drops every Manual leg (Daily Prop legs come back on the next refresh regardless) and runs a full refresh."""

    _refresh(keep_manual=False)


if __name__ == "__main__":
    main()
