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

import os

import xlwings as xw

from vol_calculator import (
    Market, Portfolio, latest_daily_prop_file, import_bushel_positions,
    fetch_latest_sofr, contracts_from_imported, prices_from_imported,
    commodity_from_contract, approx_option_expiry_days,
    build_manual_row, ADD_LEG_STATUS_CELL, compute_model_price,
    compute_underlying_price,
)
from vol_calculator.instruments import AmericanOption
from vol_calculator.excel_sync import _row_to_instrument
from vol_calculator.instruments import Future
from vol_calculator.xlwings_sync import (
    xw_reset_dropdown_validations,
    xw_write_future_spot_selector, xw_read_future_spot,
    xw_write_interest_rate_cell, xw_read_interest_rate,
    xw_write_positions_table, xw_read_positions_rows,
    xw_write_add_leg_block, xw_read_add_leg_inputs, xw_clear_add_leg_inputs,
    xw_read_live_quotes, xw_style_live_quotes_watchlist,
    xw_write_greeks_sheet,
    xw_write_spot_vol_ladder_sheet, xw_write_heat_maps_sheet,
    xw_write_stress_sheet, xw_write_strike_sheet,
    xw_write_payoff_sheet, xw_write_breakevens_sheet,
)

WORKBOOK_PATH = "Position_Analysis.xlsm"
# Anchored to this script's own folder, not the current working
# directory -- RunPython launches Python with a cwd that doesn't match
# this project folder, so a plain relative "Daily prop" silently finds
# nothing when triggered from an Excel button (works fine when this
# script is run directly from a terminal already cd'd here, which is
# why this only shows up live).
_PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
DAILY_PROP_FOLDER = os.path.join(_PROJECT_DIR, "Daily prop")
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


def _manual_rows_from_sheet(home, extra_new_rows=None):
    """
    Reads whatever Manual rows are currently on the sheet and optionally
    appends brand-new ones that don't exist there yet (used by add_leg()
    to inject a just-added leg -- or every decomposed leg of a just-
    added multi-leg structure -- before the full rebuild below sees
    them). clear_portfolio() passes nothing and gets [].
    """

    rows = [r for r in xw_read_positions_rows(home, start_row=13) if r.get("Source") == "Manual"]
    if extra_new_rows:
        rows.extend(extra_new_rows)
    return rows


def _refresh(extra_new_rows=None, keep_manual=True):
    """
    Shared body behind main()/add_leg()/clear_portfolio(). extra_new_rows
    (a list of Manual row dicts, Source="Manual" -- one row for a plain
    Future/Call/Put, 2+ for a multi-leg structure's decomposed legs) is
    folded in alongside whatever Manual rows already exist, before the
    Positions table and Portfolio are rebuilt -- so a just-added leg is
    priced in the same pass that writes it. keep_manual=False
    (clear_portfolio) drops every existing Manual row instead of
    preserving them.
    """

    path = latest_daily_prop_file(DAILY_PROP_FOLDER)
    if path is None:
        raise SystemExit(f"No Daily_Prop_Bod_Detail_*.csv found in {DAILY_PROP_FOLDER!r}")

    imported = import_bushel_positions(path)
    prices = prices_from_imported(imported)

    wb = _get_workbook()
    home = wb.sheets["Home"]

    # The Home sheet's live CQGXL watchlist -- same block
    # push_quotes_supabase.py reads to feed Streamlit. Daily Prop's own
    # settlement price always wins for a contract it covers (pinned,
    # same rule as Streamlit); this only fills in a price for a contract
    # Daily Prop doesn't have today, e.g. a brand-new leg added by hand
    # that isn't in the current book yet.
    live_quotes = xw_read_live_quotes(home)
    live_prices = {q["contract"]: q["price"] for q in live_quotes}
    for contract, price in live_prices.items():
        prices.setdefault(contract, price)

    # Defensive: a Daily Prop option whose underlying future doesn't
    # also appear in today's file (e.g. a thinly-listed month) and
    # isn't live-quoted either has no price anywhere -- same class of
    # bug that crashed Streamlit's Greeks on an unpriced leg. Drop it
    # here rather than let it take down this whole refresh.
    imported = [(instrument, meta) for instrument, meta in imported if meta["contract"] in prices]
    contracts = contracts_from_imported(imported)

    manual_rows = _manual_rows_from_sheet(home, extra_new_rows) if keep_manual else (
        list(extra_new_rows) if extra_new_rows else []
    )
    manual_instruments = [_row_to_instrument(row) for row in manual_rows]

    # Same defensive drop for Manual legs -- a leg added for a contract
    # that then turns out unpriceable (CSV doesn't cover it today, no
    # live quote either) can't be valued, so it can't stay in the book.
    unpriced_manual = {
        meta["contract"] for _, meta in manual_instruments
        if meta["contract"] not in prices and meta.get("settlement_price") is None
    }
    if unpriced_manual:
        manual_rows = [r for r in manual_rows if r["Contract"] not in unpriced_manual]
        manual_instruments = [
            (i, m) for i, m in manual_instruments if m["contract"] not in unpriced_manual
        ]

    portfolio = Portfolio()
    for instrument, _ in imported:
        portfolio.add(instrument)
    for instrument, _ in manual_instruments:
        portfolio.add(instrument)

    # Every analytic sheet (Greeks, ladders, heat maps, stress...) breaks
    # its numbers down per contract -- a manually-added leg's contract
    # has to be in this list too, or its own row just never shows up even
    # though it's correctly priced and counted in portfolio-wide totals.
    portfolio_contracts = list(contracts)
    for _, meta in manual_instruments:
        if meta["contract"] not in portfolio_contracts:
            portfolio_contracts.append(meta["contract"])

    sofr = None
    sofr_error = None
    try:
        sofr = fetch_latest_sofr()
    except Exception as e:
        sofr_error = f"{type(e).__name__}: {e}"

    # Add Leg's Contract dropdown should offer every live-quoted
    # contract too, not just today's Daily Prop book -- otherwise there
    # was never a way to add a leg for a contract Daily Prop doesn't
    # happen to carry today.
    add_leg_contracts = list(contracts)
    for contract in live_prices:
        if contract not in add_leg_contracts:
            add_leg_contracts.append(contract)

    xw_style_live_quotes_watchlist(home)

    xw_reset_dropdown_validations(home)
    xw_write_future_spot_selector(home, portfolio_contracts)
    xw_write_add_leg_block(home, add_leg_contracts)
    xw_write_interest_rate_cell(home, sofr=sofr, fallback_rate=DEFAULT_INTEREST_RATE, error=sofr_error)
    interest_rate = xw_read_interest_rate(home, fallback_rate=DEFAULT_INTEREST_RATE)

    market = Market(futures_prices=prices, interest_rate=interest_rate)
    for instrument, meta in manual_instruments:
        # Only a Future's settlement_price is the underlying contract's
        # own price -- an Option's is its own premium (Entry Price), a
        # completely different number that must never overwrite the
        # underlying's market price (that would corrupt every other leg
        # priced off the same contract).
        if isinstance(instrument, Future) and meta.get("settlement_price") is not None:
            market.futures_prices[meta["contract"]] = meta["settlement_price"]

    # A manually-added Future leg has no Underlying Price of its own
    # (it isn't CSV-imported) -- backfill the Positions table's own
    # column with whatever price the engine actually resolved for it
    # (Daily Prop if it's covered, else the live CQGXL watchlist), so
    # the row isn't left blank even though it's correctly priced.
    for row, (instrument, meta) in zip(manual_rows, manual_instruments):
        if isinstance(instrument, Future) and row.get("Underlying Price") is None:
            row["Underlying Price"] = market.futures_prices.get(meta["contract"])

    # "Mark to Market" and "Underlying Price" -- options only (a Future
    # has neither, same as Streamlit); computed fresh every refresh for
    # both Daily Prop and Manual legs. Daily Prop legs' meta dicts get
    # annotated in place so _position_row (called inside
    # xw_write_positions_table) picks them up the same way it already
    # does for settlement_price.
    for instrument, meta in imported:
        if isinstance(instrument, AmericanOption):
            meta["model_price"] = compute_model_price(instrument, meta, market)
            meta["underlying_price"] = compute_underlying_price(meta["contract"], market)
    for row, (instrument, meta) in zip(manual_rows, manual_instruments):
        if isinstance(instrument, AmericanOption):
            row["Mark to Market"] = compute_model_price(instrument, meta, market)
            row["Underlying Price"] = compute_underlying_price(meta["contract"], market)
            # No Entry Price typed in at Add Leg time -- fall back to
            # this leg's own Mark to Market, same fallback the Positions
            # table uses for any option with no recorded Entry Price.
            if row.get("Entry Price") is None:
                row["Entry Price"] = row["Mark to Market"]

    xw_write_positions_table(home, imported, extra_rows=manual_rows, start_row=13)

    future_spot = xw_read_future_spot(home, portfolio_contracts)
    xw_write_greeks_sheet(wb.sheets["Greeks"], portfolio, market, portfolio_contracts)
    xw_write_spot_vol_ladder_sheet(wb.sheets["Spot&Vol Ladder"], portfolio, market, future_spot)
    xw_write_heat_maps_sheet(wb.sheets["Heat Maps"], portfolio, market, future_spot)
    xw_write_stress_sheet(wb.sheets["Stress"], portfolio, market, future_spot)
    xw_write_strike_sheet(wb.sheets["Strike"], portfolio, market)
    xw_write_payoff_sheet(wb.sheets["Payoff"], portfolio, market, future_spot)
    xw_write_breakevens_sheet(wb.sheets["Breakevens"], portfolio, market, future_spot)

    wb.save()

    print(
        f"Refreshed (live) from {path} -- {len(imported)} Daily Prop legs + {len(manual_rows)} Manual legs, "
        f"{len(portfolio_contracts)} contracts, Future Spot={future_spot}, Interest Rate={interest_rate * 100:.2f}%"
    )


def main():
    _refresh()


def add_leg():
    """Reads the Add Leg input block, validates it, appends it as new Manual row(s) (one for a plain leg, 2+ for a multi-leg structure's decomposed legs), clears the inputs, and runs a full refresh with it included."""

    wb = _get_workbook()
    home = wb.sheets["Home"]

    inputs = xw_read_add_leg_inputs(home)

    try:
        new_rows = build_manual_row(
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

    # If a previous attempt already wrote these exact rows (its own
    # _refresh() crashed partway through -- e.g. a transient Excel/COM
    # hiccup -- before reaching the clear-inputs step below, so the
    # inputs never got wiped and this is really a retry of the same
    # click), don't append a second copy. Checking just the first leg
    # is enough -- a retry resubmits identical inputs, so the first leg
    # it builds is identical too.
    dedup_keys = ("Contract", "Type", "Quantity", "Strike", "Expiry (d)", "Vol (%)")
    already_added = any(
        all(row.get(k) == new_rows[0].get(k) for k in dedup_keys)
        for row in xw_read_positions_rows(home, start_row=13)
        if row.get("Source") == "Manual"
    )

    # Clear the inputs and write the success message only *after*
    # _refresh() actually completes -- if it raises partway through,
    # the typed inputs are still there to retry instead of being
    # silently wiped with nothing saved.
    _refresh(extra_new_rows=None if already_added else new_rows)
    xw_clear_add_leg_inputs(home)
    verb = "Already added" if already_added else "Added"
    leg_type = inputs.get("type")
    if len(new_rows) == 1:
        summary = f"{new_rows[0]['Quantity']} {new_rows[0]['Contract']} ({new_rows[0]['Type']})"
    else:
        summary = f"{leg_type} ({len(new_rows)} legs) on {inputs.get('contract')}"
    home[ADD_LEG_STATUS_CELL].value = f"{verb} {summary}"
    wb.save()


def clear_portfolio():
    """Drops every Manual leg (Daily Prop legs come back on the next refresh regardless) and runs a full refresh."""

    _refresh(keep_manual=False)


if __name__ == "__main__":
    main()
