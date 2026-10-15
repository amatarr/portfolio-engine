"""
Live refresh of the open Excel workbook via xlwings/COM. If
Position_Analysis.xlsm is already open, this attaches to that running
instance (no file lock issue, changes appear instantly); otherwise it
opens it fresh.

Home-sheet results go to the hidden _Data sheet as plain values, and
inputs are read from named cells -- see vol_calculator/xlwings_sync.py.
Nothing visible on Home is written or restyled by a refresh.
"""

import os

import xlwings as xw

from vol_calculator import (
    Market, Portfolio, latest_daily_prop_file, import_bushel_positions,
    fetch_latest_sofr, contracts_from_imported, prices_from_imported,
    commodity_from_contract, approx_option_expiry_days,
    build_manual_row, compute_model_price, compute_underlying_price,
    option_expiry_candidates,
)
from vol_calculator.instruments import AmericanOption
from vol_calculator.excel_sync import POSITIONS_HEADERS, ADD_LEG_TYPES, _position_row, _row_to_instrument
from vol_calculator.xlwings_sync import (
    xw_write_data_block, xw_read_data_block, xw_write_data_value, xw_read_data_value,
    xw_read_manual_legs, xw_write_manual_legs,
    xw_read_add_leg_inputs, xw_clear_add_leg_inputs,
    xw_read_future_spot, xw_read_interest_rate,
    xw_read_live_quotes,
    xw_write_greeks_sheet,
    xw_write_spot_vol_ladder_sheet, xw_write_heat_maps_sheet,
    xw_write_stress_sheet, xw_write_strike_sheet,
    xw_write_payoff_sheet, xw_write_breakevens_sheet,
)

# Anchored to this script's own folder, not the current working
# directory -- RunPython launches Python with a cwd that doesn't match
# this project folder, so a plain relative "Daily prop" silently finds
# nothing when triggered from an Excel button (works fine when this
# script is run directly from a terminal already cd'd here, which is
# why this only shows up live). The workbook path is absolute for the
# same reason, and so a terminal run can't attach to a same-named copy
# of the workbook open from somewhere else.
_PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
WORKBOOK_PATH = os.path.join(_PROJECT_DIR, "Position_Analysis.xlsm")
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


def _write_sofr(wb, sofr, error):
    """The visible Interest Rate cell is a formula on SofrRate, so typing a manual rate over it simply replaces the formula. A failed fetch keeps the last known rate instead of blanking it."""

    if sofr:
        xw_write_data_value(wb, "SofrRate", sofr["rate"])
        xw_write_data_value(wb, "SofrCaption", f"SOFR -- last updated {sofr['effective_date']} (click for source)")
        xw_write_data_value(wb, "SofrUrl", sofr["source_url"])
        return

    if xw_read_data_value(wb, "SofrRate") is None:
        xw_write_data_value(wb, "SofrRate", DEFAULT_INTEREST_RATE)
    xw_write_data_value(wb, "SofrCaption", f"SOFR fetch failed -- showing last known rate. Error: {error}")
    xw_write_data_value(wb, "SofrUrl", None)


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
    # push_quotes_supabase.py reads to feed Streamlit. A live quote wins
    # over the Daily Prop settlement whenever there is one: the CSV can
    # be days old, and every price derived from it (Underlying Price,
    # Mark to Market, Greeks) would be stale with it. Daily Prop only
    # covers contracts the watchlist doesn't quote.
    live_quotes = xw_read_live_quotes(home)
    live_prices = {q["contract"]: q["price"] for q in live_quotes}
    prices.update(live_prices)

    # Defensive: a Daily Prop option whose underlying future doesn't
    # also appear in today's file (e.g. a thinly-listed month) and
    # isn't live-quoted either has no price anywhere -- same class of
    # bug that crashed Streamlit's Greeks on an unpriced leg. Drop it
    # here rather than let it take down this whole refresh.
    imported = [(instrument, meta) for instrument, meta in imported if meta["contract"] in prices]
    contracts = contracts_from_imported(imported)

    manual_rows = xw_read_manual_legs(wb) if keep_manual else []
    manual_rows += list(extra_new_rows or [])
    manual_instruments = [_row_to_instrument(row) for row in manual_rows]

    # Same defensive drop for Manual legs -- a leg added for a contract
    # that then turns out unpriceable (CSV doesn't cover it today, no
    # live quote either) can't be valued, so it can't stay in the book.
    unpriced_manual = {meta["contract"] for _, meta in manual_instruments if meta["contract"] not in prices}
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

    # Dropdown sources -- the visible dropdowns point at these names, so
    # only the list contents change here, never the dropdowns themselves.
    xw_write_data_block(wb, "ContractList", [[c] for c in add_leg_contracts])
    xw_write_data_block(wb, "PortfolioContracts", [[c] for c in portfolio_contracts])
    xw_write_data_block(wb, "LegTypes", [[t] for t in ADD_LEG_TYPES])
    xw_write_data_block(wb, "OptionTypes", [["call"], ["put"]])

    _write_sofr(wb, sofr, sofr_error)
    interest_rate = xw_read_interest_rate(wb, fallback_rate=DEFAULT_INTEREST_RATE)

    market = Market(futures_prices=prices, interest_rate=interest_rate)

    # What Add Leg's suggestions are built from (in the workbook's own
    # VBA, so they fill instantly without a Python round-trip): each
    # offered contract's current price, and its real CME option expiries.
    xw_write_data_block(wb, "ContractQuotes", [
        [c, commodity_from_contract(c), market.futures_prices.get(c)] for c in add_leg_contracts
    ])
    xw_write_data_block(wb, "ExpiryCandidates", [
        [c, candidate["label"], candidate["days"]]
        for c in add_leg_contracts for candidate in option_expiry_candidates(c)
    ])

    # Underlying Price (every leg, Futures included) and Mark to Market
    # (options only, same as Streamlit) are the engine's current price,
    # recomputed every refresh. Daily Prop legs' meta dicts get annotated
    # in place for _position_row; Manual legs get a display copy, while
    # what's stored in ManualLegs stays exactly what was entered -- so a
    # blank Entry Price keeps falling back to the *current* Mark to
    # Market rather than freezing at the one from the day it was added.
    for instrument, meta in imported:
        meta["underlying_price"] = compute_underlying_price(meta["contract"], market)
        if isinstance(instrument, AmericanOption):
            meta["model_price"] = compute_model_price(instrument, meta, market)

    manual_display = []
    for row, (instrument, meta) in zip(manual_rows, manual_instruments):
        shown = dict(row)
        shown["Underlying Price"] = compute_underlying_price(meta["contract"], market)
        if isinstance(instrument, AmericanOption):
            shown["Mark to Market"] = compute_model_price(instrument, meta, market)
            if shown.get("Entry Price") is None:
                shown["Entry Price"] = shown["Mark to Market"]
        manual_display.append(shown)

    position_rows = [POSITIONS_HEADERS]
    position_rows += [_position_row(instrument, meta, source="Daily Prop") for instrument, meta in imported]
    position_rows += [[row.get(h) for h in POSITIONS_HEADERS] for row in manual_display]
    xw_write_data_block(wb, "Positions", position_rows)
    xw_write_manual_legs(
        wb, [{**row, "Underlying Price": None, "Mark to Market": None} for row in manual_rows],
    )

    future_spot = xw_read_future_spot(wb, portfolio_contracts)
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

    inputs = xw_read_add_leg_inputs(wb)

    try:
        new_rows = build_manual_row(
            inputs, commodity_lookup=commodity_from_contract, auto_expiry_fn=approx_option_expiry_days,
        )
    except ValueError as e:
        message = f"Add Leg failed: {e}"
        xw_write_data_value(wb, "AddLegStatus", message)
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
        for row in xw_read_manual_legs(wb)
    )

    # Clear the inputs and write the success message only *after*
    # _refresh() actually completes -- if it raises partway through,
    # the typed inputs are still there to retry instead of being
    # silently wiped with nothing saved.
    _refresh(extra_new_rows=None if already_added else new_rows)
    xw_clear_add_leg_inputs(wb)
    verb = "Already added" if already_added else "Added"
    leg_type = inputs.get("type")
    if len(new_rows) == 1:
        summary = f"{new_rows[0]['Quantity']} {new_rows[0]['Contract']} ({new_rows[0]['Type']})"
    else:
        summary = f"{leg_type} ({len(new_rows)} legs) on {inputs.get('contract')}"
    xw_write_data_value(wb, "AddLegStatus", f"{verb} {summary}")
    wb.save()


def remove_legs(rows):
    """
    Removes the Manual legs shown at the given Positions-table rows (a
    comma-separated string; 1 = the first leg under the header), as
    picked by the RemoveSelectedLegs macro. Daily Prop legs are left in
    place: they come back from the Daily Prop file on every refresh.
    """

    wb = _get_workbook()
    picked = sorted({int(r) for r in rows.split(",") if r.strip()})

    # Positions lists every Daily Prop leg first, then the Manual legs in
    # the same order they're stored in ManualLegs.
    shown = xw_read_data_block(wb, "Positions")[1:]
    n_daily = sum(1 for row in shown if row[-1] != "Manual")
    in_table = [i for i in picked if 1 <= i <= len(shown)]
    drop = {i - 1 - n_daily for i in in_table if shown[i - 1][-1] == "Manual"}
    skipped = len(in_table) - len(drop)

    manual = xw_read_manual_legs(wb)
    xw_write_manual_legs(wb, [leg for k, leg in enumerate(manual) if k not in drop])
    _refresh()

    message = f"Removed {len(drop)} leg(s)."
    if skipped:
        message += f" {skipped} Daily Prop leg(s) kept -- they come from the Daily Prop file."
    xw_write_data_value(wb, "AddLegStatus", message)
    wb.save()


def clear_portfolio():
    """Drops every Manual leg (Daily Prop legs come back on the next refresh regardless) and runs a full refresh."""

    _refresh(keep_manual=False)


if __name__ == "__main__":
    main()
