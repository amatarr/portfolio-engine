"""
Live refresh of the open Excel workbook via xlwings/COM. If
Position_Analysis.xlsm is already open, this attaches to that running
instance (no file lock issue, changes appear instantly); otherwise it
opens it fresh.

Every leg lives in one stored book on the hidden _Data sheet. Legs get
there by Import Daily Prop or Add Leg, and leave by Remove selected legs
or Clear Portfolio; a refresh only prices what's in the book. Results go
to _Data as plain values and inputs are read from named cells -- see
vol_calculator/xlwings_sync.py. Nothing visible on Home is written or
restyled by a refresh.
"""

import os
from datetime import date, datetime, timedelta

import xlwings as xw

from vol_calculator import (
    Market, Portfolio, import_bushel_positions, daily_prop_as_of_date,
    fetch_latest_sofr, commodity_from_contract, approx_option_expiry_days,
    build_manual_row, compute_model_price, compute_underlying_price,
    option_expiry_candidates,
)
from vol_calculator.instruments import AmericanOption, Future
from vol_calculator.excel_sync import POSITIONS_HEADERS, ADD_LEG_TYPES, _position_row, _row_to_instrument
from vol_calculator.xlwings_sync import (
    xw_write_data_block, xw_write_data_value, xw_read_data_value,
    xw_read_book, xw_write_book,
    xw_read_add_leg_inputs, xw_clear_add_leg_inputs,
    xw_read_future_spot, xw_read_interest_rate,
    xw_read_live_quotes,
    xw_write_greeks_sheet,
    xw_write_spot_vol_ladder_sheet, xw_write_heat_maps_sheet,
    xw_write_stress_sheet, xw_write_strike_sheet,
    xw_write_payoff_sheet, xw_write_breakevens_sheet,
)

# Absolute, anchored to this script's own folder: RunPython launches
# Python with an unrelated working directory, and a terminal run must not
# attach to a same-named copy of the workbook open from somewhere else.
_PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
WORKBOOK_PATH = os.path.join(_PROJECT_DIR, "Position_Analysis.xlsm")
DEFAULT_INTEREST_RATE = 0.039
ANALYTIC_SHEETS = ["Greeks", "Spot&Vol Ladder", "Heat Maps", "Stress", "Strike", "Payoff", "Breakevens"]


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


def _as_date(value):
    return value.date() if isinstance(value, datetime) else value


def _unique(items):
    return list(dict.fromkeys(items))


def _with_expiry_date(row, as_of):
    """Records an option's actual expiry date alongside its days-to-expiry, so the countdown stays right on every later day the leg is still in the book."""

    days = row.get("Expiry (d)")
    expiry_date = as_of + timedelta(days=int(days)) if row["Type"] in ("Call", "Put") and days else None
    return {**row, "Settlement": row.get("Settlement"), "Expiry Date": expiry_date}


def _imported_row(instrument, meta, as_of):
    """
    A book row for one leg of a Daily Prop file. A future's settlement is
    kept as the price fallback for its contract (an option's settlement is
    its own premium -- its Entry Price -- never a futures price).
    """

    row = dict(zip(POSITIONS_HEADERS, _position_row(instrument, meta, source="Daily Prop")))
    row["Underlying Price"] = None
    row["Mark to Market"] = None
    if isinstance(instrument, Future):
        row["Settlement"] = meta.get("settlement_price")
    return _with_expiry_date(row, as_of)


def _refresh():
    """Prices every leg in the book and rewrites the Positions table, dropdown lists and analytic sheets."""

    wb = _get_workbook()
    today = date.today()

    book = xw_read_book(wb)
    for row in book:
        if row.get("Expiry Date"):
            row["Expiry (d)"] = max((_as_date(row["Expiry Date"]) - today).days, 1)

    # A live CQGXL quote always wins over a Daily Prop settlement -- the
    # file can be days old, and every price derived from it (Underlying
    # Price, Mark to Market, Greeks) would be stale with it. Settlements
    # only cover contracts the watchlist doesn't quote.
    prices = {row["Contract"]: row["Settlement"] for row in book if row.get("Settlement") is not None}
    live_prices = {q["contract"]: q["price"] for q in xw_read_live_quotes(wb)}
    prices.update(live_prices)

    # A leg on a contract with no price anywhere stays in the book, but
    # can't be valued -- it's shown without prices and left out of the
    # portfolio rather than taking down the whole refresh.
    legs = [(row, *_row_to_instrument(row)) for row in book]
    priced = [(row, instrument, meta) for row, instrument, meta in legs if meta["contract"] in prices]
    unpriced = _unique(row["Contract"] for row, _, meta in legs if meta["contract"] not in prices)

    portfolio = Portfolio()
    for _, instrument, _ in priced:
        portfolio.add(instrument)

    portfolio_contracts = _unique(meta["contract"] for _, _, meta in priced)
    add_leg_contracts = _unique([row["Contract"] for row in book] + list(live_prices))

    sofr = None
    sofr_error = None
    try:
        sofr = fetch_latest_sofr()
    except Exception as e:
        sofr_error = f"{type(e).__name__}: {e}"

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

    # One displayed row per book row, in book order (Remove selected legs
    # relies on that). Underlying Price and Mark to Market (options only,
    # same as Streamlit) are the engine's current price; a blank Entry
    # Price shows the *current* Mark to Market, never a frozen one.
    position_rows = [POSITIONS_HEADERS]
    for row, instrument, meta in legs:
        shown = {h: row.get(h) for h in POSITIONS_HEADERS}
        if meta["contract"] in prices:
            shown["Underlying Price"] = compute_underlying_price(meta["contract"], market)
            if isinstance(instrument, AmericanOption):
                shown["Mark to Market"] = compute_model_price(instrument, meta, market)
                if shown["Entry Price"] is None:
                    shown["Entry Price"] = shown["Mark to Market"]
        position_rows.append([shown[h] for h in POSITIONS_HEADERS])
    xw_write_data_block(wb, "Positions", position_rows)

    if unpriced:
        xw_write_data_value(
            wb, "AddLegStatus",
            f"Not priced (no live quote or Daily Prop settlement): {', '.join(unpriced)}",
        )

    future_spot = xw_read_future_spot(wb, portfolio_contracts)
    if portfolio_contracts:
        xw_write_greeks_sheet(wb.sheets["Greeks"], portfolio, market, portfolio_contracts)
        xw_write_spot_vol_ladder_sheet(wb.sheets["Spot&Vol Ladder"], portfolio, market, future_spot)
        xw_write_heat_maps_sheet(wb.sheets["Heat Maps"], portfolio, market, future_spot)
        xw_write_stress_sheet(wb.sheets["Stress"], portfolio, market, future_spot)
        xw_write_strike_sheet(wb.sheets["Strike"], portfolio, market)
        xw_write_payoff_sheet(wb.sheets["Payoff"], portfolio, market, future_spot)
        xw_write_breakevens_sheet(wb.sheets["Breakevens"], portfolio, market, future_spot)
    else:
        # Nothing priced to analyse (empty book) -- clear the previous
        # book's results rather than leave them looking current.
        for name in ANALYTIC_SHEETS:
            wb.sheets[name].used_range.clear_contents()

    wb.save()

    print(
        f"Refreshed (live) -- {len(book)} legs ({len(priced)} priced), {len(portfolio_contracts)} contracts, "
        f"Future Spot={future_spot}, Interest Rate={interest_rate * 100:.2f}%"
    )


def main():
    _refresh()


def import_daily_prop(path):
    """
    Adds every leg of a Daily Prop file (picked by the ImportDailyProp
    macro) to the book, on top of whatever is already there. Days to
    expiry are measured from the file's own date, which is in its name.
    """

    wb = _get_workbook()
    as_of = daily_prop_as_of_date(path) or date.today()
    imported = import_bushel_positions(path, as_of=as_of)

    xw_write_book(wb, xw_read_book(wb) + [_imported_row(i, m, as_of) for i, m in imported])
    _refresh()

    xw_write_data_value(wb, "AddLegStatus", f"Imported {len(imported)} leg(s) from {os.path.basename(path)}")
    wb.save()


def add_leg():
    """Reads the Add Leg input block, validates it, adds it to the book (one row for a plain leg, 2+ for a multi-leg structure's decomposed legs), clears the inputs, and refreshes."""

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

    new_rows = [_with_expiry_date(row, date.today()) for row in new_rows]
    book = xw_read_book(wb)

    # The book is saved before refreshing, so if a refresh crashes partway
    # (a transient Excel/COM hiccup) and the click is retried with the
    # same inputs, the legs are already there -- don't add a second copy.
    # Checking the first leg is enough: a retry rebuilds identical legs.
    dedup_keys = ("Contract", "Type", "Quantity", "Strike", "Expiry (d)", "Vol (%)")
    already_added = any(all(row.get(k) == new_rows[0].get(k) for k in dedup_keys) for row in book)
    if not already_added:
        xw_write_book(wb, book + new_rows)

    # Clear the inputs and write the success message only *after* the
    # refresh completes -- if it raises partway through, the typed inputs
    # are still there to retry instead of being silently wiped.
    _refresh()
    xw_clear_add_leg_inputs(wb)
    verb = "Already added" if already_added else "Added"
    if len(new_rows) == 1:
        summary = f"{new_rows[0]['Quantity']} {new_rows[0]['Contract']} ({new_rows[0]['Type']})"
    else:
        summary = f"{inputs.get('type')} ({len(new_rows)} legs) on {inputs.get('contract')}"
    xw_write_data_value(wb, "AddLegStatus", f"{verb} {summary}")
    wb.save()


def remove_legs(rows):
    """Removes the legs shown at the given Positions-table rows (a comma-separated string; 1 = the first leg under the header), as picked by the RemoveSelectedLegs macro."""

    wb = _get_workbook()
    picked = {int(r) for r in rows.split(",") if r.strip()}

    # The Positions table shows exactly one row per book row, in order.
    book = xw_read_book(wb)
    kept = [row for i, row in enumerate(book, start=1) if i not in picked]
    xw_write_book(wb, kept)
    _refresh()

    xw_write_data_value(wb, "AddLegStatus", f"Removed {len(book) - len(kept)} leg(s).")
    wb.save()


def clear_portfolio():
    """Empties the whole book -- imported and manually added legs alike -- and refreshes."""

    wb = _get_workbook()
    xw_write_book(wb, [])
    _refresh()


if __name__ == "__main__":
    main()
