from openpyxl.worksheet.datavalidation import DataValidation

from vol_calculator.instruments import Future, AmericanOption
from vol_calculator.greeks import GreekEngine
from vol_calculator.scenarios import (
    spot_ladder,
    vol_ladder,
    spot_vol_surface,
    stress_report,
    strike_map,
    breakevens_quadratic,
    breakevens_numerical,
    payoff_diagram,
)


def clear_range(ws, min_row=1, max_row=500, min_col=1, max_col=20):
    """Blanks out a rectangular range -- used before rewriting a sheet/block so a shrinking table doesn't leave stale cells behind."""

    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.value = None


def write_dataframe(ws, df, start_row=1, start_col=1, label=None, include_index=False, index_label=""):
    """
    Writes a pandas DataFrame as a plain header + rows block, starting at
    (start_row, start_col). If label is given, it goes on its own row
    just above the header. With include_index=True, the DataFrame's
    index becomes the first column (for pivot-shaped outputs like
    spot_vol_surface/strike_map, where the index itself is meaningful --
    e.g. spot move or strike). Returns the row just after the last one
    written, so callers can stack multiple tables on one sheet.
    """

    row = start_row
    col_offset = start_col + (1 if include_index else 0)

    if label is not None:
        ws.cell(row=row, column=start_col, value=label)
        row += 1

    if include_index:
        ws.cell(row=row, column=start_col, value=index_label)
    for col, header in enumerate(df.columns, start=col_offset):
        ws.cell(row=row, column=col, value=header)
    row += 1

    for idx, data_row in df.iterrows():
        if include_index:
            ws.cell(row=row, column=start_col, value=idx)
        for col, value in enumerate(data_row, start=col_offset):
            ws.cell(row=row, column=col, value=round(value, 4) if isinstance(value, float) else value)
        row += 1

    return row

POSITIONS_HEADERS = [
    "Contract", "Commodity", "Type", "Quantity",
    "Strike", "Expiry (d)", "Vol (%)", "Settlement Price",
]

FUTURE_SPOT_LABEL_CELL = "A10"
FUTURE_SPOT_VALUE_CELL = "B10"


def write_future_spot_selector(ws, contracts, label_cell=FUTURE_SPOT_LABEL_CELL, value_cell=FUTURE_SPOT_VALUE_CELL):
    """
    A native Excel dropdown (data validation list) of the contracts
    currently in the book -- the "which contract is shocked" selector
    for sheets that show one contract at a time (Ladders, Surface,
    Stress, Strike, Payoff, Breakevens), same role as the Future Spot
    selector in each Streamlit tab.

    The dropdown itself works live in Excel. What's NOT live yet (no
    xlwings): picking a new value here doesn't recompute the other
    sheets by itself -- the refresh script needs to be re-run after
    changing it, and it reads whatever's currently in value_cell.
    """

    ws[label_cell] = "Future Spot:"

    if not contracts:
        return

    current = ws[value_cell].value
    ws[value_cell] = current if current in contracts else contracts[0]

    dv = DataValidation(type="list", formula1=f'"{",".join(contracts)}"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ws[value_cell])


def read_future_spot(ws, contracts, value_cell=FUTURE_SPOT_VALUE_CELL):
    """Reads the current Future Spot selection, falling back to the first contract if unset/invalid."""

    value = ws[value_cell].value
    return value if value in contracts else (contracts[0] if contracts else None)


def contracts_from_imported(imported):
    """Distinct contracts, in order of first appearance."""

    seen = []
    for _, meta in imported:
        c = meta.get("contract")
        if c and c not in seen:
            seen.append(c)
    return seen


def prices_from_imported(imported):
    """
    contract -> settlement price, from the Daily Prop file itself --
    stand-in for live CQGXL prices until the Home sheet's formula cells
    have cached values to read (requires Excel to have actually
    recalculated with CQGXL connected and been saved at least once).
    """

    prices = {}
    for _, meta in imported:
        price = meta.get("settlement_price")
        if price is not None:
            prices[meta["contract"]] = price
    return prices


def _position_row(instrument, meta):

    if isinstance(instrument, Future):
        return [
            meta["contract"], meta["commodity"], "Future", instrument.quantity,
            None, None, meta.get("vol_pct"), meta.get("settlement_price"),
        ]

    if isinstance(instrument, AmericanOption):
        return [
            meta["contract"], meta["commodity"],
            "Call" if instrument.option_type == "call" else "Put",
            instrument.quantity, instrument.strike, instrument.expiry_days,
            round(instrument.volatility * 100, 2), meta.get("settlement_price"),
        ]

    raise TypeError(f"Unsupported instrument type: {type(instrument)}")


def write_positions_table(ws, imported, start_row=11, clear_rows=500):
    """
    Writes (instrument, meta) pairs -- the same shape import_bushel_positions
    returns -- into an existing worksheet, starting at start_row (a
    "POSITIONS" label, then headers, then one row per leg). Clears a
    generous range below start_row first, so a shrinking book doesn't
    leave stale rows behind from a previous, longer day's data.
    """

    n_cols = len(POSITIONS_HEADERS)

    for row in ws.iter_rows(
        min_row=start_row, max_row=start_row + clear_rows, min_col=1, max_col=n_cols
    ):
        for cell in row:
            cell.value = None

    ws.cell(row=start_row, column=1, value="POSITIONS")

    header_row = start_row + 1
    for col, header in enumerate(POSITIONS_HEADERS, start=1):
        ws.cell(row=header_row, column=col, value=header)

    for i, (instrument, meta) in enumerate(imported):
        row_values = _position_row(instrument, meta)
        for col, value in enumerate(row_values, start=1):
            ws.cell(row=header_row + 1 + i, column=col, value=value)

    return ws


def _round_greek(key, value):
    return round(value, 3) if key == "Value ($k)" else round(value, 2)


def write_greeks_sheet(ws, portfolio, market, contracts, clear_rows=500):
    """
    One row per contract, each computed with that contract as the shocked
    "Future Spot" -- same GreekEngine.report_dollars() used by the
    Streamlit Greeks tab, same decimal precision (Value ($k) at 3dp,
    every other Greek at 2dp). Gives all contracts at once instead of
    one-at-a-time like the Streamlit selector, since this is a static
    snapshot, not a live dropdown.
    """

    if not contracts:
        return ws

    sample = GreekEngine.report_dollars(portfolio, market, contracts[0])
    greek_names = list(sample.keys())
    n_cols = 1 + len(greek_names)

    for row in ws.iter_rows(min_row=1, max_row=clear_rows, min_col=1, max_col=n_cols):
        for cell in row:
            cell.value = None

    header = ["Contract"] + greek_names
    for col, value in enumerate(header, start=1):
        ws.cell(row=1, column=col, value=value)

    for i, contract in enumerate(contracts):
        dollars = GreekEngine.report_dollars(portfolio, market, contract)
        row_values = [contract] + [_round_greek(k, v) for k, v in dollars.items()]
        for col, value in enumerate(row_values, start=1):
            ws.cell(row=2 + i, column=col, value=value)

    return ws


def write_spot_vol_ladder_sheet(ws, portfolio, market, contract, clear_rows=500, clear_cols=12):
    """
    Spot ladder table, then a gap, then vol ladder table below it --
    both for the one selected Future Spot contract, same shape/columns
    as the Streamlit Ladders tab.
    """

    clear_range(ws, min_row=1, max_row=clear_rows, min_col=1, max_col=clear_cols)

    spot_df = spot_ladder(portfolio, market, contract)
    next_row = write_dataframe(ws, spot_df, start_row=1, label="SPOT LADDER")

    vol_df = vol_ladder(portfolio, market, contract)
    write_dataframe(ws, vol_df, start_row=next_row + 1, label="VOL LADDER")

    return ws


def write_heat_maps_sheet(ws, portfolio, market, contract, metrics=("PnL", "Delta", "Gamma", "Vega"), clear_rows=200, clear_cols=20):
    """
    One spot x vol grid per metric, stacked vertically -- rows are spot
    moves, columns are vol moves, same shape as the Streamlit Surface
    tab's heatmap, just as plain numbers (conditional formatting /
    color scale is applied in Excel itself, not here).
    """

    clear_range(ws, min_row=1, max_row=clear_rows, min_col=1, max_col=clear_cols)

    row = 1
    for metric in metrics:
        df = spot_vol_surface(portfolio, market, contract, metric=metric)
        row = write_dataframe(
            ws, df, start_row=row, label=f"{metric.upper()} (rows=spot move, cols=vol move)",
            include_index=True, index_label="Spot\\Vol",
        ) + 1

    return ws


def write_stress_sheet(ws, portfolio, market, contract, clear_rows=200, clear_cols=12):
    """
    stress_report() returns a dict keyed "<greek>_<axis>" -> DataFrame --
    writes each one as its own labeled block, stacked vertically, same
    grid as the Streamlit Stress tab.
    """

    clear_range(ws, min_row=1, max_row=clear_rows, min_col=1, max_col=clear_cols)

    report = stress_report(portfolio, market, contract)

    row = 1
    for name, df in report.items():
        row = write_dataframe(
            ws, df, start_row=row, label=name.upper(),
            include_index=True, index_label=df.index.name or "",
        ) + 1

    return ws


def write_strike_sheet(ws, portfolio, market, weight="quantity", clear_rows=200, clear_cols=20):
    """
    Pivot table: rows=strike, columns=expiry_days -- spans every contract
    in the book at once (no Future Spot selection needed), same as the
    Streamlit Strike Map tab.
    """

    clear_range(ws, min_row=1, max_row=clear_rows, min_col=1, max_col=clear_cols)

    grid = strike_map(portfolio, market=market if weight == "delta" else None, weight=weight)

    if grid.empty:
        ws.cell(row=1, column=1, value="No option legs -- Strike Map excludes futures, which have no strike.")
        return ws

    write_dataframe(
        ws, grid, start_row=1, label=f"STRIKE MAP (weight={weight})",
        include_index=True, index_label="Strike\\Expiry(d)",
    )

    return ws


def write_breakevens_sheet(ws, portfolio, market, contract, clear_rows=50, clear_cols=6):
    """
    Both breakeven methods (quadratic Taylor estimate and the exact
    numerical one), side by side for comparison -- same two numbers
    shown on the Streamlit Breakevens tab.
    """

    clear_range(ws, min_row=1, max_row=clear_rows, min_col=1, max_col=clear_cols)

    quad = breakevens_quadratic(portfolio, market, contract)
    numeric = breakevens_numerical(portfolio, market, contract)

    ws.cell(row=1, column=1, value="Method")
    ws.cell(row=1, column=2, value="Downside (dF)")
    ws.cell(row=1, column=3, value="Upside (dF)")

    ws.cell(row=2, column=1, value="Quadratic (fast estimate)")
    ws.cell(row=2, column=2, value=round(quad["downside"], 4) if quad["downside"] is not None else None)
    ws.cell(row=2, column=3, value=round(quad["upside"], 4) if quad["upside"] is not None else None)

    ws.cell(row=3, column=1, value="Numerical (exact)")
    ws.cell(row=3, column=2, value=round(numeric["downside"], 4) if numeric["downside"] is not None else None)
    ws.cell(row=3, column=3, value=round(numeric["upside"], 4) if numeric["upside"] is not None else None)

    return ws


def write_payoff_sheet(ws, portfolio, market, contract, time_steps=(0, 30, 64), clear_rows=200, clear_cols=12):
    """P&L vs. spot, overlaid across time-to-expiry steps -- same columns as the Streamlit Payoff tab."""

    clear_range(ws, min_row=1, max_row=clear_rows, min_col=1, max_col=clear_cols)

    df = payoff_diagram(portfolio, market, list(time_steps), contract)
    write_dataframe(ws, df, start_row=1, label="PAYOFF")

    return ws
