from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.marker import Marker

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


INTEREST_RATE_LABEL_CELL = "D10"
INTEREST_RATE_VALUE_CELL = "E10"
SOFR_CAPTION_CELL = "D11"


def write_interest_rate_cell(ws, sofr=None, fallback_rate=0.039,
                              label_cell=INTEREST_RATE_LABEL_CELL, value_cell=INTEREST_RATE_VALUE_CELL,
                              caption_cell=SOFR_CAPTION_CELL):
    """
    Interest Rate (%) cell, pre-filled from live SOFR when available --
    same pattern as the Streamlit Market panel: this is a plain editable
    number, not a formula, so overwriting it with your own view sticks
    (re-running the refresh script won't clobber a value you've already
    changed, unless the cell is still empty).

    sofr: the dict fetch_latest_sofr() returns, or None if that call
    failed -- falls back to fallback_rate and says so in the caption.
    """

    ws[label_cell] = "Interest Rate (%):"

    if ws[value_cell].value is None:
        ws[value_cell] = round((sofr["rate"] if sofr else fallback_rate) * 100, 2)

    if sofr:
        ws[caption_cell] = (
            f"SOFR: {sofr['rate'] * 100:.2f}% (updated {sofr['effective_date']}) -- {sofr['source_url']}"
        )
    else:
        ws[caption_cell] = "SOFR unavailable -- using manually entered rate."


def read_interest_rate(ws, fallback_rate=0.039, value_cell=INTEREST_RATE_VALUE_CELL):
    """Reads the Interest Rate (%) cell as a decimal, falling back if empty/invalid."""

    value = ws[value_cell].value
    try:
        return float(value) / 100
    except (TypeError, ValueError):
        return fallback_rate


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
    tab's heatmap. Each grid gets a real red/white/green color-scale
    (negative/zero/positive), same visual as the matplotlib version,
    via Excel's native conditional formatting instead of an image.
    """

    clear_range(ws, min_row=1, max_row=clear_rows, min_col=1, max_col=clear_cols)
    ws.conditional_formatting = ConditionalFormattingList()

    row = 1
    for metric in metrics:
        df = spot_vol_surface(portfolio, market, contract, metric=metric)

        block_start = row
        row = write_dataframe(
            ws, df, start_row=block_start, label=f"{metric.upper()} (rows=spot move, cols=vol move)",
            include_index=True, index_label="Spot\\Vol",
        )

        data_start_row = block_start + 2
        data_end_row = row - 1
        data_start_col = get_column_letter(2)
        data_end_col = get_column_letter(1 + len(df.columns))
        cell_range = f"{data_start_col}{data_start_row}:{data_end_col}{data_end_row}"

        ws.conditional_formatting.add(
            cell_range,
            ColorScaleRule(
                start_type="min", start_color="F8696B",
                mid_type="num", mid_value=0, mid_color="FFFFFF",
                end_type="max", end_color="63BE7B",
            ),
        )

        row += 1

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


def write_breakevens_sheet(ws, portfolio, market, contract, low=-15, high=15, step=1, clear_rows=60, clear_cols=8):
    """
    Summary numbers (both breakeven methods) plus the actual parabola:
    Quadratic PnL(dF) = Delta*dF + 0.5*Gamma*dF^2 + Theta, charted
    alongside the exact next-day PnL curve (via payoff_diagram at
    time_steps=[1], same logic breakevens_numerical itself uses) and a
    zero line, so the breakeven points are visible as where the curves
    cross zero, not just the two numbers.
    """

    clear_range(ws, min_row=1, max_row=clear_rows, min_col=1, max_col=clear_cols)
    for chart in list(ws._charts):
        ws._charts.remove(chart)

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

    delta = GreekEngine.delta(portfolio, market, contract)
    gamma = GreekEngine.gamma(portfolio, market, contract)
    theta = GreekEngine.theta(portfolio, market)
    base_price = market.price_for(contract)

    numeric_df = payoff_diagram(portfolio, market, [1], contract, low=low, high=high, step=step)

    header_row = 5
    for col, h in enumerate(["Spot Move %", "Quadratic PnL", "Numerical PnL (T+1d)", "Zero"], start=1):
        ws.cell(row=header_row, column=col, value=h)

    row = header_row + 1
    for _, r in numeric_df.iterrows():
        pct = r["Spot Move %"]
        dF = base_price * pct / 100
        quad_pnl = delta * dF + 0.5 * gamma * dF ** 2 + theta
        ws.cell(row=row, column=1, value=pct)
        ws.cell(row=row, column=2, value=round(quad_pnl, 4))
        ws.cell(row=row, column=3, value=round(r["T+1d"], 4))
        ws.cell(row=row, column=4, value=0)
        row += 1

    data_end_row = row - 1

    chart = LineChart()
    chart.title = f"Breakeven Parabola -- {contract}"
    chart.x_axis.title = "Spot Move %"
    chart.y_axis.title = "PnL"
    chart.width, chart.height = 24, 12

    cats = Reference(ws, min_col=1, min_row=header_row + 1, max_row=data_end_row)
    for col in (2, 3, 4):
        data = Reference(ws, min_col=col, min_row=header_row, max_row=data_end_row)
        chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    for series in chart.series[:2]:  # Quadratic, Numerical -- real curves get markers
        series.marker = Marker(symbol="circle", size=5)
        series.smooth = False

    zero_line = chart.series[2]  # Zero reference -- dashed, no markers, muted gray
    zero_line.marker = Marker(symbol="none")
    zero_line.smooth = False
    zero_line.graphicalProperties.line.dashStyle = "dash"
    zero_line.graphicalProperties.line.solidFill = "999999"

    ws.add_chart(chart, f"F{header_row}")

    return ws


def write_payoff_sheet(ws, portfolio, market, contract, time_steps=(0, 30, 64), clear_rows=200, clear_cols=12):
    """P&L vs. spot, overlaid across time-to-expiry steps -- same columns as the Streamlit Payoff tab, plus a line chart."""

    clear_range(ws, min_row=1, max_row=clear_rows, min_col=1, max_col=clear_cols)
    for chart in list(ws._charts):
        ws._charts.remove(chart)

    df = payoff_diagram(portfolio, market, list(time_steps), contract)
    next_row = write_dataframe(ws, df, start_row=1, label="PAYOFF")

    header_row = 2
    data_start_row = 3
    data_end_row = next_row - 1
    n_series = len(df.columns) - 2  # exclude "Spot Move %" and "Futures Price"

    chart = LineChart()
    chart.title = f"Payoff -- {contract}"
    chart.x_axis.title = "Spot Move %"
    chart.y_axis.title = "PnL"
    chart.width, chart.height = 24, 12

    cats = Reference(ws, min_col=1, min_row=data_start_row, max_row=data_end_row)

    for col in range(3, 3 + n_series):  # PnL columns start at col 3 (Spot Move %, Futures Price, then T+Xd...)
        data = Reference(ws, min_col=col, min_row=header_row, max_row=data_end_row)
        chart.add_data(data, titles_from_data=True)

    chart.set_categories(cats)

    for series in chart.series:
        series.marker = Marker(symbol="circle", size=4)
        series.smooth = False

    ws.add_chart(chart, f"A{next_row + 2}")

    return ws
