"""
xlwings equivalents of the openpyxl writers in excel_sync.py -- these
write directly into a LIVE, already-open Excel workbook via COM, so
changes appear instantly, no file lock, no "close Excel first."

Same cell layout as excel_sync.py (imported from there, not duplicated),
so existing charts/conditional formatting/data validation -- all built
once via openpyxl -- keep referencing the right ranges; this module only
ever touches values, never structure.
"""

from vol_calculator.excel_sync import (
    FUTURE_SPOT_LABEL_CELL,
    FUTURE_SPOT_VALUE_CELL,
    INTEREST_RATE_LABEL_CELL,
    INTEREST_RATE_VALUE_CELL,
    SOFR_CAPTION_CELL,
    POSITIONS_HEADERS,
    _position_row,
    _round_greek,
)
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

_XL_VALIDATE_LIST = 3


def xw_write_future_spot_selector(sheet, contracts):

    sheet[FUTURE_SPOT_LABEL_CELL].value = "Future Spot:"

    if not contracts:
        return

    value_cell = sheet[FUTURE_SPOT_VALUE_CELL]

    if value_cell.value not in contracts:
        value_cell.value = contracts[0]

    # Clearing just this one cell's validation isn't enough -- repeated
    # automated Delete()+Add() cycles on the same cell can corrupt a
    # sheet's internal validation state until every subsequent Add()
    # silently fails, even on an untouched cell. Wiping the whole
    # sheet's validation first resets that reliably.
    try:
        sheet.api.Cells.Validation.Delete()
    except Exception:
        pass

    # AlertStyle/Operator have no real effect for a list-type validation,
    # but must be passed explicitly -- omitting them (relying on COM to
    # fill in VBA's optional-argument defaults) fails when called this
    # way, through raw COM dispatch rather than from inside VBA itself.
    value_cell.api.Validation.Add(
        Type=_XL_VALIDATE_LIST, AlertStyle=1, Operator=1, Formula1=",".join(contracts)
    )


def xw_read_future_spot(sheet, contracts):

    value = sheet[FUTURE_SPOT_VALUE_CELL].value
    return value if value in contracts else (contracts[0] if contracts else None)


def xw_write_interest_rate_cell(sheet, sofr=None, fallback_rate=0.039, error=None):

    sheet[INTEREST_RATE_LABEL_CELL].value = "Interest Rate (%):"

    if sofr:
        sheet[INTEREST_RATE_VALUE_CELL].value = round(sofr["rate"] * 100, 2)
    elif sheet[INTEREST_RATE_VALUE_CELL].value is None:
        sheet[INTEREST_RATE_VALUE_CELL].value = round(fallback_rate * 100, 2)

    if sofr:
        sheet[SOFR_CAPTION_CELL].value = (
            f"SOFR: {sofr['rate'] * 100:.2f}% (updated {sofr['effective_date']}) -- {sofr['source_url']}"
        )
    elif error:
        sheet[SOFR_CAPTION_CELL].value = f"SOFR fetch failed -- using manually entered rate. Error: {error}"
    else:
        sheet[SOFR_CAPTION_CELL].value = "SOFR unavailable -- using manually entered rate."


def xw_read_interest_rate(sheet, fallback_rate=0.039):

    value = sheet[INTEREST_RATE_VALUE_CELL].value
    try:
        return float(value) / 100
    except (TypeError, ValueError):
        return fallback_rate


def xw_write_positions_table(sheet, imported, start_row=13, clear_rows=500):

    n_cols = len(POSITIONS_HEADERS)

    clear_range = sheet.range((start_row, 1), (start_row + clear_rows, n_cols))
    clear_range.value = [[None] * n_cols for _ in range(clear_rows + 1)]

    sheet.range((start_row, 1)).value = "POSITIONS"

    header_row = start_row + 1
    sheet.range((header_row, 1)).value = POSITIONS_HEADERS

    if imported:
        rows = [_position_row(instrument, meta) for instrument, meta in imported]
        sheet.range((header_row + 1, 1)).value = rows


def xw_write_greeks_sheet(sheet, portfolio, market, contracts, clear_rows=500):

    if not contracts:
        return

    sample = GreekEngine.report_dollars(portfolio, market, contracts[0])
    greek_names = list(sample.keys())
    n_cols = 1 + len(greek_names)

    clear_range = sheet.range((1, 1), (clear_rows, n_cols))
    clear_range.value = [[None] * n_cols for _ in range(clear_rows)]

    sheet.range((1, 1)).value = ["Contract"] + greek_names

    rows = []
    for contract in contracts:
        dollars = GreekEngine.report_dollars(portfolio, market, contract)
        rows.append([contract] + [_round_greek(k, v) for k, v in dollars.items()])

    sheet.range((2, 1)).value = rows


def xw_write_dataframe(sheet, df, start_row=1, start_col=1, label=None, include_index=False, index_label=""):
    """Bulk-writes a DataFrame as label + header + rows, one COM call per block. Mirrors excel_sync.write_dataframe's layout exactly."""

    row = start_row

    if label is not None:
        sheet.range((row, start_col)).value = label
        row += 1

    header = ([index_label] if include_index else []) + list(df.columns)
    sheet.range((row, start_col)).value = header
    row += 1

    if len(df):
        data_rows = []
        for idx, data_row in df.iterrows():
            vals = [round(v, 4) if isinstance(v, float) else v for v in data_row]
            if include_index:
                vals = [idx] + vals
            data_rows.append(vals)
        sheet.range((row, start_col)).value = data_rows
        row += len(data_rows)

    return row


def xw_write_spot_vol_ladder_sheet(sheet, portfolio, market, contract, clear_rows=500, clear_cols=12):

    sheet.range((1, 1), (clear_rows, clear_cols)).clear_contents()

    spot_df = spot_ladder(portfolio, market, contract)
    next_row = xw_write_dataframe(sheet, spot_df, start_row=1, label="SPOT LADDER")

    vol_df = vol_ladder(portfolio, market, contract)
    xw_write_dataframe(sheet, vol_df, start_row=next_row + 1, label="VOL LADDER")


def xw_write_heat_maps_sheet(sheet, portfolio, market, contract, metrics=("PnL", "Delta", "Gamma", "Vega"), clear_rows=200, clear_cols=20):
    """Values only -- the color-scale conditional formatting is already baked into the file from the openpyxl pass and keeps referencing these same ranges."""

    sheet.range((1, 1), (clear_rows, clear_cols)).clear_contents()

    row = 1
    for metric in metrics:
        df = spot_vol_surface(portfolio, market, contract, metric=metric)
        row = xw_write_dataframe(
            sheet, df, start_row=row, label=f"{metric.upper()} (rows=spot move, cols=vol move)",
            include_index=True, index_label="Spot\\Vol",
        ) + 1


def xw_write_stress_sheet(sheet, portfolio, market, contract, clear_rows=200, clear_cols=12):

    sheet.range((1, 1), (clear_rows, clear_cols)).clear_contents()

    report = stress_report(portfolio, market, contract)

    row = 1
    for name, df in report.items():
        row = xw_write_dataframe(
            sheet, df, start_row=row, label=name.upper(),
            include_index=True, index_label=df.index.name or "",
        ) + 1


def xw_write_strike_sheet(sheet, portfolio, market, weight="quantity", clear_rows=200, clear_cols=20):

    sheet.range((1, 1), (clear_rows, clear_cols)).clear_contents()

    grid = strike_map(portfolio, market=market if weight == "delta" else None, weight=weight)

    if grid.empty:
        sheet.range((1, 1)).value = "No option legs -- Strike Map excludes futures, which have no strike."
        return

    xw_write_dataframe(
        sheet, grid, start_row=1, label=f"STRIKE MAP (weight={weight})",
        include_index=True, index_label="Strike\\Expiry(d)",
    )


def xw_write_payoff_sheet(sheet, portfolio, market, contract, time_steps=(0, 30, 64), clear_rows=200, clear_cols=12):
    """Values only -- the line chart is already baked into the file from the openpyxl pass and keeps referencing these same ranges."""

    sheet.range((1, 1), (clear_rows, clear_cols)).clear_contents()

    df = payoff_diagram(portfolio, market, list(time_steps), contract)
    xw_write_dataframe(sheet, df, start_row=1, label="PAYOFF")


def xw_write_breakevens_sheet(sheet, portfolio, market, contract, low=-15, high=15, step=1, clear_rows=60, clear_cols=8):
    """Values only -- the parabola chart is already baked into the file from the openpyxl pass and keeps referencing these same ranges."""

    sheet.range((1, 1), (clear_rows, clear_cols)).clear_contents()

    quad = breakevens_quadratic(portfolio, market, contract)
    numeric = breakevens_numerical(portfolio, market, contract)

    sheet.range((1, 1)).value = ["Method", "Downside (dF)", "Upside (dF)"]
    sheet.range((2, 1)).value = [
        "Quadratic (fast estimate)",
        round(quad["downside"], 4) if quad["downside"] is not None else None,
        round(quad["upside"], 4) if quad["upside"] is not None else None,
    ]
    sheet.range((3, 1)).value = [
        "Numerical (exact)",
        round(numeric["downside"], 4) if numeric["downside"] is not None else None,
        round(numeric["upside"], 4) if numeric["upside"] is not None else None,
    ]

    delta = GreekEngine.delta(portfolio, market, contract)
    gamma = GreekEngine.gamma(portfolio, market, contract)
    theta = GreekEngine.theta(portfolio, market)
    base_price = market.price_for(contract)

    numeric_df = payoff_diagram(portfolio, market, [1], contract, low=low, high=high, step=step)

    header_row = 5
    sheet.range((header_row, 1)).value = ["Spot Move %", "Quadratic PnL", "Numerical PnL (T+1d)", "Zero"]

    data_rows = []
    for _, r in numeric_df.iterrows():
        pct = r["Spot Move %"]
        dF = base_price * pct / 100
        quad_pnl = delta * dF + 0.5 * gamma * dF ** 2 + theta
        data_rows.append([pct, round(quad_pnl, 4), round(r["T+1d"], 4), 0])

    sheet.range((header_row + 1, 1)).value = data_rows
