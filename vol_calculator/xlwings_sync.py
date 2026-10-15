"""
Live (xlwings/COM) access to the open Excel workbook -- changes appear
instantly, no file lock, no "close Excel first."

The Home sheet is driven entirely through the hidden _Data sheet and
named input cells (see below): Python only writes values there and reads
inputs by name, never touching visible layout or formatting. The
analytic sheets (Greeks, ladders, heat maps, ...) are still written
directly by the xw_write_*_sheet functions further down.
"""

from vol_calculator.excel_sync import POSITIONS_HEADERS, _round_greek
from vol_calculator.greeks import GreekEngine
from vol_calculator.cqg_quotes import _convert_cqg_symbol
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

# The live CQGXL watchlist, one named block per commodity: a symbol
# column and a price column, one row per contract. Found by name, like
# every other input, so the watchlist can be moved anywhere on the
# workbook. Same quotes push_quotes_supabase.py sends to Streamlit.
WATCHLIST_NAMES = {
    "Watchlist_Corn": "Corn",
    "Watchlist_Soybean": "Soybean",
    "Watchlist_Wheat": "SRW Wheat",
}


def xw_read_live_quotes(wb):
    """Reads the live CQGXL watchlist -- [{"commodity", "contract", "price"}, ...]. Skips a cell showing text (e.g. "Disconnected from server.") instead of a number, and a commodity whose block isn't named in this workbook."""

    quotes = []

    for name, commodity in WATCHLIST_NAMES.items():
        try:
            block = wb.names[name].refers_to_range.options(ndim=2).value
        except Exception:
            continue

        for symbol, price, *_ in block:
            if not symbol or price is None:
                continue

            try:
                price = float(price)
            except (TypeError, ValueError):
                continue

            quotes.append({
                "commodity": commodity,
                "contract": _convert_cqg_symbol(str(symbol)),
                "price": price,
            })

    return quotes


_XL_CONTINUOUS = 1
_XL_THIN = 2


# --- Hidden _Data sheet ------------------------------------------------------
#
# Python writes plain values -- no styling, labels or layout -- to a
# hidden "_Data" sheet, and keeps a workbook-level name sized to each
# block's exact extent. The visible sheets reference those names with
# ordinary formulas (e.g. =Positions spills the whole table), so the
# layout and formatting of everything visible belongs to the workbook
# itself and a refresh never overwrites it.

DATA_SHEET = "_Data"

# The book: every leg, whether imported from a Daily Prop file or added
# by hand, as entered -- plus each leg's settlement from the file it came
# from (the price fallback for a contract the watchlist doesn't quote)
# and its real expiry date, so days-to-expiry can be recomputed daily.
BOOK_HEADERS = POSITIONS_HEADERS + ["Settlement", "Expiry Date"]

# name -> (first column, width); every block starts at row 1. Columns
# AO-AQ hold the workbook's own helper formulas (the filtered dropdown
# lists for Add Leg) and are never written here.
_DATA_BLOCKS = {
    "Positions": ("A", len(POSITIONS_HEADERS)),
    "Book": ("AS", len(BOOK_HEADERS)),
    "ContractList": ("Y", 1),
    "PortfolioContracts": ("Z", 1),
    "LegTypes": ("AA", 1),
    "OptionTypes": ("AB", 1),
    "ContractQuotes": ("AG", 3),    # contract, commodity, current price
    "ExpiryCandidates": ("AK", 3),  # contract, label, days to expiry
}
_DATA_CLEAR_ROWS = 2000

# name -> cell, for single values.
_DATA_VALUES = {
    "SofrRate": "AE1",
    "SofrCaption": "AE2",
    "SofrUrl": "AE3",
    "AddLegStatus": "AE4",
}


# Missing values are written as truly empty cells. A formula referencing
# an empty cell shows 0, so visible formulas wrap the name to keep them
# blank, e.g. =IF(Positions="","",Positions).

def xw_write_data_block(wb, name, rows):

    first_col, width = _DATA_BLOCKS[name]
    top_left = wb.sheets[DATA_SHEET].range(f"{first_col}1")
    top_left.resize(_DATA_CLEAR_ROWS, width).clear_contents()

    if rows:
        top_left.value = rows

    block = top_left.resize(max(len(rows), 1), width)
    wb.names.add(name, f"='{DATA_SHEET}'!{block.address}")


def xw_read_data_block(wb, name):

    try:
        block = wb.names[name].refers_to_range
    except Exception:
        return []

    return block.options(ndim=2).value


def xw_write_data_value(wb, name, value):

    cell = wb.sheets[DATA_SHEET].range(_DATA_VALUES[name])
    cell.value = value
    wb.names.add(name, f"='{DATA_SHEET}'!{cell.address}")


def xw_read_data_value(wb, name):

    return wb.sheets[DATA_SHEET].range(_DATA_VALUES[name]).value


def xw_write_book(wb, rows):

    xw_write_data_block(
        wb, "Book", [BOOK_HEADERS] + [[row.get(h) for h in BOOK_HEADERS] for row in rows],
    )


def xw_read_book(wb):
    """The book is stored on _Data, not read back off the visible Positions table -- that table can be rearranged freely without losing legs."""

    block = xw_read_data_block(wb, "Book")
    if not block or block[0] != BOOK_HEADERS:
        return []
    return [dict(zip(BOOK_HEADERS, row)) for row in block[1:] if row[0] is not None]


# --- Named input cells ---------------------------------------------------------
#
# Every input is found by workbook name, never by cell address, so the
# input cells can be moved anywhere in the workbook without touching code.

FUTURE_SPOT_NAME = "FutureSpot"
INTEREST_RATE_NAME = "InterestRate"

ADD_LEG_INPUT_NAMES = {
    "type": "AddLeg_Type",
    "contract": "AddLeg_Contract",
    "quantity": "AddLeg_Quantity",
    "quantity2": "AddLeg_Quantity2",
    "strike": "AddLeg_Strike1",
    "strike2": "AddLeg_Strike2",
    "strike3": "AddLeg_Strike3",
    "expiry_days": "AddLeg_Expiry",
    "expiry_days2": "AddLeg_Expiry2",
    "vol_pct": "AddLeg_Vol1",
    "vol_pct2": "AddLeg_Vol2",
    "vol_pct3": "AddLeg_Vol3",
    "option_type": "AddLeg_OptionType",
    "entry_price": "AddLeg_EntryPrice",
}


def xw_read_named(wb, name):
    return wb.names[name].refers_to_range.value


def xw_read_add_leg_inputs(wb):
    return {key: xw_read_named(wb, name) for key, name in ADD_LEG_INPUT_NAMES.items()}


def xw_clear_add_leg_inputs(wb):
    """Type and Contract are kept -- only the per-leg numbers are cleared after a successful add."""

    for key, name in ADD_LEG_INPUT_NAMES.items():
        if key not in ("type", "contract"):
            wb.names[name].refers_to_range.value = None


def xw_read_future_spot(wb, contracts):

    value = xw_read_named(wb, FUTURE_SPOT_NAME)
    return value if value in contracts else (contracts[0] if contracts else None)


def xw_read_interest_rate(wb, fallback_rate=0.039):

    try:
        return float(xw_read_named(wb, INTEREST_RATE_NAME)) / 100
    except (TypeError, ValueError):
        return fallback_rate


_TABLE_LABEL_RGB = (46, 117, 182)
_TABLE_HEADER_RGB = (217, 225, 242)
_TABLE_BAND_RGB = (247, 247, 247)


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

    if rows:
        sheet.range((2, 1)).value = rows

    header = sheet.range((1, 1), (1, n_cols))
    header.color = _TABLE_HEADER_RGB
    header.api.Font.Bold = True
    header.api.BorderAround(LineStyle=_XL_CONTINUOUS, Weight=_XL_THIN)

    for i in range(len(rows)):
        row_range = sheet.range((2 + i, 1), (2 + i, n_cols))
        row_range.api.BorderAround(LineStyle=_XL_CONTINUOUS, Weight=_XL_THIN)
        if i % 2 == 1:
            row_range.color = _TABLE_BAND_RGB
        if n_cols > 1:
            sheet.range((2 + i, 2), (2 + i, n_cols)).number_format = "#,##0.00"


def xw_write_dataframe(sheet, df, start_row=1, start_col=1, label=None, include_index=False, index_label=""):
    """Bulk-writes a DataFrame as a styled label + header + rows block (title bar, header row, banded/bordered data rows), one COM call per chunk. Mirrors excel_sync.write_dataframe's layout and look."""

    row = start_row
    col_offset = start_col + (1 if include_index else 0)
    n_cols = len(df.columns) + (col_offset - start_col)

    if label is not None:
        sheet.range((row, start_col)).value = label
        title = sheet.range((row, start_col), (row, start_col + n_cols - 1))
        title.api.Merge()
        title.color = _TABLE_LABEL_RGB
        title.api.Font.Color = 0xFFFFFF
        title.api.Font.Bold = True
        title.api.HorizontalAlignment = -4108  # xlCenter
        row += 1

    header_row = row
    header = ([index_label] if include_index else []) + list(df.columns)
    sheet.range((row, start_col)).value = header
    header_range = sheet.range((header_row, start_col), (header_row, start_col + n_cols - 1))
    header_range.color = _TABLE_HEADER_RGB
    header_range.api.Font.Bold = True
    header_range.api.BorderAround(LineStyle=_XL_CONTINUOUS, Weight=_XL_THIN)
    row += 1

    if len(df):
        data_rows = []
        for idx, data_row in df.iterrows():
            vals = [round(v, 4) if isinstance(v, float) else v for v in data_row]
            if include_index:
                vals = [idx] + vals
            data_rows.append(vals)
        sheet.range((row, start_col)).value = data_rows

        for i in range(len(data_rows)):
            row_range = sheet.range((row + i, start_col), (row + i, start_col + n_cols - 1))
            row_range.api.BorderAround(LineStyle=_XL_CONTINUOUS, Weight=_XL_THIN)
            if i % 2 == 1:
                row_range.color = _TABLE_BAND_RGB
        sheet.range(
            (row, col_offset), (row + len(data_rows) - 1, start_col + n_cols - 1)
        ).number_format = "#,##0.00"

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

    summary_header = sheet.range((1, 1), (1, 3))
    summary_header.color = _TABLE_HEADER_RGB
    summary_header.api.Font.Bold = True
    summary_header.api.BorderAround(LineStyle=_XL_CONTINUOUS, Weight=_XL_THIN)
    for r in (2, 3):
        row_range = sheet.range((r, 1), (r, 3))
        row_range.api.BorderAround(LineStyle=_XL_CONTINUOUS, Weight=_XL_THIN)
        if r % 2 == 1:
            row_range.color = _TABLE_BAND_RGB
    sheet.range((2, 2), (3, 3)).number_format = "#,##0.00"

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

    header_range = sheet.range((header_row, 1), (header_row, 4))
    header_range.color = _TABLE_HEADER_RGB
    header_range.api.Font.Bold = True
    header_range.api.BorderAround(LineStyle=_XL_CONTINUOUS, Weight=_XL_THIN)
    for i in range(len(data_rows)):
        row_range = sheet.range((header_row + 1 + i, 1), (header_row + 1 + i, 4))
        row_range.api.BorderAround(LineStyle=_XL_CONTINUOUS, Weight=_XL_THIN)
        if i % 2 == 1:
            row_range.color = _TABLE_BAND_RGB
    if data_rows:
        sheet.range((header_row + 1, 1), (header_row + len(data_rows), 4)).number_format = "#,##0.00"
