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
    _LIVE_QUOTES_BLOCKS, _LIVE_QUOTES_HEADER_ROW, _LIVE_QUOTES_DATA_ROWS,
    INTEREST_RATE_LABEL_CELL,
    INTEREST_RATE_VALUE_CELL,
    SOFR_CAPTION_CELL,
    POSITIONS_HEADERS,
    ADD_LEG_LABEL_CELL,
    ADD_LEG_TYPE_LABEL_CELL, ADD_LEG_TYPE_VALUE_CELL,
    ADD_LEG_CONTRACT_LABEL_CELL, ADD_LEG_CONTRACT_VALUE_CELL,
    ADD_LEG_OPTION_TYPE_LABEL_CELL, ADD_LEG_OPTION_TYPE_VALUE_CELL,
    ADD_LEG_QUANTITY_LABEL_CELL, ADD_LEG_QUANTITY_VALUE_CELL,
    ADD_LEG_STRIKE_LABEL_CELL, ADD_LEG_STRIKE_VALUE_CELL,
    ADD_LEG_EXPIRY_LABEL_CELL, ADD_LEG_EXPIRY_VALUE_CELL,
    ADD_LEG_VOL_LABEL_CELL, ADD_LEG_VOL_VALUE_CELL,
    ADD_LEG_STATUS_CELL,
    ADD_LEG_TYPES, ADD_LEG_OPTION_TYPES,
    _position_row,
    _round_greek,
)
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

_XL_VALIDATE_LIST = 3

# (commodity, symbol_col, price_col) -- matches the Home sheet's own live
# CQGXL watchlist layout (rows 3-8 under each commodity header). Same
# block push_quotes_supabase.py reads to feed Streamlit; shared here so
# xlwings_refresh.py can price contracts that aren't in today's Daily
# Prop file too (e.g. a brand-new leg added by hand), instead of only
# ever having a price when CSV-imported.
_CQG_BLOCKS = [("Corn", 1, 2), ("Soybean", 5, 6), ("SRW Wheat", 9, 10)]
_CQG_ROWS = range(3, 9)


def xw_read_live_quotes(sheet):
    """Reads the Home sheet's live CQGXL watchlist -- [{"commodity", "contract", "price"}, ...]. Skips a cell showing text (e.g. "Disconnected from server.") instead of a number."""

    quotes = []

    for commodity, symbol_col, price_col in _CQG_BLOCKS:
        for row in _CQG_ROWS:
            symbol = sheet.range((row, symbol_col)).value
            price = sheet.range((row, price_col)).value

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


_LIVE_QUOTES_HEADER_RGB = (191, 143, 0)


def xw_style_live_quotes_watchlist(sheet):
    """Live mirror of excel_sync.style_live_quotes_watchlist -- styling only, never touches the live RTD-linked price values themselves."""

    for first_col, last_col in _LIVE_QUOTES_BLOCKS:
        header = sheet.range((_LIVE_QUOTES_HEADER_ROW, first_col), (_LIVE_QUOTES_HEADER_ROW, last_col))
        header.color = _LIVE_QUOTES_HEADER_RGB
        header.api.Font.Color = 0xFFFFFF
        header.api.Font.Bold = True

        block = sheet.range((_LIVE_QUOTES_DATA_ROWS[0], first_col), (_LIVE_QUOTES_DATA_ROWS[1], last_col))
        block.api.BorderAround(LineStyle=_XL_CONTINUOUS, Weight=_XL_THIN)

        price_col = first_col + 1
        sheet.range(
            (_LIVE_QUOTES_DATA_ROWS[0], price_col), (_LIVE_QUOTES_DATA_ROWS[1], price_col)
        ).number_format = "#,##0.00"


def xw_reset_dropdown_validations(sheet):
    """
    Wipes every data-validation rule on the sheet. Repeated automated
    Delete()+Add() cycles targeting just one cell, across many refreshes,
    corrupt a sheet's internal validation state until every subsequent
    Add() silently fails, even on an untouched cell -- so the reset has
    to cover the whole sheet, and has to happen exactly once per refresh,
    before any of this module's dropdown-writing functions run (they no
    longer wipe individually).
    """

    try:
        sheet.api.Cells.Validation.Delete()
    except Exception:
        pass


def _add_list_validation(cell, items):
    """
    Best-effort: adds a list-type data validation dropdown to a cell.
    AlertStyle/Operator have no real effect for a list-type validation,
    but must be passed explicitly -- omitting them (relying on COM to
    fill in VBA's optional-argument defaults) fails when called this way,
    through raw COM dispatch rather than from inside VBA itself.

    Wrapped in try/except: the sheet's validation state can get corrupted
    by Excel/COM itself after enough Delete()+Add() cycles in one running
    session (confirmed: even Add() on a brand-new, never-validated cell
    starts failing once this happens) -- not something further automated
    Delete()/Add() calls can repair; it clears on its own when the
    workbook/Excel is closed and reopened. A dropdown failing to (re)build
    is cosmetic -- it must never take down the rest of the refresh (Greeks,
    Positions, ladders, etc. all still need to write), so this swallows
    the error rather than raising.
    """

    try:
        cell.api.Validation.Add(
            Type=_XL_VALIDATE_LIST, AlertStyle=1, Operator=1, Formula1=",".join(items)
        )
    except Exception:
        pass


def xw_write_future_spot_selector(sheet, contracts):

    sheet[FUTURE_SPOT_LABEL_CELL].value = "Future Spot:"

    if not contracts:
        return

    value_cell = sheet[FUTURE_SPOT_VALUE_CELL]

    if value_cell.value not in contracts:
        value_cell.value = contracts[0]

    # Caller must have already run xw_reset_dropdown_validations() once
    # this refresh -- see that function's docstring for why.
    _add_list_validation(value_cell, contracts)


def xw_read_future_spot(sheet, contracts):

    value = sheet[FUTURE_SPOT_VALUE_CELL].value
    return value if value in contracts else (contracts[0] if contracts else None)


def xw_write_interest_rate_cell(sheet, sofr=None, fallback_rate=0.039, error=None):

    sheet[INTEREST_RATE_LABEL_CELL].value = "IR (%):"

    if sofr:
        sheet[INTEREST_RATE_VALUE_CELL].value = round(sofr["rate"] * 100, 2)
    elif sheet[INTEREST_RATE_VALUE_CELL].value is None:
        sheet[INTEREST_RATE_VALUE_CELL].value = round(fallback_rate * 100, 2)

    if sofr:
        caption = sheet[SOFR_CAPTION_CELL]
        text = f"SOFR -- last updated {sofr['effective_date']} (click for source)"
        # A cell can't hold two hyperlinks -- clear any existing one
        # (from a previous refresh) before adding the current one.
        if caption.api.Hyperlinks.Count:
            caption.api.Hyperlinks.Delete()
        sheet.api.Hyperlinks.Add(Anchor=caption.api, Address=sofr["source_url"], TextToDisplay=text)
        caption.api.Font.Color = 0xC16305  # BGR for RGB(5,99,193)
        caption.api.Font.Underline = 2  # xlUnderlineStyleSingle
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


def xw_read_positions_rows(sheet, start_row=13, max_rows=500):
    """Live-sheet equivalent of excel_sync.read_positions_rows -- reads the Positions table back as a list of {header: value} dicts."""

    header_row = start_row + 1
    n_cols = len(POSITIONS_HEADERS)

    header_values = sheet.range((header_row, 1), (header_row, n_cols)).value
    if header_values != POSITIONS_HEADERS:
        return []

    data_block = sheet.range((header_row + 1, 1), (header_row + max_rows, n_cols)).value

    rows = []
    for raw_row in data_block:
        if not raw_row or raw_row[0] in (None, ""):
            break
        rows.append(dict(zip(POSITIONS_HEADERS, raw_row)))

    return rows


def xw_write_positions_table(sheet, imported, extra_rows=None, start_row=13, clear_rows=500):
    """
    Live-sheet equivalent of excel_sync.write_positions_table -- see that
    function's docstring. extra_rows (already-built row dicts, e.g.
    Manual-sourced legs read back via xw_read_positions_rows beforehand)
    must be supplied by the caller: this clears the whole range first, so
    it can't read "what to keep" off the sheet itself.
    """

    n_cols = len(POSITIONS_HEADERS)

    clear_range = sheet.range((start_row, 1), (start_row + clear_rows, n_cols))
    clear_range.value = [[None] * n_cols for _ in range(clear_rows + 1)]

    sheet.range((start_row, 1)).value = "POSITIONS"

    header_row = start_row + 1
    sheet.range((header_row, 1)).value = POSITIONS_HEADERS

    rows = [_position_row(instrument, meta, source="Daily Prop") for instrument, meta in (imported or [])]
    rows += [[row.get(h) for h in POSITIONS_HEADERS] for row in (extra_rows or [])]

    if rows:
        sheet.range((header_row + 1, 1)).value = rows

    _xw_style_positions_table(sheet, start_row, header_row, header_row + len(rows), n_cols)


_POSITIONS_TITLE_RGB = (55, 86, 35)
_POSITIONS_HEADER_RGB = (226, 239, 218)
_POSITIONS_BAND_RGB = (247, 247, 247)
_POSITIONS_NUMBER_FORMATS = {
    4: "#,##0.##",
    5: "#,##0.00",
    6: "0",
    7: "0.00",
    8: "#,##0.00",
}


def _xw_style_positions_table(sheet, start_row, header_row, last_data_row, n_cols):
    """Live mirror of excel_sync._style_positions_table -- row-level (not cell-level) fills/borders for fewer COM round-trips."""

    title = sheet.range((start_row, 1), (start_row, n_cols))
    title.api.Merge()
    title.color = _POSITIONS_TITLE_RGB
    title.api.Font.Color = 0xFFFFFF
    title.api.Font.Bold = True
    title.api.HorizontalAlignment = -4108  # xlCenter

    header = sheet.range((header_row, 1), (header_row, n_cols))
    header.color = _POSITIONS_HEADER_RGB
    header.api.Font.Bold = True
    header.api.BorderAround(LineStyle=_XL_CONTINUOUS, Weight=_XL_THIN)

    for row in range(header_row + 1, last_data_row + 1):
        row_range = sheet.range((row, 1), (row, n_cols))
        row_range.api.BorderAround(LineStyle=_XL_CONTINUOUS, Weight=_XL_THIN)
        if (row - header_row) % 2 == 0:
            row_range.color = _POSITIONS_BAND_RGB
        else:
            row_range.color = None

    for col, fmt in _POSITIONS_NUMBER_FORMATS.items():
        if last_data_row >= header_row + 1:
            sheet.range((header_row + 1, col), (last_data_row, col)).number_format = fmt

    # A shrinking table would otherwise leave stale banding below the
    # real data -- clear a reasonable buffer of leftover formatting too.
    buffer_range = sheet.range((last_data_row + 1, 1), (last_data_row + 50, n_cols))
    buffer_range.color = None
    try:
        buffer_range.api.Borders.LineStyle = -4142  # xlLineStyleNone
    except Exception:
        pass


_ADD_LEG_HEADER_RGB = (31, 78, 120)
_ADD_LEG_LABEL_RGB = (221, 235, 247)
_ADD_LEG_INPUT_RGB = (255, 242, 204)
_ADD_LEG_STATUS_RGB = (242, 242, 242)
_XL_CONTINUOUS = 1
_XL_THIN = 2

_ADD_LEG_ROWS = [
    (ADD_LEG_TYPE_LABEL_CELL, ADD_LEG_TYPE_VALUE_CELL),
    (ADD_LEG_CONTRACT_LABEL_CELL, ADD_LEG_CONTRACT_VALUE_CELL),
    (ADD_LEG_OPTION_TYPE_LABEL_CELL, ADD_LEG_OPTION_TYPE_VALUE_CELL),
    (ADD_LEG_QUANTITY_LABEL_CELL, ADD_LEG_QUANTITY_VALUE_CELL),
    (ADD_LEG_STRIKE_LABEL_CELL, ADD_LEG_STRIKE_VALUE_CELL),
    (ADD_LEG_EXPIRY_LABEL_CELL, ADD_LEG_EXPIRY_VALUE_CELL),
    (ADD_LEG_VOL_LABEL_CELL, ADD_LEG_VOL_VALUE_CELL),
]


def _xw_style_add_leg_block(sheet):
    """Live mirror of excel_sync._style_add_leg_block -- same colors, via COM instead of openpyxl styles."""

    # J/K are now a deliberate empty gap between the Positions table and
    # this block -- reset them back to a normal width in case an older
    # refresh widened them when the block used to sit there.
    sheet.range("J:J").column_width = 8.43
    sheet.range("K:K").column_width = 8.43
    sheet.range("L:L").column_width = 28
    sheet.range("M:M").column_width = 16

    header = sheet.range(f"{ADD_LEG_LABEL_CELL}:M13")
    header.api.Merge()
    header.color = _ADD_LEG_HEADER_RGB
    header.api.Font.Color = 0xFFFFFF
    header.api.Font.Bold = True
    header.api.HorizontalAlignment = -4108  # xlCenter
    header.api.BorderAround(LineStyle=_XL_CONTINUOUS, Weight=_XL_THIN)

    for label_cell, value_cell in _ADD_LEG_ROWS:
        label_range = sheet.range(label_cell)
        label_range.color = _ADD_LEG_LABEL_RGB
        label_range.api.Font.Bold = True
        label_range.api.BorderAround(LineStyle=_XL_CONTINUOUS, Weight=_XL_THIN)

        value_range = sheet.range(value_cell)
        value_range.color = _ADD_LEG_INPUT_RGB
        value_range.api.BorderAround(LineStyle=_XL_CONTINUOUS, Weight=_XL_THIN)

    status = sheet.range(f"{ADD_LEG_STATUS_CELL}:M21")
    status.api.Merge()
    status.color = _ADD_LEG_STATUS_RGB
    status.api.Font.Italic = True
    status.api.WrapText = True
    status.api.BorderAround(LineStyle=_XL_CONTINUOUS, Weight=_XL_THIN)


def xw_write_add_leg_block(sheet, contracts):
    """Add Leg input cells + dropdowns -- live mirror of excel_sync.write_add_leg_block. Caller must have run xw_reset_dropdown_validations() once this refresh first."""

    sheet[ADD_LEG_LABEL_CELL].value = "ADD LEG"
    sheet[ADD_LEG_TYPE_LABEL_CELL].value = "Type:"
    sheet[ADD_LEG_CONTRACT_LABEL_CELL].value = "Contract:"
    sheet[ADD_LEG_OPTION_TYPE_LABEL_CELL].value = "Call/Put (ignored for Future):"
    sheet[ADD_LEG_QUANTITY_LABEL_CELL].value = "Quantity:"
    sheet[ADD_LEG_STRIKE_LABEL_CELL].value = "Strike (blank for Future):"
    sheet[ADD_LEG_EXPIRY_LABEL_CELL].value = "Expiry d (blank=auto):"
    sheet[ADD_LEG_VOL_LABEL_CELL].value = "Vol % (blank for Future):"

    type_cell = sheet[ADD_LEG_TYPE_VALUE_CELL]
    if type_cell.value not in ADD_LEG_TYPES:
        type_cell.value = ADD_LEG_TYPES[0]
    _add_list_validation(type_cell, ADD_LEG_TYPES)

    option_type_cell = sheet[ADD_LEG_OPTION_TYPE_VALUE_CELL]
    if option_type_cell.value not in ADD_LEG_OPTION_TYPES:
        option_type_cell.value = ADD_LEG_OPTION_TYPES[0]
    _add_list_validation(option_type_cell, ADD_LEG_OPTION_TYPES)

    if contracts:
        contract_cell = sheet[ADD_LEG_CONTRACT_VALUE_CELL]
        if contract_cell.value not in contracts:
            contract_cell.value = contracts[0]
        _add_list_validation(contract_cell, contracts)

    _xw_style_add_leg_block(sheet)


def xw_read_add_leg_inputs(sheet):
    return {
        "type": sheet[ADD_LEG_TYPE_VALUE_CELL].value,
        "contract": sheet[ADD_LEG_CONTRACT_VALUE_CELL].value,
        "option_type": sheet[ADD_LEG_OPTION_TYPE_VALUE_CELL].value,
        "quantity": sheet[ADD_LEG_QUANTITY_VALUE_CELL].value,
        "strike": sheet[ADD_LEG_STRIKE_VALUE_CELL].value,
        "expiry_days": sheet[ADD_LEG_EXPIRY_VALUE_CELL].value,
        "vol_pct": sheet[ADD_LEG_VOL_VALUE_CELL].value,
    }


def xw_clear_add_leg_inputs(sheet):
    for cell in (
        ADD_LEG_QUANTITY_VALUE_CELL, ADD_LEG_STRIKE_VALUE_CELL,
        ADD_LEG_EXPIRY_VALUE_CELL, ADD_LEG_VOL_VALUE_CELL,
    ):
        sheet[cell].value = None


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
