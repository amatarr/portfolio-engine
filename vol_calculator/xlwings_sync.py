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

_XL_VALIDATE_LIST = 3


def xw_write_future_spot_selector(sheet, contracts):

    sheet[FUTURE_SPOT_LABEL_CELL].value = "Future Spot:"

    if not contracts:
        return

    value_cell = sheet[FUTURE_SPOT_VALUE_CELL]

    if value_cell.value not in contracts:
        value_cell.value = contracts[0]

    try:
        value_cell.api.Validation.Delete()
    except Exception:
        pass

    value_cell.api.Validation.Add(Type=_XL_VALIDATE_LIST, Formula1=",".join(contracts))


def xw_read_future_spot(sheet, contracts):

    value = sheet[FUTURE_SPOT_VALUE_CELL].value
    return value if value in contracts else (contracts[0] if contracts else None)


def xw_write_interest_rate_cell(sheet, sofr=None, fallback_rate=0.039):

    sheet[INTEREST_RATE_LABEL_CELL].value = "Interest Rate (%):"

    if sofr:
        sheet[INTEREST_RATE_VALUE_CELL].value = round(sofr["rate"] * 100, 2)
    elif sheet[INTEREST_RATE_VALUE_CELL].value is None:
        sheet[INTEREST_RATE_VALUE_CELL].value = round(fallback_rate * 100, 2)

    if sofr:
        sheet[SOFR_CAPTION_CELL].value = (
            f"SOFR: {sofr['rate'] * 100:.2f}% (updated {sofr['effective_date']}) -- {sofr['source_url']}"
        )
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
