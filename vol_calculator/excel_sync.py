from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

from vol_calculator.instruments import Future, AmericanOption
from vol_calculator.pricing import AmericanOptionPricer
from vol_calculator.greeks import GreekEngine
from vol_calculator.structures import (
    call_spread, put_spread, straddle, strangle, butterfly, collar, calendar_spread,
)
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


_TABLE_LABEL_FILL = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
_TABLE_LABEL_FONT = Font(color="FFFFFF", bold=True, size=11)
_TABLE_HEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
_TABLE_HEADER_FONT = Font(bold=True, size=10)
_TABLE_BAND_FILL = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")
_TABLE_BORDER = Border(*(Side(style="thin", color="D9D9D9"),) * 4)


def write_dataframe(ws, df, start_row=1, start_col=1, label=None, include_index=False, index_label=""):
    """
    Writes a pandas DataFrame as a styled header + rows block (title bar,
    header row, banded/bordered data rows -- same visual language as the
    Positions table and Add Leg block), starting at (start_row,
    start_col). If label is given, it goes on its own row just above the
    header. With include_index=True, the DataFrame's index becomes the
    first column (for pivot-shaped outputs like spot_vol_surface/
    strike_map, where the index itself is meaningful -- e.g. spot move
    or strike). Returns the row just after the last one written, so
    callers can stack multiple tables on one sheet.
    """

    row = start_row
    col_offset = start_col + (1 if include_index else 0)
    n_cols = len(df.columns) + (col_offset - start_col)

    if label is not None:
        ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=start_col + n_cols - 1)
        title_cell = ws.cell(row=row, column=start_col, value=label)
        title_cell.fill = _TABLE_LABEL_FILL
        title_cell.font = _TABLE_LABEL_FONT
        title_cell.alignment = Alignment(horizontal="center")
        row += 1

    header_row = row
    if include_index:
        ws.cell(row=row, column=start_col, value=index_label)
    for col, header in enumerate(df.columns, start=col_offset):
        ws.cell(row=row, column=col, value=header)
    for col in range(start_col, start_col + n_cols):
        cell = ws.cell(row=header_row, column=col)
        cell.fill = _TABLE_HEADER_FILL
        cell.font = _TABLE_HEADER_FONT
        cell.border = _TABLE_BORDER
    row += 1

    for band_idx, (idx, data_row) in enumerate(df.iterrows()):
        if include_index:
            ws.cell(row=row, column=start_col, value=idx)
        for col, value in enumerate(data_row, start=col_offset):
            cell = ws.cell(row=row, column=col, value=round(value, 4) if isinstance(value, float) else value)
            if isinstance(value, float):
                cell.number_format = "#,##0.00"
        for col in range(start_col, start_col + n_cols):
            cell = ws.cell(row=row, column=col)
            cell.border = _TABLE_BORDER
            if band_idx % 2 == 1:
                cell.fill = _TABLE_BAND_FILL
        row += 1

    return row

POSITIONS_HEADERS = [
    "Contract", "Commodity", "Type", "Quantity",
    "Strike", "Expiry (d)", "Vol (%)", "Entry Price", "Underlying Price", "Mark to Market", "Source",
]

FUTURE_SPOT_LABEL_CELL = "A10"
FUTURE_SPOT_VALUE_CELL = "B10"

# The live CQGXL watchlist -- rows 3-8, 3 commodity blocks (Corn A-C,
# Soybean E-G, SRW Wheat I-J), headers on row 1. Pre-existing hand-built
# layout, not written by any function here -- styling only, and the
# symbol/price cells themselves (B/F/J, the live RTD-linked values) are
# never touched, only their number format.
_LIVE_QUOTES_BLOCKS = [(1, 3), (5, 7), (9, 11)]  # (first_col, last_col) -- each is Contract/Price/ATM IV
_LIVE_QUOTES_HEADER_ROW = 1
_LIVE_QUOTES_DATA_ROWS = (3, 8)
_LIVE_QUOTES_HEADER_FILL = PatternFill(start_color="BF8F00", end_color="BF8F00", fill_type="solid")
_LIVE_QUOTES_HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
_LIVE_QUOTES_BORDER = Border(*(Side(style="thin", color="BF8F00"),) * 4)


def style_live_quotes_watchlist(ws):

    for first_col, last_col in _LIVE_QUOTES_BLOCKS:
        for col in range(first_col, last_col + 1):
            header_cell = ws.cell(row=_LIVE_QUOTES_HEADER_ROW, column=col)
            header_cell.fill = _LIVE_QUOTES_HEADER_FILL
            header_cell.font = _LIVE_QUOTES_HEADER_FONT

        for row in range(_LIVE_QUOTES_DATA_ROWS[0], _LIVE_QUOTES_DATA_ROWS[1] + 1):
            for col in range(first_col, last_col + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = _LIVE_QUOTES_BORDER
                if col == first_col + 1:  # the price column in every block
                    cell.number_format = "#,##0.00"


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
    Interest Rate (%) cell, synced to live SOFR on every refresh when
    available -- falls back to whatever's already in the cell (or
    fallback_rate if it's empty) only when the SOFR fetch itself fails.

    sofr: the dict fetch_latest_sofr() returns, or None if that call
    failed -- falls back to fallback_rate and says so in the caption.
    """

    ws[label_cell] = "IR (%):"

    if sofr:
        ws[value_cell] = round(sofr["rate"] * 100, 2)
    elif ws[value_cell].value is None:
        ws[value_cell] = round(fallback_rate * 100, 2)

    if sofr:
        caption = ws[caption_cell]
        caption.value = f"SOFR -- last updated {sofr['effective_date']} (click for source)"
        caption.hyperlink = sofr["source_url"]
        caption.font = Font(color="0563C1", underline="single", size=10)
    else:
        ws[caption_cell] = "SOFR unavailable -- using manually entered rate."


def read_interest_rate(ws, fallback_rate=0.039, value_cell=INTEREST_RATE_VALUE_CELL):
    """Reads the Interest Rate (%) cell as a decimal, falling back if empty/invalid."""

    value = ws[value_cell].value
    try:
        return float(value) / 100
    except (TypeError, ValueError):
        return fallback_rate


# Column N onward -- outside columns 1-11 that write_positions_table's
# clear_range touches (Positions is 11 columns wide now that it has
# Entry Price/Underlying Price/Mark to Market, A-K), leaving L/M as a
# deliberate two-column visual gap between the Positions table and
# this block.
ADD_LEG_LABEL_CELL = "N13"
ADD_LEG_TYPE_LABEL_CELL, ADD_LEG_TYPE_VALUE_CELL = "N14", "O14"
ADD_LEG_CONTRACT_LABEL_CELL, ADD_LEG_CONTRACT_VALUE_CELL = "N15", "O15"
# Quantity 2 is Collar's call_qty only (Quantity itself doubles as
# Collar's put_qty) -- blank/unused for every other type.
ADD_LEG_QUANTITY_LABEL_CELL, ADD_LEG_QUANTITY_VALUE_CELL = "N16", "O16"
ADD_LEG_QUANTITY2_LABEL_CELL, ADD_LEG_QUANTITY2_VALUE_CELL = "N17", "O17"
# Strike 1/2/3 are generic -- which named parameter they map to depends
# on Type (e.g. Call Spread: Strike 1/2 = k_long/k_short; Butterfly:
# Strike 1/2/3 = k1/k2/k3). See build_manual_row.
ADD_LEG_STRIKE_LABEL_CELL, ADD_LEG_STRIKE_VALUE_CELL = "N18", "O18"
ADD_LEG_STRIKE2_LABEL_CELL, ADD_LEG_STRIKE2_VALUE_CELL = "N19", "O19"
ADD_LEG_STRIKE3_LABEL_CELL, ADD_LEG_STRIKE3_VALUE_CELL = "N20", "O20"
# Expiry 2 is Calendar Spread's far_expiry_days only (Expiry itself
# doubles as its near_expiry_days) -- blank/unused otherwise.
ADD_LEG_EXPIRY_LABEL_CELL, ADD_LEG_EXPIRY_VALUE_CELL = "N21", "O21"
ADD_LEG_EXPIRY2_LABEL_CELL, ADD_LEG_EXPIRY2_VALUE_CELL = "N22", "O22"
# Vol 1/2/3 -- same generic-slot idea as Strike 1/2/3.
ADD_LEG_VOL_LABEL_CELL, ADD_LEG_VOL_VALUE_CELL = "N23", "O23"
ADD_LEG_VOL2_LABEL_CELL, ADD_LEG_VOL2_VALUE_CELL = "N24", "O24"
ADD_LEG_VOL3_LABEL_CELL, ADD_LEG_VOL3_VALUE_CELL = "N25", "O25"
# Option Type -- only Butterfly and Calendar Spread need this (every
# other structure already fixes call vs put per leg structurally).
ADD_LEG_OPTION_TYPE_LABEL_CELL, ADD_LEG_OPTION_TYPE_VALUE_CELL = "N26", "O26"
# Entry Price -- a single Call/Put only (ignored for a Future and for
# every multi-leg structure, whose decomposed legs always fall back to
# their own Mark to Market instead -- there's no one typed-in price to
# split across several legs). Left blank, falls back to Mark to Market
# once computed, same fallback the Positions table itself uses for a
# Daily Prop row with no settlement.
ADD_LEG_ENTRY_PRICE_LABEL_CELL, ADD_LEG_ENTRY_PRICE_VALUE_CELL = "N27", "O27"
ADD_LEG_STATUS_CELL = "N28"

# Direct leg types -- Call/Put are themselves the type (no generic
# "American Option" + separate selector), matching what the Positions
# table's own Type column already shows; the 7 multi-leg structures
# match Streamlit's own LEG_BUILDERS, decomposed into individual
# Call/Put/Future Positions rows on write (see build_manual_row).
ADD_LEG_TYPES = [
    "Future", "Call", "Put", "Call Spread", "Put Spread", "Straddle",
    "Strangle", "Butterfly", "Collar", "Calendar Spread",
]

# Visually separates the Add Leg block from the plain, unstyled
# Positions table it sits below -- a colored header, bold labels and
# pale-yellow input cells (the standard Excel "this is an editable
# input" convention), boxed in a border.
_ADD_LEG_HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
_ADD_LEG_HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
_ADD_LEG_LABEL_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
_ADD_LEG_LABEL_FONT = Font(bold=True, size=10)
_ADD_LEG_INPUT_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
_ADD_LEG_STATUS_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
_ADD_LEG_STATUS_FONT = Font(italic=True, size=9, color="C00000")
_ADD_LEG_THIN_BORDER = Border(*(Side(style="thin"),) * 4)

_ADD_LEG_ROWS = [
    (ADD_LEG_TYPE_LABEL_CELL, ADD_LEG_TYPE_VALUE_CELL),
    (ADD_LEG_CONTRACT_LABEL_CELL, ADD_LEG_CONTRACT_VALUE_CELL),
    (ADD_LEG_QUANTITY_LABEL_CELL, ADD_LEG_QUANTITY_VALUE_CELL),
    (ADD_LEG_QUANTITY2_LABEL_CELL, ADD_LEG_QUANTITY2_VALUE_CELL),
    (ADD_LEG_STRIKE_LABEL_CELL, ADD_LEG_STRIKE_VALUE_CELL),
    (ADD_LEG_STRIKE2_LABEL_CELL, ADD_LEG_STRIKE2_VALUE_CELL),
    (ADD_LEG_STRIKE3_LABEL_CELL, ADD_LEG_STRIKE3_VALUE_CELL),
    (ADD_LEG_EXPIRY_LABEL_CELL, ADD_LEG_EXPIRY_VALUE_CELL),
    (ADD_LEG_EXPIRY2_LABEL_CELL, ADD_LEG_EXPIRY2_VALUE_CELL),
    (ADD_LEG_VOL_LABEL_CELL, ADD_LEG_VOL_VALUE_CELL),
    (ADD_LEG_VOL2_LABEL_CELL, ADD_LEG_VOL2_VALUE_CELL),
    (ADD_LEG_VOL3_LABEL_CELL, ADD_LEG_VOL3_VALUE_CELL),
    (ADD_LEG_OPTION_TYPE_LABEL_CELL, ADD_LEG_OPTION_TYPE_VALUE_CELL),
    (ADD_LEG_ENTRY_PRICE_LABEL_CELL, ADD_LEG_ENTRY_PRICE_VALUE_CELL),
]


def _style_add_leg_block(ws):

    # K/L/M are now a deliberate empty gap between the Positions table
    # and this block -- reset them back to a normal width in case an
    # older refresh widened different columns when the block used to
    # sit elsewhere.
    ws.column_dimensions["J"].width = 8.43
    ws.column_dimensions["K"].width = 8.43
    ws.column_dimensions["L"].width = 8.43
    ws.column_dimensions["M"].width = 8.43
    ws.column_dimensions["N"].width = 28
    ws.column_dimensions["O"].width = 16

    ws.merge_cells(f"{ADD_LEG_LABEL_CELL}:O13")
    ws[ADD_LEG_LABEL_CELL].fill = _ADD_LEG_HEADER_FILL
    ws[ADD_LEG_LABEL_CELL].font = _ADD_LEG_HEADER_FONT
    ws[ADD_LEG_LABEL_CELL].alignment = Alignment(horizontal="center")

    for label_cell, value_cell in _ADD_LEG_ROWS:
        ws[label_cell].fill = _ADD_LEG_LABEL_FILL
        ws[label_cell].font = _ADD_LEG_LABEL_FONT
        ws[label_cell].border = _ADD_LEG_THIN_BORDER
        ws[value_cell].fill = _ADD_LEG_INPUT_FILL
        ws[value_cell].border = _ADD_LEG_THIN_BORDER

    ws.merge_cells(f"{ADD_LEG_STATUS_CELL}:O28")
    ws[ADD_LEG_STATUS_CELL].fill = _ADD_LEG_STATUS_FILL
    ws[ADD_LEG_STATUS_CELL].font = _ADD_LEG_STATUS_FONT
    ws[ADD_LEG_STATUS_CELL].alignment = Alignment(wrap_text=True, vertical="center")
    ws[ADD_LEG_STATUS_CELL].border = _ADD_LEG_THIN_BORDER


def write_add_leg_block(ws, contracts):
    """Add Leg input cells + dropdowns -- Future, Call, Put, and the 7 multi-leg structures, matching Streamlit's LEG_BUILDERS."""

    # Only the title (row 13) and status (row 28) cells are ever
    # merged -- a merged cell left over from an earlier column position
    # (this block has moved several times across this project's
    # history) can overlap a newly-merged range and strand a written
    # value in the OLD merge's anchor cell instead of the new one
    # (confirmed live on the xlwings path). Unmerging/clearing just
    # those two rows across every column this block has ever occupied,
    # before writing anything, guarantees a clean slate every refresh
    # -- WITHOUT touching rows 14-27, which are plain (never-merged)
    # input cells: wiping those too would destroy whatever the user
    # just typed on every single refresh, including mid-way through an
    # Add Leg submission that then hits a transient error, turning a
    # retry into a confusing false validation failure instead of a
    # clean retry.
    debris_cols = range(10, 16)  # J-O
    for merged_range in list(ws.merged_cells.ranges):
        if merged_range.min_row in (13, 28) and merged_range.min_col <= 15 and merged_range.max_col >= 10:
            ws.unmerge_cells(str(merged_range))
    for row in (13, 28):
        for col in debris_cols:
            ws.cell(row=row, column=col, value=None)

    ws[ADD_LEG_LABEL_CELL] = "ADD LEG"
    ws[ADD_LEG_TYPE_LABEL_CELL] = "Type:"
    ws[ADD_LEG_CONTRACT_LABEL_CELL] = "Contract (underlying):"
    ws[ADD_LEG_QUANTITY_LABEL_CELL] = "Quantity (Collar: Put Qty):"
    ws[ADD_LEG_QUANTITY2_LABEL_CELL] = "Quantity 2 (Collar: Call Qty):"
    ws[ADD_LEG_STRIKE_LABEL_CELL] = "Strike 1 (k1/k_long/k_put):"
    ws[ADD_LEG_STRIKE2_LABEL_CELL] = "Strike 2 (k2/k_short/k_call):"
    ws[ADD_LEG_STRIKE3_LABEL_CELL] = "Strike 3 (Butterfly k3 only):"
    ws[ADD_LEG_EXPIRY_LABEL_CELL] = "Expiry d (blank=auto; Calendar: near):"
    ws[ADD_LEG_EXPIRY2_LABEL_CELL] = "Expiry d 2 (Calendar Spread far only):"
    ws[ADD_LEG_VOL_LABEL_CELL] = "Vol % 1 (vol/vol_long/vol_call/vol_put/vol1/vol_near):"
    ws[ADD_LEG_VOL2_LABEL_CELL] = "Vol % 2 (vol_short/vol_put/vol_call/vol2/vol_far):"
    ws[ADD_LEG_VOL3_LABEL_CELL] = "Vol % 3 (Butterfly vol3 only):"
    ws[ADD_LEG_OPTION_TYPE_LABEL_CELL] = "Option Type (Butterfly/Calendar Spread only):"
    ws[ADD_LEG_ENTRY_PRICE_LABEL_CELL] = "Entry Price (blank = Mark to Mkt; single Call/Put only):"

    if ws[ADD_LEG_TYPE_VALUE_CELL].value not in ADD_LEG_TYPES:
        ws[ADD_LEG_TYPE_VALUE_CELL] = ADD_LEG_TYPES[0]

    type_dv = DataValidation(type="list", formula1=f'"{",".join(ADD_LEG_TYPES)}"', allow_blank=False)
    ws.add_data_validation(type_dv)
    type_dv.add(ws[ADD_LEG_TYPE_VALUE_CELL])

    option_type_dv = DataValidation(type="list", formula1='"call,put"', allow_blank=True)
    ws.add_data_validation(option_type_dv)
    option_type_dv.add(ws[ADD_LEG_OPTION_TYPE_VALUE_CELL])

    if contracts:
        if ws[ADD_LEG_CONTRACT_VALUE_CELL].value not in contracts:
            ws[ADD_LEG_CONTRACT_VALUE_CELL] = contracts[0]

        contract_dv = DataValidation(type="list", formula1=f'"{",".join(contracts)}"', allow_blank=False)
        ws.add_data_validation(contract_dv)
        contract_dv.add(ws[ADD_LEG_CONTRACT_VALUE_CELL])

    _style_add_leg_block(ws)


def read_add_leg_inputs(ws):
    return {
        "type": ws[ADD_LEG_TYPE_VALUE_CELL].value,
        "contract": ws[ADD_LEG_CONTRACT_VALUE_CELL].value,
        "quantity": ws[ADD_LEG_QUANTITY_VALUE_CELL].value,
        "quantity2": ws[ADD_LEG_QUANTITY2_VALUE_CELL].value,
        "strike": ws[ADD_LEG_STRIKE_VALUE_CELL].value,
        "strike2": ws[ADD_LEG_STRIKE2_VALUE_CELL].value,
        "strike3": ws[ADD_LEG_STRIKE3_VALUE_CELL].value,
        "expiry_days": ws[ADD_LEG_EXPIRY_VALUE_CELL].value,
        "expiry_days2": ws[ADD_LEG_EXPIRY2_VALUE_CELL].value,
        "vol_pct": ws[ADD_LEG_VOL_VALUE_CELL].value,
        "vol_pct2": ws[ADD_LEG_VOL2_VALUE_CELL].value,
        "vol_pct3": ws[ADD_LEG_VOL3_VALUE_CELL].value,
        "option_type": ws[ADD_LEG_OPTION_TYPE_VALUE_CELL].value,
        "entry_price": ws[ADD_LEG_ENTRY_PRICE_VALUE_CELL].value,
    }


def clear_add_leg_inputs(ws):
    for cell in (
        ADD_LEG_QUANTITY_VALUE_CELL, ADD_LEG_QUANTITY2_VALUE_CELL,
        ADD_LEG_STRIKE_VALUE_CELL, ADD_LEG_STRIKE2_VALUE_CELL, ADD_LEG_STRIKE3_VALUE_CELL,
        ADD_LEG_EXPIRY_VALUE_CELL, ADD_LEG_EXPIRY2_VALUE_CELL,
        ADD_LEG_VOL_VALUE_CELL, ADD_LEG_VOL2_VALUE_CELL, ADD_LEG_VOL3_VALUE_CELL,
        ADD_LEG_OPTION_TYPE_VALUE_CELL, ADD_LEG_ENTRY_PRICE_VALUE_CELL,
    ):
        ws[cell] = None


def build_manual_row(inputs, commodity_lookup, auto_expiry_fn):
    """
    Validates Add Leg's raw input dict and turns it into a LIST of
    POSITIONS_HEADERS-shaped {header: value} rows, Source="Manual" --
    one row for Future/Call/Put, 2+ rows for any multi-leg structure
    (same as Streamlit: a Straddle shows up in the Positions table as
    its own Call row + Put row, not one "Straddle" row -- there's no
    separate structure-level row or column). Raises ValueError with a
    human-readable message on anything missing/invalid -- the caller is
    expected to surface that directly (e.g. write it into a status
    cell), not crash the whole refresh.

    commodity_lookup: contract -> commodity, e.g. app.py's
    _commodity_from_contract. auto_expiry_fn: contract -> approx expiry
    days, e.g. vol_calculator.contracts.approx_option_expiry_days --
    used only when expiry_days is left blank for an option/structure.
    """

    leg_type = inputs.get("type")
    contract = inputs.get("contract")

    if leg_type not in ADD_LEG_TYPES:
        raise ValueError(f"Pick a Type ({'/'.join(ADD_LEG_TYPES)}).")
    if not contract:
        raise ValueError("Pick a Contract.")

    commodity = commodity_lookup(contract)

    def _required(key, label):
        value = inputs.get(key)
        if value in (None, ""):
            raise ValueError(f"{label} is required for a {leg_type}.")
        return float(value)

    def _leg_rows(*instruments):
        meta = {"commodity": commodity, "contract": contract, "vol_pct": None, "settlement_price": None}
        return [dict(zip(POSITIONS_HEADERS, _position_row(leg, meta, source="Manual"))) for leg in instruments]

    quantity = _required("quantity", "Quantity")

    if leg_type == "Future":
        return _leg_rows(Future(quantity=quantity, contract=contract))

    if leg_type in ("Call", "Put"):
        strike = _required("strike", "Strike 1")
        vol_pct = _required("vol_pct", "Vol % 1")

        expiry_days = inputs.get("expiry_days")
        if expiry_days in (None, ""):
            expiry_days = auto_expiry_fn(contract)
            if expiry_days is None:
                raise ValueError("Couldn't auto-compute Expiry -- enter it manually.")

        instrument = AmericanOption(
            quantity=quantity, strike=strike, expiry_days=int(expiry_days),
            option_type="call" if leg_type == "Call" else "put",
            volatility=vol_pct / 100, contract=contract,
        )
        rows = _leg_rows(instrument)

        # Entry Price only applies to a single Call/Put -- left as None
        # when not typed in, the caller fills it in with this leg's own
        # Mark to Market once that's computed (same fallback the
        # Positions table uses for any option with no recorded Entry
        # Price), rather than defaulting to 0 here.
        entry_price = inputs.get("entry_price")
        if entry_price not in (None, ""):
            rows[0]["Entry Price"] = float(entry_price)
        return rows

    # A multi-leg structure -- Strike 1/Vol % 1 and Expiry (d) are
    # shared entry points reused across every structure type; which
    # named parameter they map to depends on which one is selected.
    expiry_days = inputs.get("expiry_days")
    if expiry_days in (None, ""):
        expiry_days = auto_expiry_fn(contract)
        if expiry_days is None:
            raise ValueError("Couldn't auto-compute Expiry -- enter it manually.")
    expiry_days = int(expiry_days)

    strike = _required("strike", "Strike 1")
    vol_pct = _required("vol_pct", "Vol % 1")

    if leg_type in ("Call Spread", "Put Spread"):
        strike2 = _required("strike2", "Strike 2")
        vol_pct2 = _required("vol_pct2", "Vol % 2")
        builder = call_spread if leg_type == "Call Spread" else put_spread
        structure = builder(
            quantity=quantity, k_long=strike, k_short=strike2, expiry_days=expiry_days,
            vol_long=vol_pct / 100, vol_short=vol_pct2 / 100, contract=contract,
        )

    elif leg_type == "Straddle":
        vol_pct2 = _required("vol_pct2", "Vol % 2")
        structure = straddle(
            quantity=quantity, strike=strike, expiry_days=expiry_days,
            vol_call=vol_pct / 100, vol_put=vol_pct2 / 100, contract=contract,
        )

    elif leg_type == "Strangle":
        strike2 = _required("strike2", "Strike 2")
        vol_pct2 = _required("vol_pct2", "Vol % 2")
        structure = strangle(
            quantity=quantity, k_put=strike, k_call=strike2, expiry_days=expiry_days,
            vol_put=vol_pct / 100, vol_call=vol_pct2 / 100, contract=contract,
        )

    elif leg_type == "Butterfly":
        strike2 = _required("strike2", "Strike 2")
        strike3 = _required("strike3", "Strike 3")
        vol_pct2 = _required("vol_pct2", "Vol % 2")
        vol_pct3 = _required("vol_pct3", "Vol % 3")
        option_type = inputs.get("option_type")
        if option_type not in ("call", "put"):
            raise ValueError("Option Type (call/put) is required for a Butterfly.")
        structure = butterfly(
            quantity=quantity, k1=strike, k2=strike2, k3=strike3, expiry_days=expiry_days,
            option_type=option_type, vol1=vol_pct / 100, vol2=vol_pct2 / 100,
            vol3=vol_pct3 / 100, contract=contract,
        )

    elif leg_type == "Collar":
        quantity2 = _required("quantity2", "Quantity 2 (Call Qty)")
        strike2 = _required("strike2", "Strike 2")
        vol_pct2 = _required("vol_pct2", "Vol % 2")
        structure = collar(
            put_qty=quantity, call_qty=quantity2, k_put=strike, k_call=strike2,
            expiry_days=expiry_days, vol_put=vol_pct / 100, vol_call=vol_pct2 / 100,
            contract=contract,
        )

    elif leg_type == "Calendar Spread":
        expiry_days2 = inputs.get("expiry_days2")
        if expiry_days2 in (None, ""):
            raise ValueError("Expiry d 2 (far) is required for a Calendar Spread.")
        vol_pct2 = _required("vol_pct2", "Vol % 2")
        option_type = inputs.get("option_type")
        if option_type not in ("call", "put"):
            raise ValueError("Option Type (call/put) is required for a Calendar Spread.")
        structure = calendar_spread(
            quantity=quantity, strike=strike, near_expiry_days=expiry_days,
            far_expiry_days=int(expiry_days2), option_type=option_type,
            vol_near=vol_pct / 100, vol_far=vol_pct2 / 100, contract=contract,
        )

    else:
        raise ValueError(f"Unsupported Type: {leg_type!r}")

    return _leg_rows(*structure.legs)


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

    Only a Future leg's settlement_price is the underlying contract's own
    price -- an Option leg's settlement_price is that option's own
    premium, a completely different number, and must never overwrite the
    underlying price (same bug/fix as app.py's Bushel import handler).
    """

    prices = {}
    for instrument, meta in imported:
        if not isinstance(instrument, Future):
            continue
        price = meta.get("settlement_price")
        if price is not None:
            prices[meta["contract"]] = price
    return prices


def _position_row(instrument, meta, source):
    """
    Column routing matches the Streamlit Positions table exactly: a
    Future has no Entry Price or Mark to Market of its own -- its price
    lives under Underlying Price instead. An Option's Entry Price falls
    back to its Mark to Market when nothing was recorded for it (a
    blank Daily Prop settlement, or a Manual leg added before an Entry
    Price existed), rather than showing a blank cell.
    """

    if isinstance(instrument, Future):
        return [
            meta["contract"], meta["commodity"], "Future", instrument.quantity,
            None, None, meta.get("vol_pct"), None,
            meta.get("settlement_price"), None, source,
        ]

    if isinstance(instrument, AmericanOption):
        entry_price = meta.get("settlement_price")
        if entry_price is None:
            entry_price = meta.get("model_price")

        return [
            meta["contract"], meta["commodity"],
            "Call" if instrument.option_type == "call" else "Put",
            instrument.quantity, instrument.strike, instrument.expiry_days,
            round(instrument.volatility * 100, 2), entry_price,
            meta.get("underlying_price"), meta.get("model_price"), source,
        ]

    raise TypeError(f"Unsupported instrument type: {type(instrument)}")


def compute_model_price(instrument, meta, market):
    """
    The leg's own theoretical price, computed fresh from its own
    parameters using the same Black-76-based American-exercise pricer
    (AmericanOptionPricer, Barone-Adesi-Whaley on a Black process) that
    Greeks already use everywhere else -- not copied from Settlement
    Price, which can be a frozen CSV-reported value.

    For a Future, this is just its contract's current engine price
    (market.price_for) -- which can differ from a frozen Settlement
    Price if the market has moved since the CSV was generated. For a
    Call/Put, it's the actual computed premium from strike/expiry/vol/
    rate, which can likewise differ from a broker-reported settlement.

    Returns None if the contract has no price in market, or if pricing
    otherwise fails (e.g. a manually-added option with not-yet-valid
    inputs) -- callers should treat that as "can't be shown yet", not
    crash the whole refresh over it.
    """

    contract = meta.get("contract")
    try:
        underlying_price = market.price_for(contract)
    except (KeyError, ValueError):
        return None

    if isinstance(instrument, Future):
        return round(underlying_price, 4)

    if isinstance(instrument, AmericanOption):
        try:
            return round(
                AmericanOptionPricer(
                    futures_price=underlying_price,
                    strike=instrument.strike,
                    rate=market.interest_rate,
                    volatility=instrument.volatility,
                    expiry_days=instrument.expiry_days,
                    option_type=instrument.option_type,
                ).price(),
                4,
            )
        except Exception:
            return None

    return None


def compute_underlying_price(contract, market):
    """An Option's Underlying Price -- the live price of the future it's actually written on, recomputed every refresh. Returns None if the contract has no price in market."""

    try:
        return round(market.price_for(contract), 4)
    except (KeyError, ValueError):
        return None


def _row_to_instrument(row):
    """
    Inverse of _position_row -- turns a {header: value} dict (as read back
    by read_positions_rows) into an (instrument, meta) pair, same shape
    import_bushel_positions produces. Used to fold Manual rows back into
    the Portfolio the engine computes against, not just display them.
    """

    contract = row.get("Contract")
    commodity = row.get("Commodity")
    leg_type = row.get("Type")
    quantity = float(row.get("Quantity") or 0)

    if leg_type == "Future":
        instrument = Future(quantity=quantity, contract=contract)
        # A Future's own price lives under Underlying Price (it has no
        # separate Entry Price column) -- same routing _position_row uses.
        settlement_price = row.get("Underlying Price")
    elif leg_type in ("Call", "Put"):
        instrument = AmericanOption(
            quantity=quantity,
            strike=float(row["Strike"]),
            expiry_days=int(row["Expiry (d)"]),
            option_type="call" if leg_type == "Call" else "put",
            volatility=float(row.get("Vol (%)") or 0) / 100,
            contract=contract,
        )
        settlement_price = row.get("Entry Price")
    else:
        raise ValueError(f"Unknown Type in Positions row: {leg_type!r}")

    meta = {
        "commodity": commodity, "contract": contract,
        "vol_pct": None, "settlement_price": settlement_price,
    }
    return instrument, meta


def read_positions_rows(ws, start_row=13):
    """
    Reads the Positions table back as a list of {header: value} dicts --
    stops at the first row with no Contract value. Returns [] if the
    table doesn't look like POSITIONS_HEADERS yet (nothing written, or
    unexpected layout).
    """

    header_row = start_row + 1
    headers = [ws.cell(row=header_row, column=c).value for c in range(1, len(POSITIONS_HEADERS) + 1)]

    if headers != POSITIONS_HEADERS:
        return []

    rows = []
    r = header_row + 1

    while ws.cell(row=r, column=1).value is not None:
        rows.append({h: ws.cell(row=r, column=c + 1).value for c, h in enumerate(POSITIONS_HEADERS)})
        r += 1

    return rows


_POSITIONS_TITLE_FILL = PatternFill(start_color="375623", end_color="375623", fill_type="solid")
_POSITIONS_TITLE_FONT = Font(color="FFFFFF", bold=True, size=11)
_POSITIONS_HEADER_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
_POSITIONS_HEADER_FONT = Font(bold=True, size=10)
_POSITIONS_BAND_FILL = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")
_POSITIONS_THIN_BORDER = Border(*(Side(style="thin", color="D9D9D9"),) * 4)

# Column -> number_format, by POSITIONS_HEADERS position (1-indexed).
_POSITIONS_NUMBER_FORMATS = {
    4: "#,##0.##",       # Quantity
    5: "#,##0.00",       # Strike
    6: "0",              # Expiry (d)
    7: "0.00",           # Vol (%)
    8: "#,##0.00",       # Entry Price
    9: "#,##0.00",       # Underlying Price
    10: "#,##0.00",      # Mark to Market
}


def _style_positions_table(ws, start_row, header_row, last_data_row, n_cols):
    """Visually separates the Positions table from the plain, unstyled sheet around it -- same treatment as the Add Leg block, its own color scheme."""

    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=n_cols)
    title_cell = ws.cell(row=start_row, column=1)
    title_cell.fill = _POSITIONS_TITLE_FILL
    title_cell.font = _POSITIONS_TITLE_FONT
    title_cell.alignment = Alignment(horizontal="center")

    for col in range(1, n_cols + 1):
        header_cell = ws.cell(row=header_row, column=col)
        header_cell.fill = _POSITIONS_HEADER_FILL
        header_cell.font = _POSITIONS_HEADER_FONT
        header_cell.border = _POSITIONS_THIN_BORDER

    for row in range(header_row + 1, last_data_row + 1):
        banded = (row - header_row) % 2 == 0
        for col in range(1, n_cols + 1):
            cell = ws.cell(row=row, column=col)
            cell.border = _POSITIONS_THIN_BORDER
            if banded:
                cell.fill = _POSITIONS_BAND_FILL
            if col in _POSITIONS_NUMBER_FORMATS and isinstance(cell.value, (int, float)):
                cell.number_format = _POSITIONS_NUMBER_FORMATS[col]

    # A shrinking table (fewer legs today than last time this styling
    # ran) would otherwise leave stale banding/borders below the real
    # data -- clear a reasonable buffer of leftover formatting too.
    for row in range(last_data_row + 1, last_data_row + 51):
        for col in range(1, n_cols + 1):
            cell = ws.cell(row=row, column=col)
            cell.fill = PatternFill(fill_type=None)
            cell.border = Border()


def write_positions_table(ws, imported, extra_rows=None, start_row=13, clear_rows=500):
    """
    Writes (instrument, meta) pairs -- the same shape import_bushel_positions
    returns -- as "Daily Prop" rows, starting at start_row (a "POSITIONS"
    label, then headers, then one row per leg). extra_rows (already-built
    {header: value} dicts, typically preserved/new "Manual" rows from the
    caller) are appended verbatim afterward, unchanged. Clears a generous
    range below start_row first, so a shrinking book doesn't leave stale
    rows behind from a previous, longer day's data -- this is why
    extra_rows has to be read by the caller *before* calling this, not
    re-derived from the sheet in here.
    """

    extra_rows = extra_rows or []
    n_cols = len(POSITIONS_HEADERS)

    clear_range(ws, min_row=start_row, max_row=start_row + clear_rows, min_col=1, max_col=n_cols)

    ws.cell(row=start_row, column=1, value="POSITIONS")

    header_row = start_row + 1
    for col, header in enumerate(POSITIONS_HEADERS, start=1):
        ws.cell(row=header_row, column=col, value=header)

    row_idx = header_row + 1

    for instrument, meta in imported:
        row_values = _position_row(instrument, meta, source="Daily Prop")
        for col, value in enumerate(row_values, start=1):
            ws.cell(row=row_idx, column=col, value=value)
        row_idx += 1

    for extra in extra_rows:
        for col, header in enumerate(POSITIONS_HEADERS, start=1):
            ws.cell(row=row_idx, column=col, value=extra.get(header))
        row_idx += 1

    _style_positions_table(ws, start_row, header_row, row_idx - 1, n_cols)

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
        cell = ws.cell(row=1, column=col, value=value)
        cell.fill = _TABLE_HEADER_FILL
        cell.font = _TABLE_HEADER_FONT
        cell.border = _TABLE_BORDER

    for i, contract in enumerate(contracts):
        dollars = GreekEngine.report_dollars(portfolio, market, contract)
        row_values = [contract] + [_round_greek(k, v) for k, v in dollars.items()]
        for col, value in enumerate(row_values, start=1):
            cell = ws.cell(row=2 + i, column=col, value=value)
            cell.border = _TABLE_BORDER
            if i % 2 == 1:
                cell.fill = _TABLE_BAND_FILL
            if isinstance(value, float):
                cell.number_format = "#,##0.00"

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
