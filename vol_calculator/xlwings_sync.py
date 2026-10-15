"""
Live (xlwings/COM) access to the open Excel workbook -- changes appear
instantly, no file lock, no "close Excel first."

The Home sheet is driven entirely through the hidden _Data sheet and
named input cells (see below): Python only writes values there and reads
inputs by name, never touching visible layout or formatting. The analysis
sheets work the same way through _Calc -- see xlwings_analysis.py.
"""

from vol_calculator.excel_sync import POSITIONS_HEADERS
from vol_calculator.cqg_quotes import _convert_cqg_symbol

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
    """Reads the live CQGXL watchlist -- [{"commodity", "contract", "price", "cell"}, ...], where cell is the sheet-qualified address of the live price, for formulas that should tick with it. Skips a cell showing text (e.g. "Disconnected from server.") instead of a number, and a commodity whose block isn't named in this workbook."""

    quotes = []

    for name, commodity in WATCHLIST_NAMES.items():
        try:
            rng = wb.names[name].refers_to_range
        except Exception:
            continue

        for i, (symbol, price, *_) in enumerate(rng.options(ndim=2).value):
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
                "cell": f"'{rng.sheet.name}'!{rng[i, 1].address}",
            })

    return quotes




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
    "SpotChoices": ("M", 1),  # analysis sheets' spot dropdown: "Follow Home" + book contracts
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

