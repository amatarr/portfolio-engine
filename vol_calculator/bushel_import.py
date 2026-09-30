from datetime import date, datetime

import pandas as pd

from vol_calculator.instruments import Future, AmericanOption
from vol_calculator.contracts import contract_code

# Position of each field in the vendor export, by Excel column letter --
# mapping is positional (not by header name), since header text is not
# reliable across exports. Update this dict if the template changes.
_COLUMN_LETTERS = {
    "instrument": "L",
    "contract_type": "M",
    "strike": "N",
    "contract_year": "O",
    "contract_month": "P",
    "settlement_price": "V",
    "quantity": "U",
    "vol_pct": "X",
    "maturity": "FC",
}

_INSTRUMENT_TO_COMMODITY = {"C": "Corn", "S": "Soybean", "W": "SRW Wheat"}

_MONTH_NUMBER_TO_CODE = {
    1: "F", 2: "G", 3: "H", 4: "J", 5: "K", 6: "M",
    7: "N", 8: "Q", 9: "U", 10: "V", 11: "X", 12: "Z",
}

_VALID_MONTH_CODES = set(_MONTH_NUMBER_TO_CODE.values())


def _col_index(letters):
    """Excel column letters -> 0-indexed position, e.g. 'A' -> 0, 'FC' -> 158."""

    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n - 1


_COLUMNS = {name: _col_index(letters) for name, letters in _COLUMN_LETTERS.items()}


def _month_code(value):

    s = str(value).strip().upper()

    if s in _VALID_MONTH_CODES:
        return s

    try:
        return _MONTH_NUMBER_TO_CODE[int(float(s))]
    except (ValueError, KeyError):
        raise ValueError(f"Unrecognized contract month: {value!r}")


def _year_code(value):

    y = int(float(value))
    return y if y > 100 else 2000 + y


def import_bushel_positions(path):
    """
    Reads a Bushel-style export (CSV) positionally by Excel column letter,
    keeping only Corn/Soybean/SRW Wheat rows (Instrument in C/S/W).

    Each instrument carries its own .contract (e.g. "ZCH26"). Also reads the
    Settlement Price (column V) for each row, so the caller can populate
    Market.futures_prices for any newly-seen contract.

    Returns a list of (instrument, meta) tuples. meta is a plain dict --
    {"commodity", "contract", "vol_pct", "settlement_price"} -- vol_pct/
    settlement_price are display/price-seeding helpers, since Future doesn't
    store volatility and prices live on Market, not on the instrument.
    """

    raw = pd.read_csv(path, header=0)
    today = date.today()

    results = []

    for _, row in raw.iterrows():

        instrument_code = str(row.iloc[_COLUMNS["instrument"]]).strip().upper()

        if instrument_code not in _INSTRUMENT_TO_COMMODITY:
            continue

        commodity = _INSTRUMENT_TO_COMMODITY[instrument_code]
        contract_type = str(row.iloc[_COLUMNS["contract_type"]]).strip().upper()

        year = _year_code(row.iloc[_COLUMNS["contract_year"]])
        month = _month_code(row.iloc[_COLUMNS["contract_month"]])
        contract = contract_code(commodity, month, year)

        quantity = float(row.iloc[_COLUMNS["quantity"]])

        vol_raw = row.iloc[_COLUMNS["vol_pct"]]
        vol_pct = float(vol_raw) if pd.notna(vol_raw) else None

        price_raw = row.iloc[_COLUMNS["settlement_price"]]
        settlement_price = float(price_raw) if pd.notna(price_raw) else None

        if contract_type == "F":
            instrument = Future(quantity=quantity, contract=contract)
            meta = {
                "commodity": commodity, "contract": contract,
                "vol_pct": vol_pct, "settlement_price": settlement_price,
            }

        elif contract_type in ("C", "P"):

            strike = float(row.iloc[_COLUMNS["strike"]])

            maturity_raw = str(row.iloc[_COLUMNS["maturity"]]).strip().split(".")[0]
            maturity = datetime.strptime(maturity_raw, "%Y%m%d").date()
            expiry_days = max((maturity - today).days, 1)

            volatility = (vol_pct or 0.0) / 100

            instrument = AmericanOption(
                quantity=quantity,
                strike=strike,
                expiry_days=expiry_days,
                option_type="call" if contract_type == "C" else "put",
                volatility=volatility,
                contract=contract,
            )
            meta = {
                "commodity": commodity, "contract": contract,
                "vol_pct": None, "settlement_price": settlement_price,
            }

        else:
            continue

        results.append((instrument, meta))

    return results
