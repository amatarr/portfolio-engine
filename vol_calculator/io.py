import os

import pandas as pd

from vol_calculator.instruments import Portfolio, Future, AmericanOption, Structure

_FIELDS = [
    "structure_id",
    "structure_type",
    "instrument_type",
    "quantity",
    "strike",
    "expiry_days",
    "option_type",
    "volatility",
    "contract",
    "option_month",
    "entry_price",
    "underlying_price",
]


def portfolio_to_csv(portfolio, path, position_meta=None):
    """
    One row per leg. Legs belonging to the same Structure (call spread,
    straddle, etc.) share a structure_id so portfolio_from_csv can regroup
    them back into a Structure on reload, instead of a bag of anonymous
    legs. Bare (non-Structure) legs get an empty structure_id. Position
    prices are optional so existing callers can continue exporting only
    instrument data.
    """

    rows = []
    structure_id = 0

    for position_index, pos in enumerate(portfolio.positions):
        meta = position_meta[position_index] if position_meta is not None else {}
        if isinstance(pos, Future):
            entry_price = meta.get("entry_price", meta.get("settlement_price"))
            underlying_price = meta.get("settlement_price")
        else:
            entry_price = meta.get("settlement_price")
            underlying_price = None

        if isinstance(pos, Structure):

            for leg in pos.legs:
                row = _leg_to_row(leg, structure_id, pos.name or "")
                row["entry_price"] = entry_price
                row["underlying_price"] = underlying_price
                rows.append(row)

            structure_id += 1

        else:
            row = _leg_to_row(pos, "", "")
            row["entry_price"] = entry_price
            row["underlying_price"] = underlying_price
            rows.append(row)

    pd.DataFrame(rows, columns=_FIELDS).to_csv(path, index=False)


def _leg_to_row(leg, structure_id, structure_type):

    if isinstance(leg, Future):
        return {
            "structure_id": structure_id,
            "structure_type": structure_type,
            "instrument_type": "Future",
            "quantity": leg.quantity,
            "strike": "",
            "expiry_days": "",
            "option_type": "",
            "volatility": "",
            "contract": leg.contract or "",
            "option_month": "",
        }

    if isinstance(leg, AmericanOption):
        return {
            "structure_id": structure_id,
            "structure_type": structure_type,
            "instrument_type": "AmericanOption",
            "quantity": leg.quantity,
            "strike": leg.strike,
            "expiry_days": leg.expiry_days,
            "option_type": leg.option_type,
            "volatility": leg.volatility,
            "contract": leg.contract or "",
            "option_month": leg.option_month or "",
        }

    raise TypeError(f"Unsupported leg type for CSV export: {type(leg)}")


def portfolio_from_csv(path, return_meta=False):
    """
    Inverse of portfolio_to_csv. Rows sharing a structure_id are regrouped
    into one Structure, added to the portfolio at the position of that
    structure's first row, so the original leg ordering (including any
    interleaving of bare legs and structures) round-trips exactly. When
    return_meta is true, also returns saved position prices aligned with
    portfolio.positions.
    """

    df = pd.read_csv(path, dtype=str, keep_default_na=False)

    portfolio = Portfolio()
    open_structures = {}
    position_meta = []

    for _, row in df.iterrows():

        leg = _row_to_leg(row)
        sid = row["structure_id"]
        saved_entry_price = row.get("entry_price", "")
        saved_underlying_price = row.get("underlying_price", "")
        meta = {}

        if isinstance(leg, Future):
            if saved_entry_price:
                meta["entry_price"] = float(saved_entry_price)
            if saved_underlying_price:
                meta["settlement_price"] = float(saved_underlying_price)
        elif saved_entry_price:
            meta["settlement_price"] = float(saved_entry_price)

        if sid == "":
            portfolio.add(leg)
            position_meta.append(meta)
            continue

        if sid not in open_structures:
            structure = Structure([], name=row["structure_type"] or None)
            open_structures[sid] = structure
            portfolio.add(structure)
            position_meta.append(meta)

        open_structures[sid].legs.append(leg)

    if return_meta:
        return portfolio, position_meta

    return portfolio


def _row_to_leg(row):

    contract = row.get("contract") or None

    if row["instrument_type"] == "Future":
        return Future(quantity=float(row["quantity"]), contract=contract)

    if row["instrument_type"] == "AmericanOption":
        return AmericanOption(
            quantity=float(row["quantity"]),
            strike=float(row["strike"]),
            expiry_days=int(float(row["expiry_days"])),
            option_type=row["option_type"],
            volatility=float(row["volatility"]),
            contract=contract,
            option_month=row.get("option_month") or None,
        )

    raise ValueError(f"Unknown instrument_type: {row['instrument_type']}")


def export_report(df, path):
    """Thin, consistent wrapper for exporting any report DataFrame (ladders,
    stress tables, strike map, ...) to CSV -- creates the folder if needed."""

    directory = os.path.dirname(path)

    if directory:
        os.makedirs(directory, exist_ok=True)

    df.to_csv(path)
