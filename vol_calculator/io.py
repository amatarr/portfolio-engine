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
]


def portfolio_to_csv(portfolio, path):
    """
    One row per leg. Legs belonging to the same Structure (call spread,
    straddle, etc.) share a structure_id so portfolio_from_csv can regroup
    them back into a Structure on reload, instead of a bag of anonymous
    legs. Bare (non-Structure) legs get an empty structure_id.
    """

    rows = []
    structure_id = 0

    for pos in portfolio.positions:

        if isinstance(pos, Structure):

            for leg in pos.legs:
                rows.append(_leg_to_row(leg, structure_id, pos.name or ""))

            structure_id += 1

        else:
            rows.append(_leg_to_row(pos, "", ""))

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
        }

    raise TypeError(f"Unsupported leg type for CSV export: {type(leg)}")


def portfolio_from_csv(path):
    """
    Inverse of portfolio_to_csv. Rows sharing a structure_id are regrouped
    into one Structure, added to the portfolio at the position of that
    structure's first row, so the original leg ordering (including any
    interleaving of bare legs and structures) round-trips exactly.
    """

    df = pd.read_csv(path, dtype=str, keep_default_na=False)

    portfolio = Portfolio()
    open_structures = {}

    for _, row in df.iterrows():

        leg = _row_to_leg(row)
        sid = row["structure_id"]

        if sid == "":
            portfolio.add(leg)
            continue

        if sid not in open_structures:
            structure = Structure([], name=row["structure_type"] or None)
            open_structures[sid] = structure
            portfolio.add(structure)

        open_structures[sid].legs.append(leg)

    return portfolio


def _row_to_leg(row):

    if row["instrument_type"] == "Future":
        return Future(quantity=float(row["quantity"]))

    if row["instrument_type"] == "AmericanOption":
        return AmericanOption(
            quantity=float(row["quantity"]),
            strike=float(row["strike"]),
            expiry_days=int(float(row["expiry_days"])),
            option_type=row["option_type"],
            volatility=float(row["volatility"])
        )

    raise ValueError(f"Unknown instrument_type: {row['instrument_type']}")


def export_report(df, path):
    """Thin, consistent wrapper for exporting any report DataFrame (ladders,
    stress tables, strike map, ...) to CSV -- creates the folder if needed."""

    directory = os.path.dirname(path)

    if directory:
        os.makedirs(directory, exist_ok=True)

    df.to_csv(path)
