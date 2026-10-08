from vol_calculator.instruments import Future, AmericanOption

POSITIONS_HEADERS = [
    "Contract", "Commodity", "Type", "Quantity",
    "Strike", "Expiry (d)", "Vol (%)", "Settlement Price",
]


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


def write_positions_sheet(workbook, imported, sheet_name="Positions"):
    """
    Writes (instrument, meta) pairs -- the same shape import_bushel_positions
    returns -- into a sheet as a plain header + rows table. Overwrites
    whatever was there, since this always reflects the latest Daily Prop
    file, not something to hand-edit in place.
    """

    if sheet_name in workbook.sheetnames:
        del workbook[sheet_name]

    ws = workbook.create_sheet(sheet_name)

    ws.append(POSITIONS_HEADERS)

    for instrument, meta in imported:
        ws.append(_position_row(instrument, meta))

    return ws
