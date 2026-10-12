from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.marker import Marker

from vol_calculator.instruments import Future, AmericanOption
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


def clear_range(ws, min_row=1, max_row=500, min_col=1, max_col=20):
    """Blanks out a rectangular range -- used before rewriting a sheet/block so a shrinking table doesn't leave stale cells behind."""

    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
        for cell in row:
            cell.value = None


def write_dataframe(ws, df, start_row=1, start_col=1, label=None, include_index=False, index_label=""):
    """
    Writes a pandas DataFrame as a plain header + rows block, starting at
    (start_row, start_col). If label is given, it goes on its own row
    just above the header. With include_index=True, the DataFrame's
    index becomes the first column (for pivot-shaped outputs like
    spot_vol_surface/strike_map, where the index itself is meaningful --
    e.g. spot move or strike). Returns the row just after the last one
    written, so callers can stack multiple tables on one sheet.
    """

    row = start_row
    col_offset = start_col + (1 if include_index else 0)

    if label is not None:
        ws.cell(row=row, column=start_col, value=label)
        row += 1

    if include_index:
        ws.cell(row=row, column=start_col, value=index_label)
    for col, header in enumerate(df.columns, start=col_offset):
        ws.cell(row=row, column=col, value=header)
    row += 1

    for idx, data_row in df.iterrows():
        if include_index:
            ws.cell(row=row, column=start_col, value=idx)
        for col, value in enumerate(data_row, start=col_offset):
            ws.cell(row=row, column=col, value=round(value, 4) if isinstance(value, float) else value)
        row += 1

    return row

POSITIONS_HEADERS = [
    "Contract", "Commodity", "Type", "Quantity",
    "Strike", "Expiry (d)", "Vol (%)", "Settlement Price", "Source",
]

FUTURE_SPOT_LABEL_CELL = "A10"
FUTURE_SPOT_VALUE_CELL = "B10"


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

    ws[label_cell] = "Interest Rate (%):"

    if sofr:
        ws[value_cell] = round(sofr["rate"] * 100, 2)
    elif ws[value_cell].value is None:
        ws[value_cell] = round(fallback_rate * 100, 2)

    if sofr:
        ws[caption_cell] = (
            f"SOFR: {sofr['rate'] * 100:.2f}% (updated {sofr['effective_date']}) -- {sofr['source_url']}"
        )
    else:
        ws[caption_cell] = "SOFR unavailable -- using manually entered rate."


def read_interest_rate(ws, fallback_rate=0.039, value_cell=INTEREST_RATE_VALUE_CELL):
    """Reads the Interest Rate (%) cell as a decimal, falling back if empty/invalid."""

    value = ws[value_cell].value
    try:
        return float(value) / 100
    except (TypeError, ValueError):
        return fallback_rate


# Column J onward -- deliberately outside columns 1-9 that
# write_positions_table's clear_range touches, so this block survives
# every refresh regardless of how many Positions rows there are.
ADD_LEG_LABEL_CELL = "J13"
ADD_LEG_TYPE_LABEL_CELL, ADD_LEG_TYPE_VALUE_CELL = "J14", "K14"
ADD_LEG_CONTRACT_LABEL_CELL, ADD_LEG_CONTRACT_VALUE_CELL = "J15", "K15"
ADD_LEG_QUANTITY_LABEL_CELL, ADD_LEG_QUANTITY_VALUE_CELL = "J16", "K16"
ADD_LEG_STRIKE_LABEL_CELL, ADD_LEG_STRIKE_VALUE_CELL = "J17", "K17"
ADD_LEG_EXPIRY_LABEL_CELL, ADD_LEG_EXPIRY_VALUE_CELL = "J18", "K18"
ADD_LEG_VOL_LABEL_CELL, ADD_LEG_VOL_VALUE_CELL = "J19", "K19"
ADD_LEG_STATUS_CELL = "J20"

ADD_LEG_TYPES = ["Future", "American Option"]


def write_add_leg_block(ws, contracts):
    """Add Leg input cells + dropdowns -- v1 covers Future and American Option only (the only types real Daily Prop data ever contains)."""

    ws[ADD_LEG_LABEL_CELL] = "ADD LEG"
    ws[ADD_LEG_TYPE_LABEL_CELL] = "Type:"
    ws[ADD_LEG_CONTRACT_LABEL_CELL] = "Contract:"
    ws[ADD_LEG_QUANTITY_LABEL_CELL] = "Quantity:"
    ws[ADD_LEG_STRIKE_LABEL_CELL] = "Strike (blank for Future):"
    ws[ADD_LEG_EXPIRY_LABEL_CELL] = "Expiry d (blank=auto):"
    ws[ADD_LEG_VOL_LABEL_CELL] = "Vol % (blank for Future):"

    if ws[ADD_LEG_TYPE_VALUE_CELL].value not in ADD_LEG_TYPES:
        ws[ADD_LEG_TYPE_VALUE_CELL] = ADD_LEG_TYPES[0]

    type_dv = DataValidation(type="list", formula1=f'"{",".join(ADD_LEG_TYPES)}"', allow_blank=False)
    ws.add_data_validation(type_dv)
    type_dv.add(ws[ADD_LEG_TYPE_VALUE_CELL])

    if contracts:
        if ws[ADD_LEG_CONTRACT_VALUE_CELL].value not in contracts:
            ws[ADD_LEG_CONTRACT_VALUE_CELL] = contracts[0]

        contract_dv = DataValidation(type="list", formula1=f'"{",".join(contracts)}"', allow_blank=False)
        ws.add_data_validation(contract_dv)
        contract_dv.add(ws[ADD_LEG_CONTRACT_VALUE_CELL])


def read_add_leg_inputs(ws):
    return {
        "type": ws[ADD_LEG_TYPE_VALUE_CELL].value,
        "contract": ws[ADD_LEG_CONTRACT_VALUE_CELL].value,
        "quantity": ws[ADD_LEG_QUANTITY_VALUE_CELL].value,
        "strike": ws[ADD_LEG_STRIKE_VALUE_CELL].value,
        "expiry_days": ws[ADD_LEG_EXPIRY_VALUE_CELL].value,
        "vol_pct": ws[ADD_LEG_VOL_VALUE_CELL].value,
    }


def clear_add_leg_inputs(ws):
    for cell in (
        ADD_LEG_QUANTITY_VALUE_CELL, ADD_LEG_STRIKE_VALUE_CELL,
        ADD_LEG_EXPIRY_VALUE_CELL, ADD_LEG_VOL_VALUE_CELL,
    ):
        ws[cell] = None


def build_manual_row(inputs, commodity_lookup, auto_expiry_fn):
    """
    Validates Add Leg's raw input dict and turns it into a POSITIONS_
    HEADERS-shaped {header: value} row, Source="Manual". Raises ValueError
    with a human-readable message on anything missing/invalid -- the
    caller is expected to surface that directly (e.g. write it into a
    status cell), not crash the whole refresh.

    commodity_lookup: contract -> commodity, e.g. app.py's
    _commodity_from_contract. auto_expiry_fn: contract -> approx expiry
    days, e.g. vol_calculator.contracts.approx_option_expiry_days --
    used only when expiry_days is left blank for an option.
    """

    leg_type = inputs.get("type")
    contract = inputs.get("contract")

    if leg_type not in ADD_LEG_TYPES:
        raise ValueError(f"Pick a Type ({'/'.join(ADD_LEG_TYPES)}).")
    if not contract:
        raise ValueError("Pick a Contract.")

    quantity = inputs.get("quantity")
    if quantity in (None, ""):
        raise ValueError("Quantity is required.")

    commodity = commodity_lookup(contract)

    if leg_type == "Future":
        return {
            "Contract": contract, "Commodity": commodity, "Type": "Future",
            "Quantity": float(quantity), "Strike": None, "Expiry (d)": None,
            "Vol (%)": None, "Settlement Price": None, "Source": "Manual",
        }

    strike = inputs.get("strike")
    vol_pct = inputs.get("vol_pct")

    if strike in (None, ""):
        raise ValueError("Strike is required for American Option.")
    if vol_pct in (None, ""):
        raise ValueError("Vol (%) is required for American Option.")

    expiry_days = inputs.get("expiry_days")
    if expiry_days in (None, ""):
        expiry_days = auto_expiry_fn(contract)
        if expiry_days is None:
            raise ValueError("Couldn't auto-compute Expiry -- enter it manually.")

    return {
        "Contract": contract, "Commodity": commodity, "Type": "Call",
        "Quantity": float(quantity), "Strike": float(strike),
        "Expiry (d)": int(expiry_days), "Vol (%)": float(vol_pct),
        "Settlement Price": None, "Source": "Manual",
    }


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
    """

    prices = {}
    for _, meta in imported:
        price = meta.get("settlement_price")
        if price is not None:
            prices[meta["contract"]] = price
    return prices


def _position_row(instrument, meta, source):

    if isinstance(instrument, Future):
        return [
            meta["contract"], meta["commodity"], "Future", instrument.quantity,
            None, None, meta.get("vol_pct"), meta.get("settlement_price"), source,
        ]

    if isinstance(instrument, AmericanOption):
        return [
            meta["contract"], meta["commodity"],
            "Call" if instrument.option_type == "call" else "Put",
            instrument.quantity, instrument.strike, instrument.expiry_days,
            round(instrument.volatility * 100, 2), meta.get("settlement_price"), source,
        ]

    raise TypeError(f"Unsupported instrument type: {type(instrument)}")


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
    elif leg_type in ("Call", "Put"):
        instrument = AmericanOption(
            quantity=quantity,
            strike=float(row["Strike"]),
            expiry_days=int(row["Expiry (d)"]),
            option_type="call" if leg_type == "Call" else "put",
            volatility=float(row.get("Vol (%)") or 0) / 100,
            contract=contract,
        )
    else:
        raise ValueError(f"Unknown Type in Positions row: {leg_type!r}")

    meta = {
        "commodity": commodity, "contract": contract,
        "vol_pct": None, "settlement_price": row.get("Settlement Price"),
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
        ws.cell(row=1, column=col, value=value)

    for i, contract in enumerate(contracts):
        dollars = GreekEngine.report_dollars(portfolio, market, contract)
        row_values = [contract] + [_round_greek(k, v) for k, v in dollars.items()]
        for col, value in enumerate(row_values, start=1):
            ws.cell(row=2 + i, column=col, value=value)

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
