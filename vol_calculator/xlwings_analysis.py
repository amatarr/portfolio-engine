"""
The analysis sheets (Greeks, Ladders, Heat Maps, Stress, Strike, Payoff,
Breakevens): each one's tables are computed here and written as plain
values to the hidden _Calc sheet, each under a workbook name; the sheets
themselves only show those names with formulas, and their charts plot
what the formulas show. Layout and formatting belong to the workbook.

Every sheet except Strike has its own spot selector, a named cell
(Spot_Greeks, Spot_Ladder, ...) whose dropdown offers "Follow Home" plus
every contract in the book. "Follow Home" -- or a contract no longer in
the book -- uses Home's Future Spot. The contract actually used is
written to SpotUsed_<sheet> so the sheet can show it.
"""

from vol_calculator.excel_sync import _round_greek
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
    report_by_tenor,
)

CALC_SHEET = "_Calc"
FOLLOW_HOME = "Follow Home"

HEAT_METRICS = ("PnL", "Delta", "Gamma", "Vega")
STRESS_KEYS = [
    "delta_price", "delta_time", "delta_vol", "gamma_price", "gamma_time", "gamma_vol",
    "vega_price", "vega_time", "theta_price", "theta_time",
]
PAYOFF_TIME_STEPS = (0, 30, 64)

# sheet -> (selector key, its blocks). The key names the sheet's spot
# selector (Spot_<key>) and the contract it used (SpotUsed_<key>).
SHEETS = {
    "Greeks": ("Greeks", ["Greeks_Summary", "Greeks_Tenor"]),
    "Spot&Vol Ladder": ("Ladder", ["Ladder_SpotTable", "Ladder_VolTable"]),
    "Heat Maps": ("HeatMaps", [f"Heat_{m}" for m in HEAT_METRICS]),
    "Stress": ("Stress", [f"Stress_{k}" for k in STRESS_KEYS]),
    "Strike": (None, ["Strike_Map"]),
    "Payoff": ("Payoff", ["Payoff_Table"]),
    "Breakevens": ("Breakevens", ["Breakevens_Summary", "Breakevens_Curve"]),
}

# Each block gets its own fixed band of columns on _Calc, in this order,
# so a block never moves between refreshes. Band 0 (columns A-B) holds
# the SpotUsed_* values.
_BLOCKS = [block for _, blocks in SHEETS.values() for block in blocks]
_BAND = 20
_CLEAR_ROWS = 400
_SPOT_KEYS = [key for key, _ in SHEETS.values() if key]


def _write_block(wb, name, rows):

    col = 1 + _BAND * (_BLOCKS.index(name) + 1)
    top_left = wb.sheets[CALC_SHEET].range((1, col))
    top_left.resize(_CLEAR_ROWS, _BAND).clear_contents()

    if rows:
        top_left.value = rows

    width = max((len(r) for r in rows), default=1)
    block = top_left.resize(max(len(rows), 1), width)
    wb.names.add(name, f"='{CALC_SHEET}'!{block.address}")


def _write_spot_used(wb, key, contract):

    cell = wb.sheets[CALC_SHEET].range((_SPOT_KEYS.index(key) + 1, 2))
    cell.offset(0, -1).value = key
    cell.value = contract
    wb.names.add(f"SpotUsed_{key}", f"='{CALC_SHEET}'!{cell.address}")


def _frame_rows(df, index_label=None):
    """A DataFrame as rows: header first, the index as the first column when index_label is given."""

    def clean(v):
        return round(v, 4) if isinstance(v, float) else v

    header = ([index_label] if index_label is not None else []) + list(df.columns)
    rows = [header]
    for idx, values in df.iterrows():
        rows.append(([idx] if index_label is not None else []) + [clean(v) for v in values])
    return rows


def _greeks(wb, portfolio, market, spot):

    summary = GreekEngine.report_dollars(portfolio, market, spot)
    _write_block(wb, "Greeks_Summary", [["Greek", "Value"]] + [[k, _round_greek(k, v)] for k, v in summary.items()])
    tenor = report_by_tenor(portfolio, market, spot, by_contract=True)
    _write_block(wb, "Greeks_Tenor", _frame_rows(tenor, "Tenor"))


def _ladder(wb, portfolio, market, spot):

    _write_block(wb, "Ladder_SpotTable", _frame_rows(spot_ladder(portfolio, market, spot)))
    _write_block(wb, "Ladder_VolTable", _frame_rows(vol_ladder(portfolio, market, spot)))


def _heat_maps(wb, portfolio, market, spot):

    for metric in HEAT_METRICS:
        df = spot_vol_surface(portfolio, market, spot, metric=metric)
        _write_block(wb, f"Heat_{metric}", _frame_rows(df, "Spot\\Vol"))


def _stress(wb, portfolio, market, spot):

    for key, df in stress_report(portfolio, market, spot).items():
        _write_block(wb, f"Stress_{key}", _frame_rows(df, df.index.name))


def _strike(wb, portfolio, market, spot):

    grid = strike_map(portfolio, weight="quantity")
    rows = _frame_rows(grid, "Strike\\Expiry(d)") if not grid.empty else [["No option legs in the book."]]
    _write_block(wb, "Strike_Map", rows)


def _payoff(wb, portfolio, market, spot):

    _write_block(wb, "Payoff_Table", _frame_rows(payoff_diagram(portfolio, market, list(PAYOFF_TIME_STEPS), spot)))


def _breakevens(wb, portfolio, market, spot, low=-15, high=15, step=1):

    def r4(v):
        return round(v, 4) if v is not None else None

    quad = breakevens_quadratic(portfolio, market, spot)
    numeric = breakevens_numerical(portfolio, market, spot)
    _write_block(wb, "Breakevens_Summary", [
        ["Method", "Downside (dF)", "Upside (dF)"],
        ["Quadratic (fast estimate)", r4(quad["downside"]), r4(quad["upside"])],
        ["Numerical (exact)", r4(numeric["downside"]), r4(numeric["upside"])],
    ])

    delta = GreekEngine.delta(portfolio, market, spot)
    gamma = GreekEngine.gamma(portfolio, market, spot)
    theta = GreekEngine.theta(portfolio, market)
    base_price = market.price_for(spot)

    curve = [["Spot Move %", "Quadratic PnL", "Numerical PnL (T+1d)", "Zero"]]
    for _, r in payoff_diagram(portfolio, market, [1], spot, low=low, high=high, step=step).iterrows():
        dF = base_price * r["Spot Move %"] / 100
        curve.append([r["Spot Move %"], round(delta * dF + 0.5 * gamma * dF ** 2 + theta, 4), round(r["T+1d"], 4), 0])
    _write_block(wb, "Breakevens_Curve", curve)


_WRITERS = {
    "Greeks": _greeks,
    "Spot&Vol Ladder": _ladder,
    "Heat Maps": _heat_maps,
    "Stress": _stress,
    "Strike": _strike,
    "Payoff": _payoff,
    "Breakevens": _breakevens,
}


def xw_sheet_spot(wb, sheet, home_spot, contracts):
    """The contract a sheet analyses: its own selector's choice, or Home's Future Spot when it says "Follow Home" (or names a contract no longer in the book)."""

    key, _ = SHEETS[sheet]
    if key is None:
        return home_spot
    try:
        chosen = wb.names[f"Spot_{key}"].refers_to_range.value
    except Exception:
        chosen = None
    return chosen if chosen in contracts else home_spot


def xw_sheets_following_home(wb, contracts):
    return [sheet for sheet, (key, _) in SHEETS.items() if key and xw_sheet_spot(wb, sheet, None, contracts) is None]


def xw_write_analysis(wb, portfolio, market, contracts, home_spot, sheets=None):
    """Recomputes the given analysis sheets (all of them by default). An empty book clears their tables instead."""

    for sheet in sheets or SHEETS:
        key, blocks = SHEETS[sheet]
        spot = xw_sheet_spot(wb, sheet, home_spot, contracts)

        if not contracts:
            for block in blocks:
                _write_block(wb, block, [])
        else:
            _WRITERS[sheet](wb, portfolio, market, spot)

        if key:
            _write_spot_used(wb, key, spot if contracts else None)
