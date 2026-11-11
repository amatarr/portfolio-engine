"""
Interactive (Altair / Vega-Lite) charts for the Streamlit app. Each builder
takes a table the scenario functions already return and gives back an
Altair chart. Hover readouts run in the browser, so they never rerun the
app or reprice anything. plotting.py keeps the matplotlib versions for
demo.py and notebooks.
"""

import math

import altair as alt
import pandas as pd

GREEN = "#2e7d32"
RED = "#c62828"
GRAY = "gray"
EXPIRY_COLOR = "#7f7f7f"
CURVE_COLORS = ["#1f6fd1", "#ef8a17", "#7b3fbf", "#159a8c", "#c2185b", "#6d8f00", "#8d6e63", "#455a64"]

UNITS = {
    "PnL": "$K",
    "Delta": "contracts",
    "Gamma": "Δ per ¢",
    "Vega": "$K per vol pt",
    "Volga": "$K per vol pt²",
    "Vanna-D": "Δ per vol pt",
    "Vanna-V": "$K per vol pt per ¢",
    "Theta": "$K per day",
}
_MONEY = {"PnL", "Vega", "Volga", "Vanna-V", "Theta"}


def money_factor(market, metric="PnL"):
    """$K scale for monetary measures; Delta/Gamma-type measures stay as they are."""

    return market.contract_multiplier / 1000 if metric in _MONEY else 1.0


def breakeven_label(base_price, move):
    if move is None:
        return "None in range"
    move = round(move, 2) + 0.0  # no "-0.00"
    return f"{base_price + move:,.2f} ({move:+.2f}¢)"


def crossings(xs, ys):
    """Where a sampled curve crosses zero, by linear interpolation between samples."""

    found = []
    for x0, x1, y0, y1 in zip(xs, xs[1:], ys, ys[1:]):
        if any(v is None or (isinstance(v, float) and math.isnan(v)) for v in (y0, y1)):
            continue
        if y0 == 0:
            root = x0
        elif y0 * y1 < 0:
            root = x0 + (x1 - x0) * (-y0) / (y1 - y0)
        else:
            continue
        if not found or abs(root - found[-1]) > 1e-9:
            found.append(root)
    if len(ys) and ys[-1] == 0 and (not found or abs(xs[-1] - found[-1]) > 1e-9):
        found.append(xs[-1])
    return found


def _zero_rule():
    return alt.Chart().mark_rule(color=GRAY, strokeWidth=1).encode(y=alt.datum(0))


def _current_price(base_price, x):
    current = pd.DataFrame({"Futures Price": [base_price], "Label": [f"Current {base_price:,.2f}"]})
    return [
        alt.Chart(current).mark_rule(color=GRAY, strokeDash=[3, 3]).encode(x=x),
        alt.Chart(current)
        .mark_text(align="left", dx=4, dy=-6, color=GRAY, baseline="bottom")
        .encode(x=x, y=alt.value(12), text="Label:N"),
    ]


def _hover(data, x, tooltip, field):
    """A vertical rule that follows the pointer, with every column of `tooltip` in its tooltip."""

    hover = alt.selection_point(fields=[field], nearest=True, on="pointerover", empty=False)
    return (
        alt.Chart(data)
        .mark_rule(color=GRAY)
        .encode(x=x, opacity=alt.condition(hover, alt.value(0.6), alt.value(0)), tooltip=tooltip)
        .add_params(hover)
    )


def _price_axis(contract):
    return alt.X(
        "Futures Price:Q", title=f"{contract} futures price (¢/bu)", scale=alt.Scale(zero=False, nice=False)
    )


def _diverging_color(values, scheme, title, reverse=False):
    finite = [abs(v) for v in values if v is not None and not math.isnan(v)]
    extent = max(finite) if finite and max(finite) > 0 else 1.0
    return alt.Color(
        "Value:Q",
        scale=alt.Scale(scheme=scheme, domain=[-extent, extent], reverse=reverse),
        legend=alt.Legend(title=title),
    ), extent


def breakeven_chart(curve, roots, base_price, contract, days, low, high):
    """
    `curve` is breakeven_curve() output with "Exact $K" / "Quadratic $K"
    columns added; `roots` are the exact zero crossings (moves in cents).
    """

    curve = curve.copy()
    curve["Profit"] = curve["Exact $K"].clip(lower=0)
    curve["Loss"] = curve["Exact $K"].clip(upper=0)

    x = _price_axis(contract)
    y_title = f"P&L in {days} day{'s' if days != 1 else ''} ($K)"

    shading = [
        alt.Chart(curve).mark_area(color=color, opacity=0.18)
        .encode(x=x, y=alt.Y(f"{column}:Q", title=y_title), y2=alt.datum(0))
        for column, color in (("Profit", GREEN), ("Loss", RED))
    ]

    lines = (
        alt.Chart(curve)
        .transform_fold(["Exact $K", "Quadratic $K"], as_=["Curve", "PnL"])
        .transform_filter("isValid(datum.PnL)")
        .transform_calculate(Curve="datum.Curve == 'Exact $K' ? 'Exact' : 'Quadratic (estimate)'")
        .mark_line(strokeWidth=2.5)
        .encode(
            x=x,
            y=alt.Y("PnL:Q", title=y_title),
            color=alt.Color(
                "Curve:N",
                scale=alt.Scale(domain=["Exact", "Quadratic (estimate)"], range=CURVE_COLORS[:2]),
                legend=alt.Legend(title=None, orient="top"),
            ),
            strokeDash=alt.StrokeDash(
                "Curve:N",
                scale=alt.Scale(domain=["Exact", "Quadratic (estimate)"], range=[[1, 0], [6, 4]]),
                legend=None,
            ),
        )
    )

    layers = [*shading, _zero_rule(), *_current_price(base_price, x), lines]

    visible = [r for r in roots if low <= r <= high]
    if visible:
        # Put each label on the side the curve isn't: up-left where P&L rises
        # through zero, up-right where it falls.
        rising = []
        for r in visible:
            after = curve.loc[curve["Move"] > r, "Exact PnL"]
            rising.append(bool(len(after)) and after.iloc[0] > 0)
        points = pd.DataFrame({
            "Futures Price": [base_price + r for r in visible],
            "Zero": [0.0] * len(visible),
            "Label": [breakeven_label(base_price, r) for r in visible],
            "Rising": rising,
        })
        point_base = alt.Chart(points).encode(x=x, y=alt.Y("Zero:Q", title=y_title))
        layers += [
            point_base.mark_point(filled=True, size=90, color="black", stroke="white", strokeWidth=1.5),
            point_base.transform_filter("datum.Rising")
            .mark_text(align="right", dx=-8, dy=-10, fontWeight="bold").encode(text="Label:N"),
            point_base.transform_filter("!datum.Rising")
            .mark_text(align="left", dx=8, dy=-10, fontWeight="bold").encode(text="Label:N"),
        ]

    layers.append(_hover(curve, x, [
        alt.Tooltip("Futures Price:Q", title="Price (¢/bu)", format=",.2f"),
        alt.Tooltip("Move:Q", title="Move (¢)", format="+.2f"),
        alt.Tooltip("Exact $K:Q", title="Exact ($K)", format=",.2f"),
        alt.Tooltip("Quadratic $K:Q", title="Quadratic ($K)", format=",.2f"),
    ], "Futures Price"))

    return alt.layer(*layers).properties(height=420)


def payoff_chart(curves, base_price, contract, curve_names):
    """
    `curves` is payoff_curves() output already in $K: one column per name
    in curve_names (e.g. "T+0d", ..., "At Expiry"). Click a legend entry to
    highlight that curve; dots mark where each curve crosses zero.
    """

    x = _price_axis(contract)
    colors = [EXPIRY_COLOR if name == "At Expiry" else CURVE_COLORS[i % len(CURVE_COLORS)]
              for i, name in enumerate(curve_names)]
    dashes = [[6, 4] if name == "At Expiry" else [1, 0] for name in curve_names]
    color = alt.Color("Curve:N", scale=alt.Scale(domain=curve_names, range=colors),
                      legend=alt.Legend(title=None, orient="top"))
    pick = alt.selection_point(fields=["Curve"], bind="legend")
    faded = alt.condition(pick, alt.value(1), alt.value(0.15))

    long = curves.melt(id_vars=["Futures Price", "Move"], value_vars=curve_names, var_name="Curve", value_name="PnL")
    lines = (
        alt.Chart(long)
        .mark_line(strokeWidth=2.5)
        .encode(
            x=x,
            y=alt.Y("PnL:Q", title="P&L ($K)"),
            color=color,
            strokeDash=alt.StrokeDash("Curve:N", scale=alt.Scale(domain=curve_names, range=dashes), legend=None),
            opacity=faded,
        )
        .add_params(pick)
    )

    layers = [_zero_rule(), *_current_price(base_price, x), lines]

    prices = curves["Futures Price"].tolist()
    dots = [
        {"Curve": name, "Futures Price": root, "Zero": 0.0, "Label": breakeven_label(base_price, root - base_price)}
        for name in curve_names
        for root in crossings(prices, curves[name].tolist())
    ]
    if dots:
        layers.append(
            alt.Chart(pd.DataFrame(dots))
            .mark_point(filled=True, size=80, stroke="white", strokeWidth=1.5)
            .encode(
                x=x, y=alt.Y("Zero:Q"), color=color, opacity=faded,
                tooltip=[alt.Tooltip("Curve:N"), alt.Tooltip("Label:N", title="Breakeven")],
            )
        )

    layers.append(_hover(curves, x, [
        alt.Tooltip("Futures Price:Q", title="Price (¢/bu)", format=",.2f"),
        alt.Tooltip("Move:Q", title="Move (¢)", format="+.2f"),
        *[alt.Tooltip(f"{name}:Q", title=f"{name} ($K)", format=",.2f") for name in curve_names],
    ], "Futures Price"))

    return alt.layer(*layers).properties(height=420)


def _heatmap(long, x_field, y_field, y_sort, value_format, scheme, legend_title, reverse, tooltip, height):
    color, extent = _diverging_color(long["Value"].tolist(), scheme, legend_title, reverse)
    long = long.assign(Strong=long["Value"].abs() > 0.6 * extent)

    base = alt.Chart(long).encode(
        x=alt.X(f"{x_field}:O", axis=alt.Axis(labelAngle=0)),
        y=alt.Y(f"{y_field}:O", sort=y_sort),
    )
    cells = base.mark_rect(stroke="white", strokeWidth=0.5).encode(color=color, tooltip=tooltip)
    text = base.mark_text(fontSize=11).encode(
        text=alt.Text("Value:Q", format=value_format),
        color=alt.condition("datum.Strong", alt.value("white"), alt.value("black")),
        tooltip=tooltip,
    )
    return cells, text, base, height


def surface_chart(surface, metric):
    """
    `surface` is spot_vol_surface() output, already scaled (rows = price
    move in cents, columns = vol shift in points). Higher prices at the top;
    today's cell (no move, no vol shift) is outlined.
    """

    unit = UNITS.get(metric, "")
    long = pd.DataFrame(
        [
            {"Price Move (¢)": price_move, "Vol Shift (pts)": vol_move, "Value": float(surface.loc[price_move, vol_move])}
            for price_move in surface.index
            for vol_move in surface.columns
        ]
    )
    tooltip = [
        alt.Tooltip("Price Move (¢):O"),
        alt.Tooltip("Vol Shift (pts):O"),
        alt.Tooltip("Value:Q", title=f"{metric} ({unit})" if unit else metric, format=",.2f"),
    ]
    scheme = "redyellowgreen" if metric == "PnL" else "redblue"
    largest = long["Value"].abs().max() if len(long) else 0
    value_format = ",.1f" if largest >= 10 else ",.2f" if largest >= 1 else ",.3f"
    cells, text, base, height = _heatmap(
        long, "Vol Shift (pts)", "Price Move (¢)", "descending", value_format, scheme,
        f"{metric} ({unit})" if unit else metric, metric != "PnL", tooltip, max(240, 34 * len(surface.index)),
    )
    today = base.transform_filter("datum['Price Move (¢)'] == 0 && datum['Vol Shift (pts)'] == 0").mark_rect(
        fill=None, stroke="black", strokeWidth=2.5
    )
    return alt.layer(cells, text, today).properties(height=height)


def strike_map_chart(grid, weight):
    """
    `grid` is strike_map(fill_value=None) output (rows = strike, columns =
    expiry days). Pairs with no position stay blank; higher strikes on top.
    """

    label = "Net Quantity" if weight == "quantity" else "Net Delta"
    rows = []
    for strike in grid.index:
        for expiry in grid.columns:
            value = grid.loc[strike, expiry]
            if pd.notna(value):
                rows.append({"Strike": f"{strike:g}", "Expiry (d)": int(expiry), "Value": float(value)})
    long = pd.DataFrame(rows)
    strike_order = [f"{s:g}" for s in sorted(grid.index, reverse=True)]
    tooltip = [
        alt.Tooltip("Strike:N"),
        alt.Tooltip("Expiry (d):O"),
        alt.Tooltip("Value:Q", title=label, format=",.2~f"),
    ]
    cells, text, _, height = _heatmap(
        long, "Expiry (d)", "Strike", strike_order, ",.2~f", "redblue", label, True, tooltip,
        max(160, 32 * len(grid.index)),
    )
    return alt.layer(cells, text).properties(height=height)


def stress_chart(table, greek, axis_label, unit):
    """
    One stress_table(include_base=True) result, already scaled: the Greek
    against one stressed variable, with today's (unstressed) point ringed.
    """

    data = table.reset_index().rename(columns={axis_label: "Level", greek: "Value"})
    x = alt.X("Level:Q", title=axis_label, scale=alt.Scale(nice=False))
    y_title = f"{greek} ({unit})" if unit else greek
    base = alt.Chart(data).encode(x=x, y=alt.Y("Value:Q", title=y_title))

    layers = [
        _zero_rule(),
        base.mark_line(strokeWidth=2.5, color=CURVE_COLORS[0]),
        base.mark_point(filled=True, size=45, color=CURVE_COLORS[0]),
        base.transform_filter("datum.Level == 0").mark_point(size=160, color="black", strokeWidth=2),
        _hover(data, x, [
            alt.Tooltip("Level:Q", title=axis_label, format="+,.2~f"),
            alt.Tooltip("Value:Q", title=y_title, format=",.3~f"),
        ], "Level"),
    ]
    return alt.layer(*layers).properties(height=220)
