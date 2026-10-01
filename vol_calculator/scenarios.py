import copy
import math

import pandas as pd
from scipy.optimize import brentq

from vol_calculator.greeks import GreekEngine, shift_portfolio_vols, shift_time
from vol_calculator.instruments import Portfolio, Structure, Future, AmericanOption


def _bump_contract_price(market, contract, new_price):
    """Copies market and sets just one contract's price, leaving every
    other contract's price (and every other position's valuation) untouched."""

    bumped = copy.deepcopy(market)
    key = contract if contract is not None else next(iter(bumped.futures_prices))
    bumped.futures_prices[key] = new_price
    return bumped


def spot_ladder(
    portfolio,
    market,
    contract=None,
    low=-20,
    high=20,
    step=5
):

    rows = []

    current_value = portfolio.value(market)
    base_price = market.price_for(contract)

    for pct in range(low, high + step, step):

        new_price = base_price * (1 + pct / 100)
        shocked_market = _bump_contract_price(market, contract, new_price)

        value = portfolio.value(shocked_market)

        rows.append({
            "Spot Move %": pct,
            "Futures Price": new_price,
            "PnL": value - current_value,
            "Portfolio Value": value,
            "Delta": GreekEngine.delta(portfolio, shocked_market, contract),
            "Gamma": GreekEngine.gamma(portfolio, shocked_market, contract),
            "Vega": GreekEngine.vega(portfolio, shocked_market, contract),
            "Vanna-V": GreekEngine.vanna_vega(portfolio, shocked_market, contract),
            "Volga": GreekEngine.volga(portfolio, shocked_market, contract),
        })

    return pd.DataFrame(rows)


def vol_ladder(
    portfolio,
    market,
    contract=None,
    low=-10,
    high=10,
    step=2
):

    rows = []

    current_value = portfolio.value(market)

    for vol_shift in range(low, high + step, step):

        shocked_portfolio = shift_portfolio_vols(portfolio, vol_shift / 100, contract)

        value = shocked_portfolio.value(market)

        rows.append({
            "Vol Shift (pts)": vol_shift,
            "PnL": value - current_value,
            "Portfolio Value": value,
            "Delta": GreekEngine.delta(shocked_portfolio, market, contract),
            "Gamma": GreekEngine.gamma(shocked_portfolio, market, contract),
            "Vega": GreekEngine.vega(shocked_portfolio, market, contract),
            "Volga": GreekEngine.volga(shocked_portfolio, market, contract),
            "Vanna-D": GreekEngine.vanna_delta(shocked_portfolio, market, contract),
            "Vanna-V": GreekEngine.vanna_vega(shocked_portfolio, market, contract),
        })

    return pd.DataFrame(rows)


def spot_vol_surface(
    portfolio,
    market,
    contract=None,
    metric="PnL",
    spot_low=-20,
    spot_high=20,
    spot_step=5,
    vol_low=-10,
    vol_high=10,
    vol_step=2
):

    current_value = portfolio.value(market)
    base_price = market.price_for(contract)

    spot_moves = list(range(spot_low, spot_high + spot_step, spot_step))
    vol_moves = list(range(vol_low, vol_high + vol_step, vol_step))

    data = []

    for spot_move in spot_moves:

        row = []

        for vol_move in vol_moves:

            shocked_market = _bump_contract_price(market, contract, base_price + spot_move)
            shocked_portfolio = shift_portfolio_vols(portfolio, vol_move / 100, contract)

            if metric == "PnL":
                value = shocked_portfolio.value(shocked_market) - current_value
            elif metric == "Delta":
                value = GreekEngine.delta(shocked_portfolio, shocked_market, contract)
            elif metric == "Gamma":
                value = GreekEngine.gamma(shocked_portfolio, shocked_market, contract)
            elif metric == "Vega":
                value = GreekEngine.vega(shocked_portfolio, shocked_market, contract)
            elif metric == "Volga":
                value = GreekEngine.volga(shocked_portfolio, shocked_market, contract)
            elif metric == "Vanna-D":
                value = GreekEngine.vanna_delta(shocked_portfolio, shocked_market, contract)
            elif metric == "Vanna-V":
                value = GreekEngine.vanna_vega(shocked_portfolio, shocked_market, contract)
            else:
                raise ValueError(f"Unknown metric: {metric}")

            row.append(value)

        data.append(row)

    return pd.DataFrame(data, index=spot_moves, columns=vol_moves)


def breakevens_quadratic(portfolio, market, contract=None, max_move_pct=0.10):
    """
    Fast, closed-form breakeven estimate from a 2nd-order Taylor expansion:
    Delta*dF + 1/2*Gamma*dF^2 + Theta = 0. Accurate for small moves and
    smooth (non-kinked) books; degrades for large moves or structures with
    kinks (butterflies, 3-ways) -- use breakevens_numerical for those.

    Delta/Gamma are with respect to the given contract (the "Future Spot"
    being analyzed) -- other contracts' legs still contribute to Theta/Value
    but don't move as this contract's price is shocked.

    Returns the futures MOVE dF in cents from the current price (downside
    negative, upside positive), not the absolute futures level. Roots
    farther than max_move_pct of spot are returned as None: constant
    Gamma stops being a valid assumption that far out, so they are
    Taylor artifacts, not real breakevens.
    """

    delta = GreekEngine.delta(portfolio, market, contract)
    gamma = GreekEngine.gamma(portfolio, market, contract)
    theta = GreekEngine.theta(portfolio, market)

    if gamma == 0:
        return {"downside": None, "upside": None}

    discriminant = delta ** 2 - 2 * gamma * theta

    if discriminant < 0:
        return {"downside": None, "upside": None}

    sqrt_disc = math.sqrt(discriminant)

    moves = sorted([
        (-delta + sqrt_disc) / gamma,
        (-delta - sqrt_disc) / gamma
    ])

    limit = market.price_for(contract) * max_move_pct

    return {
        "downside": moves[0] if abs(moves[0]) <= limit else None,
        "upside": moves[1] if abs(moves[1]) <= limit else None,
    }


def breakevens_numerical(portfolio, market, contract=None, search_pct=0.5, steps=200):
    """
    Exact 1-day breakeven: finds the move dF in the given contract's price
    where next-day P&L (whole portfolio, repriced one day closer to expiry)
    crosses zero, via a bracket scan + brentq root-find on the real pricer.
    No Taylor error, so this is the more trustworthy number once a book has
    kinked payoffs (butterflies, 3-ways) where breakevens_quadratic can
    mislead.

    Returns the futures MOVE dF in cents from the current price (downside
    negative, upside positive), not the absolute futures level.
    """

    shifted = shift_time(portfolio, 1)
    current_value = portfolio.value(market)
    base_price = market.price_for(contract)

    def pnl(spot_move):
        shocked_market = _bump_contract_price(market, contract, base_price + spot_move)
        return shifted.value(shocked_market) - current_value

    span = base_price * search_pct
    xs = [-span + i * (2 * span) / steps for i in range(steps + 1)]
    ys = [pnl(x) for x in xs]

    downside = None
    upside = None

    for i in range(len(xs) - 1):

        if ys[i] == 0:
            root = xs[i]
        elif ys[i] * ys[i + 1] < 0:
            root = brentq(pnl, xs[i], xs[i + 1])
        else:
            continue

        if root <= 0 and downside is None:
            downside = root
        elif root > 0 and upside is None:
            upside = root

    return {"downside": downside, "upside": upside}


def _flatten_positions(positions):

    flat = []

    for pos in positions:
        if isinstance(pos, Structure):
            flat.extend(_flatten_positions(pos.legs))
        else:
            flat.append(pos)

    return flat


def _intrinsic_leg_value(leg, spot):

    if isinstance(leg, Future):
        return leg.quantity * spot

    if isinstance(leg, AmericanOption):

        if leg.option_type.lower() == "call":
            payoff = max(spot - leg.strike, 0)
        else:
            payoff = max(leg.strike - spot, 0)

        return leg.quantity * payoff

    raise TypeError(f"Unsupported leg type for intrinsic value: {type(leg)}")


def _intrinsic_portfolio_value(portfolio, market, contract, shocked_spot):
    """
    Legs on `contract` are valued at shocked_spot; legs on every other
    contract are valued at their own current (unshocked) stored price --
    same "hold everything else where it is" rule the real pricer follows.
    """

    total = 0

    for leg in _flatten_positions(portfolio.positions):
        spot = shocked_spot if leg.contract == contract else market.price_for(leg.contract)
        total += _intrinsic_leg_value(leg, spot)

    return total


def payoff_at_expiry(
    portfolio,
    market,
    contract=None,
    low=-20,
    high=20,
    step=2
):
    """
    Pure intrinsic-value payoff curve for the whole portfolio, as if every
    leg had reached its own expiry (no time value, no vol, no rate), shocking
    only `contract`'s price -- every other contract's legs are held at their
    own current price. Computed directly rather than via the QuantLib pricer
    to avoid 0-day American option edge cases.
    """

    current_value = portfolio.value(market)
    base_price = market.price_for(contract)

    rows = []

    for pct in range(low, high + step, step):

        spot = base_price * (1 + pct / 100)
        value = _intrinsic_portfolio_value(portfolio, market, contract, spot)

        rows.append({
            "Spot Move %": pct,
            "Futures Price": spot,
            "PnL": value - current_value,
            "Value at Expiry": value,
        })

    return pd.DataFrame(rows)


def payoff_diagram(
    portfolio,
    market,
    time_steps,
    contract=None,
    low=-20,
    high=20,
    step=2
):
    """
    P&L vs. spot for the whole portfolio, overlaid across several
    time-to-expiry increments (e.g. time_steps=[0, 7, 14, 30]). Each column
    shifts the book that many days forward (shift_time) and reprices with the
    real pricer, so time value is preserved for legs that haven't reached
    expiry yet -- unlike payoff_at_expiry, which is pure intrinsic value.
    Only `contract`'s price is shocked; every other contract's legs are
    priced at their own current level throughout. Call payoff_at_expiry()
    separately for the final "at expiry" curve.
    """

    current_value = portfolio.value(market)
    base_price = market.price_for(contract)

    rows = []

    for pct in range(low, high + step, step):

        spot = base_price * (1 + pct / 100)
        shocked_market = _bump_contract_price(market, contract, spot)

        row = {
            "Spot Move %": pct,
            "Futures Price": spot,
        }

        for days in time_steps:
            shifted_portfolio = shift_time(portfolio, days)
            value = shifted_portfolio.value(shocked_market)
            row[f"T+{days}d"] = value - current_value

        rows.append(row)

    return pd.DataFrame(rows)


def report_by_tenor(portfolio, market, contract=None):
    """
    GreekEngine.report_dollars(), broken out by tenor bucket (each distinct
    expiry_days value, plus a separate "Futures" bucket for legs with no
    expiry) with a "Total" row for the whole book. Structures are
    flattened to their individual legs first, so a calendar spread's two
    legs correctly land in two different tenor rows. Delta/Gamma/etc are
    with respect to `contract`, same as everywhere else. Same $k scale as
    GreekEngine.report_dollars, so it's directly comparable to the main
    Greek Summary table.
    """

    legs = _flatten_positions(portfolio.positions)

    buckets = {}

    for leg in legs:

        label = "Futures" if isinstance(leg, Future) else leg.expiry_days

        buckets.setdefault(label, Portfolio())
        buckets[label].add(leg)

    def sort_key(label):
        return (-1, 0) if label == "Futures" else (0, label)

    rows = []

    for label in sorted(buckets.keys(), key=sort_key):

        row_label = label if label == "Futures" else f"{label}d"

        report = GreekEngine.report_dollars(buckets[label], market, contract)
        report = {"Tenor": row_label, **report}

        rows.append(report)

    total = {"Tenor": "Total", **GreekEngine.report_dollars(portfolio, market, contract)}
    rows.append(total)

    return pd.DataFrame(rows).set_index("Tenor")


_STRESS_GREEKS = {
    "delta": GreekEngine.delta,
    "gamma": GreekEngine.gamma,
    "vega": GreekEngine.vega,
    "theta": GreekEngine.theta,
}

_STRESS_AXIS_LABELS = {
    "price": "Price Move (¢)",
    "time": "Days Forward",
    "vol": "Vol Shift (pts)",
}


def stress_table(portfolio, market, greek, axis, magnitudes, contract=None):
    """
    Stress one variable at a time, hold others constant, read off one
    Greek's value at each stress level.

    greek: "delta" | "gamma" | "vega" | "theta"
    axis:  "price" (absolute cents added to `contract`'s price, +/- each
                     magnitude)
           "time"  (days shifted forward via shift_time, magnitudes only
                     -- there's no "un-decaying" time; applies to the
                     whole book, not just `contract`)
           "vol"   (vol points added to `contract`'s own legs, +/- each
                     magnitude)
    magnitudes: list of positive levels, e.g. [5, 10, 15, 20]
    """

    if greek not in _STRESS_GREEKS:
        raise ValueError(f"Unknown greek: {greek}")

    if axis not in _STRESS_AXIS_LABELS:
        raise ValueError(f"Unknown axis: {axis}")

    greek_fn = _STRESS_GREEKS[greek]
    base_price = market.price_for(contract)

    if axis == "time":
        levels = list(magnitudes)
    else:
        levels = [-m for m in reversed(magnitudes)] + list(magnitudes)

    rows = []

    for level in levels:

        if axis == "price":
            shocked_market = _bump_contract_price(market, contract, base_price + level)
            value = greek_fn(portfolio, shocked_market) if greek == "theta" \
                else greek_fn(portfolio, shocked_market, contract)

        elif axis == "time":
            shocked_portfolio = shift_time(portfolio, level)
            value = greek_fn(shocked_portfolio, market) if greek == "theta" \
                else greek_fn(shocked_portfolio, market, contract)

        else:  # vol
            shocked_portfolio = shift_portfolio_vols(portfolio, level / 100, contract)
            value = greek_fn(shocked_portfolio, market) if greek == "theta" \
                else greek_fn(shocked_portfolio, market, contract)

        rows.append({_STRESS_AXIS_LABELS[axis]: level, greek.capitalize(): value})

    return pd.DataFrame(rows).set_index(_STRESS_AXIS_LABELS[axis])


def stress_report(
    portfolio,
    market,
    contract=None,
    price_magnitudes=(5, 10, 15, 20),
    time_magnitudes=(5, 10, 15, 20),
    vol_magnitudes=(1, 2, 3, 4)
):
    """
    Runs the full required stress grid in one call:
    Delta x {price, time, vol}, Gamma x {price, time, vol},
    Vega x {price, time}, Theta x {price, time}.
    (Vega/vol and Theta/vol are deliberately excluded -- those are volga
    and a veta-type cross-Greek respectively, not part of this grid.)

    Returns a dict keyed "<greek>_<axis>" -> DataFrame.
    """

    combos = [
        ("delta", "price", price_magnitudes),
        ("delta", "time", time_magnitudes),
        ("delta", "vol", vol_magnitudes),
        ("gamma", "price", price_magnitudes),
        ("gamma", "time", time_magnitudes),
        ("gamma", "vol", vol_magnitudes),
        ("vega", "price", price_magnitudes),
        ("vega", "time", time_magnitudes),
        ("theta", "price", price_magnitudes),
        ("theta", "time", time_magnitudes),
    ]

    return {
        f"{greek}_{axis}": stress_table(portfolio, market, greek, axis, magnitudes, contract)
        for greek, axis, magnitudes in combos
    }


def strike_map(portfolio, market=None, weight="quantity"):
    """
    Pivot table: rows = strike, columns = expiry_days, cell = aggregated
    quantity (or net delta if weight="delta") across legs sharing that
    strike/expiry. Structures are flattened to their legs first. Future
    legs have no strike and are excluded. Delta weighting uses each leg's
    own contract (no single "Future Spot" selection needed here, since this
    view spans every contract in the book at once).
    """

    if weight not in ("quantity", "delta"):
        raise ValueError(f"Unknown weight: {weight}")

    if weight == "delta" and market is None:
        raise ValueError("market is required for weight='delta'")

    rows = []

    for leg in _flatten_positions(portfolio.positions):

        if isinstance(leg, Future):
            continue

        if weight == "quantity":
            value = leg.quantity
        else:
            single = Portfolio()
            single.add(leg)
            value = GreekEngine.delta(single, market, leg.contract)

        rows.append({
            "Strike": leg.strike,
            "Expiry": leg.expiry_days,
            "Value": value,
        })

    if not rows:
        return pd.DataFrame()

    return pd.DataFrame(rows).pivot_table(
        index="Strike",
        columns="Expiry",
        values="Value",
        aggfunc="sum",
        fill_value=0
    )
