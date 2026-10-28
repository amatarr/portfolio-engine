import copy
import math

import pandas as pd
from scipy.optimize import brentq

from vol_calculator.greeks import GreekEngine, shift_portfolio_vols, shift_time
from vol_calculator.instruments import Portfolio, Structure, Future, AmericanOption
from vol_calculator.contracts import (
    _CME_MONTH_TO_NUMBER,
    MONTH_CODES,
    commodity_from_contract,
    option_expiry_candidates,
)


def _bump_contract_price(market, contract, new_price):
    """Copies market and sets just one contract's price, leaving every
    other contract's price (and every other position's valuation) untouched."""

    bumped = copy.deepcopy(market)
    key = contract if contract is not None else next(iter(bumped.futures_prices))
    bumped.futures_prices[key] = new_price
    return bumped


def _theta_rho(portfolio, market):

    return {
        "Theta": GreekEngine.theta(portfolio, market),
        "Rho": GreekEngine.rho(portfolio, market),
    }


def spot_ladder(
    portfolio,
    market,
    contract=None,
    low=-50,
    high=50,
    step=10,
    with_theta_rho=False
):
    """
    low/high/step are absolute cents moves on `contract`'s price, not percentages.
    with_theta_rho adds Theta and Rho (before Vega) at each price level; off by
    default so the Excel ladder sheet keeps its columns.
    """

    rows = []

    current_value = portfolio.value(market)
    base_price = market.price_for(contract)

    for cents in range(low, high + step, step):

        new_price = base_price + cents
        shocked_market = _bump_contract_price(market, contract, new_price)

        value = portfolio.value(shocked_market)

        rows.append({
            "Spot Move (¢)": cents,
            "Futures Price": new_price,
            "PnL": value - current_value,
            "Portfolio Value": value,
            "Delta": GreekEngine.delta(portfolio, shocked_market, contract),
            "Gamma": GreekEngine.gamma(portfolio, shocked_market, contract),
            **(_theta_rho(portfolio, shocked_market) if with_theta_rho else {}),
            "Vega": GreekEngine.vega(portfolio, shocked_market, contract),
            "Vanna-V": GreekEngine.vanna_vega(portfolio, shocked_market, contract),
            "Volga": GreekEngine.volga(portfolio, shocked_market, contract),
        })

    return pd.DataFrame(rows)


def spot_ladder_by_underlying(portfolio, market, commodity, metric, low=-50, high=50, step=10):
    """Return one selected price-ladder metric per underlying contract in a commodity."""

    valid_metrics = {
        "Futures Price", "PnL", "Portfolio Value", "Delta", "Gamma", "Theta", "Rho",
        "Vega", "Volga", "Vanna-V",
    }
    if metric not in valid_metrics:
        raise ValueError(f"Unknown price ladder metric: {metric}")

    legs_by_contract = {}
    for leg in _flatten_positions(portfolio.positions):
        if commodity_from_contract(leg.contract, fallback=None) == commodity:
            legs_by_contract.setdefault(leg.contract, []).append(leg)

    rows = {}
    for contract in sorted(legs_by_contract, key=_contract_sort_key):
        contract_portfolio = Portfolio()
        for leg in legs_by_contract[contract]:
            contract_portfolio.add(leg)
        ladder = spot_ladder(
            contract_portfolio, market, contract=contract, low=low, high=high, step=step,
            with_theta_rho=True,
        )
        rows[contract] = ladder[metric].tolist()

    moves = list(range(low, high + step, step))
    return pd.DataFrame.from_dict(rows, orient="index", columns=moves)


def _portfolio_for_commodity(portfolio, commodity):
    filtered = Portfolio()
    for leg in _flatten_positions(portfolio.positions):
        if commodity_from_contract(leg.contract, fallback=None) == commodity:
            filtered.add(leg)
    return filtered


def _commodity_ladder_greeks(portfolio, market, contracts):
    greeks = {}
    for label, greek in (
        ("Delta", GreekEngine.delta),
        ("Gamma", GreekEngine.gamma),
        ("Vega", GreekEngine.vega),
        ("Volga", GreekEngine.volga),
        ("Vanna-D", GreekEngine.vanna_delta),
        ("Vanna-V", GreekEngine.vanna_vega),
    ):
        greeks[label] = sum(greek(portfolio, market, contract) for contract in contracts)
    greeks.update(_theta_rho(portfolio, market))
    return greeks


def commodity_spot_ladder(portfolio, market, commodity, low=-50, high=50, step=10, with_theta_rho=False):
    """Price ladder for one commodity, moving every underlying contract in parallel."""

    commodity_portfolio = _portfolio_for_commodity(portfolio, commodity)
    contracts = sorted({leg.contract for leg in commodity_portfolio.positions})
    current_value = commodity_portfolio.value(market)
    rows = []

    for cents in range(low, high + step, step):
        shocked_market = copy.deepcopy(market)
        for contract in contracts:
            shocked_market.futures_prices[contract] += cents
        value = commodity_portfolio.value(shocked_market)
        greeks = _commodity_ladder_greeks(commodity_portfolio, shocked_market, contracts)
        row = {
            "Spot Move (¢)": cents,
            "PnL": value - current_value,
            "Portfolio Value": value,
            "Delta": greeks["Delta"],
            "Gamma": greeks["Gamma"],
            "Vega": greeks["Vega"],
            "Vanna-V": greeks["Vanna-V"],
            "Volga": greeks["Volga"],
        }
        if with_theta_rho:
            row.update({"Theta": greeks["Theta"], "Rho": greeks["Rho"]})
        rows.append(row)

    return pd.DataFrame(rows)


def vol_ladder(
    portfolio,
    market,
    contract=None,
    low=-10,
    high=10,
    step=2,
    with_theta_rho=False
):
    """with_theta_rho: same as spot_ladder's."""

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
            **(_theta_rho(shocked_portfolio, market) if with_theta_rho else {}),
            "Vega": GreekEngine.vega(shocked_portfolio, market, contract),
            "Volga": GreekEngine.volga(shocked_portfolio, market, contract),
            "Vanna-D": GreekEngine.vanna_delta(shocked_portfolio, market, contract),
            "Vanna-V": GreekEngine.vanna_vega(shocked_portfolio, market, contract),
        })

    return pd.DataFrame(rows)


def commodity_vol_ladder(portfolio, market, commodity, low=-10, high=10, step=2, with_theta_rho=False):
    """Volatility ladder for one commodity, shifting options on all its contracts in parallel."""

    commodity_portfolio = _portfolio_for_commodity(portfolio, commodity)
    contracts = sorted({leg.contract for leg in commodity_portfolio.positions})
    current_value = commodity_portfolio.value(market)
    rows = []

    for vol_shift in range(low, high + step, step):
        shift = vol_shift / 100
        shocked = Portfolio()
        for leg in commodity_portfolio.positions:
            shocked.add(leg.shift_vol(shift))

        value = shocked.value(market)
        greeks = _commodity_ladder_greeks(shocked, market, contracts)
        row = {
            "Vol Shift (pts)": vol_shift,
            "PnL": value - current_value,
            "Portfolio Value": value,
            "Delta": greeks["Delta"],
            "Gamma": greeks["Gamma"],
            "Vega": greeks["Vega"],
            "Volga": greeks["Volga"],
            "Vanna-D": greeks["Vanna-D"],
            "Vanna-V": greeks["Vanna-V"],
        }
        if with_theta_rho:
            row.update({"Theta": greeks["Theta"], "Rho": greeks["Rho"]})
        rows.append(row)

    return pd.DataFrame(rows)


def vol_ladder_by_underlying(portfolio, market, commodity, metric, low=-10, high=10, step=2):
    """Return one selected volatility-ladder metric per underlying in a commodity."""

    valid_metrics = {
        "PnL", "Portfolio Value", "Delta", "Gamma", "Theta", "Rho", "Vega", "Volga",
        "Vanna-D", "Vanna-V",
    }
    if metric not in valid_metrics:
        raise ValueError(f"Unknown volatility ladder metric: {metric}")

    legs_by_contract = {}
    for leg in _flatten_positions(portfolio.positions):
        if commodity_from_contract(leg.contract, fallback=None) == commodity:
            legs_by_contract.setdefault(leg.contract, []).append(leg)

    rows = {}
    for contract in sorted(legs_by_contract, key=_contract_sort_key):
        contract_portfolio = Portfolio()
        for leg in legs_by_contract[contract]:
            contract_portfolio.add(leg)
        ladder = vol_ladder(
            contract_portfolio, market, contract=contract, low=low, high=high, step=step,
            with_theta_rho=True,
        )
        rows[contract] = ladder[metric].tolist()

    shifts = list(range(low, high + step, step))
    return pd.DataFrame.from_dict(rows, orient="index", columns=shifts)


class _ExpiredOption:
    """
    An option leg past its expiry inside a decay scenario. In the money at
    expiry (judged at the scenario's own price) it was exercised, so it now
    behaves like the futures position it became: worth F - K (call) or
    K - F (put) per unit, delta +/-1, no gamma. Otherwise it's worthless.
    Using that fixed exercise decision rather than max(F - K, 0) keeps the
    Greeks' price bumps from straddling the payoff's kink at the strike.
    """

    def __init__(self, leg, expiry_price):
        self.leg = leg
        self.contract = leg.contract
        self.quantity = leg.quantity
        sign = 1.0 if leg.option_type == "call" else -1.0
        self._sign = sign if sign * (expiry_price - leg.strike) > 0 else 0.0

    def price(self, market):
        return self.quantity * self._sign * (market.price_for(self.contract) - self.leg.strike)

    def shift_vol(self, dv):
        return self

    def shift_time(self, days):
        return self


def _decay_leg(leg, days, market):

    if isinstance(leg, AmericanOption) and leg.expiry_days - days <= 0:
        return _ExpiredOption(leg, market.price_for(leg.contract))
    return leg.shift_time(days)


def _decay_portfolio(portfolio, days, market):
    """Moves the whole book `days` forward; options that expire on the way are exercised or lapse."""

    decayed = Portfolio()
    for pos in portfolio.positions:
        if isinstance(pos, Structure):
            decayed.add(Structure([_decay_leg(leg, days, market) for leg in pos.legs], name=pos.name))
        else:
            decayed.add(_decay_leg(pos, days, market))
    return decayed


def decay_ladder(
    portfolio,
    market,
    contract=None,
    horizon=60,
    step=5
):
    """
    PnL and Greeks as calendar days pass (0 to `horizon`, every `step`
    days), prices and vols held where they are. Delta/Gamma/Vega are
    measured against `contract`, like the other ladders.
    """

    days_list = list(range(0, horizon + 1, step))
    if days_list[-1] != horizon:
        days_list.append(horizon)

    current_value = portfolio.value(market)
    rows = []

    for days in days_list:

        decayed = _decay_portfolio(portfolio, days, market) if days else portfolio
        value = decayed.value(market)

        rows.append({
            "Days Forward": days,
            "PnL": value - current_value,
            "Portfolio Value": value,
            "Delta": GreekEngine.delta(decayed, market, contract),
            "Gamma": GreekEngine.gamma(decayed, market, contract),
            **_theta_rho(decayed, market),
            "Vega": GreekEngine.vega(decayed, market, contract),
        })

    return pd.DataFrame(rows)


def commodity_decay_ladder(portfolio, market, commodity, horizon=60, step=5):
    """Decay ladder for one commodity, advancing only its positions through time."""

    commodity_portfolio = _portfolio_for_commodity(portfolio, commodity)
    contracts = sorted({leg.contract for leg in commodity_portfolio.positions})
    days_list = list(range(0, horizon + 1, step))
    if days_list[-1] != horizon:
        days_list.append(horizon)

    current_value = commodity_portfolio.value(market)
    rows = []
    for days in days_list:
        decayed = _decay_portfolio(commodity_portfolio, days, market) if days else commodity_portfolio
        value = decayed.value(market)
        greeks = _commodity_ladder_greeks(decayed, market, contracts)
        rows.append({
            "Days Forward": days,
            "PnL": value - current_value,
            "Portfolio Value": value,
            **greeks,
        })

    return pd.DataFrame(rows)


def decay_ladder_by_underlying(portfolio, market, commodity, metric, horizon=60, step=5):
    """Return one selected decay-ladder metric per underlying in a commodity."""

    valid_metrics = {"PnL", "Portfolio Value", "Delta", "Gamma", "Theta", "Rho", "Vega"}
    if metric not in valid_metrics:
        raise ValueError(f"Unknown decay ladder metric: {metric}")

    legs_by_contract = {}
    for leg in _flatten_positions(portfolio.positions):
        if commodity_from_contract(leg.contract, fallback=None) == commodity:
            legs_by_contract.setdefault(leg.contract, []).append(leg)

    rows = {}
    for contract in sorted(legs_by_contract, key=_contract_sort_key):
        contract_portfolio = Portfolio()
        for leg in legs_by_contract[contract]:
            contract_portfolio.add(leg)
        ladder = decay_ladder(
            contract_portfolio, market, contract=contract, horizon=horizon, step=step,
        )
        rows[contract] = ladder[metric].tolist()

    days_list = list(range(0, horizon + 1, step))
    if days_list[-1] != horizon:
        days_list.append(horizon)
    return pd.DataFrame.from_dict(rows, orient="index", columns=days_list)


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


def report_by_tenor(portfolio, market, contract=None, by_contract=False, commodity=None):
    """
    GreekEngine.report_dollars(), grouped by underlying futures contract.
    Each contract's futures row comes first, followed by its option rows in
    expiry order, with a "Total" row for the included positions. Structures
    are flattened to their individual legs first, so a calendar spread's two
    legs correctly land in different tenor rows. By default, Delta/Gamma/etc
    are with respect to `contract`, same as everywhere else. Same $k scale
    as GreekEngine.report_dollars, so it's directly comparable to the main
    Greek Summary table.

    by_contract=True buckets by contract as well: each futures month gets
    its own row, and options are split by contract and expiry. Each row's
    Greeks are then taken with respect to that row's own contract, and the
    Total row is the sum of the rows.

    When commodity is set, only that commodity's legs are included. Each
    row is measured against its own underlying contract, and Total sums the
    rows, representing parallel price/volatility moves across the commodity.
    """

    legs = _flatten_positions(portfolio.positions)

    if by_contract:
        return _report_by_contract_tenor(legs, market)

    if commodity is not None:
        commodity_legs = [
            leg for leg in legs
            if commodity_from_contract(leg.contract, fallback=None) == commodity
        ]
        return _report_by_contract_tenor(commodity_legs, market, descriptive_labels=True)

    buckets = {}

    for leg in legs:
        expiry = None if isinstance(leg, Future) else leg.expiry_days
        buckets.setdefault((leg.contract, expiry), Portfolio())
        buckets[(leg.contract, expiry)].add(leg)

    def sort_key(key):
        underlying, expiry = key
        contract_order = _contract_sort_key(underlying) if underlying else ("", "", 0)
        return contract_order, -1 if expiry is None else expiry

    rows = []

    for underlying, expiry in sorted(buckets, key=sort_key):
        if underlying:
            month_code = underlying[-3] if len(underlying) >= 3 else ""
            year_code = underlying[-2:] if len(underlying) >= 2 else ""
            month = MONTH_CODES.get(month_code)
            contract_month = f"{month} 20{year_code}" if month and year_code.isdigit() else underlying
        else:
            contract_month = "Unknown contract"

        if expiry is None:
            row_label = f"{contract_month} Futures ({underlying or 'unknown'})"
        else:
            option_month = _option_month_for_bucket(
                buckets[(underlying, expiry)], underlying, expiry
            )
            month_label = f"{option_month} " if option_month else ""
            row_label = f"{month_label}Options ({expiry}d, {underlying or 'unknown'})"

        report = GreekEngine.report_dollars(buckets[(underlying, expiry)], market, contract)
        report = {"Tenor": row_label, **report}

        rows.append(report)

    total = {"Tenor": "Total", **GreekEngine.report_dollars(portfolio, market, contract)}
    rows.append(total)

    return pd.DataFrame(rows).set_index("Tenor")


def report_by_tenor_groups(portfolio, market, commodity):
    """Return each underlying contract's subtotal and its individual tenor rows."""

    legs_by_contract = {}
    for leg in _flatten_positions(portfolio.positions):
        if commodity_from_contract(leg.contract, fallback=None) == commodity:
            legs_by_contract.setdefault(leg.contract, []).append(leg)

    groups = []
    for contract in sorted(legs_by_contract, key=_contract_sort_key):
        legs = legs_by_contract[contract]
        details = _report_by_contract_tenor(legs, market, descriptive_labels=True)
        details = details.drop(index="Total")
        subtotal = details.sum().to_dict()

        month_code = contract[-3] if contract and len(contract) >= 3 else ""
        year_code = contract[-2:] if contract and len(contract) >= 2 else ""
        month = MONTH_CODES.get(month_code)
        contract_month = f"{month} 20{year_code}" if month and year_code.isdigit() else contract

        groups.append({
            "contract": contract,
            "label": f"{contract_month} ({contract})",
            "subtotal": subtotal,
            "details": details,
        })

    return groups


def _contract_sort_key(contract):
    """Chronological within a commodity: ZCZ26 before ZCH27 (plain string order would put H27 first)."""

    return contract[:-3], contract[-2:], _CME_MONTH_TO_NUMBER.get(contract[-3], 0)


def _option_month_for_bucket(bucket, contract, expiry):
    for leg in bucket.positions:
        if isinstance(leg, AmericanOption) and getattr(leg, "option_month", None):
            return leg.option_month

    if contract:
        for candidate in option_expiry_candidates(contract):
            if candidate["days"] == expiry:
                return candidate.get("option_month")

    return None


def _report_by_contract_tenor(legs, market, descriptive_labels=False):

    buckets = {}
    for leg in legs:
        expiry = None if isinstance(leg, Future) else leg.expiry_days
        buckets.setdefault((leg.contract, expiry), Portfolio()).add(leg)

    def sort_key(key):
        contract, expiry = key
        return _contract_sort_key(contract), -1 if expiry is None else expiry

    rows = []
    for contract, expiry in sorted(buckets, key=sort_key):
        if descriptive_labels:
            month_code = contract[-3] if contract and len(contract) >= 3 else ""
            year_code = contract[-2:] if contract and len(contract) >= 2 else ""
            month = MONTH_CODES.get(month_code)
            contract_month = f"{month} 20{year_code}" if month and year_code.isdigit() else contract
            if expiry is None:
                label = f"{contract_month} Futures ({contract or 'unknown'})"
            else:
                option_month = _option_month_for_bucket(
                    buckets[(contract, expiry)], contract, expiry
                )
                month_label = f"{option_month} " if option_month else ""
                label = f"{month_label}Options ({expiry}d, {contract or 'unknown'})"
        else:
            label = f"{contract} Future" if expiry is None else f"{contract} {expiry}d"
        report = GreekEngine.report_dollars(buckets[(contract, expiry)], market, contract)
        rows.append({"Tenor": label, **report})

    df = pd.DataFrame(rows).set_index("Tenor")
    df.loc["Total"] = df.sum()
    return df


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
