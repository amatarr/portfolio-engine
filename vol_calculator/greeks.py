import copy

from vol_calculator.instruments import Portfolio

VOL_POINT = 0.01

BASIS_POINT = 0.0001

SPOT_BUMP = 0.001

PERCENT = 0.01


def shift_portfolio_vols(portfolio, dv, contract=None):
    """
    Shifts vol by dv. If contract is given, only legs on that contract are
    shifted -- legs on other contracts pass through unchanged. contract=None
    shifts every leg (back-compat / single-contract books).
    """

    shocked = Portfolio()

    for pos in portfolio.positions:
        if contract is None or pos.contract == contract:
            shocked.add(pos.shift_vol(dv))
        else:
            shocked.add(pos)

    return shocked


def shift_time(portfolio, days):
    """Time decay applies to the whole book regardless of contract."""

    shocked = Portfolio()

    for pos in portfolio.positions:
        shocked.add(pos.shift_time(days))

    return shocked


def _bump_price(market, contract, ds):

    bumped = copy.deepcopy(market)
    base = bumped.price_for(contract)
    key = contract if contract is not None else next(iter(bumped.futures_prices))
    bumped.futures_prices[key] = base + ds
    return bumped


class GreekEngine:

    @staticmethod
    def delta(portfolio, market, contract=None):

        ds = market.price_for(contract) * SPOT_BUMP

        up_market = _bump_price(market, contract, ds)
        down_market = _bump_price(market, contract, -ds)

        return (
            portfolio.value(up_market)
            - portfolio.value(down_market)
        ) / (2 * ds)

    @staticmethod
    def gamma(portfolio, market, contract=None):

        ds = market.price_for(contract) * SPOT_BUMP

        up_market = _bump_price(market, contract, ds)
        down_market = _bump_price(market, contract, -ds)

        up_value = portfolio.value(up_market)
        mid_value = portfolio.value(market)
        down_value = portfolio.value(down_market)

        return (
            up_value
            - 2 * mid_value
            + down_value
        ) / (ds ** 2)

    @staticmethod
    def vega(portfolio, market, contract=None):

        dv = VOL_POINT

        up_portfolio = shift_portfolio_vols(portfolio, dv, contract)
        down_portfolio = shift_portfolio_vols(portfolio, -dv, contract)

        up_value = up_portfolio.value(market)
        down_value = down_portfolio.value(market)

        return (
            (up_value - down_value) / (2 * dv)
        ) * VOL_POINT

    @staticmethod
    def theta(portfolio, market):

        tomorrow = shift_time(
            portfolio,
            1
        )

        return (
            tomorrow.value(market)
            - portfolio.value(market)
        )

    @staticmethod
    def rho(portfolio, market):

        dr = BASIS_POINT

        up_market = copy.deepcopy(market)
        down_market = copy.deepcopy(market)

        up_market.interest_rate += dr
        down_market.interest_rate -= dr

        return (
            (portfolio.value(up_market)
            - portfolio.value(down_market))
            / (2 * dr)
        ) * PERCENT

    @staticmethod
    def volga(portfolio, market, contract=None):

        dv = VOL_POINT

        up_portfolio = shift_portfolio_vols(portfolio, dv, contract)
        down_portfolio = shift_portfolio_vols(portfolio, -dv, contract)

        up_value = up_portfolio.value(market)
        mid_value = portfolio.value(market)
        down_value = down_portfolio.value(market)

        return (
            (
                up_value
                - 2 * mid_value
                + down_value
            ) / (dv ** 2)
        ) * VOL_POINT

    @staticmethod
    def vanna_vega(portfolio, market, contract=None):

        ds = market.price_for(contract) * SPOT_BUMP

        up_market = _bump_price(market, contract, ds)
        down_market = _bump_price(market, contract, -ds)

        vega_up = GreekEngine.vega(portfolio, up_market, contract)
        vega_down = GreekEngine.vega(portfolio, down_market, contract)

        return (
            vega_up
            - vega_down
        ) / (2 * ds)

    @staticmethod
    def vanna_delta(portfolio, market, contract=None):

        dv = VOL_POINT

        up_portfolio = shift_portfolio_vols(portfolio, dv, contract)
        down_portfolio = shift_portfolio_vols(portfolio, -dv, contract)

        delta_up = GreekEngine.delta(up_portfolio, market, contract)
        delta_down = GreekEngine.delta(down_portfolio, market, contract)

        return (
            delta_up
            - delta_down
        ) / (2 * dv)

    @staticmethod
    def summary(portfolio, market, contract=None):

        report = GreekEngine.report(
            portfolio,
            market,
            contract
        )

        for k, v in report.items():

            if k == "Value":
                print(f"{k:<25}: {v:.2f}")
            else:
                print(f"{k:<25}: {v:.4f}")

    @staticmethod
    def report(portfolio, market, contract=None):

        return {
            "Value": portfolio.value(market),

            "Delta (ΔPnL/¢)": GreekEngine.delta(portfolio, market, contract),

            "Gamma (ΔDelta/¢)": GreekEngine.gamma(portfolio, market, contract),

            "Vega (ΔPnL/1 vol pt)": GreekEngine.vega(portfolio, market, contract),

            "Theta (ΔPnL/day)": GreekEngine.theta(portfolio, market),

            "Rho (ΔPnL/1% rate)": GreekEngine.rho(portfolio, market),

            "Volga (ΔVega/1 vol pt)": GreekEngine.volga(portfolio, market, contract),

            "Vanna-D (ΔDelta/1 vol pt)": GreekEngine.vanna_delta(portfolio, market, contract),

            "Vanna-V (ΔVega/¢)": GreekEngine.vanna_vega(portfolio, market, contract)
        }

    @staticmethod
    def summary_dollars(portfolio, market, contract=None):

        report = GreekEngine.report_dollars(
            portfolio,
            market,
            contract
        )

        for k, v in report.items():
            print(f"{k:<28}: {v:,.4f}")

    @staticmethod
    def report_dollars(portfolio, market, contract=None):

        to_k = market.contract_multiplier / 1000

        return {
            "Value ($k)": portfolio.value(market) * to_k,

            "Delta (contracts)": GreekEngine.delta(portfolio, market, contract),

            "Gamma (ΔDelta contracts/1¢/bu)": GreekEngine.gamma(portfolio, market, contract),

            "Vega ($k/1% vol move)": GreekEngine.vega(portfolio, market, contract) * to_k,

            "Theta ($k/1 day decay)": GreekEngine.theta(portfolio, market) * to_k,

            "Rho ($k/1% rate move)": GreekEngine.rho(portfolio, market) * to_k,

            "Volga (ΔVega/1 vol pt)": GreekEngine.volga(portfolio, market, contract),

            "Vanna-D (ΔDelta/1 vol pt)": GreekEngine.vanna_delta(portfolio, market, contract),

            "Vanna-V (ΔVega/¢)": GreekEngine.vanna_vega(portfolio, market, contract)
        }
