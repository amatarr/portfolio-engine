import copy

from vol_calculator.instruments import Portfolio

VOL_POINT = 0.01

BASIS_POINT = 0.0001

SPOT_BUMP = 0.001

PERCENT = 0.01


def shift_portfolio_vols(portfolio, dv):

    shocked = Portfolio()

    for pos in portfolio.positions:
        shocked.add(pos.shift_vol(dv))

    return shocked


def shift_time(portfolio, days):

    shocked = Portfolio()

    for pos in portfolio.positions:
        shocked.add(pos.shift_time(days))

    return shocked


class GreekEngine:

    @staticmethod
    def delta(portfolio, market):

        ds = market.futures_price * SPOT_BUMP

        up_market = copy.deepcopy(market)
        down_market = copy.deepcopy(market)

        up_market.futures_price += ds
        down_market.futures_price -= ds

        return (
            portfolio.value(up_market)
            - portfolio.value(down_market)
        ) / (2 * ds)

    @staticmethod
    def gamma(portfolio, market):

        ds = market.futures_price * SPOT_BUMP

        up_market = copy.deepcopy(market)
        down_market = copy.deepcopy(market)

        up_market.futures_price += ds
        down_market.futures_price -= ds

        up_value = portfolio.value(up_market)
        mid_value = portfolio.value(market)
        down_value = portfolio.value(down_market)

        return (
            up_value
            - 2 * mid_value
            + down_value
        ) / (ds ** 2)

    @staticmethod
    def vega(portfolio, market):

        dv = VOL_POINT

        up_portfolio = shift_portfolio_vols(portfolio, dv)
        down_portfolio = shift_portfolio_vols(portfolio, -dv)

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
    def volga(portfolio, market):

        dv = VOL_POINT

        up_portfolio = shift_portfolio_vols(
            portfolio,
            dv
        )

        down_portfolio = shift_portfolio_vols(
            portfolio,
            -dv
        )

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
    def vanna_vega(portfolio, market):

        ds = market.futures_price * SPOT_BUMP

        up_market = copy.deepcopy(market)
        down_market = copy.deepcopy(market)

        up_market.futures_price += ds
        down_market.futures_price -= ds

        vega_up = GreekEngine.vega(
            portfolio,
            up_market
        )

        vega_down = GreekEngine.vega(
            portfolio,
            down_market
        )

        return (
            vega_up
            - vega_down
        ) / (2 * ds)

    @staticmethod
    def vanna_delta(portfolio, market):

        dv = VOL_POINT

        up_portfolio = shift_portfolio_vols(
            portfolio,
            dv
        )

        down_portfolio = shift_portfolio_vols(
            portfolio,
            -dv
        )

        delta_up = GreekEngine.delta(
            up_portfolio,
            market
        )

        delta_down = GreekEngine.delta(
            down_portfolio,
            market
        )

        return (
            delta_up
            - delta_down
        ) / (2 * dv)

    @staticmethod
    def summary(portfolio, market):

        report = GreekEngine.report(
            portfolio,
            market
        )

        for k, v in report.items():

            if k == "Value":
                print(f"{k:<25}: {v:.2f}")
            else:
                print(f"{k:<25}: {v:.4f}")

    @staticmethod
    def report(portfolio, market):

        return {
            "Value": portfolio.value(market),

            "Delta (ΔPnL/pt)": GreekEngine.delta(
                portfolio,
                market
            ),

            "Gamma (ΔDelta/pt)": GreekEngine.gamma(
                portfolio,
                market
            ),

            "Vega (ΔPnL/1 vol pt)": GreekEngine.vega(
                portfolio,
                market
            ),

            "Theta (ΔPnL/day)": GreekEngine.theta(
                portfolio,
                market
            ),

            "Rho (ΔPnL/1% rate)": GreekEngine.rho(
                portfolio,
                market
            ),

            "Volga (ΔVega/1 vol pt)": GreekEngine.volga(
                portfolio,
                market
            ),

            "Vanna-D (ΔDelta/1 vol pt)": GreekEngine.vanna_delta(
                portfolio,
                market
            ),

            "Vanna-V (ΔVega/pt)": GreekEngine.vanna_vega(
                portfolio,
                market
            )
        }

    @staticmethod
    def summary_dollars(portfolio, market):

        report = GreekEngine.report_dollars(
            portfolio,
            market
        )

        for k, v in report.items():
            print(f"{k:<28}: {v:,.4f}")

    @staticmethod
    def report_dollars(portfolio, market):

        to_k = market.contract_multiplier / 1000

        return {
            "Value ($k)": portfolio.value(market) * to_k,

            "Delta (contracts)": GreekEngine.delta(
                portfolio,
                market
            ),

            "Gamma (ΔDelta contracts/1¢/bu)": GreekEngine.gamma(
                portfolio,
                market
            ),

            "Vega ($k/1% vol move)": GreekEngine.vega(
                portfolio,
                market
            ) * to_k,

            "Theta ($k/1 day decay)": GreekEngine.theta(
                portfolio,
                market
            ) * to_k,

            "Rho ($k/1% rate move)": GreekEngine.rho(
                portfolio,
                market
            ) * to_k,

            "Volga (ΔVega/1 vol pt)": GreekEngine.volga(
                portfolio,
                market
            ),

            "Vanna-D (ΔDelta/1 vol pt)": GreekEngine.vanna_delta(
                portfolio,
                market
            ),

            "Vanna-V (ΔVega/pt)": GreekEngine.vanna_vega(
                portfolio,
                market
            )
        }
