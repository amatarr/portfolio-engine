from vol_calculator import (
    Market,
    Portfolio,
    Future,
    AmericanOption,
    GreekEngine,
    spot_ladder,
    vol_ladder,
    plot_surface,
)

market = Market(
    futures_price=528.75,
    interest_rate=0.038
)

portfolio = Portfolio()

portfolio.add(Future(quantity=-1))

portfolio.add(
    AmericanOption(
        quantity=-5,
        strike=620,
        expiry_days=64,
        option_type="call",
        volatility=0.26189
    )
)

portfolio.add(
    AmericanOption(
        quantity=2,
        strike=530,
        expiry_days=64,
        option_type="call",
        volatility=0.2150
    )
)

if __name__ == "__main__":
    print(portfolio)
    print()
    GreekEngine.summary(portfolio, market)
    print()
    print(spot_ladder(portfolio, market, low=-15, high=15, step=3).round(3))
    print()
    print(vol_ladder(portfolio, market, low=-10, high=10, step=2).round(3))

    plot_surface(portfolio, market, metric="PnL")
    plot_surface(portfolio, market, metric="Delta")
    plot_surface(portfolio, market, metric="Vega")
    plot_surface(portfolio, market, metric="Vanna-V")
    plot_surface(portfolio, market, metric="Volga")
