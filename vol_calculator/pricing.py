import QuantLib as ql


class AmericanOptionPricer:

    def __init__(
        self,
        futures_price,
        strike,
        rate,
        volatility,
        expiry_days,
        option_type
    ):

        self.F = futures_price
        self.K = strike
        self.r = rate
        self.vol = volatility
        self.days = expiry_days
        self.option_type = option_type.lower()

    def price(self):

        calc_date = ql.Date.todaysDate()
        ql.Settings.instance().evaluationDate = calc_date

        expiry = calc_date + self.days

        payoff = ql.PlainVanillaPayoff(
            ql.Option.Call if self.option_type == "call"
            else ql.Option.Put,
            self.K
        )

        exercise = ql.AmericanExercise(
            calc_date,
            expiry
        )

        option = ql.VanillaOption(
            payoff,
            exercise
        )

        risk_free = ql.FlatForward(
            calc_date,
            self.r,
            ql.Actual365Fixed()
        )

        vol_curve = ql.BlackConstantVol(
            calc_date,
            ql.NullCalendar(),
            self.vol,
            ql.Actual365Fixed()
        )

        process = ql.BlackProcess(
            ql.QuoteHandle(
                ql.SimpleQuote(self.F)
            ),
            ql.YieldTermStructureHandle(
                risk_free
            ),
            ql.BlackVolTermStructureHandle(
                vol_curve
            )
        )

        engine = ql.BaroneAdesiWhaleyApproximationEngine(
            process
        )

        option.setPricingEngine(engine)

        return option.NPV()
