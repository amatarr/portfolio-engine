from vol_calculator.pricing import AmericanOptionPricer


class Portfolio:

    def __init__(self):
        self.positions = []

    def add(self, instrument):
        self.positions.append(instrument)

    def value(self, market):

        total = 0

        for position in self.positions:
            total += position.price(market)

        return total

    def __repr__(self):

        output = ["Portfolio:"]

        for p in self.positions:
            output.append(f"  {p}")

        return "\n".join(output)


class Future:

    def __init__(self, quantity, contract=None):
        self.quantity = quantity
        self.contract = contract

    def price(self, market):
        return self.quantity * market.price_for(self.contract)

    def shift_vol(self, dv):
        return self

    def shift_time(self, days):
        return self

    def __repr__(self):
        return f"Future(qty={self.quantity}, contract={self.contract})"


class AmericanOption:

    def __init__(
        self,
        quantity,
        strike,
        expiry_days,
        option_type,
        volatility,
        contract=None,
        option_month=None,
        expiry_date=None,
        entry_price=None,
        entry_volatility=None,
        reference_price=None,
    ):
        self.quantity = quantity
        self.strike = strike
        self.expiry_days = expiry_days
        self.option_type = option_type
        self.volatility = volatility
        self.contract = contract
        self.option_month = option_month
        # Calendar expiry date, when known. Pricing only uses expiry_days;
        # the Streamlit app re-derives expiry_days from this date each day.
        self.expiry_date = expiry_date
        # What this leg was actually traded at, when entered per leg (the
        # legs of a structure). Never used for pricing.
        self.entry_price = entry_price
        # Original volatility retained as a reference when current volatility changes.
        self.entry_volatility = volatility if entry_volatility is None else entry_volatility
        self.reference_price = reference_price

    def price(self, market):

        pricer = AmericanOptionPricer(
            futures_price=market.price_for(self.contract),
            strike=self.strike,
            rate=market.interest_rate,
            volatility=self.volatility,
            expiry_days=self.expiry_days,
            option_type=self.option_type
        )

        return self.quantity * pricer.price()

    def shift_vol(self, dv):
        return AmericanOption(
            quantity=self.quantity,
            strike=self.strike,
            expiry_days=self.expiry_days,
            option_type=self.option_type,
            volatility=self.volatility + dv,
            contract=self.contract,
            option_month=self.option_month,
            expiry_date=self.expiry_date,
            entry_volatility=self.entry_volatility,
            reference_price=self.reference_price,
        )

    def shift_time(self, days):
        return AmericanOption(
            quantity=self.quantity,
            strike=self.strike,
            expiry_days=max(self.expiry_days - days, 1),
            option_type=self.option_type,
            volatility=self.volatility,
            contract=self.contract,
            option_month=self.option_month,
            entry_volatility=self.entry_volatility,
            reference_price=self.reference_price,
        )

    def __repr__(self):

        return (
            f"AmericanOption("
            f"qty={self.quantity}, "
            f"type='{self.option_type}', "
            f"strike={self.strike}, "
            f"days={self.expiry_days}, "
            f"vol={self.volatility:.2%}, "
            f"contract={self.contract})"
        )


class ExpiredOption:
    """
    An option leg moved past its expiry, for Greeks. In the money at expiry
    (judged at `expiry_price`) it was exercised, so it now behaves like the
    futures position it became: worth F - K (call) or K - F (put) per unit,
    delta +/-1, no gamma. Otherwise it's worthless. Using that fixed exercise
    decision rather than max(F - K, 0) keeps the Greeks' price bumps from
    straddling the payoff's kink at the strike.
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


class IntrinsicOption:
    """
    An option leg moved past its expiry, for P&L curves: worth its intrinsic
    value, max(F - K, 0) or max(K - F, 0), at whatever price it is valued.
    """

    def __init__(self, leg):
        self.leg = leg
        self.contract = leg.contract
        self.quantity = leg.quantity

    def price(self, market):
        spot = market.price_for(self.contract)
        if self.leg.option_type.lower() == "call":
            return self.quantity * max(spot - self.leg.strike, 0)
        return self.quantity * max(self.leg.strike - spot, 0)

    def shift_vol(self, dv):
        return self

    def shift_time(self, days):
        return self


class Structure:

    def __init__(self, legs, name=None):
        self.legs = legs
        self.name = name

    @property
    def contract(self):
        contracts = {leg.contract for leg in self.legs}
        return next(iter(contracts)) if len(contracts) == 1 else None

    def price(self, market):
        return sum(leg.price(market) for leg in self.legs)

    def shift_vol(self, dv):
        return Structure(
            [leg.shift_vol(dv) for leg in self.legs],
            name=self.name
        )

    def shift_time(self, days):
        return Structure(
            [leg.shift_time(days) for leg in self.legs],
            name=self.name
        )

    def __repr__(self):

        header = self.name if self.name else "Structure"
        output = [f"{header}:"]

        for leg in self.legs:
            output.append(f"  {leg}")

        return "\n".join(output)
