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
        contract=None
    ):
        self.quantity = quantity
        self.strike = strike
        self.expiry_days = expiry_days
        self.option_type = option_type
        self.volatility = volatility
        self.contract = contract

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
            contract=self.contract
        )

    def shift_time(self, days):
        return AmericanOption(
            quantity=self.quantity,
            strike=self.strike,
            expiry_days=max(self.expiry_days - days, 1),
            option_type=self.option_type,
            volatility=self.volatility,
            contract=self.contract
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
