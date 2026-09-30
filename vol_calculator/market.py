from dataclasses import dataclass, field


@dataclass
class Market:
    # One price per contract, e.g. {"ZCZ26": 528.75, "ZSF27": 1340.0}. A
    # portfolio can hold positions on several contracts at once; each
    # instrument looks up its own price by its own .contract via price_for().
    futures_prices: dict = field(default_factory=dict)
    interest_rate: float = 0.0
    # $ per 1-cent/bu move per contract; default is the standard 5,000 bu CBOT
    # grain contract (futures_price is quoted in cents/bu, so 1c/bu x 5,000bu = $50).
    contract_multiplier: float = 50.0

    def price_for(self, contract):

        if contract is None:
            if len(self.futures_prices) == 1:
                return next(iter(self.futures_prices.values()))
            raise ValueError(
                "Instrument has no contract set and Market holds more than one "
                "(or zero) contract prices -- can't infer which price to use."
            )

        if contract not in self.futures_prices:
            raise KeyError(f"No price set for contract {contract!r}")

        return self.futures_prices[contract]
