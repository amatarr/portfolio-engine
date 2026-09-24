from dataclasses import dataclass


@dataclass
class Market:
    futures_price: float
    interest_rate: float
    # $ per 1-cent/bu move per contract; default is the standard 5,000 bu CBOT
    # grain contract (futures_price is quoted in cents/bu, so 1c/bu x 5,000bu = $50).
    contract_multiplier: float = 50.0
