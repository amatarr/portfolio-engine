# Reference data for the CME grain contracts this tool currently supports.
# All three are standard 5,000 bu contracts, so the $/pt multiplier is the
# same for each (1 cent/bu x 5,000 bu = $50) -- see Market.contract_multiplier.
#
# "months" lists the contract (delivery) months each commodity actually
# lists on CME, using standard futures month codes:
# F=Jan H=Mar K=May N=Jul Q=Aug U=Sep X=Nov Z=Dec

MONTH_CODES = {
    "F": "Jan", "H": "Mar", "K": "May", "N": "Jul",
    "Q": "Aug", "U": "Sep", "X": "Nov", "Z": "Dec",
}

COMMODITIES = {
    "Corn": {
        "symbol": "ZC",
        "months": ["H", "K", "N", "U", "Z"],
        "multiplier": 50.0,
    },
    "Soybean": {
        "symbol": "ZS",
        "months": ["F", "H", "K", "N", "Q", "U", "X"],
        "multiplier": 50.0,
    },
    "SRW Wheat": {
        "symbol": "ZW",
        "months": ["H", "K", "N", "U", "Z"],
        "multiplier": 50.0,
    },
}


def contract_code(commodity, month_code, year):
    """e.g. contract_code("Corn", "Z", 2026) -> 'ZCZ26'"""
    symbol = COMMODITIES[commodity]["symbol"]
    return f"{symbol}{month_code}{str(year)[-2:]}"


def multiplier_for(commodity):
    return COMMODITIES[commodity]["multiplier"]
