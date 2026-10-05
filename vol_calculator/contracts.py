import calendar
from datetime import date, timedelta

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

# Full CME month-letter -> month-number map (all 12), for date arithmetic --
# MONTH_CODES above only lists the months these commodities actually trade.
_CME_MONTH_TO_NUMBER = {
    "F": 1, "G": 2, "H": 3, "J": 4, "K": 5, "M": 6,
    "N": 7, "Q": 8, "U": 9, "V": 10, "X": 11, "Z": 12,
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


def approx_option_expiry_days(contract, today=None):
    """
    Rough stand-in for a grain option's real expiry date. CME's actual
    rule ("last Friday that's at least 2 business days before first
    notice day") shifts year to year and isn't derivable from the symbol
    alone -- this approximates it as the last business day of the month
    before the futures contract month, which is close enough for Greeks/
    scenario purposes until exact per-contract dates are wired in.

    contract: e.g. "ZCZ26" (2-letter root + month letter + 2-digit year).
    Returns days from today to that approximate date, floored at 1, or
    None if the contract code doesn't parse.
    """

    today = today or date.today()

    if len(contract) < 3:
        return None

    month_letter = contract[-3]
    year_part = contract[-2:]

    month_num = _CME_MONTH_TO_NUMBER.get(month_letter)

    if month_num is None or not year_part.isdigit():
        return None

    year = 2000 + int(year_part)

    prior_month = month_num - 1 or 12
    prior_year = year if month_num > 1 else year - 1

    last_day = calendar.monthrange(prior_year, prior_month)[1]
    expiry_date = date(prior_year, prior_month, last_day)

    while expiry_date.weekday() >= 5:  # Sat=5, Sun=6
        expiry_date -= timedelta(days=1)

    return max((expiry_date - today).days, 1)
