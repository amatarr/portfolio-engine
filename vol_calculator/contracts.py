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
    "SRW Wheat": {
        "symbol": "ZW",
        "months": ["H", "K", "N", "U", "Z"],
        "multiplier": 50.0,
    },
    "Soybean": {
        "symbol": "ZS",
        "months": ["F", "H", "K", "N", "Q", "U", "X"],
        "multiplier": 50.0,
    },
}


def contract_code(commodity, month_code, year):
    """e.g. contract_code("Corn", "Z", 2026) -> 'ZCZ26'"""
    symbol = COMMODITIES[commodity]["symbol"]
    return f"{symbol}{month_code}{str(year)[-2:]}"


def multiplier_for(commodity):
    return COMMODITIES[commodity]["multiplier"]


def commodity_from_contract(contract, fallback="Unknown"):
    if not contract:
        return fallback

    for name, info in COMMODITIES.items():
        if contract.startswith(info["symbol"]):
            return name

    return fallback


def _last_business_day(year, month):

    last_day = calendar.monthrange(year, month)[1]
    d = date(year, month, last_day)

    while d.weekday() >= 5:  # Sat=5, Sun=6
        d -= timedelta(days=1)

    return d


def _option_expiry_for_named_month(year, month):
    """
    CME grain options (standard and serial alike) expire on the last
    Friday that's at least 2 business days before the last business day
    of the calendar month BEFORE the option's own named month -- e.g. a
    "November" option (whether that's the serial on a December future,
    or -- for a commodity that lists a November future -- the standard
    one) expires relative to October's last business day. Verified
    directly against a real CME-sourced expiry table (12 points, Corn,
    spanning Nov 2026 - Dec 2028) -- matches exactly for near-term
    months; can drift by up to a few calendar days for months further
    out, since this doesn't know CME's actual holiday calendar (only
    weekends), unlike the exchange's own published dates.
    """

    prior_month = month - 1 or 12
    prior_year = year if month > 1 else year - 1

    anchor = _last_business_day(prior_year, prior_month)

    d = anchor
    while d.weekday() != 4:  # Friday
        d -= timedelta(days=1)

    while _count_business_days_between(d, anchor) < 2:
        d -= timedelta(days=7)

    return d


def _count_business_days_between(a, b):
    """Business days strictly between two dates, a < b."""

    n = 0
    d = a + timedelta(days=1)

    while d < b:
        if d.weekday() < 5:
            n += 1
        d += timedelta(days=1)

    return n


def option_expiry_candidates(contract, today=None):
    """
    Every plausible option expiry for a given underlying future contract
    (e.g. "ZCZ26") that hasn't passed yet: the STANDARD option (same
    named month as the future itself) plus a SERIAL option for every
    calendar month between the previous listed future and this one (a
    future month further out can have more than one serial -- e.g.
    March corn has January and February serials, since the previous
    listed corn future is December).

    This can suggest a serial a bit further out than CME has actually
    listed yet for a far-dated future (the exchange only opens serials
    on a rolling basis) -- it's a suggestion to override if it doesn't
    match what's actually tradeable yet, same as any other Add Leg
    default.

    Returns a list of {"label", "option_month", "expiry_date", "days"}
    dicts, nearest first, or [] if the contract doesn't parse or every
    candidate has already expired. "days" is calendar days from today to
    expiry.
    """

    today = today or date.today()

    if len(contract) < 3:
        return []

    month_letter = contract[-3]
    year_part = contract[-2:]

    month_num = _CME_MONTH_TO_NUMBER.get(month_letter)
    if month_num is None or not year_part.isdigit():
        return []

    year = 2000 + int(year_part)

    commodity = commodity_from_contract(contract, fallback=None)
    months_cycle = COMMODITIES[commodity]["months"] if commodity else None
    listed_months = sorted(_CME_MONTH_TO_NUMBER[m] for m in months_cycle) if months_cycle else []

    # The gap months needing a serial option are whatever's strictly
    # between the previous listed future and this one. If none of this
    # commodity's listed months come earlier in the SAME calendar year
    # (e.g. March corn, where the previous listed future is December of
    # the prior year), the gap simply starts from January instead of
    # wrapping the month number past December.
    earlier_listed = [m for m in listed_months if m < month_num]
    previous_future_month = max(earlier_listed) if earlier_listed else 0

    serial_months = list(range(previous_future_month + 1, month_num))

    candidates = []

    for named_month, is_standard in [(m, False) for m in serial_months] + [(month_num, True)]:
        expiry_date = _option_expiry_for_named_month(year, named_month)
        days = (expiry_date - today).days

        if days <= 0:
            continue

        month_name = calendar.month_abbr[named_month]
        kind = "standard" if is_standard else "serial"
        candidates.append({
            "label": f"{month_name} {year} ({kind}) -- {expiry_date.isoformat()}, {days}d",
            "option_month": f"{month_name} {year}",
            "expiry_date": expiry_date,
            "days": days,
        })

    return candidates


def option_month_for_expiry(contract, expiry_date, today=None):
    """Return the named option month for an exact candidate expiry date."""

    for candidate in option_expiry_candidates(contract, today=today):
        if candidate["expiry_date"] == expiry_date:
            return candidate["option_month"]

    return None


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
