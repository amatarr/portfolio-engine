from vol_calculator.instruments import AmericanOption, Structure


def call_spread(quantity, k_long, k_short, expiry_days, vol_long, vol_short, contract=None):

    if k_long >= k_short:
        raise ValueError("call_spread requires k_long < k_short")

    legs = [
        AmericanOption(
            quantity=quantity,
            strike=k_long,
            expiry_days=expiry_days,
            option_type="call",
            volatility=vol_long,
            contract=contract
        ),
        AmericanOption(
            quantity=-quantity,
            strike=k_short,
            expiry_days=expiry_days,
            option_type="call",
            volatility=vol_short,
            contract=contract
        ),
    ]

    return Structure(
        legs,
        name=f"CallSpread(qty={quantity}, K={k_long}/{k_short}, days={expiry_days})"
    )


def put_spread(quantity, k_long, k_short, expiry_days, vol_long, vol_short, contract=None):

    if k_long <= k_short:
        raise ValueError("put_spread requires k_long > k_short")

    legs = [
        AmericanOption(
            quantity=quantity,
            strike=k_long,
            expiry_days=expiry_days,
            option_type="put",
            volatility=vol_long,
            contract=contract
        ),
        AmericanOption(
            quantity=-quantity,
            strike=k_short,
            expiry_days=expiry_days,
            option_type="put",
            volatility=vol_short,
            contract=contract
        ),
    ]

    return Structure(
        legs,
        name=f"PutSpread(qty={quantity}, K={k_long}/{k_short}, days={expiry_days})"
    )


def straddle(quantity, strike, expiry_days, vol_call, vol_put, contract=None):

    legs = [
        AmericanOption(
            quantity=quantity,
            strike=strike,
            expiry_days=expiry_days,
            option_type="call",
            volatility=vol_call,
            contract=contract
        ),
        AmericanOption(
            quantity=quantity,
            strike=strike,
            expiry_days=expiry_days,
            option_type="put",
            volatility=vol_put,
            contract=contract
        ),
    ]

    return Structure(
        legs,
        name=f"Straddle(qty={quantity}, K={strike}, days={expiry_days})"
    )


def strangle(quantity, k_put, k_call, expiry_days, vol_put, vol_call, contract=None):

    if k_put >= k_call:
        raise ValueError("strangle requires k_put < k_call")

    legs = [
        AmericanOption(
            quantity=quantity,
            strike=k_put,
            expiry_days=expiry_days,
            option_type="put",
            volatility=vol_put,
            contract=contract
        ),
        AmericanOption(
            quantity=quantity,
            strike=k_call,
            expiry_days=expiry_days,
            option_type="call",
            volatility=vol_call,
            contract=contract
        ),
    ]

    return Structure(
        legs,
        name=f"Strangle(qty={quantity}, K={k_put}/{k_call}, days={expiry_days})"
    )


def butterfly(
    quantity,
    k1,
    k2,
    k3,
    expiry_days,
    option_type,
    vol1,
    vol2,
    vol3,
    ratio=(1, 2, 1),
    contract=None
):

    if not (k1 < k2 < k3):
        raise ValueError("butterfly requires k1 < k2 < k3")

    r1, r2, r3 = ratio

    legs = [
        AmericanOption(
            quantity=quantity * r1,
            strike=k1,
            expiry_days=expiry_days,
            option_type=option_type,
            volatility=vol1,
            contract=contract
        ),
        AmericanOption(
            quantity=-quantity * r2,
            strike=k2,
            expiry_days=expiry_days,
            option_type=option_type,
            volatility=vol2,
            contract=contract
        ),
        AmericanOption(
            quantity=quantity * r3,
            strike=k3,
            expiry_days=expiry_days,
            option_type=option_type,
            volatility=vol3,
            contract=contract
        ),
    ]

    return Structure(
        legs,
        name=f"Butterfly(qty={quantity}, K={k1}/{k2}/{k3}, days={expiry_days})"
    )


def collar(put_qty, call_qty, k_put, k_call, expiry_days, vol_put, vol_call, contract=None):

    if k_put >= k_call:
        raise ValueError("collar requires k_put < k_call")

    legs = [
        AmericanOption(
            quantity=put_qty,
            strike=k_put,
            expiry_days=expiry_days,
            option_type="put",
            volatility=vol_put,
            contract=contract
        ),
        AmericanOption(
            quantity=call_qty,
            strike=k_call,
            expiry_days=expiry_days,
            option_type="call",
            volatility=vol_call,
            contract=contract
        ),
    ]

    return Structure(
        legs,
        name=f"Collar(put_qty={put_qty}, call_qty={call_qty}, K={k_put}/{k_call}, days={expiry_days})"
    )


def calendar_spread(
    quantity,
    strike,
    near_expiry_days,
    far_expiry_days,
    option_type,
    vol_near,
    vol_far,
    contract=None,
    far_contract=None,
    far_strike=None
):
    """
    Short the near expiry, long the far one. The far leg can sit on a
    different (later) future -- e.g. short Dec options on ZCZ26, long Mar
    options on ZCH27 -- and at a different strike (a diagonal); both
    default to the near leg's.
    """

    if near_expiry_days >= far_expiry_days:
        raise ValueError("calendar_spread requires near_expiry_days < far_expiry_days")

    far_contract = far_contract or contract
    far_strike = strike if far_strike is None else far_strike

    legs = [
        AmericanOption(
            quantity=-quantity,
            strike=strike,
            expiry_days=near_expiry_days,
            option_type=option_type,
            volatility=vol_near,
            contract=contract
        ),
        AmericanOption(
            quantity=quantity,
            strike=far_strike,
            expiry_days=far_expiry_days,
            option_type=option_type,
            volatility=vol_far,
            contract=far_contract
        ),
    ]

    strikes = f"K={strike}" if far_strike == strike else f"K={strike}/{far_strike}"
    contracts = f", {contract}/{far_contract}" if far_contract != contract else ""

    return Structure(
        legs,
        name=f"CalendarSpread(qty={quantity}, {strikes}, days={near_expiry_days}/{far_expiry_days}{contracts})"
    )


def combo(legs, name=None):
    return Structure(list(legs), name=name)
