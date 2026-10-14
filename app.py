import re
import tempfile
from datetime import datetime, timezone

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st

from vol_calculator import (
    Market,
    Portfolio,
    Future,
    AmericanOption,
    Structure,
    GreekEngine,
    call_spread,
    put_spread,
    straddle,
    strangle,
    butterfly,
    collar,
    calendar_spread,
    spot_ladder,
    vol_ladder,
    report_by_tenor,
    breakevens_quadratic,
    breakevens_numerical,
    plot_surface,
    plot_strike_map,
    strike_map,
    plot_stress_report,
    plot_payoff,
    portfolio_to_csv,
    portfolio_from_csv,
    fetch_latest_sofr,
    COMMODITIES,
    approx_option_expiry_days,
    import_bushel_positions,
    import_cqg_quotes,
    compute_model_price,
    daily_prop_as_of_date,
    option_expiry_candidates,
)

st.set_page_config(page_title="Vol Calculator", layout="wide")

st.markdown(
    """
    <style>
    div[data-testid="stNumberInput"] { max-width: 170px; }
    div[data-testid="stTextInput"] { max-width: 240px; }
    div[data-testid="stSelectbox"] { max-width: 240px; }

    /* Quotes box: tight rows (label + price) instead of normal widget
       spacing, so a dozen contracts don't take a dozen widgets' worth
       of vertical room. */
    .st-key-cqg_quotes_box div[data-testid="stVerticalBlockBorderWrapper"],
    .st-key-cqg_quotes_box div[data-testid="stHorizontalBlock"] {
        margin-bottom: -0.6rem;
    }
    .st-key-cqg_quotes_box div[data-testid="stNumberInput"] input {
        padding-top: 0.25rem; padding-bottom: 0.25rem; height: 1.8rem;
    }

    .st-key-add_leg_btn button {
        background-color: #2e7d32; color: white; border-color: #1b5e20;
    }
    .st-key-add_leg_btn button:hover {
        background-color: #388e3c; border-color: #1b5e20; color: white;
    }
    .st-key-clear_btn button {
        background-color: #c62828; color: white; border-color: #8e0000;
    }
    .st-key-clear_btn button:hover {
        background-color: #d32f2f; border-color: #8e0000; color: white;
    }
    .st-key-file_btn_load button, .st-key-file_btn_download button, .st-key-file_btn_import button {
        background-color: #000000; color: white; border-color: #000000;
    }
    .st-key-file_btn_load button:hover, .st-key-file_btn_download button:hover, .st-key-file_btn_import button:hover {
        background-color: #262626; border-color: #000000; color: white;
    }
    .st-key-daily_prop_upload_box {
        background-color: #fff9c4; border-radius: 8px; padding: 0.5rem 0.75rem;
    }
    .st-key-cqg_quotes_box {
        background-color: #cfe8fc; border-radius: 8px; padding: 0.5rem 0.75rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(ttl=4 * 60 * 60)
def _cached_sofr():
    return fetch_latest_sofr()


_CQG_QUOTES_PATH = "CQG LINKS SHEET_.xlsx"
_SUPABASE_QUOTES_TABLE = "Quotes"


@st.cache_data(ttl=4)
def _fetch_supabase_quotes(url, anon_key):

    from supabase import create_client

    client = create_client(url, anon_key)
    resp = client.table(_SUPABASE_QUOTES_TABLE).select("*").execute()

    return [
        {
            "commodity": _commodity_from_contract(row["contract"]),
            "contract": row["contract"],
            "price": row["price"],
            "updated_at": row.get("updated_at"),
        }
        for row in resp.data
    ]


def _load_live_quotes():
    """
    Prefers Supabase (fed by push_quotes_supabase.py reading live CQGXL
    cells off the open workbook) -- works both locally and on the
    deployed app, since it's a normal outbound read, not a local file.
    Falls back to the local CQGXL export file (local-only), then to
    empty if neither is available.
    """

    supabase_url = st.secrets.get("supabase_url")
    supabase_anon_key = st.secrets.get("supabase_anon_key")

    if supabase_url and supabase_anon_key:
        try:
            quotes = _fetch_supabase_quotes(supabase_url, supabase_anon_key)
            if quotes:
                return quotes
        except Exception:
            pass

    try:
        return import_cqg_quotes(_CQG_QUOTES_PATH)
    except FileNotFoundError:
        return []


def _require_password():

    if st.session_state.get("authenticated"):
        return

    configured_password = st.secrets.get("app_password")

    if not configured_password:
        st.warning(
            "No app password is configured -- running unlocked. Set "
            "`app_password` in Streamlit secrets (Settings > Secrets in "
            "Streamlit Cloud, or .streamlit/secrets.toml locally) to require "
            "a login."
        )
        st.session_state.authenticated = True
        return

    st.title("Vol Calculator")
    password = st.text_input("Password", type="password")

    if st.button("Enter"):
        if password == configured_password:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password.")

    if not st.session_state.get("authenticated"):
        st.stop()


_require_password()


LEG_BUILDERS = {
    "Future": {
        "defaults": {"quantity": -1.0},
        "build": lambda **kw: Future(**kw),
    },
    "American Option": {
        "defaults": {
            "quantity": 1.0, "strike": 530.0, "expiry_days": 64,
            "option_type": "call", "volatility": 0.20,
        },
        "build": lambda **kw: AmericanOption(**kw),
    },
    "Call Spread": {
        "defaults": {
            "quantity": 1.0, "k_long": 520.0, "k_short": 560.0,
            "expiry_days": 64, "vol_long": 0.22, "vol_short": 0.20,
        },
        "build": lambda **kw: call_spread(**kw),
    },
    "Put Spread": {
        "defaults": {
            "quantity": 1.0, "k_long": 540.0, "k_short": 500.0,
            "expiry_days": 64, "vol_long": 0.23, "vol_short": 0.25,
        },
        "build": lambda **kw: put_spread(**kw),
    },
    "Straddle": {
        "defaults": {
            "quantity": 1.0, "strike": 528.75, "expiry_days": 64,
            "vol_call": 0.21, "vol_put": 0.215,
        },
        "build": lambda **kw: straddle(**kw),
    },
    "Strangle": {
        "defaults": {
            "quantity": 1.0, "k_put": 500.0, "k_call": 560.0,
            "expiry_days": 64, "vol_put": 0.24, "vol_call": 0.20,
        },
        "build": lambda **kw: strangle(**kw),
    },
    "Butterfly": {
        "defaults": {
            "quantity": 1.0, "k1": 500.0, "k2": 530.0, "k3": 560.0,
            "expiry_days": 64, "option_type": "call",
            "vol1": 0.24, "vol2": 0.21, "vol3": 0.19,
        },
        "build": lambda **kw: butterfly(**kw),
    },
    "Collar": {
        "defaults": {
            "put_qty": 1.0, "call_qty": -1.0, "k_put": 500.0, "k_call": 560.0,
            "expiry_days": 64, "vol_put": 0.24, "vol_call": 0.19,
        },
        "build": lambda **kw: collar(**kw),
    },
    "Calendar Spread": {
        "defaults": {
            "quantity": 1.0, "strike": 530.0, "near_expiry_days": 30,
            "far_expiry_days": 90, "option_type": "call",
            "vol_near": 0.20, "vol_far": 0.22,
        },
        "build": lambda **kw: calendar_spread(**kw),
    },
}


def _render_field(container, name, default):

    if name == "option_type":
        options = ["call", "put"]
        return container.selectbox(name, options, index=options.index(default))

    if "days" in name:
        return int(container.number_input(name, value=int(default), step=1))

    if "vol" in name:
        pct = container.number_input(
            f"{name} (%)", value=float(default) * 100, step=0.1, format="%.2f"
        )
        return pct / 100

    return container.number_input(name, value=float(default), step=1.0)


def _leg_label(pos):

    if isinstance(pos, Structure):
        return pos.name or "Structure"

    return repr(pos)


def _structure_kind(name):

    if not name:
        return "Structure"

    kind = name.split("(")[0]
    return re.sub(r"(?<!^)(?=[A-Z])", " ", kind)


def _leg_row(leg, position_index, structure, contract, stored_price, market, future_vol_pct=None):
    """
    stored_price is whatever meta["settlement_price"] holds -- for a
    Future that IS its Underlying Price (there's no separate Entry
    Price for a Future); for an Option it's the Entry Price (CSV
    settlement, or a manual fill price typed in at Add Leg).
    """

    if isinstance(leg, Future):
        return {
            "#": position_index, "Contract": contract, "Entry Price": None,
            "Structure": structure, "Type": "Future",
            "Quantity": leg.quantity, "Strike": None, "Expiry (d)": None, "Vol (%)": future_vol_pct,
            "Underlying Price": stored_price, "Mark to Market": None,
        }

    # Underlying Price/Mark to Market are always recomputed live from the
    # current market, never copied from Entry Price -- that's the whole
    # point of distinguishing them.
    try:
        underlying_price = market.price_for(contract)
    except (KeyError, ValueError):
        underlying_price = None

    mark_to_market = compute_model_price(leg, {"contract": contract}, market)

    return {
        "#": position_index, "Contract": contract, "Entry Price": stored_price,
        "Structure": structure,
        "Type": "Call" if leg.option_type == "call" else "Put",
        "Quantity": leg.quantity, "Strike": leg.strike,
        "Expiry (d)": leg.expiry_days, "Vol (%)": leg.volatility * 100,
        "Underlying Price": underlying_price, "Mark to Market": mark_to_market,
    }


def _positions_table_for_commodity(portfolio, meta, commodity, market):
    """meta["settlement_price"] is whatever was captured at import/add time -- a fixed snapshot, not tied to the live feed. _leg_row routes it to Underlying Price (Future) or Entry Price (Option). Underlying Price/Mark to Market for options are recomputed live every render."""

    rows = []

    for i, pos in enumerate(portfolio.positions):

        if meta[i].get("commodity") != commodity:
            continue

        contract = meta[i].get("contract") or "-"
        stored_price = meta[i].get("settlement_price")

        if isinstance(pos, Structure):
            structure = _structure_kind(pos.name)
            rows.extend(_leg_row(leg, i, structure, contract, stored_price, market) for leg in pos.legs)
        else:
            rows.append(_leg_row(pos, i, "-", contract, stored_price, market, future_vol_pct=meta[i].get("vol_pct")))

    return pd.DataFrame(rows)


def _commodities_present(meta):

    seen = []
    for m in meta:
        c = m.get("commodity", "Unknown")
        if c not in seen:
            seen.append(c)
    return seen


def _default_portfolio():
    return Portfolio()


_AUTOSAVE_PATH = "_autosave_portfolio.csv"


def _autosave_portfolio():
    """
    Writes the current portfolio to disk so an accidental browser refresh
    (which resets st.session_state) doesn't lose it. Not a real database --
    just the last-known state, overwritten on every change.
    """
    try:
        portfolio_to_csv(st.session_state.portfolio, _AUTOSAVE_PATH)
    except Exception:
        pass


def _load_autosaved_portfolio():

    try:
        portfolio = portfolio_from_csv(_AUTOSAVE_PATH)
    except FileNotFoundError:
        return _default_portfolio(), []

    meta = [
        {"commodity": _commodity_from_contract(pos.contract), "contract": pos.contract, "vol_pct": None}
        for pos in portfolio.positions
    ]
    return portfolio, meta


def _commodity_from_contract(contract, fallback="Unknown"):

    if not contract:
        return fallback

    for name, info in COMMODITIES.items():
        if contract.startswith(info["symbol"]):
            return name

    return fallback


def _download_df_button(df, label, filename):
    st.download_button(label, data=df.to_csv().encode("utf-8"), file_name=filename, mime="text/csv")


def _format_last_updated(dt):
    """
    Seconds/minutes while recent, hours+minutes once past an hour,
    "Yesterday at HH:MM" for a 1-2 day gap, and the exact date beyond
    that -- so a stale feed reads as increasingly alarming instead of
    just a climbing raw second count.
    """

    seconds = int((datetime.now(timezone.utc) - dt).total_seconds())

    if seconds < 60:
        return f"{seconds}s ago"

    if seconds < 3600:
        return f"{seconds // 60}m ago"

    if seconds < 86400:
        hours, minutes = divmod(seconds // 60, 60)
        return f"{hours}h {minutes}m ago"

    if seconds < 2 * 86400:
        return f"Yesterday at {dt.strftime('%H:%M UTC')}"

    return dt.strftime("%Y-%m-%d %H:%M:%S UTC")


def _live_quote_contracts_by_commodity(live_quotes):
    """
    Groups ONLY what's actually in live_quotes (the Supabase/CQGXL feed)
    by commodity -- for the Quotes box alone, which must never show a
    contract that only exists because it was imported or added by hand
    (e.g. a far-dated Daily Prop contract CQGXL doesn't quote).
    """

    contracts_by_commodity = {name: [] for name in COMMODITIES.keys()}

    for q in live_quotes:
        if q["commodity"] in contracts_by_commodity:
            contracts_by_commodity[q["commodity"]].append(q["contract"])

    return contracts_by_commodity


def _build_contracts_by_commodity(live_quotes):
    """
    Groups known contracts by commodity -- live quotes first, then any
    contract already tracked in Market that the live feed doesn't cover
    (e.g. manually added earlier), so it doesn't silently disappear. Used
    for the Add Leg dropdown and other places that need every contract
    in play, not just the live-quoted ones -- NOT for the Quotes box
    itself, which uses _live_quote_contracts_by_commodity instead.
    """

    contracts_by_commodity = {name: [] for name in COMMODITIES.keys()}

    for q in live_quotes:
        if q["commodity"] in contracts_by_commodity:
            contracts_by_commodity[q["commodity"]].append(q["contract"])

    for code in st.session_state.market.futures_prices.keys():
        comm = _commodity_from_contract(code, fallback=None)
        if comm in contracts_by_commodity and code not in contracts_by_commodity[comm]:
            contracts_by_commodity[comm].append(code)

    return contracts_by_commodity


@st.fragment(run_every="5s")
def _render_quotes_box():
    """
    Polls the live quotes source every 10s independently of the rest of
    the page -- a fragment rerun only re-executes this function, not the
    whole script, so the Add Leg form/other widgets don't reset while
    this ticks. Writes straight into st.session_state.market.futures_
    prices, same as a normal full rerun would.
    """

    live_quotes = _load_live_quotes()
    quote_price_lookup = {q["contract"]: q["price"] for q in live_quotes}
    contracts_by_commodity = _live_quote_contracts_by_commodity(live_quotes)

    with st.container(key="cqg_quotes_box", border=True):

        if not any(contracts_by_commodity.values()):
            st.caption("No contracts yet -- add a leg below to start one.")
            return

        timestamps = [q["updated_at"] for q in live_quotes if q.get("updated_at")]
        if timestamps:
            latest = max(timestamps)
            try:
                dt = datetime.fromisoformat(latest)
                st.caption(f"Last updated: {_format_last_updated(dt)}")
            except ValueError:
                st.caption(f"Last updated: {latest}")

        comm_cols = st.columns(len(contracts_by_commodity))

        def _price_for(code):
            # This box always shows the live feed, regardless of whether
            # the contract's calc price is pinned to a Daily Prop CSV
            # settlement -- the pin only affects the Positions table/
            # engine, never what's displayed here.
            return quote_price_lookup.get(
                code, st.session_state.market.futures_prices.get(code, 528.75)
            )

        for comm_col, (comm, codes) in zip(comm_cols, contracts_by_commodity.items()):
            with comm_col:
                st.markdown(f"**{comm.upper()}**")

                if not codes:
                    st.caption("-")
                    continue

                # One tight inline row per contract (label + editable
                # price) instead of a full data_editor grid -- no header
                # row or hover toolbar per commodity, just the numbers.
                for code in codes:
                    price = float(_price_for(code))
                    label_col, price_col = st.columns([1, 1])
                    label_col.markdown(f"<div style='padding-top:0.5rem'>{code}</div>", unsafe_allow_html=True)
                    # Keying on the live price forces Streamlit to treat
                    # a real tick as a new widget -- a number_input's
                    # value= is only its default on first creation, same
                    # reasoning as the data_editor it replaces.
                    edited_price = price_col.number_input(
                        code, value=price, step=0.25, format="%.2f",
                        label_visibility="collapsed", key=f"quote_price_{code}_{price:.4f}",
                    )

                    if code in st.session_state.csv_locked_contracts:
                        # Pinned contract: this box is informational only
                        # for it, never writes back into the engine's price.
                        continue
                    if code not in quote_price_lookup:
                        # No live quote for this contract -- leave its
                        # engine price exactly as-is rather than writing
                        # back the displayed fallback value every rerun.
                        continue
                    st.session_state.market.futures_prices[code] = float(edited_price)


if "portfolio" not in st.session_state:
    st.session_state.portfolio, st.session_state.position_meta = _load_autosaved_portfolio()

if "position_meta" not in st.session_state:
    st.session_state.position_meta = []

# Contracts whose price should stay pinned to a Daily Prop CSV's own
# settlement price, not get pulled back to the live feed -- populated at
# Bushel-import time.
if "csv_locked_contracts" not in st.session_state:
    st.session_state.csv_locked_contracts = set()

# Starting point for the deployed site, which has no CQGXL file to read --
# snapshot of the contracts/prices from the local sheet, editable afterward.
# Not "live," just a seed so the table isn't empty on first load.
_SEED_FUTURES_PRICES = {
    "ZCZ26": 500.75, "ZCH27": 515.75, "ZCK27": 522.75,
    "ZCN27": 527.25, "ZCU27": 510.25, "ZCZ27": 519.75,
    "ZSX26": 1291.25, "ZSF27": 1308.0, "ZSH27": 1317.75,
    "ZSK27": 1327.0, "ZSN27": 1333.75, "ZSQ27": 1316.25,
    "ZWZ26": 696.25, "ZWH27": 709.5, "ZWK27": 715.5,
    "ZWN27": 716.0, "ZWU27": 721.0, "ZWZ27": 728.5,
}

if "market" not in st.session_state:
    try:
        _sofr = _cached_sofr()
        _default_rate = _sofr["rate"]
    except Exception:
        _default_rate = 0.038
    st.session_state.market = Market(
        futures_prices=dict(_SEED_FUTURES_PRICES),
        interest_rate=_default_rate
    )


st.title("Vol Calculator")

with st.expander("Quick start -- how to use this tool", expanded=False):
    st.markdown(
        """
1. **Set the market** (contract, futures price, interest rate) in the control panel below.
2. **Build a portfolio**: pick a leg type, fill in strike / expiry / vol, click *Add to Portfolio*.
   Or upload a saved portfolio CSV. Select row(s) in the *Current Positions* table and click
   *Remove Selected* to take them out.
3. **Read the tabs** left to right: *Greeks* (sensitivities) -> *Ladders / Surface / Stress*
   (how risk changes if the market moves) -> *Payoff / Breakevens* (profit and loss shape).
   Each tab has a short *What is this?* box.

Units: Delta and Gamma are in contracts; Vega, Theta, Rho and Value are in $k using the contract multiplier.
Prices are futures cents per bushel; time is in calendar days.
        """
    )

# --- control panel (formerly the sidebar) ---
with st.container(border=True):

    col_market, col_addleg = st.columns(2)

    with col_market:

        st.subheader("Market")

        live_quotes = _load_live_quotes()
        contracts_by_commodity = _build_contracts_by_commodity(live_quotes)

        col_ir, col_sofr = st.columns([1, 2])

        with col_ir:
            ir_label_col, ir_value_col = st.columns([1, 1])
            ir_label_col.markdown("Interest Rate (%)")
            interest_rate_pct = ir_value_col.number_input(
                "Interest Rate (%)", value=float(st.session_state.market.interest_rate) * 100,
                step=0.01, format="%.2f", label_visibility="collapsed",
            )
        st.session_state.market.interest_rate = interest_rate_pct / 100

        with col_sofr:
            st.write("")
            try:
                _sofr = _cached_sofr()
                st.caption(
                    f"SOFR: {_sofr['rate'] * 100:.2f}% "
                    f"(updated {_sofr['effective_date']}) -- "
                    f"[NY Fed]({_sofr['source_url']})"
                )
            except Exception:
                st.caption("SOFR unavailable -- using manually entered rate.")

        _render_quotes_box()

    with col_addleg:

        st.subheader("Add Leg")

        all_contracts = [c for codes in contracts_by_commodity.values() for c in codes]

        type_label_col, type_sel_col, contract_label_col, contract_sel_col = st.columns([0.5, 1, 0.7, 1.3])

        type_label_col.markdown("Type")
        leg_type = type_sel_col.selectbox(
            "Type", list(LEG_BUILDERS.keys()), label_visibility="collapsed"
        )
        spec = LEG_BUILDERS[leg_type]

        contract_label_col.markdown("Contract")

        if not all_contracts:
            contract_sel_col.warning("No contracts -- set a price in Market above first.")
            current_contract = None
            auto_expiry = None
            underlying_price = None
        else:
            current_contract = contract_sel_col.selectbox(
                "Contract", all_contracts,
                format_func=lambda c: f"{_commodity_from_contract(c)} -- {c}",
                label_visibility="collapsed",
            )
            underlying_price = st.session_state.market.futures_prices.get(current_contract)
            auto_expiry = approx_option_expiry_days(current_contract)

        is_future = leg_type == "Future"

        # Real CME expiry candidates (standard + serial options on this
        # contract) -- outside the form so picking one immediately
        # updates the Expiry (d) field's default below, same reasoning
        # as Contract already being outside the form.
        chosen_expiry_days = None
        if not is_future and current_contract and "expiry_days" in spec["defaults"]:
            expiry_candidates = option_expiry_candidates(current_contract)
            if expiry_candidates:
                expiry_label_col, expiry_sel_col = st.columns([0.5, 2.5])
                expiry_label_col.markdown("Suggested Expiry")
                chosen_label = expiry_sel_col.selectbox(
                    "Suggested Expiry", [c["label"] for c in expiry_candidates],
                    label_visibility="collapsed", key=f"expiry_choice_{current_contract}",
                )
                chosen_expiry_days = next(
                    c["days"] for c in expiry_candidates if c["label"] == chosen_label
                )

        with st.form(f"add_{leg_type}"):

            n_cols = len(spec["defaults"]) + (1 if underlying_price is not None else 0) + (0 if is_future else 1)
            field_cols = st.columns(n_cols)

            kwargs = {}
            underlying_override = None
            col_idx = 0
            for field, default in spec["defaults"].items():
                col = field_cols[col_idx]
                col_idx += 1

                if field == "expiry_days":
                    # Pre-filled from the chosen suggestion above (real
                    # CME standard/serial expiry) if one exists, else the
                    # old rough estimate -- either way, still a normal
                    # editable field, never disabled, since the exact
                    # fill can differ from any suggestion.
                    default_expiry = chosen_expiry_days if chosen_expiry_days is not None else (auto_expiry or default)
                    kwargs[field] = int(col.number_input(field, value=int(default_expiry), step=1))
                else:
                    kwargs[field] = _render_field(col, field, default)

                if field == "quantity" and underlying_price is not None:
                    # A Future's own price is its Underlying -- there is no
                    # separate Entry Price for a Future.
                    underlying_override = field_cols[col_idx].number_input(
                        "Underlying", value=float(underlying_price),
                        key=f"underlying_display_{current_contract}",
                    )
                    col_idx += 1

            entry_price_input = None
            if not is_future and underlying_price is not None:
                # Entry Price only applies to options -- what you actually
                # paid/received for the premium, a fixed number you type in
                # and that never gets recalculated. Pre-filled with this
                # leg's own computed Mark to Market (not the underlying's
                # price) as a close starting point, same engine used for the
                # live Mark to Market column, but still fully editable since
                # the real fill price can differ from the model.
                try:
                    preview = spec["build"](**kwargs, contract=current_contract)
                    default_entry_price = compute_model_price(
                        preview, {"contract": current_contract}, st.session_state.market,
                    )
                except Exception:
                    default_entry_price = None

                entry_price_input = field_cols[col_idx].number_input(
                    "Entry Price", value=float(default_entry_price or 0.0),
                    key=f"entry_price_input_{current_contract}",
                )
                col_idx += 1

            with st.container(key="add_leg_btn"):
                submitted = st.form_submit_button("Add to Portfolio")

        if submitted:
            if current_contract is None:
                st.error("Pick a contract first.")
            else:
                try:
                    if underlying_override is not None and underlying_override != underlying_price:
                        st.session_state.market.futures_prices[current_contract] = underlying_override

                    kwargs["contract"] = current_contract
                    instrument = spec["build"](**kwargs)
                    st.session_state.portfolio.add(instrument)
                    st.session_state.position_meta.append({
                        "commodity": _commodity_from_contract(current_contract),
                        "contract": current_contract,
                        "vol_pct": None,
                        # Futures have no Entry Price -- their price lives
                        # under Underlying Price instead (same value the
                        # Underlying field above just set).
                        "settlement_price": underlying_override if is_future else entry_price_input,
                    })
                    _autosave_portfolio()
                    st.success(f"Added {leg_type}")
                except Exception as e:
                    st.error(str(e))

        st.subheader("Manage")

        mc_manage, mc_import = st.columns([1.6, 1.4])

        with mc_manage:
            with st.container(key="clear_btn"):
                if st.button("Clear Portfolio"):
                    st.session_state.portfolio = Portfolio()
                    st.session_state.position_meta = []
                    st.session_state.csv_locked_contracts = set()
                    _autosave_portfolio()
                    st.rerun()

            uploaded = st.file_uploader("Upload CSV", type="csv", key="csv_uploader")

            mc_load, mc_download = st.columns(2)

            with mc_load:
                with st.container(key="file_btn_load"):
                    if uploaded is not None and st.button("Load uploaded CSV"):
                        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tmp:
                            tmp.write(uploaded.getvalue())
                            tmp_path = tmp.name
                        st.session_state.portfolio = portfolio_from_csv(tmp_path)
                        st.session_state.position_meta = [
                            {
                                "commodity": _commodity_from_contract(pos.contract),
                                "contract": pos.contract,
                                "vol_pct": None,
                            }
                            for pos in st.session_state.portfolio.positions
                        ]
                        _autosave_portfolio()
                        st.success(f"Loaded {len(st.session_state.portfolio.positions)} position(s)")
                        st.rerun()

            with mc_download:
                with st.container(key="file_btn_download"):
                    if st.session_state.portfolio.positions:
                        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tmp:
                            portfolio_to_csv(st.session_state.portfolio, tmp.name)
                            with open(tmp.name, "rb") as f:
                                csv_bytes = f.read()
                        st.download_button(
                            "Download portfolio CSV", data=csv_bytes, file_name="portfolio.csv", mime="text/csv"
                        )

        with mc_import:
            with st.container(key="daily_prop_upload_box", border=True):

                st.markdown("**Daily_Prop_upload**")
                st.caption("Bushel export -- adds to current positions.")

                imported_file = st.file_uploader(
                    "Upload positions file", type="csv", key="bushel_uploader", label_visibility="collapsed"
                )

                with st.container(key="file_btn_import"):
                    if imported_file is not None and st.button("Add Imported Positions"):
                        with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tmp:
                            tmp.write(imported_file.getvalue())
                            tmp_path = tmp.name
                        try:
                            # The saved tempfile's own name is random (not
                            # the original "Daily_Prop_Bod_Detail_*.csv"), so
                            # the as-of date has to come from the file the
                            # browser actually uploaded, or options get
                            # priced with the wrong time-to-expiry.
                            imported = import_bushel_positions(
                                tmp_path, as_of=daily_prop_as_of_date(imported_file.name),
                            )
                            for instrument, meta in imported:
                                st.session_state.portfolio.add(instrument)
                                st.session_state.position_meta.append(meta)
                                # Only a Future leg's settlement_price is the
                                # underlying contract's own price -- an Option
                                # leg's settlement_price is that option's own
                                # premium, a completely different number, and
                                # must never overwrite the underlying price.
                                if isinstance(instrument, Future) and meta.get("settlement_price") is not None:
                                    st.session_state.market.futures_prices[meta["contract"]] = meta["settlement_price"]
                                    st.session_state.csv_locked_contracts.add(meta["contract"])
                            _autosave_portfolio()
                            st.success(f"Added {len(imported)} imported position(s)")
                            st.rerun()
                        except Exception as e:
                            st.error(f"Import failed: {e}")

    st.divider()
    st.subheader("Current Positions")

    if not st.session_state.portfolio.positions:
        st.caption("No positions.")
    else:
        meta = st.session_state.position_meta

        for group_commodity in _commodities_present(meta):

            st.markdown(f"**{group_commodity} Positions**")

            table = _positions_table_for_commodity(
                st.session_state.portfolio, meta, group_commodity, st.session_state.market,
            )
            selection = st.dataframe(
                table,
                hide_index=True,
                width="stretch",
                height=min(38 * (len(table) + 1), 400),
                column_order=[
                    "Contract", "Entry Price", "Underlying Price", "Mark to Market",
                    "Structure", "Type", "Quantity", "Strike", "Expiry (d)", "Vol (%)",
                ],
                on_select="rerun",
                selection_mode="multi-row",
                key=f"positions_table_{group_commodity}",
            )

            selected_rows = selection["selection"]["rows"]
            selected_positions = sorted({table.iloc[r]["#"] for r in selected_rows}, reverse=True)

            if selected_positions and st.button(
                f"Remove Selected ({len(selected_positions)} position(s))",
                key=f"remove_btn_{group_commodity}",
            ):
                for i in selected_positions:
                    st.session_state.portfolio.positions.pop(i)
                    st.session_state.position_meta.pop(i)
                _autosave_portfolio()
                st.rerun()


def _contracts_in_portfolio(portfolio):

    contracts = []

    for pos in portfolio.positions:
        c = pos.contract
        if c and c not in contracts:
            contracts.append(c)

    return sorted(contracts)


def _contract_selector(key):
    """
    The "Future Spot" being analyzed on this tab: only this contract's price
    (and, where relevant, vol) is shocked -- every other contract's legs are
    priced at their own current value throughout, so the whole portfolio's
    P&L is shown, not just one commodity's slice of it.

    Options come from the contracts actually present in the portfolio, not
    from market.futures_prices -- that dict can hold leftover entries (e.g.
    from a contract you cleared out of the book or just browsed past in the
    Market panel) that would otherwise show up as selectable here.
    """

    available = _contracts_in_portfolio(st.session_state.portfolio)

    if not available:
        st.info("No contract found in the current positions.")
        return None

    return st.selectbox("Future Spot", available, key=key)


def _tab_help(text, formulas=None):

    with st.expander("What is this?"):
        st.markdown(text)

        if formulas:
            st.markdown("**Formulas**")
            for f in formulas:
                st.latex(f)


# --- main area ---
portfolio = st.session_state.portfolio
market = st.session_state.market

# Defensive backfill: any contract a held position needs but that has
# no price at all (a leg that got in before this existed, an autosaved
# session from an older bug, or any other path that didn't seed one)
# gets one try against the live feed; if that's still empty, the leg
# is dropped rather than crashing every Greek/valuation calc on every
# single rerun from here on, including the very next page load.
_unpriced = [c for c in _contracts_in_portfolio(portfolio) if c not in market.futures_prices]
if _unpriced:
    _live_fallback = {q["contract"]: q["price"] for q in _load_live_quotes()}
    _still_unpriced = {c for c in _unpriced if c not in _live_fallback}

    for c in _unpriced:
        if c in _live_fallback:
            market.futures_prices[c] = _live_fallback[c]

    if _still_unpriced:
        keep_idx = [i for i, pos in enumerate(portfolio.positions) if pos.contract not in _still_unpriced]
        portfolio.positions = [portfolio.positions[i] for i in keep_idx]
        st.session_state.position_meta = [st.session_state.position_meta[i] for i in keep_idx]
        _autosave_portfolio()
        st.warning(
            f"Removed position(s) with no price available anywhere (not in the uploaded file "
            f"as a Future, and not live-quoted): {', '.join(sorted(_still_unpriced))}"
        )

tabs = st.tabs([
    "Greeks", "Ladders", "Surface",
    "Stress", "Strike Map", "Payoff", "Breakevens",
])

has_positions = bool(portfolio.positions)

with tabs[0]:

    st.subheader("Greek Summary")
    _tab_help(
        "How the portfolio reacts to small market moves.\n\n"
        "- **Delta**: effective number of futures contracts you are long/short.\n"
        "- **Gamma**: change in Delta (contracts) per 1 cent/bu move.\n"
        "- **Vega**: $k gained/lost per 1% change in vol.\n"
        "- **Theta**: $k lost per day from time passing.\n"
        "- **Rho**: $k gained/lost per 1% change in interest rates.\n\n"
        "Shown in $k using the contract multiplier (Delta/Gamma stay in contracts/cents). "
        "*By Tenor* splits the same Greeks by expiry, same $k scale.\n\n"
        "The formulas below are the European Black-76 definitions the model is built on. The tool prices "
        "American options with the Barone-Adesi-Whaley approximation and computes every Greek by bumping "
        "the input and repricing, so the numbers approximate these rather than evaluating them directly.",
        formulas=[
            r"d_1 = \frac{\ln(F/K) + \tfrac{1}{2}\sigma^2 T}{\sigma\sqrt{T}}, \qquad d_2 = d_1 - \sigma\sqrt{T}",
            r"C = e^{-rT}\left[F\,N(d_1) - K\,N(d_2)\right]",
            r"\Delta = e^{-rT} N(d_1), \qquad \Gamma = \frac{e^{-rT} N'(d_1)}{F\sigma\sqrt{T}}",
            r"\text{Vega} = F e^{-rT} N'(d_1)\sqrt{T}, \qquad \Theta = -\frac{F e^{-rT} N'(d_1)\,\sigma}{2\sqrt{T}} + rC",
            r"\rho = \frac{\partial C}{\partial r} = -T\,C",
        ],
    )

    if not has_positions:
        st.info("Add positions to see Greeks.")
    else:
        fs_contract = _contract_selector("fs_greeks")

        if fs_contract:
            dollars = GreekEngine.report_dollars(portfolio, market, fs_contract)

            formatted = {
                k: (f"{v:.3f}" if k == "Value ($k)" else f"{v:.2f}")
                for k, v in dollars.items()
            }

            st.write("$k view")
            st.table(pd.Series(formatted, name="Value").to_frame())

            st.subheader("By Tenor")
            tenor_df = report_by_tenor(portfolio, market, fs_contract)
            tenor_df = tenor_df.apply(
                lambda col: col.map(lambda v: f"{v:.3f}" if col.name == "Value ($k)" else f"{v:.2f}")
            )
            st.dataframe(tenor_df)
            _download_df_button(tenor_df, "Download tenor report CSV", "report_by_tenor.csv")

with tabs[1]:

    st.subheader("Spot Ladder")
    _tab_help(
        "Recomputes the portfolio's P&L and Greeks across a range of futures prices (Spot Ladder) and "
        "vol levels (Vol Ladder), so you can see how risk changes before the market actually moves."
    )

    fs_contract = _contract_selector("fs_ladders") if has_positions else None

    c1, c2, c3 = st.columns(3)
    spot_low = c1.number_input("Low (¢)", value=-50, step=5, key="spot_low")
    spot_high = c2.number_input("High (¢)", value=50, step=5, key="spot_high")
    spot_step = c3.number_input("Step (¢)", value=10, step=5, key="spot_step")

    if fs_contract:
        ladder = spot_ladder(
            portfolio, market, contract=fs_contract, low=spot_low, high=spot_high, step=spot_step
        ).round(3)
        st.dataframe(ladder)
        _download_df_button(ladder, "Download spot ladder CSV", "spot_ladder.csv")

    st.subheader("Vol Ladder")

    c4, c5, c6 = st.columns(3)
    vol_low = c4.number_input("Low (pts)", value=-10, step=1, key="vol_low")
    vol_high = c5.number_input("High (pts)", value=10, step=1, key="vol_high")
    vol_step = c6.number_input("Step (pts)", value=2, step=1, key="vol_step")

    if fs_contract:
        vladder = vol_ladder(
            portfolio, market, contract=fs_contract, low=vol_low, high=vol_high, step=vol_step
        ).round(3)
        st.dataframe(vladder)
        _download_df_button(vladder, "Download vol ladder CSV", "vol_ladder.csv")

with tabs[2]:

    st.subheader("Spot x Vol Surface")
    _tab_help(
        "A heatmap of P&L or a chosen Greek across price and vol at the same time, "
        "so combined moves are visible at a glance."
    )

    fs_contract = _contract_selector("fs_surface") if has_positions else None

    metric = st.selectbox(
        "Metric", ["PnL", "Delta", "Gamma", "Vega", "Volga", "Vanna-D", "Vanna-V"]
    )

    if fs_contract:
        plot_surface(portfolio, market, contract=fs_contract, metric=metric)
        st.pyplot(plt.gcf())
        plt.close()

with tabs[3]:

    st.subheader("Stress Ladders")
    _tab_help(
        "Pushes one variable at a time (price, time or vol) by set amounts and shows how a Greek moves. "
        "It reveals risk that is invisible today but appears if the market moves, e.g. Gamma that is "
        "small now but spikes after a 20 cent drop."
    )
    st.caption("Delta/Gamma x {price, time, vol}, Vega/Theta x {price, time}")

    fs_contract = _contract_selector("fs_stress") if has_positions else None

    if fs_contract:
        plot_stress_report(portfolio, market, contract=fs_contract)
        st.pyplot(plt.gcf())
        plt.close()

with tabs[4]:

    st.subheader("Strike Map")
    _tab_help(
        "A grid of strikes (rows) against expiries (columns), showing where position size is concentrated. "
        "*quantity* sums raw contracts; *delta* weights each leg by its Delta, showing where directional "
        "risk sits. Structures are split into their legs and futures are excluded."
    )

    weight = st.radio("Weight", ["quantity", "delta"], horizontal=True)

    if has_positions:
        grid = strike_map(portfolio, market=market if weight == "delta" else None, weight=weight)
        if grid.empty:
            st.info("No option legs yet -- Strike Map excludes futures, which have no strike.")
        else:
            plot_strike_map(portfolio, market=market if weight == "delta" else None, weight=weight)
            st.pyplot(plt.gcf())
            plt.close()

with tabs[5]:

    st.subheader("Payoff Diagram")
    _tab_help(
        "Profit/loss against the futures price, at chosen days forward. A time step of 0 is today; "
        "larger steps show the position closer to expiry, keeping remaining time value."
    )

    fs_contract = _contract_selector("fs_payoff") if has_positions else None

    time_steps_str = st.text_input("Time steps (days, comma-separated)", "0,30,64")
    time_steps = [int(x.strip()) for x in time_steps_str.split(",") if x.strip()]

    if fs_contract and time_steps:
        plot_payoff(portfolio, market, time_steps=time_steps, contract=fs_contract)
        st.pyplot(plt.gcf())
        plt.close()

with tabs[6]:

    st.subheader("Breakevens")
    _tab_help(
        "The futures moves, in cents from the current price (negative = down, positive = up), at which "
        "one-day P&L crosses zero. Add them to the futures price for the absolute level. Two estimates are "
        "shown: a fast quadratic approximation and an exact one that fully reprices the portfolio. "
        "They should be close; if they diverge, trust the numerical one, especially for butterflies "
        "and collars. `null` means no breakeven exists in the search range.",
        formulas=[
            r"\Delta\,\Delta F + \tfrac{1}{2}\Gamma\,\Delta F^2 + \Theta = 0",
            r"\Delta F = \frac{-\Delta \pm \sqrt{\Delta^2 - 2\Gamma\Theta}}{\Gamma}",
        ],
    )

    if has_positions:

        fs_contract = _contract_selector("fs_breakevens")

        if fs_contract:
            quad = breakevens_quadratic(portfolio, market, fs_contract)
            numeric = breakevens_numerical(portfolio, market, fs_contract)

            col1, col2 = st.columns(2)

            col1.write("Quadratic (fast, Taylor approx)")
            col1.json(quad)

            col2.write("Numerical (exact, 1-day)")
            col2.json(numeric)
