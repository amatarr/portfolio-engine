import tempfile

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
    plot_stress_report,
    plot_payoff,
    portfolio_to_csv,
    portfolio_from_csv,
)

st.set_page_config(page_title="Vol Calculator", layout="wide")


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
        "build": lambda **kw: Future(quantity=kw["quantity"]),
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


def _render_field(name, default):

    if name == "option_type":
        options = ["call", "put"]
        return st.selectbox(name, options, index=options.index(default))

    if "days" in name:
        return int(st.number_input(name, value=int(default), step=1))

    if "vol" in name:
        return st.number_input(name, value=float(default), step=0.01, format="%.4f")

    return st.number_input(name, value=float(default), step=1.0)


def _leg_label(pos):

    if isinstance(pos, Structure):
        return pos.name or "Structure"

    return repr(pos)


def _default_portfolio():

    portfolio = Portfolio()
    portfolio.add(Future(quantity=-1))
    portfolio.add(AmericanOption(quantity=-5, strike=620, expiry_days=64, option_type="call", volatility=0.26189))
    portfolio.add(AmericanOption(quantity=2, strike=530, expiry_days=64, option_type="call", volatility=0.2150))
    return portfolio


def _download_df_button(df, label, filename):
    st.download_button(label, data=df.to_csv().encode("utf-8"), file_name=filename, mime="text/csv")


if "portfolio" not in st.session_state:
    st.session_state.portfolio = _default_portfolio()

if "market" not in st.session_state:
    st.session_state.market = Market(futures_price=528.75, interest_rate=0.038)


st.title("Vol Calculator")

# --- sidebar: market ---
st.sidebar.header("Market")

futures_price = st.sidebar.number_input(
    "Futures Price", value=float(st.session_state.market.futures_price), step=0.25
)
interest_rate = st.sidebar.number_input(
    "Interest Rate", value=float(st.session_state.market.interest_rate), step=0.001, format="%.4f"
)
contract_multiplier = st.sidebar.number_input(
    "Contract Multiplier ($/pt)", value=float(st.session_state.market.contract_multiplier), step=1.0
)

st.session_state.market = Market(
    futures_price=futures_price,
    interest_rate=interest_rate,
    contract_multiplier=contract_multiplier
)

# --- sidebar: add leg ---
st.sidebar.header("Add Leg")

leg_type = st.sidebar.selectbox("Type", list(LEG_BUILDERS.keys()))
spec = LEG_BUILDERS[leg_type]

with st.sidebar.form(f"add_{leg_type}"):

    kwargs = {}
    for field, default in spec["defaults"].items():
        kwargs[field] = _render_field(field, default)

    submitted = st.form_submit_button("Add to Portfolio")

if submitted:
    try:
        instrument = spec["build"](**kwargs)
        st.session_state.portfolio.add(instrument)
        st.sidebar.success(f"Added {leg_type}")
    except Exception as e:
        st.sidebar.error(str(e))

# --- sidebar: current positions ---
st.sidebar.header("Current Positions")

if not st.session_state.portfolio.positions:
    st.sidebar.caption("No positions.")
else:
    for i, pos in enumerate(st.session_state.portfolio.positions):
        col1, col2 = st.sidebar.columns([5, 1])
        col1.text(_leg_label(pos))
        if col2.button("x", key=f"remove_{i}"):
            st.session_state.portfolio.positions.pop(i)
            st.rerun()

if st.sidebar.button("Clear Portfolio"):
    st.session_state.portfolio = Portfolio()
    st.rerun()

# --- sidebar: CSV import/export ---
st.sidebar.header("Portfolio File")

uploaded = st.sidebar.file_uploader("Upload CSV", type="csv", key="csv_uploader")

if uploaded is not None and st.sidebar.button("Load uploaded CSV"):
    with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tmp:
        tmp.write(uploaded.getvalue())
        tmp_path = tmp.name
    st.session_state.portfolio = portfolio_from_csv(tmp_path)
    st.sidebar.success(f"Loaded {len(st.session_state.portfolio.positions)} position(s)")
    st.rerun()

if st.session_state.portfolio.positions:
    with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tmp:
        portfolio_to_csv(st.session_state.portfolio, tmp.name)
        with open(tmp.name, "rb") as f:
            csv_bytes = f.read()
    st.sidebar.download_button(
        "Download portfolio CSV", data=csv_bytes, file_name="portfolio.csv", mime="text/csv"
    )


# --- main area ---
portfolio = st.session_state.portfolio
market = st.session_state.market

tabs = st.tabs([
    "Portfolio", "Greeks", "Ladders", "Surface",
    "Stress", "Strike Map", "Payoff", "Breakevens",
])

has_positions = bool(portfolio.positions)

with tabs[0]:

    st.subheader("Current Portfolio")

    if not has_positions:
        st.info("No positions yet -- add one from the sidebar or upload a CSV.")
    else:
        for i, pos in enumerate(portfolio.positions):
            st.text(f"{i}: {_leg_label(pos)}")

        st.metric("Portfolio Value", f"{portfolio.value(market):,.2f}")

with tabs[1]:

    st.subheader("Greek Summary")

    if not has_positions:
        st.info("Add positions to see Greeks.")
    else:
        raw = GreekEngine.report(portfolio, market)
        dollars = GreekEngine.report_dollars(portfolio, market)

        col1, col2 = st.columns(2)
        col1.write("Raw (contracts)")
        col1.table(pd.Series(raw, name="Value").to_frame())
        col2.write("$k view")
        col2.table(pd.Series(dollars, name="Value").to_frame())

        st.subheader("By Tenor")
        tenor_df = report_by_tenor(portfolio, market).round(4)
        st.dataframe(tenor_df)
        _download_df_button(tenor_df, "Download tenor report CSV", "report_by_tenor.csv")

with tabs[2]:

    st.subheader("Spot Ladder")

    c1, c2, c3 = st.columns(3)
    spot_low = c1.number_input("Low %", value=-15, step=1, key="spot_low")
    spot_high = c2.number_input("High %", value=15, step=1, key="spot_high")
    spot_step = c3.number_input("Step %", value=3, step=1, key="spot_step")

    if has_positions:
        ladder = spot_ladder(portfolio, market, low=spot_low, high=spot_high, step=spot_step).round(3)
        st.dataframe(ladder)
        _download_df_button(ladder, "Download spot ladder CSV", "spot_ladder.csv")

    st.subheader("Vol Ladder")

    c4, c5, c6 = st.columns(3)
    vol_low = c4.number_input("Low (pts)", value=-10, step=1, key="vol_low")
    vol_high = c5.number_input("High (pts)", value=10, step=1, key="vol_high")
    vol_step = c6.number_input("Step (pts)", value=2, step=1, key="vol_step")

    if has_positions:
        vladder = vol_ladder(portfolio, market, low=vol_low, high=vol_high, step=vol_step).round(3)
        st.dataframe(vladder)
        _download_df_button(vladder, "Download vol ladder CSV", "vol_ladder.csv")

with tabs[3]:

    st.subheader("Spot x Vol Surface")

    metric = st.selectbox(
        "Metric", ["PnL", "Delta", "Gamma", "Vega", "Volga", "Vanna-D", "Vanna-V"]
    )

    if has_positions:
        plot_surface(portfolio, market, metric=metric)
        st.pyplot(plt.gcf())
        plt.close()

with tabs[4]:

    st.subheader("Stress Ladders")
    st.caption("Delta/Gamma x {price, time, vol}, Vega/Theta x {price, time}")

    if has_positions:
        plot_stress_report(portfolio, market)
        st.pyplot(plt.gcf())
        plt.close()

with tabs[5]:

    st.subheader("Strike Map")

    weight = st.radio("Weight", ["quantity", "delta"], horizontal=True)

    if has_positions:
        plot_strike_map(portfolio, market=market if weight == "delta" else None, weight=weight)
        st.pyplot(plt.gcf())
        plt.close()

with tabs[6]:

    st.subheader("Payoff Diagram")

    time_steps_str = st.text_input("Time steps (days, comma-separated)", "0,30,64")
    time_steps = [int(x.strip()) for x in time_steps_str.split(",") if x.strip()]

    if has_positions and time_steps:
        plot_payoff(portfolio, market, time_steps=time_steps)
        st.pyplot(plt.gcf())
        plt.close()

with tabs[7]:

    st.subheader("Breakevens")

    if has_positions:

        quad = breakevens_quadratic(portfolio, market)
        numeric = breakevens_numerical(portfolio, market)

        col1, col2 = st.columns(2)

        col1.write("Quadratic (fast, Taylor approx)")
        col1.json(quad)

        col2.write("Numerical (exact, 1-day)")
        col2.json(numeric)
