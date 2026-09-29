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
    fetch_latest_sofr,
)

st.set_page_config(page_title="Vol Calculator", layout="wide")


@st.cache_data(ttl=4 * 60 * 60)
def _cached_sofr():
    return fetch_latest_sofr()


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
    try:
        _sofr = _cached_sofr()
        _default_rate = _sofr["rate"]
    except Exception:
        _default_rate = 0.038
    st.session_state.market = Market(futures_price=528.75, interest_rate=_default_rate)


st.title("Vol Calculator")

with st.expander("Quick start -- how to use this tool", expanded=False):
    st.markdown(
        """
1. **Set the market** (futures price, interest rate, contract multiplier) in the control panel below.
2. **Build a portfolio**: pick a leg type, fill in strike / expiry / vol, click *Add to Portfolio*.
   Or upload a saved portfolio CSV. Remove a position with its **x**.
3. **Read the tabs** left to right: *Portfolio* (what you hold) -> *Greeks* (sensitivities) ->
   *Ladders / Surface / Stress* (how risk changes if the market moves) -> *Payoff / Breakevens*
   (profit and loss shape). Each tab has a short *What is this?* box.

Units: Delta and Gamma are in contracts; Vega, Theta, Rho and Value are in $k using the contract multiplier.
Prices are futures cents per bushel; time is in calendar days.
        """
    )

# --- control panel (formerly the sidebar) ---
with st.container(border=True):

    col_market, col_add, col_pos = st.columns([1, 1.4, 1.4])

    with col_market:
        st.subheader("Market")

        futures_price = st.number_input(
            "Futures Price", value=float(st.session_state.market.futures_price), step=0.25
        )
        interest_rate = st.number_input(
            "Interest Rate", value=float(st.session_state.market.interest_rate), step=0.001, format="%.4f"
        )

        try:
            _sofr = _cached_sofr()
            st.caption(
                f"SOFR: {_sofr['rate'] * 100:.2f}% "
                f"(updated {_sofr['effective_date']}) -- "
                f"[NY Fed]({_sofr['source_url']})"
            )
        except Exception:
            st.caption("SOFR unavailable -- using manually entered rate.")

        contract_multiplier = st.number_input(
            "Contract Multiplier ($/pt)", value=float(st.session_state.market.contract_multiplier), step=1.0
        )

        st.session_state.market = Market(
            futures_price=futures_price,
            interest_rate=interest_rate,
            contract_multiplier=contract_multiplier
        )

    with col_add:
        st.subheader("Add Leg")

        leg_type = st.selectbox("Type", list(LEG_BUILDERS.keys()))
        spec = LEG_BUILDERS[leg_type]

        with st.form(f"add_{leg_type}"):

            kwargs = {}
            for field, default in spec["defaults"].items():
                kwargs[field] = _render_field(field, default)

            submitted = st.form_submit_button("Add to Portfolio")

        if submitted:
            try:
                instrument = spec["build"](**kwargs)
                st.session_state.portfolio.add(instrument)
                st.success(f"Added {leg_type}")
            except Exception as e:
                st.error(str(e))

    with col_pos:
        st.subheader("Current Positions")

        if not st.session_state.portfolio.positions:
            st.caption("No positions.")
        else:
            for i, pos in enumerate(st.session_state.portfolio.positions):
                c_label, c_btn = st.columns([5, 1])
                c_label.text(_leg_label(pos))
                if c_btn.button("x", key=f"remove_{i}"):
                    st.session_state.portfolio.positions.pop(i)
                    st.rerun()

        if st.button("Clear Portfolio"):
            st.session_state.portfolio = Portfolio()
            st.rerun()

        st.subheader("Portfolio File")

        uploaded = st.file_uploader("Upload CSV", type="csv", key="csv_uploader")

        if uploaded is not None and st.button("Load uploaded CSV"):
            with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tmp:
                tmp.write(uploaded.getvalue())
                tmp_path = tmp.name
            st.session_state.portfolio = portfolio_from_csv(tmp_path)
            st.success(f"Loaded {len(st.session_state.portfolio.positions)} position(s)")
            st.rerun()

        if st.session_state.portfolio.positions:
            with tempfile.NamedTemporaryFile(suffix=".csv", delete=False) as tmp:
                portfolio_to_csv(st.session_state.portfolio, tmp.name)
                with open(tmp.name, "rb") as f:
                    csv_bytes = f.read()
            st.download_button(
                "Download portfolio CSV", data=csv_bytes, file_name="portfolio.csv", mime="text/csv"
            )


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

tabs = st.tabs([
    "Portfolio", "Greeks", "Ladders", "Surface",
    "Stress", "Strike Map", "Payoff", "Breakevens",
])

has_positions = bool(portfolio.positions)

with tabs[0]:

    st.subheader("Current Portfolio")
    _tab_help(
        "Your current book, one line per position, and its total value. Check this looks right "
        "before reading any other tab."
    )

    if not has_positions:
        st.info("No positions yet -- add one in the control panel above or upload a CSV.")
    else:
        for i, pos in enumerate(portfolio.positions):
            st.text(f"{i}: {_leg_label(pos)}")

        st.metric("Portfolio Value", f"{portfolio.value(market):,.2f}")

with tabs[1]:

    st.subheader("Greek Summary")
    _tab_help(
        "How the portfolio reacts to small market moves.\n\n"
        "- **Delta**: effective number of futures contracts you are long/short.\n"
        "- **Gamma**: change in Delta (contracts) per 1 cent/bu move.\n"
        "- **Vega**: $k gained/lost per 1% change in vol.\n"
        "- **Theta**: $k lost per day from time passing.\n"
        "- **Rho**: $k gained/lost per 1% change in interest rates.\n\n"
        "The left table is raw (contracts); the right is in $k. *By Tenor* splits the same Greeks by expiry.\n\n"
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
    _tab_help(
        "Recomputes the portfolio's P&L and Greeks across a range of futures prices (Spot Ladder) and "
        "vol levels (Vol Ladder), so you can see how risk changes before the market actually moves."
    )

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
    _tab_help(
        "A heatmap of P&L or a chosen Greek across price and vol at the same time, "
        "so combined moves are visible at a glance."
    )

    metric = st.selectbox(
        "Metric", ["PnL", "Delta", "Gamma", "Vega", "Volga", "Vanna-D", "Vanna-V"]
    )

    if has_positions:
        plot_surface(portfolio, market, metric=metric)
        st.pyplot(plt.gcf())
        plt.close()

with tabs[4]:

    st.subheader("Stress Ladders")
    _tab_help(
        "Pushes one variable at a time (price, time or vol) by set amounts and shows how a Greek moves. "
        "It reveals risk that is invisible today but appears if the market moves, e.g. Gamma that is "
        "small now but spikes after a 20 cent drop."
    )
    st.caption("Delta/Gamma x {price, time, vol}, Vega/Theta x {price, time}")

    if has_positions:
        plot_stress_report(portfolio, market)
        st.pyplot(plt.gcf())
        plt.close()

with tabs[5]:

    st.subheader("Strike Map")
    _tab_help(
        "A grid of strikes (rows) against expiries (columns), showing where position size is concentrated. "
        "*quantity* sums raw contracts; *delta* weights each leg by its Delta, showing where directional "
        "risk sits. Structures are split into their legs and futures are excluded."
    )

    weight = st.radio("Weight", ["quantity", "delta"], horizontal=True)

    if has_positions:
        plot_strike_map(portfolio, market=market if weight == "delta" else None, weight=weight)
        st.pyplot(plt.gcf())
        plt.close()

with tabs[6]:

    st.subheader("Payoff Diagram")
    _tab_help(
        "Profit/loss against the futures price, at chosen days forward. A time step of 0 is today; "
        "larger steps show the position closer to expiry, keeping remaining time value."
    )

    time_steps_str = st.text_input("Time steps (days, comma-separated)", "0,30,64")
    time_steps = [int(x.strip()) for x in time_steps_str.split(",") if x.strip()]

    if has_positions and time_steps:
        plot_payoff(portfolio, market, time_steps=time_steps)
        st.pyplot(plt.gcf())
        plt.close()

with tabs[7]:

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

        quad = breakevens_quadratic(portfolio, market)
        numeric = breakevens_numerical(portfolio, market)

        col1, col2 = st.columns(2)

        col1.write("Quadratic (fast, Taylor approx)")
        col1.json(quad)

        col2.write("Numerical (exact, 1-day)")
        col2.json(numeric)
