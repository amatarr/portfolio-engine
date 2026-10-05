import re
import tempfile

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
import yfinance as yf

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
    MONTH_CODES,
    contract_code,
    multiplier_for,
    import_bushel_positions,
    import_cqg_quotes,
)

st.set_page_config(page_title="Vol Calculator", layout="wide")

st.markdown(
    """
    <style>
    div[data-testid="stNumberInput"] { max-width: 170px; }
    div[data-testid="stTextInput"] { max-width: 240px; }
    div[data-testid="stSelectbox"] { max-width: 240px; }

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


def _load_live_quotes():
    """
    Reads the CQGXL export fresh on every rerun (no caching) so prices
    stay live as the sheet updates -- only works when the file is present
    next to app.py, i.e. running locally with CQGXL writing to it.
    """

    try:
        return import_cqg_quotes(_CQG_QUOTES_PATH)
    except FileNotFoundError:
        return []


# Free, no-key placeholder source while CQG API access gets sorted out --
# continuous front-month quotes, delayed ~15-20min, not tied to a specific
# deferred contract month.
_CME_YF_TICKERS = {"Corn": "ZC=F", "Soybean": "ZS=F", "SRW Wheat": "ZW=F"}


@st.cache_data(ttl=30)
def _load_cme_quotes():

    quotes = []

    for commodity, ticker in _CME_YF_TICKERS.items():
        try:
            price = yf.Ticker(ticker).fast_info.get("lastPrice")
        except Exception:
            price = None

        if price is not None:
            quotes.append({
                "commodity": commodity,
                "contract": f"{ticker} (front month)",
                "price": float(price),
            })

    return quotes


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


def _leg_row(leg, position_index, structure, contract, future_vol_pct=None):

    if isinstance(leg, Future):
        return {
            "#": position_index, "Contract": contract, "Structure": structure, "Type": "Future",
            "Quantity": leg.quantity, "Strike": None, "Expiry (d)": None, "Vol (%)": future_vol_pct,
        }

    return {
        "#": position_index, "Contract": contract, "Structure": structure,
        "Type": "Call" if leg.option_type == "call" else "Put",
        "Quantity": leg.quantity, "Strike": leg.strike,
        "Expiry (d)": leg.expiry_days, "Vol (%)": leg.volatility * 100,
    }


def _positions_table_for_commodity(portfolio, meta, commodity):

    rows = []

    for i, pos in enumerate(portfolio.positions):

        if meta[i].get("commodity") != commodity:
            continue

        contract = meta[i].get("contract") or "-"

        if isinstance(pos, Structure):
            structure = _structure_kind(pos.name)
            rows.extend(_leg_row(leg, i, structure, contract) for leg in pos.legs)
        else:
            rows.append(_leg_row(pos, i, "-", contract, future_vol_pct=meta[i].get("vol_pct")))

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


def _commodity_from_contract(contract, fallback="Unknown"):

    if not contract:
        return fallback

    for name, info in COMMODITIES.items():
        if contract.startswith(info["symbol"]):
            return name

    return fallback


def _download_df_button(df, label, filename):
    st.download_button(label, data=df.to_csv().encode("utf-8"), file_name=filename, mime="text/csv")


if "portfolio" not in st.session_state:
    st.session_state.portfolio = _default_portfolio()

if "position_meta" not in st.session_state:
    st.session_state.position_meta = []

if "market" not in st.session_state:
    try:
        _sofr = _cached_sofr()
        _default_rate = _sofr["rate"]
    except Exception:
        _default_rate = 0.038
    st.session_state.market = Market(
        futures_prices={},
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

    col_market, col_add = st.columns([1, 2])

    with col_market:
        st.subheader("Market")

        with st.container(key="cqg_quotes_box", border=True):
            st.markdown("**Live Quotes (CME via Yahoo -- preview)**")
            st.caption("Placeholder source while CQG API access is pending. Delayed ~15-20min, front-month only.")

            live_quotes = _load_cme_quotes()

            if live_quotes:
                quotes_df = pd.DataFrame(live_quotes).rename(
                    columns={"commodity": "Commodity", "contract": "Contract", "price": "Price"}
                )
                st.dataframe(quotes_df, hide_index=True, width="stretch")
            else:
                st.caption("Live quotes unavailable right now -- using manual entry below.")

        commodity = st.selectbox("Underlying", list(COMMODITIES.keys()))
        available_months = COMMODITIES[commodity]["months"]

        c_month, c_year = st.columns(2)
        month_code = c_month.selectbox(
            "Month", available_months,
            format_func=lambda m: f"{m} ({MONTH_CODES[m]})",
        )
        year = c_year.number_input("Year", value=2026, step=1, format="%d")

        current_contract = contract_code(commodity, month_code, year)
        st.caption(f"Contract: {current_contract}")

        quote_lookup = {q["contract"]: q["price"] for q in live_quotes}

        if current_contract in quote_lookup:
            futures_price = quote_lookup[current_contract]
            st.caption(f"Futures Price: {futures_price:.2f} (live quote)")
        else:
            existing_price = st.session_state.market.futures_prices.get(current_contract, 528.75)
            futures_price = st.number_input(
                "Futures Price", value=float(existing_price), step=0.25, key=f"price_{current_contract}"
            )
            st.caption("No live quote for this contract -- manual entry.")

        interest_rate_pct = st.number_input(
            "Interest Rate (%)", value=float(st.session_state.market.interest_rate) * 100,
            step=0.01, format="%.2f",
        )
        interest_rate = interest_rate_pct / 100

        try:
            _sofr = _cached_sofr()
            st.caption(
                f"SOFR: {_sofr['rate'] * 100:.2f}% "
                f"(updated {_sofr['effective_date']}) -- "
                f"[NY Fed]({_sofr['source_url']})"
            )
        except Exception:
            st.caption("SOFR unavailable -- using manually entered rate.")

        contract_multiplier = multiplier_for(commodity)

        # Mutate in place -- a fresh Market() would wipe every other
        # contract's price, and this book can hold several at once.
        st.session_state.market.futures_prices[current_contract] = futures_price
        st.session_state.market.interest_rate = interest_rate
        st.session_state.market.contract_multiplier = contract_multiplier

    with col_add:
        st.subheader("Add Leg")

        leg_type = st.selectbox("Type", list(LEG_BUILDERS.keys()))
        spec = LEG_BUILDERS[leg_type]

        with st.form(f"add_{leg_type}"):

            field_cols = st.columns(len(spec["defaults"]))

            kwargs = {}
            for col, (field, default) in zip(field_cols, spec["defaults"].items()):
                kwargs[field] = _render_field(col, field, default)

            with st.container(key="add_leg_btn"):
                submitted = st.form_submit_button("Add to Portfolio")

        if submitted:
            try:
                kwargs["contract"] = current_contract
                instrument = spec["build"](**kwargs)
                st.session_state.portfolio.add(instrument)
                st.session_state.position_meta.append({
                    "commodity": commodity,
                    "contract": current_contract,
                    "vol_pct": None,
                })
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
                                "commodity": _commodity_from_contract(pos.contract, fallback=commodity),
                                "contract": pos.contract,
                                "vol_pct": None,
                            }
                            for pos in st.session_state.portfolio.positions
                        ]
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
                            imported = import_bushel_positions(tmp_path)
                            for instrument, meta in imported:
                                st.session_state.portfolio.add(instrument)
                                st.session_state.position_meta.append(meta)
                                if meta.get("settlement_price") is not None:
                                    st.session_state.market.futures_prices[meta["contract"]] = meta["settlement_price"]
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

            table = _positions_table_for_commodity(st.session_state.portfolio, meta, group_commodity)
            selection = st.dataframe(
                table,
                hide_index=True,
                width="stretch",
                height=min(38 * (len(table) + 1), 400),
                column_order=["Contract", "Structure", "Type", "Quantity", "Strike", "Expiry (d)", "Vol (%)"],
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

            st.write("$k view")
            st.table(pd.Series(dollars, name="Value").to_frame())

            st.subheader("By Tenor")
            tenor_df = report_by_tenor(portfolio, market, fs_contract).round(4)
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
    spot_low = c1.number_input("Low %", value=-15, step=1, key="spot_low")
    spot_high = c2.number_input("High %", value=15, step=1, key="spot_high")
    spot_step = c3.number_input("Step %", value=3, step=1, key="spot_step")

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
