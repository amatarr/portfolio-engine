# Vol Calculator

Futures + American-option portfolio risk engine (QuantLib-based Greeks, scenario ladders, multi-leg
structures, breakevens, payoff diagrams) with a Streamlit UI.

See `VolCalculator_Full_Documentation.pdf` for the complete mathematical and implementation reference.

## Local setup

```
python -m venv .venv
.venv\Scripts\pip install -e .
.venv\Scripts\python -m streamlit run app.py
```

## App access

`app.py` is gated behind a password read from Streamlit secrets (`app_password`). Locally, create
`.streamlit/secrets.toml` (already gitignored):

```
app_password = "your-password-here"
```

On Streamlit Community Cloud, set the same key under the app's Settings > Secrets instead of committing
a secrets file.
The login is presented in a centered card with a high-contrast password field and submit button.
Submitting the password form with either the Enter key or the button attempts authentication.

Monetary values in the Greek summary, analysis ladders (including their CSV downloads), surface,
stress charts, and payoff chart are displayed in thousands of dollars ($K), using the configured
contract multiplier. The Tenor and ladder headers include each measure's units:

- Delta is expressed as futures-equivalent contracts; Gamma is the change in Delta per 1¢/bu price move.
- Vega is $K per 1 volatility point (1 percentage point); Volga is $K per volatility point squared.
- Theta is $K per day; Rho is approximate $K per 0.1 percentage point (10 basis point) rate move.
- Rho uses a central 1-basis-point rate bump to estimate sensitivity, then scales that result to a 10-basis-point move.
- Vanna-D is the change in Delta per volatility point; Vanna-V is the change in Vega ($K per
  volatility point) per 1¢/bu price move.

The Greek Summary by Tenor shows a subtotal for each underlying futures contract, aggregating its
future and all options that reference it even when an option's listed month differs. Expand a
contract group to see the individual futures and option-expiry rows. Its column headers include units
and provide tooltips explaining each sensitivity and how the selected commodity affects the calculation
scope.

## Users and saved portfolios

Type your name in **User** at the top of the app (it is kept in the URL as `?user=...`). Each user's
portfolio is then autosaved online in Supabase under that name, so users no longer share one book and
it survives redeploys. **Manage Portfolio > Saved portfolios** saves named copies, lists them (yours,
or everyone's), and loads or deletes them. Without a user name, or if Supabase is unreachable, the app
falls back to the local `_autosave_portfolio.csv`.

One-time setup: run `supabase_saved_portfolios.sql` in the Supabase dashboard (SQL Editor). The app
uses the existing `supabase_url` / `supabase_anon_key` secrets.

## Portfolio screen

The Market panel shows compact quote tables for the available commodities. Portfolio management
actions (clear, upload, import, and download) are grouped under the collapsed **Manage Portfolio**
section to keep the Add Leg form and position data visible. The Add Leg form labels the position size
as **Quantity**; its Underlying field is read-only (always the live price). Current Positions show a
grouped summary with expandable lot details; short quantities are red in the detail tables. Quantity,
Entry Price, Reference Price, and Current Vol edits are autosaved and refresh analysis views.
Reference Price is the underlying price captured when an option is added;
it remains static unless edited and is blank for futures. Entry Vol is captured when an option is added
and remains as a read-only reference; Current Vol is manually editable and used to calculate TheoV and
Greeks. Existing portfolios without an entry-volatility
record use their saved volatility as the initial reference. The quote feed supplies futures prices,
not option volatility, and futures have no option volatility. Underlying Price edits update
the market price shared by every position on that contract and remain in effect until that contract's
live quote changes. For structures, quantity and Entry Price edits apply to the selected leg.
Underlying Price is shared by underlying futures contract; removing any selected structure leg
removes the full structure.
Current Positions now starts with a compact summary grouped by risk instrument: futures by contract,
and options by underlying contract, call/put, strike, and expiry. Summary quantities are net; Entry
Price, Reference Price, Entry Vol, and TheoV are absolute-quantity-weighted averages for context.
Each summary row has an inline expand/collapse control; its editable lot table appears directly below
that row. Summary rows sort by numeric Expiry (d), ascending; expired positions appear first and
positions without a defined expiry appear last. This uses native Streamlit columns rather than a
dedicated tree-grid component, so the summary rows do not have interactive column sorting or resizing.
The positions section is centered and capped at 1600 pixels wide, with compact gaps between summary
columns; Expiry, Entry Vol, Current Vol, and Trade Count values are centered below their headers.
Current Vol edits in the detail table apply to every matching option lot, while Greek calculations
continue to use the underlying positions. Legacy lots with different Current Vol values are marked
"Mixed" in the summary until edited to a shared value.
Strike and Expiry (d) display with no decimals; Entry Price, Reference Price, TheoV, and
Underlying Price display with two decimals, while Entry Vol (%) and Current Vol (%) display with two
decimals. The table
lot detail also includes a First Notice column for futures (CBOT: last business day of the month before delivery). Futures show Expiry (d) as days to the CBOT last trading day (the business day before the 15th
of the contract month). Expired options are left out of the analysis tabs. A Calendar Spread can put its far leg on a later
contract of the same commodity (e.g. Dec against Mar), each leg priced off its own underlying,
with an optional far strike. In the app, long 1 calendar means buy near, sell far; the Excel workbook's
Add Leg still sells near and buys far. In Add Leg, structures show one line per leg, each with its own Entry
Price; left blank, it is filled with that leg's TheoV when the structure is added.
It refreshes every 10 seconds with live prices (Underlying Price, TheoV). Option expiries are stored
as dates, so Expiry (d) counts down daily; expired options are greyed out until removed. Expiry values
are rendered as text in the editable table so numeric day counts and the "Expired" label work
consistently across Pandas versions. Remove selections are retained across table refreshes until
positions are removed or the portfolio is cleared or replaced. The Remove column is the rightmost
column in Current Positions. The positions editor refreshes when
an underlying price changes so incoming quotes appear correctly. Temporary quantity-edit diagnostics
write `POSITION_EDITOR_EVENT` and `POSITION_QUANTITY_COMPARE` (editor values versus stored position
quantity), `POSITION_QUANTITY_UPDATE`, `POSITION_QUANTITY_ASSIGNED`, and `PORTFOLIO_AUTOSAVE_*`
messages to the app logs; remove this instrumentation after the quantity revert issue is diagnosed.
Each detail editor stays stable across quantity edits and refreshes when a removal selection, Current
Vol, or underlying price changes so its shared selection and live quote values stay in sync.

Analysis tabs compute only when opened, using the latest prices ("Prices as of" caption). **Greek Summary by Tenor** shows current Greeks by underlying contract, expandable futures and
option tenors, and a total row. It uses each option's editable Current Vol and current underlying price.
The Payoff Diagram defaults to a -10% to +10% futures-price range in 2-percentage-point increments.
The **Price Ladder**,
**Vol Ladder** and **Decay Ladder**
each have their own tab, transposed (moves as columns, Greeks as rows with units), with PnL shaded
green/red by size and negative Greeks in red. Each ladder selector offers a commodity summary as well
as individual underlying contracts. A commodity summary isolates that commodity's positions and applies
price or volatility shocks in parallel across its contracts; the decay summary advances only that
commodity's positions through time. Price Ladder defaults are -20¢ to +20¢ in 2¢ steps, Vol Ladder
defaults are -5 to +5 vol points in 1-point steps (with step sizes adjustable to 0.1 points), and
Decay Ladder defaults to a 30-day horizon in 2-day steps. All values remain adjustable. Each ladder
includes an underlying comparison table with
one row per active contract, selectable commodity and measure, and columns matching the ladder's
scenario controls. A Total row sums the selected measure across underlyings; Futures Price in the
Price Ladder has no meaningful total. Options expiring during decay are exercised if in the money,
otherwise they lapse.

The **Breakevens** tab charts the book's P&L a chosen number of days forward (default 1) against the
selected contract's price, with the exact repriced curve, the quadratic Delta/Gamma/Theta estimate
(drawn within 10% of the price), green/red profit and loss shading, and every zero crossing marked.
Hover to read the P&L at any price. The table gives the nearest breakeven on each side as a price and
a move in cents; the chart range (default ±30¢) widens to include them.

All analysis charts are interactive (Altair, built in `vol_calculator/charts.py`; `plotting.py` keeps the
matplotlib versions for `demo.py`). Hover readouts run in the browser without rerunning the app.
**Payoff** draws each time step and the at-expiry payoff (dashed grey) over a ±% range (default 10%),
marks each curve's zero crossings, lists the nearest ones in a table, and lets you click a legend entry
to highlight a curve; invalid time steps show a warning. **Surface** is a heatmap with higher prices at
the top and today's cell outlined. **Strike Map** shows one commodity at a time, higher strikes on top,
with blank cells where there is no position. **Stress** shows one chart per Greek and stressed variable,
with today's unstressed point ringed.
The **Surface** view also offers a commodity summary, showing only that commodity's positions with
price and volatility moves applied in parallel across its contracts.

## Project layout

- `vol_calculator/` -- the core package (pricing, Greeks, scenarios, structures, plotting, CSV I/O).
- `app.py` -- the Streamlit UI, a thin layer over `vol_calculator`.
- `demo.py` -- a plain-script entry point (no UI) for the same demo portfolio.
- `PositionCalculator.ipynb` -- the original notebook, kept as a frozen historical reference; all
  further development happens in `vol_calculator/` and `app.py`, not the notebook.
- `launcher.py` -- bootstraps Streamlit programmatically; used to build a standalone executable via
  PyInstaller.
