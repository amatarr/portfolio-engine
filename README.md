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

Monetary values in the Greek summary, analysis ladders (including their CSV downloads), surface,
stress charts, and payoff chart are displayed in thousands of dollars ($K), using the configured
contract multiplier. The Tenor and ladder headers include each measure's units:

- Delta is expressed as futures-equivalent contracts; Gamma is the change in Delta per 1¢/bu price move.
- Vega is $K per 1 volatility point (1 percentage point); Volga is $K per volatility point squared.
- Theta is $K per day; Rho is $K per 1 percentage point rate move.
- Vanna-D is the change in Delta per volatility point; Vanna-V is the change in Vega ($K per
  volatility point) per 1¢/bu price move.

The Greek Summary by Tenor shows a subtotal for each underlying futures contract, aggregating its
future and all options that reference it even when an option's listed month differs. Expand a
contract group to see the individual futures and option-expiry rows. Its column headers include units
and provide tooltips explaining each sensitivity and how the selected commodity affects the calculation
scope.

## Portfolio screen

The Market panel shows compact quote tables for the available commodities. Portfolio management
actions (clear, upload, import, and download) are grouped under the collapsed **Manage Portfolio**
section to keep the Add Leg form and position data visible. The Add Leg form labels the position size
as **Quantity**; its Underlying field is read-only (always the live price). Current Positions show
short quantities in red and support direct quantity edits, which are autosaved and refresh analysis
views. For structures, quantity edits apply to the selected leg; removing any selected structure leg
removes the full structure. Strike and Expiry (d) display with no decimals; Entry Price, TheoV, and
Underlying Price display with two decimals, while Vol (%) retains its existing precision. The table
also includes a First Notice column for futures (CBOT: last business day of the month before delivery).
It refreshes every 5 seconds with live prices (Underlying Price, TheoV). Option expiries are stored
as dates, so Expiry (d) counts down daily; expired options are greyed out until removed. Expiry values
are rendered as text in the editable table so numeric day counts and the "Expired" label work
consistently across Pandas versions.

Analysis tabs compute only when opened, using the latest prices ("Prices as of" caption). **Greek
Summary by Tenor** shows negatives in red. The **Price Ladder**, **Vol Ladder** and **Decay Ladder**
each have their own tab, transposed (moves as columns, Greeks as rows with units), with PnL shaded
green/red by size and negative Greeks in red. Each ladder selector offers a commodity summary as well
as individual underlying contracts. A commodity summary isolates that commodity's positions and applies
price or volatility shocks in parallel across its contracts; the decay summary advances only that
commodity's positions through time. Price Ladder defaults are -20¢ to +20¢ in 2¢ steps, Vol Ladder
defaults are -5 to +5 vol points in 1-point steps, and Decay Ladder defaults to a 30-day horizon in
2-day steps. All values remain adjustable. Options expiring during decay are exercised if in the
money, otherwise they lapse.

## Project layout

- `vol_calculator/` -- the core package (pricing, Greeks, scenarios, structures, plotting, CSV I/O).
- `app.py` -- the Streamlit UI, a thin layer over `vol_calculator`.
- `demo.py` -- a plain-script entry point (no UI) for the same demo portfolio.
- `PositionCalculator.ipynb` -- the original notebook, kept as a frozen historical reference; all
  further development happens in `vol_calculator/` and `app.py`, not the notebook.
- `launcher.py` -- bootstraps Streamlit programmatically; used to build a standalone executable via
  PyInstaller.
