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

## Portfolio screen

The Market panel shows compact quote tables for the available commodities. Portfolio management
actions (clear, upload, import, and download) are grouped under the collapsed **Manage Portfolio**
section to keep the Add Leg form and position data visible. The Add Leg form labels the position size
as **Quantity**; its Underlying field is read-only (always the live price). Current Positions show
short quantities in red and a First Notice column for futures (CBOT: last business day of the month
before delivery). The positions table refreshes every 5 seconds with live prices (Underlying Price,
TheoV). Option expiries are stored as dates, so Expiry (d) counts down daily; expired options are
greyed out until removed.

Analysis tabs compute only when opened, using the latest prices ("Prices as of" caption). **Greek
Summary by Tenor** shows negatives in red. The **Price Ladder**, **Vol Ladder** and **Decay Ladder**
each have their own tab, transposed (moves as columns, Greeks as rows with units), with PnL shaded
green/red by size and negative Greeks in red. The Decay Ladder moves the book 20/40/60 days forward;
options expiring on the way are exercised if in the money, otherwise they lapse.

## Project layout

- `vol_calculator/` -- the core package (pricing, Greeks, scenarios, structures, plotting, CSV I/O).
- `app.py` -- the Streamlit UI, a thin layer over `vol_calculator`.
- `demo.py` -- a plain-script entry point (no UI) for the same demo portfolio.
- `PositionCalculator.ipynb` -- the original notebook, kept as a frozen historical reference; all
  further development happens in `vol_calculator/` and `app.py`, not the notebook.
- `launcher.py` -- bootstraps Streamlit programmatically; used to build a standalone executable via
  PyInstaller.
