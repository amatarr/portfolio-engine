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

## Project layout

- `vol_calculator/` -- the core package (pricing, Greeks, scenarios, structures, plotting, CSV I/O).
- `app.py` -- the Streamlit UI, a thin layer over `vol_calculator`.
- `demo.py` -- a plain-script entry point (no UI) for the same demo portfolio.
- `PositionCalculator.ipynb` -- the original notebook, kept as a frozen historical reference; all
  further development happens in `vol_calculator/` and `app.py`, not the notebook.
- `launcher.py` -- bootstraps Streamlit programmatically; used to build a standalone executable via
  PyInstaller.
