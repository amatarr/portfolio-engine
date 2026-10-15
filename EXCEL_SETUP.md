# Position Analysis workbook — setup and how it works

`Position_Analysis.xlsm` is an Excel front end to a Python pricing engine
(Black-76 / American options on CBOT grain futures). Excel shows and
collects everything; Python does the maths when you press a button.

## What's in this folder

| Item | What it is |
|---|---|
| `Position_Analysis.xlsm` | The workbook. Open it from **this** folder, since the buttons look for the Python files next to it. |
| `xlwings_refresh.py` | What the buttons run: reads the book, prices it, writes the results into the workbook. |
| `vol_calculator/` | The pricing engine: instruments, Black-76 pricing, Greeks, scenarios, Daily Prop import, CME expiry calendar. |
| `vba/` | Readable copies of the macros inside the workbook: `AddLegUI.bas` (Add Leg autofill) and `PositionsUI.bas` (Remove selected legs). |
| `Daily prop/` | Put the Daily Prop file here (see below). |
| `requirements.txt` | Python packages to install. |

## One-time setup

1. **Excel: Microsoft 365.** The workbook uses dynamic-array formulas, which
   Excel 2016/2019 don't have.
2. **Python 3.10 or newer**, then from this folder:
   ```
   pip install -r requirements.txt
   xlwings addin install
   ```
3. **Tell xlwings which Python to use.** Create the file
   `%USERPROFILE%\.xlwings\xlwings.conf` with this single line, using your own
   path to `python.exe`:
   ```
   "INTERPRETER_WIN","C:\path\to\python.exe"
   ```
   (`where python` in a terminal shows the path.)
4. **Allow the macros.** If Excel shows a security bar, click *Enable Content*.
   If it says macros are blocked, right-click the file → Properties → tick
   *Unblock*.
5. **Live prices** come from CQGXL, which must be installed and signed in.
   Without it the quote cells at the top of Home don't update, and prices
   fall back to the Daily Prop file.

## Daily use

- **Daily Prop file:** drop `Daily_Prop_Bod_Detail_YYYY_MM_DD.csv` into
  `Daily prop/`. The newest one (by the date in its name) is used. With no
  file there, a refresh stops with "No Daily_Prop_Bod_Detail_*.csv found".
- **Refresh:** runs automatically when the workbook opens. To run it on
  demand: Alt+F8 → `RefreshModel`.
- **Changing Future Spot** re-runs the model around that contract.
- **Add Leg:** pick Type, Commodity and Contract, and every field fills
  with a suggestion: live underlying, at-the-money strike, the contract's
  real CME expiries, and default vols. Overwrite anything, then click
  *Add Leg*. Structures (spreads, straddles, butterflies…) are added as
  their individual Call/Put legs. Leaving *Entry price* blank uses the
  current Mark to Market.
- **Removing legs:** the Positions table is one formula, so its rows can't be
  deleted directly. Select the row(s), using any cell in each, and click
  *Remove selected legs*. Only manually added legs can be removed, because
  Daily Prop legs come back from the file on every refresh.
- **Clear Portfolio** (Alt+F8 → `ClearPortfolio`) removes all manually added
  legs at once.
- **Prices:** a live CQGXL quote always wins over the Daily Prop settlement.
  Daily Prop is only used for contracts the watchlist doesn't quote.

## How it's built (read before editing)

- **Home's layout belongs to Excel.** Format, move, recolour and relabel
  freely; a refresh never touches it.
- **Python writes only to the hidden `_Data` sheet** (right-click a sheet tab
  → Unhide to look). Don't edit it by hand, because every refresh rewrites
  it. Each block there has a name (`Positions`, `ContractQuotes`,
  `ExpiryCandidates`, …), and Home displays them with formulas such as
  `=IF(Positions="","",Positions)`.
- **Inputs are found by name, not address** (Formulas → Name Manager:
  `FutureSpot`, `InterestRate`, `AddLeg_Type`, `AddLeg_Strike1`, …). To move
  an input, **cut and paste** it so the name moves with it. Retyping it
  elsewhere breaks the link.
- **The Interest Rate cell** is a formula on the latest SOFR. Typing a
  number over it sets a manual rate.
- **The analytic sheets** (Greeks, Ladders, Heat Maps, Stress, Strike,
  Payoff, Breakevens) are still written directly by Python. Formatting you
  add there gets overwritten for now; moving them to the same `_Data`
  approach is the next step.

## Where to change what

| To change… | Edit… |
|---|---|
| Pricing, Greeks, scenarios, structures | `vol_calculator/` |
| What a refresh computes or writes to `_Data` | `xlwings_refresh.py` and `vol_calculator/xlwings_sync.py` |
| Add Leg defaults, labels, greyed-out rows | The `AddLegUI` module in Excel (Alt+F11). Then copy it back into `vba/AddLegUI.bas` (right-click the module → Export File) so the readable copy stays current. |
| Home's look and layout | Directly in Excel |

## Working on the shared copy

- **One person in the workbook at a time.** Two people editing a macro
  workbook on OneDrive at once produces conflict copies, and changes get
  lost. Close Excel when you're done.
- **Python changes:** say what you changed. The master copy of the code is
  in a GitHub repo, and changes made here need to be merged back into it.
