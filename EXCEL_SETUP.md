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
| `vba/` | Readable copies of the macros inside the workbook: `AddLegUI.bas` (Add Leg autofill) and `PositionsUI.bas` (Import Daily Prop, Remove selected legs), `AnalysisUI.bas` (per-sheet spot switches). |
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

- **The book** is one stored list of legs, imported or added by hand. A
  refresh prices exactly what's in it.
- **Import Daily Prop:** click the button and pick a
  `Daily_Prop_Bod_Detail_YYYY_MM_DD.csv` (the picker opens in `Daily prop/`).
  Its legs are **added** to the book, so importing the same file twice
  duplicates them. Option expiries are counted from the date in the file
  name and keep counting down on later days.
- **Refresh:** runs automatically when the workbook opens. To run it on
  demand: Alt+F8 → `RefreshModel`.
- **Spot switches:** Home's *Future Spot* picks the contract the analysis
  sheets shift. Each analysis sheet (except Strike) also has its own *Spot*
  dropdown in B1: *Follow Home* (the default) tracks Home, and picking a
  contract there recomputes that sheet only. C1 shows the contract in use.
  Changing Home's Future Spot recomputes only the sheets following it.
- **Add Leg:** pick Type, Commodity and Contract, and every field fills
  with a suggestion: live underlying, at-the-money strike, the contract's
  real CME expiries, and default vols. Overwrite anything, then click
  *Add Leg*. Structures (spreads, straddles, butterflies…) are added as
  their individual Call/Put legs. Leaving *Entry price* blank uses the
  current Mark to Market.
- **Removing legs:** the Positions table is one formula, so its rows can't be
  deleted directly. Select the row(s), using any cell in each, and click
  *Remove selected legs*. This works for any leg, imported or manual.
- **Clear Portfolio** (Alt+F8 → `ClearPortfolio`) empties the whole book.
- **Prices:** a live CQGXL quote always wins. A contract that isn't on the
  watchlist uses the settlement from the Daily Prop file its legs came from.
  A leg with neither stays in the book but isn't priced, and the status line
  names it.

## How it's built (read before editing)

- **Home's layout belongs to Excel.** Format, move, recolour and relabel
  freely; a refresh never touches it.
- **Python writes only to the hidden `_Data` sheet** (right-click a sheet tab
  → Unhide to look). Don't edit it by hand, because every refresh rewrites
  it. Each block there has a name (`Positions`, `ContractQuotes`,
  `ExpiryCandidates`, …), and Home displays them with formulas such as
  `=IF(Positions="","",Positions)`.
- **Inputs are found by name, not address** (Formulas → Name Manager:
  `FutureSpot`, `InterestRate`, `AddLeg_Type`, `AddLeg_Strike1`, …, and the
  live watchlist as `Watchlist_Corn` / `Watchlist_Soybean` / `Watchlist_Wheat`,
  each a symbol column plus a price column). To move any of them, **cut and
  paste** so the name moves along. Retyping elsewhere breaks the link. To
  add contracts to a watchlist block, widen its name in Name Manager.
- **The Interest Rate cell** is a formula on the latest SOFR. Typing a
  number over it sets a manual rate.
- **The analysis sheets** (Greeks, Ladders, Heat Maps, Stress, Strike,
  Payoff, Breakevens) work the same way, through the hidden `_Calc` sheet:
  every table there has a name (`Greeks_Tenor`, `Stress_delta_price`,
  `Payoff_Table`, …), and the sheets show them with formulas. Their layout,
  formatting and charts are yours to change. Keep the charts pointing at
  the cells the formulas fill, and don't rename these sheets: the spot
  switches find them by name.
- **Greeks by tenor** has one row per futures month, with options split by
  contract and expiry. Each row's Greeks are taken against that row's own
  contract, and Total is the sum of the rows. The summary table above it is
  against the sheet's spot.

## Where to change what

| To change… | Edit… |
|---|---|
| Pricing, Greeks, scenarios, structures | `vol_calculator/` |
| What a refresh computes or writes to `_Data` | `xlwings_refresh.py` and `vol_calculator/xlwings_sync.py` |
| What each analysis sheet computes (written to `_Calc`) | `vol_calculator/xlwings_analysis.py` |
| Add Leg defaults, labels, greyed-out rows | The `AddLegUI` module in Excel (Alt+F11). Then copy it back into `vba/AddLegUI.bas` (right-click the module → Export File) so the readable copy stays current. |
| Home's look and layout | Directly in Excel |

## Working on the shared copy

- **One person in the workbook at a time.** Two people editing a macro
  workbook on OneDrive at once produces conflict copies, and changes get
  lost. Close Excel when you're done.
- **Python changes:** say what you changed. The master copy of the code is
  in a GitHub repo, and changes made here need to be merged back into it.
