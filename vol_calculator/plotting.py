import matplotlib.pyplot as plt
import seaborn as sns

from vol_calculator.scenarios import (
    spot_vol_surface,
    strike_map,
    stress_table,
    stress_report,
    payoff_diagram,
    payoff_at_expiry,
)


_MONETARY_GREEK_UNITS = {
    "PnL": "$K",
    "Vega": "$K / vol pt",
    "Volga": "$K / vol pt²",
    "Vanna-V": "$K / (vol pt·¢)",
    "Theta": "$K / day",
}


def plot_surface(
    portfolio,
    market,
    contract=None,
    metric="PnL"
):

    surface = spot_vol_surface(
        portfolio,
        market,
        contract=contract,
        metric=metric
    )
    unit = _MONETARY_GREEK_UNITS.get(metric)
    if unit:
        surface = surface * (market.contract_multiplier / 1000)
    unit_suffix = f" ({unit})" if unit else ""

    plt.figure(figsize=(8, 5))

    sns.heatmap(
        surface,
        annot=True,
        fmt=".1f",
        cmap="RdYlGn" if metric == "PnL" else "coolwarm",
        center=0,
        annot_kws={"size": 8},
        cbar_kws={"label": unit} if unit else None,
    )

    plt.title(
        f"{metric}{unit_suffix} Heatmap\n"
        f"Future Spot Move (¢) vs Vol Shift (pts) -- {contract or 'single contract'}"
    )

    plt.xlabel("Vol Shift (pts)")
    plt.ylabel("Future Spot Move (¢)")

    plt.show()


def plot_strike_map(portfolio, market=None, weight="quantity"):

    grid = strike_map(portfolio, market=market, weight=weight)

    plt.figure(figsize=(8, 5))

    sns.heatmap(
        grid,
        annot=True,
        fmt=".1f",
        cmap="coolwarm",
        center=0,
        annot_kws={"size": 8},
    )

    label = "Net Quantity" if weight == "quantity" else "Net Delta"

    plt.title(f"Strike Map ({label})\nStrike vs Expiry (days)")
    plt.xlabel("Expiry (days)")
    plt.ylabel("Strike")

    plt.show()


def plot_stress(portfolio, market, greek, axis, magnitudes, contract=None):

    table = stress_table(portfolio, market, greek, axis, magnitudes, contract)
    unit = _MONETARY_GREEK_UNITS.get(greek.capitalize())
    if unit:
        table[greek.capitalize()] *= market.contract_multiplier / 1000

    plt.figure(figsize=(8, 5))

    plt.plot(table.index, table[greek.capitalize()], marker="o")
    plt.axhline(0, color="grey", linewidth=0.8)

    plt.title(f"{greek.capitalize()} vs {axis.capitalize()}")
    plt.xlabel(table.index.name)
    plt.ylabel(f"{greek.capitalize()} ({unit})" if unit else greek.capitalize())
    plt.grid(alpha=0.3)

    plt.show()


def plot_stress_report(
    portfolio,
    market,
    contract=None,
    price_magnitudes=(5, 10, 15, 20),
    time_magnitudes=(5, 10, 15, 20),
    vol_magnitudes=(1, 2, 3, 4)
):

    report = stress_report(
        portfolio,
        market,
        contract=contract,
        price_magnitudes=price_magnitudes,
        time_magnitudes=time_magnitudes,
        vol_magnitudes=vol_magnitudes
    )

    keys = list(report.keys())
    n = len(keys)
    ncols = 5
    nrows = -(-n // ncols)

    fig, axes = plt.subplots(nrows, ncols, figsize=(4 * ncols, 3.5 * nrows))
    axes = axes.flatten()

    for ax, key in zip(axes, keys):

        table = report[key]
        greek_col = table.columns[0]
        unit = _MONETARY_GREEK_UNITS.get(greek_col)
        if unit:
            table = table.copy()
            table[greek_col] *= market.contract_multiplier / 1000

        ax.plot(table.index, table[greek_col], marker="o")
        ax.axhline(0, color="grey", linewidth=0.8)
        ax.set_title(key.replace("_", " vs ").capitalize())
        ax.set_xlabel(table.index.name)
        ax.set_ylabel(f"{greek_col} ({unit})" if unit else greek_col)
        ax.grid(alpha=0.3)

    for ax in axes[n:]:
        ax.axis("off")

    plt.tight_layout()
    plt.show()


def plot_payoff(
    portfolio,
    market,
    time_steps,
    contract=None,
    low=-20,
    high=20,
    step=2,
    include_expiry=True
):

    diagram = payoff_diagram(
        portfolio,
        market,
        time_steps=time_steps,
        contract=contract,
        low=low,
        high=high,
        step=step
    )
    money_factor = market.contract_multiplier / 1000
    payoff_columns = [f"T+{days}d" for days in time_steps]
    diagram[payoff_columns] *= money_factor

    plt.figure(figsize=(8, 5))

    for days in time_steps:
        plt.plot(
            diagram["Futures Price"],
            diagram[f"T+{days}d"],
            marker="o",
            label=f"T+{days}d"
        )

    if include_expiry:

        expiry = payoff_at_expiry(
            portfolio,
            market,
            contract=contract,
            low=low,
            high=high,
            step=step
        )
        expiry["PnL"] *= money_factor

        plt.plot(
            expiry["Futures Price"],
            expiry["PnL"],
            color="black",
            linestyle="--",
            linewidth=2,
            label="At Expiry"
        )

    plt.axhline(0, color="grey", linewidth=0.8)
    plt.axvline(market.price_for(contract), color="grey", linewidth=0.8, linestyle=":")

    plt.title(f"Payoff Diagram -- {contract or 'single contract'}")
    plt.xlabel("Future Spot")
    plt.ylabel("PnL ($K)")
    plt.legend()
    plt.grid(alpha=0.3)

    plt.show()
