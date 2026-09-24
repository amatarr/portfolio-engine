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


def plot_surface(
    portfolio,
    market,
    metric="PnL"
):

    surface = spot_vol_surface(
        portfolio,
        market,
        metric=metric
    )

    plt.figure(figsize=(12, 8))

    sns.heatmap(
        surface,
        annot=True,
        fmt=".1f",
        cmap="RdYlGn" if metric == "PnL" else "coolwarm",
        center=0
    )

    plt.title(
        f"{metric} Heatmap\nSpot Move (%) vs Vol Shift (pts)"
    )

    plt.xlabel("Vol Shift (pts)")
    plt.ylabel("Spot Move (%)")

    plt.show()


def plot_strike_map(portfolio, market=None, weight="quantity"):

    grid = strike_map(portfolio, market=market, weight=weight)

    plt.figure(figsize=(10, max(4, 0.6 * len(grid.index))))

    sns.heatmap(
        grid,
        annot=True,
        fmt=".1f",
        cmap="coolwarm",
        center=0
    )

    label = "Net Quantity" if weight == "quantity" else "Net Delta"

    plt.title(f"Strike Map ({label})\nStrike vs Expiry (days)")
    plt.xlabel("Expiry (days)")
    plt.ylabel("Strike")

    plt.show()


def plot_stress(portfolio, market, greek, axis, magnitudes):

    table = stress_table(portfolio, market, greek, axis, magnitudes)

    plt.figure(figsize=(8, 5))

    plt.plot(table.index, table[greek.capitalize()], marker="o")
    plt.axhline(0, color="grey", linewidth=0.8)

    plt.title(f"{greek.capitalize()} vs {axis.capitalize()}")
    plt.xlabel(table.index.name)
    plt.ylabel(greek.capitalize())
    plt.grid(alpha=0.3)

    plt.show()


def plot_stress_report(
    portfolio,
    market,
    price_magnitudes=(5, 10, 15, 20),
    time_magnitudes=(5, 10, 15, 20),
    vol_magnitudes=(1, 2, 3, 4)
):

    report = stress_report(
        portfolio,
        market,
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

        ax.plot(table.index, table[greek_col], marker="o")
        ax.axhline(0, color="grey", linewidth=0.8)
        ax.set_title(key.replace("_", " vs ").capitalize())
        ax.set_xlabel(table.index.name)
        ax.set_ylabel(greek_col)
        ax.grid(alpha=0.3)

    for ax in axes[n:]:
        ax.axis("off")

    plt.tight_layout()
    plt.show()


def plot_payoff(
    portfolio,
    market,
    time_steps,
    low=-20,
    high=20,
    step=2,
    include_expiry=True
):

    diagram = payoff_diagram(
        portfolio,
        market,
        time_steps=time_steps,
        low=low,
        high=high,
        step=step
    )

    plt.figure(figsize=(10, 6))

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
            low=low,
            high=high,
            step=step
        )

        plt.plot(
            expiry["Futures Price"],
            expiry["PnL"],
            color="black",
            linestyle="--",
            linewidth=2,
            label="At Expiry"
        )

    plt.axhline(0, color="grey", linewidth=0.8)
    plt.axvline(market.futures_price, color="grey", linewidth=0.8, linestyle=":")

    plt.title("Payoff Diagram")
    plt.xlabel("Futures Price")
    plt.ylabel("PnL")
    plt.legend()
    plt.grid(alpha=0.3)

    plt.show()
