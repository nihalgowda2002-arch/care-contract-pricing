"""Figures for the README and the commercial brief."""
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

HERE = Path(__file__).resolve().parent
SLOTS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
INK, INK2, GRID, SURFACE = "#0b0b0b", "#52514e", "#e4e3df", "#fcfcfb"

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 9, "text.color": INK, "axes.labelcolor": INK2,
    "xtick.color": INK2, "ytick.color": INK2, "axes.edgecolor": GRID, "figure.facecolor": SURFACE,
    "axes.facecolor": SURFACE, "savefig.facecolor": SURFACE,
})


def _axes(ax, grid_axis="x"):
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color(GRID)
    ax.grid(axis=grid_axis, color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    ax.tick_params(length=0)


def _save(fig, name):
    fig.savefig(HERE / f"{name}.png", dpi=200, bbox_inches="tight", pad_inches=0.12)
    plt.close(fig)


def rate_buildup(build: pd.DataFrame, bury_standard: float, bury_enhanced: float):
    groups = [("Pay for contact time", ["contact_pay"]), ("Pay for travel time", ["travel_pay"]),
              ("Training and supervision", ["noncontact_pay"]), ("Holiday pay", ["holiday"]),
              ("Employer NI and pension", ["oncosts"]), ("Mileage", ["mileage"]),
              ("Overheads", ["office", "nonpay"]), ("Surplus (5% of price)", ["surplus"])]
    cols = ["60 min", "45 min", "30 min", "15 min"]
    fig, ax = plt.subplots(figsize=(7.6, 3.4))
    left = pd.Series(0.0, index=cols)
    for k, (label, parts) in enumerate(groups):
        vals = build.loc[parts, cols].sum()
        ax.barh(cols, vals, left=left, color=SLOTS[k], height=0.58, edgecolor=SURFACE, linewidth=1.5, label=label)
        left += vals
    for i, c in enumerate(cols):
        ax.text(left[c] + 0.5, i, f"£{left[c]:.2f}", va="center", fontsize=9, color=INK,
                bbox=dict(facecolor=SURFACE, edgecolor="none", pad=1.5))
    for x, label, dy in [(bury_standard, f"Bury standard £{bury_standard:.2f}", 3.55),
                         (bury_enhanced, f"Bury enhanced £{bury_enhanced:.2f}", 3.95)]:
        ax.axvline(x, color=INK2, linewidth=1, linestyle=(0, (3, 2)))
        ax.text(x + 0.3, dy - 0.15, label, fontsize=8, color=INK2, va="center")
    ax.set_ylim(-0.6, 4.1)
    ax.set_xlim(0, 46)
    ax.set_xlabel("£ per contact hour, urban visit, 2026/27")
    _axes(ax)
    ax.legend(ncol=4, fontsize=7.5, frameon=False, loc="upper center", bbox_to_anchor=(0.45, -0.2))
    ax.set_title("Price per contact hour by visit length", loc="left", fontsize=10.5, fontweight="bold")
    _save(fig, "fig1_rate_buildup")


def contract_margins(profit: pd.DataFrame):
    order = profit.sort_values("margin")
    types = ["Local authority", "NHS continuing healthcare", "Self-funded"]
    colour = dict(zip(types, SLOTS[:3]))
    fig, ax = plt.subplots(figsize=(7.6, 3.6))
    labels = [f"{r.contract} {r.commissioner}" for r in order.itertuples()]
    ax.barh(labels, order["margin"] * 100, color=[colour[t] for t in order["type"]], height=0.62,
            edgecolor=SURFACE, linewidth=1.5)
    for i, m in enumerate(order["margin"] * 100):
        ax.text(m + (0.5 if m >= 0 else -0.5), i, f"{m:.1f}%", va="center", ha="left" if m >= 0 else "right",
                fontsize=8, color=INK)
    ax.axvline(0, color=INK2, linewidth=1)
    ax.set_xlim(-12, 28)
    ax.set_xlabel("Margin on 2025/26 rates (%)")
    _axes(ax)
    handles = [plt.Rectangle((0, 0), 1, 1, color=colour[t]) for t in types]
    ax.legend(handles, types, frameon=False, fontsize=8, loc="lower right")
    ax.set_title("Margin by contract (synthetic portfolio)", loc="left", fontsize=10.5, fontweight="bold")
    _save(fig, "fig2_contract_margins")


def tornado_chart(tor: pd.DataFrame, base: float):
    tor = tor.iloc[::-1]
    fig, ax = plt.subplots(figsize=(7.6, 3.2))
    ax.barh(tor["driver"], tor["change_low"], color=SLOTS[0], height=0.6, label="Low case", edgecolor=SURFACE)
    ax.barh(tor["driver"], tor["change_high"], color=SLOTS[1], height=0.6, label="High case", edgecolor=SURFACE)
    for i, r in enumerate(tor.itertuples()):
        ax.text(r.change_high + 0.06, i, f"+£{r.change_high:.2f}", va="center", fontsize=8, color=INK)
        ax.text(r.change_low - 0.06, i, f"-£{abs(r.change_low):.2f}", va="center", ha="right", fontsize=8, color=INK)
    ax.axvline(0, color=INK2, linewidth=1)
    ax.set_xlim(-2.1, 2.1)
    ax.set_xlabel(f"Change in package price per contact hour from £{base:.2f} (2026/27)")
    _axes(ax)
    ax.legend(frameon=False, fontsize=8, loc="lower right")
    ax.set_title("What moves the price most", loc="left", fontsize=10.5, fontweight="bold")
    _save(fig, "fig3_tornado")


def bury_vs_nlw(bury: pd.DataFrame):
    rate = bury["homecare_standard"] / bury["homecare_standard"].iloc[0] * 100
    nlw = bury["nlw"] / bury["nlw"].iloc[0] * 100
    fig, ax = plt.subplots(figsize=(7.6, 3.0))
    x = bury["year"]
    ax.plot(x, rate, color=SLOTS[0], linewidth=2, marker="o", markersize=5, label="Bury standard home care rate")
    ax.plot(x, nlw, color=SLOTS[1], linewidth=2, marker="o", markersize=5, label="National Living Wage")
    ax.text(5.1, rate.iloc[-1], f"{rate.iloc[-1]:.0f}", va="center", fontsize=8.5, color=INK)
    ax.text(5.1, nlw.iloc[-1], f"{nlw.iloc[-1]:.0f}", va="center", fontsize=8.5, color=INK)
    ax.set_xlim(-0.3, 5.5)
    ax.set_ylabel("Index, 2021-22 = 100")
    _axes(ax, "y")
    ax.legend(frameon=False, fontsize=8, loc="upper left")
    ax.set_title("Bury's home care rate against the National Living Wage", loc="left", fontsize=10.5,
                 fontweight="bold")
    _save(fig, "fig4_bury_rate_vs_nlw")
