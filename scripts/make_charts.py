"""Reproduce the key Excel pivot charts as static PNGs for the README.

The original analysis was done in Microsoft Excel (pivot tables + pivot charts)
and IBM Cognos Analytics. This script only re-draws a few of those views from the
same workbooks so they can be previewed directly on GitHub.

Usage:  python scripts/make_charts.py
"""
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DATA, IMG = ROOT / "data", ROOT / "images"
IMG.mkdir(exist_ok=True)

BLUE = "#2a78d6"
SERIES = {"Salish": "#2a78d6", "Beaufort": "#eb6834", "Labrador": "#1baf7a",
          "Hudson": "#eda100", "Champlain": "#e87ba4"}
INK, MUTED, GRID = "#1f1f1e", "#6b6a64", "#e6e5e0"

plt.rcParams.update({
    "font.size": 10, "axes.edgecolor": GRID, "axes.labelcolor": MUTED,
    "xtick.color": MUTED, "ytick.color": MUTED, "axes.titlecolor": INK,
    "axes.titleweight": "bold", "axes.titlesize": 12, "axes.titlelocation": "left",
    "axes.spines.top": False, "axes.spines.right": False,
    "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.8,
    "figure.facecolor": "white", "axes.axisbelow": True, "savefig.dpi": 150, "savefig.bbox": "tight",
})


def barh(series, title, xlabel, out, fmt="{:,.0f}"):
    s = series.sort_values()
    fig, ax = plt.subplots(figsize=(8, 0.42 * len(s) + 1.2))
    ax.barh(s.index.astype(str), s.values, color=BLUE, height=0.7)
    ax.grid(axis="y", visible=False)
    ax.set_title(title)
    ax.set_xlabel(xlabel)
    for y, v in enumerate(s.values):
        ax.text(v, y, " " + fmt.format(v), va="center", fontsize=9, color=INK)
    ax.set_xlim(0, s.max() * 1.15)
    fig.savefig(IMG / out)
    plt.close(fig)


# --- Car sales -----------------------------------------------------------------
sales = pd.read_excel(DATA / "CarSalesByModelEnd.xlsx", sheet_name="Sales by Model")

barh(sales.groupby("Dealer ID")["Quantity Sold"].sum(),
     "Quantity sold by dealer (2021-2025)", "Units sold",
     "excel_quantity_by_dealer.png")

sales["Period"] = sales["Date"].dt.to_period("M").dt.to_timestamp()
monthly = sales.pivot_table(index="Period", columns="Model", values="Profit", aggfunc="sum")
fig, ax = plt.subplots(figsize=(10, 4.8))
for model, color in SERIES.items():
    ax.plot(monthly.index, monthly[model] / 1e3, color=color, linewidth=2, label=model)
ax.axvspan(pd.Timestamp("2024-03-20"), pd.Timestamp("2024-06-15"), color=GRID, alpha=0.8, lw=0)
ax.text(pd.Timestamp("2024-05-01"), ax.get_ylim()[1] * 0.97, "Apr-Jun 2024\nanomaly",
        ha="center", va="top", fontsize=8.5, color=MUTED)
ax.set_title("Monthly profit by model (thousand $)")
ax.set_ylabel("Profit (k$)")
ax.legend(ncol=5, frameon=False, loc="upper left", bbox_to_anchor=(0, -0.1))
ax.set_xlim(monthly.index[0], monthly.index[-1] + pd.Timedelta(days=20))
fig.savefig(IMG / "excel_profit_by_date_and_model.png")
plt.close(fig)

# --- Montgomery County fleet ----------------------------------------------------
fleet = pd.read_excel(DATA / "Montgomery_Fleet_Equipment_Inventory_FA_PART_2_END.xlsx",
                      usecols="A:C")
barh(fleet.groupby("Department")["Equipment Count"].sum(),
     "Fleet equipment by department (pivot table 1)", "Equipment count",
     "excel_fleet_by_department.png")

print("Charts written to", IMG)
