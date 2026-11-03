"""Writes the synthetic inputs (contracts, care homes, supplier ledger, tender).

They are invented for illustration and are not any provider's data. Council fees for the care
homes are set at Bury Council's published 2026/27 standard rates for each type of home.
"""
from pathlib import Path

import numpy as np
import pandas as pd

HERE = Path(__file__).resolve().parent

CONTRACTS = [
    # id, commissioner, type, 2025/26 rate per contact hour, hours a week, mix 15/30/45/60, rural share
    ("C01", "Council A", "Local authority", 26.60, 1450, 0.10, 0.48, 0.20, 0.22, 0.10),
    ("C02", "Council A", "Local authority", 27.90, 620, 0.08, 0.42, 0.22, 0.28, 0.45),
    ("C03", "Council B", "Local authority", 27.30, 980, 0.06, 0.44, 0.24, 0.26, 0.15),
    ("C04", "Council B", "Local authority", 30.20, 410, 0.05, 0.40, 0.25, 0.30, 0.55),
    ("C05", "Council C", "Local authority", 28.40, 1150, 0.05, 0.40, 0.25, 0.30, 0.05),
    ("C06", "Council C", "Local authority", 29.80, 300, 0.04, 0.36, 0.25, 0.35, 0.40),
    ("C07", "Council D", "Local authority", 29.10, 760, 0.04, 0.38, 0.25, 0.33, 0.20),
    ("C08", "Council D", "Local authority", 29.90, 520, 0.03, 0.35, 0.27, 0.35, 0.25),
    ("C09", "NHS ICB North", "NHS continuing healthcare", 31.20, 640, 0.00, 0.30, 0.30, 0.40, 0.20),
    ("C10", "NHS ICB South", "NHS continuing healthcare", 32.10, 380, 0.00, 0.25, 0.30, 0.45, 0.35),
    ("C11", "Private clients", "Self-funded", 33.50, 540, 0.00, 0.20, 0.30, 0.50, 0.15),
    ("C12", "Private clients", "Self-funded", 34.80, 260, 0.00, 0.15, 0.25, 0.60, 0.30),
]

HOMES = [
    # home, type, beds, occupancy, council-funded share, council fee (Bury 2026/27 standard),
    # self-funder fee, fixed cost a week, variable cost per resident a week
    ("Home 1", "Residential", 42, 0.91, 0.70, 741.33, 1050, 19500, 255),
    ("Home 2", "Residential dementia", 55, 0.86, 0.65, 777.81, 1150, 28000, 290),
    ("Home 3", "Nursing", 64, 0.93, 0.60, 826.04, 1295, 36000, 330),
    ("Home 4", "Nursing dementia", 48, 0.82, 0.75, 892.45, 1350, 27500, 345),
    ("Home 5", "Residential", 36, 0.95, 0.40, 741.33, 1100, 18000, 250),
]

TENDER = {
    "name": "Council E home care framework, lot 3 (illustrative)",
    "ceiling_rate": 32.20, "hours_per_week": 850,
    "share_15": 0.05, "share_30": 0.40, "share_45": 0.25, "share_60": 0.30,
    "rural_share": 0.40, "new_staff_needed": 30, "recruitment_capacity": 24, "contract_years": 3,
}

# category, supplier, item (None for spend tracked only by invoice), unit price, units a year
# (or annual spend when item is None)
SUPPLY_LINES = [
    ("Agency staffing", "Agency 1", "Agency care worker hour", 24.50, 9000),
    ("Agency staffing", "Agency 2", "Agency care worker hour", 26.80, 6500),
    ("Agency staffing", "Agency 3", "Agency care worker hour", 29.90, 3200),
    ("PPE", "Supplier P1", "Nitrile gloves, box of 100", 4.10, 6000),
    ("PPE", "Supplier P2", "Nitrile gloves, box of 100", 4.95, 3500),
    ("PPE", "Supplier P3", "Nitrile gloves, box of 100", 5.60, 1200),
    ("PPE", "Supplier P1", "Disposable aprons, roll of 200", 6.80, 1500),
    ("PPE", "Supplier P2", "Disposable aprons, roll of 200", 8.20, 900),
    ("Continence products", "Supplier C1", "Continence pads, pack of 28", 9.30, 4200),
    ("Continence products", "Supplier C2", "Continence pads, pack of 28", 10.80, 2600),
    ("Continence products", "Supplier C3", "Continence pads, pack of 28", 11.90, 900),
    ("Food", "Supplier F1", None, None, 180000),
    ("Food", "Supplier F2", None, None, 95000),
    ("Utilities", "Supplier U1", None, None, 140000),
    ("Maintenance", "Supplier M1", None, None, 45000),
    ("Maintenance", "Supplier M2", None, None, 38000),
    ("Maintenance", "Supplier M3", None, None, 22000),
    ("Insurance", "Supplier I1", None, None, 85000),
    ("IT and call monitoring", "Supplier T1", None, None, 60000),
    ("IT and call monitoring", "Supplier T2", None, None, 18000),
    ("Training", "Supplier TR1", None, None, 28000),
    ("Training", "Supplier TR2", None, None, 15000),
]
MONTHS = pd.period_range("2026-04", "2027-03", freq="M").strftime("%Y-%m").tolist()


def contracts() -> pd.DataFrame:
    cols = ["contract", "commissioner", "type", "rate", "hours_per_week",
            "share_15", "share_30", "share_45", "share_60", "rural_share"]
    return pd.DataFrame(CONTRACTS, columns=cols)


def care_homes() -> pd.DataFrame:
    cols = ["home", "home_type", "beds", "occupancy", "council_share", "council_fee",
            "self_fee", "fixed_cost_week", "variable_cost_resident"]
    return pd.DataFrame(HOMES, columns=cols)


def supplier_ledger(seed: int = 7) -> pd.DataFrame:
    """Twelve monthly lines per supplier and item, with modest month-to-month variation."""
    rng = np.random.default_rng(seed)
    rows = []
    for category, supplier, item, price, annual in SUPPLY_LINES:
        weights = rng.uniform(0.8, 1.2, len(MONTHS))
        weights = weights / weights.sum()
        for month, w in zip(MONTHS, weights):
            if item is None:
                rows.append((month, category, supplier, None, None, None, round(annual * w, 2)))
            else:
                units = int(round(annual * w))
                rows.append((month, category, supplier, item, units, price, round(units * price, 2)))
    return pd.DataFrame(rows, columns=["month", "category", "supplier", "item", "units", "unit_price", "amount"])


def write_all() -> None:
    contracts().to_csv(HERE / "synthetic_contracts.csv", index=False)
    care_homes().to_csv(HERE / "synthetic_care_homes.csv", index=False)
    supplier_ledger().to_csv(HERE / "synthetic_supplier_ledger.csv", index=False)
    pd.Series(TENDER).to_frame("value").rename_axis("field").to_csv(HERE / "synthetic_tender.csv")


if __name__ == "__main__":
    write_all()
