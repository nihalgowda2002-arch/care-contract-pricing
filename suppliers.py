"""Supplier spend: where the money goes, price differences on like-for-like items, agency hours."""
import pandas as pd

AGENCY_ITEM = "Agency care worker hour"


def by_category(ledger: pd.DataFrame) -> pd.DataFrame:
    out = ledger.groupby("category")["amount"].sum().sort_values(ascending=False).to_frame("spend")
    out["share"] = out["spend"] / out["spend"].sum()
    return out


def price_variance_savings(ledger: pd.DataFrame) -> pd.DataFrame:
    """Saving on each product bought from more than one supplier, if every unit had been bought
    at the lowest unit price paid in the year. Agency hours are handled separately."""
    items = ledger.dropna(subset=["item"])
    items = items[items["item"] != AGENCY_ITEM]
    rows = []
    for item, g in items.groupby("item"):
        lowest = g["unit_price"].min()
        rows.append({"item": item, "suppliers": g["supplier"].nunique(), "spend": g["amount"].sum(),
                     "units": g["units"].sum(), "lowest_unit_price": lowest,
                     "saving": g["amount"].sum() - g["units"].sum() * lowest})
    return pd.DataFrame(rows).sort_values("saving", ascending=False).reset_index(drop=True)


def agency_saving(ledger: pd.DataFrame, share_converted: float, employed_cost: float) -> dict:
    """Saving from filling a share of agency hours with employed care workers."""
    agency = ledger[ledger["item"] == AGENCY_ITEM]
    hours, spend = agency["units"].sum(), agency["amount"].sum()
    avg_rate = spend / hours
    return {"hours": hours, "spend": spend, "avg_rate": avg_rate,
            "saving": hours * share_converted * (avg_rate - employed_cost)}
