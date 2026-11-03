"""Fee uplift: what each contract needs next year to hold its margin, against what is offered."""
import pandas as pd

from contracts import contract_cost_per_hour, contract_profitability
from pricing import Y2025, Y2026, Assumptions


def fee_uplift(contracts: pd.DataFrame, offered: float, now: Assumptions = Y2025,
               nxt: Assumptions = Y2026) -> pd.DataFrame:
    out = contract_profitability(contracts, now)
    out["cost_next"] = [contract_cost_per_hour(r, nxt) for _, r in out.iterrows()]
    out["uplift_to_hold"] = out["cost_next"] / out["cost_per_hour"] - 1
    out["rate_offered"] = out["rate"] * (1 + offered)
    out["margin_at_offer"] = 1 - out["cost_next"] / out["rate_offered"]
    out["uplift_to_target"] = out["cost_next"] / (1 - nxt.target_margin) / out["rate"] - 1
    out["gap_per_year"] = (out["rate"] * (1 + out["uplift_to_hold"]) - out["rate_offered"]) * out["annual_hours"]
    return out


def portfolio(table: pd.DataFrame) -> dict:
    w = table["rate"] * table["hours_per_week"]
    return {
        "margin_now": table["annual_contribution"].sum() / table["annual_revenue"].sum(),
        "uplift_to_hold": (w * (1 + table["uplift_to_hold"])).sum() / w.sum() - 1,
        "margin_at_offer": 1 - (table["cost_next"] * table["hours_per_week"]).sum()
        / (table["rate_offered"] * table["hours_per_week"]).sum(),
        "gap_per_year": table["gap_per_year"].sum(),
    }
