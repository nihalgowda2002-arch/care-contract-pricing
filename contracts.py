"""Contract profitability for home care, and break-even occupancy for care homes."""
import pandas as pd

from pricing import WEEKS_PER_YEAR, Y2025, Assumptions, blended

MIX_COLUMNS = {"share_15": 0.25, "share_30": 0.5, "share_45": 0.75, "share_60": 1.0}


def mix_of(row: pd.Series) -> dict:
    return {length: row[col] for col, length in MIX_COLUMNS.items()}


def contract_cost_per_hour(row: pd.Series, a: Assumptions) -> float:
    return blended(a, mix_of(row), row["rural_share"], "cost")


def contract_profitability(contracts: pd.DataFrame, a: Assumptions = Y2025) -> pd.DataFrame:
    out = contracts.copy()
    out["cost_per_hour"] = [contract_cost_per_hour(r, a) for _, r in out.iterrows()]
    out["margin"] = 1 - out["cost_per_hour"] / out["rate"]
    out["annual_hours"] = out["hours_per_week"] * WEEKS_PER_YEAR
    out["annual_revenue"] = out["rate"] * out["annual_hours"]
    out["annual_contribution"] = (out["rate"] - out["cost_per_hour"]) * out["annual_hours"]
    return out


def by_commissioner_type(profit: pd.DataFrame) -> pd.DataFrame:
    g = profit.groupby("type")[["annual_revenue", "annual_contribution", "hours_per_week"]].sum()
    g["margin"] = g["annual_contribution"] / g["annual_revenue"]
    g["share_of_revenue"] = g["annual_revenue"] / g["annual_revenue"].sum()
    return g


def care_home_economics(homes: pd.DataFrame) -> pd.DataFrame:
    out = homes.copy()
    out["occupied"] = out["beds"] * out["occupancy"]
    out["avg_fee"] = out["council_share"] * out["council_fee"] + (1 - out["council_share"]) * out["self_fee"]
    out["weekly_revenue"] = out["occupied"] * out["avg_fee"]
    out["weekly_cost"] = out["fixed_cost_week"] + out["occupied"] * out["variable_cost_resident"]
    out["weekly_profit"] = out["weekly_revenue"] - out["weekly_cost"]
    out["margin"] = out["weekly_profit"] / out["weekly_revenue"]
    out["breakeven_occupancy"] = out["fixed_cost_week"] / (out["avg_fee"] - out["variable_cost_resident"]) / out["beds"]
    out["residents_short"] = (out["breakeven_occupancy"] * out["beds"] - out["occupied"]).clip(lower=0)
    out["annual_profit"] = out["weekly_profit"] * WEEKS_PER_YEAR
    return out
