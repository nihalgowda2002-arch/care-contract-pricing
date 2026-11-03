"""Bid or no-bid test for a home care tender."""
from contracts import MIX_COLUMNS
from pricing import WEEKS_PER_YEAR, Y2026, Assumptions, blended


def assess_tender(t: dict, a: Assumptions = Y2026, lowest_margin: float = 0.02) -> dict:
    mix = {length: t[col] for col, length in MIX_COLUMNS.items()}
    urban = blended(a, mix, 0.0, "cost")
    rural = blended(a, mix, 1.0, "cost")
    cost = (1 - t["rural_share"]) * urban + t["rural_share"] * rural
    ceiling = t["ceiling_rate"]
    margin = 1 - cost / ceiling
    annual_hours = t["hours_per_week"] * WEEKS_PER_YEAR
    recruit_ok = t["new_staff_needed"] <= t["recruitment_capacity"]
    if margin >= a.target_margin and recruit_ok:
        decision = "Bid"
    elif margin >= lowest_margin:
        decision = "Bid with conditions"
    else:
        decision = "No bid"
    return {
        "urban_cost": urban, "rural_cost": rural, "cost_per_hour": cost,
        "our_price": cost / (1 - a.target_margin), "margin_at_ceiling": margin,
        "walk_away_rate": cost / (1 - lowest_margin),
        "max_rural_share_for_target": (ceiling * (1 - a.target_margin) - urban) / (rural - urban),
        "annual_revenue": ceiling * annual_hours, "annual_contribution": (ceiling - cost) * annual_hours,
        "contract_contribution": (ceiling - cost) * annual_hours * t["contract_years"],
        "recruitment_ok": recruit_ok, "decision": decision,
    }
