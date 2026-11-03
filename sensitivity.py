"""Which assumptions move the price of a typical package most (tornado table)."""
import pandas as pd

from pricing import Y2026, Assumptions, blended, with_changes

# label, field, how to get the low and high value from the base assumptions
DRIVERS = [
    ("Care worker pay +/-5%", "pay_rate", lambda v: v * 0.95, lambda v: v * 1.05),
    ("Travel time per visit +/-25%", "travel_minutes", lambda v: v * 0.75, lambda v: v * 1.25),
    ("Branch and central staff 9% to 15%", "office_rate", lambda v: 0.09, lambda v: 0.15),
    ("Non-pay overheads +/-25%", "nonpay_per_hour", lambda v: v * 0.75, lambda v: v * 1.25),
    ("Target margin 3% to 7%", "target_margin", lambda v: 0.03, lambda v: 0.07),
    ("Training and supervision time 3% to 7%", "noncontact_rate", lambda v: 0.03, lambda v: 0.07),
    ("Miles per visit +/-25%", "miles", lambda v: v * 0.75, lambda v: v * 1.25),
]


def tornado(a: Assumptions = Y2026) -> pd.DataFrame:
    base = blended(a)
    rows = []
    for label, field, low_fn, high_fn in DRIVERS:
        v = getattr(a, field)
        low, high = low_fn(v), high_fn(v)
        p_low = blended(with_changes(a, **{field: low}))
        p_high = blended(with_changes(a, **{field: high}))
        rows.append({"driver": label, "field": field, "low_value": low, "high_value": high,
                     "price_low": p_low, "price_high": p_high, "change_low": p_low - base,
                     "change_high": p_high - base, "swing": abs(p_high - p_low)})
    return pd.DataFrame(rows).sort_values("swing", ascending=False).reset_index(drop=True)
