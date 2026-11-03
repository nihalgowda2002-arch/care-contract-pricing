"""Runs the whole analysis: writes the synthetic inputs, the results, the figures and the workbook.

    python build.py
"""
import json
from pathlib import Path

import pandas as pd

import charts
import excel_model
import make_data as md
import pricing as p
import suppliers as sup
from bid import assess_tender
from contracts import by_commissioner_type, care_home_economics
from sensitivity import tornado
from uplift import fee_uplift, portfolio

HERE = Path(__file__).resolve().parent
OFFERED_UPLIFT = 0.036   # Bury Council's 2026/27 uplift on its standard home care rate
AGENCY_SHARE = 0.30      # share of agency hours filled by employed care workers


def run() -> dict:
    md.write_all()
    contracts = pd.read_csv(HERE / "synthetic_contracts.csv")
    homes = pd.read_csv(HERE / "synthetic_care_homes.csv")
    ledger = pd.read_csv(HERE / "synthetic_supplier_ledger.csv")
    bury = pd.read_csv(HERE / "bury_fee_rates.csv")

    schedule = p.pricing_schedule()
    build = p.build_up()
    table = fee_uplift(contracts, OFFERED_UPLIFT)
    pf = portfolio(table)
    types = by_commissioner_type(table)
    homes_out = care_home_economics(homes)
    bid = assess_tender(md.TENDER)
    tor = tornado()
    pv = sup.price_variance_savings(ledger)
    agency = sup.agency_saving(ledger, AGENCY_SHARE, p.employed_cost_per_hour())
    b26 = bury.iloc[-1]
    pkg = {f: p.blended(field=f) for f in ("direct", "cost", "price")}

    results = {
        "package_2026_27": pkg,
        "price_per_hour_15_vs_60_min": [schedule.price_per_contact_hour[0], schedule.price_per_contact_hour[3]],
        "bury_2026_27": {
            "standard_rate": b26.homecare_standard, "enhanced_rate": b26.homecare_enhanced,
            "left_after_direct_cost": b26.homecare_standard - pkg["direct"],
            "share_of_overheads_covered": (b26.homecare_standard - pkg["direct"]) / (pkg["cost"] - pkg["direct"]),
            "margin_at_enhanced_rate": 1 - pkg["cost"] / b26.homecare_enhanced,
            "rate_growth_since_2021_22": b26.homecare_standard / bury.homecare_standard.iloc[0] - 1,
            "nlw_growth_since_2021_22": b26.nlw / bury.nlw.iloc[0] - 1,
        },
        "portfolio": pf,
        "by_commissioner_type": types[["margin", "share_of_revenue"]].to_dict(orient="index"),
        "care_homes": homes_out.set_index("home")[["margin", "breakeven_occupancy", "residents_short",
                                                   "annual_profit"]].to_dict(orient="index"),
        "tender": {k: v for k, v in bid.items()},
        "tornado": tor[["driver", "swing"]].to_dict(orient="records"),
        "suppliers": {"total_spend": ledger["amount"].sum(), "like_for_like_saving": pv["saving"].sum(),
                      "agency_saving": agency["saving"], "agency_avg_rate": agency["avg_rate"]},
    }
    (HERE / "results.json").write_text(json.dumps(results, indent=2, default=float))

    charts.rate_buildup(build, b26.homecare_standard, b26.homecare_enhanced)
    charts.contract_margins(table)
    charts.tornado_chart(tor, pkg["price"])
    charts.bury_vs_nlw(bury)
    excel_model.build(contracts, homes, ledger, md.TENDER, bury)
    return results


if __name__ == "__main__":
    r = run()
    print(json.dumps({k: r[k] for k in ("package_2026_27", "bury_2026_27", "portfolio")}, indent=1, default=float))
