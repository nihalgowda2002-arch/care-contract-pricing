"""Cost and price of a home care visit, built up from a care worker's hourly pay.

Every input is a field of Assumptions. Statutory figures are cited in the README; the rest are
stated assumptions for a typical branch and should be replaced with a provider's own data.
"""
from dataclasses import dataclass, replace

import pandas as pd

WEEKS_PER_YEAR = 52.14
VISIT_LENGTHS = (0.25, 0.5, 0.75, 1.0)

# Share of contact hours delivered in visits of each length, for a typical urban package.
DEFAULT_MIX = {0.25: 0.05, 0.5: 0.40, 0.75: 0.25, 1.0: 0.30}

NLW_2025_26 = 12.21      # National Living Wage from 1 April 2025
NLW_2026_27 = 12.71      # National Living Wage from 1 April 2026
CPI_NOV_2025 = 0.032     # CPI, 12 months to November 2025, as used in Bury's 2026/27 fee report
NONPAY_2025_26 = 1.80    # assumed non-pay overheads per contact hour in 2025/26
MILEAGE_2025_26 = 0.45   # HMRC approved mileage rate, first 10,000 miles, 2025/26
MILEAGE_2026_27 = 0.55   # raised from 6 April 2026 (announced 21 May 2026)


@dataclass(frozen=True)
class Assumptions:
    pay_rate: float = NLW_2026_27
    paid_hours_per_week: float = 30.0      # including paid holiday, used for employer on-costs
    holiday_rate: float = 0.1207           # 5.6 weeks accrued on hours worked
    noncontact_rate: float = 0.05          # training, supervision and meetings
    nic_rate: float = 0.15                 # employer National Insurance
    nic_threshold: float = 5000.0          # secondary threshold, per year
    pension_rate: float = 0.03             # employer auto-enrolment minimum
    pension_lower: float = 6240.0          # qualifying earnings band
    pension_upper: float = 50270.0
    mileage_rate: float = MILEAGE_2026_27  # HMRC approved mileage rate, per mile
    office_rate: float = 0.12              # branch and central staff, share of care worker cost
    nonpay_per_hour: float = NONPAY_2025_26 * (1 + CPI_NOV_2025)  # premises, IT, insurance, CQC fees
    target_margin: float = 0.05            # surplus as a share of price
    travel_minutes: float = 10.0
    miles: float = 2.5
    rural_travel_minutes: float = 18.0
    rural_miles: float = 5.5


Y2026 = Assumptions()
Y2025 = Assumptions(pay_rate=NLW_2025_26, nonpay_per_hour=NONPAY_2025_26, mileage_rate=MILEAGE_2025_26)


def with_changes(a: Assumptions = Y2026, **changes) -> Assumptions:
    return replace(a, **changes)


def employer_oncost_rate(a: Assumptions = Y2026) -> float:
    """Employer National Insurance and pension as a share of a care worker's gross pay."""
    gross = a.pay_rate * a.paid_hours_per_week * WEEKS_PER_YEAR
    nic = a.nic_rate * max(0.0, gross - a.nic_threshold)
    pension = a.pension_rate * min(max(0.0, gross - a.pension_lower), a.pension_upper - a.pension_lower)
    return (nic + pension) / gross


def employed_cost_per_hour(a: Assumptions = Y2026) -> float:
    """Cost of one paid care worker hour, with holiday and employer on-costs."""
    return a.pay_rate * (1 + a.holiday_rate) * (1 + employer_oncost_rate(a))


def visit_cost(visit_hours: float, a: Assumptions = Y2026, rural: bool = False) -> dict:
    """Cost and price of one visit, split into the parts a commissioner asks to see (pounds)."""
    travel = (a.rural_travel_minutes if rural else a.travel_minutes) / 60
    miles = a.rural_miles if rural else a.miles
    contact_pay = visit_hours * a.pay_rate
    travel_pay = travel * a.pay_rate
    noncontact_pay = (visit_hours + travel) * a.noncontact_rate * a.pay_rate
    basic = contact_pay + travel_pay + noncontact_pay
    holiday = basic * a.holiday_rate
    oncosts = (basic + holiday) * employer_oncost_rate(a)
    care_worker = basic + holiday + oncosts
    mileage = miles * a.mileage_rate
    office = care_worker * a.office_rate
    nonpay = visit_hours * a.nonpay_per_hour
    cost = care_worker + mileage + office + nonpay
    price = cost / (1 - a.target_margin)
    return {"contact_pay": contact_pay, "travel_pay": travel_pay, "noncontact_pay": noncontact_pay,
            "holiday": holiday, "oncosts": oncosts, "care_worker": care_worker, "mileage": mileage,
            "direct": care_worker + mileage, "office": office, "nonpay": nonpay, "cost": cost,
            "surplus": price - cost, "price": price}


def per_hour(visit_hours: float, a: Assumptions = Y2026, rural: bool = False, field: str = "price") -> float:
    return visit_cost(visit_hours, a, rural)[field] / visit_hours


def blended(a: Assumptions = Y2026, mix: dict = None, rural_share: float = 0.0, field: str = "price") -> float:
    """Per contact hour figure for a package, given the share of hours in each visit length."""
    mix = DEFAULT_MIX if mix is None else mix
    return sum(share * ((1 - rural_share) * per_hour(length, a, False, field)
                        + rural_share * per_hour(length, a, True, field))
               for length, share in mix.items())


def pricing_schedule(a: Assumptions = Y2026) -> pd.DataFrame:
    """Tender pricing schedule: cost and price per visit and per contact hour, urban and rural."""
    rows = []
    for rural in (False, True):
        for length in VISIT_LENGTHS:
            c = visit_cost(length, a, rural)
            rows.append({"visit_minutes": int(round(length * 60)), "area": "Rural" if rural else "Urban",
                         "cost_per_visit": c["cost"], "price_per_visit": c["price"],
                         "cost_per_contact_hour": c["cost"] / length,
                         "price_per_contact_hour": c["price"] / length,
                         "direct_cost_per_contact_hour": c["direct"] / length})
    return pd.DataFrame(rows)


COMPONENTS = ["contact_pay", "travel_pay", "noncontact_pay", "holiday", "oncosts",
              "mileage", "office", "nonpay", "surplus"]


def build_up(a: Assumptions = Y2026, mix: dict = None, rural: bool = False) -> pd.DataFrame:
    """Price per contact hour split into its parts, for each visit length and for the package."""
    mix = DEFAULT_MIX if mix is None else mix
    table = {}
    for length in VISIT_LENGTHS:
        c = visit_cost(length, a, rural)
        table[f"{int(round(length * 60))} min"] = {k: c[k] / length for k in COMPONENTS}
    out = pd.DataFrame(table)
    out["Package"] = sum(share * out[f"{int(round(length * 60))} min"] for length, share in mix.items())
    return out
