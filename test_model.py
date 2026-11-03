"""Checks on the arithmetic, the decision rules and the agreement between Python and Excel.

    python -m pytest -q
"""
from pathlib import Path

import pandas as pd
import pytest

import make_data as md
import pricing as p
import suppliers as sup
from bid import assess_tender
from contracts import care_home_economics, contract_profitability
from sensitivity import tornado
from uplift import fee_uplift, portfolio

HERE = Path(__file__).resolve().parent


def test_oncost_rate_matches_hand_calculation():
    gross = 12.71 * 30 * 52.14
    expected = (0.15 * (gross - 5000) + 0.03 * (gross - 6240)) / gross
    assert p.employer_oncost_rate(p.Y2026) == pytest.approx(expected)


def test_build_up_adds_up_to_price():
    for length in p.VISIT_LENGTHS:
        for rural in (False, True):
            c = p.visit_cost(length, p.Y2026, rural)
            parts = sum(c[k] for k in p.COMPONENTS)
            assert parts == pytest.approx(c["price"])
            assert c["surplus"] / c["price"] == pytest.approx(p.Y2026.target_margin)


def test_short_and_rural_visits_cost_more_per_hour():
    urban = [p.per_hour(length, p.Y2026, False) for length in p.VISIT_LENGTHS]
    rural = [p.per_hour(length, p.Y2026, True) for length in p.VISIT_LENGTHS]
    assert urban == sorted(urban, reverse=True)
    assert all(r > u for r, u in zip(rural, urban))


def test_package_price_is_weighted_average_of_visit_prices():
    expected = sum(share * p.per_hour(length) for length, share in p.DEFAULT_MIX.items())
    assert p.blended() == pytest.approx(expected)
    assert sum(p.DEFAULT_MIX.values()) == pytest.approx(1.0)


def test_uplift_to_hold_restores_each_margin():
    table = fee_uplift(md.contracts(), offered=0.036)
    new_rate = table["rate"] * (1 + table["uplift_to_hold"])
    assert (1 - table["cost_next"] / new_rate).values == pytest.approx(table["margin"].values)
    assert portfolio(table)["margin_at_offer"] < portfolio(table)["margin_now"]


def test_contribution_reconciles_to_margin():
    t = contract_profitability(md.contracts())
    assert t["annual_contribution"].sum() == pytest.approx((t["annual_revenue"] * t["margin"]).sum())


def test_breakeven_occupancy_gives_zero_profit():
    homes = care_home_economics(md.care_homes())
    at_breakeven = md.care_homes().assign(occupancy=homes["breakeven_occupancy"])
    assert care_home_economics(at_breakeven)["weekly_profit"].abs().max() < 1e-6


def test_bid_rules():
    base = dict(md.TENDER)
    assert assess_tender({**base, "ceiling_rate": 40.0, "new_staff_needed": 10})["decision"] == "Bid"
    assert assess_tender({**base, "ceiling_rate": 40.0, "new_staff_needed": 99})["decision"] == "Bid with conditions"
    assert assess_tender({**base, "ceiling_rate": 25.0})["decision"] == "No bid"
    r = assess_tender(base)
    at_max = {**base, "rural_share": r["max_rural_share_for_target"]}
    assert assess_tender(at_max)["margin_at_ceiling"] == pytest.approx(p.Y2026.target_margin)


def test_supplier_savings_on_a_small_ledger():
    ledger = pd.DataFrame({
        "item": ["Gloves", "Gloves", sup.AGENCY_ITEM, sup.AGENCY_ITEM, None],
        "supplier": ["A", "B", "X", "Y", "Z"],
        "units": [100, 50, 10, 30, None], "unit_price": [4.0, 5.0, 20.0, 30.0, None],
        "amount": [400.0, 250.0, 200.0, 900.0, 1000.0], "category": ["PPE", "PPE", "Agency", "Agency", "Food"],
    })
    assert sup.price_variance_savings(ledger)["saving"].sum() == pytest.approx(50.0)
    agency = sup.agency_saving(ledger, share_converted=0.5, employed_cost=15.0)
    assert agency["avg_rate"] == pytest.approx(27.5)
    assert agency["saving"] == pytest.approx(40 * 0.5 * 12.5)


def test_tornado_is_symmetric_where_inputs_are():
    t = tornado().set_index("field")
    assert t.loc["pay_rate", "change_high"] == pytest.approx(-t.loc["pay_rate", "change_low"], rel=0.02)


def test_excel_workbook_matches_python():
    """Compares the recalculated workbook with the Python model. Skipped when the workbook has
    no cached values (Excel or LibreOffice must have recalculated and saved it first)."""
    from openpyxl import load_workbook
    path = HERE / "Care_Contract_Pricing_Model.xlsx"
    wb = load_workbook(path, data_only=True)
    if wb["Rate build-up"]["B28"].value is None:
        pytest.skip("workbook has not been recalculated")
    assert wb["Rate build-up"]["B28"].value == pytest.approx(p.blended(field="price"))
    assert wb["Rate build-up"]["B27"].value == pytest.approx(p.blended(field="cost"))
    table = fee_uplift(md.contracts(), offered=0.036)
    excel_margins = [wb["Contracts"][f"L{r}"].value for r in range(5, 17)]
    assert excel_margins == pytest.approx(table["margin"].tolist())
    assert wb["Contracts"]["Q17"].value == pytest.approx(portfolio(table)["uplift_to_hold"])
    homes = care_home_economics(md.care_homes())
    assert [wb["Care homes"][f"P{r}"].value for r in range(5, 10)] == pytest.approx(homes["breakeven_occupancy"].tolist())
    bid = assess_tender(md.TENDER)
    assert wb["Bid"]["B21"].value == pytest.approx(bid["margin_at_ceiling"])
    assert wb["Bid"]["B33"].value == bid["decision"]
