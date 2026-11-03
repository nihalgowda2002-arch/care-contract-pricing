"""Builds Care_Contract_Pricing_Model.xlsx: every result is an Excel formula that recalculates
when an input on the Inputs sheet changes. The Python modules compute the same numbers, and
test_model.py checks that the two agree."""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import make_data as md
import pandas as pd
from pricing import (CPI_NOV_2025, MILEAGE_2025_26, MILEAGE_2026_27, NLW_2025_26, NLW_2026_27,
                     NONPAY_2025_26, Y2026)
from sensitivity import tornado

HERE = Path(__file__).resolve().parent
OUT = HERE / "Care_Contract_Pricing_Model.xlsx"

F = "Arial"
BLUE = Font(name=F, size=10, color="0000FF")
BLACK = Font(name=F, size=10, color="000000")
GREEN = Font(name=F, size=10, color="008000")
BOLD = Font(name=F, size=10, bold=True)
TITLE = Font(name=F, size=13, bold=True)
NOTE = Font(name=F, size=9, italic=True, color="555555")
HEAD = PatternFill("solid", fgColor="E8EEF4")
KEY = PatternFill("solid", fgColor="FFFF00")
GBP = '£#,##0.00;(£#,##0.00);"-"'
GBP0 = '£#,##0;(£#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
PCT2 = '0.00%;(0.00%);"-"'
NUM = '#,##0;(#,##0);"-"'
NUM1 = '#,##0.0;(#,##0.0);"-"'
NUM2 = '#,##0.00;(#,##0.00);"-"'
WRAP = Alignment(wrap_text=True, vertical="top")

RB = "'Rate build-up'!"
COST26, COST25 = 22, 50          # cost per contact hour rows in the two build-up blocks
PKG26 = {"direct": 26, "cost": 27, "price": 28}
COLS = "BCDEFGHI"                 # 15, 30, 45, 60 min urban, then rural


def put(ws, ref, value, font=BLACK, fmt=None, fill=None, bold=False):
    c = ws[ref]
    c.value = value
    c.font = Font(name=F, size=10, bold=True, color=font.color) if bold else font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    return c


def header(ws, row, labels, start=1):
    for i, label in enumerate(labels):
        c = ws.cell(row=row, column=start + i, value=label)
        c.font, c.fill, c.alignment = BOLD, HEAD, WRAP


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def inputs_sheet(wb):
    ws = wb.create_sheet("Inputs")
    put(ws, "A1", "Inputs", TITLE)
    put(ws, "A2", "Blue cells are inputs you can change; black cells are formulas; yellow marks the "
                  "assumptions that matter most (see Sensitivity). Sources are on the About sheet.", NOTE)
    header(ws, 4, ["Year-specific inputs", "2025/26", "2026/27", "Unit", "Source or basis"])
    put(ws, "A5", "Care worker hourly pay")
    put(ws, "B5", NLW_2025_26, BLUE, GBP); put(ws, "C5", NLW_2026_27, BLUE, GBP)
    put(ws, "D5", "£ an hour"); put(ws, "E5", "National Living Wage, 21 and over (GOV.UK)")
    put(ws, "A6", "Non-pay overheads per contact hour")
    put(ws, "B6", NONPAY_2025_26, BLUE, GBP, KEY); put(ws, "C6", "=B6*(1+C25)", BLACK, GBP)
    put(ws, "D6", "£ a contact hour")
    put(ws, "E6", "Assumption: premises, IT, insurance, CQC fees; 2026/27 = 2025/26 plus CPI")

    header(ws, 8, ["Other inputs", "2025/26 if different", "2026/27 or both years", "Unit", "Source or basis"])
    rows = [
        (9, "Paid hours a week per care worker, including holiday", 30, NUM1, "hours", "Assumption, used for employer on-costs"),
        (10, "Weeks in a year", 52.14, NUM2, "weeks", "365.25 / 7"),
        (11, "Holiday pay accrual", 0.1207, PCT2, "of pay", "5.6 weeks statutory holiday (GOV.UK)"),
        (12, "Training, supervision and meetings", 0.05, PCT, "of visit and travel time", "Assumption"),
        (13, "Employer National Insurance rate", 0.15, PCT, "", "GOV.UK rates and thresholds 2026/27"),
        (14, "Employer NI secondary threshold", 5000, GBP0, "£ a year", "GOV.UK rates and thresholds 2026/27"),
        (15, "Employer pension contribution", 0.03, PCT, "of qualifying earnings", "Auto-enrolment minimum (The Pensions Regulator)"),
        (16, "Qualifying earnings, lower limit", 6240, GBP0, "£ a year", "DWP review for 2026/27 (GOV.UK)"),
        (17, "Qualifying earnings, upper limit", 50270, GBP0, "£ a year", "DWP review for 2026/27 (GOV.UK)"),
        (18, "Mileage paid per mile", MILEAGE_2026_27, GBP, "£ a mile", "HMRC approved mileage rate: 45p in 2025/26, 55p from 6 April 2026 (GOV.UK)"),
        (19, "Travel time per visit, urban", 10, NUM1, "minutes", "Assumption"),
        (20, "Miles per visit, urban", 2.5, NUM1, "miles", "Assumption"),
        (21, "Travel time per visit, rural", 18, NUM1, "minutes", "Assumption"),
        (22, "Miles per visit, rural", 5.5, NUM1, "miles", "Assumption"),
        (23, "Branch and central staff", 0.12, PCT, "of care worker cost", "Assumption"),
        (24, "Target margin", 0.05, PCT, "of price", "Assumption; average homecare EBITDA was 6.3% in Oct 2025 (Homecare Association)"),
        (25, "CPI, 12 months to November 2025", CPI_NOV_2025, PCT, "", "As cited in Bury Council's 2026/27 fee report"),
        (26, "Fee uplift offered for 2026/27", 0.036, PCT, "", "Bury Council's 2026/27 uplift on its standard home care rate"),
        (27, "Lowest margin we would bid at", 0.02, PCT, "of price", "Assumption"),
    ]
    key_rows = {19, 23}
    for r, label, value, fmt, unit, src in rows:
        put(ws, f"A{r}", label)
        put(ws, f"C{r}", value, BLUE, fmt, KEY if r in key_rows else None)
        put(ws, f"D{r}", unit)
        put(ws, f"E{r}", src)
    put(ws, "B18", MILEAGE_2025_26, BLUE, GBP)

    header(ws, 29, ["Calculated", "2025/26", "2026/27", "Unit", "Formula"])
    calc = [
        (30, "Annual gross pay", "={c}5*$C$9*$C$10", GBP0, "£ a year", "pay x paid hours x weeks"),
        (31, "Employer National Insurance", "=$C$13*MAX(0,{c}30-$C$14)", GBP0, "£ a year", "rate x pay above threshold"),
        (32, "Employer pension", "=$C$15*MIN(MAX(0,{c}30-$C$16),$C$17-$C$16)", GBP0, "£ a year", "rate x qualifying earnings"),
        (33, "Employer on-cost rate", "=({c}31+{c}32)/{c}30", PCT2, "of gross pay", "(NI + pension) / gross pay"),
        (34, "Cost of one paid care worker hour", "={c}5*(1+$C$11)*(1+{c}33)", GBP, "£ an hour", "pay x holiday x on-costs"),
    ]
    for r, label, f, fmt, unit, note in calc:
        put(ws, f"A{r}", label)
        put(ws, f"B{r}", f.format(c="B"), BLACK, fmt)
        put(ws, f"C{r}", f.format(c="C"), BLACK, fmt)
        put(ws, f"D{r}", unit); put(ws, f"E{r}", note)
    put(ws, "A35", "Pay increase"); put(ws, "C35", "=C5/B5-1", BLACK, PCT2)
    put(ws, "A36", "Non-pay overhead increase"); put(ws, "C36", "=C6/B6-1", BLACK, PCT2)

    header(ws, 38, ["Typical urban package", "", "Share of contact hours", "", "Note"])
    for i, (label, share) in enumerate([("15-minute visits", 0.05), ("30-minute visits", 0.40),
                                        ("45-minute visits", 0.25), ("60-minute visits", 0.30)]):
        put(ws, f"A{39 + i}", label)
        put(ws, f"C{39 + i}", share, BLUE, PCT)
    put(ws, "E39", "Assumption: share of hours, not of visits")
    put(ws, "A43", "Total (should be 100%)", bold=True); put(ws, "C43", "=SUM(C39:C42)", BLACK, PCT)
    widths(ws, {"A": 52, "B": 12, "C": 12, "D": 22, "E": 70})
    ws.freeze_panes = "A5"


def build_up_block(ws, t, year_col, title):
    pay, nonpay, oncost = f"Inputs!${year_col}$5", f"Inputs!${year_col}$6", f"Inputs!${year_col}$33"
    put(ws, f"A{t}", title, bold=True)
    labels = ["15 min urban", "30 min urban", "45 min urban", "60 min urban",
              "15 min rural", "30 min rural", "45 min rural", "60 min rural"]
    header(ws, t + 1, ["Per visit (£)"] + labels)
    names = ["Visit length (hours)", "Area", "Travel time (hours)", "Miles",
             "Pay for contact time", "Pay for travel time", "Pay for training, supervision and meetings",
             "Holiday pay", "Employer NI and pension", "Care worker cost", "Mileage",
             "Branch and central staff", "Non-pay overheads", "Total cost per visit", "Price per visit",
             "Surplus per visit", "Direct cost per contact hour (care worker and mileage)",
             "Cost per contact hour", "Price per contact hour"]
    for i, n in enumerate(names):
        put(ws, f"A{t + 2 + i}", n, bold=n in ("Care worker cost", "Total cost per visit", "Price per contact hour"))
    for j, col in enumerate(COLS):
        rural = j >= 4
        r = lambda k: f"{col}{t + k}"
        put(ws, r(2), [0.25, 0.5, 0.75, 1.0][j % 4], BLUE, NUM2)
        put(ws, r(3), "Rural" if rural else "Urban")
        put(ws, r(4), f"=Inputs!$C${21 if rural else 19}/60", BLACK, '0.000')
        put(ws, r(5), f"=Inputs!$C${22 if rural else 20}", GREEN, NUM1)
        put(ws, r(6), f"={r(2)}*{pay}", BLACK, GBP)
        put(ws, r(7), f"={r(4)}*{pay}", BLACK, GBP)
        put(ws, r(8), f"=({r(2)}+{r(4)})*Inputs!$C$12*{pay}", BLACK, GBP)
        put(ws, r(9), f"=SUM({r(6)}:{r(8)})*Inputs!$C$11", BLACK, GBP)
        put(ws, r(10), f"=SUM({r(6)}:{r(9)})*{oncost}", BLACK, GBP)
        put(ws, r(11), f"=SUM({r(6)}:{r(10)})", BLACK, GBP)
        put(ws, r(12), f"={r(5)}*Inputs!${year_col}$18", BLACK, GBP)
        put(ws, r(13), f"={r(11)}*Inputs!$C$23", BLACK, GBP)
        put(ws, r(14), f"={r(2)}*{nonpay}", BLACK, GBP)
        put(ws, r(15), f"={r(11)}+{r(12)}+{r(13)}+{r(14)}", BLACK, GBP)
        put(ws, r(16), f"={r(15)}/(1-Inputs!$C$24)", BLACK, GBP)
        put(ws, r(17), f"={r(16)}-{r(15)}", BLACK, GBP)
        put(ws, r(18), f"=({r(11)}+{r(12)})/{r(2)}", BLACK, GBP)
        put(ws, r(19), f"={r(15)}/{r(2)}", BLACK, GBP)
        put(ws, r(20), f"={r(16)}/{r(2)}", BLACK, GBP)
    put(ws, f"A{t + 22}", "Typical urban package (visit mix on Inputs), per contact hour", bold=True)
    for k, (label, src) in enumerate([("Direct cost", t + 18), ("Total cost", t + 19), ("Price at target margin", t + 20)]):
        put(ws, f"A{t + 23 + k}", label)
        put(ws, f"B{t + 23 + k}", f"=Inputs!$C$39*B{src}+Inputs!$C$40*C{src}+Inputs!$C$41*D{src}+Inputs!$C$42*E{src}",
            BLACK, GBP)


def rate_sheet(wb):
    ws = wb.create_sheet("Rate build-up")
    put(ws, "A1", "Rate build-up: cost and price of a visit", TITLE)
    build_up_block(ws, 3, "C", "2026/27: prices for new work and tenders")
    build_up_block(ws, 31, "B", "2025/26: the year the contract rates on the Contracts sheet were set")
    widths(ws, {"A": 50, **{c: 13 for c in COLS}})
    ws.freeze_panes = "B5"


def contract_cost_formula(row, cost_row):
    parts = []
    for share_col, (u, r) in zip("FGHI", [("B", "F"), ("C", "G"), ("D", "H"), ("E", "I")]):
        parts.append(f"{share_col}{row}*((1-$J{row})*{RB}${u}${cost_row}+$J{row}*{RB}${r}${cost_row})")
    return "=" + "+".join(parts)


def contracts_sheet(wb, contracts):
    ws = wb.create_sheet("Contracts")
    put(ws, "A1", "Contract profitability and the 2026/27 fee uplift", TITLE)
    put(ws, "A2", "Synthetic contracts for illustration. Rates are 2025/26 rates per contact hour; costs come "
                  "from the Rate build-up sheet for each contract's visit mix and rural share.", NOTE)
    cols = ["Contract", "Commissioner", "Type", "Rate 2025/26 (£ a contact hour)", "Hours a week",
            "Share 15 min", "Share 30 min", "Share 45 min", "Share 60 min", "Rural share",
            "Cost per hour 2025/26", "Margin 2025/26", "Annual hours", "Annual revenue", "Annual contribution",
            "Cost per hour 2026/27", "Uplift needed to hold margin", "Rate with offered uplift",
            "Margin at offered uplift", "Uplift needed for target margin", "Income gap at offered uplift (£ a year)"]
    header(ws, 4, cols)
    first, last = 5, 4 + len(contracts)
    for i, c in enumerate(contracts.itertuples(index=False)):
        r = first + i
        put(ws, f"A{r}", c.contract, BLUE); put(ws, f"B{r}", c.commissioner, BLUE); put(ws, f"C{r}", c.type, BLUE)
        put(ws, f"D{r}", c.rate, BLUE, GBP); put(ws, f"E{r}", c.hours_per_week, BLUE, NUM)
        for col, v in zip("FGHIJ", [c.share_15, c.share_30, c.share_45, c.share_60, c.rural_share]):
            put(ws, f"{col}{r}", v, BLUE, PCT)
        put(ws, f"K{r}", contract_cost_formula(r, COST25), BLACK, GBP)
        put(ws, f"L{r}", f"=1-K{r}/D{r}", BLACK, PCT)
        put(ws, f"M{r}", f"=E{r}*Inputs!$C$10", BLACK, NUM)
        put(ws, f"N{r}", f"=D{r}*M{r}", BLACK, GBP0)
        put(ws, f"O{r}", f"=(D{r}-K{r})*M{r}", BLACK, GBP0)
        put(ws, f"P{r}", contract_cost_formula(r, COST26), BLACK, GBP)
        put(ws, f"Q{r}", f"=P{r}/K{r}-1", BLACK, PCT2)
        put(ws, f"R{r}", f"=D{r}*(1+Inputs!$C$26)", BLACK, GBP)
        put(ws, f"S{r}", f"=1-P{r}/R{r}", BLACK, PCT)
        put(ws, f"T{r}", f"=P{r}/(1-Inputs!$C$24)/D{r}-1", BLACK, PCT)
        put(ws, f"U{r}", f"=(D{r}*(1+Q{r})-R{r})*M{r}", BLACK, GBP0)
    t = last + 1
    rng = lambda col: f"{col}{first}:{col}{last}"
    put(ws, f"A{t}", "Portfolio", bold=True)
    put(ws, f"E{t}", f"=SUM({rng('E')})", BLACK, NUM, bold=True)
    put(ws, f"L{t}", f"=O{t}/N{t}", BLACK, PCT, bold=True)
    for col in "MNOU":
        put(ws, f"{col}{t}", f"=SUM({rng(col)})", BLACK, NUM if col == "M" else GBP0, bold=True)
    put(ws, f"Q{t}", f"=SUMPRODUCT({rng('D')},{rng('E')},1+{rng('Q')})/SUMPRODUCT({rng('D')},{rng('E')})-1", BLACK, PCT2, bold=True)
    put(ws, f"S{t}", f"=1-SUMPRODUCT({rng('P')},{rng('E')})/SUMPRODUCT({rng('R')},{rng('E')})", BLACK, PCT, bold=True)
    put(ws, f"A{t + 1}", "Portfolio margin is contribution over revenue; portfolio uplift is weighted by revenue.", NOTE)

    s = t + 3
    header(ws, s, ["By commissioner type", "", "", "Annual revenue", "Annual contribution", "Margin", "Share of revenue"])
    for k, typ in enumerate(["Local authority", "NHS continuing healthcare", "Self-funded"]):
        r = s + 1 + k
        put(ws, f"A{r}", typ)
        put(ws, f"D{r}", f"=SUMIFS({rng('N')},{rng('C')},A{r})", BLACK, GBP0)
        put(ws, f"E{r}", f"=SUMIFS({rng('O')},{rng('C')},A{r})", BLACK, GBP0)
        put(ws, f"F{r}", f"=E{r}/D{r}", BLACK, PCT)
        put(ws, f"G{r}", f"=D{r}/$N${t}", BLACK, PCT)
    widths(ws, {"A": 11, "B": 17, "C": 25, **{get_column_letter(i): 12 for i in range(4, 22)}})
    ws.row_dimensions[4].height = 54
    ws.freeze_panes = "D5"
    return t


def homes_sheet(wb, homes):
    ws = wb.create_sheet("Care homes")
    put(ws, "A1", "Care home occupancy and break-even", TITLE)
    put(ws, "A2", "Synthetic homes for illustration. Council fees are set at Bury Council's published 2026/27 "
                  "standard weekly rate for each type of home.", NOTE)
    cols = ["Home", "Type", "Beds", "Occupancy", "Council-funded share", "Council fee (£ a week)",
            "Self-funder fee (£ a week)", "Fixed cost (£ a week)", "Variable cost per resident (£ a week)",
            "Occupied beds", "Average fee", "Weekly revenue", "Weekly cost", "Weekly profit", "Margin",
            "Break-even occupancy", "Residents short of break-even", "Annual profit"]
    header(ws, 4, cols)
    first = 5
    for i, h in enumerate(homes.itertuples(index=False)):
        r = first + i
        put(ws, f"A{r}", h.home, BLUE); put(ws, f"B{r}", h.home_type, BLUE)
        put(ws, f"C{r}", h.beds, BLUE, NUM); put(ws, f"D{r}", h.occupancy, BLUE, PCT, KEY)
        put(ws, f"E{r}", h.council_share, BLUE, PCT); put(ws, f"F{r}", h.council_fee, BLUE, GBP)
        put(ws, f"G{r}", h.self_fee, BLUE, GBP0); put(ws, f"H{r}", h.fixed_cost_week, BLUE, GBP0)
        put(ws, f"I{r}", h.variable_cost_resident, BLUE, GBP0)
        put(ws, f"J{r}", f"=C{r}*D{r}", BLACK, NUM1)
        put(ws, f"K{r}", f"=E{r}*F{r}+(1-E{r})*G{r}", BLACK, GBP)
        put(ws, f"L{r}", f"=J{r}*K{r}", BLACK, GBP0)
        put(ws, f"M{r}", f"=H{r}+J{r}*I{r}", BLACK, GBP0)
        put(ws, f"N{r}", f"=L{r}-M{r}", BLACK, GBP0)
        put(ws, f"O{r}", f"=N{r}/L{r}", BLACK, PCT)
        put(ws, f"P{r}", f"=H{r}/(K{r}-I{r})/C{r}", BLACK, PCT)
        put(ws, f"Q{r}", f"=MAX(0,P{r}*C{r}-J{r})", BLACK, NUM1)
        put(ws, f"R{r}", f"=N{r}*Inputs!$C$10", BLACK, GBP0)
    last = first + len(homes) - 1
    t = last + 1
    put(ws, f"A{t}", "All homes", bold=True)
    for col, fmt in [("C", NUM), ("J", NUM1), ("L", GBP0), ("M", GBP0), ("N", GBP0), ("R", GBP0)]:
        put(ws, f"{col}{t}", f"=SUM({col}{first}:{col}{last})", BLACK, fmt, bold=True)
    put(ws, f"D{t}", f"=J{t}/C{t}", BLACK, PCT, bold=True)
    put(ws, f"O{t}", f"=N{t}/L{t}", BLACK, PCT, bold=True)
    widths(ws, {"A": 10, "B": 20, **{get_column_letter(i): 12 for i in range(3, 19)}})
    ws.row_dimensions[4].height = 54
    ws.freeze_panes = "C5"


def bid_sheet(wb, t):
    ws = wb.create_sheet("Bid")
    put(ws, "A1", "Bid or no-bid test for a home care tender", TITLE)
    put(ws, "A2", "Illustrative tender. Costs use the 2026/27 block of the Rate build-up sheet.", NOTE)
    header(ws, 3, ["Tender inputs", "Value"])
    ins = [(4, "Tender", t["name"], None), (5, "Ceiling rate (£ a contact hour)", t["ceiling_rate"], GBP),
           (6, "Hours a week", t["hours_per_week"], NUM), (7, "Share of hours in 15-minute visits", t["share_15"], PCT),
           (8, "Share in 30-minute visits", t["share_30"], PCT), (9, "Share in 45-minute visits", t["share_45"], PCT),
           (10, "Share in 60-minute visits", t["share_60"], PCT), (11, "Rural share of hours", t["rural_share"], PCT),
           (12, "New care workers needed", t["new_staff_needed"], NUM),
           (13, "Care workers we can recruit before the start date", t["recruitment_capacity"], NUM),
           (14, "Contract length (years)", t["contract_years"], NUM)]
    for r, label, v, fmt in ins:
        put(ws, f"A{r}", label); put(ws, f"B{r}", v, BLUE, fmt)
    header(ws, 16, ["Results", "Value"])
    c = COST26
    res = [
        (17, "Urban cost per contact hour", f"=B7*{RB}$B${c}+B8*{RB}$C${c}+B9*{RB}$D${c}+B10*{RB}$E${c}", GBP),
        (18, "Rural cost per contact hour", f"=B7*{RB}$F${c}+B8*{RB}$G${c}+B9*{RB}$H${c}+B10*{RB}$I${c}", GBP),
        (19, "Blended cost per contact hour", "=(1-B11)*B17+B11*B18", GBP),
        (20, "Our price at target margin", "=B19/(1-Inputs!$C$24)", GBP),
        (21, "Margin at the ceiling rate", "=1-B19/B5", PCT),
        (22, "Walk-away rate (lowest acceptable margin)", "=B19/(1-Inputs!$C$27)", GBP),
        (23, "Highest rural share that still earns the target margin", "=(B5*(1-Inputs!$C$24)-B17)/(B18-B17)", PCT),
        (24, "Annual revenue at the ceiling rate", "=B5*B6*Inputs!$C$10", GBP0),
        (25, "Annual contribution at the ceiling rate", "=(B5-B19)*B6*Inputs!$C$10", GBP0),
        (26, "Contribution over the contract", "=B25*B14", GBP0),
    ]
    for r, label, f, fmt in res:
        put(ws, f"A{r}", label); put(ws, f"B{r}", f, BLACK, fmt)
    header(ws, 28, ["Checks", "Pass?"])
    put(ws, "A29", "Margin at ceiling meets the target margin"); put(ws, "B29", "=B21>=Inputs!$C$24")
    put(ws, "A30", "Margin at ceiling is at least the lowest we would bid at"); put(ws, "B30", "=B21>=Inputs!$C$27")
    put(ws, "A31", "Recruitment covers the new care workers needed"); put(ws, "B31", "=B12<=B13")
    put(ws, "A33", "Decision", bold=True)
    put(ws, "B33", '=IF(AND(B29,B31),"Bid",IF(B30,"Bid with conditions","No bid"))', BLACK, bold=True)
    put(ws, "A34", "What it means")
    put(ws, "B34", '=IF(B33="Bid","Margin and staffing clear both tests",IF(B33="Bid with conditions",'
                   '"Bid only with conditions, such as a rural travel supplement, fewer rural hours or a phased start",'
                   '"The ceiling rate does not cover cost at the lowest margin we would accept"))')
    widths(ws, {"A": 55, "B": 60})


def price_formula(r, length):
    return (f"(C{r}*({length}+D{r}/60)*(1+F{r})*(1+Inputs!$C$11)*(1+K{r})*(1+G{r})"
            f"+E{r}*Inputs!$C$18+H{r}*{length})/((1-I{r})*{length})")


def sensitivity_sheet(wb):
    ws = wb.create_sheet("Sensitivity")
    put(ws, "A1", "Sensitivity: price of the typical urban package, 2026/27", TITLE)
    put(ws, "A2", "Each row changes one input and keeps the rest at their Inputs values. Rows are ordered by "
                  "swing at the published inputs.", NOTE)
    header(ws, 4, ["Driver", "Case", "Pay (£ an hour)", "Travel (minutes)", "Miles per visit",
                   "Training and supervision", "Branch and central staff", "Non-pay (£ a contact hour)",
                   "Target margin", "Annual gross pay", "On-cost rate", "Price, 15 min", "Price, 30 min",
                   "Price, 45 min", "Price, 60 min", "Package price", "Change from base"])
    base = {"pay_rate": "=Inputs!$C$5", "travel_minutes": "=Inputs!$C$19", "miles": "=Inputs!$C$20",
            "noncontact_rate": "=Inputs!$C$12", "office_rate": "=Inputs!$C$23",
            "nonpay_per_hour": "=Inputs!$C$6", "target_margin": "=Inputs!$C$24"}
    col_of = dict(zip(base, "CDEFGHI"))
    fmt_of = {"C": GBP, "D": NUM1, "E": NUM2, "F": PCT, "G": PCT, "H": GBP, "I": PCT}
    override = {
        "pay_rate": ("=Inputs!$C$5*0.95", "=Inputs!$C$5*1.05"),
        "travel_minutes": ("=Inputs!$C$19*0.75", "=Inputs!$C$19*1.25"),
        "office_rate": (0.09, 0.15),
        "nonpay_per_hour": ("=Inputs!$C$6*0.75", "=Inputs!$C$6*1.25"),
        "target_margin": (0.03, 0.07),
        "noncontact_rate": (0.03, 0.07),
        "miles": ("=Inputs!$C$20*0.75", "=Inputs!$C$20*1.25"),
    }
    order = tornado()
    cases = [("Base", "Base", None, None)]
    for _, d in order.iterrows():
        cases += [(d["driver"], "Low", d["field"], 0), (d["driver"], "High", d["field"], 1)]
    for i, (label, case, field, k) in enumerate(cases):
        r = 5 + i
        put(ws, f"A{r}", label if case != "High" else "")
        put(ws, f"B{r}", case)
        for f_name, col in col_of.items():
            if f_name == field:
                v = override[field][k]
                put(ws, f"{col}{r}", v, BLUE, fmt_of[col], KEY)
            else:
                put(ws, f"{col}{r}", base[f_name], GREEN, fmt_of[col])
        put(ws, f"J{r}", f"=C{r}*Inputs!$C$9*Inputs!$C$10", BLACK, GBP0)
        put(ws, f"K{r}", f"=(Inputs!$C$13*MAX(0,J{r}-Inputs!$C$14)+Inputs!$C$15*MIN(MAX(0,J{r}-Inputs!$C$16),"
                         f"Inputs!$C$17-Inputs!$C$16))/J{r}", BLACK, PCT2)
        for col, length in zip("LMNO", ["0.25", "0.5", "0.75", "1"]):
            put(ws, f"{col}{r}", "=" + price_formula(r, length), BLACK, GBP)
        put(ws, f"P{r}", f"=Inputs!$C$39*L{r}+Inputs!$C$40*M{r}+Inputs!$C$41*N{r}+Inputs!$C$42*O{r}", BLACK, GBP)
        put(ws, f"Q{r}", f"=P{r}-$P$5", BLACK, GBP)
    last = 4 + len(cases)
    chk = last + 2
    put(ws, f"A{chk}", "Check: base package price equals the Rate build-up sheet")
    put(ws, f"B{chk}", f"=ROUND(P5-{RB}$B${PKG26['price']},9)=0")
    s = chk + 2
    header(ws, s, ["Tornado table", "Price change, low case", "Price change, high case", "Swing"])
    for j in range(len(order)):
        r, lo, hi = s + 1 + j, 6 + 2 * j, 7 + 2 * j
        put(ws, f"A{r}", f"=A{lo}", GREEN)
        put(ws, f"B{r}", f"=Q{lo}", BLACK, GBP)
        put(ws, f"C{r}", f"=Q{hi}", BLACK, GBP)
        put(ws, f"D{r}", f"=ABS(C{r}-B{r})", BLACK, GBP)
    chart = BarChart()
    chart.type, chart.grouping, chart.overlap = "bar", "clustered", 100
    chart.title = "Change in package price per contact hour (£)"
    chart.y_axis.title = "£ per contact hour"
    data = Reference(ws, min_col=2, max_col=3, min_row=s, max_row=s + len(order))
    cats = Reference(ws, min_col=1, min_row=s + 1, max_row=s + len(order))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.x_axis.scaling.orientation = "maxMin"
    chart.height, chart.width = 8, 18
    ws.add_chart(chart, f"F{s}")
    widths(ws, {"A": 36, "B": 8, **{get_column_letter(i): 12 for i in range(3, 18)}})
    ws.row_dimensions[4].height = 40


def suppliers_sheet(wb, ledger):
    ws = wb.create_sheet("Suppliers")
    put(ws, "A1", "Supplier spend and savings", TITLE)
    put(ws, "A2", "Synthetic purchase ledger for one year (April 2026 to March 2027).", NOTE)
    header(ws, 4, ["Month", "Category", "Supplier", "Item", "Units", "Unit price", "Amount"])
    first = 5
    for i, row in enumerate(ledger.itertuples(index=False)):
        r = first + i
        put(ws, f"A{r}", row.month, BLUE); put(ws, f"B{r}", row.category, BLUE); put(ws, f"C{r}", row.supplier, BLUE)
        if isinstance(row.item, str):
            put(ws, f"D{r}", row.item, BLUE); put(ws, f"E{r}", int(row.units), BLUE, NUM)
            put(ws, f"F{r}", float(row.unit_price), BLUE, GBP); put(ws, f"G{r}", f"=E{r}*F{r}", BLACK, GBP)
        else:
            put(ws, f"G{r}", float(row.amount), BLUE, GBP)
    last = first + len(ledger) - 1
    R = lambda col: f"${col}${first}:${col}${last}"

    header(ws, 4, ["Spend by category", "Spend", "Share"], start=9)
    cats = ledger.groupby("category")["amount"].sum().sort_values(ascending=False).index.tolist()
    for k, cat in enumerate(cats):
        r = 5 + k
        put(ws, f"I{r}", cat)
        put(ws, f"J{r}", f"=SUMIFS({R('G')},{R('B')},I{r})", BLACK, GBP0)
        put(ws, f"K{r}", f"=J{r}/$J${5 + len(cats)}", BLACK, PCT)
    tot = 5 + len(cats)
    put(ws, f"I{tot}", "Total", bold=True); put(ws, f"J{tot}", f"=SUM(J5:J{tot - 1})", BLACK, GBP0, bold=True)

    p = tot + 2
    header(ws, p, ["Like-for-like products", "Spend", "Units", "Lowest unit price paid", "Saving at lowest price"], start=9)
    items = [i for i in ledger["item"].dropna().unique() if i != "Agency care worker hour"]
    for k, item in enumerate(items):
        r = p + 1 + k
        put(ws, f"I{r}", item)
        put(ws, f"J{r}", f"=SUMIFS({R('G')},{R('D')},I{r})", BLACK, GBP0)
        put(ws, f"K{r}", f"=SUMIFS({R('E')},{R('D')},I{r})", BLACK, NUM)
        put(ws, f"L{r}", f"=_xlfn.MINIFS({R('F')},{R('D')},I{r})", BLACK, GBP)
        put(ws, f"M{r}", f"=J{r}-K{r}*L{r}", BLACK, GBP0)
    pv_last = p + len(items)

    a = pv_last + 2
    header(ws, a, ["Agency hours", "Value"], start=9)
    put(ws, f"I{a + 1}", "Agency hours bought"); put(ws, f"J{a + 1}", f'=SUMIFS({R("E")},{R("D")},"Agency care worker hour")', BLACK, NUM)
    put(ws, f"I{a + 2}", "Agency spend"); put(ws, f"J{a + 2}", f'=SUMIFS({R("G")},{R("D")},"Agency care worker hour")', BLACK, GBP0)
    put(ws, f"I{a + 3}", "Average agency rate"); put(ws, f"J{a + 3}", f"=J{a + 2}/J{a + 1}", BLACK, GBP)
    put(ws, f"I{a + 4}", "Share of agency hours filled by employed staff"); put(ws, f"J{a + 4}", 0.30, BLUE, PCT, KEY)
    put(ws, f"I{a + 5}", "Cost of one paid care worker hour"); put(ws, f"J{a + 5}", "=Inputs!$C$34", GREEN, GBP)
    put(ws, f"I{a + 6}", "Saving"); put(ws, f"J{a + 6}", f"=J{a + 1}*J{a + 4}*(J{a + 3}-J{a + 5})", BLACK, GBP0)

    s = a + 8
    header(ws, s, ["Savings identified", "Value"], start=9)
    put(ws, f"I{s + 1}", "Buying like-for-like products at the lowest price paid"); put(ws, f"J{s + 1}", f"=SUM(M{p + 1}:M{pv_last})", BLACK, GBP0)
    put(ws, f"I{s + 2}", "Filling agency hours with employed staff"); put(ws, f"J{s + 2}", f"=J{a + 6}", BLACK, GBP0)
    put(ws, f"I{s + 3}", "Total", bold=True); put(ws, f"J{s + 3}", f"=J{s + 1}+J{s + 2}", BLACK, GBP0, bold=True)
    put(ws, f"I{s + 4}", "Share of total spend"); put(ws, f"J{s + 4}", f"=J{s + 3}/J{tot}", BLACK, PCT)
    widths(ws, {"A": 9, "B": 21, "C": 13, "D": 30, "E": 8, "F": 10, "G": 12, "H": 3,
                "I": 48, "J": 14, "K": 12, "L": 14, "M": 14})
    ws.freeze_panes = "A5"
    return {"tot": tot, "pv": (p + 1, pv_last), "agency_saving": a + 6, "savings_total": s + 3}


def bury_sheet(wb, bury):
    ws = wb.create_sheet("Bury rates")
    put(ws, "A1", "Published council rates: Bury Council, 2021/22 to 2026/27", TITLE)
    put(ws, "A2", "Standard rates from Bury Council's adult social care fee rates page; National Living Wage "
                  "from GOV.UK. Home care is per hour; care homes are per week.", NOTE)
    header(ws, 4, ["Year", "National Living Wage", "Home care standard", "Home care enhanced",
                   "Residential standard", "Residential dementia standard", "Nursing standard",
                   "Nursing dementia standard", "Home care rate / NLW", "Home care rate change",
                   "NLW change", "Home care rate index (2021/22 = 100)", "NLW index (2021/22 = 100)"])
    for i, b in enumerate(bury.itertuples(index=False)):
        r = 5 + i
        put(ws, f"A{r}", b.year, BLUE)
        for col, v in zip("BCDEFGH", [b.nlw, b.homecare_standard, b.homecare_enhanced, b.residential_standard,
                                      b.residential_dementia_standard, b.nursing_standard, b.nursing_dementia_standard]):
            put(ws, f"{col}{r}", float(v), BLUE, GBP)
        put(ws, f"I{r}", f"=C{r}/B{r}", BLACK, '0.00"x"')
        if i:
            put(ws, f"J{r}", f"=C{r}/C{r - 1}-1", BLACK, PCT)
            put(ws, f"K{r}", f"=B{r}/B{r - 1}-1", BLACK, PCT)
        put(ws, f"L{r}", f"=C{r}/C$5*100", BLACK, NUM1)
        put(ws, f"M{r}", f"=B{r}/B$5*100", BLACK, NUM1)
    last = 4 + len(bury)
    s = last + 2
    header(ws, s, ["2026/27: Bury's home care rates against this model (typical urban package)", "Value"])
    rows = [
        ("Bury standard rate (£ a contact hour)", f"=C{last}", GREEN, GBP),
        ("Model: direct cost (care worker pay, on-costs and mileage)", f"={RB}B{PKG26['direct']}", GREEN, GBP),
        ("Model: total cost", f"={RB}B{PKG26['cost']}", GREEN, GBP),
        ("Model: price at target margin", f"={RB}B{PKG26['price']}", GREEN, GBP),
        ("Standard rate left after direct cost", f"=B{s + 1}-B{s + 2}", BLACK, GBP),
        ("Model overheads per contact hour", f"=B{s + 3}-B{s + 2}", BLACK, GBP),
        ("Share of overheads the standard rate covers", f"=B{s + 5}/B{s + 6}", BLACK, PCT),
        ("Bury enhanced rate (£ a contact hour)", f"=D{last}", GREEN, GBP),
        ("Margin at the enhanced rate", f"=1-B{s + 3}/B{s + 8}", BLACK, PCT),
    ]
    for k, (label, f, font, fmt) in enumerate(rows):
        put(ws, f"A{s + 1 + k}", label); put(ws, f"B{s + 1 + k}", f, font, fmt)
    widths(ws, {"A": 58, **{get_column_letter(i): 13 for i in range(2, 14)}})
    ws.row_dimensions[4].height = 42
    return s


def about_sheet(wb):
    ws = wb.active
    ws.title = "About"
    lines = [
        ("Care contract pricing model", TITLE),
        ("A worked example of the analysis behind home care pricing, fee uplifts and bids. Built by Nihal Naresh Kumar.", BLACK),
        ("", BLACK),
        ("How to use it", BOLD),
        ("Change the blue cells on the Inputs sheet (for example next year's National Living Wage) and every sheet recalculates.", BLACK),
        ("Blue = input, black = formula, green = link to another sheet, yellow = assumption that matters most.", BLACK),
        ("", BLACK),
        ("Sheets", BOLD),
        ("Inputs: pay, statutory on-costs, travel, overheads, margin and the visit mix.", BLACK),
        ("Rate build-up: cost and price of 15, 30, 45 and 60-minute visits, urban and rural, for 2025/26 and 2026/27.", BLACK),
        ("Contracts: margin on twelve synthetic contracts and the uplift each needed for 2026/27.", BLACK),
        ("Care homes: occupancy, margin and break-even occupancy for five synthetic homes.", BLACK),
        ("Bid: bid or no-bid test for an illustrative tender.", BLACK),
        ("Sensitivity: which assumptions move the price most, with a tornado chart.", BLACK),
        ("Suppliers: spend by category and savings from a synthetic purchase ledger.", BLACK),
        ("Bury rates: six years of a council's published fee rates against the National Living Wage.", BLACK),
        ("", BLACK),
        ("Data", BOLD),
        ("Contracts, care homes, the tender and the supplier ledger are synthetic and do not describe any provider.", BLACK),
        ("Published figures: National Living Wage (GOV.UK); employer NI rates and thresholds 2026/27 (GOV.UK);", BLACK),
        ("auto-enrolment qualifying earnings 2026/27 (DWP review on GOV.UK); HMRC approved mileage rates (45p, then 55p from April 2026);", BLACK),
        ("Bury Council adult social care fee rates page and 11 March 2026 Cabinet report on provider fees", BLACK),
        ("(3.60% uplift on the standard home care rate, CPI 3.2% for November 2025); Homecare Association,", BLACK),
        ("22 December 2025 release (average homecare EBITDA margin 6.3% in October 2025). Links are in the README.", BLACK),
        ("", BLACK),
        ("Limits", BOLD),
        ("Travel time, overheads and the visit mix are assumptions; a provider's rota and ledger data would replace them.", BLACK),
        ("Sick pay, sleep-ins, waking nights, weekend premiums and bank holiday enhancements are not modelled.", BLACK),
    ]
    for i, (text, font) in enumerate(lines, start=1):
        put(ws, f"A{i}", text, font)
    widths(ws, {"A": 120})


def build(contracts=None, homes=None, ledger=None, tender=None, bury=None, path=OUT) -> dict:
    contracts = md.contracts() if contracts is None else contracts
    homes = md.care_homes() if homes is None else homes
    ledger = md.supplier_ledger() if ledger is None else ledger
    tender = md.TENDER if tender is None else tender
    bury = pd.read_csv(HERE / "bury_fee_rates.csv") if bury is None else bury
    wb = Workbook()
    about_sheet(wb)
    inputs_sheet(wb)
    rate_sheet(wb)
    total_row = contracts_sheet(wb, contracts)
    homes_sheet(wb, homes)
    bid_sheet(wb, tender)
    sensitivity_sheet(wb)
    sup = suppliers_sheet(wb, ledger)
    bury_start = bury_sheet(wb, bury)
    wb.save(path)
    return {"contracts_total_row": total_row, "suppliers": sup, "bury_block": bury_start}


if __name__ == "__main__":
    print(build())
