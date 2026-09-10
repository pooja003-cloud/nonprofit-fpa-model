"""
05_Scenario_Model.xlsx

Base, Upside and Downside cases driven off a single scenario switch, plus the
funding sensitivity work.

The scenario switch is a data-validated dropdown feeding CHOOSE(), so the whole
five-year model re-runs when the reviewer picks a different case. The three cases
are also laid out side by side on a summary sheet so they can be compared without
toggling.

Two deliberate departures from the brief's suggested shape:

  - The horizon is five years, not three. Three years is too short for this
    organization: reserves are large enough to absorb almost any funding shock
    for three years, so a three-year window would report 'no problem' for a
    structural deficit that becomes serious later.

  - The Base case assumes foundation and government funding DECLINE. The
    2021-2024 revenue surge was driven by episodic humanitarian response
    funding. A base case that extrapolates it would be forecasting the past.
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference, Series
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter as gcl

import inputs as I
import styles as S
import engine
from common_sheets import source_data_sheet, disclaimer

OUT = Path(__file__).resolve().parents[1] / "05_Scenario_Model.xlsx"
R = {}
YEARS = engine.SENSITIVITY_HORIZON            # 2025..2029
YCOL = {y: gcl(3 + i) for i, y in enumerate(YEARS)}   # C..G
A24 = "G"
SRC = "'Source data'!"
SW = "'Scenario switch'!"


# ==========================================================================
def sheet_switch(wb):
    ws = wb.create_sheet("Scenario switch")
    S.set_widths(ws, {"A": 44, "B": 16, "C": 16, "D": 16, "E": 16, "F": 74})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Scenario switch and driver table",
                      "Pick a case in the yellow cell. Every sheet in this workbook recalculates. "
                      "The three cases are also shown side by side on 'Scenario comparison'.",
                      width=6)

    S.label(ws, r, 1, "Active scenario", bold=True, size=12)
    sel = ws.cell(row=r, column=2, value="Base")
    sel.font = Font(name="Calibri", size=12, bold=True, color="7F6000")
    sel.fill = PatternFill("solid", fgColor="FFF2CC")
    sel.border = Border(*[Side(style="medium", color="BF8F00")] * 4)
    sel.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 24
    dv = DataValidation(type="list", formula1='"Downside,Base,Upside"', allow_blank=False,
                        showDropDown=False)
    dv.error = "Choose Downside, Base or Upside."
    dv.prompt = "Choose a scenario"
    ws.add_data_validation(dv)
    dv.add(sel)
    R["sw.sel"] = r
    r += 1

    S.label(ws, r, 1, "Scenario index", size=9, italic=True, color=S.MUTED, indent=1)
    idx = ws.cell(row=r, column=2, value=f'=MATCH(B{R["sw.sel"]},{{"Downside","Base","Upside"}},0)')
    idx.number_format = "0"
    idx.font = Font(name="Calibri", size=9, italic=True, color=S.MUTED)
    idx.alignment = Alignment(horizontal="center")
    R["sw.idx"] = r
    r += 2

    r = S.section(ws, r, "Scenario drivers", width=6)
    r = S.header_row(ws, r, ["Driver", "Downside", "Base", "Upside", "Active", "Rationale"])
    for k in I.SCENARIO_DRIVERS:
        down, base, up = I.SCENARIO_DRIVERS[k]
        S.label(ws, r, 1, I.SCENARIO_DRIVER_LABELS[k], size=9.5, indent=1)
        for i, v in enumerate((down, base, up)):
            S.put(ws, r, 2 + i, v, fmt=S.PCT1, is_input=True)
        act = ws.cell(row=r, column=5, value=f"=CHOOSE($B${R['sw.idx']},B{r},C{r},D{r})")
        act.number_format = S.PCT1
        act.font = Font(name="Calibri", size=10, bold=True, color=S.NAVY)
        act.alignment = Alignment(horizontal="right")
        act.fill = PatternFill("solid", fgColor="FFF9E6")
        R[f"sw.{k}"] = r
        r += 1
    r += 1

    r = S.section(ws, r, "What each case assumes", width=6)
    for name in ("Downside", "Base", "Upside"):
        S.label(ws, r, 1, name, bold=True, size=10, indent=1)
        c = ws.cell(row=r, column=2, value=I.SCENARIO_NARRATIVE[name])
        c.font = Font(name="Calibri", size=9, color=S.INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 13 * (len(I.SCENARIO_NARRATIVE[name]) // 110 + 2)
        r += 1
    r += 1

    r = S.section(ws, r, "Base year and cost behaviour", width=6)
    consts = [
        ("FY2024 liquid reserves (opening balance)",
         f"={SRC}{A24}{R['src.assets']}*{I.LIQUID_SHARE_OF_ASSETS[2024]}", S.MONEY, "sw.liq0"),
        ("FY2024 net assets", f"={SRC}{A24}{R['src.na']}", S.MONEY, "sw.na0"),
        ("FY2024 contributions", f"={SRC}{A24}{R['src.contrib']}", S.MONEY, "sw.contrib0"),
        ("FY2024 program service revenue", f"={SRC}{A24}{R['src.progrev']}", S.MONEY, "sw.progrev0"),
        ("FY2024 program services expense", f"={SRC}{A24}{R['src.prog']}", S.MONEY, "sw.prog0"),
        ("FY2024 management and general", f"={SRC}{A24}{R['src.mg']}", S.MONEY, "sw.mg0"),
        ("FY2024 fundraising", f"={SRC}{A24}{R['src.fr']}", S.MONEY, "sw.fr0"),
        ("Average gift inflation", I.AVG_GIFT_INFLATION, S.PCT1, "sw.giftinfl"),
        ("Management and general, variable share", I.MG_VARIABLE_SHARE, S.PCT1, "sw.mgvar"),
        ("Fundraising, variable share", I.FR_VARIABLE_SHARE, S.PCT1, "sw.frvar"),
        ("Reserve floor (months of expenses)", I.ALLOCATION_CONSTRAINTS["reserve_floor_months"],
         S.MONTHS, "sw.floor"),
        ("BBB net asset ceiling (multiple of expenses)", 3.00, S.MULT, "sw.naceil"),
        ("Largest single funder, share of revenue", I.LARGEST_FUNDER_SHARE, S.PCT1, "sw.largest"),
    ]
    for name, val, fmt, key in consts:
        S.label(ws, r, 1, name, size=9.5, indent=1)
        S.put(ws, r, 2, val, fmt=fmt, is_input=not isinstance(val, str))
        R[key] = r
        r += 1
    r += 1

    r = S.section(ws, r, "Funding shock (drives the sensitivity analysis)", width=6)
    S.label(ws, r, 1, "Permanent reduction in contributions from FY2025", size=10, indent=1)
    S.put(ws, r, 2, 0.0, fmt=S.PCT1, is_input=True)
    ws.cell(row=r, column=6, value="Set this to 0.10, 0.20 or 0.30 to see the effect on the "
                                   "scenario sheets. The sensitivity table runs the same shocks "
                                   "automatically, so this cell is for exploring, not reporting.")\
        .font = Font(name="Calibri", size=8.5, italic=True, color=S.MUTED)
    R["sw.shock"] = r
    r += 2
    r = disclaimer(ws, r, width=6)
    S.finish(ws, "BF8F00")
    return ws


# ==========================================================================
def sheet_scenario(wb):
    ws = wb.create_sheet("Scenario model")
    S.set_widths(ws, {"A": 44, "B": 15, **{YCOL[y]: 15 for y in YEARS}, "H": 4, "I": 58})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Five-year scenario model",
                      "Runs on whichever case is selected on 'Scenario switch'. Column B is the "
                      "FY2024 actual starting point.", width=9)

    hdr = ws.cell(row=r - 1, column=2, value="")
    r = S.header_row(ws, r, ["Line", "FY2024 A"] + [f"FY{y} F" for y in YEARS] + ["", "Note"])

    def row(name, base, fn, fmt=S.MONEY, bold=False, indent=1, key=None, note=""):
        nonlocal r
        S.label(ws, r, 1, name, bold=bold, size=10, indent=indent)
        if base is not None:
            c = ws.cell(row=r, column=2, value=base)
            c.number_format = fmt
            c.font = Font(name="Calibri", size=10, bold=bold, italic=True, color=S.MUTED)
            c.alignment = Alignment(horizontal="right")
        for i, y in enumerate(YEARS):
            col = YCOL[y]
            pcol = "B" if i == 0 else YCOL[YEARS[i - 1]]
            c = ws.cell(row=r, column=3 + i, value=fn(y, col, pcol, i))
            c.number_format = fmt
            c.font = Font(name="Calibri", size=10, bold=bold, color=S.INK)
            c.alignment = Alignment(horizontal="right")
        if note:
            n = ws.cell(row=r, column=9, value=note)
            n.font = Font(name="Calibri", size=8.5, color=S.MUTED)
            n.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 26
        if key:
            R[key] = r
        r += 1

    gk = {"Individual donations": "growth_individual", "Corporate donations": "growth_corporate",
          "Foundation grants": "growth_foundation", "Government funding": "growth_government",
          "Fundraising events": "growth_events"}

    # --- revenue ---------------------------------------------------------
    r = S.section(ws, r, "Revenue: donor count x average gift", width=9)
    for cat in I.REVENUE_CATEGORIES:
        drow = R[f"sw.{gk[cat]}"]
        row(f"{cat}: donors", I.DONOR_DRIVERS_2024[cat]["donors"],
            lambda y, c, p, i, dr=drow: (
                f"={p}{r}*(1+{SW}$E${dr})/(1+{SW}$B${R['sw.giftinfl']})"),
            fmt=S.NUM1, indent=2, key=f"s.don_{cat}")
    for cat in I.REVENUE_CATEGORIES:
        base_gift = (I.F990[2024]["contributions"] * I.REVENUE_MIX[cat]
                     / I.DONOR_DRIVERS_2024[cat]["donors"])
        row(f"{cat}: average gift", round(base_gift, 4),
            lambda y, c, p, i: f"={p}{r}*(1+{SW}$B${R['sw.giftinfl']})",
            fmt=S.MONEY2, indent=2, key=f"s.gift_{cat}")
    r += 1
    first_rev = r
    for cat in I.REVENUE_CATEGORIES:
        dr, gr = R[f"s.don_{cat}"], R[f"s.gift_{cat}"]
        row(cat, f"={SW}B{R['sw.contrib0']}*{I.REVENUE_MIX[cat]}",
            lambda y, c, p, i, dr=dr, gr=gr: (
                f"={c}{dr}*{c}{gr}*(1-{SW}$B${R['sw.shock']})"),
            key=f"s.rev_{cat}",
            note="Donor count times average gift, less the funding shock."
            if cat == I.REVENUE_CATEGORIES[0] else "")
    row("Total contributions", f"={SW}B{R['sw.contrib0']}",
        lambda y, c, p, i: f"=SUM({c}{first_rev}:{c}{first_rev + 4})",
        bold=True, indent=0, key="s.contrib")
    row("Program service revenue", f"={SW}B{R['sw.progrev0']}",
        lambda y, c, p, i: f"={p}{r}*(1+{SW}$E${R['sw.growth_program_revenue']})",
        key="s.progrev")
    row("Investment and other income", None, lambda y, c, p, i: 0, key="s.invinc",
        note="Return on the opening reserve balance, wired below.")
    row("Total revenue", f"={SW}B{R['sw.contrib0']}+{SW}B{R['sw.progrev0']}",
        lambda y, c, p, i: f"={c}{R['s.contrib']}+{c}{R['s.progrev']}+{c}{R['s.invinc']}",
        bold=True, indent=0, key="s.rev")
    r += 1

    # --- program volume ---------------------------------------------------
    r = S.section(ws, r, "Program volume and unit cost", width=9)
    for p_ in I.PROGRAMS:
        cap = I.PROGRAM_BASE[p_]["max_participants"]
        row(f"{p_}: participants", I.PROGRAM_BASE[p_]["participants"],
            lambda y, c, pc, i, cap=cap: (
                f"=MIN({pc}{r}*(1+{SW}$E${R['sw.growth_participants']}),{cap})"),
            fmt=S.NUM1, indent=2, key=f"s.parts_{p_}")
    for p_ in I.PROGRAMS:
        base_cpp = (I.PROGRAM_BASE[p_]["program_cost"]
                    * (I.F990[2024]["total_expense"] * I.FUNCTIONAL_SPLIT[2024][0]
                       / sum(I.PROGRAM_BASE[q]["program_cost"] for q in I.PROGRAMS))
                    / I.PROGRAM_BASE[p_]["participants"])
        row(f"{p_}: cost per participant", round(base_cpp, 4),
            lambda y, c, p, i: f"={p}{r}*(1+{SW}$E${R['sw.inflation_cpp']})",
            fmt=S.MONEY2, indent=2, key=f"s.cpp_{p_}")
    r += 1
    first_cost = r
    for p_ in I.PROGRAMS:
        row(f"{p_}: program cost", None,
            lambda y, c, p, i, p_=p_: f"={c}{R[f's.parts_{p_}']}*{c}{R[f's.cpp_{p_}']}",
            indent=2, key=f"s.cost_{p_}")
    r += 1

    # --- expenses ---------------------------------------------------------
    r = S.section(ws, r, "Expenses", width=9)
    row("Program services", f"={SW}B{R['sw.prog0']}",
        lambda y, c, p, i: f"=SUM({c}{first_cost}:{c}{first_cost + 3})",
        bold=True, indent=0, key="s.prog")
    row("Management and general", f"={SW}B{R['sw.mg0']}",
        lambda y, c, p, i: (
            f"={p}{r}*(1+{SW}$E${R['sw.inflation_personnel']})*"
            f"(1+{SW}$B${R['sw.mgvar']}*({c}{R['s.prog']}/{p}{R['s.prog']}-1))"),
        key="s.mg", note="Semi-fixed. In the Downside case this does not fall as fast as revenue, "
                         "which is why the program expense ratio deteriorates exactly when the "
                         "optics are worst.")
    row("Fundraising", f"={SW}B{R['sw.fr0']}",
        lambda y, c, p, i: (
            f"={p}{r}*(1+{SW}$E${R['sw.inflation_personnel']})*"
            f"(1+{SW}$B${R['sw.frvar']}*({c}{R['s.contrib']}/{p}{R['s.contrib']}-1))"),
        key="s.fr")
    row("Total expenses", f"={SW}B{R['sw.prog0']}+{SW}B{R['sw.mg0']}+{SW}B{R['sw.fr0']}",
        lambda y, c, p, i: f"={c}{R['s.prog']}+{c}{R['s.mg']}+{c}{R['s.fr']}",
        bold=True, indent=0, key="s.exp")
    r += 1

    # --- result and reserves ---------------------------------------------
    r = S.section(ws, r, "Result, reserves and covenants", width=9)
    row("Operating surplus / (deficit)", None,
        lambda y, c, p, i: f"={c}{R['s.rev']}-{c}{R['s.exp']}", bold=True, indent=0, key="s.result")
    row("Operating margin", None, lambda y, c, p, i: f"={c}{R['s.result']}/{c}{R['s.rev']}",
        fmt=S.PCT1, key="s.margin")
    row("Opening liquid reserves", None, lambda y, c, p, i: 0, key="s.liq_open")
    row("Closing liquid reserves", f"={SW}B{R['sw.liq0']}",
        lambda y, c, p, i: f"={c}{R['s.liq_open']}+{c}{R['s.result']}", bold=True, indent=0,
        key="s.liq_close")
    row("Closing net assets", f"={SW}B{R['sw.na0']}",
        lambda y, c, p, i: f"={p}{r}+{c}{R['s.result']}", key="s.na")
    row("Average monthly expenses", None, lambda y, c, p, i: f"={c}{R['s.exp']}/12", key="s.monthly")
    row("Liquid reserve runway", None,
        lambda y, c, p, i: f"={c}{R['s.liq_close']}/{c}{R['s.monthly']}", fmt=S.MONTHS,
        key="s.runway", note="Shaded red below the six-month floor.")
    row("Reserve floor headroom",
        None, lambda y, c, p, i: f"={c}{R['s.runway']}-{SW}$B${R['sw.floor']}", fmt=S.MONTHS,
        key="s.floorgap")
    row("Net assets, multiple of expenses", None,
        lambda y, c, p, i: f"={c}{R['s.na']}/{c}{R['s.exp']}", fmt=S.MULT, key="s.namult",
        note="Shaded amber above the BBB three-times accumulation ceiling.")
    r += 1

    # --- outputs ----------------------------------------------------------
    r = S.section(ws, r, "Program outputs", width=9)
    parts_terms = "+".join(f"{{c}}{R[f's.parts_{p_}']}" for p_ in I.PROGRAMS)
    row("Participants served", None, lambda y, c, p, i: "=" + parts_terms.replace("{c}", c),
        fmt=S.NUM1, key="s.parts")
    pe = engine.program_economics()["programs"]
    out_terms = "+".join(
        f"{{c}}{R[f's.parts_{p_}']}*{pe[p_]['outcome_rate']:.10f}" for p_ in I.PROGRAMS)
    row("Successful outcomes", None, lambda y, c, p, i: "=" + out_terms.replace("{c}", c),
        fmt=S.NUM1, key="s.out")
    row("Program expense ratio", None, lambda y, c, p, i: f"={c}{R['s.prog']}/{c}{R['s.exp']}",
        fmt=S.PCT1, key="s.progratio")
    row("Cost per participant", None, lambda y, c, p, i: f"={c}{R['s.prog']}/{c}{R['s.parts']}",
        fmt=S.MONEY2, key="s.cpp")
    row("Cost per successful outcome", None,
        lambda y, c, p, i: f"={c}{R['s.prog']}/{c}{R['s.out']}", fmt=S.MONEY2, key="s.cpo")

    # wire the reserve chain and investment income
    for i, y in enumerate(YEARS):
        col = YCOL[y]
        prev = "B" if i == 0 else YCOL[YEARS[i - 1]]
        ws[f"{col}{R['s.liq_open']}"] = (f"={SW}B{R['sw.liq0']}" if i == 0
                                         else f"={prev}{R['s.liq_close']}")
        ws[f"{col}{R['s.invinc']}"] = (
            f"={col}{R['s.liq_open']}*{SW}$E${R['sw.investment_return']}")

    lo, hi = YCOL[YEARS[0]], YCOL[YEARS[-1]]
    ws.conditional_formatting.add(
        f"{lo}{R['s.result']}:{hi}{R['s.result']}",
        CellIsRule(operator="lessThan", formula=["0"],
                   fill=PatternFill("solid", bgColor="FFC7CE")))
    ws.conditional_formatting.add(
        f"{lo}{R['s.runway']}:{hi}{R['s.runway']}",
        CellIsRule(operator="lessThan", formula=[f"{SW}$B${R['sw.floor']}"],
                   fill=PatternFill("solid", bgColor="FFC7CE")))
    ws.conditional_formatting.add(
        f"{lo}{R['s.namult']}:{hi}{R['s.namult']}",
        CellIsRule(operator="greaterThan", formula=[f"{SW}$B${R['sw.naceil']}"],
                   fill=PatternFill("solid", bgColor="FFEB9C")))
    S.freeze(ws, "C5")
    S.finish(ws, S.NAVY)
    return ws


# ==========================================================================
def sheet_comparison(wb):
    """All three cases side by side, computed in Python and written as values.

    This sheet is deliberately values-only: it has to show three scenarios at
    once, and a single CHOOSE-driven model can only be in one state at a time.
    The header says so rather than letting a reviewer assume otherwise.
    """
    ws = wb.create_sheet("Scenario comparison")
    S.set_widths(ws, {"A": 40, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13, "G": 3, "H": 56})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Scenario comparison",
                      "All three cases at once. Values computed by the model engine rather than "
                      "live formulas, because a single scenario switch can only hold one case at "
                      "a time. 'Scenario model' is the live version.", width=8)

    runs = {s: engine.forecast(s, years=YEARS) for s in I.SCENARIOS}
    metrics = [
        ("Total revenue", lambda x: x.total_revenue, S.MONEY),
        ("Total expenses", lambda x: x.total_expense, S.MONEY),
        ("Operating result", lambda x: x.operating_result, S.MONEY),
        ("Operating margin", lambda x: x.derived()["operating_margin"], S.PCT1),
        ("Closing liquid reserves", lambda x: x.closing_liquid, S.MONEY),
        ("Liquid reserve runway", lambda x: x.derived()["liquid_runway_months"], S.MONTHS),
        ("Closing net assets", lambda x: x.closing_net_assets, S.MONEY),
        ("Net assets / expenses", lambda x: x.derived()["net_asset_multiple"], S.MULT),
        ("Program expense ratio", lambda x: x.derived()["program_ratio"], S.PCT1),
        ("Participants served", lambda x: x.participants, S.NUM1),
        ("Successful outcomes", lambda x: x.outcomes, S.NUM1),
        ("Cost per successful outcome", lambda x: x.derived()["cost_per_outcome"], S.MONEY),
        ("Implied FTE", lambda x: x.fte, S.NUM1),
    ]

    for scen in I.SCENARIOS:
        r = S.section(ws, r, f"{scen} case", width=8)
        r = S.header_row(ws, r, ["Metric"] + [f"FY{y}" for y in YEARS] + ["", "Note"])
        for name, fn, fmt in metrics:
            S.label(ws, r, 1, name, size=9.5, indent=1)
            for i, y in enumerate(YEARS):
                S.put(ws, r, 2 + i, fn(runs[scen][i]), fmt=fmt)
            if name == "Operating result":
                R[f"cmp.{scen}.result"] = r
            if name == "Liquid reserve runway":
                R[f"cmp.{scen}.runway"] = r
            if name == "Successful outcomes":
                R[f"cmp.{scen}.out"] = r
            if name == "Total revenue":
                R[f"cmp.{scen}.rev"] = r
            r += 1
        ws.conditional_formatting.add(
            f"B{R[f'cmp.{scen}.result']}:F{R[f'cmp.{scen}.result']}",
            CellIsRule(operator="lessThan", formula=["0"],
                       fill=PatternFill("solid", bgColor="FFC7CE")))
        r += 1

    # spread table
    r = S.section(ws, r, "Spread between cases, FY2029", width=8)
    r = S.header_row(ws, r, ["Metric", "Downside", "Base", "Upside", "Up less Down", "", "", "Note"])
    for name, fn, fmt in metrics:
        S.label(ws, r, 1, name, size=9.5, indent=1)
        vals = {s: fn(runs[s][-1]) for s in I.SCENARIOS}
        for i, s in enumerate(I.SCENARIOS):
            S.put(ws, r, 2 + i, vals[s], fmt=fmt)
        S.put(ws, r, 5, vals["Upside"] - vals["Downside"], fmt=fmt, bold=True)
        r += 1
    r += 1

    r = S.section(ws, r, "Reading the spread", width=8)
    d, b, u = (runs[s][-1] for s in I.SCENARIOS)
    cross = engine.structural_deficit_crossover("Base")
    notes = [
        f"The FY2029 operating result ranges from a deficit of ${abs(d.operating_result):,.0f} in "
        f"the Downside case to a surplus of ${u.operating_result:,.0f} in the Upside case, a spread "
        f"of ${u.operating_result - d.operating_result:,.0f}. That spread is roughly "
        f"{(u.operating_result - d.operating_result) / I.F990[2024]['total_expense']:.0%} of the "
        f"current annual expense base, which is a fair measure of how much of this organization's "
        f"future is not within its own control.",

        f"The Base case turns negative in FY{cross['year']} - "
        f"{cross['years_from_base']} years from the last filed year - with a deficit of "
        f"${abs(cross['result']):,.0f}. This is the single most important output of the model. It "
        f"is not a crisis forecast; it is the arithmetic of a cost base that grows with programs "
        f"while foundation and government funding normalize after an episodic surge.",

        f"Reserves absorb all three cases comfortably over five years. Liquid runway ends the "
        f"horizon at {d.derived()['liquid_runway_months']:.0f} months in the Downside case against "
        f"{b.derived()['liquid_runway_months']:.0f} in Base. The balance sheet is not the "
        f"constraint here; the income statement is. That distinction matters because it changes "
        f"the remedy from cost-cutting to funding diversification.",

        f"Successful outcomes range from {d.outcomes:,.0f} to {u.outcomes:,.0f} in FY2029. The "
        f"Upside case delivers {u.outcomes / d.outcomes - 1:.0%} more placements than the "
        f"Downside case for {u.total_expense / d.total_expense - 1:.0%} more spending, which is a "
        f"reasonable indication that the marginal dollar in this organization still buys real "
        f"output rather than overhead.",
    ]
    for n in notes:
        c = ws.cell(row=r, column=1, value=n)
        c.font = Font(name="Calibri", size=9.5, color=S.INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 13 * (len(n) // 118 + 1)
        r += 1

    # charts
    r += 1
    ch = LineChart()
    ch.title = "Operating result by scenario"
    ch.height, ch.width = 8, 18
    ch.y_axis.numFmt = '#,##0,,"M"'
    for scen in I.SCENARIOS:
        ref = Reference(ws, min_col=2, max_col=6, min_row=R[f"cmp.{scen}.result"])
        ch.series.append(Series(ref, title=scen))
    ch.set_categories(Reference(ws, min_col=2, max_col=6, min_row=R["cmp.Base.result"] - 6))
    ws.add_chart(ch, f"A{r}")

    ch2 = LineChart()
    ch2.title = "Successful outcomes by scenario"
    ch2.height, ch2.width = 8, 18
    for scen in I.SCENARIOS:
        ref = Reference(ws, min_col=2, max_col=6, min_row=R[f"cmp.{scen}.out"])
        ch2.series.append(Series(ref, title=scen))
    ws.add_chart(ch2, f"A{r + 17}")

    S.finish(ws, S.ACCENT)
    return ws


# ==========================================================================
def sheet_sensitivity(wb):
    ws = wb.create_sheet("Funding sensitivity")
    S.set_widths(ws, {"A": 34, "B": 15, "C": 15, "D": 15, "E": 15, "F": 15, "G": 15, "H": 3,
                      "I": 54})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Funding sensitivity and minimum viable funding",
                      "What happens to results, liquidity, program capacity and beneficiaries if "
                      "contributions fall permanently. Run on the Base case with the reserve floor "
                      "enforced, so a funding cut forces a program cut rather than an eternal "
                      "deficit.", width=9)

    sens = engine.funding_sensitivity("Base")
    r = S.section(ws, r, "Permanent reduction in contributions from FY2025", width=9)
    r = S.header_row(ws, r, ["Metric"] + [f"{s['shock']:.0%} cut" for s in sens] + ["", "Note"])

    rows = [
        ("FY2029 operating result", lambda s: s["fy_last_result"], S.MONEY,
         "After any forced program contraction."),
        ("FY2029 structural result", lambda s: s["structural_result"], S.MONEY,
         "Before contraction - what the deficit would be if the organization refused to cut."),
        ("Cumulative five-year result", lambda s: s["cumulative_result"], S.MONEY, ""),
        ("Closing liquid reserves", lambda s: s["closing_liquid"], S.MONEY, ""),
        ("Liquid reserve runway", lambda s: s["liquid_runway_months"], S.MONTHS, ""),
        ("Net assets, months", lambda s: s["net_asset_months"], S.MONTHS, ""),
        ("Participants served, FY2029", lambda s: s["participants"], S.NUM1, ""),
        ("Successful outcomes, FY2029", lambda s: s["outcomes"], S.NUM1, ""),
        ("Program capacity retained", lambda s: s["capacity_retained"], S.PCT1,
         "Against the no-shock Base case. Below 100% means programs were cut to hold the floor."),
        ("Beneficiaries forgone", lambda s: s["participants_lost"], S.NUM1, ""),
        ("Outcomes forgone", lambda s: s["outcomes_lost"], S.NUM1, ""),
        ("Years until the reserve floor", lambda s: s["years_to_floor"]["years"] or 0, "0",
         "Tested over sixteen years. This is the vulnerability measure a board can act on."),
    ]
    for name, fn, fmt, note in rows:
        S.label(ws, r, 1, name, size=9.5, indent=1)
        for i, s in enumerate(sens):
            S.put(ws, r, 2 + i, fn(s), fmt=fmt)
        if note:
            n = ws.cell(row=r, column=9, value=note)
            n.font = Font(name="Calibri", size=8.5, color=S.MUTED)
            n.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 26
        if name == "Successful outcomes, FY2029":
            R["sens.out"] = r
        if name == "Liquid reserve runway":
            R["sens.runway"] = r
        r += 1
    ws.conditional_formatting.add(
        f"B{R['sens.runway']}:G{R['sens.runway']}",
        ColorScaleRule(start_type="num", start_value=0, start_color="F8696B",
                       mid_type="num", mid_value=12, mid_color="FFEB84",
                       end_type="num", end_value=30, end_color="63BE7B"))
    r += 2

    # --- largest funder ---------------------------------------------------
    lf = engine.largest_funder_loss("Base")
    r = S.section(ws, r, "What if the largest single funder is lost", width=9)
    items = [
        ("Share of total revenue at risk", lf["revenue_share_lost"], S.PCT1),
        ("Equivalent share of contributions", lf["contribution_share_lost"], S.PCT1),
        ("Contributions lost in FY2025", lf["dollars_lost_fy1"], S.MONEY),
        ("FY2029 operating result", lf["final_operating_result"], S.MONEY),
        ("FY2029 result without the loss", lf["reference_operating_result"], S.MONEY),
        ("Cumulative five-year shortfall", lf["cumulative_gap"], S.MONEY),
        ("Closing liquid reserves", lf["closing_liquid"], S.MONEY),
        ("Liquid reserve runway at FY2029", lf["liquid_runway_months"], S.MONTHS),
        ("Years until the reserve floor", lf["years_to_floor"]["years"] or 0, "0"),
        ("Years until the floor without the loss", lf["reference_years_to_floor"]["years"] or 0, "0"),
        ("Program capacity retained", lf["capacity_retained"], S.PCT1),
    ]
    for name, val, fmt in items:
        S.label(ws, r, 1, name, size=9.5, indent=1)
        S.put(ws, r, 2, val, fmt=fmt)
        r += 1
    r += 1

    # --- thresholds -------------------------------------------------------
    mvf = engine.minimum_viable_funding("Base")
    r = S.section(ws, r, "Two different funding floors", width=9)
    S.note(ws, r, "These answer different questions and quoting either one alone is misleading. "
                  "The reserve threshold is how far funding can fall before the balance sheet "
                  "forces a program cut. The structural threshold is the funding level at which "
                  "the organization stops breaking even at all. The gap between them is the time "
                  "the balance sheet buys management to fix an income statement problem.", width=9)
    r += 1
    items = [
        ("Reserve threshold: maximum sustainable funding cut", mvf["max_sustainable_shock"], S.PCT1),
        ("  contributions at that level, FY2025", mvf["min_contributions_fy1"], S.MONEY),
        ("  headroom below the Base case", mvf["dollars_of_headroom"], S.MONEY),
        ("Structural threshold: funding uplift needed to break even by FY2029",
         mvf["required_funding_uplift"], S.PCT1),
        ("Base case turns to deficit in", mvf["crossover"]["year"], "0"),
        ("  deficit in that year", abs(mvf["crossover"]["result"]), S.MONEY),
        ("Years of reserve at Base case with no shock",
         mvf["years_to_floor_no_shock"]["years"] or 0, "0"),
    ]
    for name, val, fmt in items:
        S.label(ws, r, 1, name, size=10, indent=1 if name.startswith("  ") else 0,
                bold=not name.startswith("  "))
        S.put(ws, r, 2, val, fmt=fmt, bold=not name.startswith("  "))
        r += 1
    r += 1

    # --- expense lever ----------------------------------------------------
    er = engine.expense_reduction_required("Downside")
    r = S.section(ws, r, "The expense lever, Downside case", width=9)
    if er["required"]:
        items = [("Expense reduction needed to hold the floor", er["reduction_pct"], S.PCT1),
                 ("  in dollars, FY2025", er["reduction_dollars"], S.MONEY),
                 ("  participants forgone", er["participants_forgone"], S.NUM1),
                 ("  outcomes forgone", er["outcomes_forgone"], S.NUM1)]
    else:
        items = [("Expense reduction needed to hold the floor", 0.0, S.PCT1),
                 ("Lowest runway reached", er["worst_runway"], S.MONTHS),
                 ("In fiscal year", er["worst_year"], "0"),
                 ("FY2029 structural result", er["final_structural_result"], S.MONEY)]
    for name, val, fmt in items:
        S.label(ws, r, 1, name, size=10, indent=1 if name.startswith("  ") else 0)
        S.put(ws, r, 2, val, fmt=fmt)
        r += 1
    r += 1
    S.note(ws, r, f"Even the Downside case does not force a cut within five years: the lowest "
                  f"runway reached is {er['worst_runway']:.1f} months against a six-month floor. "
                  f"That is genuine resilience and it should be said plainly. It is also the "
                  f"reason a five-year horizon was used - on a three-year view every scenario "
                  f"looks safe, and the FY2029 structural deficit of "
                  f"${abs(er['final_structural_result']):,.0f} would never appear.", width=9)

    S.finish(ws, S.WARN)
    return ws


def build():
    wb = Workbook()
    wb.remove(wb.active)
    source_data_sheet(wb, R)
    sheet_switch(wb)
    sheet_scenario(wb)
    sheet_comparison(wb)
    sheet_sensitivity(wb)
    wb.move_sheet("Scenario switch", offset=-2)
    wb.save(OUT)
    print(f"wrote {OUT}")
    return OUT


if __name__ == "__main__":
    build()
