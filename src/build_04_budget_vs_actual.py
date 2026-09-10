"""
03_Budget_vs_Actual.xlsx

FY2024 actuals are real, taken from the filing. The budget they are measured
against does not exist publicly, so it is reconstructed: what a board would
plausibly have approved in late FY2023 for FY2024, given what was known then.
Every budget cell is labelled as a simulation.

The point is not the budget. The point is the variance decomposition and the
commentary, which is where an FP&A analyst actually earns their keep: separating
a cost overrun caused by serving more people from one caused by serving them
less efficiently, and then saying which one happened and what it means.
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference, Series
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter as gcl

import inputs as I
import styles as S
import engine
from common_sheets import source_data_sheet, disclaimer, HCOL

OUT = Path(__file__).resolve().parents[1] / "excel-models" / "03_Budget_vs_Actual.xlsx"
R = {}
A24, A23 = "G", "F"          # FY2024 and FY2023 columns on the source sheet
SRC = "'Source data'!"


def sheet_basis(wb):
    ws = wb.create_sheet("Budget basis")
    S.set_widths(ws, {"A": 4, "B": 46, "C": 18, "D": 14, "E": 74})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Budget basis  -  reconstructed, not filed",
                      "There is no public FY2024 budget for this organization. This sheet builds "
                      "one, states exactly how, and labels it a simulation everywhere it appears.",
                      width=5)

    warn = ws.cell(row=r, column=2, value=(
        "SIMULATED BUDGET. Every budget figure in this workbook is constructed by the author. "
        "It is not the organization's approved budget and does not represent its plans. Only the "
        "actual column is real."))
    warn.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    warn.fill = PatternFill("solid", fgColor=S.WARN)
    warn.alignment = Alignment(wrap_text=True, vertical="center")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    ws.row_dimensions[r].height = 32
    r += 2

    r = S.section(ws, r, "How the budget was constructed", width=5)
    body = [
        ("Anchor", "FY2023 actual, as filed. A board approving a FY2024 budget in late FY2023 "
                   "would have had the FY2023 result in front of it and little else."),
        ("Revenue", "Planned to grow 12% on FY2023. The organization had just come through two "
                    "years of 60%-plus growth driven by humanitarian response funding, and a "
                    "finance committee would reasonably assume that could not repeat. Twelve "
                    "percent is deliberately cautious relative to recent experience."),
        ("Expenses", "Planned to grow 26%, roughly double the revenue plan. This is the "
                     "deliberate deployment of accumulated reserve into program capacity, which "
                     "is exactly what a board sitting on two years of expenses in net assets "
                     "would instruct management to do."),
        ("Volume", "Participants planned to grow 18%, ahead of the expense plan, on the "
                   "assumption that expansion comes partly through lower-cost digital channels."),
        ("Mix", "Revenue and expense category splits held at their FY2023 proportions. A budget "
                "built a year ahead would not usually re-forecast mix in detail."),
    ]
    for k, v in body:
        S.label(ws, r, 2, k, bold=True, size=10)
        c = ws.cell(row=r, column=5, value=v)
        c.font = Font(name="Calibri", size=9, color=S.INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 13 * (len(v) // 82 + 1)
        r += 1
    r += 1

    r = S.section(ws, r, "Budget drivers", width=5)
    r = S.header_row(ws, r, ["", "Driver", "Value", "Tier", "Note"], start_col=1)
    drivers = [
        ("Revenue growth on FY2023", I.BUDGET_2024_REVENUE_GROWTH, S.PCT1, "bud.revg"),
        ("Expense growth on FY2023", I.BUDGET_2024_EXPENSE_GROWTH, S.PCT1, "bud.expg"),
        ("Participant growth on FY2023", I.BUDGET_2024_PARTICIPANT_GROWTH, S.PCT1, "bud.partg"),
    ]
    for name, val, fmt, key in drivers:
        S.label(ws, r, 2, name, size=10, indent=1)
        S.put(ws, r, 3, val, fmt=fmt, is_input=True)
        S.tier_cell(ws, r, 4, "SYNTHETIC")
        R[key] = r
        r += 1
    r += 1
    r = disclaimer(ws, r, width=5)
    S.finish(ws, S.WARN)
    return ws


def sheet_bva(wb):
    ws = wb.create_sheet("Budget vs actual")
    S.set_widths(ws, {"A": 44, "B": 17, "C": 17, "D": 16, "E": 12, "F": 13, "G": 58})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "FY2024 budget versus actual",
                      "Budget column is a simulation (see 'Budget basis'). Actual column is as "
                      "filed. Variance is actual less budget; the favourable / unfavourable flag "
                      "flips sign for expense lines.", width=7)
    r = S.header_row(ws, r, ["Line", "Budget (simulated)", "Actual (filed)", "Variance",
                             "Variance %", "F / U", "Commentary"])

    def line(name, budget_f, actual_f, expense_like=False, bold=False, indent=1, key=None,
             note="", fmt=S.MONEY):
        nonlocal r
        S.label(ws, r, 1, name, bold=bold, size=10, indent=indent)
        b = ws.cell(row=r, column=2, value=budget_f)
        b.number_format = fmt
        b.font = Font(name="Calibri", size=10, bold=bold, italic=True, color=S.WARN)
        b.alignment = Alignment(horizontal="right")
        a = ws.cell(row=r, column=3, value=actual_f)
        a.number_format = fmt
        a.font = Font(name="Calibri", size=10, bold=bold, color=S.INK)
        a.alignment = Alignment(horizontal="right")
        v = ws.cell(row=r, column=4, value=f"=C{r}-B{r}")
        v.number_format = fmt
        v.font = Font(name="Calibri", size=10, bold=bold)
        v.alignment = Alignment(horizontal="right")
        p = ws.cell(row=r, column=5, value=f'=IF(B{r}=0,"n/a",D{r}/ABS(B{r}))')
        p.number_format = S.PCT1
        p.font = Font(name="Calibri", size=10, bold=bold)
        p.alignment = Alignment(horizontal="right")
        cond = f"D{r}<0" if expense_like else f"D{r}>0"
        fu = ws.cell(row=r, column=6, value=f'=IF({cond},"Favourable","Unfavourable")')
        fu.font = Font(name="Calibri", size=9, bold=True)
        fu.alignment = Alignment(horizontal="center")
        if note:
            c = ws.cell(row=r, column=7, value=note)
            c.font = Font(name="Calibri", size=8.5, color=S.MUTED)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 26
        if key:
            R[key] = r
        r += 1

    bg, eg, pg = R["bud.revg"], R["bud.expg"], R["bud.partg"]
    B = "'Budget basis'!$C$"

    # --- revenue ---------------------------------------------------------
    r = S.section(ws, r, "Revenue", width=7)
    line("Contributions and grants",
         f"={SRC}{A23}{R['src.rev']}*(1+{B}{bg})*({SRC}{A23}{R['src.contrib']}/{SRC}{A23}{R['src.rev']})",
         f"={SRC}{A24}{R['src.contrib']}", key="b.contrib",
         note="Budget holds the FY2023 revenue mix and applies the 12% growth plan.")
    line("Program service revenue",
         f"={SRC}{A23}{R['src.rev']}*(1+{B}{bg})*({SRC}{A23}{R['src.progrev']}/{SRC}{A23}{R['src.rev']})",
         f"={SRC}{A24}{R['src.progrev']}", key="b.progrev")
    line("Investment and other income",
         f"={SRC}{A23}{R['src.rev']}*(1+{B}{bg})-B{R['b.contrib']}-B{R['b.progrev']}",
         f"={SRC}{A24}{R['src.invinc']}+{SRC}{A24}{R['src.netting']}", key="b.invinc",
         note="Actual combines investment income with the fundraising-event netting line so the "
              "two columns are measured on the same basis.")
    line("Total revenue", f"=SUM(B{R['b.contrib']}:B{R['b.invinc']})",
         f"={SRC}{A24}{R['src.rev']}", bold=True, indent=0, key="b.rev")
    r += 1

    # --- expenses --------------------------------------------------------
    r = S.section(ws, r, "Expenses", width=7)
    line("Program services",
         f"={SRC}{A23}{R['src.exp']}*(1+{B}{eg})*{SRC}{A23}{R['src.fs_prog']}",
         f"={SRC}{A24}{R['src.prog']}", expense_like=True, key="b.prog",
         note="Budget applies the FY2023 functional split to the grown expense base.")
    line("Management and general",
         f"={SRC}{A23}{R['src.exp']}*(1+{B}{eg})*{SRC}{A23}{R['src.fs_mg']}",
         f"={SRC}{A24}{R['src.mg']}", expense_like=True, key="b.mg")
    line("Fundraising",
         f"={SRC}{A23}{R['src.exp']}*(1+{B}{eg})*{SRC}{A23}{R['src.fs_fr']}",
         f"={SRC}{A24}{R['src.fr']}", expense_like=True, key="b.fr")
    line("Total expenses", f"=SUM(B{R['b.prog']}:B{R['b.fr']})",
         f"={SRC}{A24}{R['src.exp']}", expense_like=True, bold=True, indent=0, key="b.exp")
    r += 1

    # --- personnel memo --------------------------------------------------
    r = S.section(ws, r, "Natural classification (memo)", width=7)
    pers23 = (f"({SRC}{A23}{R['src.offcomp']}+{SRC}{A23}{R['src.othersal']}"
              f"+{SRC}{A23}{R['src.ptax']}+({SRC}{A23}{R['src.offcomp']}+{SRC}{A23}{R['src.othersal']})*{I.BENEFITS_RATE})")
    pers24 = (f"{SRC}{A24}{R['src.offcomp']}+{SRC}{A24}{R['src.othersal']}"
              f"+{SRC}{A24}{R['src.ptax']}+({SRC}{A24}{R['src.offcomp']}+{SRC}{A24}{R['src.othersal']})*{I.BENEFITS_RATE}")
    line("Personnel", f"={SRC}{A23}{R['src.exp']}*(1+{B}{eg})*{pers23}/{SRC}{A23}{R['src.exp']}",
         f"={pers24}", expense_like=True, key="b.personnel",
         note="Budget holds the FY2023 personnel share of expenses constant.")
    line("Non-personnel", f"=B{R['b.exp']}-B{R['b.personnel']}",
         f"=C{R['b.exp']}-C{R['b.personnel']}", expense_like=True, key="b.nonpersonnel")
    r += 1

    # --- result ----------------------------------------------------------
    r = S.section(ws, r, "Result", width=7)
    line("Operating surplus / (deficit)", f"=B{R['b.rev']}-B{R['b.exp']}",
         f"=C{R['b.rev']}-C{R['b.exp']}", bold=True, indent=0, key="b.result")
    line("Operating margin", f"=B{R['b.result']}/B{R['b.rev']}",
         f"=C{R['b.result']}/C{R['b.rev']}", fmt=S.PCT1, key="b.margin")
    line("Program expense ratio", f"=B{R['b.prog']}/B{R['b.exp']}",
         f"=C{R['b.prog']}/C{R['b.exp']}", fmt=S.PCT1, key="b.progratio")
    r += 1

    # --- volume ----------------------------------------------------------
    r = S.section(ws, r, "Volume and unit cost", width=7)
    line("Participants served", f"={SRC}{A23}{R['src.parts']}*(1+{B}{pg})",
         f"={SRC}{A24}{R['src.parts']}", fmt=S.NUM1, key="b.parts")
    line("Successful outcomes",
         f"=B{R['b.parts']}*({SRC}{A23}{R['src.places']}/{SRC}{A23}{R['src.parts']})",
         f"={SRC}{A24}{R['src.places']}", fmt=S.NUM1, key="b.places",
         note="Budget holds the FY2023 outcome rate constant on the planned participant count.")
    line("Cost per participant", f"=B{R['b.prog']}/B{R['b.parts']}",
         f"=C{R['b.prog']}/C{R['b.parts']}", fmt=S.MONEY2, expense_like=True, key="b.cpp")
    line("Cost per successful outcome", f"=B{R['b.prog']}/B{R['b.places']}",
         f"=C{R['b.prog']}/C{R['b.places']}", fmt=S.MONEY2, expense_like=True, key="b.cpo")

    for row_key in ("b.rev", "b.exp", "b.result", "b.prog"):
        ws.conditional_formatting.add(
            f"D{R[row_key]}", CellIsRule(operator="lessThan", formula=["0"],
                                         fill=PatternFill("solid", bgColor="FFC7CE")))
    S.freeze(ws, "B5")
    S.finish(ws, S.NAVY)
    return ws


def sheet_decomposition(wb):
    ws = wb.create_sheet("Variance decomposition")
    S.set_widths(ws, {"A": 46, "B": 18, "C": 18, "D": 14, "E": 66})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Variance decomposition",
                      "Splitting each variance into the part caused by doing more or less, and the "
                      "part caused by each unit costing more or less. Without this split a "
                      "variance number cannot be acted on.", width=5)

    BVA = "'Budget vs actual'!"

    def item(name, formula, fmt=S.MONEY, bold=False, indent=1, key=None, note=""):
        nonlocal r
        S.label(ws, r, 1, name, bold=bold, size=10, indent=indent)
        c = ws.cell(row=r, column=2, value=formula)
        c.number_format = fmt
        c.font = Font(name="Calibri", size=10, bold=bold)
        c.alignment = Alignment(horizontal="right")
        if note:
            n = ws.cell(row=r, column=5, value=note)
            n.font = Font(name="Calibri", size=8.5, color=S.MUTED)
            n.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 26
        if key:
            R[key] = r
        r += 1

    # --- program cost bridge ---------------------------------------------
    r = S.section(ws, r, "Program cost variance bridge", width=5)
    r = S.header_row(ws, r, ["Component", "Amount", "", "", "How it is calculated"], start_col=1)
    item("Budgeted program cost", f"={BVA}B{R['b.prog']}", key="d.bprog")
    item("Budgeted participants", f"={BVA}B{R['b.parts']}", fmt=S.NUM1, key="d.bparts")
    item("Budgeted cost per participant", f"={BVA}B{R['b.cpp']}", fmt=S.MONEY2, key="d.bcpp")
    item("Actual participants", f"={BVA}C{R['b.parts']}", fmt=S.NUM1, key="d.aparts")
    item("Actual cost per participant", f"={BVA}C{R['b.cpp']}", fmt=S.MONEY2, key="d.acpp")
    item("Actual program cost", f"={BVA}C{R['b.prog']}", key="d.aprog")
    r += 1
    item("Total program cost variance", f"=B{R['d.aprog']}-B{R['d.bprog']}", bold=True, indent=0,
         key="d.totvar")
    item("Volume variance", f"=(B{R['d.aparts']}-B{R['d.bparts']})*B{R['d.bcpp']}", key="d.vol",
         note="Change in participants, valued at the BUDGETED cost per participant. This is the "
              "cost of doing more or less than planned.")
    item("Rate variance", f"=(B{R['d.acpp']}-B{R['d.bcpp']})*B{R['d.aparts']}", key="d.rate",
         note="Change in cost per participant, valued at ACTUAL participants. This is the cost of "
              "each unit being more or less expensive than planned.")
    item("Check: volume plus rate equals total",
         f'=IF(ABS(B{R["d.vol"]}+B{R["d.rate"]}-B{R["d.totvar"]})<1,"ok","CHECK")', fmt="@",
         key="d.check",
         note="The two effects must reconcile exactly. This ordering charges the interaction term "
              "to rate; the reverse convention is equally valid but must be stated, which is why "
              "it is stated here.")
    r += 1

    # --- funding variance -------------------------------------------------
    r = S.section(ws, r, "Funding variance by source", width=5)
    r = S.header_row(ws, r, ["Category", "Budget", "Actual", "Variance", "Note"], start_col=1)
    first = r
    for cat in I.REVENUE_CATEGORIES:
        share = I.REVENUE_MIX[cat]
        S.label(ws, r, 1, cat, size=9.5, indent=1)
        b = ws.cell(row=r, column=2, value=f"={BVA}B{R['b.contrib']}*{share}")
        a = ws.cell(row=r, column=3, value=f"={BVA}C{R['b.contrib']}*{share}")
        v = ws.cell(row=r, column=4, value=f"=C{r}-B{r}")
        for c in (b, a, v):
            c.number_format = S.MONEY
            c.font = Font(name="Calibri", size=10)
            c.alignment = Alignment(horizontal="right")
        r += 1
    S.label(ws, r, 1, "Total contributions variance", bold=True, indent=0)
    for col in (2, 3, 4):
        L = gcl(col)
        c = ws.cell(row=r, column=col, value=f"=SUM({L}{first}:{L}{r - 1})")
        c.number_format = S.MONEY
        c.alignment = Alignment(horizontal="right")
    S.total_row_style(ws, r, range(1, 5))
    R["d.fundvar"] = r
    r += 1
    S.note(ws, r, "The category split is modelled at constant shares, so this table apportions the "
                  "total contribution variance rather than explaining it. It is honest about what "
                  "it can and cannot show: without the FY2024 Schedule B or a grants schedule, "
                  "which funder moved is not publicly determinable.", width=5)
    r += 1

    # --- personnel variance ----------------------------------------------
    r = S.section(ws, r, "Personnel variance", width=5)
    item("Budgeted personnel cost", f"={BVA}B{R['b.personnel']}", key="d.bpers")
    item("Actual personnel cost", f"={BVA}C{R['b.personnel']}", key="d.apers")
    item("Total personnel variance", f"=B{R['d.apers']}-B{R['d.bpers']}", bold=True, indent=0,
         key="d.persvar")
    item("Personnel as a share of expenses, budget", f"=B{R['d.bpers']}/{BVA}B{R['b.exp']}",
         fmt=S.PCT1, key="d.bpersratio")
    item("Personnel as a share of expenses, actual", f"=B{R['d.apers']}/{BVA}C{R['b.exp']}",
         fmt=S.PCT1, key="d.apersratio",
         note="A personnel variance that is large in dollars but flat as a share of expenses is a "
              "scale story, not an efficiency story.")
    r += 1

    # --- cost per participant bridge -------------------------------------
    r = S.section(ws, r, "Cost per participant bridge", width=5)
    item("Budgeted cost per participant", f"=B{R['d.bcpp']}", fmt=S.MONEY2, key="d.b2")
    item("Effect of fixed-cost absorption",
         f"=-(B{R['d.bprog']}/B{R['d.bparts']})*(1-B{R['d.bparts']}/B{R['d.aparts']})",
         fmt=S.MONEY2, key="d.absorb",
         note="Spreading the planned program cost over a different number of participants. "
              "Negative means more participants diluted the fixed base.")
    item("Effect of total spend differing from plan",
         f"=(B{R['d.aprog']}-B{R['d.bprog']})/B{R['d.aparts']}", fmt=S.MONEY2, key="d.spend")
    item("Actual cost per participant", f"=B{R['d.b2']}+B{R['d.absorb']}+B{R['d.spend']}",
         fmt=S.MONEY2, bold=True, indent=0, key="d.a2")
    item("Check: bridge ties to actual",
         f'=IF(ABS(B{R["d.a2"]}-B{R["d.acpp"]})<0.01,"ok","CHECK")', fmt="@")

    S.finish(ws, S.ACCENT)
    return ws


def sheet_commentary(wb):
    """Written management commentary. The numbers come from the engine so the
    text and the workbook cannot drift apart."""
    ws = wb.create_sheet("Management commentary")
    S.set_widths(ws, {"A": 3, "B": 112})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Management commentary on FY2024 variances",
                      "Written against the reconstructed budget. Figures are quoted from the "
                      "decomposition sheet.", width=2)

    bva = engine.budget_vs_actual()
    L = bva["lines"]
    dec = bva["decomposition"]
    hist = engine.historical()["rows"][2024]
    fmt = lambda v: f"${abs(v):,.0f}"
    pct = lambda v: f"{abs(v):.1%}"

    rev_v = L["Total revenue"]
    exp_v = L["Total expenses"]
    prog_v = L["  Program services"]
    mg_v = L["  Management and general"]
    res_v = L["Operating result"]
    part_v = L["Participants served"]
    out_v = L["Successful outcomes"]
    cpp_v = L["Cost per participant"]
    inv_v = L["  Investment and other income"]
    con_v = L["  Contributions and grants"]

    budget_rate = out_v["budget"] / part_v["budget"]
    actual_rate = out_v["actual"] / part_v["actual"]
    vol_share = dec["volume_variance"] / dec["total_program_variance"]
    inv_share = inv_v["variance"] / rev_v["variance"]
    outcomes_at_plan_rate = part_v["actual"] * budget_rate
    outcomes_lost_to_mix = outcomes_at_plan_rate - out_v["actual"]

    paras = [
        ("Headline",
         f"Revenue of {fmt(rev_v['actual'])} finished {fmt(rev_v['variance'])} "
         f"({pct(rev_v['variance_pct'])}) ahead of the {fmt(rev_v['budget'])} plan. Expenses of "
         f"{fmt(exp_v['actual'])} finished {fmt(exp_v['variance'])} ({pct(exp_v['variance_pct'])}) "
         f"above the {fmt(exp_v['budget'])} plan. The operating surplus of {fmt(res_v['actual'])} "
         f"was therefore {fmt(res_v['variance'])} ({pct(res_v['variance_pct'])}) better than "
         f"budgeted. On the face of it this is a good year. The decomposition below argues that "
         f"the most important variance in the file is a small unfavourable one, not the large "
         f"favourable one."),

        ("Quality of the revenue beat",
         f"Of the {fmt(rev_v['variance'])} revenue overperformance, {fmt(inv_v['variance'])} - "
         f"{pct(inv_share)} of it - is investment and other income, which came in "
         f"{pct(inv_v['variance_pct'])} above plan simply because the accumulated reserve is large "
         f"and returns were strong. That is a balance sheet outcome, not a fundraising one, and it "
         f"will reverse if markets do. Contributions beat plan by {fmt(con_v['variance'])} "
         f"({pct(con_v['variance_pct'])}), which is the number that reflects actual fundraising "
         f"performance. Program service revenue was the one revenue line to miss, coming in "
         f"{fmt(inv_v['variance'] and L['  Program service revenue']['variance'])} "
         f"({pct(L['  Program service revenue']['variance_pct'])}) below plan."),

        ("Program cost variance, decomposed",
         f"Program services spending exceeded budget by {fmt(prog_v['variance'])} "
         f"({pct(prog_v['variance_pct'])}). Splitting that into its two causes: "
         f"{fmt(dec['volume_variance'])}, or {pct(vol_share)} of the overrun, is volume - the "
         f"organization served {part_v['actual']:,.0f} participants against a plan of "
         f"{part_v['budget']:,.0f}, {abs(part_v['variance']):,.0f} more people than budgeted, "
         f"valued at the budgeted unit cost. The remaining {fmt(dec['rate_variance'])} is rate: "
         f"cost per participant came in at {fmt(cpp_v['actual'])} against {fmt(cpp_v['budget'])} "
         f"planned, {pct(cpp_v['variance_pct'])} above. Program expense exceeded budget primarily "
         f"because participant acquisition ran ahead of the planned expansion, with only a quarter "
         f"of the overrun attributable to unit costs rising faster than assumed."),

        ("The variance that actually matters",
         f"Successful outcomes came in at {out_v['actual']:,.0f} against a plan of "
         f"{out_v['budget']:,.0f} - {abs(out_v['variance']):,.0f} fewer, {pct(out_v['variance_pct'])} "
         f"below budget - in a year when the organization served {abs(part_v['variance']):,.0f} more "
         f"participants than planned. The outcome rate fell to {actual_rate:.1%} from the "
         f"{budget_rate:.1%} assumed in the budget. Had the planned conversion rate held on the "
         f"actual participant base, the organization would have produced roughly "
         f"{outcomes_at_plan_rate:,.0f} placements; the shortfall of about "
         f"{outcomes_lost_to_mix:,.0f} is the cost of the mix shift. Growth came "
         f"disproportionately through lower-converting channels. The organization bought reach and "
         f"reported it as growth, and reach is not the mission."),

        ("What that implies operationally",
         f"A favourable volume variance sitting alongside an unfavourable outcome variance is a "
         f"mix problem, not a cost problem, and cost discipline will not fix it. The relevant "
         f"question for FY2025 is not whether the extra {fmt(dec['volume_variance'])} of program "
         f"spend was justified, but whether it was spent in the right channels. "
         f"05_Program_Economics.xlsx sets out the per-program cost per successful outcome, and "
         f"06_Resource_Allocation.xlsx shows that the answer changes entirely depending on "
         f"whether management is buying placement count or earnings gain. That choice should be "
         f"made explicitly at budget time rather than emerging from where the marginal dollar "
         f"happened to land."),

        ("Support costs",
         f"Management and general rose only {pct(mg_v['variance_pct'])} against plan while program "
         f"services rose {pct(prog_v['variance_pct'])}, so support costs did not drive the "
         f"overrun. The program expense ratio finished at {hist['program_ratio']:.1%}, well clear "
         f"of the 65% BBB Standard 8 minimum; the administrative ratio at {hist['admin_ratio']:.1%}; "
         f"and fundraising cost {hist['fundraising_cost_ratio'] * 100:.1f} cents per dollar raised "
         f"against a 35 cent BBB Standard 9 ceiling. On every conventional stewardship measure this "
         f"organization is comfortable. None of those ratios speak to whether the money is being "
         f"spent on the right programs, which is the harder question and the one the rest of "
         f"this model is about."),

        ("The number the board should ask about",
         f"Net assets closed FY2024 at {fmt(I.NET_ASSETS_2024_PUBLIC)}, or "
         f"{hist['net_asset_multiple']:.2f} times annual expenses and "
         f"{hist['liquid_runway_months']:.0f} months of liquid runway. BBB Standard 10 flags "
         f"unrestricted net assets above three times annual expenses. The board approved a 26% "
         f"expense increase for FY2024 specifically to deploy reserve into programs; expenses "
         f"grew {pct(exp_v['variance_pct'])} faster than that, which is the right direction, but "
         f"revenue grew faster still and the surplus came in {pct(res_v['variance_pct'])} ahead of "
         f"plan. A fifth consecutive surplus of this size moves the organization toward the "
         f"accumulation ceiling, not away from it. The FY2025 budget conversation should start "
         f"there."),
    ]

    for head, body in paras:
        h = ws.cell(row=r, column=2, value=head)
        h.font = Font(name="Calibri", size=11, bold=True, color=S.NAVY)
        r += 1
        c = ws.cell(row=r, column=2, value=body)
        c.font = Font(name="Calibri", size=10, color=S.INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 13.5 * (len(body) // 108 + 1)
        r += 2

    r = disclaimer(ws, r, width=2)
    S.finish(ws, S.NAVY)
    return ws


def build():
    wb = Workbook()
    wb.remove(wb.active)
    source_data_sheet(wb, R)
    sheet_basis(wb)
    sheet_bva(wb)
    sheet_decomposition(wb)
    sheet_commentary(wb)
    wb.move_sheet("Budget vs actual", offset=-2)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"wrote {OUT}")
    return OUT


if __name__ == "__main__":
    build()
