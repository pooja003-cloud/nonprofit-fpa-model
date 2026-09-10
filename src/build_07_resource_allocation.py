"""
07_Resource_Allocation_Model.xlsx

Allocate an incremental $1,000,000 across the four programs.

The sheet is built for Excel Solver: yellow cells are the decision variables, one
cell is the objective, and the constraints are laid out as explicit left-hand
side / operator / right-hand side rows so they can be entered into Solver
verbatim. The decision cells are pre-seeded with the optimal answer so the
workbook is readable without running Solver, and a linear-programming solution
computed independently in Python sits alongside for comparison.

The substantive point of this workbook: the problem is solved three times under
three objectives, and they disagree sharply. Maximising placement count sends
money to the cheapest channel; maximising earnings gain sends it to the most
expensive one. Presenting a single "optimal" allocation would hide the only
decision that actually belongs to management.
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter as gcl

import inputs as I
import styles as S
import engine
from common_sheets import source_data_sheet, disclaimer

OUT = Path(__file__).resolve().parents[1] / "07_Resource_Allocation_Model.xlsx"
R = {}
A24 = "G"
SRC = "'Source data'!"
PCOL = {p: gcl(2 + i) for i, p in enumerate(I.PROGRAMS)}
YELLOW = PatternFill("solid", fgColor="FFF2CC")
SOLVER_BORDER = Border(*[Side(style="medium", color="BF8F00")] * 4)


def sheet_setup(wb):
    ws = wb.create_sheet("Allocation model")
    S.set_widths(ws, {"A": 46, "B": 17, "C": 17, "D": 17, "E": 17, "F": 18, "G": 3, "H": 56})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Resource allocation model",
                      "Yellow cells are Solver decision variables. The objective cell is outlined "
                      "in navy. Constraints are listed explicitly below so they can be entered "
                      "into Solver as written.", width=8)

    pe = engine.program_economics()["programs"]
    opt = engine.allocate(objective="outcomes")

    # --- problem statement -----------------------------------------------
    S.label(ws, r, 1, "The question", bold=True, size=11, indent=0)
    c = ws.cell(row=r, column=2, value=I.ALLOCATION_NOTE)
    c.font = Font(name="Calibri", size=9.5, color=S.INK)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    ws.row_dimensions[r].height = 13 * (len(I.ALLOCATION_NOTE) // 108 + 1)
    r += 2

    # --- budget -----------------------------------------------------------
    r = S.section(ws, r, "Budget", width=8)
    for name, val, fmt, key, note in (
            ("Incremental funding available", I.ALLOCATION_BUDGET, S.MONEY, "a.budget", ""),
            ("Administrative load", I.ALLOCATION_CONSTRAINTS["admin_load"], S.PCT1, "a.adminload",
             "A share of any new funding must support the administrative capacity to deliver it. "
             "Ignoring this is how allocation models produce answers that cannot be executed."),
            ("Administrative provision", None, S.MONEY, "a.admin", ""),
            ("Deployable to programs", None, S.MONEY, "a.deployable", "")):
        S.label(ws, r, 1, name, size=10, indent=1,
                bold=name == "Deployable to programs")
        if val is None:
            f = (f"=B{R['a.budget']}*B{R['a.adminload']}" if key == "a.admin"
                 else f"=B{R['a.budget']}-B{R['a.admin']}")
            S.put(ws, r, 2, f, fmt=fmt, bold=key == "a.deployable")
        else:
            S.put(ws, r, 2, val, fmt=fmt, is_input=True)
        if note:
            n = ws.cell(row=r, column=8, value=note)
            n.font = Font(name="Calibri", size=8.5, color=S.MUTED)
            n.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 26
        R[key] = r
        r += 1
    r += 1

    # --- program parameters ------------------------------------------------
    r = S.section(ws, r, "Program parameters (from 06_Program_Economics)", width=8)
    r = S.header_row(ws, r, ["Parameter"] + I.PROGRAMS + ["Total", "", "Note"])

    def prow(name, fn, fmt, key, total=True, note=""):
        nonlocal r
        S.label(ws, r, 1, name, size=9.5, indent=1)
        for i, p in enumerate(I.PROGRAMS):
            S.put(ws, r, 2 + i, fn(p), fmt=fmt, is_input=True)
        if total:
            c = ws.cell(row=r, column=6, value=f"=SUM(B{r}:E{r})")
            c.number_format = fmt
            c.font = Font(name="Calibri", size=10, bold=True, color=S.NAVY)
            c.alignment = Alignment(horizontal="right")
        if note:
            n = ws.cell(row=r, column=8, value=note)
            n.font = Font(name="Calibri", size=8.5, color=S.MUTED)
            n.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 26
        R[key] = r
        r += 1

    prow("Marginal cost per participant", lambda p: pe[p]["cost_per_participant"], S.MONEY2,
         "a.cpp", total=False,
         note="Assumed constant at current average cost. A real expansion would face rising "
              "marginal cost as the easiest-to-reach participants are exhausted; that "
              "non-linearity is acknowledged and not modelled.")
    prow("Marginal cost per successful outcome", lambda p: pe[p]["cost_per_outcome"], S.MONEY2,
         "a.cpo", total=False)
    prow("Average earnings gain per outcome", lambda p: pe[p]["salary_gain"], S.MONEY,
         "a.gain", total=False)
    prow("Staff hours per participant", lambda p: pe[p]["staff_hours_per_participant"], S.NUM1,
         "a.hours", total=False)
    prow("Current participants", lambda p: pe[p]["participants"], S.NUM, "a.curparts")
    prow("Maximum servable participants", lambda p: pe[p]["max_participants"], S.NUM, "a.maxparts")
    prow("Capacity headroom (participants)",
         lambda p: pe[p]["max_participants"] - pe[p]["participants"], S.NUM, "a.headroom",
         note="What the allocation model can actually buy in each program.")
    r += 1

    # --- decision variables ------------------------------------------------
    r = S.section(ws, r, "Decision variables  -  Solver changes these", width=8)
    S.label(ws, r, 1, "Allocation to each program", bold=True, size=10, indent=1)
    for i, p in enumerate(I.PROGRAMS):
        c = ws.cell(row=r, column=2 + i, value=round(opt["detail"][p]["allocation"], 2))
        c.number_format = S.MONEY
        c.font = Font(name="Calibri", size=11, bold=True, color="7F6000")
        c.fill = YELLOW
        c.border = SOLVER_BORDER
        c.alignment = Alignment(horizontal="right")
    tot = ws.cell(row=r, column=6, value=f"=SUM(B{r}:E{r})")
    tot.number_format = S.MONEY
    tot.font = Font(name="Calibri", size=11, bold=True, color=S.NAVY)
    tot.alignment = Alignment(horizontal="right")
    ws.row_dimensions[r].height = 22
    R["a.x"] = r
    n = ws.cell(row=r, column=8, value="Pre-seeded with the optimal answer so the workbook reads "
                                       "correctly without running Solver. Clear them to zero and "
                                       "re-solve to verify.")
    n.font = Font(name="Calibri", size=8.5, italic=True, color=S.MUTED)
    n.alignment = Alignment(wrap_text=True, vertical="top")
    r += 1
    S.label(ws, r, 1, "Share of deployable funding", size=9.5, indent=2)
    for i in range(4):
        L = gcl(2 + i)
        c = ws.cell(row=r, column=2 + i, value=f"={L}{R['a.x']}/$B${R['a.deployable']}")
        c.number_format = S.PCT1
        c.alignment = Alignment(horizontal="right")
    R["a.share"] = r
    r += 2

    # --- outcomes of the allocation ---------------------------------------
    r = S.section(ws, r, "What the allocation buys", width=8)

    def crow(name, fn, fmt, key, bold=False, note=""):
        nonlocal r
        S.label(ws, r, 1, name, size=10, indent=1, bold=bold)
        for i in range(4):
            L = gcl(2 + i)
            c = ws.cell(row=r, column=2 + i, value=fn(L))
            c.number_format = fmt
            c.font = Font(name="Calibri", size=10, bold=bold)
            c.alignment = Alignment(horizontal="right")
        c = ws.cell(row=r, column=6, value=f"=SUM(B{r}:E{r})")
        c.number_format = fmt
        c.font = Font(name="Calibri", size=10, bold=True, color=S.NAVY)
        c.alignment = Alignment(horizontal="right")
        if note:
            nn = ws.cell(row=r, column=8, value=note)
            nn.font = Font(name="Calibri", size=8.5, color=S.MUTED)
            nn.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 26
        R[key] = r
        r += 1

    crow("Additional participants", lambda L: f"={L}{R['a.x']}/{L}{R['a.cpp']}", S.NUM1,
         "a.addparts")
    crow("Additional successful outcomes", lambda L: f"={L}{R['a.x']}/{L}{R['a.cpo']}", S.NUM1,
         "a.addout", bold=True)
    crow("Additional earnings gain",
         lambda L: f"={L}{R['a.addout']}*{L}{R['a.gain']}", S.MONEY, "a.addgain")
    crow("Additional staff hours required",
         lambda L: f"={L}{R['a.addparts']}*{L}{R['a.hours']}", S.NUM1, "a.addhours")
    crow("Participants after allocation",
         lambda L: f"={L}{R['a.curparts']}+{L}{R['a.addparts']}", S.NUM1, "a.newparts")
    crow("Capacity utilisation after allocation",
         lambda L: f"={L}{R['a.newparts']}/{L}{R['a.maxparts']}", S.PCT1, "a.newutil",
         note="Must not exceed 100%. That is the capacity constraint.")
    r += 1

    # --- objective ---------------------------------------------------------
    r = S.section(ws, r, "Objective cell  -  Solver maximises this", width=8)
    S.label(ws, r, 1, "Total additional successful outcomes", bold=True, size=11, indent=1)
    obj = ws.cell(row=r, column=2, value=f"=F{R['a.addout']}")
    obj.number_format = S.NUM1
    obj.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    obj.fill = PatternFill("solid", fgColor=S.NAVY)
    obj.border = Border(*[Side(style="medium", color=S.NAVY)] * 4)
    obj.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[r].height = 24
    R["a.obj"] = r
    r += 1
    S.label(ws, r, 1, "Total additional earnings gain (alternative objective)", size=10, indent=1)
    S.put(ws, r, 2, f"=F{R['a.addgain']}", fmt=S.MONEY, bold=True)
    R["a.obj_gain"] = r
    r += 1
    S.label(ws, r, 1, "Average earnings gain per outcome funded", size=10, indent=1)
    S.put(ws, r, 2, f"=F{R['a.addgain']}/F{R['a.addout']}", fmt=S.MONEY)
    R["a.avggain"] = r
    r += 1
    S.label(ws, r, 1, "Blended cost per additional outcome", size=10, indent=1)
    S.put(ws, r, 2, f"=B{R['a.deployable']}/F{R['a.addout']}", fmt=S.MONEY2)
    R["a.blendedcpo"] = r
    r += 2

    # --- constraints -------------------------------------------------------
    r = S.section(ws, r, "Constraints  -  enter these into Solver as written", width=8)
    r = S.header_row(ws, r, ["Constraint", "Cell reference", "Operator", "Limit", "Current value",
                             "Status", "", "Why it exists"])
    C = I.ALLOCATION_CONSTRAINTS
    cons = []
    cons.append(("Total allocation within deployable budget", f"$F${R['a.x']}", "<=",
                 f"=B{R['a.deployable']}", f"=F{R['a.x']}",
                 "The organization cannot spend money it does not have."))
    for i, p in enumerate(I.PROGRAMS):
        L = gcl(2 + i)
        cons.append((f"{p}: minimum share", f"${L}${R['a.x']}", ">=",
                     f"=$B${R['a.deployable']}*{C['min_share_per_program']}", f"={L}{R['a.x']}",
                     "No program is starved to zero. Abrupt defunding destroys delivery teams that "
                     "take years to rebuild." if i == 0 else ""))
    for i, p in enumerate(I.PROGRAMS):
        L = gcl(2 + i)
        cons.append((f"{p}: maximum share", f"${L}${R['a.x']}", "<=",
                     f"=$B${R['a.deployable']}*{C['max_share_per_program']}", f"={L}{R['a.x']}",
                     "Concentration limit. A single-program bet is not a portfolio." if i == 0 else ""))
    for i, p in enumerate(I.PROGRAMS):
        L = gcl(2 + i)
        cons.append((f"{p}: capacity ceiling", f"${L}${R['a.newparts']}", "<=",
                     f"=${L}${R['a.maxparts']}", f"={L}{R['a.newparts']}",
                     "Demand and delivery ceiling. Money cannot buy participants who are not "
                     "there." if i == 0 else ""))
    cons.append(("Incremental staff hours available", f"$F${R['a.addhours']}", "<=",
                 C["staff_hours_available"], f"=F{R['a.addhours']}",
                 "Staff capacity is a separate scarce input from money and binds independently."))
    cons.append(("All allocations non-negative", f"$B${R['a.x']}:$E${R['a.x']}", ">=", 0,
                 f"=MIN(B{R['a.x']}:E{R['a.x']})",
                 "Funding cannot be withdrawn from a program through this model."))

    first_con = r
    for name, ref, op, limit, cur, why in cons:
        S.label(ws, r, 1, name, size=9.5, indent=1)
        c = ws.cell(row=r, column=2, value=ref)
        c.font = Font(name="Consolas", size=9, color=S.MUTED)
        c.alignment = Alignment(horizontal="center")
        c = ws.cell(row=r, column=3, value=op)
        c.font = Font(name="Consolas", size=10, bold=True)
        c.alignment = Alignment(horizontal="center")
        c = ws.cell(row=r, column=4, value=limit)
        c.number_format = S.MONEY if isinstance(limit, str) and "deployable" in limit else S.NUM1
        c.alignment = Alignment(horizontal="right")
        c = ws.cell(row=r, column=5, value=cur)
        c.number_format = S.NUM1
        c.alignment = Alignment(horizontal="right")
        # The tolerance has to be added for <= and subtracted for >=, otherwise a
        # constraint sitting exactly on its limit reads as violated rather than binding.
        satisfied = f"E{r}<=D{r}+0.01" if op == "<=" else f"E{r}>=D{r}-0.01"
        st = ws.cell(row=r, column=6, value=(
            f'=IF({satisfied},IF(ABS(E{r}-D{r})<0.01,"BINDING","ok"),"VIOLATED")'))
        st.font = Font(name="Calibri", size=9, bold=True)
        st.alignment = Alignment(horizontal="center")
        if why:
            n = ws.cell(row=r, column=8, value=why)
            n.font = Font(name="Calibri", size=8.5, color=S.MUTED)
            n.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 26
        r += 1
    last_con = r - 1
    R["a.con_first"], R["a.con_last"] = first_con, last_con

    ws.conditional_formatting.add(
        f"F{first_con}:F{last_con}",
        CellIsRule(operator="equal", formula=['"VIOLATED"'],
                   fill=PatternFill("solid", bgColor="FFC7CE")))
    ws.conditional_formatting.add(
        f"F{first_con}:F{last_con}",
        CellIsRule(operator="equal", formula=['"BINDING"'],
                   fill=PatternFill("solid", bgColor="FFEB9C")))
    r += 1
    S.label(ws, r, 1, "All constraints satisfied", bold=True, size=10)
    c = ws.cell(row=r, column=6,
                value=f'=IF(COUNTIF(F{first_con}:F{last_con},"VIOLATED")=0,"FEASIBLE","INFEASIBLE")')
    c.font = Font(name="Calibri", size=11, bold=True)
    c.alignment = Alignment(horizontal="center")
    r += 2
    r = disclaimer(ws, r, width=8)
    S.finish(ws, "BF8F00")
    return ws


def sheet_objectives(wb):
    ws = wb.create_sheet("Objective comparison")
    S.set_widths(ws, {"A": 40, "B": 18, "C": 18, "D": 18, "E": 18, "F": 3, "G": 56})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "The same money, three different objectives",
                      "Each column is a separate optimisation under the same constraints. They "
                      "disagree, and the disagreement is the finding.", width=7)

    comp = engine.allocation_objective_comparison()
    rows = {row["objective"]: row for row in comp["rows"]}
    order = ["outcomes", "balanced", "earnings"]
    labels = {"outcomes": "Maximise placements", "balanced": "Balanced (quality floor)",
              "earnings": "Maximise earnings gain"}

    r = S.header_row(ws, r, ["", "Maximise placements", "Balanced (quality floor)",
                             "Maximise earnings gain", "Spread", "", "Note"])

    r = S.section(ws, r, "Allocation", width=7)
    first_alloc = r
    for p in I.PROGRAMS:
        S.label(ws, r, 1, p, size=10, indent=1)
        vals = [rows[o]["allocation"][p] for o in order]
        for i, v in enumerate(vals):
            S.put(ws, r, 2 + i, v, fmt=S.MONEY)
        S.put(ws, r, 5, max(vals) - min(vals), fmt=S.MONEY, bold=True)
        R[f"obj.alloc_{p}"] = r
        r += 1
    S.label(ws, r, 1, "Total deployable", bold=True, indent=0)
    for i in range(3):
        L = gcl(2 + i)
        c = ws.cell(row=r, column=2 + i, value=f"=SUM({L}{first_alloc}:{L}{r - 1})")
        c.number_format = S.MONEY
        c.alignment = Alignment(horizontal="right")
    S.total_row_style(ws, r, range(1, 6))
    r += 2

    r = S.section(ws, r, "What each buys", width=7)
    metrics = [
        ("Additional placements", "outcomes", S.NUM1,
         "The count of people placed in skill-aligned work."),
        ("Additional earnings gain", "earnings", S.MONEY,
         "First-year earnings increase across all placements funded."),
        ("Average gain per placement", "avg_gain", S.MONEY,
         "The quality of the placements bought, not just the quantity."),
        ("Cost per additional placement", "cost_per_outcome", S.MONEY2, ""),
        ("Placements forgone vs best", "outcomes_forgone", S.NUM1,
         "What choosing this objective costs in placement count."),
        ("Earnings gain forgone vs best", "earnings_forgone", S.MONEY,
         "What choosing this objective costs in income mobility."),
    ]
    for name, field, fmt, note in metrics:
        S.label(ws, r, 1, name, size=10, indent=1)
        vals = [rows[o][field] for o in order]
        for i, v in enumerate(vals):
            S.put(ws, r, 2 + i, v, fmt=fmt)
        S.put(ws, r, 5, max(vals) - min(vals), fmt=fmt, bold=True)
        if note:
            n = ws.cell(row=r, column=7, value=note)
            n.font = Font(name="Calibri", size=8.5, color=S.MUTED)
            n.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 26
        if field == "outcomes":
            R["obj.out"] = r
        if field == "earnings":
            R["obj.earn"] = r
        r += 1
    r += 2

    r = S.section(ws, r, "Reference policies, for comparison", width=7)
    base = comp["runs"]["outcomes"]["comparison"]
    r = S.header_row(ws, r, ["Policy", "Placements", "Earnings gain", "Average gain", "", "", "Note"])
    pol = [
        ("Optimised for placements", base["optimized"],
         rows["outcomes"]["earnings"], rows["outcomes"]["avg_gain"], ""),
        ("Equal split across programs", base["equal_detail"]["outcomes"],
         base["equal_detail"]["earnings"], base["equal_detail"]["avg_gain"],
         "The default when no one wants to choose."),
        ("Pro rata to current program size", base["pro_rata_detail"]["outcomes"],
         base["pro_rata_detail"]["earnings"], base["pro_rata_detail"]["avg_gain"],
         "The default when the budget is rolled forward. It entrenches whatever the current mix "
         "happens to be."),
    ]
    for name, o, e, a, note in pol:
        S.label(ws, r, 1, name, size=10, indent=1)
        S.put(ws, r, 2, o, fmt=S.NUM1)
        S.put(ws, r, 3, e, fmt=S.MONEY)
        S.put(ws, r, 4, a, fmt=S.MONEY)
        if note:
            n = ws.cell(row=r, column=7, value=note)
            n.font = Font(name="Calibri", size=8.5, color=S.MUTED)
            n.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 26
        r += 1
    r += 2

    r = S.section(ws, r, "What this says", width=7)
    ro, rb, re_ = rows["outcomes"], rows["balanced"], rows["earnings"]
    notes = [
        f"Optimising for placement count produces {ro['outcomes']:.0f} additional placements at an "
        f"average earnings gain of ${ro['avg_gain']:,.0f} each. Optimising for earnings gain "
        f"produces only {re_['outcomes']:.0f} placements - {ro['outcomes'] - re_['outcomes']:.0f} "
        f"fewer - but at ${re_['avg_gain']:,.0f} of gain each, and "
        f"${re_['earnings'] - ro['earnings']:,.0f} more earnings gain in total. Same million "
        f"dollars, same constraints, materially different programs funded.",

        f"The balanced case shows the trade is not all-or-nothing. Holding placement count as the "
        f"objective while requiring an average gain of at least "
        f"${I.ALLOCATION_CONSTRAINTS['min_avg_salary_gain']:,} per placement gives "
        f"{rb['outcomes']:.0f} placements - only {ro['outcomes'] - rb['outcomes']:.1f} fewer than "
        f"the placement-maximising answer - while lifting average gain from "
        f"${ro['avg_gain']:,.0f} to ${rb['avg_gain']:,.0f}. Roughly "
        f"{(ro['outcomes'] - rb['outcomes']) / ro['outcomes']:.1%} of placement count buys a "
        f"{rb['avg_gain'] / ro['avg_gain'] - 1:.0%} improvement in placement quality. On these "
        f"assumptions that is a cheap trade, and it is the allocation this analysis recommends.",

        f"Both optimised answers beat the two default policies. An equal split delivers "
        f"{base['equal_detail']['outcomes']:.0f} placements and a pro rata roll-forward "
        f"{base['pro_rata_detail']['outcomes']:.0f}, against {ro['outcomes']:.0f} for the "
        f"optimised allocation - so the analysis is worth roughly "
        f"{ro['outcomes'] - base['pro_rata_detail']['outcomes']:.0f} additional placements a year "
        f"against the most likely default. That is the value of doing this exercise at all.",

        f"A caution on all of it. The model assumes marginal cost per outcome stays at current "
        f"average cost as programs expand. Real expansion faces rising marginal cost: the "
        f"easiest-to-serve participants are already being served. Because the cheapest program is "
        f"also the one most exposed to that effect, the true optimum almost certainly sits closer "
        f"to the balanced answer than the placement-maximising one. The linear model overstates "
        f"the case for concentration, and knowing the direction of that bias is more useful than "
        f"pretending it is not there.",
    ]
    for n in notes:
        c = ws.cell(row=r, column=1, value=n)
        c.font = Font(name="Calibri", size=9.5, color=S.INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
        ws.row_dimensions[r].height = 13 * (len(n) // 128 + 1)
        r += 1
    r += 1

    ch = BarChart()
    ch.type = "col"
    ch.grouping = "clustered"
    ch.title = "Allocation by objective"
    ch.height, ch.width = 8, 17
    ch.y_axis.numFmt = '$#,##0'
    for i, o in enumerate(order):
        ref = Reference(ws, min_col=2 + i, min_row=first_alloc, max_row=first_alloc + 3)
        ch.series.append(Series(ref, title=labels[o]))
    ch.set_categories(Reference(ws, min_col=1, min_row=first_alloc, max_row=first_alloc + 3))
    ws.add_chart(ch, f"A{r}")

    S.finish(ws, S.NAVY)
    return ws


def sheet_frontier(wb):
    ws = wb.create_sheet("Trade-off frontier")
    S.set_widths(ws, {"A": 24, "B": 16, "C": 18, "D": 18, "E": 18, "F": 3, "G": 56})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Trade-off and budget frontiers",
                      "How much placement count each level of placement quality costs, and what "
                      "happens as the budget grows.", width=7)

    tr = engine.allocation_tradeoff_frontier()
    r = S.section(ws, r, "Quality floor against placement count", width=7)
    r = S.header_row(ws, r, ["Minimum average gain", "Placements", "Total earnings gain",
                             "Average gain achieved", "Cost per placement", "", "Note"])
    first = r
    for t in tr:
        S.label(ws, r, 1, f"${t['floor']:,}", size=10, indent=1)
        if t.get("feasible"):
            S.put(ws, r, 2, t["outcomes"], fmt=S.NUM1)
            S.put(ws, r, 3, t["earnings"], fmt=S.MONEY)
            S.put(ws, r, 4, t["avg_gain"], fmt=S.MONEY)
            S.put(ws, r, 5, t["cost_per_outcome"], fmt=S.MONEY2)
        else:
            c = ws.cell(row=r, column=2, value="infeasible")
            c.font = Font(name="Calibri", size=9, italic=True, color=S.WARN)
            c.alignment = Alignment(horizontal="right")
            n = ws.cell(row=r, column=7,
                        value="No allocation satisfying the minimum and maximum share constraints "
                              "can reach this average gain. The constraint set, not the money, is "
                              "what binds here.")
            n.font = Font(name="Calibri", size=8.5, color=S.MUTED)
            n.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 26
        r += 1
    last = r - 1
    R["fr.first"], R["fr.last"] = first, last
    r += 1

    ch = LineChart()
    ch.title = "Placements against the earnings-gain quality floor"
    ch.height, ch.width = 8, 17
    feas_last = max(rr for rr in range(first, last + 1)
                    if isinstance(ws.cell(row=rr, column=2).value, (int, float)))
    ch.series.append(Series(Reference(ws, min_col=2, min_row=first, max_row=feas_last),
                            title="Placements"))
    ch.set_categories(Reference(ws, min_col=1, min_row=first, max_row=feas_last))
    ws.add_chart(ch, f"A{r}")
    r += 17

    fr = engine.allocation_efficient_frontier()
    r = S.section(ws, r, "Budget frontier", width=7)
    r = S.header_row(ws, r, ["Budget", "Placements", "Earnings gain", "Cost per placement",
                             "Staff hours used", "", "Note"])
    bfirst = r
    for f in fr:
        S.label(ws, r, 1, f"${f['budget']:,.0f}", size=10, indent=1)
        S.put(ws, r, 2, f["outcomes"], fmt=S.NUM1)
        S.put(ws, r, 3, f["earnings"], fmt=S.MONEY)
        S.put(ws, r, 4, f["cost_per_outcome"], fmt=S.MONEY2)
        S.put(ws, r, 5, f["hours_utilization"], fmt=S.PCT1)
        if f["hours_binding"]:
            n = ws.cell(row=r, column=7, value="Staff hours fully consumed - beyond this point "
                                               "more money does not buy more outcomes without "
                                               "more staff.")
            n.font = Font(name="Calibri", size=8.5, color=S.WARN)
            n.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    blast = r - 1
    r += 1

    ch2 = LineChart()
    ch2.title = "Additional placements as the budget grows"
    ch2.height, ch2.width = 8, 17
    ch2.series.append(Series(Reference(ws, min_col=2, min_row=bfirst, max_row=blast),
                             title="Placements"))
    ch2.set_categories(Reference(ws, min_col=1, min_row=bfirst, max_row=blast))
    ws.add_chart(ch2, f"A{r}")
    r += 17

    hours_binds = [f for f in fr if f["hours_binding"]]
    r = S.section(ws, r, "Where the binding constraint sits", width=7)
    msg = (f"Staff hours reach full utilisation at a budget of "
           f"${hours_binds[0]['budget']:,.0f}." if hours_binds else
           f"Staff hours never fully bind across the budget range tested, peaking at "
           f"{max(f['hours_utilization'] for f in fr):.0%} utilisation. Within this range the "
           f"binding constraints are the minimum and maximum share limits, not delivery capacity "
           f"or money. That is worth knowing: it means the allocation answer is being driven by "
           f"policy choices the board set, not by operational limits, and the board could revisit "
           f"them if it wanted a different answer.")
    c = ws.cell(row=r, column=1, value=msg)
    c.font = Font(name="Calibri", size=9.5, color=S.INK)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    ws.row_dimensions[r].height = 13 * (len(msg) // 128 + 1)

    S.finish(ws, S.ACCENT)
    return ws


def sheet_solver(wb):
    ws = wb.create_sheet("How to run Solver")
    S.set_widths(ws, {"A": 3, "B": 26, "C": 88})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Running this in Excel Solver",
                      "The decision cells are pre-seeded with the optimal answer. These steps "
                      "reproduce it from scratch.", width=3)

    steps = [
        ("Enable Solver",
         "File > Options > Add-ins > Manage Excel Add-ins > Go, then tick Solver Add-in. It "
         "appears on the Data tab. On Mac it is Tools > Excel Add-ins."),
        ("Clear the decision cells",
         f"On 'Allocation model', set B{R['a.x']}:E{R['a.x']} to zero. If Solver reproduces the "
         "seeded values from a cold start, the model is doing what it claims."),
        ("Set the objective",
         f"Data > Solver. Set Objective to $B${R['a.obj']} on 'Allocation model' and choose Max."),
        ("Set the decision variables",
         f"By Changing Variable Cells: $B${R['a.x']}:$E${R['a.x']}."),
        ("Add the constraints",
         f"Add each row from the constraints block on 'Allocation model' (rows "
         f"{R['a.con_first']} to {R['a.con_last']}). Cell reference is column B, operator column "
         f"C, and the limit is the cell in column D. For example: $F${R['a.x']} <= "
         f"$B${R['a.deployable']}."),
        ("Choose the solving method",
         "Simplex LP. The model is linear in the decision variables - every outcome and cost term "
         "is a constant multiplied by an allocation - so Simplex is both correct and fast. Tick "
         "'Make Unconstrained Variables Non-Negative'."),
        ("Solve",
         "Solver should report an optimal solution and reproduce the seeded allocation. The "
         "Answer and Sensitivity reports are worth generating: the Sensitivity report gives "
         "shadow prices, which tell you what one more dollar or one more staff hour would be "
         "worth at the margin."),
        ("Change the objective",
         f"To reproduce the earnings-maximising answer on 'Objective comparison', set the "
         f"objective to $B${R['a.obj_gain']} instead and re-solve. To reproduce the balanced "
         f"answer, keep $B${R['a.obj']} as the objective and add one further constraint: "
         f"$B${R['a.avggain']} >= {I.ALLOCATION_CONSTRAINTS['min_avg_salary_gain']}."),
    ]
    for i, (head, body) in enumerate(steps, start=1):
        S.label(ws, r, 2, f"{i}. {head}", bold=True, size=10.5)
        c = ws.cell(row=r, column=3, value=body)
        c.font = Font(name="Calibri", size=9.5, color=S.INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 13 * (len(body) // 92 + 1)
        r += 1
    r += 1

    r = S.section(ws, r, "Independent check", width=3)
    opt = engine.allocate(objective="outcomes")
    body = (f"The same problem is solved independently in Python using the CBC linear programming "
            f"solver, in src/engine.py. It reports status '{opt['status']}' with a total of "
            f"{opt['total_additional_outcomes']:.2f} additional placements from "
            f"${opt['deployable']:,.0f} of deployable funding. The automated test suite in "
            f"tests/ recalculates this workbook and checks that the Excel objective cell agrees "
            f"with the Python solution to within a rounding tolerance. Two independent "
            f"implementations agreeing is a much stronger claim than one implementation being "
            f"self-consistent.")
    c = ws.cell(row=r, column=2, value=body)
    c.font = Font(name="Calibri", size=9.5, color=S.INK)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    ws.row_dimensions[r].height = 13 * (len(body) // 112 + 1)
    r += 2

    r = S.section(ws, r, "Known limitations of this formulation", width=3)
    lims = [
        "Marginal cost per outcome is assumed constant at current average cost. Real expansion "
        "faces rising marginal cost, and the effect is strongest for the cheapest program, so the "
        "linear model overstates the case for concentrating funding there.",
        "Outcome rates are held at their current levels. A program that doubles in size may well "
        "convert at a lower rate as it reaches further down the demand curve.",
        "The model allocates one year of incremental funding. It says nothing about the multi-year "
        "commitments that program expansion actually requires - hiring, leases, partner contracts - "
        "which are far harder to reverse than an annual budget line.",
        "Earnings gain is treated as the only measure of outcome quality. Job durability, career "
        "trajectory, and the value of serving harder-to-place participants are all excluded, and "
        "all three would plausibly favour the more intensive programs.",
    ]
    for l in lims:
        c = ws.cell(row=r, column=2, value="- " + l)
        c.font = Font(name="Calibri", size=9, color=S.INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ws.row_dimensions[r].height = 13 * (len(l) // 112 + 1)
        r += 1

    S.finish(ws, S.MUTED)
    return ws


def build():
    wb = Workbook()
    wb.remove(wb.active)
    source_data_sheet(wb, R)
    sheet_setup(wb)
    sheet_objectives(wb)
    sheet_frontier(wb)
    sheet_solver(wb)
    wb.move_sheet("Allocation model", offset=-2)
    wb.save(OUT)
    print(f"wrote {OUT}")
    return OUT


if __name__ == "__main__":
    build()
