"""
05_Program_Economics.xlsx

Unit economics for each of the four programs, and the comparison that tells
management where a marginal dollar does the most good.

The headline finding is uncomfortable and worth stating up front: the four
programs cannot be ranked on a single measure. Digital Learning is the cheapest
per placement and Credential Support is nearly twice as expensive, but Credential
Support delivers more than twice the earnings gain per placement. Which one is
"better" depends entirely on whether the organization is buying placement counts
or income mobility, and that is a board decision, not an analytical one.

The social return section is last, deliberately, and is reported as a range
rather than a point estimate.
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, ScatterChart, Reference, Series
from openpyxl.chart.marker import Marker
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.utils import get_column_letter as gcl

import inputs as I
import styles as S
import engine
from common_sheets import source_data_sheet, disclaimer

OUT = Path(__file__).resolve().parents[1] / "excel-models" / "05_Program_Economics.xlsx"
R = {}
A24 = "G"
SRC = "'Source data'!"
PIN = "'Program inputs'!"
PCOL = {p: gcl(2 + i) for i, p in enumerate(I.PROGRAMS)}   # B..E
TOTCOL = gcl(2 + len(I.PROGRAMS))                           # F


def sheet_program_inputs(wb):
    ws = wb.create_sheet("Program inputs")
    S.set_widths(ws, {"A": 42, "B": 17, "C": 17, "D": 17, "E": 17, "F": 17, "G": 11, "H": 56})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Program inputs",
                      "The four-program portfolio. Participant and outcome counts reconcile to the "
                      "FY2024 totals derived from published data; costs reconcile to the modelled "
                      "program expense pool. The split across programs is an analyst assumption.",
                      width=8)
    r = S.header_row(ws, r, ["Input"] + I.PROGRAMS + ["Total", "Tier", "Note"])

    def row(name, field, fmt=S.NUM, key=None, note="", total=True, tier="SYNTHETIC"):
        nonlocal r
        S.label(ws, r, 1, name, size=9.5, indent=1)
        for p in I.PROGRAMS:
            S.put(ws, r, 2 + I.PROGRAMS.index(p), I.PROGRAM_BASE[p][field], fmt=fmt, is_input=True)
        if total:
            c = ws.cell(row=r, column=6, value=f"=SUM(B{r}:E{r})")
            c.number_format = fmt
            c.font = Font(name="Calibri", size=10, bold=True, color=S.NAVY)
            c.alignment = Alignment(horizontal="right")
        S.tier_cell(ws, r, 7, tier)
        if note:
            n = ws.cell(row=r, column=8, value=note)
            n.font = Font(name="Calibri", size=8.5, color=S.MUTED)
            n.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 26
        if key:
            R[key] = r
        r += 1

    r = S.section(ws, r, "Volume", width=8)
    row("Participants enrolled", "participants", S.NUM, "pi.parts",
        "Sums to the FY2024 total interpolated from the published 2022 and 2025 observations.")
    row("Successful outcomes (placements)", "placements", S.NUM, "pi.out",
        "Sums to the FY2024 placement total on the same basis.")
    row("Completion rate", "completion_rate", S.PCT1, "pi.comp", total=False,
        note="Share of enrolled participants who finish the program. Distinct from the outcome rate.")
    r += 1

    r = S.section(ws, r, "Cost", width=8)
    row("Program cost", "program_cost", S.MONEY, "pi.cost",
        "Sums to the modelled FY2024 program expense pool.")
    row("Direct cost share", "direct_cost_share", S.PCT1, "pi.directshare", total=False,
        note="Remainder is allocated indirect cost - supervision, platform, evaluation, shared space.")
    r += 1

    r = S.section(ws, r, "Delivery and capacity", width=8)
    row("Staff hours per participant", "staff_hours_per_participant", S.NUM1, "pi.hours",
        total=False,
        note="Set so total delivery hours tie to the staffing model: 121 FTE x 74% program share x "
             "1,680 productive hours.")
    row("Maximum servable participants", "max_participants", S.NUM, "pi.maxparts",
        "Demand and delivery ceiling. Binds the allocation model in 07.")
    row("Minimum viable funding", "min_funding", S.MONEY, "pi.minfund",
        "Below this the program cannot retain its delivery team.")
    r += 1

    r = S.section(ws, r, "Outcome value", width=8)
    row("Average salary gain per placement", "salary_gain", S.MONEY, "pi.gain", total=False,
        note="Increase in annual earnings against the participant's pre-program position.")
    r += 1

    r = S.section(ws, r, "Reconciliation to public data", width=8)
    checks = [
        ("Participants tie to the FY2024 derived total",
         f"=B{R['pi.parts']}+C{R['pi.parts']}+D{R['pi.parts']}+E{R['pi.parts']}",
         f"={SRC}{A24}{R['src.parts']}"),
        ("Outcomes tie to the FY2024 derived total",
         f"=B{R['pi.out']}+C{R['pi.out']}+D{R['pi.out']}+E{R['pi.out']}",
         f"={SRC}{A24}{R['src.places']}"),
        ("Program costs tie to the modelled expense pool",
         f"=B{R['pi.cost']}+C{R['pi.cost']}+D{R['pi.cost']}+E{R['pi.cost']}",
         f"={SRC}{A24}{R['src.prog']}"),
    ]
    r = S.header_row(ws, r, ["Check", "Modelled", "Control total", "Difference", "Result", "", "", ""])
    for name, modelled, control in checks:
        S.label(ws, r, 1, name, size=9.5, indent=1)
        ws.cell(row=r, column=2, value=modelled).number_format = S.MONEY2
        ws.cell(row=r, column=3, value=control).number_format = S.MONEY2
        ws.cell(row=r, column=4, value=f"=B{r}-C{r}").number_format = S.MONEY2
        c = ws.cell(row=r, column=5, value=f'=IF(ABS(D{r})<1,"PASS","FAIL")')
        c.font = Font(name="Calibri", size=10, bold=True)
        c.alignment = Alignment(horizontal="center")
        for col in (2, 3, 4):
            ws.cell(row=r, column=col).alignment = Alignment(horizontal="right")
        r += 1
    r += 1

    r = S.section(ws, r, "What each program is", width=8)
    for p in I.PROGRAMS:
        S.label(ws, r, 1, f"{p} ({I.PROGRAM_CODES[p]})", bold=True, size=10, indent=1)
        c = ws.cell(row=r, column=2, value=I.PROGRAM_DESCRIPTION[p])
        c.font = Font(name="Calibri", size=9, color=S.INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 13 * (len(I.PROGRAM_DESCRIPTION[p]) // 100 + 1)
        r += 1
    r += 1
    r = disclaimer(ws, r, width=8)
    S.finish(ws, S.ACCENT)
    return ws


def sheet_economics(wb):
    ws = wb.create_sheet("Program economics")
    S.set_widths(ws, {"A": 44, "B": 17, "C": 17, "D": 17, "E": 17, "F": 17, "G": 3, "H": 58})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Program unit economics, FY2024",
                      "Every figure calculated from 'Program inputs'. Change an input there and "
                      "this sheet moves.", width=8)
    r = S.header_row(ws, r, ["Measure"] + I.PROGRAMS + ["Portfolio", "", "Note"])

    def row(name, fn, total_fn=None, fmt=S.MONEY, bold=False, indent=1, key=None, note=""):
        nonlocal r
        S.label(ws, r, 1, name, bold=bold, size=10, indent=indent)
        for p in I.PROGRAMS:
            col = PCOL[p]
            c = ws.cell(row=r, column=2 + I.PROGRAMS.index(p), value=fn(col))
            c.number_format = fmt
            c.font = Font(name="Calibri", size=10, bold=bold)
            c.alignment = Alignment(horizontal="right")
        if total_fn:
            c = ws.cell(row=r, column=6, value=total_fn())
            c.number_format = fmt
            c.font = Font(name="Calibri", size=10, bold=True, color=S.NAVY)
            c.alignment = Alignment(horizontal="right")
            c.fill = PatternFill("solid", fgColor=S.NAVY_LIGHT)
        if note:
            n = ws.cell(row=r, column=8, value=note)
            n.font = Font(name="Calibri", size=8.5, color=S.MUTED)
            n.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 26
        if key:
            R[key] = r
        r += 1

    P = PIN
    pp, po, pc = R["pi.parts"], R["pi.out"], R["pi.cost"]

    r = S.section(ws, r, "Scale", width=8)
    row("Participants enrolled", lambda c: f"={P}{c}{pp}",
        lambda: f"=SUM(B{r}:E{r})", fmt=S.NUM, key="e.parts")
    row("Participants completing", lambda c: f"={P}{c}{pp}*{P}{c}{R['pi.comp']}",
        lambda: f"=SUM(B{r}:E{r})", fmt=S.NUM1, key="e.comp")
    row("Successful outcomes", lambda c: f"={P}{c}{po}",
        lambda: f"=SUM(B{r}:E{r})", fmt=S.NUM, key="e.out")
    row("Program cost", lambda c: f"={P}{c}{pc}", lambda: f"=SUM(B{r}:E{r})", key="e.cost")
    row("Share of program spend", lambda c: f"={c}{R['e.cost']}/$F${R['e.cost']}",
        lambda: f"=SUM(B{r}:E{r})", fmt=S.PCT1, key="e.costshare")
    r += 1

    r = S.section(ws, r, "Unit cost", width=8)
    row("Cost per participant", lambda c: f"={c}{R['e.cost']}/{c}{R['e.parts']}",
        lambda: f"=F{R['e.cost']}/F{R['e.parts']}", fmt=S.MONEY2, key="e.cpp",
        note="Program cost divided by enrolled participants.")
    row("Cost per completer", lambda c: f"={c}{R['e.cost']}/{c}{R['e.comp']}",
        lambda: f"=F{R['e.cost']}/F{R['e.comp']}", fmt=S.MONEY2, key="e.cpc")
    row("Cost per successful outcome", lambda c: f"={c}{R['e.cost']}/{c}{R['e.out']}",
        lambda: f"=F{R['e.cost']}/F{R['e.out']}", fmt=S.MONEY2, bold=True, indent=0, key="e.cpo",
        note="The single most decision-relevant cost measure in this workbook.")
    r += 1

    r = S.section(ws, r, "Conversion", width=8)
    row("Completion rate", lambda c: f"={P}{c}{R['pi.comp']}",
        lambda: f"=F{R['e.comp']}/F{R['e.parts']}", fmt=S.PCT1, key="e.comprate")
    row("Outcome rate (of enrolled)", lambda c: f"={c}{R['e.out']}/{c}{R['e.parts']}",
        lambda: f"=F{R['e.out']}/F{R['e.parts']}", fmt=S.PCT1, key="e.outrate")
    row("Outcome rate (of completers)", lambda c: f"={c}{R['e.out']}/{c}{R['e.comp']}",
        lambda: f"=F{R['e.out']}/F{R['e.comp']}", fmt=S.PCT1, key="e.outcomprate",
        note="Separating these two matters. A low outcome rate caused by drop-out is a retention "
             "problem; one caused by completers not being placed is a placement problem. They have "
             "different fixes.")
    r += 1

    r = S.section(ws, r, "Cost structure", width=8)
    row("Direct cost", lambda c: f"={c}{R['e.cost']}*{P}{c}{R['pi.directshare']}",
        lambda: f"=SUM(B{r}:E{r})", key="e.direct")
    row("Indirect cost", lambda c: f"={c}{R['e.cost']}-{c}{R['e.direct']}",
        lambda: f"=SUM(B{r}:E{r})", key="e.indirect")
    row("Indirect as a share of program cost",
        lambda c: f"={c}{R['e.indirect']}/{c}{R['e.cost']}",
        lambda: f"=F{R['e.indirect']}/F{R['e.cost']}", fmt=S.PCT1, key="e.indirectshare")
    r += 1

    r = S.section(ws, r, "Delivery capacity", width=8)
    row("Staff hours consumed", lambda c: f"={P}{c}{pp}*{P}{c}{R['pi.hours']}",
        lambda: f"=SUM(B{r}:E{r})", fmt=S.NUM1, key="e.hours")
    row("Implied programme FTE",
        lambda c: f"={c}{R['e.hours']}/{I.STAFF_HOURS_PER_FTE}",
        lambda: f"=SUM(B{r}:E{r})", fmt=S.NUM1, key="e.fte")
    row("Staff hours per successful outcome",
        lambda c: f"={c}{R['e.hours']}/{c}{R['e.out']}",
        lambda: f"=F{R['e.hours']}/F{R['e.out']}", fmt=S.NUM1, key="e.hoursout",
        note="The capacity equivalent of cost per outcome. Money is not the only scarce input.")
    row("Capacity utilisation", lambda c: f"={c}{R['e.parts']}/{P}{c}{R['pi.maxparts']}",
        lambda: f"=F{R['e.parts']}/SUM({P}B{R['pi.maxparts']}:{P}E{R['pi.maxparts']})",
        fmt=S.PCT1, key="e.util",
        note="How full each programme is against its ceiling. Headroom here is what the "
             "allocation model can actually buy.")
    row("Funding utilisation against minimum viable",
        lambda c: f"={c}{R['e.cost']}/{P}{c}{R['pi.minfund']}",
        lambda: f"=F{R['e.cost']}/SUM({P}B{R['pi.minfund']}:{P}E{R['pi.minfund']})",
        fmt=S.MULT, key="e.fundutil",
        note="How far above its funding floor each programme sits. A ratio near 1.0 means the "
             "programme has no room to absorb a cut without losing its team.")
    r += 1

    r = S.section(ws, r, "Outcome value", width=8)
    row("Average salary gain per placement", lambda c: f"={P}{c}{R['pi.gain']}",
        lambda: f"=F{r + 1}/F{R['e.out']}", key="e.gain")
    row("Total first-year earnings gain",
        lambda c: f"={c}{R['e.out']}*{c}{R['e.gain']}", lambda: f"=SUM(B{r}:E{r})", key="e.totgain")
    row("Earnings gain per dollar of programme cost",
        lambda c: f"={c}{R['e.totgain']}/{c}{R['e.cost']}",
        lambda: f"=F{R['e.totgain']}/F{R['e.cost']}", fmt=S.MULT, key="e.gainperdollar",
        note="First-year gross earnings gain per dollar spent. Not a social return figure - no "
             "attribution, deadweight, drop-off or discounting has been applied yet.")

    ws.conditional_formatting.add(
        f"B{R['e.cpo']}:E{R['e.cpo']}",
        ColorScaleRule(start_type="min", start_color="63BE7B",
                       end_type="max", end_color="F8696B"))
    ws.conditional_formatting.add(
        f"B{R['e.gainperdollar']}:E{R['e.gainperdollar']}",
        ColorScaleRule(start_type="min", start_color="F8696B",
                       end_type="max", end_color="63BE7B"))
    ws.conditional_formatting.add(
        f"B{R['e.util']}:E{R['e.util']}",
        DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color=S.NAVY))
    S.freeze(ws, "B5")
    S.finish(ws, S.NAVY)
    return ws


def sheet_comparison(wb):
    ws = wb.create_sheet("Program comparison")
    S.set_widths(ws, {"A": 26, "B": 15, "C": 15, "D": 15, "E": 15, "F": 15, "G": 15, "H": 15,
                      "I": 3, "J": 54})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Program comparison",
                      "The same four programs ranked on each measure. They rank differently on "
                      "each one, which is the point.", width=10)

    E = "'Program economics'!"
    r = S.header_row(ws, r, ["Program", "Cost", "Participants", "Outcomes", "Cost per participant",
                             "Cost per outcome", "Outcome rate", "Gain per $", "", "Note"])
    first = r
    for p in I.PROGRAMS:
        col = PCOL[p]
        S.label(ws, r, 1, p, size=10, indent=1)
        cells = [(2, f"={E}{col}{R['e.cost']}", S.MONEY),
                 (3, f"={E}{col}{R['e.parts']}", S.NUM),
                 (4, f"={E}{col}{R['e.out']}", S.NUM),
                 (5, f"={E}{col}{R['e.cpp']}", S.MONEY2),
                 (6, f"={E}{col}{R['e.cpo']}", S.MONEY2),
                 (7, f"={E}{col}{R['e.outrate']}", S.PCT1),
                 (8, f"={E}{col}{R['e.gainperdollar']}", S.MULT)]
        for c_, f_, fmt in cells:
            cell = ws.cell(row=r, column=c_, value=f_)
            cell.number_format = fmt
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(horizontal="right")
        R[f"cmp.{p}"] = r
        r += 1
    last = r - 1
    S.label(ws, r, 1, "Portfolio", bold=True, indent=0)
    for c_, key, fmt in ((2, "e.cost", S.MONEY), (3, "e.parts", S.NUM), (4, "e.out", S.NUM),
                         (5, "e.cpp", S.MONEY2), (6, "e.cpo", S.MONEY2),
                         (7, "e.outrate", S.PCT1), (8, "e.gainperdollar", S.MULT)):
        cell = ws.cell(row=r, column=c_, value=f"={E}F{R[key]}")
        cell.number_format = fmt
        cell.alignment = Alignment(horizontal="right")
    S.total_row_style(ws, r, range(1, 9), double=True)
    r += 2

    r = S.section(ws, r, "Rank on each measure (1 = best)", width=10)
    r = S.header_row(ws, r, ["Program", "", "", "", "Cost per participant", "Cost per outcome",
                             "Outcome rate", "Gain per $", "", "Combined rank"])
    rank_first = r
    for p in I.PROGRAMS:
        row_ = R[f"cmp.{p}"]
        S.label(ws, r, 1, p, size=10, indent=1)
        for c_, asc in ((5, True), (6, True), (7, False), (8, False)):
            L = gcl(c_)
            f = (f"=RANK({L}{row_},${L}${first}:${L}${last},{1 if asc else 0})")
            cell = ws.cell(row=r, column=c_, value=f)
            cell.number_format = "0"
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(horizontal="center")
        cell = ws.cell(row=r, column=10, value=f"=SUM(E{r}:H{r})")
        cell.number_format = "0"
        cell.font = Font(name="Calibri", size=10, bold=True, color=S.NAVY)
        cell.alignment = Alignment(horizontal="center")
        r += 1
    r += 1
    S.note(ws, r, "A combined rank is shown because reviewers always ask for one, and it is the "
                  "least useful number on the sheet. Adding ordinal ranks across measures with "
                  "different units implies they matter equally, which no one has decided. The "
                  "honest reading is the four columns to its left.", width=10)
    r += 2

    # charts
    ch = BarChart()
    ch.type = "col"
    ch.title = "Cost per successful outcome"
    ch.height, ch.width = 8, 16
    ch.y_axis.numFmt = '$#,##0'
    ch.series.append(Series(Reference(ws, min_col=6, min_row=first, max_row=last), title="Cost per outcome"))
    ch.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
    ws.add_chart(ch, f"A{r}")

    ch2 = BarChart()
    ch2.type = "col"
    ch2.title = "First-year earnings gain per dollar of program cost"
    ch2.height, ch2.width = 8, 16
    ch2.series.append(Series(Reference(ws, min_col=8, min_row=first, max_row=last), title="Gain per $"))
    ch2.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
    ws.add_chart(ch2, f"A{r + 17}")

    r += 34
    r = S.section(ws, r, "What the comparison says", width=10)
    pe = engine.program_economics()
    P = pe["programs"]
    cheapest = min(I.PROGRAMS, key=lambda p: P[p]["cost_per_outcome"])
    dearest = max(I.PROGRAMS, key=lambda p: P[p]["cost_per_outcome"])
    best_gain = max(I.PROGRAMS, key=lambda p: P[p]["earnings_gain_per_dollar"])
    notes = [
        f"{cheapest} produces a placement for ${P[cheapest]['cost_per_outcome']:,.0f} against "
        f"${P[dearest]['cost_per_outcome']:,.0f} for {dearest} - a "
        f"{P[dearest]['cost_per_outcome'] / P[cheapest]['cost_per_outcome']:.1f}x difference. On "
        f"cost per placement alone the allocation decision looks obvious.",

        f"It is not obvious. {cheapest} converts only "
        f"{P[cheapest]['outcome_rate']:.1%} of enrolled participants and its placements carry an "
        f"average earnings gain of ${P[cheapest]['salary_gain']:,.0f}. Credential Support converts "
        f"{P['Credential Support']['outcome_rate']:.1%} at "
        f"${P['Credential Support']['salary_gain']:,.0f} of gain - "
        f"{P['Credential Support']['salary_gain'] / P[cheapest]['salary_gain']:.1f} times as much "
        f"income mobility per person placed.",

        f"Measured on earnings gain generated per dollar spent, the ranking changes: "
        f"{best_gain} leads at ${P[best_gain]['earnings_gain_per_dollar']:,.2f} of first-year "
        f"gain per dollar. Cost per placement and value per placement point in different "
        f"directions, and no amount of further analysis resolves that. It is a choice about what "
        f"the organization is for.",

        f"Capacity is the constraint that decides how much the choice is worth. "
        f"{cheapest} is running at {P[cheapest]['capacity_utilization']:.0%} of its ceiling and "
        f"Credential Support at {P['Credential Support']['capacity_utilization']:.0%}. Headroom, "
        f"not preference, sets the upper bound on how far funding can be tilted toward either. "
        f"06_Resource_Allocation.xlsx solves that problem explicitly.",

        f"Partner Capacity is the weakest program on every financial measure - "
        f"${P['Partner Capacity']['cost_per_outcome']:,.0f} per placement, the second-lowest "
        f"conversion rate, and the lowest capacity utilisation at "
        f"{P['Partner Capacity']['capacity_utilization']:.0%}. Its case rests on sector leverage "
        f"and on outcomes delivered through partners that this model attributes at a discount. "
        f"That case may well be sound, but it is a strategic argument and should be defended as "
        f"one rather than hidden inside a blended portfolio average.",
    ]
    for n in notes:
        c = ws.cell(row=r, column=1, value=n)
        c.font = Font(name="Calibri", size=9.5, color=S.INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
        ws.row_dimensions[r].height = 13 * (len(n) // 128 + 1)
        r += 1

    S.finish(ws, S.ACCENT)
    return ws


def sheet_calibration(wb):
    ws = wb.create_sheet("Calibration check")
    S.set_widths(ws, {"A": 50, "B": 20, "C": 20, "D": 14, "E": 60})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Calibration against published impact data",
                      "The program mix is synthetic. This is the one place it can be tested "
                      "against something real.", width=5)

    E = "'Program economics'!"
    r = S.header_row(ws, r, ["Measure", "Model", "Published", "Variance", "Note"])
    checks = [
        ("Average salary gain per placement", f"={E}F{R['e.gain']}",
         I.IMPACT_PUBLIC[2025]["avg_salary_gain"],
         "The organization published $58,790 for 2025. The synthetic program mix was tuned so the "
         "blended figure it implies lands within about 1% of that. This is a weak test - one "
         "number, from a different year, that the mix was fitted to - but it is a real constraint "
         "and it is better than none."),
        ("Participants served, FY2024", f"={E}F{R['e.parts']}",
         None,
         "Ties by construction to the count interpolated between the published 2022 and 2025 "
         "observations."),
        ("Successful outcomes, FY2024", f"={E}F{R['e.out']}", None,
         "Ties to the interpolated placement total on the same basis."),
        ("Blended outcome rate", f"={E}F{R['e.outrate']}", None,
         "Implied by the two lines above. The 2022 published figures give 15.9% and the 2025 "
         "figures 13.0%, so a FY2024 rate between those is consistent with the published record."),
    ]
    src_parts = f"={SRC}{A24}{R['src.parts']}"
    src_places = f"={SRC}{A24}{R['src.places']}"
    controls = [I.IMPACT_PUBLIC[2025]["avg_salary_gain"], src_parts, src_places,
                f"={SRC}{A24}{R['src.places']}/{SRC}{A24}{R['src.parts']}"]
    fmts = [S.MONEY, S.NUM1, S.NUM1, S.PCT1]
    for (name, model_f, _pub, note), control, fmt in zip(checks, controls, fmts):
        S.label(ws, r, 1, name, size=10, indent=1)
        ws.cell(row=r, column=2, value=model_f).number_format = fmt
        c = ws.cell(row=r, column=3, value=control)
        c.number_format = fmt
        ws.cell(row=r, column=4, value=f"=IFERROR(B{r}/C{r}-1,\"n/a\")").number_format = S.PCT2
        for col in (2, 3, 4):
            ws.cell(row=r, column=col).alignment = Alignment(horizontal="right")
        n = ws.cell(row=r, column=5, value=note)
        n.font = Font(name="Calibri", size=8.5, color=S.MUTED)
        n.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 13 * (len(note) // 68 + 1)
        r += 1
    r += 1
    S.note(ws, r, "What this check cannot do: it cannot validate the split of participants, costs "
                  "or outcomes across the four programs, because no public data exists at that "
                  "level. Everything on 'Program comparison' should be read as an internally "
                  "consistent illustration of how a portfolio like this behaves, not as a "
                  "measurement of how this organization's programs actually perform.", width=5)
    S.finish(ws, S.MUTED)
    return ws


def sheet_sroi(wb):
    ws = wb.create_sheet("Social return (caveated)")
    S.set_widths(ws, {"A": 48, "B": 16, "C": 16, "D": 16, "E": 16, "F": 16, "G": 3, "H": 54})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Social return on investment - low confidence, reported as a range",
                      "This is the least reliable analysis in the project and is placed last "
                      "deliberately.", width=8)

    warn = ws.cell(row=r, column=1, value=(
        "Read this before quoting any number below. Every parameter driving these ratios - "
        "attribution, deadweight, drop-off, persistence, benefit horizon - is an assumption with no "
        "empirical support in this project. The headline ratio moves by more than a factor of five "
        "across plausible values. Cost per successful outcome, on 'Program economics', rests on far "
        "fewer assumptions and is the measure management should steer on."))
    warn.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    warn.fill = PatternFill("solid", fgColor=S.WARN)
    warn.alignment = Alignment(wrap_text=True, vertical="center")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.row_dimensions[r].height = 56
    r += 2

    r = S.section(ws, r, "Parameters", width=8)
    r = S.header_row(ws, r, ["Parameter", "Value", "", "", "", "", "", "What it does"])
    param_notes = {
        "benefit_horizon_years": "Number of years the earnings gain is counted for. Doubling it "
                                 "roughly doubles the ratio.",
        "discount_rate": "Applied to future years of benefit.",
        "attribution_rate": "Share of the earnings gain credited to the program rather than to "
                            "the participant's own effort, credentials and labour market "
                            "conditions. The single most consequential assumption here.",
        "deadweight_rate": "Share of the gain that would have occurred anyway without the program.",
        "dropoff_rate": "Annual decay in the attributable benefit as the program's influence fades.",
        "persistence_rate": "Share of placements still employed at twelve months.",
    }
    for k, v in I.SROI_PARAMS.items():
        S.label(ws, r, 1, k.replace("_", " ").capitalize(), size=9.5, indent=1)
        S.put(ws, r, 2, v, fmt="#,##0" if "years" in k else S.PCT1, is_input=True)
        n = ws.cell(row=r, column=8, value=param_notes.get(k, ""))
        n.font = Font(name="Calibri", size=8.5, color=S.MUTED)
        n.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 26
        R[f"sroi.{k}"] = r
        r += 1
    r += 1

    sr = engine.sroi()
    r = S.section(ws, r, "Result at the stated parameters", width=8)
    r = S.header_row(ws, r, ["Measure"] + I.PROGRAMS + ["Portfolio", "", ""])
    lines = [
        ("Gross first-year earnings gain", "gross_annual_gain", S.MONEY),
        ("Present value over the horizon", "present_value", S.MONEY),
        ("Attributable value after deadweight", "attributable_value", S.MONEY),
        ("Program cost", "program_cost", S.MONEY),
        ("Social return ratio", "sroi_ratio", S.MULT),
    ]
    for name, field, fmt in lines:
        S.label(ws, r, 1, name, size=10, indent=1)
        for i, p in enumerate(I.PROGRAMS):
            S.put(ws, r, 2 + i, sr[p][field], fmt=fmt)
        if field in sr["_portfolio"]:
            S.put(ws, r, 6, sr["_portfolio"][field], fmt=fmt, bold=True)
        r += 1
    r += 2

    rng = engine.sroi_sensitivity()
    r = S.section(ws, r, "How unstable that number is", width=8)
    r = S.header_row(ws, r, ["Attribution rate"] + [f"{h}-year horizon" for h in (1, 3, 5)]
                     + ["", "", "", ""])
    grid = {}
    for g_ in rng["grid"]:
        grid[(g_["attribution"], g_["horizon"])] = g_["ratio"]
    first_grid = r
    for a in (0.35, 0.45, 0.55, 0.65, 0.75):
        S.label(ws, r, 1, f"{a:.0%}", size=10, indent=1)
        for i, h in enumerate((1, 3, 5)):
            S.put(ws, r, 2 + i, grid[(a, h)], fmt=S.MULT)
        r += 1
    ws.conditional_formatting.add(
        f"B{first_grid}:D{r - 1}",
        ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="F8696B"))
    r += 1

    S.note(ws, r, f"The portfolio ratio ranges from {rng['low']:.2f} to {rng['high']:.2f} across "
                  f"this grid alone - a spread of {rng['spread']:.1f} times - and the grid varies "
                  f"only two of the six parameters. A social return figure quoted without its "
                  f"assumptions is not a finding; it is a decoration. That is why the management "
                  f"report leads on cost per successful outcome instead.", width=8)
    r += 2

    r = S.section(ws, r, "What is deliberately excluded", width=8)
    excl = [
        "Fiscal effects. Higher earnings generate additional tax revenue and reduce transfer "
        "payments. Both are real and both are commonly used to inflate social return figures. "
        "Neither is included here.",
        "Second-order and household effects. Effects on participants' children, on household "
        "stability, on health outcomes. Plausibly large, entirely unmeasured in this project.",
        "Employer-side value. The value to employers of filling skilled roles they could not "
        "otherwise fill. Real, and not estimable from public data.",
        "Earnings gains beyond the stated horizon. Placements plainly produce value past year "
        "three or five; the horizon is a modelling convenience, not a claim that the benefit stops.",
        "Any counterfactual evidence at all. There is no comparison group here. Attribution and "
        "deadweight are assumptions standing in for a study that has not been done, and no amount "
        "of arithmetic can substitute for one.",
    ]
    for e in excl:
        c = ws.cell(row=r, column=1, value="- " + e)
        c.font = Font(name="Calibri", size=9, color=S.INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 13 * (len(e) // 118 + 1)
        r += 1

    S.finish(ws, S.WARN)
    return ws


def build():
    wb = Workbook()
    wb.remove(wb.active)
    source_data_sheet(wb, R)
    sheet_program_inputs(wb)
    sheet_economics(wb)
    sheet_comparison(wb)
    sheet_calibration(wb)
    sheet_sroi(wb)
    wb.move_sheet("Program economics", offset=-3)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"wrote {OUT}")
    return OUT


if __name__ == "__main__":
    build()
