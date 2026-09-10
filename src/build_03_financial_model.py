"""
03_Nonprofit_Financial_Model.xlsx  -  the centre of the project.

Historical analysis of six years of filed accounts, then a driver-based
three-year forecast built on top of the last actual year.

Every calculated cell is a live Excel formula referencing the input sheets. If a
reviewer changes a driver on 'Model inputs', the whole forecast, the ratios, the
cash roll-forward and the dashboard all move. That is the difference between a
model and a report, and it is the point of the exercise.
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import LineChart, BarChart, Reference, Series
from openpyxl.formatting.rule import CellIsRule, DataBarRule
from openpyxl.utils import get_column_letter as gcl
from openpyxl.workbook.defined_name import DefinedName

import inputs as I
import styles as S

OUT = Path(__file__).resolve().parents[1] / "03_Nonprofit_Financial_Model.xlsx"

HY = I.HIST_YEARS                      # 2019..2024
FY = I.FCST_YEARS                      # 2025..2027
HCOL = {y: gcl(2 + i) for i, y in enumerate(HY)}          # B..G
FCOL = {y: gcl(3 + i) for i, y in enumerate(FY)}          # C..E (B holds FY2024A)
BASECOL = "B"

R = {}   # row registry: R['sheet.line'] = row number


# ==========================================================================
# Sheet 1 - Form 990 data
# ==========================================================================

def sheet_990(wb):
    ws = wb.create_sheet("Form 990 data")
    S.set_widths(ws, {"A": 40, **{HCOL[y]: 15 for y in HY}, "H": 11, "I": 62})
    ws.sheet_view.showGridLines = False

    r = S.title_block(ws, "Form 990 data",
                      f"{I.ORG_NAME}, EIN {I.ORG_EIN}. Figures as filed. Blue cells are inputs; "
                      "nothing on this sheet is calculated except the two identity checks.", width=9)

    r = S.header_row(ws, r, ["Line"] + [f"FY{y}" for y in HY] + ["Tier", "Note"])

    def line(name, field, tier="PUBLIC", note="", key=None, bold=False):
        nonlocal r
        S.label(ws, r, 1, name, bold=bold, size=10)
        for y in HY:
            t = tier if not (y == 2024 and field in ("contributions", "program_rev",
                                                     "invest_inc", "total_liab", "payroll_tax")) \
                else "DERIVED"
            S.put(ws, r, ws[HCOL[y] + "1"].column, I.F990[y][field], fmt=S.MONEY, is_input=True)
        S.tier_cell(ws, r, 8, tier)
        c = ws.cell(row=r, column=9, value=note)
        c.font = Font(name="Calibri", size=8.5, color=S.MUTED)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 24
        if key:
            R[key] = r
        r += 1

    r = S.section(ws, r, "Revenue", width=9)
    line("Contributions and grants", "contributions", "PUBLIC",
         "Form 990 Part VIII line 1h. FY2024 derived from the published 91% revenue mix.",
         key="p990.contrib")
    line("Program service revenue", "program_rev", "PUBLIC",
         "Form 990 Part VIII line 2g. FY2024 derived from the published 4% share.",
         key="p990.progrev")
    line("Investment and other income", "invest_inc", "PUBLIC",
         "Form 990 Part VIII line 3 and other. FY2024 is the residual of the derived mix.",
         key="p990.invinc")
    line("Less: fundraising event costs and net asset-sale losses", "revenue_netting", "DERIVED",
         I.REVENUE_NETTING_METHOD[:190], key="p990.netting")
    line("Total revenue", "total_revenue", "PUBLIC", "Form 990 Part I line 12, stated net.",
         key="p990.rev", bold=True)

    # identity check
    S.label(ws, r, 1, "Check: components equal total", italic=True, size=9, color=S.MUTED)
    for y in HY:
        col = HCOL[y]
        f = (f'=IF(ABS({col}{R["p990.contrib"]}+{col}{R["p990.progrev"]}+{col}{R["p990.invinc"]}'
             f'+{col}{R["p990.netting"]}-{col}{R["p990.rev"]})<1,"ok","CHECK")')
        c = ws.cell(row=r, column=ws[col + "1"].column, value=f)
        c.font = Font(name="Calibri", size=8.5, italic=True, color=S.ACCENT)
        c.alignment = Alignment(horizontal="right")
    R["p990.revcheck"] = r
    r += 1
    S.note(ws, r, "Form 990 reports gross contributions, gross program service revenue and gross "
                  "investment income, but states total revenue net of the direct expenses of "
                  "fundraising events and of losses on asset sales. The three gross lines therefore "
                  "sum to more than reported total revenue in FY2019 through FY2023. Carrying that "
                  "difference on its own line is what makes the analysis tie to the filing instead "
                  "of quietly overstating revenue by between $34,000 and $135,000 a year.", width=9)
    r += 1

    r = S.section(ws, r, "Expenses and personnel", width=9)
    line("Officer compensation", "officer_comp", "PUBLIC", "Form 990 Part IX line 5.",
         key="p990.offcomp")
    line("Other salaries and wages", "other_salaries", "PUBLIC", "Form 990 Part IX line 7.",
         key="p990.othersal")
    line("Payroll taxes", "payroll_tax", "PUBLIC",
         "Form 990 Part IX line 10. FY2024 derived at the FY2023 effective rate of 7.82%.",
         key="p990.ptax")
    line("Total functional expenses", "total_expense", "PUBLIC", "Form 990 Part I line 18.",
         key="p990.exp", bold=True)
    r += 1

    r = S.section(ws, r, "Balance sheet", width=9)
    line("Total assets", "total_assets", "PUBLIC", "Form 990 Part I line 20, end of year.",
         key="p990.assets")
    line("Total liabilities", "total_liab", "PUBLIC",
         "Form 990 Part I line 21. FY2024 derived as assets less published net assets.",
         key="p990.liab")

    S.label(ws, r, 1, "Net assets", bold=True, size=10)
    for y in HY:
        col = HCOL[y]
        c = ws.cell(row=r, column=ws[col + "1"].column,
                    value=f"={col}{R['p990.assets']}-{col}{R['p990.liab']}")
        c.number_format = S.MONEY
        c.font = Font(name="Calibri", size=10, bold=True, color=S.INK)
        c.alignment = Alignment(horizontal="right")
    S.tier_cell(ws, r, 8, "DERIVED")
    ws.cell(row=r, column=9, value="Assets less liabilities. FY2024 ties to the published "
                                   f"net asset figure of ${I.NET_ASSETS_2024_PUBLIC:,}.").font = \
        Font(name="Calibri", size=8.5, color=S.MUTED)
    R["p990.netassets"] = r
    r += 1

    S.label(ws, r, 1, "Check: FY2024 ties to filed net assets", italic=True, size=9, color=S.MUTED)
    c = ws.cell(row=r, column=7,
                value=f'=IF(ABS(G{R["p990.netassets"]}-{I.NET_ASSETS_2024_PUBLIC})<1,"ok","CHECK")')
    c.font = Font(name="Calibri", size=8.5, italic=True, color=S.ACCENT)
    c.alignment = Alignment(horizontal="right")
    r += 2

    r = S.section(ws, r, "Published impact metrics", width=9)
    S.label(ws, r, 1, "Participants served", size=10)
    for y in HY:
        col = HCOL[y]
        if y == 2022:
            S.put(ws, r, ws[col + "1"].column, I.IMPACT_PUBLIC[2022]["participants"],
                  fmt=S.NUM, is_input=True)
        else:
            # constant-growth interpolation anchored on the two published observations
            S.put(ws, r, ws[col + "1"].column,
                  f'=$B${r + 4}*(1+$B${r + 6})^({y}-2022)', fmt=S.NUM1)
    S.tier_cell(ws, r, 8, "DERIVED")
    ws.cell(row=r, column=9, value="Published for 2022 and 2025 only; other years interpolated at "
                                   "constant growth between those two observations.").font = \
        Font(name="Calibri", size=8.5, color=S.MUTED)
    R["p990.parts"] = r
    r += 1

    S.label(ws, r, 1, "Successful outcomes (placements)", size=10)
    for y in HY:
        col = HCOL[y]
        if y == 2022:
            S.put(ws, r, ws[col + "1"].column, I.IMPACT_PUBLIC[2022]["placements"],
                  fmt=S.NUM, is_input=True)
        else:
            S.put(ws, r, ws[col + "1"].column,
                  f'=$B${r + 4}*(1+$B${r + 6})^({y}-2022)', fmt=S.NUM1)
    S.tier_cell(ws, r, 8, "DERIVED")
    R["p990.places"] = r
    r += 2

    r = S.section(ws, r, "Interpolation anchors", width=9)
    anchors = [
        ("2022 participants (published)", I.IMPACT_PUBLIC[2022]["participants"], S.NUM),
        ("2025 participants (published)", I.IMPACT_PUBLIC[2025]["participants"], S.NUM),
        ("Participant growth per year", None, S.PCT1),
        ("2022 placements (published)", I.IMPACT_PUBLIC[2022]["placements"], S.NUM),
        ("2025 placements (published)", I.IMPACT_PUBLIC[2025]["placements"], S.NUM),
        ("Placement growth per year", None, S.PCT1),
        ("2025 average starting salary", I.IMPACT_PUBLIC[2025]["avg_starting_salary"], S.MONEY),
        ("2025 average salary gain", I.IMPACT_PUBLIC[2025]["avg_salary_gain"], S.MONEY),
    ]
    anchor_start = r
    for i, (name, val, fmt) in enumerate(anchors):
        S.label(ws, r, 1, name, size=9.5, indent=1)
        if val is None:
            base = anchor_start + (0 if i == 2 else 3)
            S.put(ws, r, 2, f"=(B{base + 1}/B{base})^(1/3)-1", fmt=fmt)
        else:
            S.put(ws, r, 2, val, fmt=fmt, is_input=True)
        S.tier_cell(ws, r, 8, "PUBLIC" if val is not None else "DERIVED")
        r += 1

    # rewire the interpolation formulas now that anchor rows are known
    p_anchor, p_growth = anchor_start, anchor_start + 2
    l_anchor, l_growth = anchor_start + 3, anchor_start + 5
    for y in HY:
        col = HCOL[y]
        if y != 2022:
            ws[f"{col}{R['p990.parts']}"] = f"=$B${p_anchor}*(1+$B${p_growth})^({y}-2022)"
            ws[f"{col}{R['p990.places']}"] = f"=$B${l_anchor}*(1+$B${l_growth})^({y}-2022)"
    R["p990.anchor_gain"] = anchor_start + 7

    r += 1
    S.note(ws, r, "The interpolation has no free parameters: it is fixed by two published observations "
                  "and the assumption of constant growth between them. The back-cast to FY2019 through "
                  "FY2021 is a different matter and is the weakest link in the participant series, "
                  "because it assumes the same growth rate held through a period that includes the "
                  "pandemic and the start of the humanitarian response surge.", width=9)

    S.finish(ws, S.NAVY)
    return ws


# ==========================================================================
# Sheet 2 - Model inputs
# ==========================================================================

def sheet_inputs(wb):
    ws = wb.create_sheet("Model inputs")
    S.set_widths(ws, {"A": 42, "B": 15, "C": 15, "D": 15, "E": 15, "F": 15, "G": 15,
                      "H": 11, "I": 62})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Model inputs",
                      "Every driver the historical analysis and the forecast run on. Blue cells are "
                      "inputs you can change; the rest of the workbook recalculates.", width=9)

    def block(title, rows_spec, headers=None):
        nonlocal r
        r = S.section(ws, r, title, width=9)
        if headers:
            r = S.header_row(ws, r, headers)
        for spec in rows_spec:
            name, values, fmt, tier, note, key = spec
            S.label(ws, r, 1, name, size=9.5, indent=1)
            if isinstance(values, (list, tuple)):
                for i, v in enumerate(values):
                    S.put(ws, r, 2 + i, v, fmt=fmt, is_input=True)
            else:
                S.put(ws, r, 2, values, fmt=fmt, is_input=True)
            S.tier_cell(ws, r, 8, tier)
            c = ws.cell(row=r, column=9, value=note)
            c.font = Font(name="Calibri", size=8.5, color=S.MUTED)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 22
            if key:
                R[key] = r
            r += 1
        r += 1

    # --- functional split ------------------------------------------------
    block("Functional expense allocation", [
        ("Program services share", [I.FUNCTIONAL_SPLIT[y][0] for y in HY], S.PCT1, "SYNTHETIC",
         I.FUNCTIONAL_SPLIT_METHOD[:150], "in.fs_prog"),
        ("Management and general share", [I.FUNCTIONAL_SPLIT[y][1] for y in HY], S.PCT1, "SYNTHETIC",
         "Semi-fixed; grows more slowly than programs as the organization scales.", "in.fs_mg"),
        ("Fundraising share", [I.FUNCTIONAL_SPLIT[y][2] for y in HY], S.PCT1, "SYNTHETIC",
         "Rises slightly in FY2024 as the organization invests in replacing surge funding.",
         "in.fs_fr"),
    ], headers=["Driver"] + [f"FY{y}" for y in HY] + ["Tier", "Note"])

    # --- liquidity -------------------------------------------------------
    block("Liquidity composition", [
        ("Cash as a share of total assets", [I.CASH_SHARE_OF_ASSETS[y] for y in HY], S.PCT1,
         "SYNTHETIC", "Strict cash measure.", "in.cashshare"),
        ("Liquid reserves as a share of total assets", [I.LIQUID_SHARE_OF_ASSETS[y] for y in HY],
         S.PCT1, "SYNTHETIC", "Cash plus short-term investments - the measure a board treats as "
                              "available. Both are reported everywhere; they must never be mixed.",
         "in.liqshare"),
    ], headers=["Driver"] + [f"FY{y}" for y in HY] + ["Tier", "Note"])

    # --- revenue mix -----------------------------------------------------
    spec = [(f"{c}", I.REVENUE_MIX[c], S.PCT1, "SYNTHETIC", I.REVENUE_MIX_METHOD[:130],
             f"in.mix_{c}") for c in I.REVENUE_CATEGORIES]
    block("Revenue composition (share of contributions)", spec,
          headers=["Category", "Share", "", "", "", "", "", "Tier", "Note"])
    R["in.mix_check"] = r - 1

    # --- donor drivers ---------------------------------------------------
    r = S.section(ws, r, "Revenue drivers: donor count and average gift", width=9)
    r = S.header_row(ws, r, ["Category", "FY2024 donors", "Retention", "FY2024 revenue",
                             "Average gift", "", "", "Tier", "Note"])
    donor_first = r
    for c in I.REVENUE_CATEGORIES:
        S.label(ws, r, 1, c, size=9.5, indent=1)
        S.put(ws, r, 2, I.DONOR_DRIVERS_2024[c]["donors"], fmt=S.NUM, is_input=True)
        S.put(ws, r, 3, I.DONOR_DRIVERS_2024[c]["retention"], fmt=S.PCT1, is_input=True)
        mix_row = R[f"in.mix_{c}"]
        S.put(ws, r, 4, f"='Form 990 data'!G{R['p990.contrib']}*B{mix_row}", fmt=S.MONEY)
        S.put(ws, r, 5, f"=D{r}/B{r}", fmt=S.MONEY2)
        S.tier_cell(ws, r, 8, "SYNTHETIC")
        ws.cell(row=r, column=9, value=I.DONOR_DRIVER_METHOD[:120]).font = \
            Font(name="Calibri", size=8.5, color=S.MUTED)
        R[f"in.donor_{c}"] = r
        r += 1
    S.label(ws, r, 1, "Total", bold=True, indent=1)
    ws.cell(row=r, column=4, value=f"=SUM(D{donor_first}:D{r - 1})").number_format = S.MONEY
    S.total_row_style(ws, r, [1, 2, 3, 4, 5])
    R["in.donor_total"] = r
    r += 1
    S.label(ws, r, 1, "Check: ties to filed contributions", italic=True, size=9, color=S.MUTED)
    c = ws.cell(row=r, column=4,
                value=f"=IF(ABS(D{R['in.donor_total']}-'Form 990 data'!G{R['p990.contrib']})<1,\"ok\",\"CHECK\")")
    c.font = Font(name="Calibri", size=8.5, italic=True, color=S.ACCENT)
    c.alignment = Alignment(horizontal="right")
    r += 2

    # --- program portfolio ----------------------------------------------
    r = S.section(ws, r, "Program portfolio, FY2024 base year", width=9)
    r = S.header_row(ws, r, ["Program", "Participants", "Outcomes", "Program cost",
                             "Completion rate", "Salary gain", "Staff hrs / participant",
                             "Tier", "Note"])
    prog_first = r
    for p in I.PROGRAMS:
        d = I.PROGRAM_BASE[p]
        S.label(ws, r, 1, p, size=9.5, indent=1)
        S.put(ws, r, 2, d["participants"], fmt=S.NUM, is_input=True)
        S.put(ws, r, 3, d["placements"], fmt=S.NUM, is_input=True)
        S.put(ws, r, 4, d["program_cost"], fmt=S.MONEY, is_input=True)
        S.put(ws, r, 5, d["completion_rate"], fmt=S.PCT1, is_input=True)
        S.put(ws, r, 6, d["salary_gain"], fmt=S.MONEY, is_input=True)
        S.put(ws, r, 7, d["staff_hours_per_participant"], fmt=S.NUM1, is_input=True)
        S.tier_cell(ws, r, 8, "SYNTHETIC")
        ws.cell(row=r, column=9, value=I.PROGRAM_DESCRIPTION[p][:110]).font = \
            Font(name="Calibri", size=8.5, color=S.MUTED)
        ws.row_dimensions[r].height = 24
        R[f"in.prog_{p}"] = r
        r += 1
    S.label(ws, r, 1, "Total", bold=True, indent=1)
    for col in (2, 3, 4, 7):
        L = gcl(col)
        ws.cell(row=r, column=col,
                value=f"=SUM({L}{prog_first}:{L}{r - 1})").number_format = \
            S.NUM if col in (2, 3) else (S.MONEY if col == 4 else S.NUM1)
    S.total_row_style(ws, r, range(1, 8))
    R["in.prog_total"] = r
    r += 1
    S.label(ws, r, 1, "Check: costs tie to the program expense pool", italic=True, size=9,
            color=S.MUTED)
    c = ws.cell(row=r, column=4, value=(
        f"=IF(ABS(D{R['in.prog_total']}-'Form 990 data'!G{R['p990.exp']}*G{R['in.fs_prog']})<1,"
        f"\"ok\",\"CHECK\")"))
    c.font = Font(name="Calibri", size=8.5, italic=True, color=S.ACCENT)
    c.alignment = Alignment(horizontal="right")
    r += 2

    # --- personnel and cost behaviour ------------------------------------
    block("Personnel", [
        ("Full-time equivalents, FY2024", I.FTE_2024, S.NUM, "SYNTHETIC",
         I.PERSONNEL_METHOD[:140], "in.fte"),
        ("Payroll tax rate", I.PAYROLL_TAX_RATE, S.PCT2, "DERIVED",
         "FY2023 payroll tax divided by FY2023 salaries, both filed figures.", "in.ptaxrate"),
        ("Benefits load", I.BENEFITS_RATE, S.PCT1, "SYNTHETIC",
         "Health, retirement and other benefits as a share of salary.", "in.benefits"),
        ("Productive hours per FTE", I.STAFF_HOURS_PER_FTE, S.NUM, "SYNTHETIC",
         "2,080 gross hours less leave, training and non-delivery time.", "in.hoursfte"),
        ("Share of FTE in program delivery", I.PROGRAM_FTE_SHARE, S.PCT1, "SYNTHETIC",
         "Balance sits in administration and fundraising.", "in.progfte"),
    ])

    block("Cost behaviour", [
        ("Average gift inflation", I.AVG_GIFT_INFLATION, S.PCT1, "SYNTHETIC",
         I.AVG_GIFT_INFLATION_NOTE[:140], "in.giftinfl"),
        ("Management and general, variable share", I.MG_VARIABLE_SHARE, S.PCT1, "SYNTHETIC",
         I.SEMI_VARIABLE_NOTE[:140], "in.mgvar"),
        ("Fundraising, variable share", I.FR_VARIABLE_SHARE, S.PCT1, "SYNTHETIC",
         "Scales with contribution volume; the rest is standing fundraising infrastructure.",
         "in.frvar"),
    ])

    block("Base case forecast drivers", [
        (I.SCENARIO_DRIVER_LABELS[k], I.SCENARIO_DRIVERS[k][1], S.PCT1, "SYNTHETIC",
         I.SCENARIO_NARRATIVE["Base"][:130], f"in.drv_{k}")
        for k in I.SCENARIO_DRIVERS
    ])

    block("Benchmarks and policy limits", [
        ("Minimum program expense ratio (BBB Standard 8)", 0.65, S.PCT1, "BENCHMARK",
         "BBB Wise Giving Alliance: at least 65% of total expenses on programs.", "in.bm_prog"),
        ("Maximum fundraising cost ratio (BBB Standard 9)", 0.35, S.PCT1, "BENCHMARK",
         "No more than 35 cents spent to raise a dollar.", "in.bm_fr"),
        ("Maximum net assets, multiple of expenses (BBB Standard 10)", 3.00, S.MULT, "BENCHMARK",
         "Unrestricted net assets available for use should not exceed three times annual expenses. "
         "This is the binding constraint at the top of the range for this organization.", "in.bm_na"),
        ("Operating reserve target, low", 3.0, S.MONTHS, "BENCHMARK",
         "Propel Nonprofits: a commonly used goal is three to six months.", "in.bm_res_lo"),
        ("Operating reserve target, high", 6.0, S.MONTHS, "BENCHMARK",
         "Also the reserve floor used in the sensitivity and allocation models.", "in.bm_res_hi"),
        ("Reserve ceiling", 24.0, S.MONTHS, "BENCHMARK",
         "Propel Nonprofits: reserves should not exceed two years of budget.", "in.bm_res_ceil"),
    ])

    S.finish(ws, S.ACCENT)
    return ws


# ==========================================================================
# Sheet 3 - Historical analysis
# ==========================================================================

def sheet_historical(wb):
    ws = wb.create_sheet("Historical analysis")
    S.set_widths(ws, {"A": 44, **{HCOL[y]: 15 for y in HY}, "H": 13, "I": 60})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Historical financial analysis, FY2019 - FY2024",
                      "All figures calculated from the Form 990 data and model input sheets. "
                      "Nothing on this sheet is typed in.", width=9)

    D = "'Form 990 data'!"
    M = "'Model inputs'!"

    def row(name, formula_fn, fmt=S.MONEY, bold=False, indent=1, key=None, note="",
            first_year_blank=False, tier=None):
        nonlocal r
        S.label(ws, r, 1, name, bold=bold, size=10, indent=indent)
        for i, y in enumerate(HY):
            col = HCOL[y]
            if first_year_blank and i == 0:
                c = ws.cell(row=r, column=ws[col + "1"].column, value="n/a")
                c.font = Font(name="Calibri", size=9, italic=True, color=S.MUTED)
                c.alignment = Alignment(horizontal="right")
                continue
            f = formula_fn(y, col, i)
            c = ws.cell(row=r, column=ws[col + "1"].column, value=f)
            c.number_format = fmt
            c.font = Font(name="Calibri", size=10, bold=bold, color=S.INK)
            c.alignment = Alignment(horizontal="right")
        if tier:
            S.tier_cell(ws, r, 8, tier)
        if note:
            c = ws.cell(row=r, column=9, value=note)
            c.font = Font(name="Calibri", size=8.5, color=S.MUTED)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 24
        if key:
            R[key] = r
        r += 1

    r = S.header_row(ws, r, ["Line"] + [f"FY{y}" for y in HY] + ["", "Note"])

    # --- P&L -------------------------------------------------------------
    r = S.section(ws, r, "Statement of activities", width=9)
    row("Contributions and grants", lambda y, c, i: f"={D}{c}{R['p990.contrib']}", key="h.contrib")
    row("Program service revenue", lambda y, c, i: f"={D}{c}{R['p990.progrev']}", key="h.progrev")
    row("Investment and other income", lambda y, c, i: f"={D}{c}{R['p990.invinc']}", key="h.invinc")
    row("Less: fundraising event costs and asset-sale losses",
        lambda y, c, i: f"={D}{c}{R['p990.netting']}", key="h.netting", tier="DERIVED",
        note="Form 990 states total revenue net of these. Without this line the revenue total "
             "would not tie to the filing.")
    row("Total revenue", lambda y, c, i: f"=SUM({c}{R['h.contrib']}:{c}{R['h.netting']})",
        bold=True, indent=0, key="h.rev")
    row("Check: ties to filed total revenue",
        lambda y, c, i: f'=IF(ABS({c}{R["h.rev"]}-{D}{c}{R["p990.rev"]})<1,"ok","CHECK")',
        fmt="@", indent=2)
    r += 1
    row("Program services", lambda y, c, i: f"={D}{c}{R['p990.exp']}*{M}{c}{R['in.fs_prog']}",
        key="h.prog", tier="SYNTHETIC",
        note="Filed total expense multiplied by the modelled program share.")
    row("Management and general", lambda y, c, i: f"={D}{c}{R['p990.exp']}*{M}{c}{R['in.fs_mg']}",
        key="h.mg", tier="SYNTHETIC")
    row("Fundraising", lambda y, c, i: f"={D}{c}{R['p990.exp']}*{M}{c}{R['in.fs_fr']}",
        key="h.fr", tier="SYNTHETIC")
    row("Total expenses", lambda y, c, i: f"=SUM({c}{R['h.prog']}:{c}{R['h.fr']})",
        bold=True, indent=0, key="h.exp")
    row("Check: ties to filed total expenses",
        lambda y, c, i: f'=IF(ABS({c}{R["h.exp"]}-{D}{c}{R["p990.exp"]})<1,"ok","CHECK")',
        fmt="@", indent=2)
    r += 1
    row("Operating surplus / (deficit)", lambda y, c, i: f"={c}{R['h.rev']}-{c}{R['h.exp']}",
        bold=True, indent=0, key="h.result")
    row("Operating margin", lambda y, c, i: f"={c}{R['h.result']}/{c}{R['h.rev']}",
        fmt=S.PCT1, key="h.margin")
    r += 1

    # --- natural classification ------------------------------------------
    r = S.section(ws, r, "Natural classification (memo)", width=9)
    row("Salaries and officer compensation",
        lambda y, c, i: f"={D}{c}{R['p990.offcomp']}+{D}{c}{R['p990.othersal']}", key="h.sal")
    row("Payroll taxes", lambda y, c, i: f"={D}{c}{R['p990.ptax']}", key="h.ptax")
    row("Benefits", lambda y, c, i: f"={c}{R['h.sal']}*{M}$B${R['in.benefits']}", key="h.ben",
        tier="SYNTHETIC", note="Assumed load; not separately reported in the extract used here.")
    row("Total personnel", lambda y, c, i: f"=SUM({c}{R['h.sal']}:{c}{R['h.ben']})",
        bold=True, indent=0, key="h.personnel")
    row("Non-personnel expenses", lambda y, c, i: f"={c}{R['h.exp']}-{c}{R['h.personnel']}",
        key="h.nonpersonnel")
    row("Personnel as a share of expenses",
        lambda y, c, i: f"={c}{R['h.personnel']}/{c}{R['h.exp']}", fmt=S.PCT1, key="h.persratio")
    r += 1

    # --- balance sheet ---------------------------------------------------
    r = S.section(ws, r, "Financial position", width=9)
    row("Total assets", lambda y, c, i: f"={D}{c}{R['p990.assets']}", key="h.assets")
    row("Total liabilities", lambda y, c, i: f"={D}{c}{R['p990.liab']}", key="h.liab")
    row("Net assets", lambda y, c, i: f"={c}{R['h.assets']}-{c}{R['h.liab']}", bold=True,
        indent=0, key="h.na")
    row("Cash and equivalents",
        lambda y, c, i: f"={c}{R['h.assets']}*{M}{c}{R['in.cashshare']}", key="h.cash",
        tier="SYNTHETIC")
    row("Liquid reserves (cash plus short-term investments)",
        lambda y, c, i: f"={c}{R['h.assets']}*{M}{c}{R['in.liqshare']}", key="h.liquid",
        tier="SYNTHETIC",
        note="The measure a board treats as available. Roughly a year more runway than cash alone.")
    r += 1

    # --- growth ----------------------------------------------------------
    r = S.section(ws, r, "Growth", width=9)
    prev = lambda c: gcl(ws[c + "1"].column - 1)
    row("Revenue growth",
        lambda y, c, i: f"={c}{R['h.rev']}/{prev(c)}{R['h.rev']}-1", fmt=S.PCT1,
        first_year_blank=True, key="h.revg")
    row("Expense growth",
        lambda y, c, i: f"={c}{R['h.exp']}/{prev(c)}{R['h.exp']}-1", fmt=S.PCT1,
        first_year_blank=True, key="h.expg")
    row("Contribution growth",
        lambda y, c, i: f"={c}{R['h.contrib']}/{prev(c)}{R['h.contrib']}-1", fmt=S.PCT1,
        first_year_blank=True, key="h.cong")
    row("Growth gap (revenue less expense)",
        lambda y, c, i: f"={c}{R['h.revg']}-{c}{R['h.expg']}", fmt=S.PCT1,
        first_year_blank=True, key="h.gapg",
        note="Positive means revenue outran spending - the pattern that built the reserve.")
    r += 1

    # --- ratios ----------------------------------------------------------
    r = S.section(ws, r, "Efficiency and stewardship ratios", width=9)
    row("Program expense ratio", lambda y, c, i: f"={c}{R['h.prog']}/{c}{R['h.exp']}",
        fmt=S.PCT1, key="h.progratio",
        note="BBB Standard 8 requires at least 65%.")
    row("  vs BBB minimum (65%)",
        lambda y, c, i: f"={c}{R['h.progratio']}-{M}$B${R['in.bm_prog']}", fmt=S.PCT1, indent=2,
        key="h.progratio_gap")
    row("Administrative cost ratio", lambda y, c, i: f"={c}{R['h.mg']}/{c}{R['h.exp']}",
        fmt=S.PCT1, key="h.adminratio")
    row("Fundraising cost ratio (cost to raise $1)",
        lambda y, c, i: f"={c}{R['h.fr']}/{c}{R['h.contrib']}", fmt=S.PCT2, key="h.frratio",
        note="BBB Standard 9 caps this at 35 cents. Anything under 10 cents is strong.")
    row("Fundraising return (contributions per $1 spent)",
        lambda y, c, i: f"={c}{R['h.contrib']}/{c}{R['h.fr']}", fmt=S.MULT, key="h.frroi")
    r += 1

    # --- liquidity -------------------------------------------------------
    r = S.section(ws, r, "Liquidity and reserves", width=9)
    row("Average monthly expenses", lambda y, c, i: f"={c}{R['h.exp']}/12", key="h.monthly")
    row("Cash runway", lambda y, c, i: f"={c}{R['h.cash']}/{c}{R['h.monthly']}",
        fmt=S.MONTHS, key="h.cashrunway",
        note="Strict measure: cash and equivalents only.")
    row("Liquid reserve runway",
        lambda y, c, i: f"={c}{R['h.liquid']}/{c}{R['h.monthly']}", fmt=S.MONTHS, key="h.liqrunway",
        note="Cash plus short-term investments. This is the number the board would use.")
    row("Net assets, months of expenses",
        lambda y, c, i: f"={c}{R['h.na']}/{c}{R['h.monthly']}", fmt=S.MONTHS, key="h.namonths")
    row("Net assets, multiple of annual expenses",
        lambda y, c, i: f"={c}{R['h.na']}/{c}{R['h.exp']}", fmt=S.MULT, key="h.namult",
        note="BBB Standard 10 caps this at three times. Watch the trend, not just the level.")
    row("  headroom to the BBB 3x limit",
        lambda y, c, i: f"={M}$B${R['in.bm_na']}-{c}{R['h.namult']}", fmt=S.MULT, indent=2,
        key="h.naheadroom")
    r += 1

    # --- program economics ------------------------------------------------
    r = S.section(ws, r, "Program economics", width=9)
    row("Participants served", lambda y, c, i: f"={D}{c}{R['p990.parts']}", fmt=S.NUM1,
        key="h.parts", tier="DERIVED")
    row("Successful outcomes", lambda y, c, i: f"={D}{c}{R['p990.places']}", fmt=S.NUM1,
        key="h.places", tier="DERIVED")
    row("Outcome rate", lambda y, c, i: f"={c}{R['h.places']}/{c}{R['h.parts']}",
        fmt=S.PCT1, key="h.outrate")
    row("Cost per participant", lambda y, c, i: f"={c}{R['h.prog']}/{c}{R['h.parts']}",
        fmt=S.MONEY, key="h.cpp")
    row("Cost per successful outcome", lambda y, c, i: f"={c}{R['h.prog']}/{c}{R['h.places']}",
        fmt=S.MONEY, key="h.cpo")
    row("Total revenue per participant",
        lambda y, c, i: f"={c}{R['h.rev']}/{c}{R['h.parts']}", fmt=S.MONEY, key="h.revpp",
        note="Rising faster than cost per participant is the surge in one line.")
    r += 2

    # --- summary ----------------------------------------------------------
    r = S.section(ws, r, "Six-year summary", width=9)
    n = len(HY) - 1
    summ = [
        ("Revenue CAGR, FY2019 - FY2024", f"=(G{R['h.rev']}/B{R['h.rev']})^(1/{n})-1", S.PCT1),
        ("Expense CAGR, FY2019 - FY2024", f"=(G{R['h.exp']}/B{R['h.exp']})^(1/{n})-1", S.PCT1),
        ("Contribution CAGR, FY2019 - FY2024",
         f"=(G{R['h.contrib']}/B{R['h.contrib']})^(1/{n})-1", S.PCT1),
        ("Net asset CAGR, FY2019 - FY2024", f"=(G{R['h.na']}/B{R['h.na']})^(1/{n})-1", S.PCT1),
        ("Cumulative operating surplus", f"=SUM(B{R['h.result']}:G{R['h.result']})", S.MONEY),
        ("Years in surplus", f"=COUNTIF(B{R['h.result']}:G{R['h.result']},\">0\")", "0"),
        ("Net assets added over the period", f"=G{R['h.na']}-B{R['h.na']}", S.MONEY),
        ("Participants added over the period", f"=G{R['h.parts']}-B{R['h.parts']}", S.NUM1),
    ]
    for name, f, fmt in summ:
        S.label(ws, r, 1, name, size=10, indent=1)
        c = ws.cell(row=r, column=2, value=f)
        c.number_format = fmt
        c.font = Font(name="Calibri", size=10, bold=True, color=S.NAVY)
        c.alignment = Alignment(horizontal="right")
        r += 1

    # conditional formatting on the ratio rows
    ws.conditional_formatting.add(
        f"B{R['h.progratio']}:G{R['h.progratio']}",
        CellIsRule(operator="lessThan", formula=[f"'Model inputs'!$B${R['in.bm_prog']}"],
                   fill=PatternFill("solid", bgColor="FFC7CE")))
    ws.conditional_formatting.add(
        f"B{R['h.namult']}:G{R['h.namult']}",
        CellIsRule(operator="greaterThan", formula=[f"'Model inputs'!$B${R['in.bm_na']}"],
                   fill=PatternFill("solid", bgColor="FFC7CE")))
    ws.conditional_formatting.add(
        f"B{R['h.result']}:G{R['h.result']}",
        CellIsRule(operator="lessThan", formula=["0"],
                   fill=PatternFill("solid", bgColor="FFC7CE")))

    S.freeze(ws, "B5")
    S.finish(ws, S.NAVY)
    return ws


# ==========================================================================
# Sheet 4 - Forecast
# ==========================================================================

def sheet_forecast(wb):
    ws = wb.create_sheet("Forecast")
    S.set_widths(ws, {"A": 44, "B": 16, "C": 16, "D": 16, "E": 16, "F": 4, "G": 13, "H": 60})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Three-year forecast, FY2025 - FY2027 (Base case)",
                      "Driver based: donation revenue is donor count times average gift, and program "
                      "cost is participants times cost per participant. Change a driver on "
                      "'Model inputs' and every number here moves. Scenarios live in "
                      "05_Scenario_Model.xlsx.", width=8)

    M = "'Model inputs'!"
    H = "'Historical analysis'!"
    cols = [BASECOL] + [FCOL[y] for y in FY]

    def row(name, base_val, fcst_fn, fmt=S.MONEY, bold=False, indent=1, key=None, note="",
            tier=None):
        nonlocal r
        S.label(ws, r, 1, name, bold=bold, size=10, indent=indent)
        if base_val is not None:
            c = ws.cell(row=r, column=2, value=base_val)
            c.number_format = fmt
            c.font = Font(name="Calibri", size=10, bold=bold, italic=True, color=S.MUTED)
            c.alignment = Alignment(horizontal="right")
        for i, y in enumerate(FY):
            col = FCOL[y]
            pcol = cols[i]
            c = ws.cell(row=r, column=ws[col + "1"].column, value=fcst_fn(y, col, pcol, i))
            c.number_format = fmt
            c.font = Font(name="Calibri", size=10, bold=bold, color=S.INK)
            c.alignment = Alignment(horizontal="right")
        if tier:
            S.tier_cell(ws, r, 7, tier)
        if note:
            c = ws.cell(row=r, column=8, value=note)
            c.font = Font(name="Calibri", size=8.5, color=S.MUTED)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 24
        if key:
            R[key] = r
        r += 1

    r = S.header_row(ws, r, ["Line", "FY2024 A", "FY2025 E", "FY2026 F", "FY2027 F", "", "", "Note"])

    # ---- revenue drivers -------------------------------------------------
    r = S.section(ws, r, "Revenue drivers", width=8)
    for cat in I.REVENUE_CATEGORIES:
        gk = {"Individual donations": "growth_individual",
              "Corporate donations": "growth_corporate",
              "Foundation grants": "growth_foundation",
              "Government funding": "growth_government",
              "Fundraising events": "growth_events"}[cat]
        drow = R[f"in.drv_{gk}"]
        dn_row = R[f"in.donor_{cat}"]
        row(f"{cat}: donors", f"={M}B{dn_row}",
            lambda y, c, p, i, dr=drow: f"={p}{r}*(1+{M}$B${dr})/(1+{M}$B${R['in.giftinfl']})",
            fmt=S.NUM1, indent=2, key=f"f.don_{cat}")
        row(f"{cat}: average gift", f"={M}E{dn_row}",
            lambda y, c, p, i: f"={p}{r}*(1+{M}$B${R['in.giftinfl']})",
            fmt=S.MONEY2, indent=2, key=f"f.gift_{cat}")
    r += 1

    # ---- revenue ---------------------------------------------------------
    r = S.section(ws, r, "Revenue", width=8)
    first_cat = r
    for cat in I.REVENUE_CATEGORIES:
        dr, gr = R[f"f.don_{cat}"], R[f"f.gift_{cat}"]
        row(cat, f"={M}D{R[f'in.donor_{cat}']}",
            lambda y, c, p, i, dr=dr, gr=gr: f"={c}{dr}*{c}{gr}",
            key=f"f.rev_{cat}",
            note="Donor count times average gift." if cat == I.REVENUE_CATEGORIES[0] else "")
    row("Total contributions and grants", f"={H}G{R['h.contrib']}",
        lambda y, c, p, i: f"=SUM({c}{first_cat}:{c}{first_cat + 4})",
        bold=True, indent=0, key="f.contrib")
    row("Program service revenue", f"={H}G{R['h.progrev']}",
        lambda y, c, p, i: f"={p}{r}*(1+{M}$B${R['in.drv_growth_program_revenue']})",
        key="f.progrev")
    row("Investment and other income", f"={H}G{R['h.invinc']}",
        lambda y, c, p, i: f"={p}{R['f.liq_open'] if 'f.liq_open' in R else 0}*"
                           f"{M}$B${R['in.drv_investment_return']}",
        key="f.invinc",
        note="Return on the opening liquid reserve balance. This line grows as the reserve grows, "
             "which flatters the operating result and is worth separating out.")
    row("Total revenue", f"={H}G{R['h.rev']}",
        lambda y, c, p, i: f"={c}{R['f.contrib']}+{c}{R['f.progrev']}+{c}{R['f.invinc']}",
        bold=True, indent=0, key="f.rev")
    r += 1

    # ---- program volume and cost ----------------------------------------
    r = S.section(ws, r, "Program volume and unit cost", width=8)
    for p_ in I.PROGRAMS:
        prow = R[f"in.prog_{p_}"]
        row(f"{p_}: participants", f"={M}B{prow}",
            lambda y, c, pc, i, pr=prow: (
                f"=MIN({pc}{r}*(1+{M}$B${R['in.drv_growth_participants']}),"
                f"{I.PROGRAM_BASE[[k for k in I.PROGRAMS][0]]['max_participants']})"),
            fmt=S.NUM1, indent=2, key=f"f.parts_{p_}")
    r += 1
    for p_ in I.PROGRAMS:
        prow = R[f"in.prog_{p_}"]
        row(f"{p_}: cost per participant", f"={M}D{prow}/{M}B{prow}",
            lambda y, c, pc, i: f"={pc}{r}*(1+{M}$B${R['in.drv_inflation_cpp']})",
            fmt=S.MONEY2, indent=2, key=f"f.cpp_{p_}")
    r += 1
    first_cost = r
    for p_ in I.PROGRAMS:
        row(f"{p_}: program cost", f"={M}D{R[f'in.prog_{p_}']}",
            lambda y, c, pc, i, p_=p_: f"={c}{R[f'f.parts_{p_}']}*{c}{R[f'f.cpp_{p_}']}",
            indent=2, key=f"f.cost_{p_}",
            note="Participants times cost per participant." if p_ == I.PROGRAMS[0] else "")
    r += 1

    # ---- expenses --------------------------------------------------------
    r = S.section(ws, r, "Expenses", width=8)
    row("Program services", f"={H}G{R['h.prog']}",
        lambda y, c, p, i: f"=SUM({c}{first_cost}:{c}{first_cost + 3})",
        bold=True, indent=0, key="f.prog")
    row("Management and general", f"={H}G{R['h.mg']}",
        lambda y, c, p, i: (
            f"={p}{r}*(1+{M}$B${R['in.drv_inflation_personnel']})*"
            f"(1+{M}$B${R['in.mgvar']}*({c}{R['f.prog']}/{p}{R['f.prog']}-1))"),
        key="f.mg",
        note="Semi-fixed: only the variable share moves with program scale.")
    row("Fundraising", f"={H}G{R['h.fr']}",
        lambda y, c, p, i: (
            f"={p}{r}*(1+{M}$B${R['in.drv_inflation_personnel']})*"
            f"(1+{M}$B${R['in.frvar']}*({c}{R['f.contrib']}/{p}{R['f.contrib']}-1))"),
        key="f.fr")
    row("Total expenses", f"={H}G{R['h.exp']}",
        lambda y, c, p, i: f"={c}{R['f.prog']}+{c}{R['f.mg']}+{c}{R['f.fr']}",
        bold=True, indent=0, key="f.exp")
    r += 1

    # ---- personnel -------------------------------------------------------
    r = S.section(ws, r, "Personnel (implied by delivery hours)", width=8)
    hours_terms = " + ".join(
        f"{{c}}{R[f'f.parts_{p_}']}*{M}$G${R[f'in.prog_{p_}']}" for p_ in I.PROGRAMS)
    row("Program delivery hours", None,
        lambda y, c, p, i: "=" + hours_terms.replace("{c}", c), fmt=S.NUM1, key="f.hours",
        note="Participants times staff hours per participant, summed across programs.")
    row("Implied FTE", f"={M}B{R['in.fte']}",
        lambda y, c, p, i: f"={c}{R['f.hours']}/{M}$B${R['in.hoursfte']}/{M}$B${R['in.progfte']}",
        fmt=S.NUM1, key="f.fte")
    row("Average salary", f"={M}B{R['in.fte']}*0+{I.AVG_SALARY_2024:.4f}",
        lambda y, c, p, i: f"={p}{r}*(1+{M}$B${R['in.drv_inflation_personnel']})",
        fmt=S.MONEY, key="f.salary")
    row("Fully loaded personnel cost", None,
        lambda y, c, p, i: (
            f"={c}{R['f.fte']}*{c}{R['f.salary']}*"
            f"(1+{M}$B${R['in.ptaxrate']}+{M}$B${R['in.benefits']})"),
        key="f.personnel")
    row("Personnel as a share of expenses", None,
        lambda y, c, p, i: f"={c}{R['f.personnel']}/{c}{R['f.exp']}", fmt=S.PCT1,
        key="f.persratio",
        note="A sanity check on the expense build. If this drifts far from the FY2024 level of "
             "about 66%, the participant or cost drivers are implying a staffing level the "
             "expense lines do not fund.")
    r += 1

    # ---- result and reserves --------------------------------------------
    r = S.section(ws, r, "Result and reserves", width=8)
    row("Operating surplus / (deficit)", f"={H}G{R['h.result']}",
        lambda y, c, p, i: f"={c}{R['f.rev']}-{c}{R['f.exp']}", bold=True, indent=0,
        key="f.result")
    row("Operating margin", f"={H}G{R['h.margin']}",
        lambda y, c, p, i: f"={c}{R['f.result']}/{c}{R['f.rev']}", fmt=S.PCT1, key="f.margin")
    # placeholder; the opening-balance chain is wired up after the closing row exists
    row("Opening liquid reserves", None, lambda y, c, p, i: 0, key="f.liq_open",
        note="Prior year closing balance. The reserve roll-forward is the spine of the "
             "liquidity analysis, so it is built as an explicit opening / movement / closing "
             "chain rather than a single netted line.")
    row("Closing liquid reserves", f"={H}G{R['h.liquid']}",
        lambda y, c, p, i: f"={c}{R['f.liq_open']}+{c}{R['f.result']}", bold=True, indent=0,
        key="f.liq_close")
    row("Closing net assets", f"={H}G{R['h.na']}",
        lambda y, c, p, i: f"={p}{r}+{c}{R['f.result']}", key="f.na")
    row("Average monthly expenses", f"={H}G{R['h.monthly']}",
        lambda y, c, p, i: f"={c}{R['f.exp']}/12", key="f.monthly")
    row("Liquid reserve runway", f"={H}G{R['h.liqrunway']}",
        lambda y, c, p, i: f"={c}{R['f.liq_close']}/{c}{R['f.monthly']}", fmt=S.MONTHS,
        key="f.runway")
    row("Net assets, multiple of expenses", f"={H}G{R['h.namult']}",
        lambda y, c, p, i: f"={c}{R['f.na']}/{c}{R['f.exp']}", fmt=S.MULT, key="f.namult",
        note="BBB Standard 10 caps this at three times annual expenses.")
    r += 1

    # ---- outputs ---------------------------------------------------------
    r = S.section(ws, r, "Program outputs", width=8)
    parts_terms = "+".join(f"{{c}}{R[f'f.parts_{p_}']}" for p_ in I.PROGRAMS)
    row("Participants served", f"={H}G{R['h.parts']}",
        lambda y, c, p, i: "=" + parts_terms.replace("{c}", c), fmt=S.NUM1, key="f.parts")
    out_terms = "+".join(
        f"{{c}}{R[f'f.parts_{p_}']}*{M}$C${R[f'in.prog_{p_}']}/{M}$B${R[f'in.prog_{p_}']}"
        for p_ in I.PROGRAMS)
    row("Successful outcomes", f"={H}G{R['h.places']}",
        lambda y, c, p, i: "=" + out_terms.replace("{c}", c), fmt=S.NUM1, key="f.out",
        note="Participants times each program's base-year outcome rate, held constant.")
    row("Program expense ratio", f"={H}G{R['h.progratio']}",
        lambda y, c, p, i: f"={c}{R['f.prog']}/{c}{R['f.exp']}", fmt=S.PCT1, key="f.progratio")
    row("Cost per participant", f"={H}G{R['h.cpp']}",
        lambda y, c, p, i: f"={c}{R['f.prog']}/{c}{R['f.parts']}", key="f.cpp")
    row("Cost per successful outcome", f"={H}G{R['h.cpo']}",
        lambda y, c, p, i: f"={c}{R['f.prog']}/{c}{R['f.out']}", key="f.cpo")

    # fix the opening-reserve chain and investment-income reference
    ws[f"C{R['f.liq_open']}"] = f"='Historical analysis'!G{R['h.liquid']}"
    ws[f"D{R['f.liq_open']}"] = f"=C{R['f.liq_close']}"
    ws[f"E{R['f.liq_open']}"] = f"=D{R['f.liq_close']}"
    for y in FY:
        col = FCOL[y]
        ws[f"{col}{R['f.invinc']}"] = f"={col}{R['f.liq_open']}*{M}$B${R['in.drv_investment_return']}"

    # participant caps per program
    for p_ in I.PROGRAMS:
        cap = I.PROGRAM_BASE[p_]["max_participants"]
        for i, y in enumerate(FY):
            col = FCOL[y]
            pcol = cols[i]
            ws[f"{col}{R[f'f.parts_{p_}']}"] = (
                f"=MIN({pcol}{R[f'f.parts_{p_}']}*"
                f"(1+{M}$B${R['in.drv_growth_participants']}),{cap})")

    ws.conditional_formatting.add(
        f"C{R['f.result']}:E{R['f.result']}",
        CellIsRule(operator="lessThan", formula=["0"],
                   fill=PatternFill("solid", bgColor="FFC7CE")))
    ws.conditional_formatting.add(
        f"C{R['f.runway']}:E{R['f.runway']}",
        CellIsRule(operator="lessThan", formula=[f"'Model inputs'!$B${R['in.bm_res_hi']}"],
                   fill=PatternFill("solid", bgColor="FFC7CE")))
    ws.conditional_formatting.add(
        f"C{R['f.namult']}:E{R['f.namult']}",
        CellIsRule(operator="greaterThan", formula=[f"'Model inputs'!$B${R['in.bm_na']}"],
                   fill=PatternFill("solid", bgColor="FFEB9C")))

    S.freeze(ws, "B5")
    S.finish(ws, S.ACCENT)
    return ws


# ==========================================================================
# Sheet 5 - Dashboard
# ==========================================================================

def sheet_dashboard(wb):
    ws = wb.create_sheet("Dashboard", 0)
    S.set_widths(ws, {"A": 2, "B": 38, "C": 16, "D": 16, "E": 16, "F": 16, "G": 16, "H": 16,
                      "I": 2, "J": 20})
    ws.sheet_view.showGridLines = False

    r = S.title_block(ws, "Financial dashboard",
                      f"{I.ORG_NAME}  |  FY2019 - FY2024 actual, FY2025 - FY2027 Base case forecast",
                      width=8)

    H = "'Historical analysis'!"
    F = "'Forecast'!"

    # KPI strip
    kpis = [
        ("FY2024 total revenue", f"={H}G{R['h.rev']}", S.MONEY0),
        ("FY2024 operating result", f"={H}G{R['h.result']}", S.MONEY0),
        ("Program expense ratio", f"={H}G{R['h.progratio']}", S.PCT1),
        ("Cost to raise $1", f"={H}G{R['h.frratio']}", '$0.00'),
        ("Liquid reserve runway", f"={H}G{R['h.liqrunway']}", S.MONTHS),
        ("Net assets / expenses", f"={H}G{R['h.namult']}", S.MULT),
    ]
    kr = r
    for i, (name, f, fmt) in enumerate(kpis):
        col = 2 + i
        c = ws.cell(row=kr, column=col, value=name)
        c.font = Font(name="Calibri", size=8.5, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=S.NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        v = ws.cell(row=kr + 1, column=col, value=f)
        v.number_format = fmt
        v.font = Font(name="Calibri", size=13, bold=True, color=S.NAVY)
        v.alignment = Alignment(horizontal="center", vertical="center")
        v.fill = PatternFill("solid", fgColor=S.NAVY_LIGHT)
        v.border = Border(bottom=Side(style="medium", color=S.NAVY))
    ws.row_dimensions[kr].height = 28
    ws.row_dimensions[kr + 1].height = 30
    r = kr + 3

    # combined actual + forecast series for charting
    r = S.section(ws, r, "Revenue, expenses and result", width=8)
    hdr = r
    S.label(ws, r, 2, "", bold=True)
    allyears = HY + FY
    for i, y in enumerate(allyears):
        c = ws.cell(row=r, column=3 + i, value=f"FY{y}")
        c.font = Font(name="Calibri", size=9, bold=True, color=S.NAVY)
        c.alignment = Alignment(horizontal="center")
    r += 1
    series = [
        ("Total revenue", R["h.rev"], R["f.rev"], S.MONEY),
        ("Total expenses", R["h.exp"], R["f.exp"], S.MONEY),
        ("Operating result", R["h.result"], R["f.result"], S.MONEY),
        ("Liquid reserves", R["h.liquid"], R["f.liq_close"], S.MONEY),
        ("Participants", R["h.parts"], R["f.parts"], S.NUM1),
        ("Successful outcomes", R["h.places"], R["f.out"], S.NUM1),
        ("Program expense ratio", R["h.progratio"], R["f.progratio"], S.PCT1),
        ("Liquid reserve runway", R["h.liqrunway"], R["f.runway"], S.MONTHS),
        ("Cost per outcome", R["h.cpo"], R["f.cpo"], S.MONEY),
    ]
    chart_rows = {}
    for name, hrow, frow, fmt in series:
        S.label(ws, r, 2, name, size=9.5)
        for i, y in enumerate(allyears):
            col = 3 + i
            if y in HY:
                f = f"={H}{HCOL[y]}{hrow}"
            else:
                f = f"={F}{FCOL[y]}{frow}"
            c = ws.cell(row=r, column=col, value=f)
            c.number_format = fmt
            c.font = Font(name="Calibri", size=9,
                          color=S.INK if y in HY else S.MUTED,
                          italic=(y in FY))
            c.alignment = Alignment(horizontal="right")
        chart_rows[name] = r
        r += 1
    r += 1

    # charts
    ch = LineChart()
    ch.title = "Revenue and expenses, actual then Base case forecast"
    ch.height, ch.width = 7.6, 17
    ch.y_axis.numFmt = '#,##0,,"M"'
    ch.y_axis.title = "USD"
    for name in ("Total revenue", "Total expenses"):
        ref = Reference(ws, min_col=3, max_col=2 + len(allyears), min_row=chart_rows[name])
        s = Series(ref, title=name)
        ch.series.append(s)
    ch.set_categories(Reference(ws, min_col=3, max_col=2 + len(allyears), min_row=hdr))
    ws.add_chart(ch, f"B{r}")

    ch2 = BarChart()
    ch2.type = "col"
    ch2.title = "Operating surplus / (deficit)"
    ch2.height, ch2.width = 7.6, 17
    ch2.y_axis.numFmt = '#,##0,,"M"'
    ref = Reference(ws, min_col=3, max_col=2 + len(allyears), min_row=chart_rows["Operating result"])
    ch2.series.append(Series(ref, title="Operating result"))
    ch2.set_categories(Reference(ws, min_col=3, max_col=2 + len(allyears), min_row=hdr))
    ws.add_chart(ch2, f"B{r + 16}")

    ch3 = LineChart()
    ch3.title = "Liquid reserve runway, months"
    ch3.height, ch3.width = 7.6, 17
    ref = Reference(ws, min_col=3, max_col=2 + len(allyears), min_row=chart_rows["Liquid reserve runway"])
    ch3.series.append(Series(ref, title="Runway (months)"))
    ch3.set_categories(Reference(ws, min_col=3, max_col=2 + len(allyears), min_row=hdr))
    ws.add_chart(ch3, f"B{r + 32}")

    ch4 = LineChart()
    ch4.title = "Participants served and successful outcomes"
    ch4.height, ch4.width = 7.6, 17
    for name in ("Participants", "Successful outcomes"):
        ref = Reference(ws, min_col=3, max_col=2 + len(allyears), min_row=chart_rows[name])
        ch4.series.append(Series(ref, title=name))
    ch4.set_categories(Reference(ws, min_col=3, max_col=2 + len(allyears), min_row=hdr))
    ws.add_chart(ch4, f"B{r + 48}")

    S.finish(ws, S.NAVY)
    return ws


def build():
    wb = Workbook()
    wb.remove(wb.active)
    sheet_990(wb)
    sheet_inputs(wb)
    sheet_historical(wb)
    sheet_forecast(wb)
    sheet_dashboard(wb)
    wb.move_sheet("Dashboard", offset=-4)
    wb.save(OUT)
    print(f"wrote {OUT}")
    return OUT


if __name__ == "__main__":
    build()
