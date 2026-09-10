"""
01_Assumptions.xlsx

The control document for the whole project. Every input used anywhere in the
model appears here exactly once, with its provenance tier, its value, its unit,
its source and a note explaining why it is what it is.

If a reviewer only opens one file, it should be this one.
"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter

import inputs as I
import styles as S
from provenance import TIER_ORDER, TIER_COLOR, TIER_DESCRIPTION

OUT = Path(__file__).resolve().parents[1] / "excel-models" / "01_Assumptions.xlsx"


def sheet_readme(wb):
    ws = wb.create_sheet("Read me first")
    S.set_widths(ws, {"A": 3, "B": 108})
    ws.sheet_view.showGridLines = False

    r = 2
    c = ws.cell(row=r, column=2, value="Nonprofit Financial Planning, Scenario Forecasting and Impact Measurement")
    c.font = Font(name="Calibri", size=16, bold=True, color=S.NAVY)
    r += 1
    c = ws.cell(row=r, column=2, value=f"Assumptions register  |  {I.ORG_NAME} (EIN {I.ORG_EIN})  |  as of {I.AS_OF}")
    c.font = Font(name="Calibri", size=10, italic=True, color=S.MUTED)
    r += 2

    blocks = [
        ("What this project is",
         "An independent financial planning and impact measurement model for a real US nonprofit, built "
         "entirely from public data plus clearly labelled analyst assumptions. It answers one question: "
         "how should this organization allocate limited funding across its programs while staying "
         "financially sustainable and maximizing measurable impact?"),
        ("What is real",
         "All historical financial totals are taken verbatim from IRS Form 990 filings for FY2019 through "
         "FY2024. All published impact metrics - participants served, placements, average starting salary, "
         "average salary gain - are taken from the organization's own public reporting. Sector benchmarks "
         "are cited to BBB Wise Giving Alliance and Propel Nonprofits."),
        ("What is not real",
         "Everything below the reported totals is modelled: the split of expenses into program, "
         "administrative and fundraising categories; the four-program portfolio and its unit economics; "
         "the composition of the balance sheet; the donor-count and average-gift drivers; and every "
         "forward-looking number in the model. None of the forecasts are the organization's plans, budget "
         "or guidance, and none should be attributed to it."),
        ("The rule that keeps this honest",
         "Where a public control total exists, the modelled composition must reconcile to it exactly. The "
         "four programs' costs sum to the modelled program expense pool, which is a fixed percentage of "
         "the filed total expense figure. The five revenue categories sum to the filed contributions "
         "total. The balance sheet lines sum to filed total assets. The composition is assumed; the "
         "totals are real."),
        ("How to read the tiers",
         "Every row in the register carries one of four provenance tags. PUBLIC means verbatim from a "
         "filing. DERIVED means arithmetic on public values only, by a stated method. BENCHMARK means a "
         "published sector standard. SYNTHETIC means an analyst assumption. Filter the register by tier "
         "to see exactly how much of the model rests on judgement."),
        ("Where to go next",
         "02_Financial_Model.xlsx holds the historical analysis and the three-year forecast and "
         "is the centre of the project. 04_Scenario_Model.xlsx holds the Base, Upside and Downside cases "
         "and the funding sensitivity work. 05_Program_Economics.xlsx compares the programs. "
         "06_Resource_Allocation.xlsx solves the allocation question. Management_Report.pdf "
         "explains what all of it means."),
    ]
    for head, body in blocks:
        h = ws.cell(row=r, column=2, value=head)
        h.font = Font(name="Calibri", size=11.5, bold=True, color=S.NAVY)
        r += 1
        b = ws.cell(row=r, column=2, value=body)
        b.font = Font(name="Calibri", size=10, color=S.INK)
        b.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 14 * (len(body) // 105 + 1)
        r += 2

    h = ws.cell(row=r, column=2, value="Provenance key")
    h.font = Font(name="Calibri", size=11.5, bold=True, color=S.NAVY)
    r += 1
    for t in TIER_ORDER:
        cell = ws.cell(row=r, column=2, value=f"{t}   -   {TIER_DESCRIPTION[t]}")
        cell.font = Font(name="Calibri", size=10, bold=True, color=TIER_COLOR[t])
        r += 1
    r += 1

    counts = I.REGISTER.tier_counts()
    total = sum(counts.values())
    h = ws.cell(row=r, column=2, value="Composition of this model")
    h.font = Font(name="Calibri", size=11.5, bold=True, color=S.NAVY)
    r += 1
    for t in TIER_ORDER:
        cell = ws.cell(row=r, column=2,
                       value=f"{t:<10}  {counts[t]:>4} inputs   ({counts[t]/total:.0%})")
        cell.font = Font(name="Consolas", size=10, color=TIER_COLOR[t])
        r += 1
    cell = ws.cell(row=r, column=2, value=f"{'TOTAL':<10}  {total:>4} inputs")
    cell.font = Font(name="Consolas", size=10, bold=True, color=S.INK)
    r += 2

    d = ws.cell(row=r, column=2, value=(
        "Disclaimer. This is an independent analytical exercise prepared from public information for "
        "portfolio and educational purposes. It is not affiliated with, endorsed by, or reviewed by "
        f"{I.ORG_NAME}. It is not financial advice and it is not a representation of the organization's "
        "financial position, plans or prospects. Forward-looking figures are the author's assumptions."))
    d.font = Font(name="Calibri", size=9, italic=True, color=S.WARN)
    d.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 42

    S.finish(ws, S.NAVY)
    return ws


def sheet_register(wb):
    ws = wb.create_sheet("Assumptions register")
    S.set_widths(ws, {"A": 30, "B": 46, "C": 18, "D": 20, "E": 12, "F": 58, "G": 74})
    ws.sheet_view.showGridLines = False

    r = S.title_block(ws, "Assumptions register",
                      "Every input used anywhere in this project, with its provenance. "
                      "Filter column E to isolate a tier.", width=7)

    headers = ["Section", "Input", "Value", "Unit", "Tier", "Source", "Why it is what it is"]
    hr = r
    r = S.header_row(ws, r, headers, widths=None)
    first_data = r

    for inp in I.REGISTER:
        S.label(ws, r, 1, inp.section, size=9)
        S.label(ws, r, 2, inp.label, size=9.5)
        v = inp.value
        if isinstance(v, str) and inp.fmt == "@":
            c = ws.cell(row=r, column=3, value=(v[:60] + "..." if len(v) > 63 else v))
            c.font = Font(name="Calibri", size=9, italic=True, color=S.MUTED)
            c.alignment = Alignment(horizontal="left", wrap_text=False)
        else:
            S.put(ws, r, 3, v, fmt=inp.fmt, is_input=(inp.tier == "SYNTHETIC"))
        S.label(ws, r, 4, inp.unit, size=9, color=S.MUTED)
        S.tier_cell(ws, r, 5, inp.tier)
        c = ws.cell(row=r, column=6, value=inp.source)
        c.font = Font(name="Calibri", size=8, color=S.MUTED)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c = ws.cell(row=r, column=7, value=inp.note)
        c.font = Font(name="Calibri", size=8.5, color=S.INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 26
        if (r - first_data) % 2 == 1:
            S.band(ws, r, range(1, 8))
        r += 1

    last = r - 1
    ref = f"A{hr}:G{last}"
    t = Table(displayName="AssumptionsRegister", ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
    ws.add_table(t)
    ws.auto_filter.ref = ref
    S.freeze(ws, "C" + str(first_data))
    S.finish(ws, S.ACCENT)
    return ws


def sheet_sources(wb):
    ws = wb.create_sheet("Sources")
    S.set_widths(ws, {"A": 4, "B": 34, "C": 96, "D": 16})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Sources",
                      "Every external source relied on, with what was taken from it.", width=4)

    rows = [
        ("IRS Form 990, EIN 94-3346127",
         I.SRC_990,
         "PUBLIC",
         "Total revenue, total functional expenses, total assets, total liabilities, contributions and "
         "grants, program service revenue, investment income, officer compensation, other salaries and "
         "wages, payroll taxes, and net assets, for fiscal years 2019 through 2024."),
        ("Upwardly Global impact reporting (2025)",
         I.SRC_IMPACT_2025,
         "PUBLIC",
         "13,350 jobseekers supported, 1,730+ placed, $67,580 average starting salary, $58,790 average "
         "salary gain, 100+ employer partners, 87 workforce partnerships, 188 countries represented."),
        ("Upwardly Global 2022 Annual Report",
         I.SRC_IMPACT_2022,
         "PUBLIC",
         "7,041 program participants, 1,116 placed in thriving-wage jobs, $66,481 average annual "
         "starting salary, $74M estimated annual economic contribution, 300 corporate partners."),
        ("BBB Wise Giving Alliance Standards",
         I.SRC_BBB,
         "BENCHMARK",
         "Standard 8: at least 65% of total expenses on programs. Standard 9: no more than 35% of "
         "related contributions on fundraising. Standard 10: unrestricted net assets available for use "
         "should not exceed three times prior-year expenses or current-year budget."),
        ("Propel Nonprofits, operating reserves",
         I.SRC_PROPEL,
         "BENCHMARK",
         "A commonly used reserve goal is three to six months of expenses. At the low end reserves "
         "should cover at least one full payroll including taxes. At the high end they should not "
         "exceed two years of budget."),
        ("Charity Navigator rating profile",
         I.SRC_CN,
         "PUBLIC",
         "4 of 4 stars, 96% overall score, all four beacons complete. Used as an upper sanity bound on "
         "the modelled functional expense split - a 4-star organization is very unlikely to sit near "
         "the BBB minimum program ratio."),
    ]

    r = S.header_row(ws, r, ["", "Source", "What was taken from it", "Tier"], start_col=1)
    for i, (name, url, tier, what) in enumerate(rows, start=1):
        S.label(ws, r, 1, str(i), size=9, color=S.MUTED)
        c = ws.cell(row=r, column=2, value=name)
        c.font = Font(name="Calibri", size=10, bold=True, color=S.NAVY)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c2 = ws.cell(row=r, column=3, value=what)
        c2.font = Font(name="Calibri", size=9, color=S.INK)
        c2.alignment = Alignment(wrap_text=True, vertical="top")
        S.tier_cell(ws, r, 4, tier)
        ws.row_dimensions[r].height = 52
        r += 1
        u = ws.cell(row=r, column=3, value=url)
        u.font = Font(name="Calibri", size=8, color="0563C1", underline="single")
        u.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 24
        r += 1

    r += 1
    r = S.section(ws, r, "Data quality notes", width=4)
    notes = [
        "The FY2024 filing exposes total revenue, total expenses, total assets, net assets, officer "
        "compensation and other salaries in machine-readable form, but not the revenue mix. "
        "Contributions, program service revenue and investment income for FY2024 are therefore derived "
        "from the published approximate mix of 91% contributions and 4% program services, and are "
        "tagged DERIVED rather than PUBLIC.",
        "Form 990 Part IX splits expenses into program, management and general, and fundraising. That "
        "split is not exposed in the machine-readable extract used here, so it is modelled. This is the "
        "single largest piece of judgement in the historical analysis and it drives the program expense "
        "ratio, the administrative ratio and the fundraising cost ratio directly.",
        "Published participant and placement counts exist for 2022 and 2025 only. Intervening and "
        "earlier years are interpolated at constant growth between those two observations. The "
        "back-cast to 2019 through 2021 is the weakest part of that chain and should be read as "
        "illustrative rather than as an estimate of what actually happened.",
        "The organization's own reporting uses at least two different all-time participant definitions "
        "- 13,460 individuals and families with improved lives in the 2022 report against 35,000+ "
        "jobseekers on the 2025 impact page. Definitions of who counts as served clearly changed. Only "
        "the single-year figures are used here, and only from the two dates cited.",
    ]
    for n in notes:
        c = ws.cell(row=r, column=2, value=n)
        c.font = Font(name="Calibri", size=9, color=S.INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        ws.row_dimensions[r].height = 14 * (len(n) // 120 + 1)
        r += 1

    S.finish(ws, S.NAVY)
    return ws


def sheet_reconciliation(wb):
    """Live proof that the modelled composition ties to the filed totals."""
    ws = wb.create_sheet("Reconciliation checks")
    S.set_widths(ws, {"A": 46, "B": 18, "C": 18, "D": 14, "E": 12, "F": 62})
    ws.sheet_view.showGridLines = False
    r = S.title_block(ws, "Reconciliation checks",
                      "Live formulas proving that every modelled split sums back to a filed total. "
                      "Any FAIL here invalidates the model.", width=6)

    r = S.header_row(ws, r, ["Check", "Modelled", "Filed / control", "Difference",
                             "Result", "What it proves"], start_col=1)

    checks = []

    # functional expense split -> filed total expenses
    for y in I.HIST_YEARS:
        p, m, f = I.FUNCTIONAL_SPLIT[y]
        exp = I.F990[y]["total_expense"]
        checks.append((
            f"FY{y} functional expense split sums to filed total",
            exp * p + exp * m + exp * f, exp,
            "Program + management and general + fundraising must equal the Form 990 total expense line."))

    # revenue mix -> filed contributions
    for y in (2023, 2024):
        contrib = I.F990[y]["contributions"]
        modelled = sum(contrib * s for s in I.REVENUE_MIX.values())
        checks.append((
            f"FY{y} revenue category split sums to contributions",
            modelled, contrib,
            "The five funding categories must sum to the filed contributions and grants total."))

    # program costs -> program expense pool
    pool = I.F990[2024]["total_expense"] * I.FUNCTIONAL_SPLIT[2024][0]
    modelled = sum(I.PROGRAM_BASE[p]["program_cost"] for p in I.PROGRAMS)
    checks.append((
        "FY2024 program costs sum to the program expense pool",
        modelled, pool,
        "The four programs must sum to the modelled program services pool, which is itself a fixed "
        "share of the filed total expense."))

    # participants -> interpolated total
    p22 = I.IMPACT_PUBLIC[2022]["participants"]
    p25 = I.IMPACT_PUBLIC[2025]["participants"]
    interp = p22 * (p25 / p22) ** (2 / 3)
    checks.append((
        "FY2024 program participants sum to the interpolated total",
        sum(I.PROGRAM_BASE[p]["participants"] for p in I.PROGRAMS), interp,
        "The four programs' participant counts must sum to the count interpolated between the two "
        "published observations."))

    # placements
    l22 = I.IMPACT_PUBLIC[2022]["placements"]
    l25 = I.IMPACT_PUBLIC[2025]["placements"]
    interp_l = l22 * (l25 / l22) ** (2 / 3)
    checks.append((
        "FY2024 successful outcomes sum to the interpolated total",
        sum(I.PROGRAM_BASE[p]["placements"] for p in I.PROGRAMS), interp_l,
        "Rounded to whole placements, so a difference under one person is expected."))

    # balance sheet
    checks.append((
        "FY2024 balance sheet composition sums to filed total assets",
        sum(I.BALANCE_SHEET_2024.values()), I.F990[2024]["total_assets"],
        "The modelled asset classes must sum to the filed total assets figure."))

    # net assets identity
    checks.append((
        "FY2024 assets less liabilities equals filed net assets",
        I.F990[2024]["total_assets"] - I.F990[2024]["total_liab"], I.NET_ASSETS_2024_PUBLIC,
        "Confirms the derived FY2024 liability figure is consistent with the two public values it "
        "was backed out of."))

    # revenue mix shares
    checks.append((
        "Revenue category shares sum to 100%",
        sum(I.REVENUE_MIX.values()) * 1_000_000, 1_000_000,
        "Scaled by a million so the tolerance test is meaningful in the same units as the others."))

    # personnel
    sal24 = I.F990[2024]["officer_comp"] + I.F990[2024]["other_salaries"]
    checks.append((
        "FY2024 implied average salary times FTE equals filed salaries",
        I.FTE_2024 * I.AVG_SALARY_2024, sal24,
        "Confirms the assumed headcount and the solved average salary reproduce the filed salary total."))

    # staff hours
    hours = sum(I.PROGRAM_BASE[p]["participants"] * I.PROGRAM_BASE[p]["staff_hours_per_participant"]
                for p in I.PROGRAMS)
    checks.append((
        "Program delivery hours equal FTE times program share times hours per FTE",
        hours, I.FTE_2024 * I.PROGRAM_FTE_SHARE * I.STAFF_HOURS_PER_FTE,
        "Ties the capacity constraint used in the allocation model to the personnel cost in the P&L."))

    first = r
    for name, modelled, control, what in checks:
        S.label(ws, r, 1, name, size=9.5)
        S.put(ws, r, 2, modelled, fmt=S.MONEY2)
        S.put(ws, r, 3, control, fmt=S.MONEY2)
        ws.cell(row=r, column=4, value=f"=B{r}-C{r}").number_format = S.MONEY2
        ws.cell(row=r, column=4).font = Font(name="Calibri", size=10)
        ws.cell(row=r, column=4).alignment = Alignment(horizontal="right")
        res = ws.cell(row=r, column=5, value=f'=IF(ABS(D{r})<1,"PASS","FAIL")')
        res.font = Font(name="Calibri", size=10, bold=True)
        res.alignment = Alignment(horizontal="center")
        c = ws.cell(row=r, column=6, value=what)
        c.font = Font(name="Calibri", size=8.5, color=S.MUTED)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 28
        r += 1

    last = r - 1
    r += 1
    S.label(ws, r, 1, "All checks pass", bold=True)
    allc = ws.cell(row=r, column=5,
                   value=f'=IF(COUNTIF(E{first}:E{last},"FAIL")=0,"PASS","FAIL")')
    allc.font = Font(name="Calibri", size=11, bold=True)
    allc.alignment = Alignment(horizontal="center")
    r += 2
    S.note(ws, r, "Tolerance is one dollar, or one person for count checks. The participant and outcome "
                  "checks compare integer program counts against a continuous interpolation, so a "
                  "sub-unit difference there is arithmetic rounding rather than a modelling error.",
           width=6)

    S.finish(ws, S.ACCENT)
    return ws


def build():
    wb = Workbook()
    wb.remove(wb.active)
    sheet_readme(wb)
    sheet_register(wb)
    sheet_sources(wb)
    sheet_reconciliation(wb)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT)
    print(f"wrote {OUT}")
    return OUT


if __name__ == "__main__":
    build()
