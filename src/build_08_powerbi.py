"""
powerbi/

Generates a Power BI-ready star schema, a DAX measure library and a build guide.

A .pbix file is a proprietary binary and cannot be authored programmatically, so
what is produced here is everything that goes inside one: a clean dimensional
model, the measures written out in DAX, and a page-by-page assembly spec.
Importing the data and pasting the measures reproduces the dashboard.

The model is emitted in two forms. Sixteen CSVs, which diff properly in version
control and suit Power BI Desktop's folder connector; and one workbook holding
the same sixteen tables as named Excel Tables, because Power BI in the browser
imports a single file at a time. Both come from the same rows, so they cannot
drift apart.

The schema is a proper star: narrow fact tables keyed to conformed dimensions,
no snowflaking, and one shared date and scenario dimension so a single slicer
filters every visual on the page. That is the part worth getting right - a
dashboard built on a wide flat table works until the first time someone asks a
question the table was not shaped for.
"""

import csv
from pathlib import Path

import inputs as I
import engine

ROOT = Path(__file__).resolve().parents[1] / "powerbi"
DATA = ROOT / "data"

# Every table is emitted twice: once as a CSV, and once as a sheet in a single
# workbook. The CSVs are the better artefact in version control - they diff
# line by line, where a binary .xlsx shows only "changed" - and they are what
# Power BI Desktop's folder connector expects. The workbook exists because
# Power BI in the browser imports one file at a time, so sixteen CSVs would
# mean sixteen separate imports.
TABLES: list[tuple[str, list[str], list[list]]] = []


def write_csv(name: str, header: list[str], rows: list[list]):
    DATA.mkdir(parents=True, exist_ok=True)
    path = DATA / name
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    TABLES.append((name.removesuffix(".csv"), header, rows))
    print(f"  {name:<32} {len(rows):>5} rows")
    return path


def write_workbook():
    """
    One workbook, one Excel Table per sheet, for importing in the browser.

    Written with xlsxwriter rather than openpyxl. openpyxl's fast writer (used
    whenever lxml is installed) emits numeric cells as t="n" and strings as
    inline <is> runs with no sharedStrings.xml part. That is legal OOXML but it
    is not the shape Excel itself writes, and Power Query rejects it outright
    with "We were unable to load this Excel file because we couldn't understand
    its format" - a parser error, not a data error, so it gives no clue which
    cell is at fault. xlsxwriter produces Excel-shaped output: a real shared
    string table, and no type attribute on numbers.

    Each range is a genuine named Excel Table, so Power BI lists it under Tables
    rather than Sheets. Missing values are written as blanks rather than empty
    strings, so a numeric column with a gap is not typed as text.
    """
    import xlsxwriter

    out = ROOT / "PowerBI_Data_Model.xlsx"
    wb = xlsxwriter.Workbook(str(out), {"in_memory": True, "strings_to_numbers": False})

    head = wb.add_format({"bold": True, "font_color": "#12324F", "font_size": 14})
    key = wb.add_format({"bold": True, "font_color": "#12324F", "font_size": 10})
    body = wb.add_format({"font_size": 10, "text_wrap": True, "valign": "top"})
    mono = wb.add_format({"font_name": "Consolas", "font_size": 9})
    hdr = wb.add_format({"bold": True, "font_size": 10})

    # Plain sheet, deliberately NOT an Excel Table, so it appears under
    # "Sheets" rather than "Tables" in the import dialog and is easy to skip.
    ws = wb.add_worksheet("_README")
    ws.hide_gridlines(2)
    ws.set_column("A:A", 34)
    ws.set_column("B:B", 96)
    ndim = sum(1 for n, _, _ in TABLES if n.startswith("dim"))
    nfact = sum(1 for n, _, _ in TABLES if n.startswith("fact"))
    lines = [
        ("Power BI data model", ""),
        ("", ""),
        ("What this is", "Every table of the Power BI star schema, one per sheet, as a named "
                         "Excel Table. Import this single file rather than the sixteen CSVs "
                         "in the same folder - they hold identical data."),
        ("How to use it", "Power BI: Get data > Excel workbook > Import. In the Navigator tick "
                          "the sixteen TABLES listed below - not the sheets they sit on, and "
                          "not this sheet. The table names are what the DAX measures expect."),
        ("Then", "Build the relationships, then add the measures - measures_dax_query.dax "
                 "in the browser, measures.dax in Desktop. "
                 "Both steps are in BUILD_GUIDE.md - relationships are NOT carried over "
                 "by the import and must be created by hand."),
        ("", ""),
        ("Tables in this workbook", f"{len(TABLES)} - {ndim} dimensions, {nfact} facts"),
        ("", ""),
        ("Provenance", "Historical figures come from IRS Form 990 filings. Program-level detail "
                       "and all forward-looking figures are analyst assumptions, not the "
                       "organization's guidance. See data/SOURCES.md."),
    ]
    for i, (k, v) in enumerate(lines):
        ws.write_string(i, 0, k, head if i == 0 else key)
        if v:
            ws.write_string(i, 1, v, body)
            ws.set_row(i, 13 * (len(v) // 92 + 1))
    r = len(lines) + 1
    ws.write_string(r, 0, "Table to import", key)
    ws.write_string(r, 1, "Rows", key)
    ws.write_string(r, 2, "Lives on", key)
    for idx, (name, _h, rows) in enumerate(TABLES, start=1):
        r += 1
        ws.write_string(r, 0, name, mono)
        ws.write_number(r, 1, len(rows), mono)
        ws.write_string(r, 2, f"on sheet {idx:02d} {name}", mono)

    # Sheet names are deliberately NOT the table names. Excel allows a sheet and
    # a table to share a name, but Power Query then has to disambiguate them in
    # its Navigator and appends the table's id - so "dim_year" imports as
    # "dim_year1", and every DAX measure written against "dim_year" breaks. The
    # numeric prefix keeps the tabs in load order for anyone browsing the file
    # in Excel, and leaves the clean name free for the table, which is the thing
    # Power BI actually imports.
    for idx, (name, header, rows) in enumerate(TABLES, start=1):
        ws = wb.add_worksheet(f"{idx:02d} {name}"[:31])
        ws.freeze_panes(1, 0)
        for j, h in enumerate(header):
            ws.set_column(j, j, min(max(len(str(h)) + 4, 12), 46))

        for i, row in enumerate(rows, start=1):
            for j, v in enumerate(row):
                if v == "" or v is None:
                    ws.write_blank(i, j, None)
                elif isinstance(v, bool):
                    ws.write_number(i, j, int(v))
                elif isinstance(v, (int, float)):
                    ws.write_number(i, j, v)
                else:
                    ws.write_string(i, j, str(v))

        ws.add_table(0, 0, len(rows), len(header) - 1, {
            "name": name,
            "style": "Table Style Light 9",
            "columns": [{"header": h, "header_format": hdr} for h in header],
        })

    wb.close()
    total = sum(len(r) for _n, _h, r in TABLES)
    print(f"  {'PowerBI_Data_Model.xlsx':<32} {len(TABLES):>5} tables, {total:,} rows")
    return out


# ==========================================================================
# Dimensions
# ==========================================================================

def dim_year():
    rows = []
    for y in I.HIST_YEARS:
        rows.append([y, f"FY{y}", "Actual", 1 if y == I.BASE_YEAR else 0, y - I.HIST_YEARS[0] + 1])
    for i, y in enumerate(engine.SENSITIVITY_HORIZON):
        label = "Estimate" if i == 0 else "Forecast"
        rows.append([y, f"FY{y}", label, 0, len(I.HIST_YEARS) + i + 1])
    return write_csv("dim_year.csv",
                     ["Year", "YearLabel", "PeriodType", "IsBaseYear", "SortOrder"], rows)


def dim_scenario():
    rows = [[0, "Actual", 0, "Filed results, not a scenario."]]
    for i, s in enumerate(I.SCENARIOS, start=1):
        rows.append([i, s, i, I.SCENARIO_NARRATIVE[s]])
    return write_csv("dim_scenario.csv",
                     ["ScenarioKey", "Scenario", "SortOrder", "Narrative"], rows)


def dim_program():
    pe = engine.program_economics()["programs"]
    rows = []
    for i, p in enumerate(I.PROGRAMS, start=1):
        rows.append([i, p, I.PROGRAM_CODES[p], I.PROGRAM_DESCRIPTION[p],
                     round(pe[p]["salary_gain"], 2), pe[p]["max_participants"],
                     pe[p]["min_funding"], i])
    return write_csv("dim_program.csv",
                     ["ProgramKey", "Program", "Code", "Description", "SalaryGainPerOutcome",
                      "MaxParticipants", "MinViableFunding", "SortOrder"], rows)


def dim_funding_source():
    kind = {"Individual donations": "Private, diversified",
            "Corporate donations": "Private, concentrated",
            "Foundation grants": "Institutional, episodic",
            "Government funding": "Public, appropriation-driven",
            "Fundraising events": "Private, diversified"}
    rows = []
    for i, c in enumerate(I.REVENUE_CATEGORIES, start=1):
        rows.append([i, c, kind[c], round(I.REVENUE_MIX[c], 4),
                     I.DONOR_DRIVERS_2024[c]["retention"], i])
    return write_csv("dim_funding_source.csv",
                     ["SourceKey", "FundingSource", "SourceType", "ShareOfContributions",
                      "DonorRetention", "SortOrder"], rows)


def dim_account():
    rows = [
        [1, "Statement of Activities", "Revenue", "Contributions and grants", 0, 10],
        [2, "Statement of Activities", "Revenue", "Program service revenue", 0, 20],
        [3, "Statement of Activities", "Revenue", "Investment and other income", 0, 30],
        [4, "Statement of Activities", "Revenue", "Fundraising event costs and netting", 0, 40],
        [5, "Statement of Activities", "Revenue", "Total revenue", 0, 50],
        [6, "Statement of Activities", "Expense", "Program services", 1, 60],
        [7, "Statement of Activities", "Expense", "Management and general", 1, 70],
        [8, "Statement of Activities", "Expense", "Fundraising", 1, 80],
        [9, "Statement of Activities", "Expense", "Total expenses", 1, 90],
        [10, "Statement of Activities", "Result", "Operating surplus / (deficit)", 0, 100],
        [11, "Natural Classification", "Expense", "Personnel", 1, 110],
        [12, "Natural Classification", "Expense", "Non-personnel", 1, 120],
        [13, "Financial Position", "Asset", "Total assets", 0, 130],
        [14, "Financial Position", "Liability", "Total liabilities", 0, 140],
        [15, "Financial Position", "Net assets", "Net assets", 0, 150],
        [16, "Financial Position", "Liquidity", "Cash and equivalents", 0, 160],
        [17, "Financial Position", "Liquidity", "Liquid reserves", 0, 170],
    ]
    return write_csv("dim_account.csv",
                     ["AccountKey", "Statement", "Category", "LineItem", "IsExpense", "SortOrder"],
                     rows)


def dim_provenance():
    from provenance import TIER_ORDER, TIER_DESCRIPTION
    rows = [[i, t, TIER_DESCRIPTION[t], i] for i, t in enumerate(TIER_ORDER, start=1)]
    return write_csv("dim_provenance.csv",
                     ["TierKey", "Tier", "TierDescription", "SortOrder"], rows)


# ==========================================================================
# Facts
# ==========================================================================

ACC = {name: key for key, name in [
    (1, "Contributions and grants"), (2, "Program service revenue"),
    (3, "Investment and other income"), (4, "Fundraising event costs and netting"),
    (5, "Total revenue"), (6, "Program services"), (7, "Management and general"),
    (8, "Fundraising"), (9, "Total expenses"), (10, "Operating surplus / (deficit)"),
    (11, "Personnel"), (12, "Non-personnel"), (13, "Total assets"), (14, "Total liabilities"),
    (15, "Net assets"), (16, "Cash and equivalents"), (17, "Liquid reserves"),
]}


def fact_financials():
    """Long-format financial fact: one row per year, scenario and account."""
    rows = []
    hist = engine.historical()["rows"]
    for y in I.HIST_YEARS:
        h = hist[y]
        vals = {
            "Contributions and grants": h["contributions"],
            "Program service revenue": h["program_revenue"],
            "Investment and other income": h["investment_other"],
            "Fundraising event costs and netting": h["revenue_netting"],
            "Total revenue": h["revenue"],
            "Program services": h["program_expense"],
            "Management and general": h["mg_expense"],
            "Fundraising": h["fundraising_expense"],
            "Total expenses": h["expenses"],
            "Operating surplus / (deficit)": h["operating_result"],
            "Personnel": h["personnel"],
            "Non-personnel": h["non_personnel"],
            "Total assets": h["total_assets"],
            "Total liabilities": h["total_liabilities"],
            "Net assets": h["net_assets"],
            "Cash and equivalents": h["cash"],
            "Liquid reserves": h["liquid_reserves"],
        }
        for name, v in vals.items():
            rows.append([y, 0, ACC[name], round(v, 2)])

    for si, scen in enumerate(I.SCENARIOS, start=1):
        for f in engine.forecast(scen, years=engine.SENSITIVITY_HORIZON):
            vals = {
                "Contributions and grants": f.contributions,
                "Program service revenue": f.program_revenue,
                "Investment and other income": f.investment_income,
                "Fundraising event costs and netting": 0.0,
                "Total revenue": f.total_revenue,
                "Program services": f.program_expense,
                "Management and general": f.mg_expense,
                "Fundraising": f.fundraising_expense,
                "Total expenses": f.total_expense,
                "Operating surplus / (deficit)": f.operating_result,
                "Personnel": f.personnel,
                "Non-personnel": f.non_personnel,
                "Net assets": f.closing_net_assets,
                "Liquid reserves": f.closing_liquid,
                "Cash and equivalents": f.closing_cash,
            }
            for name, v in vals.items():
                rows.append([f.year, si, ACC[name], round(v, 2)])
    return write_csv("fact_financials.csv", ["Year", "ScenarioKey", "AccountKey", "Amount"], rows)


def fact_program():
    rows = []
    pe = engine.program_economics()["programs"]
    pk = {p: i for i, p in enumerate(I.PROGRAMS, start=1)}
    hist = engine.historical()["rows"]

    # FY2024 actual, by program
    for p in I.PROGRAMS:
        d = pe[p]
        rows.append([I.BASE_YEAR, 0, pk[p], round(d["participants"], 2), round(d["completers"], 2),
                     round(d["outcomes"], 2), round(d["program_cost"], 2),
                     round(d["direct_cost"], 2), round(d["indirect_cost"], 2),
                     round(d["staff_hours"], 2), round(d["total_earnings_gain"], 2)])

    for si, scen in enumerate(I.SCENARIOS, start=1):
        for f in engine.forecast(scen, years=engine.SENSITIVITY_HORIZON):
            for p in I.PROGRAMS:
                parts = f.participants_by_program[p]
                outs = f.outcomes_by_program[p]
                cost = f.cost_by_program[p]
                comp = parts * I.PROGRAM_BASE[p]["completion_rate"]
                direct = cost * I.PROGRAM_BASE[p]["direct_cost_share"]
                hrs = parts * I.PROGRAM_BASE[p]["staff_hours_per_participant"]
                gain = outs * I.PROGRAM_BASE[p]["salary_gain"]
                rows.append([f.year, si, pk[p], round(parts, 2), round(comp, 2), round(outs, 2),
                             round(cost, 2), round(direct, 2), round(cost - direct, 2),
                             round(hrs, 2), round(gain, 2)])
    return write_csv("fact_program.csv",
                     ["Year", "ScenarioKey", "ProgramKey", "Participants", "Completers",
                      "Outcomes", "ProgramCost", "DirectCost", "IndirectCost", "StaffHours",
                      "EarningsGain"], rows)


def fact_funding():
    rows = []
    sk = {c: i for i, c in enumerate(I.REVENUE_CATEGORIES, start=1)}
    hist = engine.historical()["rows"]
    for y in I.HIST_YEARS:
        contrib = hist[y]["contributions"]
        for c in I.REVENUE_CATEGORIES:
            amt = contrib * I.REVENUE_MIX[c]
            rows.append([y, 0, sk[c], round(amt, 2), "", ""])
    for si, scen in enumerate(I.SCENARIOS, start=1):
        for f in engine.forecast(scen, years=engine.SENSITIVITY_HORIZON):
            for c in I.REVENUE_CATEGORIES:
                rows.append([f.year, si, sk[c], round(f.revenue_by_category[c], 2),
                             round(f.donors_by_category[c], 2),
                             round(f.avg_gift_by_category[c], 2)])
    return write_csv("fact_funding.csv",
                     ["Year", "ScenarioKey", "SourceKey", "Amount", "Donors", "AverageGift"], rows)


def fact_budget_variance():
    bva = engine.budget_vs_actual()
    mapping = {
        "  Contributions and grants": "Contributions and grants",
        "  Program service revenue": "Program service revenue",
        "  Investment and other income": "Investment and other income",
        "Total revenue": "Total revenue",
        "  Program services": "Program services",
        "  Management and general": "Management and general",
        "  Fundraising": "Fundraising",
        "Total expenses": "Total expenses",
        "Operating result": "Operating surplus / (deficit)",
        "Personnel (memo)": "Personnel",
        "Non-personnel (memo)": "Non-personnel",
    }
    rows = []
    for label, acc in mapping.items():
        v = bva["lines"][label]
        rows.append([I.BASE_YEAR, ACC[acc], round(v["budget"], 2), round(v["actual"], 2),
                     round(v["variance"], 2),
                     round(v["variance_pct"], 6) if v["variance_pct"] is not None else "",
                     "Favourable" if v["favourable"] else "Unfavourable"])
    return write_csv("fact_budget_variance.csv",
                     ["Year", "AccountKey", "Budget", "Actual", "Variance", "VariancePct",
                      "Direction"], rows)


def fact_variance_bridge():
    bva = engine.budget_vs_actual()
    d = bva["decomposition"]
    L = bva["lines"]
    rows = [
        [1, "Budgeted program cost", "Start", round(d["total_program_variance"] * 0 + L["  Program services"]["budget"], 2)],
        [2, "Volume variance", "Movement", round(d["volume_variance"], 2)],
        [3, "Rate variance", "Movement", round(d["rate_variance"], 2)],
        [4, "Actual program cost", "End", round(L["  Program services"]["actual"], 2)],
    ]
    return write_csv("fact_variance_bridge.csv", ["StepOrder", "Step", "StepType", "Amount"], rows)


def fact_sensitivity():
    rows = []
    for scen in I.SCENARIOS:
        for s in engine.funding_sensitivity(scen):
            rows.append([
                scen, round(s["shock"], 4),
                round(s["fy_last_result"], 2), round(s["structural_result"], 2),
                round(s["cumulative_result"], 2), round(s["closing_liquid"], 2),
                round(s["liquid_runway_months"], 3), round(s["participants"], 2),
                round(s["outcomes"], 2), round(s["capacity_retained"], 4),
                round(s["participants_lost"], 2), round(s["outcomes_lost"], 2),
                s["years_to_floor"]["years"] or "",
                1 if s["breaches_reserve_floor"] else 0,
            ])
    return write_csv("fact_sensitivity.csv",
                     ["Scenario", "FundingShock", "FinalYearResult", "StructuralResult",
                      "CumulativeResult", "ClosingLiquidReserves", "LiquidRunwayMonths",
                      "Participants", "Outcomes", "CapacityRetained", "ParticipantsLost",
                      "OutcomesLost", "YearsToReserveFloor", "BreachesFloor"], rows)


def fact_allocation():
    comp = engine.allocation_objective_comparison()
    pk = {p: i for i, p in enumerate(I.PROGRAMS, start=1)}
    rows = []
    for row in comp["rows"]:
        run = comp["runs"][row["objective"]]
        for p in I.PROGRAMS:
            d = run["detail"][p]
            rows.append([row["objective"], I.ALLOCATION_OBJECTIVES[row["objective"]], pk[p],
                         round(d["allocation"], 2), round(d["share"], 4),
                         round(d["additional_participants"], 2),
                         round(d["additional_outcomes"], 3),
                         round(d["additional_earnings"], 2),
                         round(d["additional_hours"], 2),
                         round(d["capacity_after"], 4),
                         1 if d["binding_capacity"] else 0,
                         1 if d["binding_max_share"] else 0,
                         1 if d["binding_min_share"] else 0])
    return write_csv("fact_allocation.csv",
                     ["Objective", "ObjectiveLabel", "ProgramKey", "Allocation", "Share",
                      "AdditionalParticipants", "AdditionalOutcomes", "AdditionalEarnings",
                      "AdditionalHours", "CapacityAfter", "BindingCapacity", "BindingMaxShare",
                      "BindingMinShare"], rows)


def fact_tradeoff():
    rows = []
    for t in engine.allocation_tradeoff_frontier():
        if t.get("feasible"):
            rows.append([t["floor"], 1, round(t["outcomes"], 3), round(t["earnings"], 2),
                         round(t["avg_gain"], 2), round(t["cost_per_outcome"], 2)])
        else:
            rows.append([t["floor"], 0, "", "", "", ""])
    return write_csv("fact_quality_tradeoff.csv",
                     ["QualityFloor", "Feasible", "Outcomes", "EarningsGain", "AverageGain",
                      "CostPerOutcome"], rows)


def fact_benchmarks():
    rows = [
        ["Program expense ratio", "Minimum", 0.65, "BBB Standard 8", "Higher is better"],
        ["Fundraising cost ratio", "Maximum", 0.35, "BBB Standard 9", "Lower is better"],
        ["Net assets / expenses", "Maximum", 3.00, "BBB Standard 10", "Lower is better"],
        ["Liquid reserve runway", "Minimum", 3.00, "Propel Nonprofits", "Higher is better"],
        ["Liquid reserve runway", "Target", 6.00, "Propel Nonprofits", "Higher is better"],
        ["Liquid reserve runway", "Maximum", 24.00, "Propel Nonprofits", "Lower is better"],
    ]
    return write_csv("dim_benchmark.csv",
                     ["Metric", "LimitType", "Value", "Source", "Direction"], rows)


def fact_provenance_summary():
    counts = I.REGISTER.tier_counts()
    total = sum(counts.values())
    rows = [[t, n, round(n / total, 4)] for t, n in counts.items()]
    return write_csv("fact_provenance.csv", ["Tier", "InputCount", "Share"], rows)


# ==========================================================================
# DAX
# ==========================================================================

DAX = r"""// =====================================================================
// Nonprofit Financial Planning - DAX measure library
//
// Paste these into Power BI Desktop. Create a blank measure and paste one
// block at a time, or use Tabular Editor to import the file wholesale.
//
// Naming convention: measures are grouped by a leading comment block. A few
// measures depend on others; those are noted where it matters.
//
// The model expects these relationships (all single-direction, many-to-one
// from fact to dimension):
//   fact_financials[Year]        -> dim_year[Year]
//   fact_financials[ScenarioKey] -> dim_scenario[ScenarioKey]
//   fact_financials[AccountKey]  -> dim_account[AccountKey]
//   fact_program[Year]           -> dim_year[Year]
//   fact_program[ScenarioKey]    -> dim_scenario[ScenarioKey]
//   fact_program[ProgramKey]     -> dim_program[ProgramKey]
//   fact_funding[Year]           -> dim_year[Year]
//   fact_funding[ScenarioKey]    -> dim_scenario[ScenarioKey]
//   fact_funding[SourceKey]      -> dim_funding_source[SourceKey]
//   fact_budget_variance[Year]       -> dim_year[Year]
//   fact_budget_variance[AccountKey] -> dim_account[AccountKey]
//   fact_allocation[ProgramKey]  -> dim_program[ProgramKey]
//
// fact_sensitivity and fact_quality_tradeoff are standalone; slice them on
// their own columns rather than wiring them into the star.
// =====================================================================


// ---------------------------------------------------------------------
// 1. Base amount selectors
// ---------------------------------------------------------------------

Total Amount =
SUM ( fact_financials[Amount] )

_Line =
// Helper pattern. Not a measure to expose; shown so the ones below read clearly.
// Every line measure filters fact_financials to a single account.
BLANK ()

Total Revenue =
CALCULATE ( [Total Amount], dim_account[LineItem] = "Total revenue" )

Total Expenses =
CALCULATE ( [Total Amount], dim_account[LineItem] = "Total expenses" )

Contributions =
CALCULATE ( [Total Amount], dim_account[LineItem] = "Contributions and grants" )

Program Service Revenue =
CALCULATE ( [Total Amount], dim_account[LineItem] = "Program service revenue" )

Investment Income =
CALCULATE ( [Total Amount], dim_account[LineItem] = "Investment and other income" )

Program Expense =
CALCULATE ( [Total Amount], dim_account[LineItem] = "Program services" )

Management and General =
CALCULATE ( [Total Amount], dim_account[LineItem] = "Management and general" )

Fundraising Expense =
CALCULATE ( [Total Amount], dim_account[LineItem] = "Fundraising" )

Personnel Cost =
CALCULATE ( [Total Amount], dim_account[LineItem] = "Personnel" )

Net Assets =
CALCULATE ( [Total Amount], dim_account[LineItem] = "Net assets" )

Liquid Reserves =
CALCULATE ( [Total Amount], dim_account[LineItem] = "Liquid reserves" )

Cash =
CALCULATE ( [Total Amount], dim_account[LineItem] = "Cash and equivalents" )

Operating Result =
[Total Revenue] - [Total Expenses]

Operating Margin =
DIVIDE ( [Operating Result], [Total Revenue] )


// ---------------------------------------------------------------------
// 2. Stewardship ratios
// ---------------------------------------------------------------------

Program Expense Ratio =
DIVIDE ( [Program Expense], [Total Expenses] )

Administrative Ratio =
DIVIDE ( [Management and General], [Total Expenses] )

Fundraising Cost Ratio =
DIVIDE ( [Fundraising Expense], [Contributions] )

Cost to Raise a Dollar =
// Same quantity as the ratio above, formatted as currency for a card visual.
DIVIDE ( [Fundraising Expense], [Contributions] )

Fundraising Return =
DIVIDE ( [Contributions], [Fundraising Expense] )

Personnel Ratio =
DIVIDE ( [Personnel Cost], [Total Expenses] )


// ---------------------------------------------------------------------
// 3. Liquidity and reserves
// ---------------------------------------------------------------------

Monthly Expenses =
DIVIDE ( [Total Expenses], 12 )

Cash Runway Months =
DIVIDE ( [Cash], [Monthly Expenses] )

Liquid Runway Months =
DIVIDE ( [Liquid Reserves], [Monthly Expenses] )

Net Asset Multiple =
DIVIDE ( [Net Assets], [Total Expenses] )

Net Asset Months =
DIVIDE ( [Net Assets], [Monthly Expenses] )

BBB Accumulation Headroom =
// Positive means there is room below the three-times ceiling.
3 - [Net Asset Multiple]

Reserve Floor Breach =
IF ( [Liquid Runway Months] < 6, 1, 0 )


// ---------------------------------------------------------------------
// 4. Growth
// ---------------------------------------------------------------------

Revenue PY =
// dim_year[Year] is a whole number and dim_year is not a marked date table, so
// the time-intelligence functions do not apply. Offsetting the key by one is
// the equivalent, and is the pattern every prior-year measure below uses.
CALCULATE ( [Total Revenue], dim_year[Year] = SELECTEDVALUE ( dim_year[Year] ) - 1 )

Revenue Growth =
VAR Prior =
    CALCULATE ( [Total Revenue], dim_year[Year] = SELECTEDVALUE ( dim_year[Year] ) - 1 )
RETURN
    DIVIDE ( [Total Revenue] - Prior, Prior )

Expense Growth =
VAR Prior =
    CALCULATE ( [Total Expenses], dim_year[Year] = SELECTEDVALUE ( dim_year[Year] ) - 1 )
RETURN
    DIVIDE ( [Total Expenses] - Prior, Prior )

Contribution Growth =
VAR Prior =
    CALCULATE ( [Contributions], dim_year[Year] = SELECTEDVALUE ( dim_year[Year] ) - 1 )
RETURN
    DIVIDE ( [Contributions] - Prior, Prior )

Growth Gap =
// Revenue growth less expense growth. Positive is what built the reserve.
[Revenue Growth] - [Expense Growth]

Revenue CAGR =
VAR FirstYear = MIN ( dim_year[Year] )
VAR LastYear  = MAX ( dim_year[Year] )
VAR Periods   = LastYear - FirstYear
VAR StartVal  = CALCULATE ( [Total Revenue], dim_year[Year] = FirstYear )
VAR EndVal    = CALCULATE ( [Total Revenue], dim_year[Year] = LastYear )
RETURN
    IF ( Periods > 0 && StartVal > 0, ( EndVal / StartVal ) ^ ( 1 / Periods ) - 1 )

Expense CAGR =
VAR FirstYear = MIN ( dim_year[Year] )
VAR LastYear  = MAX ( dim_year[Year] )
VAR Periods   = LastYear - FirstYear
VAR StartVal  = CALCULATE ( [Total Expenses], dim_year[Year] = FirstYear )
VAR EndVal    = CALCULATE ( [Total Expenses], dim_year[Year] = LastYear )
RETURN
    IF ( Periods > 0 && StartVal > 0, ( EndVal / StartVal ) ^ ( 1 / Periods ) - 1 )

Cumulative Operating Result =
CALCULATE (
    [Operating Result],
    FILTER ( ALLSELECTED ( dim_year ), dim_year[Year] <= MAX ( dim_year[Year] ) )
)


// ---------------------------------------------------------------------
// 5. Program economics
// ---------------------------------------------------------------------

Total Participants =
SUM ( fact_program[Participants] )

Total Completers =
SUM ( fact_program[Completers] )

Total Outcomes =
SUM ( fact_program[Outcomes] )

Program Cost =
SUM ( fact_program[ProgramCost] )

Staff Hours =
SUM ( fact_program[StaffHours] )

Earnings Gain =
SUM ( fact_program[EarningsGain] )

Completion Rate =
DIVIDE ( [Total Completers], [Total Participants] )

Outcome Rate =
DIVIDE ( [Total Outcomes], [Total Participants] )

Outcome Rate of Completers =
DIVIDE ( [Total Outcomes], [Total Completers] )

Cost per Participant =
DIVIDE ( [Program Cost], [Total Participants] )

Cost per Completer =
DIVIDE ( [Program Cost], [Total Completers] )

Cost per Outcome =
DIVIDE ( [Program Cost], [Total Outcomes] )

Staff Hours per Outcome =
DIVIDE ( [Staff Hours], [Total Outcomes] )

Earnings Gain per Dollar =
DIVIDE ( [Earnings Gain], [Program Cost] )

Average Gain per Outcome =
DIVIDE ( [Earnings Gain], [Total Outcomes] )

Capacity Utilisation =
DIVIDE ( [Total Participants], SUM ( dim_program[MaxParticipants] ) )

Indirect Cost Share =
DIVIDE ( SUM ( fact_program[IndirectCost] ), [Program Cost] )

Cost per Outcome Rank =
RANKX ( ALLSELECTED ( dim_program[Program] ), [Cost per Outcome], , ASC, DENSE )


// ---------------------------------------------------------------------
// 6. Budget versus actual
// ---------------------------------------------------------------------

Total Budget =
SUM ( fact_budget_variance[Budget] )

Total Actual =
SUM ( fact_budget_variance[Actual] )

Total Variance =
[Total Actual] - [Total Budget]

Variance Pct =
DIVIDE ( [Total Variance], ABS ( [Total Budget] ) )

Variance Direction =
// Expense lines are favourable when actual comes in below budget.
VAR IsExpense = SELECTEDVALUE ( dim_account[IsExpense] )
RETURN
    SWITCH (
        TRUE (),
        [Total Variance] = 0, "On plan",
        IsExpense = 1 && [Total Variance] < 0, "Favourable",
        IsExpense = 1 && [Total Variance] > 0, "Unfavourable",
        [Total Variance] > 0, "Favourable",
        "Unfavourable"
    )

Variance Colour =
SWITCH ( [Variance Direction], "Favourable", "#1B5E20", "Unfavourable", "#B71C1C", "#5F6B7A" )

Absolute Variance =
ABS ( [Total Variance] )


// ---------------------------------------------------------------------
// 7. Funding mix and concentration
// ---------------------------------------------------------------------

Funding Amount =
SUM ( fact_funding[Amount] )

Total Donors =
SUM ( fact_funding[Donors] )

Average Gift =
DIVIDE ( [Funding Amount], [Total Donors] )

Funding Share =
DIVIDE (
    [Funding Amount],
    CALCULATE ( [Funding Amount], ALL ( dim_funding_source ) )
)

Institutional Funding Share =
DIVIDE (
    CALCULATE (
        [Funding Amount],
        dim_funding_source[FundingSource] IN { "Foundation grants", "Government funding" }
    ),
    CALCULATE ( [Funding Amount], ALL ( dim_funding_source ) )
)

Largest Source Share =
VAR Amounts =
    ADDCOLUMNS (
        ALLSELECTED ( dim_funding_source[FundingSource] ),
        "@amt", [Funding Amount]
    )
// Top and Total look like ordinary variable names, but the DAX query parser
// treats them as reserved and rejects the definition with "The syntax for
// 'Top' is incorrect" - an error the measure editor in Desktop never raises,
// because it parses an expression rather than a query.
VAR Largest = MAXX ( Amounts, [@amt] )
VAR AllSources = SUMX ( Amounts, [@amt] )
RETURN
    DIVIDE ( Largest, AllSources )


// ---------------------------------------------------------------------
// 8. Scenario selection
// ---------------------------------------------------------------------

Selected Scenario =
SELECTEDVALUE ( dim_scenario[Scenario], "All scenarios" )

Scenario Narrative =
SELECTEDVALUE ( dim_scenario[Narrative] )

Base Case Result =
CALCULATE ( [Operating Result], dim_scenario[Scenario] = "Base" )

Result vs Base =
[Operating Result] - [Base Case Result]

Scenario Spread =
VAR Up   = CALCULATE ( [Operating Result], dim_scenario[Scenario] = "Upside" )
VAR Down = CALCULATE ( [Operating Result], dim_scenario[Scenario] = "Downside" )
RETURN
    Up - Down

First Deficit Year =
CALCULATE (
    MIN ( dim_year[Year] ),
    FILTER (
        ALLSELECTED ( dim_year ),
        CALCULATE ( [Operating Result] ) < 0
    )
)


// ---------------------------------------------------------------------
// 9. Funding sensitivity  (standalone table)
// ---------------------------------------------------------------------

Sensitivity Final Result =
SUM ( fact_sensitivity[FinalYearResult] )

Sensitivity Structural Result =
SUM ( fact_sensitivity[StructuralResult] )

Sensitivity Runway =
AVERAGE ( fact_sensitivity[LiquidRunwayMonths] )

Sensitivity Outcomes =
SUM ( fact_sensitivity[Outcomes] )

Sensitivity Outcomes Lost =
SUM ( fact_sensitivity[OutcomesLost] )

Sensitivity Capacity Retained =
AVERAGE ( fact_sensitivity[CapacityRetained] )

Years to Reserve Floor =
AVERAGE ( fact_sensitivity[YearsToReserveFloor] )

Max Sustainable Shock =
CALCULATE (
    MAX ( fact_sensitivity[FundingShock] ),
    fact_sensitivity[BreachesFloor] = 0
)


// ---------------------------------------------------------------------
// 10. Resource allocation  (standalone table)
// ---------------------------------------------------------------------

Total Allocation =
SUM ( fact_allocation[Allocation] )

Additional Outcomes =
SUM ( fact_allocation[AdditionalOutcomes] )

Additional Earnings =
SUM ( fact_allocation[AdditionalEarnings] )

Allocation Avg Gain =
DIVIDE ( [Additional Earnings], [Additional Outcomes] )

Allocation Cost per Outcome =
DIVIDE ( [Total Allocation], [Additional Outcomes] )

Outcomes Forgone vs Best =
VAR Best =
    MAXX (
        ALLSELECTED ( fact_allocation[Objective] ),
        CALCULATE ( SUM ( fact_allocation[AdditionalOutcomes] ) )
    )
RETURN
    Best - [Additional Outcomes]


// ---------------------------------------------------------------------
// 11. Benchmark status  (drives conditional formatting and KPI cards)
// ---------------------------------------------------------------------

Program Ratio Status =
IF ( [Program Expense Ratio] >= 0.65, "Meets BBB Standard 8", "Below BBB Standard 8" )

Program Ratio Colour =
IF ( [Program Expense Ratio] >= 0.65, "#1B5E20", "#B71C1C" )

Fundraising Ratio Status =
IF ( [Fundraising Cost Ratio] <= 0.35, "Meets BBB Standard 9", "Above BBB Standard 9" )

Accumulation Status =
SWITCH (
    TRUE (),
    [Net Asset Multiple] > 3, "Above BBB Standard 10 ceiling",
    [Net Asset Multiple] > 2, "Approaching the ceiling",
    "Within range"
)

Accumulation Colour =
SWITCH (
    TRUE (),
    [Net Asset Multiple] > 3, "#B71C1C",
    [Net Asset Multiple] > 2, "#9A3412",
    "#1B5E20"
)

Runway Status =
SWITCH (
    TRUE (),
    [Liquid Runway Months] < 3,  "Below minimum reserve",
    [Liquid Runway Months] < 6,  "Below target reserve",
    [Liquid Runway Months] > 24, "Above the reserve ceiling",
    "Within target range"
)

Runway Colour =
SWITCH (
    TRUE (),
    [Liquid Runway Months] < 6,  "#B71C1C",
    [Liquid Runway Months] > 24, "#9A3412",
    "#1B5E20"
)


// ---------------------------------------------------------------------
// 12. Report furniture
// ---------------------------------------------------------------------

Period Type =
SELECTEDVALUE ( dim_year[PeriodType], "Mixed" )

Forecast Flag =
IF ( SELECTEDVALUE ( dim_year[PeriodType] ) <> "Actual", "Forecast - analyst assumptions", "" )

Data Provenance Note =
"Historical figures are from IRS Form 990 filings. Program-level detail and all "
    & "forward-looking figures are analyst assumptions and are not the organization's guidance."

Selected Year Label =
SELECTEDVALUE ( dim_year[YearLabel], "All years" )
"""


def write_dax():
    ROOT.mkdir(parents=True, exist_ok=True)
    p = ROOT / "measures.dax"
    p.write_text(DAX, encoding="utf-8")
    n = len(_parse_measures(DAX))
    print(f"  {'measures.dax':<32} {n:>5} measures")
    return p


# --------------------------------------------------------------------------
# The same library in DAX query view form.
#
# Power BI Desktop's New measure box takes one definition at a time, which is
# what measures.dax is shaped for. The browser has no Tabular Editor, and a
# hundred measures pasted one at a time is an hour of clicking. DAX query view
# (in the Service: open the semantic model > Write DAX queries) accepts a DEFINE
# block and offers an "Update model: Add new measures" action that creates all
# of them at once.
#
# Home tables are cosmetic - they decide only where a measure sits in the Data
# pane. They are resolved to the fact table a measure actually reads, following
# measure-to-measure references where the body names no table itself.
# --------------------------------------------------------------------------
# Column-0 tokens that open a statement inside a measure body rather than
# starting a new measure.
DAX_KEYWORDS = ("VAR ", "RETURN", "EVALUATE", "DEFINE", "MEASURE ", "COLUMN ",
                "TABLE ", "ORDER ")

FACT_TABLES = [
    "fact_financials", "fact_program", "fact_funding", "fact_budget_variance",
    "fact_allocation", "fact_sensitivity", "fact_quality_tradeoff",
    "fact_provenance",
]


def _parse_measures(text: str) -> list[tuple[str, list[str]]]:
    """Split the library into (name, body-lines). Skips the _Line placeholder."""
    import re
    lines = text.splitlines()
    hdr = re.compile(r"^([A-Za-z_][A-Za-z0-9 _%\-.()/&]*?)\s*=\s*$")
    # A measure header is a bare "Name =" in column 0. So is the "VAR Thing ="
    # of a multi-statement measure, which is why the keywords are excluded -
    # without this, a measure using VAR is silently split in two and the half
    # after the VAR becomes a measure named "VAR Thing".
    heads = [(i, m.group(1)) for i, ln in enumerate(lines)
             if (m := hdr.match(ln)) and not ln.startswith(DAX_KEYWORDS)]
    out = []
    for k, (i, name) in enumerate(heads):
        end = heads[k + 1][0] if k + 1 < len(heads) else len(lines)
        body = lines[i + 1:end]
        # Trailing blanks and the next group's comment banner belong to neither.
        while body and (not body[-1].strip() or body[-1].lstrip().startswith("//")):
            body.pop()
        if name != "_Line":
            out.append((name, body))
    return out


def _home_tables(items: list[tuple[str, list[str]]]) -> dict[str, str]:
    """Resolve each measure to the fact table it reads.

    A measure that names a fact table directly is settled. One that only calls
    other measures inherits from them, so [Total Revenue] - which filters
    dim_account but sums through [Total Amount] - lands on fact_financials
    rather than on the dimension it happens to mention.
    """
    import re
    from collections import Counter
    names = {n for n, _ in items}
    home: dict[str, str] = {}
    pending = []
    for name, body in items:
        txt = "\n".join(body)
        direct = next((t for t in FACT_TABLES if t in txt), None)
        if direct:
            home[name] = direct
        else:
            pending.append((name, [r for r in re.findall(r"\[([^\]]+)\]", txt)
                                   if r in names and r != name]))
    for _ in range(len(pending) + 1):
        if not any(n not in home for n, _ in pending):
            break
        for name, refs in pending:
            if name in home:
                continue
            known = [home[r] for r in refs if r in home]
            if known:
                home[name] = Counter(known).most_common(1)[0][0]
    for name, _ in items:
        home.setdefault(name, "fact_financials")
    return home


def write_dax_query():
    items = _parse_measures(DAX)
    home = _home_tables(items)
    out = [
        "// =====================================================================",
        "// Nonprofit Financial Planning - DAX measure library, query view form",
        "//",
        "// For the Power BI Service in a browser. Open the semantic model, choose",
        "// Write DAX queries, paste this whole file, then click the",
        "// 'Update model: Add new measures' link that appears above DEFINE.",
        f"// All {len(items)} measures are created in one action.",
        "//",
        "// The home table on each MEASURE line is cosmetic - it decides where the",
        "// measure appears in the Data pane. Measures evaluate model-wide.",
        "//",
        "// Create the twelve relationships first. Many of these filter on",
        "// dimension columns and return errors until the model is wired.",
        "//",
        "// measures.dax holds the same library in the one-at-a-time paste form",
        "// that Power BI Desktop's New measure box expects.",
        "// =====================================================================",
        "",
        "DEFINE",
    ]
    for name, body in items:
        out.append(f"    MEASURE '{home[name]}'[{name}] =")
        out.extend(("        " + b) if b.strip() else "" for b in body)
        out.append("")
    out += ["EVALUATE", f'    ROW ( "Measures added", {len(items)} )', ""]
    ROOT.mkdir(parents=True, exist_ok=True)
    p = ROOT / "measures_dax_query.dax"
    p.write_text("\n".join(out), encoding="utf-8")
    print(f"  {'measures_dax_query.dax':<32} {len(items):>5} measures")
    return p



# --------------------------------------------------------------------------
# A query that checks the deployed model against the filed figures.
#
# Everything up to here verifies the files. This verifies the thing Power BI
# actually built from them - which is not the same claim, as four separate
# failures in this project demonstrated. It runs in DAX query view against the
# live semantic model and returns one row per check with the difference.
# --------------------------------------------------------------------------
def write_verify_query():
    """
    Expected values come from the rows this build actually ships, not from the
    engine's continuous figures. The two differ slightly by design - outcomes
    are whole people once rounded per program, so the engine's 1,494.8 is 1,494
    in the fact table - and a check that quoted the engine would fail against a
    model that is behaving correctly. Tying the shipped rows back to the filed
    990 is a separate job, and the test suite already does it.
    """
    rows = dict(
        (name, (header, data)) for name, header, data in TABLES
    )
    acc_h, acc_r = rows["dim_account"]
    line_of = {r[acc_h.index("AccountKey")]: r[acc_h.index("LineItem")] for r in acc_r}
    fin_h, fin_r = rows["fact_financials"]
    fin = {}
    for r in fin_r:
        if str(r[fin_h.index("Year")]) == "2024":
            fin[line_of[r[fin_h.index("AccountKey")]]] = float(r[fin_h.index("Amount")])
    pg_h, pg_r = rows["fact_program"]
    pg = {c: sum(float(r[pg_h.index(c)]) for r in pg_r
                 if str(r[pg_h.index("Year")]) == "2024")
          for c in ("Participants", "Outcomes", "ProgramCost")}

    expenses = fin["Total expenses"]
    checks = [
        ("Total revenue",          "Total Revenue",          fin["Total revenue"]),
        ("Total expenses",         "Total Expenses",         expenses),
        ("Operating result",       "Operating Result",       fin["Total revenue"] - expenses),
        ("Program expense ratio",  "Program Expense Ratio",  fin["Program services"] / expenses),
        ("Administrative ratio",   "Administrative Ratio",   fin["Management and general"] / expenses),
        ("Fundraising cost ratio", "Fundraising Cost Ratio", fin["Fundraising"] / fin["Contributions and grants"]),
        ("Net assets",             "Net Assets",             fin["Net assets"]),
        ("Net asset multiple",     "Net Asset Multiple",     fin["Net assets"] / expenses),
        ("Liquid runway months",   "Liquid Runway Months",   fin["Liquid reserves"] / (expenses / 12)),
        ("Participants",           "Total Participants",     pg["Participants"]),
        ("Outcomes",               "Total Outcomes",         pg["Outcomes"]),
        ("Cost per outcome",       "Cost per Outcome",       pg["ProgramCost"] / pg["Outcomes"]),
    ]

    out = [
        "// =====================================================================",
        "// Nonprofit Financial Planning - model verification query",
        "//",
        "// Run this in DAX query view once the measures are created. It checks",
        "// the deployed semantic model against the FY2024 figures and returns",
        "// one row per check.",
        "//",
        "// Every row should read OK and every Difference should be 0. A whole",
        "// column of wrong numbers points at the relationships; a single wrong",
        "// row points at that measure.",
        "//",
        "// FY2024 is an actual, so it carries one scenario and needs no scenario",
        "// filter. Forecast years hold Downside, Base and Upside side by side, so",
        "// a visual spanning those years without a scenario slicer sums all three",
        "// and reads about three times too high. That is the model working as",
        "// designed, not a defect - but it is the first thing to suspect when a",
        "// forecast number looks implausible.",
        "// =====================================================================",
        "",
        "EVALUATE",
        "VAR Yr = 2024",
        "VAR Tol = 0.000001",
    ]
    for i, (_l, measure, _e) in enumerate(checks):
        out.append(f"VAR v{i} = CALCULATE ( [{measure}], dim_year[Year] = Yr )")
    out += ["RETURN", "    UNION ("]
    blocks = []
    for i, (label, _m, exp) in enumerate(checks):
        e = repr(float(exp))
        blocks.append(
            "        ROW (\n"
            f'            "Check", "{label}",\n'
            f'            "Expected", {e},\n'
            f'            "Power BI", v{i},\n'
            f'            "Difference", v{i} - {e},\n'
            f'            "Status", IF ( ABS ( v{i} - {e} ) <= Tol * ABS ( {e} ) + 0.005, "OK", "CHECK" )\n'
            "        )"
        )
    out.append(",\n".join(blocks))
    out += ["    )", ""]
    ROOT.mkdir(parents=True, exist_ok=True)
    path = ROOT / "verify_model.dax"
    path.write_text("\n".join(out), encoding="utf-8")
    print(f"  {'verify_model.dax':<32} {len(checks):>5} checks")
    return path

# ==========================================================================
def write_build_guide():
    nm = len(_parse_measures(DAX))
    hist = engine.historical()["rows"]
    h24 = hist[2024]
    guide = f"""# Power BI build guide

A `.pbix` file is a proprietary binary and cannot be generated programmatically,
so this folder contains everything that goes inside one: the data model, the
measures, and the assembly spec. Importing the workbook, wiring the twelve
relationships and adding the measures reproduces the dashboard.

Power BI Desktop is Windows-only, so on macOS the route is the Power BI service
in a browser at app.powerbi.com. Everything below can be done there. Import
`PowerBI_Data_Model.xlsx`; the sixteen CSVs in `data/` hold identical data and
are for the Desktop folder connector.

---

## 1. Load the data

### In the browser (Power BI service) — recommended on macOS

Use **`PowerBI_Data_Model.xlsx`**. It holds all sixteen tables as named Excel
Tables, one per sheet, so the whole model imports in a single step. The browser
imports one file at a time, which is why this workbook exists.

1. Go to app.powerbi.com and open a workspace (My Workspace is fine).
2. **New > Semantic model**, or **Get data > Files > Local File**.
3. Choose **Import** — *not* Upload. Upload opens the workbook in Excel Online
   and gives you nothing to build a report on; Import creates the semantic model
   you actually need. This is the single most common wrong turn here.
4. Tick the sixteen tables. Leave `_README` unticked — it is a plain sheet
   rather than an Excel Table, so it appears under *Sheets* and is easy to skip.
5. Then go to step 2 below. **The import does not create relationships** — Power
   BI never infers them from a workbook, so all twelve must be drawn by hand.

Creating relationships and measures in the browser needs write access to the
semantic model, which you have on a model you just created in your own
workspace. If the modelling view is greyed out, the model is being viewed rather
than edited — open it from the workspace list and choose **Open data model**.

### In Power BI Desktop (Windows only)

Use the CSVs instead: **Get data > Folder**, point at `data/`, and load all
sixteen in one step. Identical data — the CSVs are kept because they diff
properly in version control, where a binary `.xlsx` shows only "changed".

### What you are loading

| Table | Grain | Rows |
|---|---|---|
| `dim_year` | one row per fiscal year | actual and forecast years |
| `dim_scenario` | one row per scenario | Actual, Downside, Base, Upside |
| `dim_program` | one row per program | 4 |
| `dim_funding_source` | one row per funding category | 5 |
| `dim_account` | one row per financial statement line | 17 |
| `dim_benchmark` | one row per published benchmark | 6 |
| `dim_provenance` | one row per provenance tier | 4 |
| `fact_financials` | year x scenario x account | long format |
| `fact_program` | year x scenario x program | |
| `fact_funding` | year x scenario x funding source | |
| `fact_budget_variance` | FY2024 x account | |
| `fact_variance_bridge` | one row per waterfall step | 4 |
| `fact_sensitivity` | scenario x funding shock | |
| `fact_allocation` | objective x program | |
| `fact_quality_tradeoff` | one row per quality floor tested | |
| `fact_provenance` | one row per tier | 4 |

Check the data types after loading. `Year` should be a whole number; `Amount`,
`VariancePct`, `Share`, `CapacityRetained` and `FundingShock` should be decimal
numbers, not text. The workbook writes genuine numeric cells and leaves missing
values blank rather than empty strings, precisely so type inference does not
silently turn a numeric column into text — but it is worth a glance, because
every measure downstream depends on it.

## 2. Build the relationships

All single-direction, many-to-one, from fact to dimension. In the browser this
is **Open data model**, then drag the fact column onto the dimension column. In
Desktop it is the Model view.

```
fact_financials[Year]            ->  dim_year[Year]
fact_financials[ScenarioKey]     ->  dim_scenario[ScenarioKey]
fact_financials[AccountKey]      ->  dim_account[AccountKey]

fact_program[Year]               ->  dim_year[Year]
fact_program[ScenarioKey]        ->  dim_scenario[ScenarioKey]
fact_program[ProgramKey]         ->  dim_program[ProgramKey]

fact_funding[Year]               ->  dim_year[Year]
fact_funding[ScenarioKey]        ->  dim_scenario[ScenarioKey]
fact_funding[SourceKey]          ->  dim_funding_source[SourceKey]

fact_budget_variance[Year]       ->  dim_year[Year]
fact_budget_variance[AccountKey] ->  dim_account[AccountKey]

fact_allocation[ProgramKey]      ->  dim_program[ProgramKey]
```

`fact_sensitivity` and `fact_quality_tradeoff` are deliberately left
disconnected. They are scenario-shock cross-sections rather than time series, and
wiring them into the star would create ambiguous filter paths. Slice them on
their own columns.

Sort `dim_year` by `SortOrder`, `dim_scenario` by `SortOrder`, `dim_program` by
`SortOrder`, `dim_account` by `SortOrder`, `dim_funding_source` by `SortOrder`
(Column tools > Sort by column). Without this, everything sorts alphabetically
and the statement lines come out in the wrong order.

## 3. Add the measures

{nm} measures in twelve groups, supplied in two forms. Use the one that matches
where you are building.

**In the browser — `measures_dax_query.dax`, all at once.** Open the semantic
model, choose **Write DAX queries**, paste the whole file, and click the
**Update model: Add new measures** link that appears above `DEFINE`. Every
measure is created in one action. This is the only sane route in the Service:
the web modelling canvas has a **New measure** button but no bulk import, and
{nm} definitions pasted one at a time is an hour of clicking.

**In Power BI Desktop — `measures.dax`, one at a time.** Create a blank measure
and paste a definition (the name above the `=`, the expression below). Tabular
Editor imports the file wholesale, but it is Windows-only.

The two files hold the same library. The query-view form adds a home table to
each definition, which decides only where the measure sits in the Data pane;
measures evaluate across the whole model regardless of where they live.

Nine measures carry a `Total ` prefix — `[Total Amount]`, `[Total Participants]`,
`[Total Budget]` and so on — because a measure may not share its name with a
column in the same table, and each of those aggregates a column of the bare
name. Renaming one back will fail on create.

Check group 1 before trusting anything else. Almost every measure downstream is
built from `[Total Amount]` and the line measures beneath it. `[Total Revenue]`
should read **$28,580,411** for FY2024 with no scenario filter applied.

Better than checking one measure by eye: open a new query tab and run
`verify_model.dax`. It returns twelve rows comparing the deployed model against
the FY2024 figures, with a Difference and an OK / CHECK on each. Twelve OKs
means the load, the relationships and the base measures are all correct. A whole
column of wrong numbers points at the relationships; one wrong row points at
that measure.

Ignore the `_Line` placeholder in `measures.dax`. It documents the pattern the
line measures use and is not meant to be created; the query-view file omits it.

Changes to a model edited in the browser save automatically with no undo, so if
something goes badly wrong use the semantic model's version history rather than
trying to unpick it by hand.

## 4. Build the pages

### The scenario slicer, before anything else

`dim_scenario` has four members and they do not overlap in time. **Actual**
exists only for FY2019-FY2024, because filed results are not a scenario.
**Downside**, **Base** and **Upside** exist only for FY2025-FY2029.

So a scenario slicer set to Base alone shows nothing before FY2025, and the
FY2024 cards go blank - `[Total Revenue]` has no Base row to sum. Select
**Actual and Base together** for any page that spans history and forecast. No
year holds both, so there is exactly one series per year and nothing is double
counted.

Leaving the slicer unset is worse than wrong: the forecast years sum all three
scenarios and read about three times too high, with no error to warn you.

Page 5 inverts this. There the three scenarios are the subject, so select
Downside, Base and Upside and put `dim_scenario[Scenario]` on the legend. Leave
Actual off - it would draw a fourth line that stops at FY2024.

Set the slicer to allow multiple selections: Format > Slicer settings >
Selection, with **Single select** off and **Multi-select with CTRL** off, so
plain clicks toggle each member.

### Page 1 - Financial overview

Purpose: can a director see the shape of the organization in ten seconds.

- KPI card row: `Total Revenue`, `Total Expenses`, `Operating Result`,
  `Program Expense Ratio`, `Liquid Runway Months`, `Net Asset Multiple`
- Line chart: `Total Revenue` and `Total Expenses` by `dim_year[YearLabel]`
- Column chart: `Operating Result` by year, conditionally coloured on sign
- Line chart: `Liquid Runway Months` by year, with reference lines at 3, 6 and 24
  months from `dim_benchmark`
- Slicers: `dim_year[PeriodType]`, `dim_scenario[Scenario]`
- Text box bound to `Data Provenance Note`

Set the KPI cards' font colour with `Program Ratio Colour`, `Runway Colour` and
`Accumulation Colour` (Format > Callout value > fx > Field value).

### Page 2 - Budget performance

- Matrix: rows `dim_account[LineItem]`, values `Total Budget`, `Total Actual`,
  `Total Variance`, `Variance Pct`, conditionally formatted on
  `Variance Colour`
- Waterfall: `fact_variance_bridge`, category `fact_variance_bridge[Step]`,
  value `fact_variance_bridge[Amount]`, breakdown `fact_variance_bridge[StepType]`
  - this is the volume-versus-rate decomposition
- Card: `Variance Direction` for the selected line
- Text box for the commentary from `03_Budget_vs_Actual.xlsx`

The point of this page is the waterfall. A variance table tells a reader that
program cost was over budget; the waterfall tells them {abs(engine.budget_vs_actual()["decomposition"]["volume_variance"]) / abs(engine.budget_vs_actual()["decomposition"]["total_program_variance"]):.0%}
of the overrun was serving more people and the rest was unit cost.

### Page 3 - Programs

- Table: `dim_program[Program]` with `Total Participants`, `Total Outcomes`,
  `Outcome Rate`, `Cost per Participant`, `Cost per Outcome`,
  `Earnings Gain per Dollar`
- Scatter: X `Cost per Outcome`, Y `Average Gain per Outcome`, size
  `Program Cost`, legend `dim_program[Program]`. This single visual carries the
  whole finding - the cheap programs and the high-value programs are different
  programs.
- Bar: `Capacity Utilisation` by program, with a reference line at 100%
- Slicer: `dim_year[Year]`, `dim_scenario[Scenario]`

### Page 4 - Forecast

- Line chart: `Total Revenue`, `Total Expenses` by year, split actual and
  forecast using `dim_year[PeriodType]` on the legend
- Line chart: `Liquid Reserves` by year
- Card: `First Deficit Year` - the crossover into structural deficit
- Line chart: `Total Participants` and `Total Outcomes` by year

### Page 5 - Scenarios

- Line chart: `Operating Result` by year, legend `dim_scenario[Scenario]`
- Line chart: `Liquid Runway Months` by year, legend scenario, reference line at 6
- Card: `Scenario Spread`
- Table: scenario, `Total Revenue`, `Total Expenses`, `Operating Result`,
  `Total Outcomes` for the final forecast year
- Text box bound to `Scenario Narrative`

### Page 6 - Funding and sensitivity

- Donut or bar: `Funding Amount` by `dim_funding_source[FundingSource]`
- Card: `Institutional Funding Share` - the concentration measure that matters
  most for this organization
- Line chart from `fact_sensitivity`: `Sensitivity Runway` by
  `fact_sensitivity[FundingShock]`, legend `fact_sensitivity[Scenario]`,
  reference line at 6 months
- Column chart: `Sensitivity Outcomes Lost` by `fact_sensitivity[FundingShock]`
- Card: `Max Sustainable Shock`

### Page 7 - Allocation

- Clustered bar from `fact_allocation`: `Total Allocation` by program, legend
  `fact_allocation[ObjectiveLabel]`
- Table: `fact_allocation[ObjectiveLabel]`, `Additional Outcomes`,
  `Additional Earnings`, `Allocation Avg Gain`, `Outcomes Forgone vs Best`
- Line chart from `fact_quality_tradeoff`: the column
  `fact_quality_tradeoff[Outcomes]` by `fact_quality_tradeoff[QualityFloor]` -
  this table is disconnected, so use its own columns, not `[Total Outcomes]`

## 5. Formatting

Number formats: currency with no decimals for amounts, one decimal for
percentages, one decimal plus a " mo" suffix for runway measures, two decimals
plus an "x" suffix for multiples.

Keep the theme restrained. This is a management report, not a marketing asset:
one accent colour, one alert colour, and everything else in greys. Red should
mean a threshold has been breached and nothing else.

## 6. What to put on the page that is not a number

Every page needs a visible provenance note. The historical figures are real and
the forward-looking ones are not, and a dashboard that presents both in the same
visual style without saying so is misleading regardless of what the appendix
says. `Forecast Flag` returns a warning string whenever the selected period is
not an actual; bind it to a text box on every page that shows forecast data.

---

## Troubleshooting

**"DataFormat.Error: We were unable to load this Excel file because we couldn't
understand its format. File contains corrupted data."**

This is a parser error, not a data error - Power Query could not read the file
at all, which is why it names no sheet, no column and no cell. The file is not
corrupt in any ordinary sense; it opens fine in Excel.

The cause, if you regenerate the workbook yourself: `openpyxl` switches to a
faster writer whenever `lxml` is installed, and that writer emits numeric cells
as `t="n"` and strings as inline runs with no shared string table. Both are
legal OOXML and neither is what Excel writes, and Power Query refuses the file.
`src/build_08_powerbi.py` therefore writes this workbook with `xlsxwriter`,
which produces Excel-shaped output. Seven tests in `tests/` pin that shape, so
a regression fails the suite rather than reaching Power BI.

If you hit it anyway, fall back to the sixteen CSVs in `data/` - same data, and
CSV has no format to misread. Import them one at a time, or use Power BI
Desktop's folder connector if you have Windows access.

**The Navigator lists `dim_year1`, `dim_scenario2`, `fact_financials8` …**

You are using a workbook generated before this was fixed. Excel permits a sheet
and a table to share a name, but Power Query then appends the table's id to tell
them apart - and `dim_year1` breaks every measure written against `dim_year`.

Regenerate the workbook (`python src/build_08_powerbi.py`), or import the
un-numbered items instead: those are the sheets, and they carry the correct
names. In the current workbook the sheets are prefixed `01 `, `02 ` and so on,
so there is no collision and the Tables import under their own names.

**Measures return blank after loading.**

Almost always a missing relationship rather than a broken measure. The import
does not create relationships - check all twelve from section 2 exist, and
that each points from the fact table to the dimension and not the reverse.

## Reference figures

These come from the FY2024 filing and should match what the dashboard shows once
built. If they do not, something is wrong with the load or the relationships.

| Measure | FY2024 |
|---|---|
| Total revenue | ${h24['revenue']:,.0f} |
| Total expenses | ${h24['expenses']:,.0f} |
| Operating result | ${h24['operating_result']:,.0f} |
| Program expense ratio | {h24['program_ratio']:.1%} |
| Administrative ratio | {h24['admin_ratio']:.1%} |
| Fundraising cost ratio | {h24['fundraising_cost_ratio']:.1%} |
| Net assets | ${h24['net_assets']:,.0f} |
| Net asset multiple | {h24['net_asset_multiple']:.2f}x |
| Liquid runway | {h24['liquid_runway_months']:.1f} months |
| Participants | {h24['participants']:,.0f} |
| Outcomes | {h24['placements']:,.0f} |
| Cost per outcome | ${h24['cost_per_outcome']:,.0f} |
"""
    p = ROOT / "BUILD_GUIDE.md"
    p.write_text(guide, encoding="utf-8")
    print(f"  BUILD_GUIDE.md")
    return p


def build():
    print("Power BI package:")
    TABLES.clear()
    DATA.mkdir(parents=True, exist_ok=True)
    dim_year(); dim_scenario(); dim_program(); dim_funding_source()
    dim_account(); dim_provenance(); fact_benchmarks()
    fact_financials(); fact_program(); fact_funding()
    fact_budget_variance(); fact_variance_bridge()
    fact_sensitivity(); fact_allocation(); fact_tradeoff()
    fact_provenance_summary()
    write_workbook()
    write_dax()
    write_dax_query()
    write_verify_query()
    write_build_guide()
    print(f"wrote {ROOT}")
    return ROOT


if __name__ == "__main__":
    build()
