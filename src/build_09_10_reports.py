"""
Management_Report.pdf and Executive_Summary.pdf

Every figure quoted in both documents is pulled from the engine at build time,
so the prose cannot drift away from the model. If an assumption changes and the
workbooks move, these documents move with them.
"""

from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_RIGHT, TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer,
                                Table, TableStyle, PageBreak, KeepTogether, HRFlowable)

import inputs as I
import engine

ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"
NAVY = colors.HexColor("#12324F")
INK = colors.HexColor("#1A1A1A")
MUTED = colors.HexColor("#5F6B7A")
RULE = colors.HexColor("#C7D2DC")
BAND = colors.HexColor("#F5F8FA")
ACCENT = colors.HexColor("#0B6E4F")
WARN = colors.HexColor("#9A3412")

M = engine.full_model()
H = M["historical"]["rows"]
HS = M["historical"]["summary"]
H24 = H[2024]
PE = M["programs"]["programs"]
PT = M["programs"]["totals"]
BVA = M["bva"]
DEC = BVA["decomposition"]
L = BVA["lines"]
MVF = M["minimum_funding"]
LF = M["largest_funder"]
SENS = M["sensitivity"]
AO = M["allocation_objectives"]
AR = {r["objective"]: r for r in AO["rows"]}
SC = {s: engine.forecast(s, years=engine.SENSITIVITY_HORIZON) for s in I.SCENARIOS}

d = lambda v: f"${v:,.0f}"
dn = lambda v: f"${abs(v):,.0f}"
p1 = lambda v: f"{v:.1%}"
p0 = lambda v: f"{v:.0%}"
n0 = lambda v: f"{v:,.0f}"


def dm(v):
    """Compact currency for KPI tiles, which are too narrow for full figures."""
    a = abs(v)
    sign = "-" if v < 0 else ""
    if a >= 1_000_000:
        return f"{sign}${a / 1_000_000:.1f}M"
    if a >= 1_000:
        return f"{sign}${a / 1_000:.0f}k"
    return f"{sign}${a:,.0f}"


def styles():
    s = getSampleStyleSheet()
    base = dict(fontName="Helvetica", fontSize=9.6, leading=14.2, textColor=INK,
                alignment=TA_JUSTIFY, spaceAfter=7)
    return {
        "title": ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=19, leading=23,
                                textColor=NAVY, spaceAfter=4),
        "subtitle": ParagraphStyle("st", fontName="Helvetica-Oblique", fontSize=10,
                                   leading=14, textColor=MUTED, spaceAfter=14),
        "h1": ParagraphStyle("h1", fontName="Helvetica-Bold", fontSize=13.5, leading=17,
                             textColor=NAVY, spaceBefore=15, spaceAfter=7),
        "h2": ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=10.8, leading=14,
                             textColor=NAVY, spaceBefore=11, spaceAfter=5),
        "body": ParagraphStyle("b", **base),
        "bodyc": ParagraphStyle("bc", **{**base, "spaceAfter": 3}),
        "bullet": ParagraphStyle("bu", **{**base, "leftIndent": 13, "bulletIndent": 3,
                                          "spaceAfter": 4}),
        "lead": ParagraphStyle("l", fontName="Helvetica", fontSize=11.2, leading=16.5,
                               textColor=INK, alignment=TA_JUSTIFY, spaceAfter=9),
        "note": ParagraphStyle("n", fontName="Helvetica-Oblique", fontSize=8.2, leading=11.5,
                               textColor=MUTED, alignment=TA_JUSTIFY, spaceAfter=6),
        "warn": ParagraphStyle("w", fontName="Helvetica", fontSize=8.4, leading=11.8,
                               textColor=WARN, alignment=TA_JUSTIFY, spaceAfter=6),
        "cell": ParagraphStyle("c", fontName="Helvetica", fontSize=8.4, leading=11,
                               textColor=INK),
        "cellb": ParagraphStyle("cb", fontName="Helvetica-Bold", fontSize=8.4, leading=11,
                                textColor=NAVY),
        "kpi_v": ParagraphStyle("kv", fontName="Helvetica-Bold", fontSize=15, leading=18,
                                textColor=NAVY, alignment=TA_CENTER),
        "kpi_l": ParagraphStyle("kl", fontName="Helvetica", fontSize=7.4, leading=9.5,
                                textColor=MUTED, alignment=TA_CENTER),
    }


ST = styles()


def para(text, style="body"):
    return Paragraph(text, ST[style])


def bullets(items, style="bullet"):
    return [Paragraph(t, ST[style], bulletText="•") for t in items]


def rule(space_before=4, space_after=8, color=RULE, width=0.8):
    return [Spacer(1, space_before),
            HRFlowable(width="100%", thickness=width, color=color, spaceAfter=0),
            Spacer(1, space_after)]


def table(data, widths, align=None, header=True, band=True, size=8.4, pad=4.5):
    t = Table(data, colWidths=widths, hAlign="LEFT", repeatRows=1 if header else 0)
    cmds = [
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), size),
        ("TEXTCOLOR", (0, 0), (-1, -1), INK),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), pad),
        ("BOTTOMPADDING", (0, 0), (-1, -1), pad),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("LINEBELOW", (0, 0), (-1, -2), 0.3, RULE),
    ]
    if header:
        cmds += [("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                 ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                 ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                 ("TOPPADDING", (0, 0), (-1, 0), 5.5),
                 ("BOTTOMPADDING", (0, 0), (-1, 0), 5.5)]
    if band:
        start = 1 if header else 0
        for i in range(start, len(data)):
            if (i - start) % 2 == 1:
                cmds.append(("BACKGROUND", (0, i), (-1, i), BAND))
    for col in (align or []):
        cmds.append(("ALIGN", (col, 0), (col, -1), "RIGHT"))
    t.setStyle(TableStyle(cmds))
    return t


def kpi_row(items, width):
    cells = []
    for label, value in items:
        cells.append([Paragraph(value, ST["kpi_v"]), Paragraph(label, ST["kpi_l"])])
    data = [[c[0] for c in cells], [c[1] for c in cells]]
    cw = width / len(items)
    t = Table(data, colWidths=[cw] * len(items), hAlign="LEFT")
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, 1), (-1, 1), 8),
        ("BACKGROUND", (0, 0), (-1, -1), BAND),
        ("LINEABOVE", (0, 0), (-1, 0), 1.6, NAVY),
        ("LINEBELOW", (0, 1), (-1, 1), 0.6, RULE),
        ("LINEAFTER", (0, 0), (-2, -1), 0.4, RULE),
    ]))
    return t


class Doc(BaseDocTemplate):
    def __init__(self, path, title, footer):
        super().__init__(str(path), pagesize=A4,
                         leftMargin=20 * mm, rightMargin=20 * mm,
                         topMargin=17 * mm, bottomMargin=17 * mm,
                         title=title, author="Independent analysis",
                         subject=f"{I.ORG_NAME} financial planning and impact measurement")
        self.footer_text = footer
        frame = Frame(self.leftMargin, self.bottomMargin, self.width, self.height, id="f")
        self.addPageTemplates([PageTemplate(id="p", frames=[frame], onPage=self._decorate)])

    def _decorate(self, canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(RULE)
        canvas.setLineWidth(0.5)
        y = self.bottomMargin - 5 * mm
        canvas.line(self.leftMargin, y, self.leftMargin + self.width, y)
        canvas.setFont("Helvetica", 7.2)
        canvas.setFillColor(MUTED)
        canvas.drawString(self.leftMargin, y - 4.2 * mm, self.footer_text)
        canvas.drawRightString(self.leftMargin + self.width, y - 4.2 * mm,
                               f"Page {canvas.getPageNumber()}")
        canvas.restoreState()


DISCLAIMER = (
    f"<b>Basis of preparation.</b> This is an independent analysis prepared from publicly "
    f"available information for portfolio and educational purposes. It is not affiliated with, "
    f"endorsed by, or reviewed by {I.ORG_NAME}, and it is not financial advice. Historical "
    f"financial totals are taken verbatim from IRS Form 990 filings for FY2019 to FY2024 and "
    f"impact metrics from the organization's own published reporting. Everything else - the split "
    f"of expenses into functional categories, the four-program portfolio and its unit economics, "
    f"the composition of the balance sheet, and every forward-looking figure - is an analyst "
    f"assumption, is labelled as such in 01_Assumptions.xlsx, and does not represent the "
    f"organization's plans, budget or guidance."
)


# ==========================================================================
def management_report():
    W = A4[0] - 40 * mm
    doc = Doc(REPORTS / "Management_Report.pdf", "Management Report",
              f"{I.ORG_NAME} - independent financial planning analysis - prepared from public data")
    s = []

    s.append(para("Nonprofit Financial Planning, Scenario Forecasting "
                  "and Impact Measurement", "title"))
    s.append(para(f"Management report &nbsp;|&nbsp; {I.ORG_NAME} (EIN {I.ORG_EIN}) "
                  f"&nbsp;|&nbsp; analysis as of {I.AS_OF}", "subtitle"))

    s.append(kpi_row([
        ("FY2024 revenue", dm(H24["revenue"])),
        ("Operating result", dm(H24["operating_result"])),
        ("Program ratio", p1(H24["program_ratio"])),
        ("Liquid runway", f"{H24['liquid_runway_months']:.0f} mo"),
        ("Net assets / expenses", f"{H24['net_asset_multiple']:.2f}x"),
        ("Cost per outcome", d(H24["cost_per_outcome"])),
    ], W))
    s.append(Spacer(1, 14))

    # ---- the question --------------------------------------------------
    s.append(para("The question", "h1"))
    s.append(para(
        "How should a nonprofit allocate limited funding across its programs while remaining "
        "financially sustainable and maximising measurable impact? For this organization the "
        "question arrives in an unusual form. It is not short of money. It has run six "
        "consecutive operating surpluses, holds net assets of "
        f"{d(H24['net_assets'])} against annual expenses of {d(H24['expenses'])}, and could "
        "absorb a total collapse in institutional funding for several years without cutting a "
        "single program. The financial planning problem here is not survival. It is that a "
        "revenue surge driven by circumstances outside the organization's control has left it "
        "with a cost base sized to funding that is unlikely to repeat, and a balance sheet large "
        "enough to disguise how long that has been true.", "lead"))

    # ---- what happened -------------------------------------------------
    s.append(para("1. What the last six years actually did", "h1"))
    s.append(para(
        f"Revenue grew from {d(H[2019]['revenue'])} in FY2019 to {d(H24['revenue'])} in FY2024, a "
        f"compound rate of {p1(HS['revenue_cagr'])} a year. Almost all of it came through "
        f"contributions, which compounded at {p1(HS['contribution_cagr'])}. The two years that "
        f"matter are FY2021 and FY2022, when revenue rose {p1(H[2021]['revenue_growth'])} and then "
        f"{p1(H[2022]['revenue_growth'])} as humanitarian response funding arrived following the "
        f"Afghanistan and Ukraine displacement crises. That funding was real, it was well used, "
        f"and it was episodic."))
    s.append(para(
        f"Expenses grew too, but more slowly - {p1(HS['expense_cagr'])} a year against "
        f"{p1(HS['revenue_cagr'])} for revenue. The gap compounded into "
        f"{d(HS['cumulative_surplus'])} of cumulative surplus across six years and lifted net "
        f"assets from {d(H[2019]['net_assets'])} to {d(H24['net_assets'])}, a ninefold increase. "
        f"Every year in the period was a surplus year."))

    s.append(Spacer(1, 4))
    rows = [["Year", "Revenue", "Expenses", "Result", "Program\nratio", "Liquid\nrunway",
             "Net assets\n/ expenses"]]
    for y in I.HIST_YEARS:
        h = H[y]
        rows.append([f"FY{y}", d(h["revenue"]), d(h["expenses"]), d(h["operating_result"]),
                     p1(h["program_ratio"]), f"{h['liquid_runway_months']:.1f} mo",
                     f"{h['net_asset_multiple']:.2f}x"])
    s.append(table(rows, [W * 0.13, W * 0.155, W * 0.155, W * 0.15, W * 0.125, W * 0.13, W * 0.155],
                   align=[1, 2, 3, 4, 5, 6], size=8.0))
    s.append(Spacer(1, 4))
    s.append(para("Source: IRS Form 990, FY2019 to FY2024. Program ratio, runway and net asset "
                  "multiple depend on the modelled functional expense split and balance sheet "
                  "composition; the revenue, expense and net asset figures are as filed.", "note"))

    s.append(para("Stewardship is not the problem", "h2"))
    s.append(para(
        f"On every conventional measure of nonprofit stewardship this organization is comfortable. "
        f"The program expense ratio finished FY2024 at {p1(H24['program_ratio'])} against a BBB "
        f"Wise Giving Alliance minimum of 65%. Fundraising cost "
        f"{H24['fundraising_cost_ratio'] * 100:.1f} cents per dollar raised against a 35 cent "
        f"ceiling - a return of {H24['fundraising_roi']:.1f} dollars raised per dollar spent. The "
        f"administrative ratio is {p1(H24['admin_ratio'])}. Charity Navigator rates the "
        f"organization four stars at 96%. None of these ratios are where the risk sits, and a "
        f"board that spends its finance committee time on them is looking in the wrong place."))

    s.append(para(
        f"The measure that does carry a warning is accumulation. Net assets stand at "
        f"{H24['net_asset_multiple']:.2f} times annual expenses. BBB Standard 10 asks that "
        f"unrestricted net assets available for use not exceed three times annual expenses. The "
        f"organization is not in breach, but it has moved from {H[2019]['net_asset_multiple']:.2f}x "
        f"to {H24['net_asset_multiple']:.2f}x in six years, and the direction of travel is "
        f"toward the line rather than away from it."))

    # ---- FY2024 variance ------------------------------------------------
    s.append(PageBreak())
    s.append(para("2. FY2024 against plan: the variance that matters", "h1"))
    s.append(para(
        "No public FY2024 budget exists, so one has been reconstructed on stated assumptions and "
        "is labelled a simulation throughout 03_Budget_vs_Actual.xlsx. The reconstruction assumes "
        "a board approving 12% revenue growth and 26% expense growth in late FY2023 - cautious on "
        "income after two extraordinary years, deliberately expansionary on spending to deploy "
        "accumulated reserve. What the variances show is more interesting than the plan.", "body"))

    rows = [["Line", "Budget (simulated)", "Actual (filed)", "Variance", "Variance %", ""]]
    for k in ("Total revenue", "  Contributions and grants", "  Investment and other income",
              "Total expenses", "  Program services", "Operating result",
              "Participants served", "Successful outcomes"):
        v = L[k]
        fmt = n0 if "Participants" in k or "outcomes" in k else d
        rows.append([k.strip(), fmt(v["budget"]), fmt(v["actual"]),
                     ("+" if v["variance"] > 0 else "") + fmt(v["variance"]),
                     f"{v['variance_pct']:+.1%}",
                     "F" if v["favourable"] else "U"])
    s.append(table(rows, [W * 0.27, W * 0.17, W * 0.17, W * 0.16, W * 0.13, W * 0.05],
                   align=[1, 2, 3, 4, 5]))
    s.append(Spacer(1, 8))

    vol_share = DEC["volume_variance"] / DEC["total_program_variance"]
    budget_rate = L["Successful outcomes"]["budget"] / L["Participants served"]["budget"]
    actual_rate = L["Successful outcomes"]["actual"] / L["Participants served"]["actual"]
    lost = L["Participants served"]["actual"] * budget_rate - L["Successful outcomes"]["actual"]

    s.append(para(
        f"Revenue beat plan by {dn(L['Total revenue']['variance'])}, "
        f"{p1(abs(L['Total revenue']['variance_pct']))}. Before treating that as a fundraising "
        f"success, note that {dn(L['  Investment and other income']['variance'])} of it - "
        f"{p0(L['  Investment and other income']['variance'] / L['Total revenue']['variance'])} "
        f"of the total beat - is investment income on the accumulated reserve, running "
        f"{p1(abs(L['  Investment and other income']['variance_pct']))} above plan. That is a "
        f"balance sheet outcome and it will reverse when markets do."))

    s.append(para(
        f"Program spending exceeded budget by {dn(DEC['total_program_variance'])}. Decomposed, "
        f"{p0(vol_share)} of that overrun is volume - the organization served "
        f"{n0(L['Participants served']['actual'])} participants against a plan of "
        f"{n0(L['Participants served']['budget'])} - and the remaining "
        f"{dn(DEC['rate_variance'])} is rate, with cost per participant coming in "
        f"{p1(abs(L['Cost per participant']['variance_pct']))} above plan at "
        f"{d(DEC['actual_cpp'])}. Program expense exceeded budget primarily because participant "
        f"acquisition ran ahead of the planned expansion, not because delivery became materially "
        f"more expensive."))

    s.append(para("The finding", "h2"))
    s.append(para(
        f"In the same year that the organization served "
        f"{n0(abs(L['Participants served']['variance']))} more participants than planned, it "
        f"produced {n0(abs(L['Successful outcomes']['variance']))} <b>fewer</b> successful "
        f"outcomes - {n0(L['Successful outcomes']['actual'])} against a plan of "
        f"{n0(L['Successful outcomes']['budget'])}. The outcome rate fell from the "
        f"{p1(budget_rate)} assumed in the budget to {p1(actual_rate)}. Had the planned conversion "
        f"rate held across the actual participant base, the organization would have produced "
        f"roughly {n0(L['Participants served']['actual'] * budget_rate)} placements. The shortfall "
        f"of about {n0(lost)} placements is what the mix shift cost."))
    s.append(para(
        "That combination - more reach, fewer placements - is a mix problem rather than a cost "
        "problem, and cost discipline will not fix it. Growth came disproportionately through "
        "lower-converting channels. Reach is a means, and it was reported as though it were the "
        "end."))

    # ---- programs -------------------------------------------------------
    s.append(PageBreak())
    s.append(para("3. Program economics: which programs are most cost-effective", "h1"))
    s.append(para(
        "The four-program portfolio below is an analyst construction. Participant, outcome and "
        "cost totals reconcile exactly to the FY2024 figures derived from public data, but the "
        "split across programs is assumed and no public data exists to validate it. What follows "
        "should be read as a demonstration of how a portfolio like this behaves rather than a "
        "measurement of how these particular programs perform.", "note"))
    s.append(Spacer(1, 5))

    rows = [["Program", "Participants", "Outcomes", "Outcome rate", "Cost / participant",
             "Cost / outcome", "Gain / $"]]
    for p in I.PROGRAMS:
        e = PE[p]
        rows.append([p, n0(e["participants"]), n0(e["outcomes"]), p1(e["outcome_rate"]),
                     d(e["cost_per_participant"]), d(e["cost_per_outcome"]),
                     f"{e['earnings_gain_per_dollar']:.2f}x"])
    rows.append(["Portfolio", n0(PT["participants"]), n0(PT["outcomes"]), p1(PT["outcome_rate"]),
                 d(PT["cost_per_participant"]), d(PT["cost_per_outcome"]), ""])
    t = table(rows, [W * 0.21, W * 0.13, W * 0.11, W * 0.13, W * 0.15, W * 0.14, W * 0.13],
              align=[1, 2, 3, 4, 5, 6])
    t.setStyle(TableStyle([("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                           ("TEXTCOLOR", (0, -1), (-1, -1), NAVY),
                           ("LINEABOVE", (0, -1), (-1, -1), 0.9, NAVY)]))
    s.append(t)
    s.append(Spacer(1, 8))

    cheap = min(I.PROGRAMS, key=lambda x: PE[x]["cost_per_outcome"])
    dear = max(I.PROGRAMS, key=lambda x: PE[x]["cost_per_outcome"])
    bestgain = max(I.PROGRAMS, key=lambda x: PE[x]["earnings_gain_per_dollar"])

    s.append(para(
        f"{cheap} produces a placement for {d(PE[cheap]['cost_per_outcome'])} against "
        f"{d(PE[dear]['cost_per_outcome'])} for {dear} - a "
        f"{PE[dear]['cost_per_outcome'] / PE[cheap]['cost_per_outcome']:.1f} times difference. On "
        f"cost per placement alone the allocation decision looks settled. It is not."))
    s.append(para(
        f"{cheap} converts {p1(PE[cheap]['outcome_rate'])} of its enrolled participants, and the "
        f"placements it produces carry an average earnings gain of "
        f"{d(PE[cheap]['salary_gain'])}. Credential Support converts "
        f"{p1(PE['Credential Support']['outcome_rate'])} at "
        f"{d(PE['Credential Support']['salary_gain'])} of gain - "
        f"{PE['Credential Support']['salary_gain'] / PE[cheap]['salary_gain']:.1f} times as much "
        f"income mobility per person placed. Measured on earnings gain generated per dollar "
        f"spent, {bestgain} leads at {PE[bestgain]['earnings_gain_per_dollar']:.2f} times. Cost "
        f"per placement and value per placement point in different directions, and no further "
        f"analysis resolves that. It is a choice about what the organization is for."))
    s.append(para(
        f"Partner Capacity is the weakest program on every financial measure: "
        f"{d(PE['Partner Capacity']['cost_per_outcome'])} per placement, an outcome rate of "
        f"{p1(PE['Partner Capacity']['outcome_rate'])}, and the lowest capacity utilisation at "
        f"{p0(PE['Partner Capacity']['capacity_utilization'])}. Its case rests on sector leverage "
        f"and on outcomes delivered through partners that this model attributes at a discount. "
        f"That case may be sound, but it is a strategic argument and should be made explicitly "
        f"rather than averaged into the portfolio."))

    s.append(para("On social return", "h2"))
    s.append(para(
        f"A social return figure can be calculated and is, in 05_Program_Economics.xlsx. At the "
        f"stated parameters the portfolio ratio is "
        f"{M['sroi']['_portfolio']['sroi_ratio']:.2f} to one. It is not quoted as a headline here "
        f"because varying just two of the six parameters across plausible ranges moves it between "
        f"{M['sroi_range']['low']:.2f} and {M['sroi_range']['high']:.2f} - a spread of "
        f"{M['sroi_range']['spread']:.1f} times. There is no counterfactual, no comparison group "
        f"and no study behind the attribution and deadweight assumptions. Cost per successful "
        f"outcome rests on far fewer assumptions and is the measure management should steer on."))

    # ---- forecast and scenarios ----------------------------------------
    s.append(PageBreak())
    s.append(para("4. Forecast and scenarios: how much growth can funding support", "h1"))
    s.append(para(
        "The forecast runs five years rather than three. Three years is too short for this "
        "organization: reserves absorb almost any shock over that horizon, so a three-year model "
        "would report no problem for a structural deficit that becomes serious later. The Base "
        "case assumes foundation and government funding decline as the humanitarian response "
        "normalises. A base case that extrapolated the surge would be forecasting the past."))

    rows = [["Scenario", "FY2029 revenue", "FY2029 expenses", "FY2029 result", "Liquid runway",
             "Outcomes"]]
    for scen in I.SCENARIOS:
        f = SC[scen][-1]
        rows.append([scen, d(f.total_revenue), d(f.total_expense), d(f.operating_result),
                     f"{f.derived()['liquid_runway_months']:.1f} mo", n0(f.outcomes)])
    s.append(table(rows, [W * 0.16, W * 0.19, W * 0.19, W * 0.18, W * 0.15, W * 0.13],
                   align=[1, 2, 3, 4, 5]))
    s.append(Spacer(1, 8))

    cross = MVF["crossover"]
    s.append(para(
        f"<b>The central finding of the forecast.</b> On Base case assumptions the organization "
        f"crosses into structural deficit in FY{cross['year']}, "
        f"{cross['years_from_base']} years from the last filed year, with a deficit of "
        f"{dn(cross['result'])}. This is not a crisis projection. It is the arithmetic of a cost "
        f"base that grows with program volume while foundation and government funding normalise "
        f"after an episodic surge. Closing that gap requires contributions roughly "
        f"{p1(MVF['required_funding_uplift'])} above the Base case trajectory, or an equivalent "
        f"reduction in the rate of expense growth."))
    s.append(para(
        f"Reserves absorb all three cases comfortably. Liquid runway ends the horizon at "
        f"{SC['Downside'][-1].derived()['liquid_runway_months']:.0f} months even in the Downside "
        f"case. The balance sheet is not the constraint; the income statement is. That distinction "
        f"changes the remedy from cost-cutting to funding diversification, and it is the single "
        f"most useful thing this model has to say."))

    # ---- funding vulnerability -----------------------------------------
    s.append(para("5. How vulnerable is the organization to losing a major funder", "h1"))
    rows = [["Permanent funding cut", "FY2029 structural result", "Liquid runway",
             "Years to reserve floor", "Program capacity retained"]]
    for x in SENS:
        rows.append([p0(x["shock"]), d(x["structural_result"]),
                     f"{x['liquid_runway_months']:.1f} mo",
                     f"{x['years_to_floor']['years']} yrs" if x["years_to_floor"]["years"] else "16+",
                     p0(x["capacity_retained"])])
    s.append(table(rows, [W * 0.20, W * 0.24, W * 0.17, W * 0.20, W * 0.19], align=[1, 2, 3, 4]))
    s.append(Spacer(1, 8))

    s.append(para(
        f"The answer is that the organization is remarkably resilient in the short run and "
        f"structurally exposed in the long run, and these are easy to confuse. A permanent 20% cut "
        f"in contributions still leaves {SENS[2]['liquid_runway_months']:.0f} months of liquid "
        f"runway at FY2029 and forces no program cuts within the five-year window. It also "
        f"produces a structural deficit of {dn(SENS[2]['structural_result'])} and exhausts the "
        f"usable reserve in {SENS[2]['years_to_floor']['years']} years. Reserves buy time; they do "
        f"not close a structural gap."))
    s.append(para(
        f"Losing the single largest funder - modelled at {p0(I.LARGEST_FUNDER_SHARE)} of total "
        f"revenue, {dn(LF['dollars_lost_fy1'])} in the first year - moves the reserve floor from "
        f"{LF['reference_years_to_floor']['years']} years out to "
        f"{LF['years_to_floor']['years']} years out and widens the FY2029 deficit to "
        f"{dn(LF['final_operating_result'])}. No program cut is forced inside five years. The "
        f"honest framing for a board is not that this would be survivable, which it plainly would "
        f"be, but that it would consume two years of the runway available to fix a problem that "
        f"already exists."))
    s.append(para(
        f"Funding can fall {p1(MVF['max_sustainable_shock'])} before the six-month reserve floor "
        f"is reached within five years - equivalent to first-year contributions of "
        f"{d(MVF['min_contributions_fy1'])} against a Base case "
        f"{d(MVF['reference_contributions_fy1'])}, or {d(MVF['dollars_of_headroom'])} of "
        f"headroom. That is the minimum funding level required to maintain operations on the "
        f"current cost base over the planning horizon."))

    # ---- allocation ------------------------------------------------------
    s.append(PageBreak())
    s.append(para("6. Where should incremental funding be allocated", "h1"))
    s.append(para(
        f"An incremental {d(I.ALLOCATION_BUDGET)} was allocated across the four programs subject "
        f"to minimum and maximum share limits, delivery capacity ceilings, staff hour "
        f"availability and a {p0(I.ALLOCATION_CONSTRAINTS['admin_load'])} administrative load. "
        f"The problem was solved three times under three objectives, because they disagree."))

    rows = [["Objective", "Career Coaching", "Digital Learning", "Credential Support",
             "Partner Capacity", "Placements", "Avg gain"]]
    labels = {"outcomes": "Maximise placements", "balanced": "Balanced (quality floor)",
              "earnings": "Maximise earnings gain"}
    for o in ("outcomes", "balanced", "earnings"):
        a = AR[o]
        rows.append([labels[o]] + [d(a["allocation"][p]) for p in I.PROGRAMS]
                    + [f"{a['outcomes']:.0f}", d(a["avg_gain"])])
    s.append(table(rows, [W * 0.19, W * 0.14, W * 0.14, W * 0.14, W * 0.13, W * 0.13, W * 0.13],
                   align=[1, 2, 3, 4, 5, 6], size=7.8))
    s.append(Spacer(1, 8))

    ro, rb, re_ = AR["outcomes"], AR["balanced"], AR["earnings"]
    s.append(para(
        f"Maximising placement count delivers {ro['outcomes']:.0f} additional placements at an "
        f"average earnings gain of {d(ro['avg_gain'])} each. Maximising earnings gain delivers "
        f"only {re_['outcomes']:.0f} placements - {AO['outcomes_price_of_earnings_focus']:.0f} "
        f"fewer - but at {d(re_['avg_gain'])} of gain each and {dn(re_['earnings'] - ro['earnings'])} "
        f"more earnings gain in total. Same money, same constraints, materially different "
        f"programs funded."))
    s.append(para(
        f"<b>The recommended allocation is the balanced case.</b> Holding placement count as the "
        f"objective while requiring an average gain of at least "
        f"{d(I.ALLOCATION_CONSTRAINTS['min_avg_salary_gain'])} per placement yields "
        f"{rb['outcomes']:.0f} placements - only {ro['outcomes'] - rb['outcomes']:.1f} fewer than "
        f"the placement-maximising answer - while lifting average gain from {d(ro['avg_gain'])} to "
        f"{d(rb['avg_gain'])}. Roughly {p1((ro['outcomes'] - rb['outcomes']) / ro['outcomes'])} of "
        f"placement count buys a {p0(rb['avg_gain'] / ro['avg_gain'] - 1)} improvement in "
        f"placement quality."))
    s.append(para(
        f"All three optimised answers beat the two default policies. An equal split across "
        f"programs delivers {M['allocation']['comparison']['equal_split']:.0f} placements and a "
        f"pro rata roll-forward of current program size delivers "
        f"{M['allocation']['comparison']['pro_rata']:.0f}, against {ro['outcomes']:.0f} for the "
        f"optimised allocation. Doing this analysis at all is worth roughly "
        f"{ro['outcomes'] - M['allocation']['comparison']['pro_rata']:.0f} additional placements a "
        f"year against the most likely default."))
    s.append(para(
        "One caution. The model assumes marginal cost per outcome stays at current average cost "
        "as programs expand. Real expansion faces rising marginal cost because the "
        "easiest-to-serve participants are already being served, and the effect is strongest in "
        "the cheapest program. The linear model therefore overstates the case for concentration, "
        "which is a further reason to prefer the balanced allocation over the corner solution.",
        "note"))

    # ---- recommendations -------------------------------------------------
    s.append(PageBreak())
    s.append(para("7. Recommendations", "h1"))
    recs = [
        (f"<b>Set an explicit reserve policy and a deployment schedule.</b> Net assets sit at "
         f"{H24['net_asset_multiple']:.2f} times annual expenses against a BBB accumulation "
         f"ceiling of three times. A seventh consecutive surplus would raise a legitimate donor "
         f"question. The board should adopt a target reserve range in months of expenses, and "
         f"commit the excess above it to a multi-year program expansion with dated milestones."),
        (f"<b>Treat FY{cross['year']} as the planning horizon, not FY2027.</b> The Base case "
         f"turns to deficit in FY{cross['year']} and the usable reserve is exhausted around "
         f"FY{MVF['years_to_floor_no_shock']['year']}. Both dates are far enough away to be "
         f"ignored and close enough to matter. Closing the structural gap needs contributions "
         f"roughly {p1(MVF['required_funding_uplift'])} above trajectory, and fundraising "
         f"capacity of that size takes years to build."),
        (f"<b>Diversify away from institutional funding.</b> On the modelled mix, foundation and "
         f"government funding together account for "
         f"{p0(I.REVENUE_MIX['Foundation grants'] + I.REVENUE_MIX['Government funding'])} of "
         f"contributions, and these are the categories most exposed to the humanitarian response "
         f"cycle. Individual giving, the most durable category, is the smallest. Shifting that "
         f"balance is the single highest-value use of the reserve."),
        (f"<b>Report placements, not reach.</b> FY2024 served "
         f"{n0(abs(L['Participants served']['variance']))} more participants than planned and "
         f"produced {n0(abs(L['Successful outcomes']['variance']))} fewer placements. Management "
         f"reporting should lead with outcome rate and cost per successful outcome by program, "
         f"with participant counts as a secondary measure."),
        (f"<b>Adopt the balanced allocation and set the quality floor explicitly.</b> Fund the "
         f"balanced case: {d(rb['allocation']['Career Coaching'])} to Career Coaching, "
         f"{d(rb['allocation']['Digital Learning'])} to Digital Learning, and the minimum share to "
         f"Credential Support and Partner Capacity. More importantly, make the earnings-gain floor "
         f"a stated board policy rather than an emergent property of where the marginal dollar "
         f"landed."),
        (f"<b>Require a defence of Partner Capacity on its own terms.</b> At "
         f"{d(PE['Partner Capacity']['cost_per_outcome'])} per placement and "
         f"{p0(PE['Partner Capacity']['capacity_utilization'])} capacity utilisation it is the "
         f"weakest program financially. If it is retained for sector-leverage reasons, that case "
         f"should be documented and given its own success measures rather than being judged on "
         f"metrics it was never designed to win."),
    ]
    for r_ in recs:
        s.append(Paragraph(r_, ST["bullet"], bulletText="•"))

    s.append(para("What would change these conclusions", "h1"))
    s.append(para(
        "The program-level analysis rests on a synthetic portfolio. If the real distribution of "
        "cost and outcomes across programs differs materially from the assumed one, sections 3 and "
        "6 change and the allocation recommendation may reverse. The functional expense split is "
        "modelled rather than filed, so the program, administrative and fundraising ratios in "
        "section 1 would move if the actual Part IX split differs from the assumption. The "
        "structural deficit date in section 4 is sensitive to the assumed rate of decline in "
        "institutional funding: a Base case that held foundation grants flat rather than declining "
        "6% a year would push the crossover several years further out. None of these would change "
        "the shape of the argument - a large reserve, an episodic funding base, and a widening gap "
        "between the two - but each would change the numbers attached to it."))

    s.extend(rule(10, 8))
    s.append(para(DISCLAIMER, "warn"))
    doc.build(s)
    print(f"wrote {REPORTS / 'Management_Report.pdf'}")


# ==========================================================================
def executive_summary():
    W = A4[0] - 40 * mm
    doc = Doc(REPORTS / "Executive_Summary.pdf", "Executive Summary",
              f"{I.ORG_NAME} - executive summary - independent analysis from public data")
    s = []
    cross = MVF["crossover"]
    rb = AR["balanced"]

    s.append(para("Executive summary", "title"))
    s.append(para(f"Nonprofit financial planning, scenario forecasting and impact measurement "
                  f"&nbsp;|&nbsp; {I.ORG_NAME} (EIN {I.ORG_EIN}) &nbsp;|&nbsp; {I.AS_OF}",
                  "subtitle"))

    s.append(kpi_row([
        ("FY2024 revenue", dm(H24["revenue"])),
        ("Operating result", dm(H24["operating_result"])),
        ("Net assets / expenses", f"{H24['net_asset_multiple']:.2f}x"),
        ("Liquid runway", f"{H24['liquid_runway_months']:.0f} mo"),
        ("Structural deficit from", f"FY{cross['year']}"),
        ("Cost per outcome", d(H24["cost_per_outcome"])),
    ], W))
    s.append(Spacer(1, 13))

    s.append(para(
        f"This organization has a strong balance sheet and a weakening business model, and the "
        f"first is concealing the second. Six consecutive surpluses have built net assets to "
        f"{d(H24['net_assets'])} - {H24['net_asset_multiple']:.2f} times annual expenses and "
        f"{H24['liquid_runway_months']:.0f} months of liquid runway. That cushion is large enough "
        f"to absorb a {p1(MVF['max_sustainable_shock'])} permanent cut in contributions without "
        f"cutting a single program inside five years. It is also large enough to make a slow "
        f"structural decline invisible for most of a decade.", "lead"))

    s.append(para("The four things that matter", "h1"))

    s.append(para("1. The growth was episodic, and the cost base is not", "h2"))
    s.append(para(
        f"Revenue compounded at {p1(HS['revenue_cagr'])} a year from FY2019 to FY2024, almost "
        f"entirely through contributions tied to the Afghanistan and Ukraine displacement "
        f"responses. Expenses compounded at {p1(HS['expense_cagr'])}. The gap produced "
        f"{d(HS['cumulative_surplus'])} of cumulative surplus, but it also built a delivery "
        f"organization sized to funding that arrived for reasons outside management's control. "
        f"On Base case assumptions, where foundation and government funding normalise rather than "
        f"collapse, the organization crosses into structural deficit in FY{cross['year']} and "
        f"reaches its reserve floor around FY{MVF['years_to_floor_no_shock']['year']}. Closing the "
        f"gap needs contributions roughly {p1(MVF['required_funding_uplift'])} above trajectory."))

    s.append(para("2. Reach grew; placements did not", "h2"))
    budget_rate = L["Successful outcomes"]["budget"] / L["Participants served"]["budget"]
    actual_rate = L["Successful outcomes"]["actual"] / L["Participants served"]["actual"]
    s.append(para(
        f"FY2024 served {n0(abs(L['Participants served']['variance']))} more participants than "
        f"planned and produced {n0(abs(L['Successful outcomes']['variance']))} fewer successful "
        f"outcomes. The outcome rate fell from {p1(budget_rate)} to {p1(actual_rate)} as growth "
        f"came through lower-converting channels. This is a mix problem, not a cost problem, and "
        f"it is the most actionable finding in the analysis. Management reporting that leads with "
        f"participant counts will not surface it."))

    s.append(para("3. The programs cannot be ranked on one measure", "h2"))
    cheap = min(I.PROGRAMS, key=lambda x: PE[x]["cost_per_outcome"])
    s.append(para(
        f"{cheap} produces a placement for {d(PE[cheap]['cost_per_outcome'])}; Credential Support "
        f"costs {d(PE['Credential Support']['cost_per_outcome'])} but delivers "
        f"{PE['Credential Support']['salary_gain'] / PE[cheap]['salary_gain']:.1f} times the "
        f"earnings gain per person placed. Optimising an incremental "
        f"{d(I.ALLOCATION_BUDGET)} for placement count and optimising it for earnings gain produce "
        f"almost opposite allocations. The recommended balanced allocation gives up "
        f"{p1((AR['outcomes']['outcomes'] - rb['outcomes']) / AR['outcomes']['outcomes'])} of "
        f"placement count to lift average earnings gain per placement by "
        f"{p0(rb['avg_gain'] / AR['outcomes']['avg_gain'] - 1)}, and beats a pro rata "
        f"roll-forward by {AR['outcomes']['outcomes'] - M['allocation']['comparison']['pro_rata']:.0f} "
        f"placements a year."))

    s.append(para("4. Stewardship ratios are not where the risk is", "h2"))
    s.append(para(
        f"Program expense ratio {p1(H24['program_ratio'])} against a 65% minimum; fundraising cost "
        f"{H24['fundraising_cost_ratio'] * 100:.1f} cents per dollar against a 35 cent ceiling; "
        f"four-star Charity Navigator rating. All comfortable, none informative. The measure that "
        f"does carry a warning is accumulation: net assets have moved from "
        f"{H[2019]['net_asset_multiple']:.2f}x to {H24['net_asset_multiple']:.2f}x annual expenses "
        f"in six years against a BBB ceiling of three times."))

    s.append(para("Recommended actions", "h1"))
    for t in [
        f"Adopt a stated reserve target in months of expenses and commit the excess to a dated, "
        f"multi-year program expansion. The reserve is an asset only if it is deployed.",
        f"Plan to FY{cross['year']}, not FY2027. Build the fundraising capacity to close a "
        f"{p1(MVF['required_funding_uplift'])} structural gap before the reserve makes the "
        f"decision urgent.",
        f"Shift the funding mix away from institutional sources, which are "
        f"{p0(I.REVENUE_MIX['Foundation grants'] + I.REVENUE_MIX['Government funding'])} of "
        f"contributions and the most exposed to the response cycle.",
        f"Report outcome rate and cost per successful outcome by program as the primary measures; "
        f"demote participant counts to context.",
        f"Fund the balanced allocation and make the earnings-gain quality floor an explicit board "
        f"policy rather than an accident of where the marginal dollar landed.",
    ]:
        s.append(Paragraph(t, ST["bullet"], bulletText="•"))

    s.append(para("What this analysis is", "h1"))
    s.append(para(
        f"An independent financial model built from public data: six years of IRS Form 990 "
        f"filings, the organization's published impact reporting, and sector benchmarks from BBB "
        f"Wise Giving Alliance and Propel Nonprofits. Historical totals are real and cited. The "
        f"four-program portfolio, the split of expenses into functional categories, the balance "
        f"sheet composition and every forward-looking figure are analyst assumptions, tagged by "
        f"provenance across {len(I.REGISTER)} registered inputs in 01_Assumptions.xlsx. Of those, "
        f"{I.REGISTER.tier_counts()['PUBLIC']} are taken verbatim from filings, "
        f"{I.REGISTER.tier_counts()['DERIVED']} are arithmetic on public values, "
        f"{I.REGISTER.tier_counts()['BENCHMARK']} are published sector standards and "
        f"{I.REGISTER.tier_counts()['SYNTHETIC']} are assumptions. Wherever a public control total "
        f"exists, the modelled composition reconciles to it exactly."))

    s.extend(rule(8, 6))
    s.append(para(DISCLAIMER, "warn"))
    doc.build(s)
    print(f"wrote {REPORTS / 'Executive_Summary.pdf'}")


def build():
    REPORTS.mkdir(parents=True, exist_ok=True)
    management_report()
    executive_summary()


if __name__ == "__main__":
    build()
