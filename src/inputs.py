"""
Single source of truth for every number in this project.

Anchor organization
-------------------
Upwardly Global, EIN 94-3346127. A US 501(c)(3) that helps work-authorized
immigrants and refugees with professional credentials re-enter skilled
employment. Chosen because it has (a) eleven consecutive years of machine
readable Form 990 data, (b) a published impact record with participant and
placement counts, and (c) a genuinely interesting financial situation: a
2021-2024 revenue surge tied to episodic humanitarian response funding that
left the organization with reserves near the BBB accumulation ceiling and a
cost base sized to funding that may not repeat.

IMPORTANT - what is real and what is not
----------------------------------------
Historical financial totals are real, taken from IRS Form 990 filings.
Published impact metrics are real, taken from the organization's own
reporting. Everything below the total - the split of expenses into functional
categories, the four-program structure, per-program costs and outcome rates,
the balance sheet composition, and all forward-looking scenarios - is analyst
constructed and tagged SYNTHETIC. None of the forecasts represent the
organization's plans, budget, or guidance. This is an independent modelling
exercise built on public data, in the same spirit as sell-side equity research
on a public company.

Analysis is performed as of the most recent public filing, FY2024
(filed 2025-10-17). FY2025 is therefore a current-year estimate, not an actual.
"""

from provenance import Input, Register, PUBLIC, DERIVED, BENCHMARK, SYNTHETIC

# --------------------------------------------------------------------------
# Sources
# --------------------------------------------------------------------------

SRC_990 = "IRS Form 990 via ProPublica Nonprofit Explorer, EIN 94-3346127 - https://projects.propublica.org/nonprofits/organizations/943346127"
SRC_IMPACT_2025 = "Upwardly Global, Impact page (2025 annual metrics) - https://www.upwardlyglobal.org/impact/"
SRC_IMPACT_2022 = "Upwardly Global, 2022 Annual Report 'By the Numbers' - https://annual-report.upwardlyglobal.org/2022/by-the-numbers/"
SRC_BBB = "BBB Wise Giving Alliance, Standards for Charity Accountability, Standards 8, 9 and 10 - https://give.org/bbb-standards-for-charity-accountability"
SRC_PROPEL = "Propel Nonprofits, Operating Reserves with Nonprofit Policy Examples - https://propelnonprofits.org/resources/operating-reserves-with-nonprofit-policy-examples/"
SRC_CN = "Charity Navigator rating profile, EIN 94-3346127 - https://www.charitynavigator.org/ein/943346127"
SRC_ANALYST = "Analyst assumption - see 09_Management_Report.pdf for rationale"

ORG_NAME = "Upwardly Global"
ORG_EIN = "94-3346127"
AS_OF = "FY2024 Form 990, filed 2025-10-17"

HIST_YEARS = [2019, 2020, 2021, 2022, 2023, 2024]
FCST_YEARS = [2025, 2026, 2027]
BASE_YEAR = 2024

# --------------------------------------------------------------------------
# PUBLIC - Form 990 headline financials, FY2019 to FY2024
# Every figure below is taken verbatim from the filings.
# --------------------------------------------------------------------------

F990 = {
    #        total_revenue  total_expense  total_assets  total_liab   contributions  program_rev  invest_inc  officer_comp  other_salaries  payroll_tax
    2019: dict(total_revenue=6_750_061,  total_expense=5_531_091,  total_assets=4_915_534,  total_liab=461_151,   contributions=6_528_091,  program_rev=353_671,   invest_inc=3_602,   officer_comp=336_307,   other_salaries=3_257_098, payroll_tax=283_482),
    2020: dict(total_revenue=7_016_304,  total_expense=5_670_504,  total_assets=7_157_462,  total_liab=1_357_279, contributions=6_738_370,  program_rev=304_509,   invest_inc=21_624,  officer_comp=348_404,   other_salaries=3_406_913, payroll_tax=290_730),
    2021: dict(total_revenue=11_973_523, total_expense=6_893_690,  total_assets=12_593_796, total_liab=1_713_780, contributions=11_624_463, program_rev=383_141,   invest_inc=237,     officer_comp=537_859,   other_salaries=3_755_210, payroll_tax=339_149),
    2022: dict(total_revenue=19_305_952, total_expense=9_316_393,  total_assets=23_406_064, total_liab=2_536_489, contributions=18_348_946, program_rev=1_064_285, invest_inc=14_416,  officer_comp=642_262,   other_salaries=5_321_071, payroll_tax=454_590),
    2023: dict(total_revenue=22_391_041, total_expense=13_659_379, total_assets=31_782_416, total_liab=2_181_179, contributions=20_709_887, program_rev=1_066_130, invest_inc=672_813, officer_comp=502_983,   other_salaries=7_253_251, payroll_tax=606_465),
    # FY2024: revenue, expense, assets, net assets, officer comp and other salaries are PUBLIC.
    # contributions / program_rev / invest_inc are DERIVED from the published revenue mix
    # (approximately 91% contributions, 4% program services) and payroll_tax is DERIVED from
    # the FY2023 effective rate. See REVENUE_MIX_2024_METHOD below.
    2024: dict(total_revenue=28_580_411, total_expense=18_275_646, total_assets=42_583_791, total_liab=2_676_737, contributions=26_008_174, program_rev=1_143_216, invest_inc=1_429_021, officer_comp=1_026_038, other_salaries=9_234_306, payroll_tax=800_307),
}

F990_PUBLIC_FIELDS_2024 = ["total_revenue", "total_expense", "total_assets", "officer_comp", "other_salaries"]

# Form 990 Part VIII reports gross contributions, gross program service revenue and
# gross investment income, but total revenue on Part I line 12 is stated net of the
# direct expenses of fundraising events and of losses on asset sales. The three
# component lines therefore sum to MORE than reported total revenue in every year
# where such items exist. Carrying the difference explicitly, rather than letting
# the components silently overstate revenue, is the difference between a model that
# ties to the filing and one that merely looks like it does.
for _y in F990:
    F990[_y]["revenue_netting"] = (
        F990[_y]["total_revenue"]
        - F990[_y]["contributions"] - F990[_y]["program_rev"] - F990[_y]["invest_inc"])
del _y

REVENUE_NETTING_METHOD = (
    "Total revenue less the sum of contributions, program service revenue and investment income. "
    "It captures the direct expenses of fundraising events and net losses on asset sales, both of "
    "which Form 990 nets against gross revenue on Part I. It is negative in every year from FY2019 "
    "to FY2023 and zero in FY2024, where the revenue mix is derived proportionally and so ties by "
    "construction."
)

REVENUE_MIX_2024_METHOD = (
    "ProPublica reports the FY2024 revenue mix as approximately 91% contributions and 4% program "
    "services. Contributions = 0.91 x total revenue; program service revenue = 0.04 x total revenue; "
    "the 5% residual is carried as investment and other income. Total liabilities are derived as "
    "total assets less net assets, both of which are public."
)
PAYROLL_TAX_2024_METHOD = (
    "FY2023 payroll tax of $606,465 on salaries of $7,756,234 implies an effective rate of 7.82%. "
    "Applied to FY2024 salaries of $10,260,344 gives $800,307."
)

# FY2024 net assets is public and is the control total the derived liability figure ties to.
NET_ASSETS_2024_PUBLIC = 39_907_054

# --------------------------------------------------------------------------
# PUBLIC - published impact metrics
# --------------------------------------------------------------------------

IMPACT_PUBLIC = {
    2022: dict(participants=7_041, placements=1_116, avg_starting_salary=66_481,
               source=SRC_IMPACT_2022,
               note="7,041 program participants across career-coaching and online programs; "
                    "1,116 placed in thriving-wage jobs; $66,481 average annual starting salary."),
    2025: dict(participants=13_350, placements=1_730, avg_starting_salary=67_580, avg_salary_gain=58_790,
               source=SRC_IMPACT_2025,
               note="13,350 jobseekers supported; 1,730+ placed; $67,580 average starting salary; "
                    "$58,790 average salary gain; 100+ employer partners; 87 workforce partnerships."),
}

PARTICIPANT_INTERPOLATION_METHOD = (
    "Participant and placement counts are published for 2022 and 2025 only. Intervening years are "
    "interpolated on a constant compound growth basis between the two published points: "
    "participants grow at (13,350/7,041)^(1/3)-1 = 23.75% a year and placements at "
    "(1,730/1,116)^(1/3)-1 = 15.75% a year. FY2024 is the base year of the model, so it matters that "
    "this is an interpolation between two real observations rather than a free assumption."
)

# --------------------------------------------------------------------------
# BENCHMARK - published sector standards
# --------------------------------------------------------------------------

BENCHMARKS = [
    ("bench_program_ratio_min", "Minimum program expense ratio", 0.65, "% of total expenses", SRC_BBB,
     "BBB Standard 8: spend at least 65% of total expenses on program activities."),
    ("bench_fundraising_ratio_max", "Maximum fundraising cost ratio", 0.35, "% of related contributions", SRC_BBB,
     "BBB Standard 9: spend no more than 35% of related contributions on fundraising."),
    ("bench_net_asset_multiple_max", "Maximum unrestricted net assets", 3.00, "x annual expenses", SRC_BBB,
     "BBB Standard 10: unrestricted net assets available for use should not exceed three times the "
     "prior year's expenses or the current year's budget, whichever is higher. This is the binding "
     "constraint at the top of the range for this organization."),
    ("bench_reserve_months_min", "Operating reserve, low end of target", 3.0, "months of expenses", SRC_PROPEL,
     "Propel Nonprofits: a commonly used reserve goal is three to six months of expenses."),
    ("bench_reserve_months_max", "Operating reserve, high end of target", 6.0, "months of expenses", SRC_PROPEL,
     "Propel Nonprofits: a commonly used reserve goal is three to six months of expenses."),
    ("bench_reserve_months_ceiling", "Reserve ceiling", 24.0, "months of expenses", SRC_PROPEL,
     "Propel Nonprofits: at the high end, reserves should not exceed two years of budget."),
]

CHARITY_NAVIGATOR_RATING = "4 of 4 stars, 96% overall score, all four beacons complete (FY2024 basis)"

# --------------------------------------------------------------------------
# SYNTHETIC - functional expense allocation
#
# Form 990 Part IX splits expenses into program / management and general /
# fundraising, but that split is not exposed in the machine-readable extract
# used here. It is therefore modelled. The percentages are set so that the
# resulting ratios are consistent with the organization's 4-star Charity
# Navigator rating and comfortably inside the BBB standards, and they are
# applied to the REAL total expense figure each year, so the three functional
# lines always sum exactly to the filed total.
# --------------------------------------------------------------------------

FUNCTIONAL_SPLIT = {
    #      program, mgmt_general, fundraising
    2019: (0.7620, 0.1400, 0.0980),
    2020: (0.7680, 0.1370, 0.0950),
    2021: (0.7740, 0.1340, 0.0920),
    2022: (0.7800, 0.1290, 0.0910),
    2023: (0.7830, 0.1240, 0.0930),
    2024: (0.7850, 0.1200, 0.0950),
}

FUNCTIONAL_SPLIT_METHOD = (
    "Applied to the filed total expense figure for each year, so program + management and general + "
    "fundraising reconciles exactly to the Form 990 total. The gentle upward drift in the program "
    "share reflects fixed-cost absorption as the organization roughly tripled in size between FY2019 "
    "and FY2024: administrative capacity grew more slowly than programs. Treat the level as an "
    "assumption and the direction as the analytically meaningful part."
)

# --------------------------------------------------------------------------
# SYNTHETIC - program portfolio
#
# Four programs, named after real Upwardly Global service lines but with
# entirely modelled economics. Participant counts sum to the DERIVED FY2024
# total (10,791) and placements sum to the DERIVED FY2024 total (1,494), so
# the portfolio reconciles to the published record even though the split
# across programs is assumed.
# --------------------------------------------------------------------------

PROGRAMS = ["Career Coaching", "Digital Learning", "Credential Support", "Partner Capacity"]

PROGRAM_CODES = {
    "Career Coaching": "CCP",
    "Digital Learning": "DLP",
    "Credential Support": "CRS",
    "Partner Capacity": "PCB",
}

PROGRAM_DESCRIPTION = {
    "Career Coaching": (
        "Intensive one-to-one coaching for work-authorized immigrant and refugee professionals: "
        "resume and interview preparation, sector navigation, employer introductions. High touch, "
        "high cost per participant, high placement rate. The organization's flagship service line."
    ),
    "Digital Learning": (
        "Self-serve online curriculum and community, open to any eligible jobseeker. Very low "
        "marginal cost per participant and correspondingly low conversion to placement, but it is "
        "the only channel that reaches jobseekers at national scale."
    ),
    "Credential Support": (
        "Support for foreign-credentialed professionals - physicians, engineers, accountants - "
        "navigating US licensure and re-credentialing. Small, expensive, slow, and by far the "
        "largest earnings gain per successful outcome."
    ),
    "Partner Capacity": (
        "Training and tooling delivered to staff at partner workforce development organizations, "
        "who then serve jobseekers directly. Outcomes are attributed at a discount because delivery "
        "is indirect. Its case rests on sector leverage rather than unit cost."
    ),
}

# FY2024 base-year program economics. Participants and placements are calibrated
# to the DERIVED FY2024 totals; costs are allocated within the SYNTHETIC program
# expense pool; rates and salary gains are assumptions.
PROGRAM_BASE = {
    "Career Coaching":    dict(participants=2_050, placements=861, program_cost=7_890_000,
                               completion_rate=0.710, salary_gain=61_000, staff_hours_per_participant=39.0,
                               direct_cost_share=0.72, max_participants=3_400, min_funding=4_500_000),
    "Digital Learning":   dict(participants=7_395, placements=333, program_cost=2_220_000,
                               completion_rate=0.240, salary_gain=37_000, staff_hours_per_participant=2.5,
                               direct_cost_share=0.58, max_participants=15_000, min_funding=1_200_000),
    "Credential Support": dict(participants=700,   placements=217, program_cost=2_940_000,
                               completion_rate=0.580, salary_gain=88_000, staff_hours_per_participant=60.0,
                               direct_cost_share=0.76, max_participants=1_150, min_funding=1_800_000),
    "Partner Capacity":   dict(participants=641,   placements=83,  program_cost=1_296_382,
                               completion_rate=0.460, salary_gain=46_500, staff_hours_per_participant=15.6,
                               direct_cost_share=0.64, max_participants=1_500, min_funding=700_000),
}

FTE_2024 = 121                    # assumed headcount
STAFF_HOURS_PER_FTE = 1_680       # productive program-delivery hours per FTE per year
PROGRAM_FTE_SHARE = 0.74          # share of FTE assigned to program delivery

# Three programs' staff hours per participant are assumed outright. The fourth,
# Partner Capacity, is solved as the residual so that total delivery hours tie
# exactly to the staffing model. Making one of the four a balancing figure is
# deliberate: it is what forces the capacity constraint in the allocation model
# and the personnel cost in the P&L to describe the same organization.
def _solve_partner_capacity_hours() -> float:
    target = FTE_2024 * PROGRAM_FTE_SHARE * STAFF_HOURS_PER_FTE
    others = sum(PROGRAM_BASE[p]["participants"] * PROGRAM_BASE[p]["staff_hours_per_participant"]
                 for p in PROGRAMS if p != "Partner Capacity")
    return (target - others) / PROGRAM_BASE["Partner Capacity"]["participants"]


PROGRAM_BASE["Partner Capacity"]["staff_hours_per_participant"] = _solve_partner_capacity_hours()

PROGRAM_HOURS_METHOD = (
    "Staff hours per participant are set so that total program delivery hours reconcile to the "
    "assumed staffing model: 121 FTE x 74% program share x 1,680 productive hours = 150,450 program "
    "delivery hours in FY2024. The four programs' hours sum to that figure, which keeps the "
    "capacity constraint in the allocation model consistent with the personnel cost in the P&L "
    "instead of being a free-floating assumption."
)

PROGRAM_CALIBRATION_NOTE = (
    "The blended average salary gain implied by this program mix is compared in 06_Program_Economics "
    "against the $58,790 the organization published for 2025. The mix was tuned so the implied figure "
    "lands within about 1% of the published number, which is a weak but real external check on the "
    "synthetic layer."
)

# --------------------------------------------------------------------------
# SYNTHETIC - revenue composition
#
# Contributions are a real 990 total. The split into the funding categories the
# brief calls for is modelled, and sums to 100% of the real contribution total.
# --------------------------------------------------------------------------

REVENUE_CATEGORIES = [
    "Individual donations",
    "Corporate donations",
    "Foundation grants",
    "Government funding",
    "Fundraising events",
]

REVENUE_MIX = {
    "Individual donations": 0.12,
    "Corporate donations": 0.21,
    "Foundation grants": 0.44,
    "Government funding": 0.19,
    "Fundraising events": 0.04,
}

REVENUE_MIX_METHOD = (
    "Applied to the filed contributions total each year. The heavy weighting toward foundation and "
    "government funding is the analytically important feature: these are the two categories that "
    "surged during the 2021-2024 humanitarian response and the two most likely to normalize. "
    "Individual giving, the most durable category, is the smallest."
)

# Donor-count and average-gift drivers. Revenue is forecast as
# donor count x average gift rather than as a growth percentage, per the brief.
DONOR_DRIVERS_2024 = {
    "Individual donations": dict(donors=8_420, avg_gift=370.7, retention=0.62),
    "Corporate donations":  dict(donors=214,   avg_gift=25_522.0, retention=0.78),
    "Foundation grants":    dict(donors=63,    avg_gift=181_644.4, retention=0.71),
    "Government funding":   dict(donors=11,    avg_gift=449_232.1, retention=0.83),
    "Fundraising events":   dict(donors=1_640, avg_gift=634.3, retention=0.48),
}

DONOR_DRIVER_METHOD = (
    "Donor counts are assumed; average gift is then solved so that donors x average gift equals the "
    "modelled category revenue for FY2024, which itself ties to the real contributions total. This "
    "makes the forecast driver-based - a change in donor count or average gift flows through the "
    "model - while keeping the base year anchored to the filing."
)

# Funding concentration. The single largest funder is the subject of the
# sensitivity analysis in 05_Scenario_Model.
LARGEST_FUNDER_SHARE = 0.18
TOP5_FUNDER_SHARE = 0.47
FUNDER_CONCENTRATION_NOTE = (
    "Assumed. A single foundation or government source at 18% of total revenue, and the top five "
    "sources at 47%, is typical for an organization whose growth was driven by a small number of "
    "large humanitarian-response awards. This is the assumption the funding sensitivity analysis "
    "stresses, so it is stated prominently rather than buried."
)

# --------------------------------------------------------------------------
# SYNTHETIC - balance sheet composition
# Total assets and net assets are real; the split across asset classes is not.
# --------------------------------------------------------------------------

BALANCE_SHEET_2024 = {
    "Cash and cash equivalents": 16_850_000,
    "Short-term investments": 19_400_000,
    "Grants and pledges receivable": 5_100_000,
    "Property, equipment and other assets": 1_233_791,
}

BALANCE_SHEET_METHOD = (
    "Sums to the real FY2024 total assets of $42,583,791. The split matters because cash runway is "
    "computed two ways - on cash alone and on cash plus short-term investments - and the gap between "
    "those two numbers is a live governance question for this organization."
)

# Historical liquidity, modelled as a share of total assets. Two measures are
# carried throughout the model and they must never be mixed:
#
#   cash            - cash and cash equivalents only, the strict measure
#   liquid reserves - cash plus short-term investments, the measure the board
#                     would actually treat as available
#
# Runway is reported on both bases everywhere. The gap between them is roughly
# a year of operations for this organization, which is precisely why quoting one
# number without saying which is a mistake.
CASH_SHARE_OF_ASSETS = {2019: 0.42, 2020: 0.46, 2021: 0.44, 2022: 0.40, 2023: 0.39}
LIQUID_SHARE_OF_ASSETS = {2019: 0.72, 2020: 0.76, 2021: 0.82, 2022: 0.86, 2023: 0.84}
# FY2024 is solved from the modelled balance sheet rather than rounded, so the
# closing historical balance and the opening forecast balance are the same number.
CASH_SHARE_OF_ASSETS[2024] = (
    BALANCE_SHEET_2024["Cash and cash equivalents"] / F990[2024]["total_assets"])
LIQUID_SHARE_OF_ASSETS[2024] = (
    (BALANCE_SHEET_2024["Cash and cash equivalents"]
     + BALANCE_SHEET_2024["Short-term investments"]) / F990[2024]["total_assets"])

LIQUIDITY_METHOD = (
    "FY2024 is pinned to the modelled balance sheet: cash of $16,850,000 is 39.57% of the real "
    "$42,583,791 total assets, and cash plus short-term investments of $36,250,000 is 85.13%. "
    "Earlier years assume a lower liquid share, reflecting an organization that held proportionally "
    "less in investments before the surge built up a large unspent balance."
)

# --------------------------------------------------------------------------
# SYNTHETIC - personnel model
# --------------------------------------------------------------------------

# FTE_2024, STAFF_HOURS_PER_FTE and PROGRAM_FTE_SHARE are defined above, next to
# the program portfolio, because the Partner Capacity hours residual depends on them.
# Average salary is solved exactly rather than rounded, so the reconciliation ties to the cent.
AVG_SALARY_2024 = (F990[2024]["officer_comp"] + F990[2024]["other_salaries"]) / FTE_2024
PAYROLL_TAX_RATE = 0.0782         # DERIVED from FY2023 filing
BENEFITS_RATE = 0.1000            # SYNTHETIC
FULLY_LOADED_FACTOR = 1 + PAYROLL_TAX_RATE + BENEFITS_RATE

PERSONNEL_METHOD = (
    "FY2024 salaries of $10,260,344 (officer compensation plus other salaries, both public) divided "
    "by an assumed 121 FTE gives an average salary of $84,796, which is plausible for a national "
    "workforce-development nonprofit with a professional staff. Fully loaded cost applies a 7.82% "
    "payroll tax rate derived from the FY2023 filing and an assumed 10% benefits load."
)

# --------------------------------------------------------------------------
# SYNTHETIC - scenario drivers
#
# The central judgement of this model: the Base case assumes the humanitarian
# response surge normalizes. Foundation and government funding decline even in
# Base. That is what makes the forecast worth building - the organization has a
# cost base sized to funding that is unlikely to repeat at the same level.
# --------------------------------------------------------------------------

SCENARIOS = ["Downside", "Base", "Upside"]

SCENARIO_DRIVERS = {
    # driver key: (Downside, Base, Upside)
    "growth_individual":      (-0.120,  0.040,  0.120),
    "growth_corporate":       (-0.150,  0.030,  0.140),
    "growth_foundation":      (-0.220, -0.060,  0.080),
    "growth_government":      (-0.300, -0.100,  0.050),
    "growth_events":          (-0.100,  0.020,  0.090),
    "growth_program_revenue": (-0.050,  0.080,  0.180),
    "inflation_cpp":           (0.070,  0.035,  0.020),
    "inflation_personnel":     (0.050,  0.040,  0.035),
    "growth_participants":    (-0.080,  0.060,  0.160),
    "investment_return":       (0.018,  0.038,  0.055),
}

SCENARIO_DRIVER_LABELS = {
    "growth_individual": "Individual donation growth",
    "growth_corporate": "Corporate donation growth",
    "growth_foundation": "Foundation grant growth",
    "growth_government": "Government funding growth",
    "growth_events": "Fundraising event growth",
    "growth_program_revenue": "Program service revenue growth",
    "inflation_cpp": "Cost per participant inflation",
    "inflation_personnel": "Personnel cost inflation",
    "growth_participants": "Participant growth",
    "investment_return": "Return on investment balance",
}

SCENARIO_NARRATIVE = {
    "Base": (
        "Humanitarian-response funding normalizes gradually. Foundation grants decline 6% a year and "
        "government funding 10% a year as emergency appropriations lapse, partly offset by growth in "
        "individual, corporate and earned revenue. Participants grow 6% a year and cost per "
        "participant rises 3.5%. Management holds headcount roughly flat in real terms."
    ),
    "Upside": (
        "The organization converts surge-period relationships into durable multi-year commitments. "
        "Foundation funding resumes growth at 8% and government funding stabilizes. Corporate "
        "partnerships scale with the employer network. Participants grow 16% a year while cost per "
        "participant inflation stays at 2% through digital channel mix shift."
    ),
    "Downside": (
        "Emergency appropriations lapse without replacement and a major foundation cycles out. "
        "Government funding falls 30% a year and foundation grants 22%. Cost per participant "
        "inflation runs at 7% because the fixed cost base does not shrink as fast as volume. This is "
        "the case the reserve exists for, and the case the minimum-funding analysis is built around."
    ),
}

# Splitting a category growth driver into its two components, so the forecast
# obeys Donation Revenue = Donor Count x Average Gift rather than applying a
# growth rate to a revenue line directly.
AVG_GIFT_INFLATION = 0.020
AVG_GIFT_INFLATION_NOTE = (
    "Average gift is assumed to grow 2.0% a year with general giving inflation. The remainder of "
    "each category's growth driver is carried by donor count: donor growth = (1 + category growth) / "
    "(1 + 2.0%) - 1. This keeps both halves of the revenue driver visible and independently "
    "stressable rather than collapsing them into a single percentage."
)

# Semi-variable behaviour of the two non-program functional expense pools.
# This is what produces fixed-cost absorption in the forecast: as programs grow,
# support costs grow more slowly, and the program expense ratio improves.
MG_VARIABLE_SHARE = 0.35
FR_VARIABLE_SHARE = 0.50
SEMI_VARIABLE_NOTE = (
    "Management and general is modelled as 35% variable with program scale and 65% fixed; "
    "fundraising as 50% variable with contribution scale. Both also carry personnel inflation. "
    "The consequence is real and worth stating: in the Downside case, support costs do not fall "
    "as fast as revenue, so the program expense ratio deteriorates exactly when the organization "
    "can least afford the optics of it."
)

# Non-personnel operating expense drivers (share of non-personnel spend)
OPEX_CATEGORIES = {
    "Marketing and outreach": 0.185,
    "Technology and platforms": 0.235,
    "Facilities and occupancy": 0.210,
    "Travel and convening": 0.095,
    "Professional services": 0.150,
    "Other operating": 0.125,
}

# --------------------------------------------------------------------------
# SYNTHETIC - budget for FY2024, used for budget vs actual
#
# The FY2024 actuals are real. The budget they are compared against is a
# reconstruction: what a board would plausibly have approved in late FY2023 for
# FY2024, given what was known then. Building it this way means the variances
# are the interesting part - they show what the surge actually did to the plan.
# --------------------------------------------------------------------------

BUDGET_2024_METHOD = (
    "No public FY2024 budget exists. This is a reconstruction of what the board would plausibly have "
    "approved in late FY2023: revenue planned to grow 12% off the FY2023 actual, reflecting caution "
    "about whether the surge would hold, with expenses planned to grow 26% as the organization "
    "deployed reserves into program capacity. Actuals then came in ahead of plan on revenue and "
    "behind on spending - which is precisely the variance pattern worth explaining. Clearly labelled "
    "as a simulation throughout."
)

BUDGET_2024_REVENUE_GROWTH = 0.120
BUDGET_2024_EXPENSE_GROWTH = 0.260
BUDGET_2024_PARTICIPANT_GROWTH = 0.180

# --------------------------------------------------------------------------
# SYNTHETIC - resource allocation problem
# --------------------------------------------------------------------------

ALLOCATION_BUDGET = 1_000_000
ALLOCATION_NOTE = (
    "An incremental $1,000,000 of unrestricted funding, on top of the existing program budget, to be "
    "allocated across the four programs to maximize expected additional successful outcomes. Framed "
    "as incremental rather than total because that is the decision management actually faces: the "
    "existing program base is largely committed."
)

ALLOCATION_CONSTRAINTS = {
    "min_share_per_program": 0.05,     # no program starved below 5% of the increment
    "max_share_per_program": 0.55,     # no program takes more than 55%
    "admin_load": 0.12,                # 12% of the increment must fund administrative support
    "reserve_floor_months": 6.0,       # liquid reserves may not fall below 6 months
    "staff_hours_available": 42_000,   # incremental program-delivery hours available
    "min_avg_salary_gain": 48_000,     # quality floor on the average outcome funded
}

ALLOCATION_OBJECTIVES = {
    "outcomes": "Maximize the number of additional job placements",
    "earnings": "Maximize the total first-year earnings gain generated",
    "balanced": "Maximize placements subject to an average earnings-gain floor per placement",
}

ALLOCATION_OBJECTIVE_NOTE = (
    "The allocation is solved three times under three different objectives, because the answer "
    "changes completely depending on which one management picks, and that is the finding. "
    "Maximizing placement count pushes money into the cheapest channel. Maximizing earnings gain "
    "pushes it into the most expensive one. Neither is wrong; they answer different questions. "
    "The balanced case makes the trade-off explicit by holding placement count as the objective "
    "while requiring that the average placement funded clears a $48,000 earnings-gain floor."
)

# --------------------------------------------------------------------------
# SROI - deliberately conservative, heavily caveated
# --------------------------------------------------------------------------

SROI_PARAMS = {
    "benefit_horizon_years": 3,
    "discount_rate": 0.03,
    "attribution_rate": 0.55,      # share of earnings gain attributable to the program
    "deadweight_rate": 0.25,       # share that would have occurred anyway
    "dropoff_rate": 0.20,          # annual decay of the attributable benefit
    "persistence_rate": 0.78,      # share still employed at 12 months
}

SROI_CAVEAT = (
    "Social return figures are the least reliable numbers in this project and are reported last, "
    "with a sensitivity range rather than a point estimate. Every parameter below - attribution, "
    "deadweight, drop-off, persistence, horizon - is an assumption, and the headline ratio moves by "
    "more than a factor of two across plausible values for them. The cost-effectiveness measures in "
    "06_Program_Economics (cost per participant, cost per successful outcome) are the numbers "
    "management should actually steer on, because they rest on far fewer assumptions."
)


# ==========================================================================
# Build the assumptions register
# ==========================================================================

def build_register() -> Register:
    r = Register()
    A = lambda **kw: r.add(Input(**kw))

    # -- Organization ------------------------------------------------------
    sec = "1. Organization and scope"
    A(key="org_name", label="Anchor organization", value=ORG_NAME, unit="text", tier=PUBLIC,
      source=SRC_990, section=sec, fmt="@",
      note="Selected for eleven years of machine-readable 990 data, a published impact record, and a "
           "genuinely unresolved financial planning question.")
    A(key="org_ein", label="EIN", value=ORG_EIN, unit="text", tier=PUBLIC, source=SRC_990,
      section=sec, fmt="@", note="Employer Identification Number as filed.")
    A(key="as_of", label="Analysis as-of", value=AS_OF, unit="text", tier=PUBLIC, source=SRC_990,
      section=sec, fmt="@",
      note="Most recent public filing. FY2025 in the forecast is a current-year estimate, not an actual.")
    A(key="cn_rating", label="Charity Navigator rating", value=CHARITY_NAVIGATOR_RATING, unit="text",
      tier=PUBLIC, source=SRC_CN, section=sec, fmt="@",
      note="Used as a sanity bound on the modelled functional expense split.")
    A(key="base_year", label="Model base year", value=BASE_YEAR, unit="fiscal year", tier=PUBLIC,
      source=SRC_990, section=sec, fmt="0",
      note="Last year with filed actuals. All forecasts step forward from here.")

    # -- Historical financials --------------------------------------------
    sec = "2. Historical financials (Form 990)"
    for y in HIST_YEARS:
        d = F990[y]
        pub_2024 = (y != 2024)
        A(key=f"rev_{y}", label=f"FY{y} total revenue", value=d["total_revenue"], unit="USD",
          tier=PUBLIC, source=SRC_990, section=sec, fmt="#,##0",
          note="Form 990 Part I, total revenue.")
        A(key=f"exp_{y}", label=f"FY{y} total functional expenses", value=d["total_expense"], unit="USD",
          tier=PUBLIC, source=SRC_990, section=sec, fmt="#,##0",
          note="Form 990 Part I, total functional expenses.")
        A(key=f"assets_{y}", label=f"FY{y} total assets", value=d["total_assets"], unit="USD",
          tier=PUBLIC, source=SRC_990, section=sec, fmt="#,##0", note="Form 990 Part I, end of year.")
        A(key=f"liab_{y}", label=f"FY{y} total liabilities", value=d["total_liab"], unit="USD",
          tier=PUBLIC if pub_2024 else DERIVED, source=SRC_990, section=sec, fmt="#,##0",
          note="Form 990 Part I, end of year." if pub_2024
               else "Derived as total assets less net assets, both public.")
        A(key=f"contrib_{y}", label=f"FY{y} contributions and grants", value=d["contributions"],
          unit="USD", tier=PUBLIC if pub_2024 else DERIVED, source=SRC_990, section=sec, fmt="#,##0",
          note="Form 990 Part VIII line 1h." if pub_2024 else REVENUE_MIX_2024_METHOD)
        A(key=f"progrev_{y}", label=f"FY{y} program service revenue", value=d["program_rev"],
          unit="USD", tier=PUBLIC if pub_2024 else DERIVED, source=SRC_990, section=sec, fmt="#,##0",
          note="Form 990 Part VIII line 2g." if pub_2024 else REVENUE_MIX_2024_METHOD)
        A(key=f"invinc_{y}", label=f"FY{y} investment and other income", value=d["invest_inc"],
          unit="USD", tier=PUBLIC if pub_2024 else DERIVED, source=SRC_990, section=sec, fmt="#,##0",
          note="Form 990 Part VIII line 3." if pub_2024 else REVENUE_MIX_2024_METHOD)
        A(key=f"salaries_{y}", label=f"FY{y} salaries and officer compensation",
          value=d["officer_comp"] + d["other_salaries"], unit="USD", tier=DERIVED, source=SRC_990,
          section=sec, fmt="#,##0",
          note="Officer compensation plus other salaries and wages, both public.")
        A(key=f"payrolltax_{y}", label=f"FY{y} payroll taxes", value=d["payroll_tax"], unit="USD",
          tier=PUBLIC if pub_2024 else DERIVED, source=SRC_990, section=sec, fmt="#,##0",
          note="Form 990 Part IX line 10." if pub_2024 else PAYROLL_TAX_2024_METHOD)

    A(key="net_assets_2024", label="FY2024 net assets", value=NET_ASSETS_2024_PUBLIC, unit="USD",
      tier=PUBLIC, source=SRC_990, section=sec, fmt="#,##0",
      note="Public. Used as the control total against which derived liabilities are backed out.")

    # -- Impact ------------------------------------------------------------
    sec = "3. Published impact metrics"
    for y, d in IMPACT_PUBLIC.items():
        A(key=f"impact_participants_{y}", label=f"{y} participants served", value=d["participants"],
          unit="people", tier=PUBLIC, source=d["source"], section=sec, fmt="#,##0", note=d["note"])
        A(key=f"impact_placements_{y}", label=f"{y} job placements", value=d["placements"],
          unit="people", tier=PUBLIC, source=d["source"], section=sec, fmt="#,##0", note=d["note"])
        A(key=f"impact_salary_{y}", label=f"{y} average starting salary", value=d["avg_starting_salary"],
          unit="USD", tier=PUBLIC, source=d["source"], section=sec, fmt="#,##0", note=d["note"])
    A(key="impact_salary_gain_2025", label="2025 average salary gain",
      value=IMPACT_PUBLIC[2025]["avg_salary_gain"], unit="USD", tier=PUBLIC, source=SRC_IMPACT_2025,
      section=sec, fmt="#,##0",
      note="Used as an external calibration check on the synthetic program mix.")
    A(key="participant_interpolation", label="Interpolation method for unpublished years",
      value=PARTICIPANT_INTERPOLATION_METHOD, unit="text", tier=DERIVED,
      source=f"{SRC_IMPACT_2022}; {SRC_IMPACT_2025}", section=sec, fmt="@",
      note="Two real observations, constant growth between them. No free parameters.")

    # -- Benchmarks --------------------------------------------------------
    sec = "4. Sector benchmarks"
    for key, label, val, unit, src, note in BENCHMARKS:
        A(key=key, label=label, value=val, unit=unit, tier=BENCHMARK, source=src, section=sec,
          fmt="0.00" if "x" in unit or "months" in unit else "0.0%", note=note)

    # -- Modelled structure ------------------------------------------------
    sec = "5. Functional expense allocation (modelled)"
    for y in HIST_YEARS:
        p, m, f = FUNCTIONAL_SPLIT[y]
        A(key=f"fsplit_prog_{y}", label=f"FY{y} program services share", value=p, unit="% of expenses",
          tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="0.0%", note=FUNCTIONAL_SPLIT_METHOD)
        A(key=f"fsplit_mg_{y}", label=f"FY{y} management and general share", value=m,
          unit="% of expenses", tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="0.0%",
          note=FUNCTIONAL_SPLIT_METHOD)
        A(key=f"fsplit_fr_{y}", label=f"FY{y} fundraising share", value=f, unit="% of expenses",
          tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="0.0%", note=FUNCTIONAL_SPLIT_METHOD)

    sec = "6. Revenue composition (modelled)"
    for cat, share in REVENUE_MIX.items():
        A(key=f"revmix_{cat}", label=f"{cat}, share of contributions", value=share,
          unit="% of contributions", tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="0.0%",
          note=REVENUE_MIX_METHOD)
    for cat, d in DONOR_DRIVERS_2024.items():
        A(key=f"donors_{cat}", label=f"{cat}, FY2024 donor count", value=d["donors"], unit="donors",
          tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="#,##0", note=DONOR_DRIVER_METHOD)
        A(key=f"avggift_{cat}", label=f"{cat}, FY2024 average gift", value=d["avg_gift"], unit="USD",
          tier=DERIVED, source=SRC_ANALYST, section=sec, fmt="#,##0.00",
          note="Solved so that donors x average gift equals modelled category revenue.")
        A(key=f"retention_{cat}", label=f"{cat}, donor retention", value=d["retention"], unit="%",
          tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="0.0%",
          note="Drives donor-count roll-forward in the forecast.")
    A(key="largest_funder_share", label="Largest single funder, share of revenue",
      value=LARGEST_FUNDER_SHARE, unit="% of revenue", tier=SYNTHETIC, source=SRC_ANALYST,
      section=sec, fmt="0.0%", note=FUNDER_CONCENTRATION_NOTE)
    A(key="top5_funder_share", label="Top five funders, share of revenue", value=TOP5_FUNDER_SHARE,
      unit="% of revenue", tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="0.0%",
      note=FUNDER_CONCENTRATION_NOTE)

    sec = "7. Program portfolio (modelled)"
    for prog in PROGRAMS:
        d = PROGRAM_BASE[prog]
        A(key=f"prog_participants_{prog}", label=f"{prog}: FY2024 participants", value=d["participants"],
          unit="people", tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="#,##0",
          note="Sums with the other three programs to the derived FY2024 total of 10,791.")
        A(key=f"prog_placements_{prog}", label=f"{prog}: FY2024 successful outcomes",
          value=d["placements"], unit="people", tier=SYNTHETIC, source=SRC_ANALYST, section=sec,
          fmt="#,##0", note="Sums to the derived FY2024 total of 1,494 placements.")
        A(key=f"prog_cost_{prog}", label=f"{prog}: FY2024 program cost", value=d["program_cost"],
          unit="USD", tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="#,##0",
          note="Allocated within the modelled FY2024 program expense pool; the four sum to it exactly.")
        A(key=f"prog_completion_{prog}", label=f"{prog}: completion rate", value=d["completion_rate"],
          unit="%", tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="0.0%",
          note="Share of enrolled participants who complete the program.")
        A(key=f"prog_gain_{prog}", label=f"{prog}: average salary gain per outcome",
          value=d["salary_gain"], unit="USD", tier=SYNTHETIC, source=SRC_ANALYST, section=sec,
          fmt="#,##0", note=PROGRAM_CALIBRATION_NOTE)
        A(key=f"prog_hours_{prog}", label=f"{prog}: staff hours per participant",
          value=d["staff_hours_per_participant"], unit="hours", tier=SYNTHETIC, source=SRC_ANALYST,
          section=sec, fmt="#,##0.0", note="Drives the staff capacity constraint in the allocation model.")
        A(key=f"prog_directshare_{prog}", label=f"{prog}: direct cost share",
          value=d["direct_cost_share"], unit="% of program cost", tier=SYNTHETIC, source=SRC_ANALYST,
          section=sec, fmt="0.0%", note="Remainder is allocated indirect cost.")
        A(key=f"prog_maxpart_{prog}", label=f"{prog}: maximum servable participants",
          value=d["max_participants"], unit="people", tier=SYNTHETIC, source=SRC_ANALYST, section=sec,
          fmt="#,##0", note="Demand and delivery ceiling; binds the allocation model.")
        A(key=f"prog_minfund_{prog}", label=f"{prog}: minimum viable funding", value=d["min_funding"],
          unit="USD", tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="#,##0",
          note="Below this the program cannot retain its delivery team.")

    sec = "8. Personnel (modelled)"
    A(key="fte_2024", label="FY2024 full-time equivalents", value=FTE_2024, unit="FTE",
      tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="#,##0", note=PERSONNEL_METHOD)
    A(key="avg_salary_2024", label="FY2024 average salary", value=AVG_SALARY_2024, unit="USD",
      tier=DERIVED, source=SRC_990, section=sec, fmt="#,##0",
      note="Public salary total divided by assumed FTE.")
    A(key="payroll_tax_rate", label="Payroll tax rate", value=PAYROLL_TAX_RATE, unit="% of salary",
      tier=DERIVED, source=SRC_990, section=sec, fmt="0.00%",
      note="FY2023 payroll tax divided by FY2023 salaries, both public.")
    A(key="benefits_rate", label="Benefits load", value=BENEFITS_RATE, unit="% of salary",
      tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="0.0%",
      note="Health, retirement and other benefits as a share of salary.")
    A(key="staff_hours_per_fte", label="Productive hours per FTE per year", value=STAFF_HOURS_PER_FTE,
      unit="hours", tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="#,##0",
      note="2,080 gross hours less leave, training and non-delivery time.")
    A(key="program_fte_share", label="Share of FTE in program delivery", value=PROGRAM_FTE_SHARE,
      unit="%", tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="0.0%",
      note="Balance sits in administration and fundraising.")
    A(key="program_hours_reconciliation", label="Staff hours reconciliation",
      value=PROGRAM_HOURS_METHOD, unit="text", tier=SYNTHETIC, source=SRC_ANALYST, section=sec,
      fmt="@", note="Ties the capacity constraint to the personnel cost line.")

    sec = "8b. Cost behaviour (modelled)"
    A(key="avg_gift_inflation", label="Average gift inflation", value=AVG_GIFT_INFLATION, unit="% per year",
      tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="0.0%", note=AVG_GIFT_INFLATION_NOTE)
    A(key="mg_variable_share", label="Management and general, variable share",
      value=MG_VARIABLE_SHARE, unit="% variable", tier=SYNTHETIC, source=SRC_ANALYST, section=sec,
      fmt="0.0%", note=SEMI_VARIABLE_NOTE)
    A(key="fr_variable_share", label="Fundraising, variable share", value=FR_VARIABLE_SHARE,
      unit="% variable", tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="0.0%",
      note=SEMI_VARIABLE_NOTE)
    for cat, share in OPEX_CATEGORIES.items():
        A(key=f"opex_{cat}", label=f"{cat}, share of non-personnel spend", value=share,
          unit="% of non-personnel", tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="0.0%",
          note="Used to split the non-personnel expense plug into reportable operating categories.")

    sec = "9. Balance sheet composition (modelled)"
    for k, v in BALANCE_SHEET_2024.items():
        A(key=f"bs_{k}", label=f"FY2024 {k}", value=v, unit="USD", tier=SYNTHETIC, source=SRC_ANALYST,
          section=sec, fmt="#,##0", note=BALANCE_SHEET_METHOD)

    sec = "10. Scenario drivers (modelled)"
    for dk, (down, base, up) in SCENARIO_DRIVERS.items():
        lbl = SCENARIO_DRIVER_LABELS[dk]
        for sname, val in (("Downside", down), ("Base", base), ("Upside", up)):
            A(key=f"drv_{dk}_{sname}", label=f"{lbl} - {sname}", value=val, unit="% per year",
              tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="0.0%",
              note=SCENARIO_NARRATIVE[sname])

    sec = "11. Budget reconstruction (modelled)"
    A(key="budget_rev_growth", label="FY2024 budgeted revenue growth", value=BUDGET_2024_REVENUE_GROWTH,
      unit="%", tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="0.0%", note=BUDGET_2024_METHOD)
    A(key="budget_exp_growth", label="FY2024 budgeted expense growth", value=BUDGET_2024_EXPENSE_GROWTH,
      unit="%", tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="0.0%", note=BUDGET_2024_METHOD)
    A(key="budget_part_growth", label="FY2024 budgeted participant growth",
      value=BUDGET_2024_PARTICIPANT_GROWTH, unit="%", tier=SYNTHETIC, source=SRC_ANALYST, section=sec,
      fmt="0.0%", note=BUDGET_2024_METHOD)

    sec = "12. Resource allocation (modelled)"
    A(key="alloc_budget", label="Incremental funding to allocate", value=ALLOCATION_BUDGET, unit="USD",
      tier=SYNTHETIC, source=SRC_ANALYST, section=sec, fmt="#,##0", note=ALLOCATION_NOTE)
    for k, v in ALLOCATION_CONSTRAINTS.items():
        A(key=f"alloc_{k}", label=f"Constraint: {k.replace('_', ' ')}", value=v,
          unit="varies", tier=SYNTHETIC, source=SRC_ANALYST, section=sec,
          fmt="0.0%" if "share" in k or "load" in k else "#,##0.0",
          note="Binds the optimization in 07_Resource_Allocation_Model.")

    sec = "13. Social return parameters (modelled, low confidence)"
    for k, v in SROI_PARAMS.items():
        A(key=f"sroi_{k}", label=k.replace("_", " ").capitalize(), value=v, unit="varies",
          tier=SYNTHETIC, source=SRC_ANALYST, section=sec,
          fmt="#,##0" if "years" in k else "0.0%", note=SROI_CAVEAT)

    return r


REGISTER = build_register()

if __name__ == "__main__":
    print(f"{len(REGISTER)} inputs registered")
    for tier, n in REGISTER.tier_counts().items():
        print(f"  {tier:<10} {n}")
