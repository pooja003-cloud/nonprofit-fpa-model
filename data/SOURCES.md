# Sources

Every external source this project relies on, what was taken from it, and what
could not be obtained.

## Primary financial data

**IRS Form 990, Upwardly Global, EIN 94-3346127**
<https://projects.propublica.org/nonprofits/organizations/943346127>

Accessed via the ProPublica Nonprofit Explorer, which republishes the IRS
machine-readable Form 990 extracts. Fiscal years 2019 through 2024.

Taken verbatim: total revenue, total functional expenses, total assets, total
liabilities, net assets, contributions and grants, program service revenue,
investment income, officer compensation, other salaries and wages, and payroll
taxes.

Two adjustments were needed and both are documented in the workbooks:

1. For FY2024 the extract exposes total revenue, total expenses, total assets,
   net assets, officer compensation and other salaries, but not the revenue mix.
   Contributions, program service revenue and investment income for that year
   are derived from the published approximate mix of 91% contributions and 4%
   program services, and are tagged DERIVED rather than PUBLIC. Total liabilities
   for FY2024 are derived as total assets less published net assets. Payroll tax
   is derived by applying the FY2023 effective rate of 7.82%.

2. Form 990 reports gross contributions, gross program service revenue and gross
   investment income, but states total revenue net of the direct expenses of
   fundraising events and of losses on asset sales. The three gross lines
   therefore sum to more than reported total revenue in every year from FY2019 to
   FY2023, by between $34,318 and $135,303. That difference is carried on its own
   line rather than ignored. Without it the analysis would overstate revenue.

## Published impact metrics

**Upwardly Global impact reporting, 2025**
<https://www.upwardlyglobal.org/impact/>

13,350 jobseekers supported, 1,730+ placed, $67,580 average starting salary,
$58,790 average salary gain, 100+ employer partners, 87 workforce partnerships.

**Upwardly Global 2022 Annual Report, "By the Numbers"**
<https://annual-report.upwardlyglobal.org/2022/by-the-numbers/>

7,041 program participants across career-coaching and online programs, 1,116
placed in thriving-wage jobs, $66,481 average annual starting salary, $74M
estimated annual economic contribution, 300 corporate partners.

Participant and placement counts exist for 2022 and 2025 only. Intervening years
are interpolated at constant growth between the two published observations, which
fixes the series with no free parameters. The back-cast to FY2019 through FY2021
uses the same rate and is the weakest link in the chain: it assumes a growth rate
observed across 2022-2025 also held through a period containing the pandemic and
the beginning of the humanitarian response surge. It should be read as
illustrative rather than as an estimate of what happened.

## Sector benchmarks

**BBB Wise Giving Alliance, Standards for Charity Accountability**
<https://give.org/bbb-standards-for-charity-accountability>

- Standard 8: spend at least 65% of total expenses on program activities.
- Standard 9: spend no more than 35% of related contributions on fundraising.
- Standard 10: unrestricted net assets available for use should not exceed three
  times the past year's expenses or three times the current year's budget,
  whichever is higher.

**Propel Nonprofits, operating reserves guidance**
<https://propelnonprofits.org/resources/operating-reserves-with-nonprofit-policy-examples/>

A commonly used reserve goal is three to six months of expenses. At the low end,
reserves should cover at least one full payroll including taxes. At the high end,
they should not exceed two years of budget.

**Charity Navigator rating profile**
<https://www.charitynavigator.org/ein/943346127>

Four of four stars, 96% overall score, all four beacons complete. Used only as an
upper sanity bound on the modelled functional expense split: a four-star
organization is very unlikely to be sitting near the BBB minimum program ratio.

## What could not be obtained

**The Form 990 Part IX functional expense split.** Program services, management
and general, and fundraising are reported on the filed return but are not exposed
in the machine-readable extract used here, and the filing PDFs are not
retrievable through the tooling available for this project. The split is
therefore modelled, applied to the real total expense figure so the three lines
always reconcile to the filing. This is the single largest piece of judgement in
the historical analysis and it drives the program expense ratio, the
administrative ratio and the fundraising cost ratio directly.

**Program-level financial and operational data.** No public source breaks out
cost, participants or outcomes by program. The four-program portfolio used
throughout this project is an analyst construction whose totals reconcile to
public figures but whose composition cannot be validated.

**Balance sheet composition.** Total assets and net assets are filed; the split
between cash, short-term investments, receivables and fixed assets is not
available in the extract and is modelled.

**Funder-level detail.** Schedule B is not public in a form that identifies
individual funders, so funding concentration is assumed rather than measured. The
18% largest-funder share used in the sensitivity analysis is an assumption
selected to be plausible for an organization whose growth came through a small
number of large awards, and it is the input the funding-loss analysis is most
sensitive to.

**Any FY2024 budget.** No public budget exists. The budget used in
03_Budget_vs_Actual.xlsx is a reconstruction on stated assumptions and is
labelled a simulation throughout.

## Data quality note

The organization's own reporting uses at least two different all-time participant
definitions: 13,460 individuals and families with improved lives in the 2022
report against 35,000+ jobseekers on the 2025 impact page. Definitions of who
counts as served evidently changed between the two. Only single-year figures are
used in this project, and only from the two dates cited above.
