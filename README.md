# Nonprofit Financial Planning, Scenario Forecasting & Impact Measurement

An FP&A model for a real US nonprofit, built entirely from public data plus
clearly labelled analyst assumptions. Excel is the deliverable; Python builds and
verifies it.

**Anchor organization:** Upwardly Global (EIN 94-3346127) — a 501(c)(3) helping
work-authorized immigrants and refugees with professional credentials re-enter
skilled employment.
**Analysis as of:** FY2024 Form 990, filed 17 October 2025.

> **Basis of preparation.** This is an independent analysis prepared from public
> information for portfolio and educational purposes. It is not affiliated with,
> endorsed by, or reviewed by Upwardly Global, and it is not financial advice.
> Historical financial totals are taken verbatim from IRS Form 990 filings.
> Program-level detail and all forward-looking figures are analyst assumptions
> and do not represent the organization's plans, budget, or guidance.

---

## The business question

> How should a nonprofit allocate limited funding across its programs while
> remaining financially sustainable and maximising measurable impact?

For this organization the question arrives in an unusual form. It is not short of
money. Six consecutive operating surpluses have built net assets to **$39.9M
against annual expenses of $18.3M** — 2.18× cover and roughly 24 months of liquid
runway. It could absorb a permanent 32% cut in contributions without cutting a
single program inside five years.

The planning problem is that a revenue surge driven by circumstances outside the
organization's control has left it with a cost base sized to funding that is
unlikely to repeat — and a balance sheet large enough to disguise how long that
has been true.

## What the analysis found

**1. The growth was episodic; the cost base is not.**
Revenue compounded at 33.5% a year from FY2019 to FY2024 (from $6.75M to $28.6M),
almost entirely through contributions tied to the Afghanistan and Ukraine
displacement responses. Expenses compounded at 27.0%. The gap produced $36.7M of
cumulative surplus. On Base case assumptions — where institutional funding
normalises rather than collapses — the organization **crosses into structural
deficit in FY2029** and reaches its reserve floor around FY2033. Closing the gap
needs contributions roughly **6.3% above trajectory**.

**2. Reach grew; placements did not.**
FY2024 served **503 more participants than planned and produced 29 fewer
successful outcomes.** The outcome rate fell from 14.8% to 13.9% as growth came
through lower-converting channels. Decomposed, 76% of the program cost overrun
was volume and 24% was rate — the organization served more people, less
effectively. This is a mix problem, not a cost problem, and cost discipline will
not fix it.

**3. The programs cannot be ranked on one measure.**
Digital Learning produces a placement for $6,667; Credential Support costs
$13,548 but delivers 2.4× the earnings gain per person placed. Optimising an
incremental $1M for **placement count** and for **earnings gain** produce almost
opposite allocations — 112 placements at $45,898 average gain, versus 85
placements at $65,879. The recommended balanced allocation gives up 2.8% of
placement count to lift average gain by 5%, and beats a pro-rata roll-forward by
21 placements a year.

**4. Stewardship ratios are not where the risk is.**
Program expense ratio 78.5% (BBB minimum 65%); fundraising cost 6.7¢ per dollar
raised (BBB ceiling 35¢); four-star Charity Navigator rating. All comfortable,
none informative. The measure that carries a warning is **accumulation**: net
assets moved from 0.81× to 2.18× annual expenses in six years, against a BBB
Standard 10 ceiling of 3×.

---

## Deliverables

**The Excel models are the deliverable.** Read them in order — each builds on the
one before.

| `excel-models/` | What it contains |
|---|---|
| `01_Assumptions.xlsx` | All 216 model inputs, each tagged `PUBLIC` / `DERIVED` / `BENCHMARK` / `SYNTHETIC` with its source and rationale. Includes 16 live reconciliation checks proving every modelled split ties to a filed control total. **Start here.** |
| `02_Financial_Model.xlsx` | The centrepiece. Six years of historical analysis and a driver-based three-year forecast, with a dashboard. Every calculated cell is a live formula. |
| `03_Budget_vs_Actual.xlsx` | FY2024 actuals against a clearly-labelled reconstructed budget, with volume/rate variance decomposition and written management commentary. |
| `04_Scenario_Model.xlsx` | Base / Upside / Downside on a `CHOOSE`-driven scenario switch, five-year horizon, plus funding sensitivity, largest-funder loss, and two distinct minimum-funding thresholds. |
| `05_Program_Economics.xlsx` | Per-program cost per participant, cost per outcome, completion and outcome rates, capacity utilisation, and a deliberately caveated SROI section reported as a range. |
| `06_Resource_Allocation.xlsx` | Solver-ready optimisation of $1M across four programs under three competing objectives, with a quality-floor trade-off frontier. |

| Supporting | What it contains |
|---|---|
| `reports/Management_Report.pdf` | Seven-page management report answering the six recommendation questions. |
| `reports/Executive_Summary.pdf` | Two-page summary for a board or hiring manager. |
| `powerbi/` | A validated star schema (7 dimensions, 9 fact tables), a 98-measure DAX library, and a page-by-page build guide. `PowerBI_Data_Model.xlsx` holds all sixteen tables in one workbook for importing in the browser; the CSVs in `powerbi/data/` are the same data for Power BI Desktop's folder connector. |
| `data/` | Raw Form 990 extracts, published impact metrics, sector benchmarks, and processed analysis outputs as CSV. `SOURCES.md` records every source and — importantly — what could not be obtained. |
| `src/`, `tests/` | The build and verification harness. See [Repository layout](#repository-layout). |

> **A note on the language bar.** GitHub reports this repository as Python,
> because GitHub's language detection has no category for `.xlsx` files — it
> cannot see the six workbooks at all. The Python exists to generate and verify
> the Excel models, not the other way round.

## Modelling decisions worth defending

A few choices depart from the obvious approach. Each was made deliberately.

**The Base case assumes institutional funding declines.** Foundation grants fall
6% a year and government funding 10%. A base case that extrapolated the 2021–2024
surge would be forecasting the past. This single choice is what makes the
forecast worth building.

**The horizon is five years, not three.** On a three-year view every scenario for
this organization looks safe, because reserves absorb almost any shock over that
period. A three-year model would report "no problem" for a structural deficit
that is serious by year six.

**Runway is reported on two bases and never mixed.** Cash alone gives 11.1 months;
cash plus short-term investments gives 23.8. The gap is roughly a year of
operations, so quoting one without saying which is a mistake.

**Funding cuts force program cuts.** Without this, a revenue shock produces an
eternal deficit no board would permit, and beneficiary counts never move. With
the reserve floor enforced, a large enough shock contracts program spending — and
the participants and outcomes forgone become the number the sensitivity analysis
actually reports.

**The allocation problem is solved three times.** Presenting one "optimal"
allocation would hide the only decision that genuinely belongs to management.
The objectives disagree, and the disagreement is the finding.

**SROI is reported last, as a range.** At the stated parameters the portfolio
ratio is 5.24×. Varying just two of six parameters moves it between **1.61× and
9.38×**. There is no counterfactual and no comparison group behind the
attribution assumptions. Cost per successful outcome rests on far fewer
assumptions and is what the report leads on.

## How the honesty is enforced

The project's central discipline: **the composition is assumed; the totals are
real.** Wherever a public control total exists, the modelled split must reconcile
to it exactly.

- The four programs' costs sum to the modelled program expense pool, which is a
  fixed share of the filed total expense figure.
- The five revenue categories sum to filed contributions.
- Balance sheet lines sum to filed total assets.
- Program delivery hours tie to the staffing model that drives the personnel line.
- Participant and outcome counts sum to totals interpolated between two published
  observations.

One example of what this catches: Form 990 reports gross contributions, program
service revenue and investment income, but states total revenue **net** of
fundraising-event costs and asset-sale losses. The gross lines therefore sum to
more than reported revenue — by $34k to $135k a year. The model carries that
difference on its own line rather than quietly overstating revenue. The
discrepancy was found by a reconciliation check, not by reading the form.

## Verification

139 automated tests. Every workbook is recalculated headlessly with LibreOffice —
which forces evaluation of formulas `openpyxl` writes but never computes — and
the results are checked against an independent Python implementation of the same
model.

```bash
python -m pytest tests -q     # 139 passed
```

The tests confirm: no formula errors in any workbook; Excel and Python agree to 1
part in 10⁶ across every historical, forecast and program line; the scenario
switch correctly re-runs all three cases; the variance bridge reconciles; the
Excel Solver objective matches the CBC linear-programming solution; the cash
roll-forward closes; scenarios are correctly ordered; the reserve floor is
actually held; and the minimum-viable-funding threshold is a real boundary
(holding just inside it, breaking just outside).

Two independent implementations agreeing is a considerably stronger claim than
one implementation being self-consistent.

## Running it

```bash
pip install -r requirements.txt
python src/build_all.py       # rebuilds every deliverable from source (~1s)
python -m pytest tests -q     # 139 tests verify them
```

Requires LibreOffice for the verification step (`soffice` on PATH). The build
itself does not.

## Repository layout

```
excel-models/       The deliverable — six workbooks, live formulas throughout
reports/            Management report and executive summary, as PDFs
powerbi/            Star schema, DAX measure library, build guide
data/
  raw/              Form 990 extracts, impact metrics, sector benchmarks
  processed/        Analysis outputs as CSV
  SOURCES.md        Every source, and what could not be obtained

src/
  inputs.py         Single source of truth — every number, with provenance
  provenance.py     Provenance tagging and the assumptions register
  engine.py         Reference implementation: historical, forecast, scenarios,
                    sensitivity, budget variance, program economics, allocation LP
  styles.py         Shared openpyxl formatting
  common_sheets.py  Source-data sheet shared across workbooks
  build_0*.py       One builder per deliverable
  build_all.py      Rebuilds everything
  recalc.py         LibreOffice recalculation harness
tests/
  test_model_integrity.py
```

Python's role is deliberately supporting: it generates the workbooks so they are
reproducible when an assumption changes, and it provides the independent
implementation the tests check Excel against. The Excel models carry live
formulas throughout — change a driver on an inputs sheet and the forecast, the
ratios, the cash roll-forward and the dashboard all move.

## Known limitations

- **The functional expense split is modelled.** Form 990 Part IX reports program,
  administrative and fundraising expenses, but that split is not exposed in the
  machine-readable extract used here. It drives three headline ratios directly
  and is the largest single piece of judgement in the historical analysis.
- **The program portfolio is an analyst construction.** No public source breaks
  out cost, participants or outcomes by program. The totals reconcile to public
  figures; the composition cannot be validated. Sections 3 and 6 of the report
  would change if the real distribution differs materially.
- **Marginal cost is assumed constant** in the allocation model. Real expansion
  faces rising marginal cost, and the effect is strongest in the cheapest
  program — so the linear model overstates the case for concentration. The
  direction of that bias is stated rather than hidden.
- **Participant counts before FY2022 are back-cast** at the growth rate observed
  between 2022 and 2025, through a period containing the pandemic. Illustrative
  rather than estimated.
- **Funder concentration is assumed, not measured.** Schedule B does not identify
  individual funders publicly. The 18% largest-funder share is the input the
  funding-loss analysis is most sensitive to.

## Sources

- [IRS Form 990 filings via ProPublica Nonprofit Explorer, EIN 94-3346127](https://projects.propublica.org/nonprofits/organizations/943346127)
- [Upwardly Global impact reporting, 2025](https://www.upwardlyglobal.org/impact/)
- [Upwardly Global 2022 Annual Report, "By the Numbers"](https://annual-report.upwardlyglobal.org/2022/by-the-numbers/)
- [BBB Wise Giving Alliance, Standards for Charity Accountability](https://give.org/bbb-standards-for-charity-accountability)
- [Propel Nonprofits, Operating Reserves guidance](https://propelnonprofits.org/resources/operating-reserves-with-nonprofit-policy-examples/)
- [Charity Navigator rating profile, EIN 94-3346127](https://www.charitynavigator.org/ein/943346127)

Full detail, including what could not be obtained, is in
[`data/SOURCES.md`](data/SOURCES.md).
