# Power BI build guide

A `.pbix` file is a proprietary binary and cannot be generated programmatically,
so this folder contains everything that goes inside one. Loading `data/`, wiring
the relationships below and pasting `measures.dax` reproduces the dashboard.

Power BI Desktop is Windows-only. On macOS the options are Power BI Service
(app.powerbi.com, which will consume this star schema directly), a Windows VM, or
Parallels. Nothing in this folder depends on which route is taken.

---

## 1. Load the data

Get Data > Text/CSV, and load all files from `data/`. Or Get Data > Folder to
pull the whole directory in one step.

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

Set data types explicitly after loading. Power BI will usually infer `Year` as a
whole number, which is what you want; check that `Amount`, `VariancePct`,
`Share`, `CapacityRetained` and `FundingShock` come in as decimal numbers rather
than text.

## 2. Build the relationships

All single-direction, many-to-one, from fact to dimension.

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

Paste from `measures.dax`. Twelve groups, roughly seventy measures. If you use
Tabular Editor, the file can be imported wholesale; otherwise create a blank
measure and paste one block at a time.

Ignore the `_Line` placeholder - it documents the pattern the line measures use
and is not meant to be created.

## 4. Build the pages

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

- Matrix: rows `dim_account[LineItem]`, values `Budget`, `Actual`, `Variance`,
  `Variance Pct`, conditionally formatted on `Variance Colour`
- Waterfall: `fact_variance_bridge`, category `Step`, value `Amount`, breakdown
  by `StepType` - this is the volume-versus-rate decomposition
- Card: `Variance Direction` for the selected line
- Text box for the commentary from `04_Budget_vs_Actual.xlsx`

The point of this page is the waterfall. A variance table tells a reader that
program cost was over budget; the waterfall tells them 76%
of the overrun was serving more people and the rest was unit cost.

### Page 3 - Programs

- Table: `dim_program[Program]` with `Participants`, `Outcomes`, `Outcome Rate`,
  `Cost per Participant`, `Cost per Outcome`, `Earnings Gain per Dollar`
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
- Line chart: `Participants` and `Outcomes` by year

### Page 5 - Scenarios

- Line chart: `Operating Result` by year, legend `dim_scenario[Scenario]`
- Line chart: `Liquid Runway Months` by year, legend scenario, reference line at 6
- Card: `Scenario Spread`
- Table: scenario, `Total Revenue`, `Total Expenses`, `Operating Result`,
  `Outcomes` for the final forecast year
- Text box bound to `Scenario Narrative`

### Page 6 - Funding and sensitivity

- Donut or bar: `Funding Amount` by `dim_funding_source[FundingSource]`
- Card: `Institutional Funding Share` - the concentration measure that matters
  most for this organization
- Line chart from `fact_sensitivity`: `Sensitivity Runway` by `FundingShock`,
  legend `Scenario`, reference line at 6 months
- Column chart: `Sensitivity Outcomes Lost` by `FundingShock`
- Card: `Max Sustainable Shock`

### Page 7 - Allocation

- Clustered bar from `fact_allocation`: `Allocation` by program, legend
  `ObjectiveLabel`
- Table: objective, `Additional Outcomes`, `Additional Earnings`,
  `Allocation Avg Gain`, `Outcomes Forgone vs Best`
- Line chart from `fact_quality_tradeoff`: `Outcomes` by `QualityFloor`

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

## Reference figures

These come from the FY2024 filing and should match what the dashboard shows once
built. If they do not, something is wrong with the load or the relationships.

| Measure | FY2024 |
|---|---|
| Total revenue | $28,580,411 |
| Total expenses | $18,275,646 |
| Operating result | $10,304,765 |
| Program expense ratio | 78.5% |
| Administrative ratio | 12.0% |
| Fundraising cost ratio | 6.7% |
| Net assets | $39,907,054 |
| Net asset multiple | 2.18x |
| Liquid runway | 23.8 months |
| Participants | 10,786 |
| Outcomes | 1,495 |
| Cost per outcome | $9,597 |
