# Power BI build guide

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

98 measures in twelve groups, supplied in two forms. Use the one that matches
where you are building.

**In the browser — `measures_dax_query.dax`, all at once.** Open the semantic
model, choose **Write DAX queries**, paste the whole file, and click the
**Update model: Add new measures** link that appears above `DEFINE`. Every
measure is created in one action. This is the only sane route in the Service:
the web modelling canvas has a **New measure** button but no bulk import, and
98 definitions pasted one at a time is an hour of clicking.

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
- Text box for the commentary from `03_Budget_vs_Actual.xlsx`

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
