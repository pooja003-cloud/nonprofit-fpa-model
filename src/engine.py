"""
Reference calculation engine.

The Excel workbooks are the deliverable and they carry live formulas. This
module computes the same thing in Python for three reasons: to verify that the
recalculated workbooks agree with an independent implementation, to generate the
Power BI star schema, and to give the written report real numbers to quote.

If Excel and this module ever disagree, one of them is wrong, and the test suite
in tests/ is what catches it.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field

import inputs as I


# ==========================================================================
# Historical analysis
# ==========================================================================

def participants_series() -> dict[int, float]:
    """
    Participant counts by fiscal year.

    Published for 2022 (7,041) and 2025 (13,350) only. Everything between is a
    constant-growth interpolation between two real observations; 2019-2021 is
    back-cast at the same rate, which is weaker and flagged as such in the notes.
    """
    p22 = I.IMPACT_PUBLIC[2022]["participants"]
    p25 = I.IMPACT_PUBLIC[2025]["participants"]
    g = (p25 / p22) ** (1 / 3) - 1
    return {y: p22 * (1 + g) ** (y - 2022) for y in I.HIST_YEARS}, g


def placements_series() -> dict[int, float]:
    p22 = I.IMPACT_PUBLIC[2022]["placements"]
    p25 = I.IMPACT_PUBLIC[2025]["placements"]
    g = (p25 / p22) ** (1 / 3) - 1
    return {y: p22 * (1 + g) ** (y - 2022) for y in I.HIST_YEARS}, g


def historical() -> dict:
    """Full historical analysis, FY2019-FY2024."""
    parts, part_g = participants_series()
    places, place_g = placements_series()

    rows = {}
    prev = None
    for y in I.HIST_YEARS:
        d = I.F990[y]
        rev, exp = d["total_revenue"], d["total_expense"]
        assets, liab = d["total_assets"], d["total_liab"]
        net_assets = assets - liab
        pshare, mshare, fshare = I.FUNCTIONAL_SPLIT[y]

        prog_exp = exp * pshare
        mg_exp = exp * mshare
        fr_exp = exp * fshare
        salaries = d["officer_comp"] + d["other_salaries"]
        personnel = salaries + d["payroll_tax"] + salaries * I.BENEFITS_RATE
        cash = assets * I.CASH_SHARE_OF_ASSETS[y]
        liquid = assets * I.LIQUID_SHARE_OF_ASSETS[y]

        r = dict(
            year=y,
            revenue=rev,
            expenses=exp,
            operating_result=rev - exp,
            operating_margin=(rev - exp) / rev,
            contributions=d["contributions"],
            program_revenue=d["program_rev"],
            investment_other=d["invest_inc"],
            revenue_netting=d["revenue_netting"],
            total_assets=assets,
            total_liabilities=liab,
            net_assets=net_assets,
            cash=cash,
            liquid_reserves=liquid,
            program_expense=prog_exp,
            mg_expense=mg_exp,
            fundraising_expense=fr_exp,
            salaries=salaries,
            personnel=personnel,
            non_personnel=exp - personnel,
            program_ratio=prog_exp / exp,
            admin_ratio=mg_exp / exp,
            fundraising_ratio=fr_exp / exp,
            fundraising_cost_ratio=fr_exp / d["contributions"],
            cost_to_raise_a_dollar=fr_exp / d["contributions"],
            fundraising_roi=d["contributions"] / fr_exp,
            personnel_ratio=personnel / exp,
            monthly_expenses=exp / 12,
            cash_runway_months=cash / (exp / 12),
            liquid_runway_months=liquid / (exp / 12),
            net_asset_months=net_assets / (exp / 12),
            net_asset_multiple=net_assets / exp,
            current_ratio=assets / liab if liab else float("nan"),
            participants=parts[y],
            placements=places[y],
            placement_rate=places[y] / parts[y],
            cost_per_participant=prog_exp / parts[y],
            cost_per_outcome=prog_exp / places[y],
            revenue_per_participant=rev / parts[y],
            contribution_concentration=I.LARGEST_FUNDER_SHARE,
        )
        if prev:
            r["revenue_growth"] = rev / prev["revenue"] - 1
            r["expense_growth"] = exp / prev["expenses"] - 1
            r["contribution_growth"] = d["contributions"] / prev["contributions"] - 1
            r["participant_growth"] = parts[y] / prev["participants"] - 1
            r["cpp_growth"] = r["cost_per_participant"] / prev["cost_per_participant"] - 1
        else:
            for k in ("revenue_growth", "expense_growth", "contribution_growth",
                      "participant_growth", "cpp_growth"):
                r[k] = None
        rows[y] = r
        prev = r

    first, last = I.HIST_YEARS[0], I.HIST_YEARS[-1]
    n = last - first
    summary = dict(
        revenue_cagr=(rows[last]["revenue"] / rows[first]["revenue"]) ** (1 / n) - 1,
        expense_cagr=(rows[last]["expenses"] / rows[first]["expenses"]) ** (1 / n) - 1,
        contribution_cagr=(rows[last]["contributions"] / rows[first]["contributions"]) ** (1 / n) - 1,
        net_asset_cagr=(rows[last]["net_assets"] / rows[first]["net_assets"]) ** (1 / n) - 1,
        participant_cagr=part_g,
        placement_cagr=place_g,
        cumulative_surplus=sum(rows[y]["operating_result"] for y in I.HIST_YEARS),
        years_in_surplus=sum(1 for y in I.HIST_YEARS if rows[y]["operating_result"] > 0),
    )
    return dict(rows=rows, summary=summary, participant_growth=part_g, placement_growth=place_g)


# ==========================================================================
# Program economics, base year
# ==========================================================================

def program_economics() -> dict:
    hist = historical()
    base = hist["rows"][I.BASE_YEAR]
    prog_pool = base["program_expense"]

    # PROGRAM_BASE costs are stated against the modelled FY2024 program pool.
    stated = sum(I.PROGRAM_BASE[p]["program_cost"] for p in I.PROGRAMS)
    scale = prog_pool / stated  # forces exact reconciliation to the modelled pool

    out = {}
    for p in I.PROGRAMS:
        d = I.PROGRAM_BASE[p]
        cost = d["program_cost"] * scale
        parts = d["participants"]
        places = d["placements"]
        completers = parts * d["completion_rate"]
        direct = cost * d["direct_cost_share"]
        hours = parts * d["staff_hours_per_participant"]
        out[p] = dict(
            program=p,
            code=I.PROGRAM_CODES[p],
            participants=parts,
            completers=completers,
            outcomes=places,
            program_cost=cost,
            direct_cost=direct,
            indirect_cost=cost - direct,
            cost_per_participant=cost / parts,
            cost_per_completer=cost / completers,
            cost_per_outcome=cost / places,
            completion_rate=d["completion_rate"],
            outcome_rate=places / parts,
            outcome_rate_of_completers=places / completers,
            salary_gain=d["salary_gain"],
            total_earnings_gain=places * d["salary_gain"],
            earnings_gain_per_dollar=places * d["salary_gain"] / cost,
            staff_hours=hours,
            staff_hours_per_participant=d["staff_hours_per_participant"],
            staff_hours_per_outcome=hours / places,
            implied_fte=hours / I.STAFF_HOURS_PER_FTE,
            max_participants=d["max_participants"],
            capacity_utilization=parts / d["max_participants"],
            min_funding=d["min_funding"],
            funding_headroom=cost - d["min_funding"],
            marginal_cost_per_outcome=cost / places,  # linear assumption, stated in the workbook
        )

    tot_parts = sum(o["participants"] for o in out.values())
    tot_places = sum(o["outcomes"] for o in out.values())
    tot_cost = sum(o["program_cost"] for o in out.values())
    tot_hours = sum(o["staff_hours"] for o in out.values())
    blended_gain = sum(o["total_earnings_gain"] for o in out.values()) / tot_places

    published_gain = I.IMPACT_PUBLIC[2025]["avg_salary_gain"]
    totals = dict(
        participants=tot_parts,
        outcomes=tot_places,
        program_cost=tot_cost,
        staff_hours=tot_hours,
        implied_program_fte=tot_hours / I.STAFF_HOURS_PER_FTE,
        implied_total_fte=tot_hours / I.STAFF_HOURS_PER_FTE / I.PROGRAM_FTE_SHARE,
        cost_per_participant=tot_cost / tot_parts,
        cost_per_outcome=tot_cost / tot_places,
        outcome_rate=tot_places / tot_parts,
        blended_salary_gain=blended_gain,
        published_salary_gain=published_gain,
        calibration_variance=blended_gain / published_gain - 1,
        program_pool_reconciled=abs(tot_cost - prog_pool) < 0.01,
        participants_vs_derived=tot_parts - base["participants"],
        outcomes_vs_derived=tot_places - base["placements"],
    )
    return dict(programs=out, totals=totals, base=base, scale=scale)


# ==========================================================================
# Social return - conservative, ranged, reported last
# ==========================================================================

def sroi(program_key: str | None = None, params: dict | None = None) -> dict:
    p = dict(I.SROI_PARAMS)
    if params:
        p.update(params)
    pe = program_economics()
    progs = pe["programs"]
    keys = [program_key] if program_key else I.PROGRAMS

    out = {}
    for k in keys:
        d = progs[k]
        gross = d["outcomes"] * d["salary_gain"]
        pv = 0.0
        for t in range(p["benefit_horizon_years"]):
            annual = gross * (1 - p["dropoff_rate"]) ** t * (p["persistence_rate"] if t else 1.0)
            pv += annual / (1 + p["discount_rate"]) ** t
        net = pv * p["attribution_rate"] * (1 - p["deadweight_rate"])
        out[k] = dict(
            gross_annual_gain=gross,
            present_value=pv,
            attributable_value=net,
            program_cost=d["program_cost"],
            sroi_ratio=net / d["program_cost"],
        )
    tot_net = sum(v["attributable_value"] for v in out.values())
    tot_cost = sum(v["program_cost"] for v in out.values())
    out["_portfolio"] = dict(attributable_value=tot_net, program_cost=tot_cost,
                             sroi_ratio=tot_net / tot_cost)
    return out


def sroi_sensitivity() -> list[dict]:
    """The point of this table is to show how unstable the headline ratio is."""
    grid = []
    for attribution in (0.35, 0.45, 0.55, 0.65, 0.75):
        for horizon in (1, 3, 5):
            r = sroi(params=dict(attribution_rate=attribution, benefit_horizon_years=horizon))
            grid.append(dict(attribution=attribution, horizon=horizon,
                             ratio=r["_portfolio"]["sroi_ratio"]))
    ratios = [g["ratio"] for g in grid]
    return dict(grid=grid, low=min(ratios), high=max(ratios), spread=max(ratios) / min(ratios))


# ==========================================================================
# Forecast
# ==========================================================================

@dataclass
class YearResult:
    year: int
    scenario: str
    revenue_by_category: dict = field(default_factory=dict)
    donors_by_category: dict = field(default_factory=dict)
    avg_gift_by_category: dict = field(default_factory=dict)
    contributions: float = 0.0
    program_revenue: float = 0.0
    investment_income: float = 0.0
    total_revenue: float = 0.0
    participants_by_program: dict = field(default_factory=dict)
    cpp_by_program: dict = field(default_factory=dict)
    cost_by_program: dict = field(default_factory=dict)
    outcomes_by_program: dict = field(default_factory=dict)
    program_expense: float = 0.0
    mg_expense: float = 0.0
    fundraising_expense: float = 0.0
    total_expense: float = 0.0
    personnel: float = 0.0
    non_personnel: float = 0.0
    fte: float = 0.0
    avg_salary: float = 0.0
    operating_result: float = 0.0
    opening_cash: float = 0.0
    closing_cash: float = 0.0
    opening_liquid: float = 0.0
    closing_liquid: float = 0.0
    opening_net_assets: float = 0.0
    closing_net_assets: float = 0.0
    participants: float = 0.0
    outcomes: float = 0.0
    contraction_applied: float = 0.0   # share of planned program spend cut to hold the floor

    def derived(self) -> dict:
        m = self.total_expense / 12
        return dict(
            program_ratio=self.program_expense / self.total_expense,
            admin_ratio=self.mg_expense / self.total_expense,
            fundraising_cost_ratio=self.fundraising_expense / self.contributions,
            personnel_ratio=self.personnel / self.total_expense,
            operating_margin=self.operating_result / self.total_revenue,
            monthly_expenses=m,
            cash_runway_months=self.closing_cash / m,
            liquid_runway_months=self.closing_liquid / m,
            net_asset_months=self.closing_net_assets / m,
            net_asset_multiple=self.closing_net_assets / self.total_expense,
            cost_per_participant=self.program_expense / self.participants,
            cost_per_outcome=self.program_expense / self.outcomes,
            outcome_rate=self.outcomes / self.participants,
        )


def _base_year_state() -> dict:
    """FY2024 starting point, tied to the filing."""
    hist = historical()
    b = hist["rows"][I.BASE_YEAR]
    pe = program_economics()

    rev_by_cat = {c: b["contributions"] * I.REVENUE_MIX[c] for c in I.REVENUE_CATEGORIES}
    donors = {c: I.DONOR_DRIVERS_2024[c]["donors"] for c in I.REVENUE_CATEGORIES}
    avg_gift = {c: rev_by_cat[c] / donors[c] for c in I.REVENUE_CATEGORIES}

    return dict(
        revenue_by_category=rev_by_cat,
        donors=donors,
        avg_gift=avg_gift,
        contributions=b["contributions"],
        program_revenue=b["program_revenue"],
        investment_income=b["investment_other"],
        total_revenue=b["revenue"],
        participants={p: pe["programs"][p]["participants"] for p in I.PROGRAMS},
        cpp={p: pe["programs"][p]["cost_per_participant"] for p in I.PROGRAMS},
        outcome_rate={p: pe["programs"][p]["outcome_rate"] for p in I.PROGRAMS},
        program_expense=b["program_expense"],
        mg_expense=b["mg_expense"],
        fundraising_expense=b["fundraising_expense"],
        total_expense=b["expenses"],
        liquid=sum(I.BALANCE_SHEET_2024[k] for k in
                   ("Cash and cash equivalents", "Short-term investments")),
        cash=I.BALANCE_SHEET_2024["Cash and cash equivalents"],
        net_assets=b["net_assets"],
        avg_salary=I.AVG_SALARY_2024,
        fte=I.FTE_2024,
        participants_total=sum(pe["programs"][p]["participants"] for p in I.PROGRAMS),
        outcomes_total=sum(pe["programs"][p]["outcomes"] for p in I.PROGRAMS),
    )


def forecast(scenario: str, funding_shock: float = 0.0, shock_year: int | None = None,
             years: list[int] | None = None, enforce_reserve_floor: bool = False,
             reserve_floor_months: float | None = None,
             expense_reduction: float = 0.0) -> list[YearResult]:
    """
    Driver-based three-year forecast.

    funding_shock applies a permanent proportional reduction to contributions
    from shock_year onward. It is how the funding sensitivity and
    minimum-viable-funding analyses are run.

    enforce_reserve_floor is the switch that makes a funding cut hurt. Without
    it, expenses are driven entirely by participant and cost drivers and a
    revenue shock simply produces a bigger deficit forever, which no board would
    permit. With it, planned program spending is cut back far enough to keep
    liquid reserves at the floor - and participants and outcomes fall with it.
    That contraction is the number the brief is really asking for when it says
    'calculate effect on program capacity and beneficiaries'.

    Support costs are treated as semi-fixed in the contraction: only their
    variable share can be cut in-year, so program delivery absorbs a
    disproportionate share of any squeeze. That is realistic and it is the
    reason the program expense ratio falls in the Downside case.
    """
    years = years or I.FCST_YEARS
    floor = reserve_floor_months if reserve_floor_months is not None \
        else I.ALLOCATION_CONSTRAINTS["reserve_floor_months"]
    drv = {k: I.SCENARIO_DRIVERS[k][I.SCENARIOS.index(scenario)] for k in I.SCENARIO_DRIVERS}
    st = _base_year_state()

    donors = dict(st["donors"])
    avg_gift = dict(st["avg_gift"])
    parts = dict(st["participants"])
    cpp = dict(st["cpp"])
    prog_rev = st["program_revenue"]
    mg = st["mg_expense"]
    fr = st["fundraising_expense"]
    cash = st["cash"]
    liquid = st["liquid"]
    net_assets = st["net_assets"]
    avg_salary = st["avg_salary"]
    prev_prog_exp = st["program_expense"]
    prev_contrib = st["contributions"]

    growth_key = {
        "Individual donations": "growth_individual",
        "Corporate donations": "growth_corporate",
        "Foundation grants": "growth_foundation",
        "Government funding": "growth_government",
        "Fundraising events": "growth_events",
    }

    results = []
    for y in years:
        # --- revenue: donor count x average gift -------------------------
        rev_by_cat, don_out, gift_out = {}, {}, {}
        for c in I.REVENUE_CATEGORIES:
            g = drv[growth_key[c]]
            gift_growth = I.AVG_GIFT_INFLATION
            donor_growth = (1 + g) / (1 + gift_growth) - 1
            donors[c] = donors[c] * (1 + donor_growth)
            avg_gift[c] = avg_gift[c] * (1 + gift_growth)
            rev_by_cat[c] = donors[c] * avg_gift[c]
            don_out[c], gift_out[c] = donors[c], avg_gift[c]

        contributions = sum(rev_by_cat.values())
        if shock_year is not None and y >= shock_year:
            contributions *= (1 - funding_shock)
            rev_by_cat = {c: v * (1 - funding_shock) for c, v in rev_by_cat.items()}

        prog_rev *= (1 + drv["growth_program_revenue"])
        investment_income = liquid * drv["investment_return"]
        total_revenue = contributions + prog_rev + investment_income

        # --- program expense: participants x cost per participant --------
        cost_by_prog, out_by_prog = {}, {}
        for p in I.PROGRAMS:
            parts[p] = min(parts[p] * (1 + drv["growth_participants"]),
                           I.PROGRAM_BASE[p]["max_participants"])
            cpp[p] = cpp[p] * (1 + drv["inflation_cpp"])
            cost_by_prog[p] = parts[p] * cpp[p]
            out_by_prog[p] = parts[p] * st["outcome_rate"][p]

        program_expense = sum(cost_by_prog.values())

        # --- semi-variable support costs ---------------------------------
        prog_growth = program_expense / prev_prog_exp - 1
        contrib_growth = contributions / prev_contrib - 1
        mg = mg * (1 + drv["inflation_personnel"]) * (1 + I.MG_VARIABLE_SHARE * prog_growth)
        fr = fr * (1 + drv["inflation_personnel"]) * (1 + I.FR_VARIABLE_SHARE * contrib_growth)

        if expense_reduction:
            program_expense *= (1 - expense_reduction)
            mg *= (1 - expense_reduction * I.MG_VARIABLE_SHARE)
            fr *= (1 - expense_reduction * I.FR_VARIABLE_SHARE)
            k = 1 - expense_reduction
            cost_by_prog = {p: v * k for p, v in cost_by_prog.items()}
            parts = {p: v * k for p, v in parts.items()}
            out_by_prog = {p: v * k for p, v in out_by_prog.items()}

        total_expense = program_expense + mg + fr

        # --- reserve floor: cut planned spending if the floor would break ---
        contraction = 0.0
        if enforce_reserve_floor:
            # closing_liquid = opening + revenue - expense >= floor/12 * expense
            max_expense = (liquid + total_revenue) / (1 + floor / 12)
            if total_expense > max_expense + 0.01:
                cuttable = program_expense + mg * I.MG_VARIABLE_SHARE + fr * I.FR_VARIABLE_SHARE
                needed = total_expense - max_expense
                k = max(0.0, 1 - needed / cuttable) if cuttable > 0 else 0.0
                contraction = 1 - k
                program_expense *= k
                mg = mg * (1 - I.MG_VARIABLE_SHARE) + mg * I.MG_VARIABLE_SHARE * k
                fr = fr * (1 - I.FR_VARIABLE_SHARE) + fr * I.FR_VARIABLE_SHARE * k
                cost_by_prog = {p: v * k for p, v in cost_by_prog.items()}
                parts = {p: v * k for p, v in parts.items()}
                out_by_prog = {p: v * k for p, v in out_by_prog.items()}
                total_expense = program_expense + mg + fr

        # --- personnel, implied from delivery hours ----------------------
        hours = sum(parts[p] * I.PROGRAM_BASE[p]["staff_hours_per_participant"] for p in I.PROGRAMS)
        fte = hours / I.STAFF_HOURS_PER_FTE / I.PROGRAM_FTE_SHARE
        avg_salary *= (1 + drv["inflation_personnel"])
        personnel = fte * avg_salary * I.FULLY_LOADED_FACTOR

        operating_result = total_revenue - total_expense
        opening_cash, opening_liq, opening_na = cash, liquid, net_assets
        liquid = liquid + operating_result
        cash = cash + operating_result * (st["cash"] / st["liquid"])
        net_assets = net_assets + operating_result

        r = YearResult(
            year=y, scenario=scenario,
            revenue_by_category=rev_by_cat, donors_by_category=don_out,
            avg_gift_by_category=gift_out,
            contributions=contributions, program_revenue=prog_rev,
            investment_income=investment_income, total_revenue=total_revenue,
            participants_by_program=dict(parts), cpp_by_program=dict(cpp),
            cost_by_program=cost_by_prog, outcomes_by_program=out_by_prog,
            program_expense=program_expense, mg_expense=mg, fundraising_expense=fr,
            total_expense=total_expense, personnel=personnel,
            non_personnel=total_expense - personnel, fte=fte, avg_salary=avg_salary,
            operating_result=operating_result,
            opening_cash=opening_cash, closing_cash=cash,
            opening_liquid=opening_liq, closing_liquid=liquid,
            opening_net_assets=opening_na, closing_net_assets=net_assets,
            participants=sum(parts.values()), outcomes=sum(out_by_prog.values()),
            contraction_applied=contraction,
        )
        results.append(r)
        prev_prog_exp, prev_contrib = program_expense, contributions

    return results


def all_scenarios() -> dict[str, list[YearResult]]:
    return {s: forecast(s) for s in I.SCENARIOS}


# ==========================================================================
# Funding sensitivity and minimum viable funding
# ==========================================================================

SENSITIVITY_HORIZON = [2025, 2026, 2027, 2028, 2029]
LONG_HORIZON = list(range(2025, 2041))


def years_to_reserve_floor(scenario: str = "Base", shock: float = 0.0,
                           floor: float | None = None) -> dict:
    """
    How many years until liquid reserves fall to the floor.

    This is the most intuitive way to express funding vulnerability for an
    organization with a very strong balance sheet. Saying 'a 20% funding cut
    leaves reserves at 12 months in FY2029' is true but hard to act on. Saying
    'a 20% funding cut exhausts the usable reserve in seven years' is the same
    fact in the unit a board thinks in.
    """
    floor = floor if floor is not None else I.ALLOCATION_CONSTRAINTS["reserve_floor_months"]
    f = forecast(scenario, funding_shock=shock, shock_year=LONG_HORIZON[0], years=LONG_HORIZON)
    for i, r in enumerate(f, start=1):
        if r.closing_liquid / (r.total_expense / 12) < floor:
            return dict(reached=True, years=i, year=r.year,
                        runway_at_that_point=r.closing_liquid / (r.total_expense / 12))
    return dict(reached=False, years=None, year=None,
                runway_at_that_point=f[-1].closing_liquid / (f[-1].total_expense / 12),
                horizon_tested=len(LONG_HORIZON))


def structural_deficit_crossover(scenario: str = "Base") -> dict:
    """First forecast year in which the organization stops breaking even."""
    f = forecast(scenario, years=LONG_HORIZON)
    for r in f:
        if r.operating_result < 0:
            return dict(crosses=True, year=r.year, result=r.operating_result,
                        years_from_base=r.year - I.BASE_YEAR)
    return dict(crosses=False, year=None, result=f[-1].operating_result)


def funding_sensitivity(scenario: str = "Base",
                        shocks: tuple = (0.0, 0.10, 0.20, 0.30, 0.40, 0.50),
                        years: list[int] | None = None) -> list[dict]:
    """
    Effect of a permanent funding reduction on results, liquidity, program
    capacity and beneficiaries.

    Run over five years rather than three. Three years understates the risk
    badly for this organization: reserves are large enough to absorb almost any
    shock for three years, so a three-year window would report 'no problem' for
    a structural deficit that is in fact fatal by year six.
    """
    years = years or SENSITIVITY_HORIZON
    ref = forecast(scenario, years=years, enforce_reserve_floor=True)
    out = []
    for s in shocks:
        f = forecast(scenario, funding_shock=s, shock_year=years[0], years=years,
                     enforce_reserve_floor=True)
        last = f[-1]
        d = last.derived()
        unconstrained = forecast(scenario, funding_shock=s, shock_year=years[0], years=years)
        out.append(dict(
            shock=s,
            fy1_result=f[0].operating_result,
            fy_last_result=last.operating_result,
            structural_result=unconstrained[-1].operating_result,
            cumulative_result=sum(x.operating_result for x in f),
            closing_liquid=last.closing_liquid,
            liquid_runway_months=d["liquid_runway_months"],
            closing_net_assets=last.closing_net_assets,
            net_asset_months=d["net_asset_months"],
            participants=last.participants,
            outcomes=last.outcomes,
            participants_lost=ref[-1].participants - last.participants,
            outcomes_lost=ref[-1].outcomes - last.outcomes,
            capacity_retained=last.participants / ref[-1].participants,
            revenue=last.total_revenue,
            expense=last.total_expense,
            contraction=max(x.contraction_applied for x in f),
            first_contraction_year=next((x.year for x in f if x.contraction_applied > 1e-9), None),
            breaches_reserve_floor=any(x.contraction_applied > 1e-9 for x in f),
            years_to_floor=years_to_reserve_floor(scenario, s),
        ))
    return out


def largest_funder_loss(scenario: str = "Base", years: list[int] | None = None) -> dict:
    """The specific question: what if the single largest funder walks."""
    years = years or SENSITIVITY_HORIZON
    share = I.LARGEST_FUNDER_SHARE
    base_state = _base_year_state()
    contrib_share = share * base_state["total_revenue"] / base_state["contributions"]
    f = forecast(scenario, funding_shock=contrib_share, shock_year=years[0], years=years,
                 enforce_reserve_floor=True)
    ref = forecast(scenario, years=years, enforce_reserve_floor=True)
    last, ref_last = f[-1], ref[-1]
    return dict(
        revenue_share_lost=share,
        contribution_share_lost=contrib_share,
        dollars_lost_fy1=ref[0].contributions - f[0].contributions,
        final_year=last.year,
        final_operating_result=last.operating_result,
        reference_operating_result=ref_last.operating_result,
        cumulative_gap=sum(r.operating_result for r in ref) - sum(r.operating_result for r in f),
        closing_liquid=last.closing_liquid,
        liquid_runway_months=last.derived()["liquid_runway_months"],
        outcomes_lost=ref_last.outcomes - last.outcomes,
        participants_lost=ref_last.participants - last.participants,
        capacity_retained=last.participants / ref_last.participants,
        first_contraction_year=next((x.year for x in f if x.contraction_applied > 1e-9), None),
        years_to_floor=years_to_reserve_floor(scenario, contrib_share),
        reference_years_to_floor=years_to_reserve_floor(scenario, 0.0),
    )


def structural_breakeven(scenario: str = "Base", years: list[int] | None = None) -> dict:
    """
    The funding level at which the final forecast year breaks even before any
    reserve is drawn. This is the sustainability question, and it is different
    from the runway question: reserves buy time to close a structural gap, they
    do not close it.
    """
    years = years or SENSITIVITY_HORIZON

    def gap(shock: float) -> float:
        f = forecast(scenario, funding_shock=shock, shock_year=years[0], years=years)
        return f[-1].operating_result

    # The shock is allowed to go negative, which represents a funding uplift.
    # It has to: in the Base case this organization is already in structural
    # deficit by the end of the horizon, so the meaningful answer is not 'how
    # much can funding fall' but 'how much must it rise'.
    lo, hi = -0.80, 0.95
    negative_at_zero = gap(0.0) < 0
    for _ in range(90):
        mid = (lo + hi) / 2
        if gap(mid) >= 0:
            lo = mid
        else:
            hi = mid
    at = forecast(scenario, funding_shock=lo, shock_year=years[0], years=years)
    ref = forecast(scenario, years=years)
    return dict(
        already_structurally_negative=negative_at_zero,
        max_shock_at_breakeven=lo,
        required_funding_uplift=max(0.0, -lo),
        scenario=scenario,
        final_year=years[-1],
        contributions_at_breakeven=at[0].contributions,
        reference_contributions=ref[0].contributions,
        dollars_of_structural_headroom=ref[0].contributions - at[0].contributions,
        crossover=structural_deficit_crossover(scenario),
    )


def minimum_viable_funding(scenario: str = "Base", years: list[int] | None = None) -> dict:
    """
    Two distinct thresholds, reported together because quoting either one alone
    is misleading:

      reserve threshold   - the funding cut at which liquid reserves reach the
                            six-month floor within the horizon, forcing the
                            organization to cut programs
      structural threshold - the funding cut at which the organization stops
                            breaking even in the final year, regardless of how
                            much reserve it holds

    The gap between them is the runway the balance sheet buys management to fix
    a structural problem. It is the single most decision-relevant number here.
    """
    years = years or SENSITIVITY_HORIZON
    floor = I.ALLOCATION_CONSTRAINTS["reserve_floor_months"]

    def holds(shock: float) -> bool:
        f = forecast(scenario, funding_shock=shock, shock_year=years[0], years=years)
        return all(r.closing_liquid / (r.total_expense / 12) >= floor for r in f)

    lo, hi = 0.0, 0.95
    if holds(hi):
        max_shock = hi
    else:
        for _ in range(80):
            mid = (lo + hi) / 2
            if holds(mid):
                lo = mid
            else:
                hi = mid
        max_shock = lo

    ref = forecast(scenario, years=years)
    at_limit = forecast(scenario, funding_shock=max_shock, shock_year=years[0], years=years)
    struct = structural_breakeven(scenario, years)

    return dict(
        horizon_years=len(years),
        final_year=years[-1],
        reserve_floor_months=floor,
        max_sustainable_shock=max_shock,
        min_contributions_fy1=at_limit[0].contributions,
        reference_contributions_fy1=ref[0].contributions,
        dollars_of_headroom=ref[0].contributions - at_limit[0].contributions,
        final_runway=at_limit[-1].closing_liquid / (at_limit[-1].total_expense / 12),
        structural_shock=struct["max_shock_at_breakeven"],
        structural_contributions=struct.get("contributions_at_breakeven"),
        buffer_between_thresholds=max_shock - struct["max_shock_at_breakeven"],
        required_funding_uplift=struct["required_funding_uplift"],
        crossover=struct["crossover"],
        years_to_floor_no_shock=years_to_reserve_floor(scenario, 0.0),
    )


def expense_reduction_required(scenario: str = "Downside", years: list[int] | None = None) -> dict:
    """Complement to the funding question: how much cost must come out to hold the floor."""
    years = years or SENSITIVITY_HORIZON
    floor = I.ALLOCATION_CONSTRAINTS["reserve_floor_months"]
    f = forecast(scenario, years=years)
    worst = min(r.closing_liquid / (r.total_expense / 12) for r in f)
    worst_year = min(f, key=lambda r: r.closing_liquid / (r.total_expense / 12)).year
    if worst >= floor:
        return dict(required=False, worst_runway=worst, worst_year=worst_year,
                    reduction_pct=0.0, reduction_dollars=0.0,
                    final_structural_result=f[-1].operating_result)

    lo, hi = 0.0, 0.75
    for _ in range(80):
        mid = (lo + hi) / 2
        trial = forecast(scenario, years=years, expense_reduction=mid)
        if all(r.closing_liquid / (r.total_expense / 12) >= floor for r in trial):
            hi = mid
        else:
            lo = mid
    pct = hi
    trial = forecast(scenario, years=years, expense_reduction=pct)
    return dict(required=True, worst_runway=worst, worst_year=worst_year, reduction_pct=pct,
                reduction_dollars=f[0].total_expense * pct,
                participants_forgone=f[-1].participants - trial[-1].participants,
                outcomes_forgone=f[-1].outcomes - trial[-1].outcomes,
                final_structural_result=f[-1].operating_result)


# ==========================================================================
# Budget vs actual, FY2024
# ==========================================================================

def budget_vs_actual() -> dict:
    """
    FY2024 actuals are real. The budget is a reconstruction of what would
    plausibly have been approved in late FY2023, clearly labelled as such.
    Variance is then decomposed into volume and rate effects.
    """
    hist = historical()
    a24 = hist["rows"][2024]
    a23 = hist["rows"][2023]
    pe = program_economics()

    b_rev = a23["revenue"] * (1 + I.BUDGET_2024_REVENUE_GROWTH)
    b_exp = a23["expenses"] * (1 + I.BUDGET_2024_EXPENSE_GROWTH)
    b_parts = a23["participants"] * (1 + I.BUDGET_2024_PARTICIPANT_GROWTH)

    b_contrib = b_rev * (a23["contributions"] / a23["revenue"])
    b_progrev = b_rev * (a23["program_revenue"] / a23["revenue"])
    b_other = b_rev - b_contrib - b_progrev

    p23, m23, f23 = I.FUNCTIONAL_SPLIT[2023]
    b_prog = b_exp * p23
    b_mg = b_exp * m23
    b_fr = b_exp * f23
    b_personnel = b_exp * a23["personnel_ratio"]
    b_cpp = b_prog / b_parts

    a_parts = a24["participants"]
    a_cpp = a24["program_expense"] / a_parts

    def var(actual, budget):
        return dict(actual=actual, budget=budget, variance=actual - budget,
                    variance_pct=(actual - budget) / budget if budget else None,
                    favourable=None)

    lines = {
        "Total revenue": var(a24["revenue"], b_rev),
        "  Contributions and grants": var(a24["contributions"], b_contrib),
        "  Program service revenue": var(a24["program_revenue"], b_progrev),
        "  Investment and other income": var(a24["investment_other"], b_other),
        "Total expenses": var(a24["expenses"], b_exp),
        "  Program services": var(a24["program_expense"], b_prog),
        "  Management and general": var(a24["mg_expense"], b_mg),
        "  Fundraising": var(a24["fundraising_expense"], b_fr),
        "Personnel (memo)": var(a24["personnel"], b_personnel),
        "Non-personnel (memo)": var(a24["non_personnel"], b_exp - b_personnel),
        "Operating result": var(a24["operating_result"], b_rev - b_exp),
        "Participants served": var(a_parts, b_parts),
        "Successful outcomes": var(a24["placements"],
                                   b_parts * (a23["placements"] / a23["participants"])),
        "Cost per participant": var(a_cpp, b_cpp),
    }
    for k, v in lines.items():
        expense_like = ("expense" in k.lower() or "Personnel" in k or "personnel" in k
                        or "Cost per" in k or k.strip().startswith(("Program services",
                                                                    "Management", "Fundraising")))
        v["favourable"] = (v["variance"] < 0) if expense_like else (v["variance"] > 0)

    # --- program cost variance decomposition -----------------------------
    prog_var = a24["program_expense"] - b_prog
    volume_var = (a_parts - b_parts) * b_cpp
    rate_var = (a_cpp - b_cpp) * a_parts
    decomposition = dict(
        total_program_variance=prog_var,
        volume_variance=volume_var,
        rate_variance=rate_var,
        check=abs(prog_var - (volume_var + rate_var)) < 0.01,
        budget_participants=b_parts,
        actual_participants=a_parts,
        budget_cpp=b_cpp,
        actual_cpp=a_cpp,
    )

    funding_var = dict(
        total=a24["contributions"] - b_contrib,
        by_category={c: (a24["contributions"] * I.REVENUE_MIX[c]) - (b_contrib * I.REVENUE_MIX[c])
                     for c in I.REVENUE_CATEGORIES},
    )
    personnel_var = dict(
        total=a24["personnel"] - b_personnel,
        rate_effect=a24["personnel"] - b_personnel,  # single-rate simplification, stated in workbook
    )

    return dict(lines=lines, decomposition=decomposition, funding_variance=funding_var,
                personnel_variance=personnel_var, budget=dict(
                    revenue=b_rev, expense=b_exp, participants=b_parts, cpp=b_cpp,
                    program=b_prog, mg=b_mg, fundraising=b_fr, personnel=b_personnel,
                    contributions=b_contrib, program_revenue=b_progrev, other=b_other),
                actual=a24)


# ==========================================================================
# Resource allocation
# ==========================================================================

def allocate(budget: float | None = None, constraints: dict | None = None,
             objective: str = "outcomes") -> dict:
    """
    Allocate incremental funding across programs, subject to the stated
    constraints, under one of three objectives.

    The three objectives matter because they disagree. Maximizing placement
    count sends money to the cheapest channel; maximizing earnings gain sends it
    to the most expensive one. Management has to choose which it is buying, and
    the model's job is to price that choice rather than hide it behind a single
    'optimal' answer.

    Solved with PuLP as a linear program. The Excel version in
    06_Resource_Allocation.xlsx is set up for Solver and reproduces these
    answers; the test suite checks that it does.
    """
    import pulp

    budget = budget or I.ALLOCATION_BUDGET
    c = dict(I.ALLOCATION_CONSTRAINTS)
    if constraints:
        c.update(constraints)

    pe = program_economics()
    progs = pe["programs"]
    deployable = budget * (1 - c["admin_load"])

    prob = pulp.LpProblem(f"allocation_{objective}", pulp.LpMaximize)
    x = {p: pulp.LpVariable(f"alloc_{I.PROGRAM_CODES[p]}", lowBound=0) for p in I.PROGRAMS}

    outcomes_expr = pulp.lpSum(x[p] / progs[p]["cost_per_outcome"] for p in I.PROGRAMS)
    earnings_expr = pulp.lpSum(x[p] / progs[p]["cost_per_outcome"] * progs[p]["salary_gain"]
                               for p in I.PROGRAMS)

    if objective == "outcomes":
        prob += outcomes_expr
    elif objective == "earnings":
        prob += earnings_expr
    elif objective == "balanced":
        prob += outcomes_expr
        # average earnings gain per placement funded must clear the quality floor,
        # linearised as: sum(outcomes_p * gain_p) >= floor * sum(outcomes_p)
        prob += pulp.lpSum(x[p] / progs[p]["cost_per_outcome"]
                           * (progs[p]["salary_gain"] - c["min_avg_salary_gain"])
                           for p in I.PROGRAMS) >= 0, "quality_floor"
    else:
        raise ValueError(f"unknown objective {objective!r}")

    prob += pulp.lpSum(x.values()) <= deployable, "budget"
    for p in I.PROGRAMS:
        prob += x[p] >= c["min_share_per_program"] * deployable, f"min_{I.PROGRAM_CODES[p]}"
        prob += x[p] <= c["max_share_per_program"] * deployable, f"max_{I.PROGRAM_CODES[p]}"
        headroom = progs[p]["max_participants"] - progs[p]["participants"]
        prob += x[p] <= headroom * progs[p]["cost_per_participant"], f"cap_{I.PROGRAM_CODES[p]}"
    prob += pulp.lpSum(x[p] / progs[p]["cost_per_participant"]
                       * progs[p]["staff_hours_per_participant"]
                       for p in I.PROGRAMS) <= c["staff_hours_available"], "staff_hours"

    prob.solve(pulp.PULP_CBC_CMD(msg=0))

    alloc = {p: (x[p].value() or 0.0) for p in I.PROGRAMS}
    detail = {}
    for p in I.PROGRAMS:
        a = alloc[p]
        add_parts = a / progs[p]["cost_per_participant"]
        add_out = a / progs[p]["cost_per_outcome"]
        detail[p] = dict(
            allocation=a,
            share=a / deployable if deployable else 0,
            additional_participants=add_parts,
            additional_outcomes=add_out,
            additional_earnings=add_out * progs[p]["salary_gain"],
            additional_hours=add_parts * progs[p]["staff_hours_per_participant"],
            marginal_cost_per_outcome=progs[p]["cost_per_outcome"],
            salary_gain=progs[p]["salary_gain"],
            new_participants=progs[p]["participants"] + add_parts,
            capacity_after=(progs[p]["participants"] + add_parts) / progs[p]["max_participants"],
            binding_capacity=abs(progs[p]["participants"] + add_parts
                                 - progs[p]["max_participants"]) < 1.0,
            binding_max_share=abs(a - c["max_share_per_program"] * deployable) < 1.0,
            binding_min_share=abs(a - c["min_share_per_program"] * deployable) < 1.0,
        )

    total_out = sum(d["additional_outcomes"] for d in detail.values())
    total_earn = sum(d["additional_earnings"] for d in detail.values())
    total_hours = sum(d["additional_hours"] for d in detail.values())

    def policy(weights: dict) -> dict:
        tot = sum(weights.values())
        alloc_p = {p: deployable * w / tot for p, w in weights.items()}
        outs = sum(alloc_p[p] / progs[p]["cost_per_outcome"] for p in I.PROGRAMS)
        earn = sum(alloc_p[p] / progs[p]["cost_per_outcome"] * progs[p]["salary_gain"]
                   for p in I.PROGRAMS)
        hrs = sum(alloc_p[p] / progs[p]["cost_per_participant"]
                  * progs[p]["staff_hours_per_participant"] for p in I.PROGRAMS)
        return dict(allocation=alloc_p, outcomes=outs, earnings=earn, hours=hrs,
                    avg_gain=earn / outs if outs else 0,
                    feasible_hours=hrs <= c["staff_hours_available"])

    equal = policy({p: 1 for p in I.PROGRAMS})
    pro_rata = policy({p: progs[p]["program_cost"] for p in I.PROGRAMS})

    return dict(
        objective=objective,
        objective_label=I.ALLOCATION_OBJECTIVES[objective],
        status=pulp.LpStatus[prob.status],
        budget=budget,
        admin_load=budget * c["admin_load"],
        deployable=deployable,
        detail=detail,
        total_additional_outcomes=total_out,
        total_additional_earnings=total_earn,
        avg_earnings_gain_per_outcome=total_earn / total_out if total_out else None,
        total_additional_hours=total_hours,
        hours_constraint=c["staff_hours_available"],
        hours_binding=abs(total_hours - c["staff_hours_available"]) < 1.0,
        hours_utilization=total_hours / c["staff_hours_available"],
        blended_cost_per_outcome=deployable / total_out if total_out else None,
        comparison=dict(
            optimized=total_out,
            equal_split=equal["outcomes"],
            pro_rata=pro_rata["outcomes"],
            gain_vs_equal=total_out - equal["outcomes"],
            gain_vs_pro_rata=total_out - pro_rata["outcomes"],
            equal_detail=equal,
            pro_rata_detail=pro_rata,
        ),
        constraints=c,
    )


def allocation_objective_comparison() -> dict:
    """Solve all three objectives and price the difference between them."""
    runs = {o: allocate(objective=o) for o in I.ALLOCATION_OBJECTIVES}
    best_out = runs["outcomes"]["total_additional_outcomes"]
    best_earn = runs["earnings"]["total_additional_earnings"]
    rows = []
    for o, r in runs.items():
        rows.append(dict(
            objective=o,
            label=I.ALLOCATION_OBJECTIVES[o],
            outcomes=r["total_additional_outcomes"],
            earnings=r["total_additional_earnings"],
            avg_gain=r["avg_earnings_gain_per_outcome"],
            cost_per_outcome=r["blended_cost_per_outcome"],
            outcomes_forgone=best_out - r["total_additional_outcomes"],
            earnings_forgone=best_earn - r["total_additional_earnings"],
            allocation={p: d["allocation"] for p, d in r["detail"].items()},
        ))
    return dict(runs=runs, rows=rows,
                outcomes_price_of_earnings_focus=best_out - runs["earnings"]["total_additional_outcomes"],
                earnings_price_of_outcome_focus=best_earn - runs["outcomes"]["total_additional_earnings"])


def allocation_efficient_frontier(steps: int = 11) -> list[dict]:
    """Outcomes achievable at increasing budget levels - shows where capacity binds."""
    out = []
    for i in range(steps):
        b = I.ALLOCATION_BUDGET * (0.2 + 0.16 * i)
        r = allocate(budget=b)
        out.append(dict(budget=b, outcomes=r["total_additional_outcomes"],
                        earnings=r["total_additional_earnings"],
                        cost_per_outcome=r["blended_cost_per_outcome"],
                        hours_utilization=r["hours_utilization"],
                        hours_binding=r["hours_binding"]))
    return out


def allocation_tradeoff_frontier(steps: int = 9) -> list[dict]:
    """
    Sweep the earnings-gain quality floor and record what each level costs in
    placement count. This is the curve management should actually look at.
    """
    out = []
    for i in range(steps):
        floor = 37_000 + i * 5_000
        try:
            r = allocate(objective="balanced", constraints=dict(min_avg_salary_gain=floor))
            if r["status"] != "Optimal":
                out.append(dict(floor=floor, feasible=False))
                continue
            out.append(dict(floor=floor, feasible=True,
                            outcomes=r["total_additional_outcomes"],
                            earnings=r["total_additional_earnings"],
                            avg_gain=r["avg_earnings_gain_per_outcome"],
                            cost_per_outcome=r["blended_cost_per_outcome"]))
        except Exception:
            out.append(dict(floor=floor, feasible=False))
    return out


# ==========================================================================
# Convenience
# ==========================================================================

def full_model() -> dict:
    return dict(
        historical=historical(),
        programs=program_economics(),
        scenarios=all_scenarios(),
        sensitivity=funding_sensitivity(),
        largest_funder=largest_funder_loss(),
        minimum_funding=minimum_viable_funding(),
        structural_breakeven=structural_breakeven(),
        expense_reduction=expense_reduction_required(),
        bva=budget_vs_actual(),
        allocation=allocate(),
        allocation_objectives=allocation_objective_comparison(),
        allocation_frontier=allocation_efficient_frontier(),
        allocation_tradeoff=allocation_tradeoff_frontier(),
        sroi=sroi(),
        sroi_range=sroi_sensitivity(),
    )


if __name__ == "__main__":
    import json
    m = full_model()
    h = m["historical"]
    print("=== Historical ===")
    for y in I.HIST_YEARS:
        r = h["rows"][y]
        print(f"FY{y}  rev {r['revenue']:>12,.0f}  exp {r['expenses']:>12,.0f}  "
              f"result {r['operating_result']:>11,.0f}  prog% {r['program_ratio']:.1%}  "
              f"cash {r['cash_runway_months']:.1f}mo  liq {r['liquid_runway_months']:.1f}mo  "
              f"NA x{r['net_asset_multiple']:.2f}  parts {r['participants']:>8,.0f}")
    print("\nCAGR rev %.1f%%  exp %.1f%%  cumulative surplus %s" % (
        h["summary"]["revenue_cagr"] * 100, h["summary"]["expense_cagr"] * 100,
        f"{h['summary']['cumulative_surplus']:,.0f}"))

    print("\n=== Program economics FY2024 ===")
    pe = m["programs"]
    for p, d in pe["programs"].items():
        print(f"{p:<20} parts {d['participants']:>7,.0f}  out {d['outcomes']:>6,.0f}  "
              f"CPP {d['cost_per_participant']:>9,.0f}  CPO {d['cost_per_outcome']:>10,.0f}  "
              f"rate {d['outcome_rate']:.1%}")
    t = pe["totals"]
    print(f"TOTALS parts {t['participants']:,.0f}  outcomes {t['outcomes']:,.0f}  "
          f"CPO {t['cost_per_outcome']:,.0f}  blended gain {t['blended_salary_gain']:,.0f} "
          f"vs published {t['published_salary_gain']:,.0f} ({t['calibration_variance']:+.2%})  "
          f"FTE {t['implied_total_fte']:.1f}")

    print("\n=== Scenarios ===")
    for s, rows in m["scenarios"].items():
        for r in rows:
            d = r.derived()
            print(f"{s:<9} FY{r.year}  rev {r.total_revenue:>12,.0f}  exp {r.total_expense:>12,.0f}  "
                  f"result {r.operating_result:>11,.0f}  liquid {r.closing_liquid:>12,.0f}  "
                  f"runway {d['liquid_runway_months']:>5.1f}mo  prog% {d['program_ratio']:.1%}  "
                  f"out {r.outcomes:>6,.0f}")

    print("\n=== Funding sensitivity (Base, 5-year, reserve floor enforced) ===")
    for s in m["sensitivity"]:
        print(f"shock {s['shock']:>5.0%}  FY{2029} result {s['fy_last_result']:>12,.0f}  "
              f"structural {s['structural_result']:>12,.0f}  "
              f"liquid {s['closing_liquid']:>12,.0f}  runway {s['liquid_runway_months']:>5.1f}mo  "
              f"outcomes {s['outcomes']:>6,.0f} ({s['capacity_retained']:.0%} capacity)  "
              f"cut from FY{s['first_contraction_year'] or '-'}  "
              f"floor in {s['years_to_floor']['years'] or '>15'}y")

    print("\n=== Minimum viable funding ===")
    print(json.dumps({k: (round(v, 4) if isinstance(v, float) else v)
                      for k, v in m["minimum_funding"].items()}, indent=2))

    print("\n=== Largest funder loss ===")
    print(json.dumps({k: (round(v, 2) if isinstance(v, float) else v)
                      for k, v in m["largest_funder"].items()}, indent=2))

    print("\n=== Expense reduction required (Downside) ===")
    print(json.dumps({k: (round(v, 4) if isinstance(v, float) else v)
                      for k, v in m["expense_reduction"].items()}, indent=2))

    print("\n=== Allocation, three objectives ===")
    for row in m["allocation_objectives"]["rows"]:
        print(f"{row['objective']:<10} outcomes {row['outcomes']:>7.1f}  "
              f"earnings {row['earnings']:>13,.0f}  avg gain {row['avg_gain']:>8,.0f}  "
              f"CPO {row['cost_per_outcome']:>8,.0f}")
        for p, v in row["allocation"].items():
            print(f"       {p:<20} {v:>10,.0f}")
    ao = m["allocation_objectives"]
    print(f"  price of an earnings focus: {ao['outcomes_price_of_earnings_focus']:.1f} placements")
    print(f"  price of an outcome focus:  ${ao['earnings_price_of_outcome_focus']:,.0f} earnings gain")

    print("\n=== Quality-floor trade-off ===")
    for t in m["allocation_tradeoff"]:
        if t.get("feasible"):
            print(f"  floor {t['floor']:>7,}  outcomes {t['outcomes']:>7.1f}  "
                  f"avg gain {t['avg_gain']:>8,.0f}  earnings {t['earnings']:>13,.0f}")
        else:
            print(f"  floor {t['floor']:>7,}  INFEASIBLE")

    print("\n=== SROI (reported last, low confidence) ===")
    print(f"portfolio ratio {m['sroi']['_portfolio']['sroi_ratio']:.2f}  "
          f"range {m['sroi_range']['low']:.2f} to {m['sroi_range']['high']:.2f} "
          f"(spread {m['sroi_range']['spread']:.1f}x)")
