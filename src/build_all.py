"""
Rebuild every deliverable from scratch.

    python src/build_all.py

Order matters only in that the reports read the engine, and the engine reads
inputs.py. Nothing reads a workbook, so the builders are independent of each
other and could run in any sequence.
"""

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import build_01_public_data
import build_02_assumptions
import build_03_financial_model
import build_04_budget_vs_actual
import build_05_scenario_model
import build_06_program_economics
import build_07_resource_allocation
import build_08_powerbi
import build_09_10_reports

STEPS = [
    ("01  Public data", build_01_public_data.build),
    ("02  Assumptions register", build_02_assumptions.build),
    ("03  Financial model", build_03_financial_model.build),
    ("04  Budget vs actual", build_04_budget_vs_actual.build),
    ("05  Scenario model", build_05_scenario_model.build),
    ("06  Program economics", build_06_program_economics.build),
    ("07  Resource allocation", build_07_resource_allocation.build),
    ("08  Power BI package", build_08_powerbi.build),
    ("09/10  Reports", build_09_10_reports.build),
]


def main():
    t0 = time.time()
    for name, fn in STEPS:
        print(f"\n=== {name} " + "=" * (56 - len(name)))
        fn()
    print(f"\nAll deliverables rebuilt in {time.time() - t0:.1f}s")
    print("Run 'python -m pytest tests -q' to verify the workbooks recalculate correctly.")


if __name__ == "__main__":
    main()
