"""
Provenance tagging for every input in the model.

The single most important discipline in this project is that no number ever
appears in a workbook without a label saying where it came from. Four tiers:

    PUBLIC     - taken verbatim from a public filing or publication
    DERIVED    - arithmetic on PUBLIC values only (interpolation, subtotals,
                 ratios). No judgement applied beyond the stated method.
    BENCHMARK  - a published sector standard or peer statistic
    SYNTHETIC  - an analyst assumption. Not the organization's data. Not its
                 guidance. Constructed to be plausible and internally
                 consistent, and disclosed as an assumption.

Anything tagged SYNTHETIC must still reconcile to a PUBLIC control total
wherever a control total exists. That is what keeps the model honest: the
composition is assumed, the totals are real.
"""

from dataclasses import dataclass, field
from typing import Any

PUBLIC = "PUBLIC"
DERIVED = "DERIVED"
BENCHMARK = "BENCHMARK"
SYNTHETIC = "SYNTHETIC"

TIER_ORDER = [PUBLIC, DERIVED, BENCHMARK, SYNTHETIC]

TIER_DESCRIPTION = {
    PUBLIC: "Verbatim from a public filing or publication",
    DERIVED: "Arithmetic on public values only, by the stated method",
    BENCHMARK: "Published sector standard or peer statistic",
    SYNTHETIC: "Analyst assumption - not organizational data or guidance",
}

TIER_COLOR = {
    PUBLIC: "1B5E20",       # green
    DERIVED: "0D47A1",      # blue
    BENCHMARK: "4A148C",    # purple
    SYNTHETIC: "B71C1C",    # red
}


@dataclass
class Input:
    """A single model input carrying its own provenance."""

    key: str
    label: str
    value: Any
    unit: str
    tier: str
    source: str
    note: str = ""
    section: str = ""
    fmt: str = "#,##0"

    def __post_init__(self):
        if self.tier not in TIER_ORDER:
            raise ValueError(f"{self.key}: unknown provenance tier {self.tier!r}")
        if not self.source.strip():
            raise ValueError(f"{self.key}: every input needs a source")


class Register:
    """Ordered collection of Inputs, addressable by key."""

    def __init__(self):
        self._items: list[Input] = []
        self._index: dict[str, Input] = {}

    def add(self, inp: Input) -> Input:
        if inp.key in self._index:
            raise ValueError(f"duplicate input key: {inp.key}")
        self._items.append(inp)
        self._index[inp.key] = inp
        return inp

    def __getitem__(self, key: str) -> Any:
        return self._index[key].value

    def meta(self, key: str) -> Input:
        return self._index[key]

    def __iter__(self):
        return iter(self._items)

    def __len__(self):
        return len(self._items)

    def by_section(self) -> dict[str, list[Input]]:
        out: dict[str, list[Input]] = {}
        for i in self._items:
            out.setdefault(i.section, []).append(i)
        return out

    def tier_counts(self) -> dict[str, int]:
        counts = {t: 0 for t in TIER_ORDER}
        for i in self._items:
            counts[i.tier] += 1
        return counts
