from __future__ import annotations

from dataclasses import dataclass
from typing import Any, TypeVar


_T = TypeVar("_T")


def previous_value_for(
    parts: list[str],
    by_part: dict[str, _T],
    by_group: dict[tuple[str, ...], _T],
) -> _T | None:
    """Use the same identity for carried inputs and the Improve baseline."""
    keys = [str(part).strip() for part in parts if part is not None and str(part).strip()]
    if keys and keys[0] in by_part:
        return by_part[keys[0]]
    return by_group.get(tuple(sorted(set(keys)))) if keys else None


@dataclass(frozen=True)
class PreviousExcessInputs:
    values_by_part: dict[str, dict[str, Any]]
    values_by_group: dict[tuple[str, ...], dict[str, Any]]
    customer_sheet_rows: tuple[tuple[Any, ...], ...] | None = None

    def values_for(self, parts: list[str]) -> dict[str, Any]:
        return previous_value_for(parts, self.values_by_part, self.values_by_group) or {}
