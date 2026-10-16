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
class CustomerLookup:
    sheet_rows: tuple[tuple[Any, ...], ...]
    customers_by_model: dict[str, str]
    conflicting_models: frozenset[str]

    @classmethod
    def from_rows(cls, rows: tuple[tuple[Any, ...], ...]) -> CustomerLookup:
        header_row = None
        for index, row in enumerate(rows[:20]):
            headers = [str(value).strip().casefold() for value in row]
            if "model" in headers and "customer" in headers:
                header_row = index
                model_column = headers.index("model")
                customer_column = headers.index("customer")
                break
        if header_row is None:
            raise ValueError("Customer sheet must contain MODEL and Customer headers")

        customers: dict[str, str] = {}
        conflicts: set[str] = set()
        for row in rows[header_row + 1:]:
            model = _text(row[model_column]) if model_column < len(row) else ""
            customer = _text(row[customer_column]) if customer_column < len(row) else ""
            if not model or not customer:
                continue
            if model in customers and customers[model] != customer:
                conflicts.add(model)
            customers.setdefault(model, customer)
        return cls(rows, customers, frozenset(conflicts))

    def resolve(self, model: Any) -> tuple[str | None, str | None]:
        key = _text(model)
        if not key:
            return None, "MODEL 空白"
        if key in self.conflicting_models:
            return None, "Customer 對應衝突"
        if key in self.customers_by_model:
            return self.customers_by_model[key], None

        models = list(dict.fromkeys(part.strip() for part in key.split("/") if part.strip()))
        if not models:
            return None, "查無 MODEL"
        if any(part in self.conflicting_models for part in models):
            return None, "Customer 對應衝突"
        matched = [self.customers_by_model[part] for part in models if part in self.customers_by_model]
        if len(matched) != len(models):
            return None, "部分機種查不到" if matched else "查無 MODEL"
        return "/".join(dict.fromkeys(matched)), None


@dataclass(frozen=True)
class PreviousExcessInputs:
    values_by_part: dict[str, dict[str, Any]]
    values_by_group: dict[tuple[str, ...], dict[str, Any]]
    customer_lookup: CustomerLookup | None = None

    def values_for(self, parts: list[str]) -> dict[str, Any]:
        return previous_value_for(parts, self.values_by_part, self.values_by_group) or {}


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()
