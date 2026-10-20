from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Iterator, TypeVar

from openpyxl.utils import get_column_letter


CUSTOMER_LOOKUP_SHEET_NAME = "_CustomerLookup"
CUSTOMER_LOOKUP_ERROR = "error"
CUSTOMER_LOOKUP_MIN_MODEL_SLOTS = 16

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
    header_row: int = 1
    model_column: int = 1
    customer_column: int = 2

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
        return cls(
            rows, customers, frozenset(conflicts),
            header_row + 1, model_column + 1, customer_column + 1,
        )

    def last_row_formula(self) -> str:
        column = get_column_letter(self.model_column)
        source = f"'Customer'!${column}:${column}"
        first_row = self.header_row + 1
        return (
            f'=MAX({first_row},IFERROR('
            f'LOOKUP(2,1/({source}<>""),ROW({source})),{first_row}))'
        )

    def formula_for(self, model_ref: str, last_row: int) -> str:
        """Use a scalar VLOOKUP supported by Excel 2016 and later."""
        key = _lookup_key(f'TRIM({model_ref}&"")')
        table = f"'{CUSTOMER_LOOKUP_SHEET_NAME}'!$A$3:$B${max(3, last_row)}"
        return f'=IFERROR(VLOOKUP({key},{table},2,FALSE),"{CUSTOMER_LOOKUP_ERROR}")'

    def helper_rows(self, model_column: str, models: list[Any]) -> Iterator[list[Any]]:
        """Split and join models using ordinary scalar helper cells, without CSE."""
        maximum = max((str(model or "").count("/") + 1 for model in models), default=1)
        slots = max(CUSTOMER_LOOKUP_MIN_MODEL_SLOTS, maximum + 4)
        if 5 + slots * 4 > 16384:
            raise ValueError("MODEL 機種數超過 Excel 輔助欄位容量")

        def source_range(column: int) -> str:
            letter = get_column_letter(column)
            return (
                f"'Customer'!${letter}${self.header_row + 1}:"
                f"INDEX('Customer'!${letter}:${letter},"
                f"'{CUSTOMER_LOOKUP_SHEET_NAME}'!$A$1)"
            )

        keys = f'IFERROR(TRIM({source_range(self.model_column)}&""),"")'
        customers = f'IFERROR(TRIM({source_range(self.customer_column)}&""),"")'
        error = f'"{CUSTOMER_LOOKUP_ERROR}"'

        def candidate(key: str) -> str:
            return (
                f'IFERROR(LOOKUP(2,1/(({keys}={key})*({customers}<>"")),'
                f'{customers}),"")'
            )

        def resolved(key: str, value: str) -> str:
            conflicts = (
                f'SUMPRODUCT(--({keys}={key}),--({customers}<>""),'
                f'--({customers}<>{value}))'
            )
            return f'IF({value}="",{error},IF({conflicts}=0,{value},{error}))'

        yield [self.last_row_formula(), slots]
        yield [
            "MODEL", "Customer", "Exact candidate", "Exact customer", "Model count",
            *(f"Model {index}" for index in range(1, slots + 1)),
            *(f"Candidate {index}" for index in range(1, slots + 1)),
            *(f"Customer {index}" for index in range(1, slots + 1)),
            *(f"Combined {index}" for index in range(1, slots + 1)),
        ]
        for row_number in range(3, len(models) + 3):
            model = f"$A{row_number}"
            exact_candidate = f"$C{row_number}"
            exact_customer = f"$D{row_number}"
            model_count = f"$E{row_number}"
            last_combined = f"{get_column_letter(5 + slots * 4)}{row_number}"
            row = [
                f'=TRIM(\'Excess\'!${model_column}{row_number}&"")',
                f'=IF({model}="",{error},IF({exact_candidate}<>"",{exact_customer},'
                f'IF(OR({model_count}>{slots},{last_combined}=""),{error},{last_combined})))',
                f'=IF({model}="","",{candidate(model)})',
                f'={resolved(model, exact_candidate)}',
                f'=IF({model}="",0,LEN({model})-LEN(SUBSTITUTE({model},"/",""))+1)',
            ]
            part_keys, candidates, results, combined = [], [], [], []
            for index in range(slots):
                key = f"{get_column_letter(6 + index)}{row_number}"
                value = f"{get_column_letter(6 + slots + index)}{row_number}"
                result = f"{get_column_letter(6 + slots * 2 + index)}{row_number}"
                previous = (
                    f"{get_column_letter(5 + slots * 3 + index)}{row_number}"
                    if index else '""'
                )
                duplicates = (
                    f'SUMPRODUCT(--(${get_column_letter(6 + slots * 2)}{row_number}:'
                    f'{get_column_letter(5 + slots * 2 + index)}{row_number}={result}))>0'
                    if index else "FALSE"
                )
                part_keys.append(
                    f'=IF(AND({exact_candidate}="",{index + 1}<={model_count},'
                    f'{model_count}<={slots}),TRIM(MID(SUBSTITUTE({model},"/",'
                    f'REPT(" ",LEN({model}))),{index}*LEN({model})+1,LEN({model}))),"")'
                )
                candidates.append(f'=IF({key}="","",{candidate(key)})')
                results.append(f'=IF({key}="","",{resolved(key, value)})')
                combined.append(
                    f'=IF({key}="",{previous},IF(OR({previous}={error},{result}={error}),'
                    f'{error},IF({duplicates},{previous},'
                    f'IF({previous}="",{result},{previous}&"/"&{result}))))'
                )
            yield row + part_keys + candidates + results + combined

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


def _lookup_key(expression: str) -> str:
    # Escape VLOOKUP wildcard characters while retaining an exact text match.
    return (
        f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({expression},'
        f'"~","~~"),"*","~*"),"?","~?")'
    )
