from __future__ import annotations

from dataclasses import dataclass
from typing import Any, TypeVar

from openpyxl.utils import get_column_letter


CUSTOMER_LOOKUP_SHEET_NAME = "_CustomerLookup"
CUSTOMER_LOOKUP_ERROR = "error"

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

    def formula_for(self, model_ref: str) -> str:
        """Excel 2021 lookup, including complete multi-model matches and conflicts."""
        def source_range(column: int) -> str:
            letter = get_column_letter(column)
            return (
                f"'Customer'!${letter}${self.header_row + 1}:"
                f"INDEX('Customer'!${letter}:${letter},"
                f"'{CUSTOMER_LOOKUP_SHEET_NAME}'!$A$1)"
            )

        def lookup_key(expression: str) -> str:
            # VLOOKUP treats these characters as wildcards even with FALSE.
            return (
                f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({expression},'
                f'"~","~~"),"*","~*"),"?","~?")'
            )

        return (
            '=IFERROR(_xlfn.LET('
            f'modelText,TRIM({model_ref}&""),'
            f'sourceModels,IFERROR(TRIM({source_range(self.model_column)}&""),""),'
            f'sourceCustomers,IFERROR(TRIM({source_range(self.customer_column)}&""),""),'
            'validRows,(sourceModels<>"")*(sourceCustomers<>""),'
            'modelKeys,_xlfn._xlws.FILTER(sourceModels,validRows,""),'
            'customerValues,_xlfn._xlws.FILTER(sourceCustomers,validRows,""),'
            'lookupTable,CHOOSE({1,2},modelKeys,customerValues),'
            f'exactCustomer,IFERROR(VLOOKUP({lookup_key("modelText")},lookupTable,2,FALSE),""),'
            'IF(modelText="",NA(),IF(exactCustomer<>"",'
            'IF(SUM((modelKeys=modelText)*(customerValues<>exactCustomer))=0,exactCustomer,NA()),'
            '_xlfn.LET('
            'rawParts,TRIM(MID(SUBSTITUTE(modelText,"/",REPT(" ",LEN(modelText))),'
            '_xlfn.SEQUENCE(LEN(modelText)-LEN(SUBSTITUTE(modelText,"/",""))+1,,1,LEN(modelText)),LEN(modelText))),'
            'parts,_xlfn.UNIQUE(_xlfn._xlws.FILTER(rawParts,rawParts<>"")), '
            f'partCustomers,IFERROR(VLOOKUP({lookup_key("parts")},lookupTable,2,FALSE),""),'
            'IF(AND(SUM(--(partCustomers=""))=0,'
            'SUM((modelKeys=TRANSPOSE(parts))*(customerValues<>TRANSPOSE(partCustomers)))=0),'
            '_xlfn.TEXTJOIN("/",TRUE,_xlfn.UNIQUE(partCustomers)),NA()))))),'
            f'"{CUSTOMER_LOOKUP_ERROR}")'
        )

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
