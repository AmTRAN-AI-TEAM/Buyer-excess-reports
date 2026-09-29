from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from openpyxl.utils import get_column_letter

from bom_alternates import BomAlternates
from workbook_io import ItemWorkbookData, ShortageRecord


INFO_HEADERS = ["Desc", "Planner", "GSD", "LT", "MOQ", "Price（USD)"]
EXCESS_METRIC_HEADERS = [
    "Excess\nstockQTY",
    "Excess\nstockAmount",
    "Excess\nPOQty",
    "Excess\nPOAmount",
    "Excess\nTotalAMT",
]
TRAILING_HEADERS = [
    "Vendor",
    "reason&action",
    "Customer\n客户",
    "MODEL",
    "MODELRemark\n机种",
    "Category\n分类",
    "Improve",
    "Previous Excess Total AMT",
    "是否呆",
]
BLANK_HEADER = ""
DEFAULT_DEMAND_LABELS = ("WO外demand",)


@dataclass(frozen=True)
class ExcessBuildResult:
    rows: list[list[Any]]
    row1_headers: list[Any]
    row2_headers: list[str]
    part_column_count: int
    input_part_count: int
    output_row_count: int
    demand_labels: tuple[str, ...]


@dataclass(frozen=True)
class _PreparedExcessRow:
    parts: list[str]
    main_record: ShortageRecord
    price: float | None
    overshortage_values: list[int]
    wo_values_by_period: list[list[int]]
    open_po_values: list[int]
    overshortage_total: int
    wo_totals: list[int]
    open_po_total: int
    customer: str | None
    model: str | None
    model_remark: str | None


def build_excess_rows(
    shortage_records: list[ShortageRecord],
    alternates: BomAlternates,
    item_data: ItemWorkbookData | None = None,
    *,
    blank_columns_before_excess: int = 0,
) -> ExcessBuildResult:
    item_data = item_data or ItemWorkbookData(metadata_by_part={}, prices_by_part={})
    demand_labels = _demand_labels(shortage_records)
    records_by_part = {record.part_no: record for record in shortage_records}
    known_parts = set(records_by_part)
    active_parts = {
        record.part_no
        for record in shortage_records
        if record.overshortage != 0 or record.wo_demand != 0 or record.open_po != 0
    }
    assigned_parts: set[str] = set()
    prepared_rows: list[_PreparedExcessRow] = []
    part_column_count = 1

    for record in shortage_records:
        if record.part_no in assigned_parts:
            continue

        full_group = _shortage_aware_group(
            record.part_no,
            alternates,
            records_by_part,
            known_parts,
            active_parts,
            assigned_parts,
        )

        displayed_parts = _ordered_parts(full_group)

        overshortage_values = [
            _metric(records_by_part, part, "overshortage") for part in displayed_parts
        ]
        wo_values_by_period = [
            [_demand_metric(records_by_part, part, index) for part in displayed_parts]
            for index in range(len(demand_labels))
        ]
        open_po_values = [_metric(records_by_part, part, "open_po") for part in displayed_parts]

        overshortage_total = sum(overshortage_values)
        wo_totals = [sum(values) for values in wo_values_by_period]
        open_po_total = sum(open_po_values)

        if overshortage_total == 0 and all(total == 0 for total in wo_totals) and open_po_total == 0:
            assigned_parts.update(part for part in displayed_parts if part in records_by_part)
            continue

        main_record = _first_record(displayed_parts, records_by_part)
        customer = item_data.customer_summary(displayed_parts)
        model = alternates.primary_model_for(
            record.part_no,
            known_parts=known_parts,
            active_parts=active_parts,
        ) or item_data.model_summary(displayed_parts)
        model_remark = alternates.model_remark_summary(displayed_parts)

        prepared_rows.append(
            _PreparedExcessRow(
                parts=displayed_parts,
                main_record=main_record,
                price=item_data.price_for(main_record.part_no),
                overshortage_values=overshortage_values,
                wo_values_by_period=wo_values_by_period,
                open_po_values=open_po_values,
                overshortage_total=overshortage_total,
                wo_totals=wo_totals,
                open_po_total=open_po_total,
                customer=customer,
                model=model,
                model_remark=model_remark,
            )
        )
        part_column_count = max(part_column_count, len(displayed_parts))
        assigned_parts.update(part for part in displayed_parts if part in records_by_part)

    row1_headers = build_row1_headers(
        part_column_count,
        demand_labels,
        blank_columns_before_excess=blank_columns_before_excess,
    )
    row2_headers = build_row2_headers(
        part_column_count,
        demand_labels,
        blank_columns_before_excess=blank_columns_before_excess,
    )
    rows = [
        _build_output_row(
            row,
            part_column_count,
            demand_labels,
            blank_columns_before_excess,
            row_number=index + 3,
            row2_headers=row2_headers,
        )
        for index, row in enumerate(prepared_rows)
    ]

    return ExcessBuildResult(
        rows=rows,
        row1_headers=row1_headers,
        row2_headers=row2_headers,
        part_column_count=part_column_count,
        input_part_count=len(shortage_records),
        output_row_count=len(rows),
        demand_labels=demand_labels,
    )


def build_row1_headers(
    part_column_count: int,
    demand_labels: tuple[str, ...],
    *,
    blank_columns_before_excess: int = 0,
) -> list[Any]:
    headers: list[Any] = [None] * (part_column_count + len(INFO_HEADERS))
    headers += ["Total"] + [None] * part_column_count
    for label in demand_labels:
        headers += [_group_header("Total", label, demand_labels)] + [None] * part_column_count
    headers += ["Total"] + [None] * part_column_count
    headers += ["Total"]
    headers += [None] * blank_columns_before_excess
    for label in demand_labels:
        headers += [_group_header("Excess", label, demand_labels)] + [None] * (len(EXCESS_METRIC_HEADERS) - 1)
    headers += [None] * len(TRAILING_HEADERS)
    return headers


def build_row2_headers(
    part_column_count: int,
    demand_labels: tuple[str, ...],
    *,
    blank_columns_before_excess: int = 0,
) -> list[str]:
    headers = (
        _part_headers(part_column_count)
        + INFO_HEADERS
        + ["Overshortage"]
        + [f"Overshortage{index}" for index in range(1, part_column_count + 1)]
    )
    for label in demand_labels:
        headers += [_demand_total_header(label, demand_labels)]
        headers += [
            _demand_part_header(label, index, demand_labels)
            for index in range(1, part_column_count + 1)
        ]
    headers += ["Openpo"] + [f"Open po{index}" for index in range(1, part_column_count + 1)]
    headers += ["Riskbuy"] + [BLANK_HEADER] * blank_columns_before_excess
    for label in demand_labels:
        headers += [_excess_header(header, label, demand_labels) for header in EXCESS_METRIC_HEADERS]
    headers += TRAILING_HEADERS
    return headers


def _part_headers(part_column_count: int) -> list[str]:
    return ["PartNo1"] + [
        f"替代料{index}" for index in range(2, part_column_count + 1)
    ]


def _build_output_row(
    row: _PreparedExcessRow,
    part_column_count: int,
    demand_labels: tuple[str, ...],
    blank_columns_before_excess: int,
    row_number: int,
    row2_headers: list[str],
) -> list[Any]:
    previous_total = None
    formulas = _row_formulas(row_number, row2_headers, part_column_count, demand_labels)

    values: list[Any] = (
        _pad(row.parts, part_column_count)
        + [
            row.main_record.description,
            row.main_record.planner,
            row.main_record.buyer,
            row.main_record.lt,
            row.main_record.moq,
            row.price,
            formulas["overshortage_total"],
        ]
        + _pad(row.overshortage_values, part_column_count, fill=0)
    )

    for formula, demand_values in zip(formulas["wo_totals"], row.wo_values_by_period):
        values += [formula] + _pad(demand_values, part_column_count, fill=0)

    values += [formulas["open_po_total"]]
    values += _pad(row.open_po_values, part_column_count, fill=0)
    values += [None]
    values += [None] * blank_columns_before_excess

    for excess_formulas in formulas["excess_groups"]:
        values += [
            excess_formulas["stock_qty"],
            excess_formulas["stock_amount"],
            excess_formulas["po_qty"],
            excess_formulas["po_amount"],
            excess_formulas["total_amount"],
        ]

    values += [
        None,
        None,
        row.customer,
        row.model,
        row.model_remark,
        None,
        formulas["improve"],
        previous_total,
        None,
    ]
    return values


def _row_formulas(
    row_number: int,
    row2_headers: list[str],
    part_column_count: int,
    demand_labels: tuple[str, ...],
) -> dict[str, Any]:
    overshortage_total = _cell_ref(row2_headers, "Overshortage", row_number)
    overshortage_first = _cell_ref(row2_headers, "Overshortage1", row_number)
    overshortage_last = _cell_ref(
        row2_headers,
        f"Overshortage{part_column_count}",
        row_number,
    )
    open_po_total = _cell_ref(row2_headers, "Openpo", row_number)
    open_po_first = _cell_ref(row2_headers, "Open po1", row_number)
    open_po_last = _cell_ref(row2_headers, f"Open po{part_column_count}", row_number)
    price = _cell_ref(row2_headers, "Price（USD)", row_number)
    previous_total = _cell_ref(row2_headers, "Previous Excess Total AMT", row_number)

    wo_total_formulas: list[str] = []
    excess_groups: list[dict[str, str]] = []
    first_total_amount_ref = ""
    for label in demand_labels:
        wo_total = _cell_ref(row2_headers, _demand_total_header(label, demand_labels), row_number)
        wo_first = _cell_ref(row2_headers, _demand_part_header(label, 1, demand_labels), row_number)
        wo_last = _cell_ref(
            row2_headers,
            _demand_part_header(label, part_column_count, demand_labels),
            row_number,
        )
        stock_qty = _cell_ref(row2_headers, _excess_header("Excess\nstockQTY", label, demand_labels), row_number)
        stock_amount = _cell_ref(row2_headers, _excess_header("Excess\nstockAmount", label, demand_labels), row_number)
        po_qty = _cell_ref(row2_headers, _excess_header("Excess\nPOQty", label, demand_labels), row_number)
        po_amount = _cell_ref(row2_headers, _excess_header("Excess\nPOAmount", label, demand_labels), row_number)
        total_amount = _cell_ref(row2_headers, _excess_header("Excess\nTotalAMT", label, demand_labels), row_number)
        if not first_total_amount_ref:
            first_total_amount_ref = total_amount

        wo_total_formulas.append(f"=ROUND(SUM({wo_first}:{wo_last}),0)")
        excess_groups.append(
            {
                "stock_qty": f"=ROUND({overshortage_total}-{wo_total},0)",
                "stock_amount": f'=IF({price}="","",ROUND(IF({stock_qty}<0,0,{stock_qty}*{price}),0))',
                "po_qty": f"=ROUND(IF({stock_qty}>0,{open_po_total},{stock_qty}+{open_po_total}),0)",
                "po_amount": f'=IF({price}="","",ROUND(IF({po_qty}>0,{po_qty}*{price},0),0))',
                "total_amount": f'=IF(OR({po_amount}="",{stock_amount}=""),-1,ROUND({po_amount}+{stock_amount},0))',
            }
        )

    return {
        "overshortage_total": f"=ROUND(SUM({overshortage_first}:{overshortage_last}),0)",
        "wo_totals": wo_total_formulas,
        "open_po_total": f"=ROUND(SUM({open_po_first}:{open_po_last}),0)",
        "excess_groups": excess_groups,
        "improve": f'=IF(OR({first_total_amount_ref}<0,{previous_total}=""),"",ROUND({first_total_amount_ref}-{previous_total},0))',
    }


def _cell_ref(headers: list[str], header: str, row_number: int) -> str:
    column = headers.index(header) + 1
    return f"{get_column_letter(column)}{row_number}"


def _shortage_aware_group(
    part_no: str,
    alternates: BomAlternates,
    records_by_part: dict[str, ShortageRecord],
    known_parts: set[str],
    active_parts: set[str],
    assigned_parts: set[str],
) -> list[str]:
    group = list(
        alternates.ordered_group_for(
            part_no,
            known_parts=known_parts,
            active_parts=active_parts,
        )
    )
    if part_no not in group:
        group.insert(0, part_no)

    visible_parts: list[str] = []
    for part in group:
        if part != part_no and part in assigned_parts:
            continue
        if part not in visible_parts:
            visible_parts.append(part)

    if part_no not in visible_parts:
        visible_parts.insert(0, part_no)

    return visible_parts


def _ordered_parts(group: list[str]) -> list[str]:
    unique_group = []
    for part in group:
        if part not in unique_group:
            unique_group.append(part)
    return unique_group


def _first_record(
    parts: list[str],
    records_by_part: dict[str, ShortageRecord],
) -> ShortageRecord:
    for part in parts:
        record = records_by_part.get(part)
        if record is not None:
            return record
    raise ValueError("Cannot build an Excess row without at least one shortage record")


def _metric(
    records_by_part: dict[str, ShortageRecord],
    part: str,
    field_name: str,
) -> int:
    record = records_by_part.get(part)
    if record is None:
        return 0
    return _round_integer(float(getattr(record, field_name)))


def _demand_metric(
    records_by_part: dict[str, ShortageRecord],
    part: str,
    demand_index: int,
) -> int:
    record = records_by_part.get(part)
    if record is None:
        return 0
    return _round_integer(float(record.wo_demand_at(demand_index)))


def _round_integer(value: float) -> int:
    if value >= 0:
        return int(value + 0.5)
    return int(value - 0.5)


def _pad(values: list[Any], length: int, fill: Any = None) -> list[Any]:
    return values[:length] + [fill] * max(0, length - len(values))


def _demand_labels(shortage_records: list[ShortageRecord]) -> tuple[str, ...]:
    for record in shortage_records:
        if record.demand_labels:
            return record.demand_labels
    return DEFAULT_DEMAND_LABELS


def _group_header(base: str, label: str, demand_labels: tuple[str, ...]) -> str:
    if len(demand_labels) == 1:
        return base
    return f"{base}\n{label}"


def _demand_total_header(label: str, demand_labels: tuple[str, ...]) -> str:
    if len(demand_labels) == 1:
        return "WO外demand"
    return f"WO外demand\n{label}"


def _demand_part_header(label: str, index: int, demand_labels: tuple[str, ...]) -> str:
    if len(demand_labels) == 1:
        return f"WO 外demand{index}"
    return f"WO 外demand{index}\n{label}"


def _excess_header(header: str, label: str, demand_labels: tuple[str, ...]) -> str:
    if len(demand_labels) == 1:
        return header
    return f"{header}\n{label}"
