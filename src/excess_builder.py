from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from bom_alternates import BomAlternates
from workbook_io import ShortageRecord


MAX_PART_COLUMNS = 7

ROW1_HEADERS = [
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    "Total",
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    "Total",
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    "Total",
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    "Total",
    "Excess",
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
    None,
]

ROW2_HEADERS = [
    "PartNo1",
    "替代料2",
    "替代料3",
    "替代料4",
    "替代料5",
    "替代料6",
    "替代料7",
    "Desc",
    "Planner",
    "GSD",
    "LT",
    "MOQ",
    "Price（USD)",
    "Overshortage",
    "Overshortage1",
    "Overshortage2",
    "Overshortage3",
    "Overshortage4",
    "Overshortage5",
    "Overshortage6",
    "Overshortage7",
    "WO外demand",
    "WO 外demand1",
    "WO 外demand2",
    "WO 外demand3",
    "WO 外demand4",
    "WO 外demand5",
    "WO 外demand6",
    "WO 外demand7",
    "Openpo",
    "Open po1",
    "Open po2",
    "Open po3",
    "Open po4",
    "Open po5",
    "Open po6",
    "Open po7",
    "Riskbuy",
    "Excess\nstockQTY",
    "Excess\nstockAmount",
    "Excess\nPOQty",
    "Excess\nPOAmount",
    "Excess\nTotalAMT",
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


@dataclass(frozen=True)
class ExcessBuildResult:
    rows: list[list[Any]]
    input_part_count: int
    output_row_count: int
    truncated_group_count: int


def build_excess_rows(
    shortage_records: list[ShortageRecord],
    alternates: BomAlternates,
) -> ExcessBuildResult:
    records_by_part = {record.part_no: record for record in shortage_records}
    source_order = {record.part_no: index for index, record in enumerate(shortage_records)}
    known_parts = set(records_by_part)
    active_parts = {
        record.part_no
        for record in shortage_records
        if record.overshortage != 0 or record.wo_demand != 0 or record.open_po != 0
    }
    assigned_parts: set[str] = set()
    rows: list[list[Any]] = []
    truncated_group_count = 0

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

        displayed_parts, was_truncated = _displayed_parts(
            record.part_no,
            full_group,
            source_order,
            alternates,
        )
        if was_truncated:
            truncated_group_count += 1

        overshortage_values = [_metric(records_by_part, part, "overshortage") for part in displayed_parts]
        wo_values = [_metric(records_by_part, part, "wo_demand") for part in displayed_parts]
        open_po_values = [_metric(records_by_part, part, "open_po") for part in displayed_parts]

        overshortage_total = sum(overshortage_values)
        wo_total = sum(wo_values)
        open_po_total = sum(open_po_values)

        if overshortage_total == 0 and wo_total == 0 and open_po_total == 0:
            assigned_parts.update(part for part in displayed_parts if part in records_by_part)
            continue

        main_record = _first_record(displayed_parts, records_by_part)
        stock_qty = overshortage_total - wo_total
        price = None
        stock_amount = 0
        po_qty = open_po_total if stock_qty > 0 else stock_qty + open_po_total
        po_amount = 0
        total_amount = stock_amount + po_amount
        previous_total = None
        improve = total_amount - (previous_total or 0)
        customer = alternates.customer_summary(displayed_parts)
        model = alternates.model_summary(displayed_parts)
        note = None
        if was_truncated:
            note = f"BOM alternate segment has more than {MAX_PART_COLUMNS} parts; only listed parts are totaled."

        rows.append(
            _pad(displayed_parts, MAX_PART_COLUMNS)
            + [
                main_record.description,
                main_record.planner,
                main_record.buyer,
                main_record.lt,
                main_record.moq,
                price,
                overshortage_total,
            ]
            + _pad(overshortage_values, MAX_PART_COLUMNS, fill=0)
            + [wo_total]
            + _pad(wo_values, MAX_PART_COLUMNS, fill=0)
            + [open_po_total]
            + _pad(open_po_values, MAX_PART_COLUMNS, fill=0)
            + [
                0,
                stock_qty,
                stock_amount,
                po_qty,
                po_amount,
                total_amount,
                None,
                note,
                customer,
                model,
                model,
                None,
                improve,
                previous_total,
                None,
            ]
        )
        assigned_parts.update(part for part in displayed_parts if part in records_by_part)

    return ExcessBuildResult(
        rows=rows,
        input_part_count=len(shortage_records),
        output_row_count=len(rows),
        truncated_group_count=truncated_group_count,
    )


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


def _displayed_parts(
    main_part: str,
    group: list[str],
    source_order: dict[str, int],
    alternates: BomAlternates,
) -> tuple[list[str], bool]:
    unique_group = []
    for part in group:
        if part not in unique_group:
            unique_group.append(part)

    if len(unique_group) <= MAX_PART_COLUMNS:
        return unique_group, False

    remaining = [part for part in unique_group if part != main_part]
    remaining.sort(
        key=lambda part: (
            0 if part in source_order else 1,
            source_order.get(part, 10**9),
            alternates.first_seen_order.get(part, 10**9),
        )
    )
    return [main_part] + remaining[: MAX_PART_COLUMNS - 1], True


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
) -> float:
    record = records_by_part.get(part)
    if record is None:
        return 0.0
    return float(getattr(record, field_name))


def _pad(values: list[Any], length: int, fill: Any = None) -> list[Any]:
    return values[:length] + [fill] * max(0, length - len(values))
