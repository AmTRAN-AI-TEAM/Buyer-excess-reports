from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


@dataclass
class BomAlternates:
    groups_by_part: dict[str, tuple[str, ...]]
    first_seen_order: dict[str, int]
    customers_by_part: dict[str, tuple[str, ...]] = field(default_factory=dict)
    models_by_part: dict[str, tuple[str, ...]] = field(default_factory=dict)

    def group_for(self, part_no: str) -> tuple[str, ...]:
        return self.groups_by_part.get(part_no, (part_no,))

    def ordered_group_for(self, part_no: str) -> tuple[str, ...]:
        group = self.group_for(part_no)
        return tuple(sorted(group, key=lambda part: self.first_seen_order.get(part, 10**9)))

    def customer_summary(self, parts: list[str]) -> str | None:
        return _join_limited(_unique_values(self.customers_by_part, parts))

    def model_summary(self, parts: list[str]) -> str | None:
        return _join_limited(_unique_values(self.models_by_part, parts))


class _UnionFind:
    def __init__(self) -> None:
        self.parent: dict[str, str] = {}

    def find(self, item: str) -> str:
        self.parent.setdefault(item, item)
        if self.parent[item] != item:
            self.parent[item] = self.find(self.parent[item])
        return self.parent[item]

    def union(self, left: str, right: str) -> None:
        left_root = self.find(left)
        right_root = self.find(right)
        if left_root != right_root:
            self.parent[right_root] = left_root


def load_bom_alternates(bom_path: Path) -> BomAlternates:
    workbook = load_workbook(bom_path, read_only=True, data_only=True)
    bom_sheet = _find_sheet(workbook.sheetnames, "BOM")
    bom_ws = workbook[bom_sheet]

    finished_goods = _load_finished_goods(workbook)
    union_find = _UnionFind()
    first_seen: dict[str, int] = {}
    customers: dict[str, list[str]] = defaultdict(list)
    models: dict[str, list[str]] = defaultdict(list)

    current_context: tuple[Any, Any, Any] | None = None
    current_base: str | None = None
    order = 0

    for row in bom_ws.iter_rows(min_row=2, values_only=True):
        top_assembly = row[0] if len(row) > 0 else None
        bill_level = row[2] if len(row) > 2 else None
        assembly_item = row[3] if len(row) > 3 else None
        item_seq = row[6] if len(row) > 6 else None
        component = row[7] if len(row) > 7 else None

        if component in (None, ""):
            continue

        part_no = str(component).strip()
        if part_no not in first_seen:
            first_seen[part_no] = order
            order += 1

        union_find.find(part_no)
        _collect_finished_good_metadata(
            part_no,
            top_assembly,
            finished_goods,
            customers,
            models,
        )

        context = (top_assembly, bill_level, assembly_item)
        if _is_replacement_marker(item_seq):
            if current_base and current_context == context:
                union_find.union(current_base, part_no)
            continue

        current_context = context
        current_base = part_no

    grouped: dict[str, list[str]] = defaultdict(list)
    for part_no in union_find.parent:
        grouped[union_find.find(part_no)].append(part_no)

    groups_by_part: dict[str, tuple[str, ...]] = {}
    for members in grouped.values():
        ordered_members = tuple(
            sorted(members, key=lambda part: first_seen.get(part, 10**9))
        )
        for part_no in ordered_members:
            groups_by_part[part_no] = ordered_members

    workbook.close()
    return BomAlternates(
        groups_by_part=groups_by_part,
        first_seen_order=first_seen,
        customers_by_part={key: tuple(values) for key, values in customers.items()},
        models_by_part={key: tuple(values) for key, values in models.items()},
    )


def _find_sheet(sheet_names: list[str], expected_name: str) -> str:
    for sheet_name in sheet_names:
        if sheet_name.lower() == expected_name.lower():
            return sheet_name
    raise ValueError(f"Cannot find required sheet {expected_name}")


def _load_finished_goods(workbook: Any) -> dict[str, tuple[Any, Any]]:
    sheet_name = next(
        (name for name in workbook.sheetnames if name.lower() == "成品料号"),
        None,
    )
    if sheet_name is None:
        return {}

    worksheet = workbook[sheet_name]
    mapping: dict[str, tuple[Any, Any]] = {}
    for row in worksheet.iter_rows(min_row=2, values_only=True):
        customer = row[0] if len(row) > 0 else None
        finished_good = row[1] if len(row) > 1 else None
        model = row[2] if len(row) > 2 else None
        if finished_good not in (None, ""):
            mapping[str(finished_good).strip()] = (customer, model)
    return mapping


def _collect_finished_good_metadata(
    part_no: str,
    top_assembly: Any,
    finished_goods: dict[str, tuple[Any, Any]],
    customers: dict[str, list[str]],
    models: dict[str, list[str]],
) -> None:
    if top_assembly in (None, ""):
        return

    customer, model = finished_goods.get(str(top_assembly).strip(), (None, None))
    _append_unique(customers[part_no], customer)
    _append_unique(models[part_no], model)


def _is_replacement_marker(value: Any) -> bool:
    if value is None:
        return False
    normalized = str(value).strip().replace("*", "").upper()
    return normalized == "R"


def _append_unique(values: list[str], value: Any) -> None:
    if value in (None, ""):
        return
    text = str(value).strip()
    if text and text not in values:
        values.append(text)


def _unique_values(source: dict[str, tuple[str, ...]], parts: list[str]) -> list[str]:
    values: list[str] = []
    for part_no in parts:
        for value in source.get(part_no, ()):
            if value not in values:
                values.append(value)
    return values


def _join_limited(values: list[str], limit: int = 6) -> str | None:
    if not values:
        return None
    if len(values) <= limit:
        return "/".join(values)
    return "/".join(values[:limit]) + f"/+{len(values) - limit} more"
