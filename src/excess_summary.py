from __future__ import annotations

from dataclasses import dataclass, replace
from math import ceil
from pathlib import Path
import unicodedata
from typing import Any

from openpyxl import Workbook
from openpyxl.cell.rich_text import CellRichText, InlineFont, TextBlock
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from bom_alternates import BomComponentUsage
from excess_builder import ExcessSummaryMetrics, ExcessSummarySourceRow
from summary_rules import is_allowed_summary_item_type
from workbook_io import ItemWorkbookData


SUMMARY_HEADERS = [
    "AVTC",
    "ON HAND",
    "OPO",
    "ON HAND AMT USD",
    "OPO AMT\nUSD",
    "TTL AMT",
    "Gating",
    "AVTC建议加单",
    "消耗率",
    "建議料號可消耗\nUSD",
    "所有相關料號原有\nExcess USD（加單前）",
    "所有相關料號可消耗\nUSD",
    "加單後剩餘\nExcess USD",
    "料號供應狀態",
    "需額外採購",
    "AVTC 呆滯 预计TO-BE",
]
CONSUMPTION_HEADERS = [
    "建議編號",
    "Excess來源 MODEL",
    "BOM MODEL",
    "TOP_ASSEMBLY_ITEM",
    "Top排名",
    "觸發料號／替代料群",
    "建議加單台數",
    "建議料號可消耗 USD",
    "所有相關料號原有 Excess USD（加單前）",
    "所有相關料號可消耗 USD",
    "加單後剩餘 Excess USD",
    "消耗率",
    "料號供應狀態",
    "需額外採購摘要",
    "計算說明",
]
SUPPLY_DETAIL_HEADERS = [
    "建議編號",
    "BOM MODEL",
    "TOP_ASSEMBLY_ITEM",
    "必要料號／替代料群",
    "對應 BOM 列",
    "每台用量",
    "加單需求量",
    "現有 Excess stock QTY",
    "現有 Excess USD",
    "現有 OPO QTY",
    "可消耗數量",
    "需額外採購數量",
    "加單後剩餘 QTY",
    "單價（USD）",
    "可消耗金額 USD",
    "剩餘 Excess USD",
    "判定",
]
DETAIL_FIXED_HEADERS = [
    "排名",
    "來源列",
    "MODEL",
]
DETAIL_TAIL_HEADERS = [
    "Price（USD)",
    "Excess stock QTY",
    "Excess stock Amount USD",
    "Excess POQTY",
    "Excess POAmount USD",
    "Excess TotalAMT USD",
]
UNCLASSIFIED_MODEL = "未分類（MODEL空白）"
TOP_FIVE_LIMIT = 5
SUMMARY_TOP_LIMIT = 2
GATING_TOP_LIMIT = 3
SELECTED_MODEL_MARKER = "此料號對應使用者所選model"
_BomRowDetail = tuple[int, str, float]


@dataclass
class _ModelSummary:
    key: str
    model: str
    on_hand_qty: int = 0
    opo_qty: int = 0
    on_hand_amount: int = 0
    opo_amount: int = 0
    total_amount: int = 0


@dataclass(frozen=True)
class _TopFiveEntry:
    rank: int
    source_row: int
    source: ExcessSummarySourceRow
    stock_qty: int
    stock_amount: int
    po_qty: int
    po_amount: int | None
    total_amount: int


@dataclass(frozen=True)
class _BomMatch:
    model_key: str
    target_model: str
    calculation_top_assembly: str
    matched_part: str
    usage_qty: float
    bom_rows: tuple[_BomRowDetail, ...]
    slot_key: tuple[str, str, frozenset[str]]


@dataclass(frozen=True)
class _InventoryAllocation:
    source_row: int
    source: ExcessSummarySourceRow


@dataclass(frozen=True)
class _MaterialCheck:
    parts: tuple[str, ...]
    bom_rows: tuple[_BomRowDetail, ...]
    usage_qty: float
    required_qty: float
    available_qty: float
    po_qty: float
    consumed_qty: float
    extra_purchase_qty: float
    remaining_qty: float
    available_amount: float | None
    consumed_amount: float | None
    remaining_amount: float | None
    price: float | None
    status: str
    note: str | None = None


@dataclass(frozen=True)
class _ProductionSuggestion:
    suggestion_id: str
    target_model: str
    calculation_top_assembly: str
    quantity: int
    source_rank: int
    trigger_parts: tuple[str, ...]
    trigger_consumed_amount: float | None
    total_available_amount: float | None
    total_consumed_amount: float | None
    total_remaining_amount: float | None
    consumption_rate: float | None
    supply_status: str
    shortage_details: tuple[str, ...]
    material_checks: tuple[_MaterialCheck, ...]
    bom_rows: tuple[_BomRowDetail, ...] = ()


@dataclass(frozen=True)
class _BomUsageSlot:
    model_key: str
    model_name: str
    finished_good_key: str
    finished_good_name: str
    parts: frozenset[str]
    part_usage: dict[str, float]
    bom_rows: tuple[_BomRowDetail, ...]
    row_number: int


def write_excess_summary(
    source_rows: tuple[ExcessSummarySourceRow, ...],
    period_label: str,
    output_file: Path,
    *,
    component_usages: tuple[BomComponentUsage, ...] = (),
    customer: str | None = None,
    selected_models: tuple[str, ...] | None = None,
    item_data: ItemWorkbookData | None = None,
) -> tuple[int, int]:
    """Write a model summary and each model's five highest stock amounts.

    Values are calculated before the Excess workbook's formulas are written, so
    this workbook contains concrete values and does not depend on Excel's formula
    cache being populated.
    """
    output_file.parent.mkdir(parents=True, exist_ok=True)
    summary_by_model: dict[str, _ModelSummary] = {}
    all_bom_matches_by_source_row = _bom_matches_by_source_model(
        source_rows,
        component_usages,
    )
    eligible_bom_matches_by_source_row = _filter_bom_matches_by_item_type(
        all_bom_matches_by_source_row,
        item_data,
    )
    top_five_by_model = _top_five_entries(
        source_rows,
        eligible_source_rows=set(eligible_bom_matches_by_source_row),
    )
    usage_slots = _bom_usage_slots(component_usages)

    for source in source_rows:
        metrics = _last_period_metrics(source)
        model_key = _model_key(source.model)
        model_display = model_key or UNCLASSIFIED_MODEL
        normalized_key = model_key or ""
        summary = summary_by_model.setdefault(
            normalized_key,
            _ModelSummary(normalized_key, model_display),
        )
        summary.on_hand_qty += max(metrics.stock_qty, 0)
        summary.opo_qty += metrics.po_qty
        summary.on_hand_amount += metrics.stock_amount or 0
        summary.opo_amount += metrics.po_amount or 0
        summary.total_amount += metrics.total_amount

    workbook = Workbook()
    recommendations_ws = workbook.active
    recommendations_ws.title = "加單建議"
    summary_sheet_title = (
        "Raken summary"
        if customer and customer.upper() == "RAKEN"
        else "AVTC Summary"
    )
    summary_ws = workbook.create_sheet(summary_sheet_title)
    summary_ws["A1"] = f"選取區間：{period_label}"
    summary_headers = list(SUMMARY_HEADERS)
    if customer and customer.upper() == "RAKEN":
        summary_headers[0] = "Raken"
        summary_headers[7] = "Raken建议加单"
        summary_headers[15] = "Raken 呆滯 预计TO-BE"
    for column, header in enumerate(summary_headers, start=1):
        summary_ws.cell(3, column, header)

    summaries = sorted(
        summary_by_model.values(),
        key=lambda item: (-item.on_hand_amount, item.model.casefold()),
    )
    (
        _recommendable_entries_by_model,
        suggestions_by_model,
    ) = _recommend_production_orders(
        top_five_by_model,
        source_rows,
        all_bom_matches_by_source_row,
        usage_slots,
        item_data,
    )
    suggestions_by_model = _number_suggestions(suggestions_by_model)
    top_five = [
        entry
        for summary in summaries
        for entry in top_five_by_model.get(summary.key, ())
    ]
    for summary in summaries:
        suggestions = suggestions_by_model.get(summary.key, ())
        summary_top_entries = top_five_by_model.get(summary.key, ())[:GATING_TOP_LIMIT]
        gating = _gating_rich_text(
            summary,
            summary_top_entries,
            eligible_bom_matches_by_source_row,
            item_data,
        )
        summary_ws.append(
            _summary_values(
                summary.model,
                summary.on_hand_qty,
                summary.opo_qty,
                summary.on_hand_amount,
                summary.opo_amount,
                summary.total_amount,
                gating,
                suggestions,
            )
        )

    total_row = summary_ws.max_row + 1
    summary_ws.append(
        [
            "總計",
            sum(item.on_hand_qty for item in summaries),
            sum(item.opo_qty for item in summaries),
            sum(item.on_hand_amount for item in summaries),
            sum(item.opo_amount for item in summaries),
            sum(item.total_amount for item in summaries),
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
    )
    _format_summary_sheet(summary_ws, total_row)
    _write_recommendations_sheet(
        recommendations_ws,
        summaries,
        suggestions_by_model,
        period_label,
        selected_models,
        summary_headers,
        top_five_by_model,
        eligible_bom_matches_by_source_row,
        item_data,
    )

    consumption_ws = workbook.create_sheet("消耗建議")
    _write_consumption_sheet(
        consumption_ws,
        summaries,
        suggestions_by_model,
        period_label,
    )

    supply_ws = workbook.create_sheet("供料明細")
    _write_supply_detail_sheet(
        supply_ws,
        summaries,
        suggestions_by_model,
        period_label,
    )

    detail_ws = workbook.create_sheet("Top 5明細")
    detail_ws["A1"] = f"每個 MODEL 依最後一期 Excess stock Amount 取前五名：{period_label}"
    detail_headers = _detail_headers(top_five)
    for column, header in enumerate(detail_headers, start=1):
        detail_ws.cell(3, column, header)
    for entry in top_five:
        detail_ws.append(_detail_values(entry, len(detail_headers)))
    _format_detail_sheet(detail_ws, len(detail_headers), detail_ws.max_row)

    workbook.save(output_file)
    return len(summaries), len(top_five)


def _filter_bom_matches_by_item_type(
    bom_matches_by_source_row: dict[int, _BomMatch],
    item_data: ItemWorkbookData | None,
) -> dict[int, _BomMatch]:
    """Keep BOM matches whose actual matched part has an allowed Item TYPE."""
    if item_data is None:
        return {}

    filtered: dict[int, _BomMatch] = {}
    for source_row, match in bom_matches_by_source_row.items():
        item_type = item_data.item_type_for(match.matched_part)
        if is_allowed_summary_item_type(item_type):
            filtered[source_row] = match
    return filtered


def _top_five_entries(
    source_rows: tuple[ExcessSummarySourceRow, ...],
    *,
    eligible_source_rows: set[int] | None = None,
) -> dict[str, list[_TopFiveEntry]]:
    candidates_by_model: dict[str, list[tuple[int, int, ExcessSummarySourceRow]]] = {}
    for index, source in enumerate(source_rows):
        source_row = index + 3
        if eligible_source_rows is not None and source_row not in eligible_source_rows:
            continue
        if not source.period_metrics:
            continue
        metrics = source.period_metrics[-1]
        if metrics.stock_amount is not None and metrics.stock_amount > 0:
            model_key = _model_key(source.model) or ""
            candidates_by_model.setdefault(model_key, []).append(
                (metrics.stock_amount, index, source)
            )

    entries_by_model: dict[str, list[_TopFiveEntry]] = {}
    for model_key, candidates in candidates_by_model.items():
        candidates.sort(key=lambda item: (-item[0], item[1]))
        entries: list[_TopFiveEntry] = []
        for rank, (stock_amount, index, source) in enumerate(
            candidates[:TOP_FIVE_LIMIT],
            start=1,
        ):
            metrics = source.period_metrics[-1]
            entries.append(
                _TopFiveEntry(
                    rank=rank,
                    source_row=index + 3,
                    source=source,
                    stock_qty=metrics.stock_qty,
                    stock_amount=stock_amount,
                    po_qty=metrics.po_qty,
                    po_amount=metrics.po_amount,
                    total_amount=metrics.total_amount,
                )
            )
        entries_by_model[model_key] = entries
    return entries_by_model


def _bom_usage_slots(
    component_usages: tuple[BomComponentUsage, ...],
) -> tuple[_BomUsageSlot, ...]:
    raw_slots: dict[
        tuple[str, str, frozenset[str]],
        dict[str, Any],
    ] = {}
    for usage in component_usages:
        model_name = (usage.model or "").strip()
        model_key = model_name.casefold()
        finished_good_name = (usage.finished_good or "").strip()
        finished_good_key = finished_good_name.casefold()
        parts = frozenset(_part_key(part) for part in usage.parts if _part_key(part))
        if not model_key or not finished_good_key or not parts:
            continue

        slot_key = (model_key, finished_good_key, parts)
        slot = raw_slots.setdefault(
            slot_key,
            {
                "model_name": model_name,
                "finished_good_name": finished_good_name,
                "part_usage": {},
                "bom_rows": [],
                "row_number": usage.row_number,
            },
        )
        slot["row_number"] = min(slot["row_number"], usage.row_number)
        part_usage_rows = usage.part_usage_rows or tuple(
            (part, usage.row_number, quantity)
            for part, quantity in (
                usage.part_usage_qtys
                or tuple((part, usage.usage_qty) for part in usage.parts)
            )
        )
        for part, row_number, quantity in part_usage_rows:
            part_key = _part_key(part)
            if part_key not in parts or quantity is None or quantity <= 0:
                continue
            part_usage = slot["part_usage"]
            part_usage[part_key] = part_usage.get(part_key, 0.0) + float(quantity)
            slot["bom_rows"].append((row_number, str(part).strip(), float(quantity)))

    return tuple(
        _BomUsageSlot(
            model_key=model_key,
            model_name=slot["model_name"],
            finished_good_key=finished_good_key,
            finished_good_name=slot["finished_good_name"],
            parts=parts,
            part_usage=slot["part_usage"],
            bom_rows=tuple(slot["bom_rows"]),
            row_number=slot["row_number"],
        )
        for (model_key, finished_good_key, parts), slot in raw_slots.items()
    )


def _match_source_to_model(
    source: ExcessSummarySourceRow,
    model_key: str,
    usage_slots: tuple[_BomUsageSlot, ...],
) -> _BomMatch | None:
    part_names = {
        _part_key(part): str(part).strip()
        for part in source.parts
        if _part_key(part)
    }
    pool_parts = frozenset(part_names)
    if not pool_parts:
        return None
    primary_part = next(iter(part_names), "")

    candidates: list[tuple[tuple[Any, ...], str, _BomUsageSlot, float]] = []
    for slot_index, slot in enumerate(usage_slots):
        if slot.model_key != model_key or not pool_parts.issubset(slot.part_usage):
            continue
        usage_values = {slot.part_usage[part] for part in pool_parts}
        if len(usage_values) > 1:
            continue
        for matched_part in pool_parts:
            quantity = slot.part_usage[matched_part]
            sort_key = (
                -quantity,
                slot.row_number,
                0 if matched_part == primary_part else 1,
                matched_part,
                slot.finished_good_key,
                tuple(sorted(slot.parts)),
                slot_index,
            )
            candidates.append((sort_key, matched_part, slot, quantity))

    if not candidates:
        return None
    _sort_key, matched_part, slot, usage_qty = min(candidates, key=lambda item: item[0])
    return _BomMatch(
        model_key=slot.model_key,
        target_model=slot.model_name,
        calculation_top_assembly=slot.finished_good_name,
        matched_part=matched_part,
        usage_qty=usage_qty,
        bom_rows=tuple(
            row
            for row in slot.bom_rows
            if _part_key(row[1]) == matched_part
        ),
        slot_key=(slot.model_key, slot.finished_good_key, slot.parts),
    )


def _bom_matches_by_source_model(
    source_rows: tuple[ExcessSummarySourceRow, ...],
    component_usages: tuple[BomComponentUsage, ...],
) -> dict[int, _BomMatch]:
    slot_lists_by_model: dict[str, list[_BomUsageSlot]] = {}
    for slot in _bom_usage_slots(component_usages):
        slot_lists_by_model.setdefault(slot.model_key, []).append(slot)
    usage_slots_by_model = {
        model_key: tuple(slots)
        for model_key, slots in slot_lists_by_model.items()
    }

    matches: dict[int, _BomMatch] = {}
    for index, source in enumerate(source_rows):
        model_key = _model_key(source.model)
        if not model_key:
            continue
        match = _match_source_to_model(
            source,
            model_key.casefold(),
            usage_slots_by_model.get(model_key.casefold(), ()),
        )
        if match is not None:
            matches[index + 3] = match
    return matches


def _recommend_production_orders(
    top_five_by_model: dict[str, list[_TopFiveEntry]],
    source_rows: tuple[ExcessSummarySourceRow, ...],
    bom_matches_by_source_row: dict[int, _BomMatch],
    usage_slots: tuple[_BomUsageSlot, ...],
    item_data: ItemWorkbookData | None,
) -> tuple[
    dict[str, tuple[_TopFiveEntry, ...]],
    dict[str, tuple[_ProductionSuggestion, ...]],
]:
    """Select up to two entries per model and calculate each independently."""
    if not bom_matches_by_source_row:
        return {}, {}

    source_by_row = {
        index + 3: source
        for index, source in enumerate(source_rows)
    }
    allocations_by_slot: dict[
        tuple[str, str, frozenset[str]],
        list[_InventoryAllocation],
    ] = {}
    for source_row, match in bom_matches_by_source_row.items():
        source = source_by_row.get(source_row)
        if source is not None:
            allocations_by_slot.setdefault(match.slot_key, []).append(
                _InventoryAllocation(source_row, source)
            )

    assignments = {
        entry.source_row: bom_matches_by_source_row[entry.source_row]
        for entries in top_five_by_model.values()
        for entry in entries
        if entry.source_row in bom_matches_by_source_row
    }

    recommendable_entries_by_model: dict[str, tuple[_TopFiveEntry, ...]] = {}
    for model, entries in top_five_by_model.items():
        source_model_key = model.casefold()
        recommendable_entries = tuple(
            entry
            for entry in entries
            if (
                entry.source_row in assignments
                and assignments[entry.source_row].model_key == source_model_key
            )
        )[:SUMMARY_TOP_LIMIT]
        if recommendable_entries:
            recommendable_entries_by_model[model] = recommendable_entries

    suggestions_by_model: dict[str, list[_ProductionSuggestion]] = {}
    for entry in (
        entry
        for entries in recommendable_entries_by_model.values()
        for entry in entries
    ):
        match = assignments[entry.source_row]
        if entry.stock_qty <= 0 or match.usage_qty <= 0:
            continue

        quantity = ceil(entry.stock_qty / match.usage_qty)
        scenario_slots = tuple(
            slot
            for slot in usage_slots
            if (
                slot.model_key == match.model_key
                and slot.finished_good_key
                == match.slot_key[1]
            )
        )
        material_checks = _material_checks_for_scenario(
            scenario_slots,
            allocations_by_slot,
            quantity,
            item_data,
        )
        trigger_consumed_amount = _source_amount_for_consumed_qty(
            entry.source,
            min(max(entry.stock_qty, 0), quantity * match.usage_qty),
        )
        total_available_amount = _sum_optional_amount(
            check.available_amount for check in material_checks
        )
        total_consumed_amount = _sum_optional_amount(
            check.consumed_amount for check in material_checks
        )
        total_remaining_amount = _sum_optional_amount(
            check.remaining_amount for check in material_checks
        )
        consumption_rate = (
            None
            if (
                total_available_amount is None
                or total_consumed_amount is None
                or total_available_amount <= 0
            )
            else total_consumed_amount / total_available_amount
        )
        shortages = tuple(
            f"{_display_parts(check.parts)} +{_format_number(check.extra_purchase_qty)}"
            for check in material_checks
            if check.extra_purchase_qty > 1e-9
        )
        supply_status = "不足" if shortages else "足夠"
        owner = _model_key(entry.source.model) or ""
        suggestions_by_model.setdefault(owner, []).append(
            _ProductionSuggestion(
                suggestion_id="",
                target_model=match.target_model,
                calculation_top_assembly=match.calculation_top_assembly,
                quantity=quantity,
                source_rank=entry.rank,
                trigger_parts=tuple(entry.source.parts),
                trigger_consumed_amount=trigger_consumed_amount,
                total_available_amount=total_available_amount,
                total_consumed_amount=total_consumed_amount,
                total_remaining_amount=total_remaining_amount,
                consumption_rate=consumption_rate,
                supply_status=supply_status,
                shortage_details=shortages,
                material_checks=material_checks,
                bom_rows=match.bom_rows,
            )
        )

    return (
        recommendable_entries_by_model,
        {
            model: tuple(
                sorted(
                    suggestions,
                    key=lambda suggestion: (
                        -(
                            suggestion.trigger_consumed_amount
                            if suggestion.trigger_consumed_amount is not None
                            else -1
                        ),
                        suggestion.quantity,
                        suggestion.target_model.casefold(),
                        suggestion.calculation_top_assembly.casefold(),
                    ),
                )
            )
            for model, suggestions in suggestions_by_model.items()
        },
    )


def _material_checks_for_scenario(
    scenario_slots: tuple[_BomUsageSlot, ...],
    allocations_by_slot: dict[
        tuple[str, str, frozenset[str]],
        list[_InventoryAllocation],
    ],
    order_quantity: int,
    item_data: ItemWorkbookData | None,
) -> tuple[_MaterialCheck, ...]:
    checks: list[_MaterialCheck] = []
    for slot in sorted(scenario_slots, key=lambda item: item.row_number):
        usage_qty, usage_note = _slot_usage_qty(slot)
        if usage_qty <= 0:
            continue
        allocations = allocations_by_slot.get(
            (slot.model_key, slot.finished_good_key, slot.parts),
            [],
        )
        required_qty = order_quantity * usage_qty
        available_qty = sum(
            max(_last_period_metrics(allocation.source).stock_qty, 0)
            for allocation in allocations
        )
        po_qty = sum(
            max(_last_period_metrics(allocation.source).po_qty, 0)
            for allocation in allocations
        )
        remaining_need = required_qty
        consumed_qty = 0.0
        consumed_amounts: list[float | None] = []
        remaining_amounts: list[float | None] = []
        available_amounts: list[float | None] = []
        for allocation in allocations:
            metrics = _last_period_metrics(allocation.source)
            source_qty = max(metrics.stock_qty, 0)
            consumed_from_source = min(source_qty, max(remaining_need, 0))
            remaining_need -= consumed_from_source
            consumed_qty += consumed_from_source
            available_amounts.append(
                metrics.stock_amount if source_qty > 0 else 0.0
            )
            consumed_amounts.append(
                _source_amount_for_consumed_qty(
                    allocation.source,
                    consumed_from_source,
                )
            )
            remaining_amounts.append(
                _source_amount_for_consumed_qty(
                    allocation.source,
                    max(source_qty - consumed_from_source, 0),
                )
            )

        if not allocations:
            available_amount = 0.0
            consumed_amount = 0.0
            remaining_amount = 0.0
        else:
            available_amount = _sum_optional_amount(available_amounts)
            consumed_amount = _sum_optional_amount(consumed_amounts)
            remaining_amount = _sum_optional_amount(remaining_amounts)

        extra_purchase_qty = max(required_qty - available_qty, 0)
        remaining_qty = max(available_qty - required_qty, 0)
        note_parts: list[str] = []
        if not allocations:
            note_parts.append("找不到此必要料號的 Excess 來源")
        if usage_note:
            note_parts.append(usage_note)
        if any(
            amount is None
            for amount in (*available_amounts, *consumed_amounts, *remaining_amounts)
        ):
            note_parts.append("缺少單價，USD 金額無法完整計算")
        price = next(
            (
                allocation.source.price
                for allocation in allocations
                if allocation.source.price is not None
            ),
            _fallback_price_for_parts(item_data, slot.parts),
        )
        checks.append(
            _MaterialCheck(
                parts=tuple(sorted(slot.parts)),
                bom_rows=slot.bom_rows,
                usage_qty=usage_qty,
                required_qty=required_qty,
                available_qty=available_qty,
                po_qty=po_qty,
                consumed_qty=consumed_qty,
                extra_purchase_qty=extra_purchase_qty,
                remaining_qty=remaining_qty,
                available_amount=available_amount,
                consumed_amount=consumed_amount,
                remaining_amount=remaining_amount,
                price=price,
                status="不足" if extra_purchase_qty > 1e-9 else "足夠",
                note="；".join(note_parts) or None,
            )
        )
    return tuple(checks)


def _slot_usage_qty(slot: _BomUsageSlot) -> tuple[float, str | None]:
    values = tuple(value for value in slot.part_usage.values() if value > 0)
    if not values:
        return 0.0, "BOM 缺少有效 EXTENDED_QT"
    if len({round(value, 9) for value in values}) == 1:
        return values[0], None
    return max(values), "替代料 EXTENDED_QT 不一致，採用較高用量"


def _source_amount_for_consumed_qty(
    source: ExcessSummarySourceRow,
    consumed_qty: float,
) -> float | None:
    metrics = _last_period_metrics(source)
    source_qty = max(metrics.stock_qty, 0)
    if consumed_qty <= 0:
        return 0.0
    if source_qty <= 0:
        return 0.0
    if metrics.stock_amount is None:
        return None
    return metrics.stock_amount * min(consumed_qty, source_qty) / source_qty


def _sum_optional_amount(values: Any) -> float | None:
    values = tuple(values)
    if any(value is None for value in values):
        return None
    return sum(float(value) for value in values)


def _fallback_price_for_parts(
    item_data: ItemWorkbookData | None,
    parts: frozenset[str],
) -> float | None:
    if item_data is None:
        return None
    price_by_key = {
        str(part).strip().casefold(): price
        for part, price in item_data.prices_by_part.items()
    }
    for part in sorted(parts):
        price = price_by_key.get(part.casefold())
        if price is not None:
            return price
    return None


def _number_suggestions(
    suggestions_by_model: dict[str, tuple[_ProductionSuggestion, ...]],
) -> dict[str, tuple[_ProductionSuggestion, ...]]:
    counter = 1
    numbered: dict[str, tuple[_ProductionSuggestion, ...]] = {}
    for model, suggestions in suggestions_by_model.items():
        model_suggestions: list[_ProductionSuggestion] = []
        for suggestion in suggestions:
            model_suggestions.append(
                replace(suggestion, suggestion_id=f"S{counter:03d}")
            )
            counter += 1
        numbered[model] = tuple(model_suggestions)
    return numbered


def _summary_values(
    model: str,
    on_hand_qty: int,
    opo_qty: int,
    on_hand_amount: int,
    opo_amount: int,
    total_amount: int,
    gating: CellRichText | None,
    suggestions: tuple[_ProductionSuggestion, ...],
) -> list[Any]:
    return [
        model,
        on_hand_qty,
        opo_qty,
        on_hand_amount,
        opo_amount,
        total_amount,
        gating,
        _suggestion_lines(suggestions, _suggestion_order_text),
        _suggestion_lines(suggestions, _suggestion_rate_text),
        _suggestion_lines(suggestions, _suggestion_trigger_amount_text),
        _suggestion_lines(suggestions, _suggestion_total_available_text),
        _suggestion_lines(suggestions, _suggestion_total_consumed_text),
        _suggestion_lines(suggestions, _suggestion_remaining_text),
        _suggestion_lines(suggestions, lambda suggestion: suggestion.supply_status),
        _suggestion_lines(suggestions, _suggestion_shortage_text),
        None,
    ]


def _suggestion_lines(
    suggestions: tuple[_ProductionSuggestion, ...],
    formatter: Any,
) -> str | None:
    if not suggestions:
        return None
    return "\n".join(
        f"{suggestion.suggestion_id}｜{formatter(suggestion)}"
        for suggestion in suggestions
    )


def _suggestion_rate_text(suggestion: _ProductionSuggestion) -> str:
    if suggestion.consumption_rate is None:
        return "無法計算"
    return f"{suggestion.consumption_rate:.2%}"


def _suggestion_trigger_amount_text(suggestion: _ProductionSuggestion) -> str:
    return _format_usd(suggestion.trigger_consumed_amount)


def _suggestion_total_consumed_text(suggestion: _ProductionSuggestion) -> str:
    return _format_usd(suggestion.total_consumed_amount)


def _suggestion_total_available_text(suggestion: _ProductionSuggestion) -> str:
    return _format_usd(suggestion.total_available_amount)


def _suggestion_remaining_text(suggestion: _ProductionSuggestion) -> str:
    return _format_usd(suggestion.total_remaining_amount)


def _suggestion_shortage_text(suggestion: _ProductionSuggestion) -> str:
    if not suggestion.shortage_details:
        return "無"
    preview_limit = 5
    preview = "；".join(suggestion.shortage_details[:preview_limit])
    remaining = len(suggestion.shortage_details) - preview_limit
    if remaining > 0:
        return (
            f"共 {len(suggestion.shortage_details)} 筆料號／替代料群不足："
            f"{preview}；另有 {remaining} 筆，詳見供料明細"
        )
    return f"共 {len(suggestion.shortage_details)} 筆料號／替代料群不足：{preview}"


def _format_usd(value: float | None) -> str:
    if value is None:
        return "無法計算"
    return f"USD {value:,.2f}"


def _write_consumption_sheet(
    worksheet: Any,
    summaries: list[_ModelSummary],
    suggestions_by_model: dict[str, tuple[_ProductionSuggestion, ...]],
    period_label: str,
) -> None:
    worksheet["A1"] = (
        f"消耗建議：{period_label}；每筆建議獨立計算，庫存基準為最後一期 Excess stock QTY"
    )
    for column, header in enumerate(CONSUMPTION_HEADERS, start=1):
        worksheet.cell(3, column, header)
    for summary in summaries:
        for suggestion in suggestions_by_model.get(summary.key, ()):
            worksheet.append(
                [
                    suggestion.suggestion_id,
                    summary.model,
                    suggestion.target_model,
                    suggestion.calculation_top_assembly,
                    suggestion.source_rank,
                    _display_parts(suggestion.trigger_parts),
                    suggestion.quantity,
                    suggestion.trigger_consumed_amount,
                    suggestion.total_available_amount,
                    suggestion.total_consumed_amount,
                    suggestion.total_remaining_amount,
                    suggestion.consumption_rate,
                    suggestion.supply_status,
                    _suggestion_shortage_text(suggestion),
                    (
                        f"Top {suggestion.source_rank} 料號庫存 ÷ BOM 每台用量，"
                        "向上取整；消耗率 = 所有相關料號可消耗 USD ÷ "
                        "所有相關料號原有 Excess USD（加單前）；"
                        "供料檢查請見「供料明細」"
                    ),
                ]
            )
    _format_consumption_sheet(worksheet, worksheet.max_row)


def _write_supply_detail_sheet(
    worksheet: Any,
    summaries: list[_ModelSummary],
    suggestions_by_model: dict[str, tuple[_ProductionSuggestion, ...]],
    period_label: str,
) -> None:
    worksheet["A1"] = (
        f"供料明細：{period_label}；OPO 僅供參考，不納入現有 Excess 消耗率"
    )
    for column, header in enumerate(SUPPLY_DETAIL_HEADERS, start=1):
        worksheet.cell(3, column, header)
    for summary in summaries:
        for suggestion in suggestions_by_model.get(summary.key, ()):
            for check in suggestion.material_checks:
                worksheet.append(
                    [
                        suggestion.suggestion_id,
                        suggestion.target_model,
                        suggestion.calculation_top_assembly,
                        _display_parts(check.parts),
                        _format_bom_rows(check.bom_rows),
                        check.usage_qty,
                        check.required_qty,
                        check.available_qty,
                        check.available_amount,
                        check.po_qty,
                        check.consumed_qty,
                        check.extra_purchase_qty,
                        check.remaining_qty,
                        check.price,
                        check.consumed_amount,
                        check.remaining_amount,
                        "；".join(
                            value
                            for value in (check.status, check.note)
                            if value
                        ),
                    ]
                )
    _format_supply_detail_sheet(worksheet, worksheet.max_row)


def _write_recommendations_sheet(
    worksheet: Any,
    summaries: list[_ModelSummary],
    suggestions_by_source_model: dict[str, tuple[_ProductionSuggestion, ...]],
    period_label: str,
    selected_models: tuple[str, ...] | None,
    headers: list[str],
    top_five_by_model: dict[str, list[_TopFiveEntry]],
    bom_matches_by_source_row: dict[int, _BomMatch],
    item_data: ItemWorkbookData | None,
) -> None:
    worksheet["A1"] = (
        f"選取區間：{period_label}；依 Excess 來源 MODEL 篩選，"
        "Gating 列 Top 1～Top 3 庫存及 BOM MODEL 對應料號"
    )
    for column, header in enumerate(headers, start=1):
        worksheet.cell(3, column, header)

    selected_keys = (
        {model.casefold() for model in selected_models}
        if selected_models is not None
        else None
    )
    displayed_summaries: list[_ModelSummary] = []
    displayed_source_models: set[str] = set()
    for summary in summaries:
        if selected_keys is not None and summary.key.casefold() not in selected_keys:
            continue
        displayed_summaries.append(summary)
        displayed_source_models.add(summary.key.casefold())
        suggestions = suggestions_by_source_model.get(summary.key, ())
        top_entries = tuple(top_five_by_model.get(summary.key, ())[:GATING_TOP_LIMIT])
        gating = _gating_rich_text(
            summary,
            top_entries,
            bom_matches_by_source_row,
            item_data,
        )
        worksheet.append(
            _summary_values(
                summary.model,
                summary.on_hand_qty,
                summary.opo_qty,
                summary.on_hand_amount,
                summary.opo_amount,
                summary.total_amount,
                gating,
                suggestions,
            )
        )

    if selected_models is not None:
        for model in selected_models:
            if model.casefold() in displayed_source_models:
                continue
            worksheet.append(
                [
                    None,
                    None,
                    None,
                    None,
                    None,
                    None,
                    f"{model}：找不到該來源 MODEL 的 Excess 摘要資料。",
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
            )
    if worksheet.max_row == 3:
        worksheet.append(
            [
                None,
                None,
                None,
                None,
                None,
                None,
                "目前沒有符合條件的資料。",
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
        )

    total_row = worksheet.max_row + 1
    worksheet.append(
        [
            "總計",
            sum(summary.on_hand_qty for summary in displayed_summaries),
            sum(summary.opo_qty for summary in displayed_summaries),
            sum(summary.on_hand_amount for summary in displayed_summaries),
            sum(summary.opo_amount for summary in displayed_summaries),
            sum(summary.total_amount for summary in displayed_summaries),
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
    )
    _format_summary_sheet(worksheet, total_row)


def _gating_rich_text(
    summary: _ModelSummary,
    top_entries: tuple[_TopFiveEntry, ...],
    bom_matches_by_source_row: dict[int, _BomMatch],
    item_data: ItemWorkbookData | None,
) -> CellRichText | None:
    if not top_entries:
        return None

    rich_text = CellRichText()
    selected_matches: list[tuple[_TopFiveEntry, _BomMatch]] = []
    for index, entry in enumerate(top_entries):
        if index:
            _append_rich_text(rich_text, "\n")
        part_text = entry.source.parts[0] if entry.source.parts else ""
        match = bom_matches_by_source_row.get(entry.source_row)
        item_type = (
            item_data.item_type_for(match.matched_part)
            if match is not None and item_data is not None
            else None
        )
        type_suffix = f"（TYPE: {item_type}）" if item_type else ""
        _append_rich_text(
            rich_text,
            f"Top {entry.rank}：{part_text} USD {entry.stock_amount:,}{type_suffix}",
        )
        _append_rich_text(rich_text, "\n")
        if match is None:
            _append_rich_text(rich_text, "此料號未找到可辨識的 BOM MODEL")
            continue
        if match.model_key == summary.key.casefold():
            _append_rich_text(
                rich_text,
                f"{SELECTED_MODEL_MARKER}：{summary.model}",
            )
            selected_matches.append((entry, match))
        else:
            _append_rich_text(
                rich_text,
                f"此料號對應 BOM MODEL：{match.target_model}",
            )

    if selected_matches:
        _append_rich_text(rich_text, "\n\n建議邏輯")
        for index, (entry, match) in enumerate(selected_matches):
            _append_rich_text(rich_text, "\n\n" if index else "\n")
            part_text = entry.source.parts[0] if entry.source.parts else ""
            _append_rich_text(rich_text, f"Top {entry.rank} 料號：{part_text}\n")
            _append_rich_text(
                rich_text,
                f"對應 BOM MODEL：{match.target_model}\n",
            )
            _append_rich_text(
                rich_text,
                f"對應 TOP_ASSEMBLY_ITEM：{match.calculation_top_assembly}\n",
            )
            _append_rich_text(
                rich_text,
                f"對應 BOM 列：{_format_bom_rows(match.bom_rows)}\n",
            )
            raw_quantity = entry.stock_qty / match.usage_qty
            _append_rich_text(
                rich_text,
                f"EXTENDED_QT 使用量：{_format_bom_usage(match.bom_rows, match.usage_qty)}\n",
            )
            _append_rich_text(
                rich_text,
                "計算：Excess stock QTY "
                f"{_format_number(entry.stock_qty)} ÷ EXTENDED_QT "
                f"{_format_number(match.usage_qty)} = {_format_number(raw_quantity)} 台，"
                f"向上取整為 {ceil(raw_quantity)} 台",
            )
    return rich_text


def _append_rich_text(
    rich_text: CellRichText,
    text: str,
    *,
    red: bool = False,
) -> None:
    # Keep line breaks inside a real text run.  A newline-only TextBlock is
    # serialized by openpyxl as <t>\n</t> without xml:space="preserve".
    # Excel may then discard that run when rendering rich text, which makes
    # the Gating cell appear as one continuous paragraph.  Appending the
    # whitespace to the preceding run makes openpyxl emit
    # xml:space="preserve" while retaining the formatting of the preceding
    # text.  The newline itself is not visible, so inheriting that color does
    # not affect the displayed result.
    if text and not text.strip() and rich_text:
        previous = rich_text[-1]
        if isinstance(previous, TextBlock):
            previous.text += text
            return

    rich_text.append(
        TextBlock(
            InlineFont(
                rFont="Arial",
                sz=9,
                color="FFFF0000" if red else "FF000000",
            ),
            text,
        )
    )


def _format_bom_rows(rows: tuple[_BomRowDetail, ...]) -> str:
    if not rows:
        return "未保留 BOM 列資訊"
    return "、".join(
        f"第{row_number}列（{part}，EXTENDED_QT {_format_number(quantity)} 件/台）"
        for row_number, part, quantity in rows
    )


def _format_bom_usage(
    rows: tuple[_BomRowDetail, ...],
    total_usage: float,
) -> str:
    if len(rows) <= 1:
        return f"{_format_number(total_usage)} 件/台"
    quantities = " + ".join(
        _format_number(quantity)
        for _row_number, _part, quantity in rows
    )
    return f"{quantities} = {_format_number(total_usage)} 件/台"


def _format_number(value: float) -> str:
    rounded = round(value)
    if abs(value - rounded) < 1e-9:
        return f"{int(rounded):,}"
    return f"{value:,.3f}".rstrip("0").rstrip(".")


def _suggestion_order_text(suggestion: _ProductionSuggestion) -> str:
    return f"建議對BOM MODEL {suggestion.target_model} 加單{suggestion.quantity}台"


def _part_key(value: str) -> str:
    return str(value).strip().casefold()


def _display_parts(parts: Any) -> str:
    values = [str(part).strip() for part in parts if str(part).strip()]
    return "／".join(values) or "未提供料號"


def _last_period_metrics(source: ExcessSummarySourceRow) -> ExcessSummaryMetrics:
    if not source.period_metrics:
        raise ValueError("Summary source row has no excess-period metrics")
    return source.period_metrics[-1]


def _model_key(value: str | None) -> str | None:
    if value is None:
        return None
    normalized = str(value).strip()
    return normalized or None


def _detail_headers(top_five: list[_TopFiveEntry]) -> list[str]:
    max_part_count = max(
        (len(entry.source.parts) for entry in top_five),
        default=1,
    )
    part_headers = ["PartNo1"] + [
        f"替代料{index}" for index in range(2, max_part_count + 1)
    ]
    return DETAIL_FIXED_HEADERS + part_headers + DETAIL_TAIL_HEADERS


def _detail_values(entry: _TopFiveEntry, column_count: int) -> list[Any]:
    source = entry.source
    metrics = source.period_metrics[-1]
    part_count = column_count - len(DETAIL_FIXED_HEADERS) - len(DETAIL_TAIL_HEADERS)
    parts: list[Any] = list(source.parts[:part_count])
    parts.extend([None] * (part_count - len(parts)))
    return (
        [
            entry.rank,
            entry.source_row,
            _model_key(source.model) or UNCLASSIFIED_MODEL,
        ]
        + parts
        + [
            source.price,
            entry.stock_qty,
            entry.stock_amount,
            entry.po_qty,
            entry.po_amount,
            entry.total_amount,
        ]
    )


def _format_summary_sheet(worksheet: Any, total_row: int) -> None:
    _format_common_sheet(worksheet, len(SUMMARY_HEADERS), total_row)
    worksheet.freeze_panes = "A4"
    worksheet.auto_filter.ref = f"A3:{get_column_letter(len(SUMMARY_HEADERS))}{total_row - 1}"
    worksheet.row_dimensions[3].height = 44
    worksheet.column_dimensions["A"].width = 32
    worksheet.column_dimensions["B"].width = 12
    worksheet.column_dimensions["C"].width = 12
    worksheet.column_dimensions["D"].width = 19
    worksheet.column_dimensions["E"].width = 17
    worksheet.column_dimensions["F"].width = 14
    worksheet.column_dimensions["G"].width = 64
    worksheet.column_dimensions["H"].width = 30
    worksheet.column_dimensions["I"].width = 14
    worksheet.column_dimensions["J"].width = 22
    worksheet.column_dimensions["K"].width = 26
    worksheet.column_dimensions["L"].width = 24
    worksheet.column_dimensions["M"].width = 22
    worksheet.column_dimensions["N"].width = 20
    worksheet.column_dimensions["O"].width = 36
    worksheet.column_dimensions["P"].width = 23

    for row in range(4, total_row):
        line_count = max(
            _wrapped_line_count(
                worksheet.cell(row, column).value,
                worksheet.column_dimensions[get_column_letter(column)].width or 18,
            )
            for column in range(7, 17)
        )
        worksheet.row_dimensions[row].height = max(
            30,
            min(409.5, line_count * 18 + 8),
        )
    for row in range(4, total_row + 1):
        for column in range(2, 7):
            worksheet.cell(row, column).number_format = "#,##0;[Red]-#,##0;0"
    for cell in worksheet[total_row]:
        cell.font = Font(name="Arial", size=9, bold=True, color="FF0070C0")


def _format_consumption_sheet(worksheet: Any, last_row: int) -> None:
    _format_common_sheet(worksheet, len(CONSUMPTION_HEADERS), last_row)
    worksheet.freeze_panes = "A4"
    worksheet.auto_filter.ref = (
        f"A3:{get_column_letter(len(CONSUMPTION_HEADERS))}{max(last_row, 3)}"
    )
    worksheet.row_dimensions[3].height = 44
    widths = [
        12, 24, 24, 28, 10, 28, 14, 20, 26, 23, 21, 12, 16, 34, 42
    ]
    for column, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(column)].width = width
    for row in range(4, last_row + 1):
        worksheet.row_dimensions[row].height = 54
        for column in (7,):
            worksheet.cell(row, column).number_format = "#,##0.###;[Red]-#,##0.###;0"
        for column in (8, 9, 10, 11):
            worksheet.cell(row, column).number_format = "#,##0.00;[Red]-#,##0.00;0"
        worksheet.cell(row, 12).number_format = "0.00%"
        if str(worksheet.cell(row, 13).value).startswith("不足"):
            worksheet.cell(row, 13).font = Font(name="Arial", size=9, color="FFFF0000")


def _format_supply_detail_sheet(worksheet: Any, last_row: int) -> None:
    _format_common_sheet(worksheet, len(SUPPLY_DETAIL_HEADERS), last_row)
    worksheet.freeze_panes = "A4"
    worksheet.auto_filter.ref = (
        f"A3:{get_column_letter(len(SUPPLY_DETAIL_HEADERS))}{max(last_row, 3)}"
    )
    worksheet.row_dimensions[3].height = 44
    widths = [
        12, 24, 28, 30, 48, 12, 14, 18, 20, 14, 14, 16, 16, 15, 20, 20, 38
    ]
    for column, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(column)].width = width
    for row in range(4, last_row + 1):
        worksheet.row_dimensions[row].height = 42
        for column in (6, 7, 8, 10, 11, 12, 13):
            worksheet.cell(row, column).number_format = "#,##0.###;[Red]-#,##0.###;0"
        worksheet.cell(row, 9).number_format = "#,##0.00;[Red]-#,##0.00;0"
        worksheet.cell(row, 14).number_format = "#,##0.0000;[Red]-#,##0.0000;0"
        for column in (15, 16):
            worksheet.cell(row, column).number_format = "#,##0.00;[Red]-#,##0.00;0"
        if str(worksheet.cell(row, 17).value).startswith("不足"):
            worksheet.cell(row, 17).font = Font(name="Arial", size=9, color="FFFF0000")


def _format_detail_sheet(worksheet: Any, column_count: int, last_row: int) -> None:
    _format_common_sheet(worksheet, column_count, last_row)
    worksheet.freeze_panes = "A4"
    worksheet.auto_filter.ref = f"A3:{get_column_letter(column_count)}{max(last_row, 3)}"
    worksheet.row_dimensions[3].height = 44
    no_fill = PatternFill(fill_type=None)
    for column in range(1, column_count + 1):
        worksheet.cell(3, column).fill = no_fill

    previous_model: Any = object()
    model_group = -1
    group_fill = PatternFill("solid", fgColor="F2F2F2")
    for row in range(4, last_row + 1):
        model = worksheet.cell(row, 3).value
        if model != previous_model:
            model_group += 1
            previous_model = model
        fill = group_fill if model_group % 2 == 0 else no_fill
        for column in range(1, column_count + 1):
            worksheet.cell(row, column).fill = fill

    for column in range(1, column_count + 1):
        header = worksheet.cell(3, column).value
        if header and (header.startswith("PartNo") or header.startswith("替代料")):
            worksheet.column_dimensions[get_column_letter(column)].width = 18
        elif header in {"模型", "MODEL"}:
            worksheet.column_dimensions[get_column_letter(column)].width = 26
        elif header in {"來源列", "排名"}:
            worksheet.column_dimensions[get_column_letter(column)].width = 11
        else:
            worksheet.column_dimensions[get_column_letter(column)].width = 18
    for row in range(4, last_row + 1):
        worksheet.row_dimensions[row].height = 36
        for column in range(1, column_count + 1):
            header = worksheet.cell(3, column).value
            if header in {
                "Excess stock Amount USD",
                "Excess POAmount USD",
                "Excess TotalAMT USD",
            }:
                worksheet.cell(row, column).number_format = "#,##0.00;[Red]-#,##0.00;0"
            elif header == "Price（USD)":
                worksheet.cell(row, column).number_format = "#,##0.0000;[Red]-#,##0.0000;0"
            elif header in {
                "排名",
                "來源列",
                "Excess stock QTY",
                "Excess POQTY",
            }:
                worksheet.cell(row, column).number_format = "#,##0;[Red]-#,##0;0"


def _format_common_sheet(worksheet: Any, column_count: int, last_row: int) -> None:
    thin_gray = Side(style="thin", color="BFBFBF")
    border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    header_fills = [
        "D9EAF7",
        "D9EAF7",
        "D9EAF7",
        "FCE4D6",
        "FCE4D6",
        "FFF2CC",
        "E7E6E6",
        "D9EAF7",
        "FFF2CC",
    ]
    for row in worksheet.iter_rows(min_row=1, max_row=max(last_row, 3), max_col=column_count):
        for cell in row:
            cell.font = Font(name="Arial", size=9)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    worksheet["A1"].font = Font(name="Arial", size=10, bold=True, color="FF704016")
    for column in range(1, column_count + 1):
        cell = worksheet.cell(3, column)
        cell.fill = PatternFill("solid", fgColor=header_fills[(column - 1) % len(header_fills)])
        cell.font = Font(name="Arial", size=9, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in range(4, last_row + 1):
        for column in range(1, column_count + 1):
            cell = worksheet.cell(row, column)
            if column == 1 or (
                worksheet.title in {"AVTC Summary", "Raken summary", "加單建議"}
                and column in {7, 8, 9, 10, 11, 12, 13, 14, 15, 16}
            ):
                cell.alignment = Alignment(
                    horizontal="left",
                    vertical="center",
                    wrap_text=True,
                )
            else:
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )


def _wrapped_line_count(value: Any, column_width: float) -> int:
    if value in (None, ""):
        return 1
    usable_width = max(float(column_width) - 2, 1.0)
    return sum(
        max(1, ceil(_display_width(line) / usable_width))
        for line in str(value).split("\n")
    )


def _display_width(text: str) -> int:
    return sum(
        2 if unicodedata.east_asian_width(character) in {"W", "F"} else 1
        for character in text
    )
