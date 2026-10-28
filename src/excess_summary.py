from __future__ import annotations

from dataclasses import dataclass, field
from math import ceil
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from bom_alternates import BomComponentUsage
from excess_builder import ExcessSummaryMetrics, ExcessSummarySourceRow


SUMMARY_HEADERS = [
    "AVTC",
    "ON HAND",
    "OPO",
    "ON HAND AMT USD",
    "OPO AMT\nUSD",
    "TTL AMT",
    "Gating",
    "AVTC建议加单",
    "AVTC 呆滞 预计TO-BE",
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


@dataclass
class _ModelSummary:
    key: str
    model: str
    on_hand_qty: int = 0
    opo_qty: int = 0
    on_hand_amount: int = 0
    opo_amount: int = 0
    total_amount: int = 0
    gating_notes: list[str] = field(default_factory=list)


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
class _ProductionSuggestion:
    target_finished_good: str
    quantity: int
    consumed_amount: float
    consumed_items: tuple[tuple[int, str, str, float, float], ...]
    shared_order: bool = False


@dataclass
class _InventoryPool:
    parts: frozenset[str]
    quantity: float


def write_excess_summary(
    source_rows: tuple[ExcessSummarySourceRow, ...],
    period_label: str,
    output_file: Path,
    *,
    component_usages: tuple[BomComponentUsage, ...] = (),
) -> tuple[int, int]:
    """Write a model summary and each model's five highest stock amounts.

    Values are calculated before the Excess workbook's formulas are written, so
    this workbook contains concrete values and does not depend on Excel's formula
    cache being populated.
    """
    output_file.parent.mkdir(parents=True, exist_ok=True)
    summary_by_model: dict[str, _ModelSummary] = {}
    top_five_by_model = _top_five_entries(source_rows)

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
    summary_ws = workbook.active
    summary_ws.title = "AVTC Summary"
    summary_ws["A1"] = f"選取區間：{period_label}"
    for column, header in enumerate(SUMMARY_HEADERS, start=1):
        summary_ws.cell(3, column, header)

    summaries = sorted(
        summary_by_model.values(),
        key=lambda item: (-item.on_hand_amount, item.model.casefold()),
    )
    summary_top_by_model = {
        model: entries[:SUMMARY_TOP_LIMIT]
        for model, entries in top_five_by_model.items()
    }
    suggestions_by_model = _recommend_production_orders(
        summary_top_by_model,
        component_usages,
    )
    top_five = [
        entry
        for summary in summaries
        for entry in top_five_by_model.get(summary.key, ())
    ]
    for summary in summaries:
        suggestions = suggestions_by_model.get(summary.key, ())
        summary_top_entries = top_five_by_model.get(summary.key, ())[:SUMMARY_TOP_LIMIT]
        for entry in summary_top_entries:
            part_no = entry.source.parts[0] if entry.source.parts else ""
            summary.gating_notes.append(
                f"Top {entry.rank}：{part_no} USD {entry.stock_amount:,}"
            )
        gating = ""
        if suggestions:
            summary.gating_notes.append("生產加單建議：")
            summary.gating_notes.extend(
                _suggestion_text(suggestion) for suggestion in suggestions
            )
        elif summary.gating_notes:
            if not _has_usable_component_usage(component_usages):
                summary.gating_notes.append(
                    "無法計算加單：BOM 缺少可辨識的 TOP_ASSEMBLY_ITEM 或正數 EXTENDED_QT 用量。"
                )
            else:
                summary.gating_notes.append(
                    "暫不建議加單：前兩名庫存沒有找到完整相符的 BOM 成品／替代料用料群及正數 EXTENDED_QT，或替代料用量不一致而無法可靠換算。"
                )
        if summary.gating_notes:
            gating = "\n".join(summary.gating_notes)

        order_text = "\n".join(
            _suggestion_order_text(suggestion)
            for suggestion in suggestions
        ) or None
        summary_ws.append(
            [
                summary.model,
                summary.on_hand_qty,
                summary.opo_qty,
                summary.on_hand_amount,
                summary.opo_amount,
                summary.total_amount,
                gating,
                order_text,
                None,
            ]
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
        ]
    )
    _format_summary_sheet(summary_ws, total_row)

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


def _top_five_entries(
    source_rows: tuple[ExcessSummarySourceRow, ...],
) -> dict[str, list[_TopFiveEntry]]:
    candidates_by_model: dict[str, list[tuple[int, int, ExcessSummarySourceRow]]] = {}
    for index, source in enumerate(source_rows):
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


def _recommend_production_orders(
    top_five_by_model: dict[str, list[_TopFiveEntry]],
    component_usages: tuple[BomComponentUsage, ...],
) -> dict[str, tuple[_ProductionSuggestion, ...]]:
    """Assign each Top 5 inventory pool once to its best BOM consumption option."""
    if not _has_usable_component_usage(component_usages):
        return {}

    pools: dict[int, _InventoryPool] = {}
    owner_by_pool: dict[int, str] = {}
    entries_by_pool: dict[int, _TopFiveEntry] = {}
    part_names_by_pool: dict[int, dict[str, str]] = {}
    for source_model, entries in top_five_by_model.items():
        for entry in entries:
            key = entry.source_row
            quantity = max(entry.stock_qty, 0)
            part_names = {
                _part_key(part): str(part).strip()
                for part in entry.source.parts
                if _part_key(part)
            }
            parts = frozenset(part_names)
            if quantity <= 0 or not parts:
                continue
            pools[key] = _InventoryPool(parts, float(quantity))
            owner_by_pool[key] = source_model
            entries_by_pool[key] = entry
            part_names_by_pool[key] = part_names

    usage_slots: dict[
        tuple[str, frozenset[str]],
        dict[str, Any],
    ] = {}
    finished_good_names: dict[str, str] = {}
    for usage in component_usages:
        finished_good_name = (usage.finished_good or "").strip()
        finished_good_key = finished_good_name.casefold()
        parts = frozenset(_part_key(part) for part in usage.parts if _part_key(part))
        if not finished_good_key or not parts:
            continue
        finished_good_names.setdefault(finished_good_key, finished_good_name)
        slot_key = (finished_good_key, parts)
        slot = usage_slots.setdefault(
            slot_key,
            {"part_usage": {}, "row_number": usage.row_number},
        )
        slot["row_number"] = min(slot["row_number"], usage.row_number)
        part_usage_qtys = usage.part_usage_qtys or tuple(
            (part, usage.usage_qty) for part in usage.parts
        )
        for part, quantity in part_usage_qtys:
            part_key = _part_key(part)
            if part_key not in parts or quantity is None or quantity <= 0:
                continue
            part_usage = slot["part_usage"]
            part_usage[part_key] = part_usage.get(part_key, 0.0) + quantity

    assignments: dict[int, tuple[str, frozenset[str], str, float]] = {}
    for pool_key, pool in pools.items():
        primary_part = next(iter(part_names_by_pool[pool_key]), "")
        candidates: list[
            tuple[float, int, str, int, str, tuple[str, ...], frozenset[str]]
        ] = []
        for (finished_good_key, slot_parts), slot in usage_slots.items():
            part_usage = slot["part_usage"]
            if not pool.parts.issubset(part_usage):
                continue
            matched_parts = pool.parts
            usage_values = {part_usage[part] for part in matched_parts}
            if len(usage_values) > 1:
                continue
            for matched_part in matched_parts:
                quantity = part_usage[matched_part]
                candidates.append(
                    (
                        -quantity,
                        slot["row_number"],
                        finished_good_key,
                        0 if matched_part == primary_part else 1,
                        matched_part,
                        tuple(sorted(slot_parts)),
                        slot_parts,
                    )
                )
        if not candidates:
            continue
        (
            negative_usage,
            _row_number,
            finished_good_key,
            _primary_order,
            matched_part,
            _slot_parts_order,
            slot_parts,
        ) = min(candidates)
        assignments[pool_key] = (
            finished_good_key,
            slot_parts,
            matched_part,
            -negative_usage,
        )

    assigned_by_product_slot_part: dict[
        tuple[str, frozenset[str], str], list[int]
    ] = {}
    for pool_key, (finished_good_key, slot_parts, matched_part, _usage_qty) in assignments.items():
        assigned_by_product_slot_part.setdefault(
            (finished_good_key, slot_parts, matched_part), []
        ).append(pool_key)

    product_order_qty: dict[str, int] = {}
    for finished_good_key in finished_good_names:
        slot_order_qtys: list[int] = []
        for (slot_finished_good, slot_parts), slot in usage_slots.items():
            if slot_finished_good != finished_good_key:
                continue
            alternative_order_qty = 0
            part_usage = slot["part_usage"]
            for part in part_usage:
                pool_keys = assigned_by_product_slot_part.get(
                    (finished_good_key, slot_parts, part), ()
                )
                if not pool_keys:
                    continue
                stock_qty = sum(pools[key].quantity for key in pool_keys)
                usage_qty = part_usage[part]
                if stock_qty > 0 and usage_qty > 0:
                    alternative_order_qty += ceil(stock_qty / usage_qty)
            if alternative_order_qty:
                slot_order_qtys.append(alternative_order_qty)
        if slot_order_qtys:
            product_order_qty[finished_good_key] = max(slot_order_qtys)

    pools_by_product: dict[str, list[int]] = {}
    for pool_key, (finished_good_key, _slot_parts, _part, _usage_qty) in assignments.items():
        pools_by_product.setdefault(finished_good_key, []).append(pool_key)

    suggestions_by_model: dict[str, list[_ProductionSuggestion]] = {}
    for finished_good_key in sorted(pools_by_product):
        order_qty = product_order_qty.get(finished_good_key, 0)
        if order_qty <= 0:
            continue
        pool_keys = pools_by_product[finished_good_key]
        owners = {owner_by_pool[key] for key in pool_keys}
        shared_order = len(owners) > 1
        target_name = finished_good_names.get(finished_good_key, finished_good_key)
        for owner in sorted(owners):
            owner_pool_keys = [key for key in pool_keys if owner_by_pool[key] == owner]
            consumed_amount = sum(
                entries_by_pool[key].stock_amount for key in owner_pool_keys
            )
            consumed_items = tuple(
                sorted(
                    (
                        entries_by_pool[key].rank,
                        "／".join(entries_by_pool[key].source.parts),
                        part_names_by_pool[key].get(matched_part, matched_part),
                        pools[key].quantity,
                        usage_qty,
                    )
                    for key in owner_pool_keys
                    for _finished_good, _slot_parts, matched_part, usage_qty in [
                        assignments[key]
                    ]
                )
            )
            suggestions_by_model.setdefault(owner, []).append(
                _ProductionSuggestion(
                    target_finished_good=target_name,
                    quantity=order_qty,
                    consumed_amount=consumed_amount,
                    consumed_items=consumed_items,
                    shared_order=shared_order,
                )
            )

    return {
        model: tuple(
            sorted(
                suggestions,
                key=lambda suggestion: (
                    -suggestion.consumed_amount,
                    suggestion.quantity,
                    suggestion.target_finished_good.casefold(),
                ),
            )
        )
        for model, suggestions in suggestions_by_model.items()
    }


def _suggestion_text(suggestion: _ProductionSuggestion) -> str:
    used_items = "、".join(
        f"Top {rank} {part_label}"
        + (f"（對應 BOM {matched_part}）" if part_label != matched_part else "")
        + f"庫存 {_format_quantity(quantity)} 件 ÷ BOM EXTENDED_QT {_format_quantity(usage_qty)} 件/台"
        for rank, part_label, matched_part, quantity, usage_qty in suggestion.consumed_items
    )
    shared_order_note = (
        "（與其他 MODEL 共用同一筆加單，台數勿重複加總）"
        if suggestion.shared_order
        else ""
    )
    return (
        f"建議對成品料號 {suggestion.target_finished_good} 加排 {suggestion.quantity:,} 台；"
        f"依現有 Excess 庫存估算可消耗 USD {int(round(suggestion.consumed_amount)):,}"
        f"（{used_items}）。{shared_order_note}"
    )


def _suggestion_order_text(suggestion: _ProductionSuggestion) -> str:
    ranks = sorted({rank for rank, *_rest in suggestion.consumed_items})
    rank_text = "、".join(f"Top {rank}" for rank in ranks)
    shared_order_note = (
        "（共用同一筆加單，台數勿重複加總）" if suggestion.shared_order else ""
    )
    return (
        f"{rank_text}：{suggestion.target_finished_good} × {suggestion.quantity:,} 台"
        f"{shared_order_note}"
    )


def _format_quantity(quantity: float) -> str:
    rounded = round(quantity)
    if abs(quantity - rounded) < 1e-9:
        return f"{int(rounded):,}"
    return f"{quantity:,.3f}".rstrip("0").rstrip(".")


def _has_usable_component_usage(
    component_usages: tuple[BomComponentUsage, ...],
) -> bool:
    return any(
        usage.finished_good
        and usage.parts
        and any(
            quantity is not None and quantity > 0
            for _part, quantity in (
                usage.part_usage_qtys
                or tuple((part, usage.usage_qty) for part in usage.parts)
            )
        )
        for usage in component_usages
    )


def _part_key(value: str) -> str:
    return str(value).strip().casefold()


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
    worksheet.auto_filter.ref = f"A3:I{total_row - 1}"
    worksheet.row_dimensions[3].height = 44
    worksheet.column_dimensions["A"].width = 32
    worksheet.column_dimensions["B"].width = 12
    worksheet.column_dimensions["C"].width = 12
    worksheet.column_dimensions["D"].width = 19
    worksheet.column_dimensions["E"].width = 17
    worksheet.column_dimensions["F"].width = 14
    worksheet.column_dimensions["G"].width = 64
    worksheet.column_dimensions["H"].width = 25
    worksheet.column_dimensions["I"].width = 23

    for row in range(4, total_row):
        gating_text = worksheet.cell(row, 7).value or ""
        line_count = str(gating_text).count("\n") + 1 if gating_text else 1
        worksheet.row_dimensions[row].height = max(24, min(160, line_count * 20))
    for row in range(4, total_row + 1):
        for column in range(2, 7):
            worksheet.cell(row, column).number_format = "#,##0;[Red]-#,##0;0"
    for cell in worksheet[total_row]:
        cell.font = Font(name="Arial", size=9, bold=True, color="FF0070C0")


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
            if column == 1 or (worksheet.title == "AVTC Summary" and column == 7):
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
