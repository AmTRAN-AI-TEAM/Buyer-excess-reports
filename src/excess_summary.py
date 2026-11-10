from __future__ import annotations

from dataclasses import dataclass
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


@dataclass(frozen=True)
class _ProductionSuggestion:
    target_model: str
    calculation_top_assembly: str
    quantity: int
    consumed_amount: float
    consumed_items: tuple[tuple[int, str, str, float, float], ...]
    bom_rows: tuple[_BomRowDetail, ...] = ()


@dataclass
class _InventoryPool:
    parts: frozenset[str]
    quantity: float


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
) -> tuple[int, int]:
    """Write a model summary and each model's five highest stock amounts.

    Values are calculated before the Excess workbook's formulas are written, so
    this workbook contains concrete values and does not depend on Excel's formula
    cache being populated.
    """
    output_file.parent.mkdir(parents=True, exist_ok=True)
    summary_by_model: dict[str, _ModelSummary] = {}
    bom_matches_by_source_row = _bom_matches_by_source_model(
        source_rows,
        component_usages,
    )
    top_five_by_model = _top_five_entries(
        source_rows,
        eligible_source_rows=set(bom_matches_by_source_row),
    )

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
        summary_headers[8] = "Raken 呆滯 预计TO-BE"
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
        bom_matches_by_source_row,
    )
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
            bom_matches_by_source_row,
        )

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
    _write_recommendations_sheet(
        recommendations_ws,
        summaries,
        suggestions_by_model,
        period_label,
        selected_models,
        summary_headers,
        top_five_by_model,
        bom_matches_by_source_row,
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
    bom_matches_by_source_row: dict[int, _BomMatch],
) -> tuple[
    dict[str, tuple[_TopFiveEntry, ...]],
    dict[str, tuple[_ProductionSuggestion, ...]],
]:
    """Select up to two model-matched entries per model and calculate each independently."""
    if not bom_matches_by_source_row:
        return {}, {}

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

    assignments = {
        entry.source_row: bom_matches_by_source_row[entry.source_row]
        for entries in top_five_by_model.values()
        for entry in entries
        if entry.source_row in bom_matches_by_source_row
    }

    recommendable_entries_by_model: dict[str, tuple[_TopFiveEntry, ...]] = {}
    selected_pool_keys: set[int] = set()
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
            selected_pool_keys.update(entry.source_row for entry in recommendable_entries)
    selected_assignments = {
        pool_key: assignment
        for pool_key, assignment in assignments.items()
        if pool_key in selected_pool_keys
    }

    suggestions_by_model: dict[str, list[_ProductionSuggestion]] = {}
    for pool_key, match in selected_assignments.items():
        pool = pools[pool_key]
        if pool.quantity <= 0 or match.usage_qty <= 0:
            continue
        entry = entries_by_pool[pool_key]
        owner = owner_by_pool[pool_key]
        consumed_item = (
            entry.rank,
            "／".join(entry.source.parts),
            part_names_by_pool[pool_key].get(match.matched_part, match.matched_part),
            pool.quantity,
            match.usage_qty,
        )
        suggestions_by_model.setdefault(owner, []).append(
            _ProductionSuggestion(
                target_model=match.target_model,
                calculation_top_assembly=match.calculation_top_assembly,
                quantity=ceil(pool.quantity / match.usage_qty),
                consumed_amount=entry.stock_amount,
                consumed_items=(consumed_item,),
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
                        -suggestion.consumed_amount,
                        suggestion.quantity,
                        suggestion.target_model.casefold(),
                        suggestion.calculation_top_assembly.casefold(),
                    ),
                )
            )
            for model, suggestions in suggestions_by_model.items()
        },
    )


def _write_recommendations_sheet(
    worksheet: Any,
    summaries: list[_ModelSummary],
    suggestions_by_source_model: dict[str, tuple[_ProductionSuggestion, ...]],
    period_label: str,
    selected_models: tuple[str, ...] | None,
    headers: list[str],
    top_five_by_model: dict[str, list[_TopFiveEntry]],
    bom_matches_by_source_row: dict[int, _BomMatch],
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
        )
        order_text = "\n".join(
            _suggestion_order_text(suggestion)
            for suggestion in suggestions
        ) or None
        worksheet.append(
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
                ]
            )
    if worksheet.max_row == 3:
        worksheet.append(
            [None, None, None, None, None, None, "目前沒有符合條件的資料。", None, None]
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
        ]
    )
    _format_summary_sheet(worksheet, total_row)


def _gating_rich_text(
    summary: _ModelSummary,
    top_entries: tuple[_TopFiveEntry, ...],
    bom_matches_by_source_row: dict[int, _BomMatch],
) -> CellRichText | None:
    if not top_entries:
        return None

    rich_text = CellRichText()
    selected_matches: list[tuple[_TopFiveEntry, _BomMatch]] = []
    for index, entry in enumerate(top_entries):
        if index:
            _append_rich_text(rich_text, "\n")
        part_text = entry.source.parts[0] if entry.source.parts else ""
        _append_rich_text(
            rich_text,
            f"Top {entry.rank}：{part_text} USD {entry.stock_amount:,}",
        )
        match = bom_matches_by_source_row.get(entry.source_row)
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
            _append_rich_text(
                rich_text,
                f"EXTENDED_QT 使用量：{_format_bom_usage(match.bom_rows, match.usage_qty)}\n",
            )
            raw_quantity = entry.stock_qty / match.usage_qty
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
    usage_qty: float,
) -> str:
    if not rows:
        return f"合計 {_format_number(usage_qty)} 件/台"
    quantities = " + ".join(
        _format_number(quantity) for _row_number, _part, quantity in rows
    )
    return f"{quantities} = {_format_number(usage_qty)} 件/台"


def _format_number(value: float) -> str:
    rounded = round(value)
    if abs(value - rounded) < 1e-9:
        return f"{int(rounded):,}"
    return f"{value:,.3f}".rstrip("0").rstrip(".")


def _suggestion_order_text(suggestion: _ProductionSuggestion) -> str:
    return f"建議對BOM MODEL {suggestion.target_model}\n加單{suggestion.quantity}台"


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
    worksheet.column_dimensions["H"].width = 30
    worksheet.column_dimensions["I"].width = 23

    for row in range(4, total_row):
        gating_text = worksheet.cell(row, 7).value
        order_text = worksheet.cell(row, 8).value
        gating_lines = _wrapped_line_count(
            gating_text,
            worksheet.column_dimensions["G"].width or 64,
        )
        order_lines = _wrapped_line_count(
            order_text,
            worksheet.column_dimensions["H"].width or 25,
        )
        line_count = max(gating_lines, order_lines)
        worksheet.row_dimensions[row].height = max(
            30,
            min(409.5, line_count * 18 + 8),
        )
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
            if column == 1 or (
                worksheet.title in {"AVTC Summary", "Raken summary", "加單建議"}
                and column == 7
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
