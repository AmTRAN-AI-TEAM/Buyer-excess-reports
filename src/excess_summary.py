from __future__ import annotations

from dataclasses import dataclass, replace
from math import ceil
from pathlib import Path
import unicodedata
from typing import Any, Callable

from openpyxl import Workbook, load_workbook
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
    "庫存數量是否可滿足加單",
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
    "庫存數量是否可滿足加單",
    "需額外採購摘要",
    "計算說明",
]
SUPPLY_DETAIL_HEADERS = [
    "建議編號",
    "BOM MODEL",
    "TOP_ASSEMBLY_ITEM",
    "必要料號／替代料群",
    "供料角色",
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
SHARED_MODEL = "多機種共用料"
UNMATCHED_BOM_MODEL = "BOM 未配對"
TOP_FIVE_LIMIT = 5
SUMMARY_TOP_LIMIT = 2
GATING_TOP_LIMIT = 3
SELECTED_RECOMMENDATIONS_SHEET = "加單建議（選）"
FULL_RECOMMENDATIONS_SHEET = "加單建議（全）"
SUMMARY_PROGRESS_LABELS = (
    "Read Excess report",
    "Build BOM matches",
    "Filter summary item types",
    "Aggregate model summary",
    "Build stock-amount Top 5",
    "Build recommendations",
    "Write summary sheets",
    "Write detail sheets",
    "Save Excess summary",
)
SUMMARY_PROGRESS_STEPS = len(SUMMARY_PROGRESS_LABELS)
_BomRowDetail = tuple[int, str, float]


@dataclass
class _ModelSummary:
    key: str
    model: str
    customer: str | None = None
    on_hand_qty: int = 0
    opo_qty: int = 0
    on_hand_amount: int = 0
    opo_amount: int = 0
    total_amount: int = 0


@dataclass(frozen=True)
class _TopFiveEntry:
    rank: int
    source_row: int
    model: str
    model_key: str
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
    slot_key: tuple[str, str, str, frozenset[str]] | None
    match_kind: str = "component"
    assembly_key: tuple[str, str, str] | None = None
    assembly_row: int | None = None


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
    role: str = "一般物料"


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
    assembly_item_key: str
    assembly_item_name: str
    assembly_key: tuple[str, str, str]
    parts: frozenset[str]
    part_usage: dict[str, float]
    bom_rows: tuple[_BomRowDetail, ...]
    row_number: int


def write_excess_summary(
    source_rows: tuple[ExcessSummarySourceRow, ...],
    period_label: str,
    output_file: Path,
    *,
    excess_report_file: Path | None = None,
    component_usages: tuple[BomComponentUsage, ...] = (),
    customer: str | None = None,
    selected_models: tuple[str, ...] | None = None,
    item_data: ItemWorkbookData | None = None,
    progress_step: Callable[[str], None] | None = None,
) -> tuple[int, int]:
    """Write a model summary and each model's five highest stock amounts.

    Values are calculated before the Excess workbook's formulas are written, so
    this workbook contains concrete values and does not depend on Excel's formula
    cache being populated.
    """
    if progress_step is not None:
        progress_step(SUMMARY_PROGRESS_LABELS[0])
    if excess_report_file is not None:
        source_rows = _source_rows_from_excess_report(source_rows, excess_report_file)

    if progress_step is not None:
        progress_step(SUMMARY_PROGRESS_LABELS[1])
    output_file.parent.mkdir(parents=True, exist_ok=True)
    summary_by_model: dict[str, _ModelSummary] = {}
    usage_slots = _bom_usage_slots(component_usages)
    summary_models_by_source_row = _summary_models_by_source_row(
        source_rows,
    )
    bom_models_by_source_row = _bom_models_by_source_row(source_rows, usage_slots)
    all_bom_matches_by_source_row = _bom_matches_by_source_model(
        source_rows,
        usage_slots,
        bom_models_by_source_row,
    )

    if progress_step is not None:
        progress_step(SUMMARY_PROGRESS_LABELS[2])
    eligible_bom_matches_by_source_row = _filter_bom_matches_by_item_type(
        all_bom_matches_by_source_row,
        item_data,
    )

    if progress_step is not None:
        progress_step(SUMMARY_PROGRESS_LABELS[3])
    for index, source in enumerate(source_rows):
        metrics = _last_period_metrics(source)
        model_display = summary_models_by_source_row[index + 3]
        normalized_key = model_display.casefold()
        summary = summary_by_model.setdefault(
            normalized_key,
            _ModelSummary(normalized_key, model_display),
        )
        summary.customer = _merge_customer_values(summary.customer, source.customer)
        summary.on_hand_qty += max(metrics.stock_qty, 0)
        summary.opo_qty += metrics.po_qty
        summary.on_hand_amount += metrics.stock_amount or 0
        summary.opo_amount += metrics.po_amount or 0
        summary.total_amount += metrics.total_amount

    if progress_step is not None:
        progress_step(SUMMARY_PROGRESS_LABELS[4])
    top_five_by_model = _top_five_entries(
        source_rows,
        summary_models_by_source_row,
        eligible_source_rows=set(eligible_bom_matches_by_source_row),
    )

    if progress_step is not None:
        progress_step(SUMMARY_PROGRESS_LABELS[5])
    workbook = Workbook()
    recommendations_ws = workbook.active
    recommendations_ws.title = SELECTED_RECOMMENDATIONS_SHEET
    summary_sheet_title = FULL_RECOMMENDATIONS_SHEET
    summary_ws = workbook.create_sheet(summary_sheet_title)
    summary_ws["A1"] = f"選取區間：{period_label}"
    summary_headers = list(SUMMARY_HEADERS)
    if customer and customer.upper() == "RAKEN":
        summary_headers[0] = "Raken"
        summary_headers[7] = "Raken建议加单"
        summary_headers[15] = "Raken 呆滯 预计TO-BE"
    summary_customer_column = bool(customer and customer.upper() == "RAKEN")
    summary_sheet_headers = (
        [summary_headers[0], "Customer\n客戶", *summary_headers[1:]]
        if summary_customer_column
        else summary_headers
    )
    recommendation_headers = [
        summary_headers[0],
        "Customer\n客戶",
        *summary_headers[1:],
    ]
    for column, header in enumerate(summary_sheet_headers, start=1):
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
    selected_report_models = _selected_report_models(
        selected_models,
    )
    top_five = [
        entry
        for summary in summaries
        for entry in top_five_by_model.get(summary.key, ())
    ]

    if progress_step is not None:
        progress_step(SUMMARY_PROGRESS_LABELS[6])
    for summary in summaries:
        suggestions = suggestions_by_model.get(summary.key, ())
        summary_top_entries = top_five_by_model.get(summary.key, ())[:GATING_TOP_LIMIT]
        gating = _gating_rich_text(
            summary_top_entries,
            eligible_bom_matches_by_source_row,
            item_data,
        )
        summary_values = _summary_values(
            summary.model,
            summary.on_hand_qty,
            summary.opo_qty,
            summary.on_hand_amount,
            summary.opo_amount,
            summary.total_amount,
            gating,
            suggestions,
        )
        if summary_customer_column:
            summary_values = [summary_values[0], summary.customer, *summary_values[1:]]
        summary_ws.append(summary_values)

    total_row = summary_ws.max_row + 1
    total_values = [
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
    if summary_customer_column:
        total_values = [total_values[0], None, *total_values[1:]]
    summary_ws.append(total_values)
    _format_summary_sheet(
        summary_ws,
        total_row,
        customer_column=bool(summary_customer_column),
    )
    _write_recommendations_sheet(
        recommendations_ws,
        summaries,
        suggestions_by_model,
        period_label,
        selected_report_models,
        recommendation_headers,
        top_five_by_model,
        eligible_bom_matches_by_source_row,
        item_data,
    )

    if progress_step is not None:
        progress_step(SUMMARY_PROGRESS_LABELS[7])
    supply_ws = workbook.create_sheet("供料明細")
    _write_supply_detail_sheet(
        supply_ws,
        summaries,
        suggestions_by_model,
        period_label,
    )

    consumption_ws = workbook.create_sheet("消耗建議")
    _write_consumption_sheet(
        consumption_ws,
        summaries,
        suggestions_by_model,
        period_label,
    )

    detail_ws = workbook.create_sheet("庫存金額Top5")
    detail_ws["A1"] = f"每個 MODEL 依最後一期 Excess stock Amount 取前五名：{period_label}"
    detail_headers = _detail_headers(top_five)
    for column, header in enumerate(detail_headers, start=1):
        detail_ws.cell(3, column, header)
    for entry in top_five:
        detail_ws.append(_detail_values(entry, len(detail_headers)))
    _format_detail_sheet(detail_ws, len(detail_headers), detail_ws.max_row)

    if progress_step is not None:
        progress_step(SUMMARY_PROGRESS_LABELS[8])
    workbook.save(output_file)
    return len(summaries), len(top_five)


def _filter_bom_matches_by_item_type(
    bom_matches_by_source_row: dict[int, _BomMatch],
    item_data: ItemWorkbookData | None,
) -> dict[int, _BomMatch]:
    """Keep eligible BOM matches, exempting matched semi-finished assemblies from TYPE filtering."""
    if item_data is None:
        return {}

    filtered: dict[int, _BomMatch] = {}
    for source_row, match in bom_matches_by_source_row.items():
        # A semi-finished candidate is already proven to be a 3-prefixed
        # L ASSEMBLY_ITEM by _bom_matches_by_source_model().  It is therefore
        # allowed to proceed without an Item TYPE value; normal component
        # candidates continue to use the existing TYPE allow-list.
        if match.match_kind == "semi":
            filtered[source_row] = match
            continue
        item_type = item_data.item_type_for(match.matched_part)
        if is_allowed_summary_item_type(item_type):
            filtered[source_row] = match
    return filtered


def _source_rows_from_excess_report(
    source_rows: tuple[ExcessSummarySourceRow, ...],
    excess_report_file: Path,
) -> tuple[ExcessSummarySourceRow, ...]:
    """Use the fields actually written to the current Excess report."""
    workbook = load_workbook(excess_report_file, read_only=True, data_only=True)
    try:
        worksheet = workbook["Excess"] if "Excess" in workbook.sheetnames else workbook.active
        header_row = None
        headers: tuple[Any, ...] = ()
        for row_number, row in enumerate(
            worksheet.iter_rows(min_row=1, max_row=20, values_only=True),
            start=1,
        ):
            if any(str(value or "").strip() == "MODEL" for value in row):
                header_row = row_number
                headers = tuple(row)
                break
        if header_row is None:
            raise ValueError("Cannot find MODEL column in the generated Excess report")

        model_column = next(
            index
            for index, value in enumerate(headers, start=1)
            if str(value or "").strip() == "MODEL"
        )
        customer_column = next(
            (
                index
                for index, value in enumerate(headers, start=1)
                if str(value or "").strip().casefold().startswith("customer")
            ),
            None,
        )
        report_rows = list(
            worksheet.iter_rows(
                min_row=header_row + 1,
                max_row=header_row + len(source_rows),
                values_only=True,
            )
        )
        if len(report_rows) != len(source_rows):
            raise ValueError(
                "Generated Excess report rows do not match the summary source rows"
            )

        updated_rows: list[ExcessSummarySourceRow] = []
        for source, row in zip(source_rows, report_rows):
            model = row[model_column - 1] if model_column <= len(row) else None
            customer = (
                row[customer_column - 1]
                if customer_column is not None and customer_column <= len(row)
                else source.customer
            )
            updated_rows.append(
                replace(
                    source,
                    model=None if model is None else str(model).strip(),
                    customer=None if customer is None else str(customer).strip(),
                )
            )
        return tuple(updated_rows)
    finally:
        workbook.close()


def _top_five_entries(
    source_rows: tuple[ExcessSummarySourceRow, ...],
    summary_models_by_source_row: dict[int, str],
    *,
    eligible_source_rows: set[int] | None = None,
) -> dict[str, list[_TopFiveEntry]]:
    candidates_by_model: dict[
        str,
        list[tuple[int, int, str, ExcessSummarySourceRow]],
    ] = {}
    for index, source in enumerate(source_rows):
        source_row = index + 3
        if eligible_source_rows is not None and source_row not in eligible_source_rows:
            continue
        if not source.period_metrics:
            continue
        metrics = source.period_metrics[-1]
        if metrics.stock_amount is not None and metrics.stock_amount > 0:
            model = summary_models_by_source_row[source_row]
            model_key = model.casefold()
            candidates_by_model.setdefault(model_key, []).append(
                (metrics.stock_amount, index, model, source)
            )

    entries_by_model: dict[str, list[_TopFiveEntry]] = {}
    for model_key, candidates in candidates_by_model.items():
        candidates.sort(key=lambda item: (-item[0], item[1]))
        entries: list[_TopFiveEntry] = []
        for rank, (stock_amount, index, model, source) in enumerate(
            candidates[:TOP_FIVE_LIMIT],
            start=1,
        ):
            metrics = source.period_metrics[-1]
            entries.append(
                _TopFiveEntry(
                    rank=rank,
                    source_row=index + 3,
                    model=model,
                    model_key=model.casefold(),
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
        tuple[str, str, str, frozenset[str]],
        dict[str, Any],
    ] = {}
    for usage in component_usages:
        model_name = (usage.model or "").strip()
        model_key = model_name.casefold()
        finished_good_name = (usage.finished_good or "").strip()
        finished_good_key = finished_good_name.casefold()
        assembly_item_name = (usage.assembly_item or usage.finished_good or "").strip()
        assembly_item_key = assembly_item_name.casefold()
        parts = frozenset(_part_key(part) for part in usage.parts if _part_key(part))
        if not model_key or not finished_good_key or not assembly_item_key or not parts:
            continue

        slot_key = (model_key, finished_good_key, assembly_item_key, parts)
        slot = raw_slots.setdefault(
            slot_key,
            {
                "model_name": model_name,
                "finished_good_name": finished_good_name,
                "assembly_item_name": assembly_item_name,
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
            assembly_item_key=assembly_item_key,
            assembly_item_name=slot["assembly_item_name"],
            assembly_key=(model_key, finished_good_key, assembly_item_key),
            parts=parts,
            part_usage=slot["part_usage"],
            bom_rows=tuple(slot["bom_rows"]),
            row_number=slot["row_number"],
        )
        for (model_key, finished_good_key, assembly_item_key, parts), slot
        in raw_slots.items()
    )


def _summary_models_by_source_row(
    source_rows: tuple[ExcessSummarySourceRow, ...],
) -> dict[int, str]:
    return {
        index + 3: _report_model_for_source(source)
        for index, source in enumerate(source_rows)
    }


def _report_model_for_source(source: ExcessSummarySourceRow) -> str:
    """Use the MODEL value that is written to the Excess report as the owner."""
    model = str(source.model or "").strip()
    return model or UNMATCHED_BOM_MODEL


def _bom_models_by_source_row(
    source_rows: tuple[ExcessSummarySourceRow, ...],
    usage_slots: tuple[_BomUsageSlot, ...],
) -> dict[int, str]:
    """Keep BOM classification separate from the Excess-report MODEL owner."""
    return {
        index + 3: _summary_model_for_source(source, usage_slots)
        for index, source in enumerate(source_rows)
    }


def _selected_report_models(
    selected_models: tuple[str, ...] | None,
) -> tuple[str, ...] | None:
    """Normalize selections that came directly from the Excess report MODEL column."""
    if selected_models is None:
        return None

    selected: list[str] = []
    selected_keys: set[str] = set()
    for model in selected_models:
        normalized = str(model).strip().casefold()
        if normalized and normalized not in selected_keys:
            selected.append(str(model).strip())
            selected_keys.add(normalized)
    return tuple(selected)


def _summary_model_for_source(
    source: ExcessSummarySourceRow,
    usage_slots: tuple[_BomUsageSlot, ...],
) -> str:
    part_keys = frozenset(
        _part_key(part)
        for part in source.parts
        if _part_key(part)
    )
    if not part_keys:
        return UNMATCHED_BOM_MODEL

    models_by_key: dict[str, str] = {}
    for slot in usage_slots:
        if slot.model_key in {"#n/a", "n/a"}:
            continue
        if part_keys.issubset(slot.part_usage):
            models_by_key.setdefault(slot.model_key, slot.model_name)

    if len(models_by_key) == 1:
        model_name = next(iter(models_by_key.values()))
        return SHARED_MODEL if "/" in model_name else model_name
    if len(models_by_key) > 1:
        return SHARED_MODEL
    return UNMATCHED_BOM_MODEL


def _is_semifinished_part(part: str) -> bool:
    """A 3-prefixed part is treated as a BOM assembly/半成品."""
    return str(part).strip().startswith("3")


def _is_semifinished_source(source: ExcessSummarySourceRow) -> bool:
    """Use the primary report part to decide whether D 欄 matching applies."""
    return bool(source.parts) and _is_semifinished_part(source.parts[0])


def _model_match_score(report_model: str, bom_model: str) -> int:
    """Prefer an exact/contained report MODEL match, but allow BOM aliases."""
    report_key = str(report_model or "").strip().casefold()
    bom_key = str(bom_model or "").strip().casefold()
    if report_key and report_key == bom_key:
        return 0
    if report_key and (report_key in bom_key or bom_key in report_key):
        return 1
    return 2


def _assembly_usage_qty(
    target_slot: _BomUsageSlot,
    usage_slots: tuple[_BomUsageSlot, ...],
    active_keys: frozenset[tuple[str, str, str]] = frozenset(),
) -> tuple[float, str | None]:
    """Return an assembly's quantity per finished good by walking D -> H paths.

    A top-level D equal to A is one assembly per finished good.  Nested
    assemblies multiply the quantities along their parent component path.
    """
    if target_slot.assembly_item_key == target_slot.finished_good_key:
        return 1.0, None
    if target_slot.assembly_key in active_keys:
        return 1.0, "BOM 組裝階層出現循環，半成品用量暫以 1 件/台估算"

    parent_candidates: list[tuple[int, _BomUsageSlot, float]] = []
    for parent in usage_slots:
        if (
            parent.model_key != target_slot.model_key
            or parent.finished_good_key != target_slot.finished_good_key
            or parent.assembly_item_key == target_slot.assembly_item_key
        ):
            continue
        quantity = parent.part_usage.get(target_slot.assembly_item_key)
        if quantity is not None and quantity > 0:
            parent_candidates.append((parent.row_number, parent, quantity))
    if not parent_candidates:
        return 1.0, "找不到半成品的上階 BOM 用量，暫以 1 件/台估算"

    _row_number, parent, child_quantity = min(
        parent_candidates,
        key=lambda item: item[0],
    )
    parent_quantity, parent_note = _assembly_usage_qty(
        parent,
        usage_slots,
        active_keys | {target_slot.assembly_key},
    )
    note = parent_note
    return parent_quantity * child_quantity, note


def _match_semifinished_source(
    source: ExcessSummarySourceRow,
    usage_slots: tuple[_BomUsageSlot, ...],
    slots_by_assembly_item: dict[str, tuple[_BomUsageSlot, ...]] | None = None,
) -> _BomMatch | None:
    """Match a 3-prefixed report part against BOM D/L ASSEMBLY_ITEM."""
    part_keys = {
        _part_key(part): str(part).strip()
        for part in source.parts
        if _part_key(part)
    }
    candidates: list[tuple[tuple[Any, ...], _BomUsageSlot]] = []
    report_model = _report_model_for_source(source)
    candidate_slots = (
        tuple(
            slot
            for part_key in part_keys
            for slot in slots_by_assembly_item.get(part_key, ())
        )
        if slots_by_assembly_item is not None
        else usage_slots
    )
    for slot in candidate_slots:
        if slot.assembly_item_key not in part_keys:
            continue
        score = _model_match_score(report_model, slot.model_name)
        candidates.append(
            (
                (
                    score,
                    slot.row_number,
                    0 if slot.assembly_item_key == _part_key(source.parts[0]) else 1,
                    slot.finished_good_key,
                ),
                slot,
            )
        )
    if not candidates:
        return None

    _sort_key, slot = min(candidates, key=lambda item: item[0])
    usage_qty, _usage_note = _assembly_usage_qty(slot, usage_slots)
    matched_part = slot.assembly_item_name
    return _BomMatch(
        model_key=slot.model_key,
        target_model=slot.model_name,
        calculation_top_assembly=slot.finished_good_name,
        matched_part=matched_part,
        usage_qty=usage_qty,
        bom_rows=((slot.row_number, matched_part, usage_qty),),
        slot_key=None,
        match_kind="semi",
        assembly_key=slot.assembly_key,
        assembly_row=slot.row_number,
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
        slot_key=(
            slot.model_key,
            slot.finished_good_key,
            slot.assembly_item_key,
            slot.parts,
        ),
        assembly_key=slot.assembly_key,
    )


def _bom_matches_by_source_model(
    source_rows: tuple[ExcessSummarySourceRow, ...],
    usage_slots: tuple[_BomUsageSlot, ...],
    bom_models_by_source_row: dict[int, str],
) -> dict[int, _BomMatch]:
    slot_lists_by_model: dict[str, list[_BomUsageSlot]] = {}
    for slot in usage_slots:
        slot_lists_by_model.setdefault(slot.model_key, []).append(slot)
    usage_slots_by_model = {
        model_key: tuple(slots)
        for model_key, slots in slot_lists_by_model.items()
    }
    slots_by_assembly_item: dict[str, list[_BomUsageSlot]] = {}
    for slot in usage_slots:
        slots_by_assembly_item.setdefault(slot.assembly_item_key, []).append(slot)
    indexed_assembly_slots = {
        part_key: tuple(slots)
        for part_key, slots in slots_by_assembly_item.items()
    }

    matches: dict[int, _BomMatch] = {}
    for index, source in enumerate(source_rows):
        if _is_semifinished_source(source):
            match = _match_semifinished_source(
                source,
                usage_slots,
                indexed_assembly_slots,
            )
            if match is not None:
                matches[index + 3] = match
            continue
        model_key = _model_key(bom_models_by_source_row.get(index + 3))
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
        tuple[str, str, str, frozenset[str]],
        list[_InventoryAllocation],
    ] = {}
    semifinished_allocations_by_assembly: dict[
        tuple[str, str, str],
        list[_InventoryAllocation],
    ] = {}
    for source_row, match in bom_matches_by_source_row.items():
        source = source_by_row.get(source_row)
        if source is None:
            continue
        allocation = _InventoryAllocation(source_row, source)
        if match.match_kind == "semi":
            if match.assembly_key is not None:
                semifinished_allocations_by_assembly.setdefault(
                    match.assembly_key,
                    [],
                ).append(allocation)
        elif match.slot_key is not None:
            allocations_by_slot.setdefault(match.slot_key, []).append(allocation)

    assignments = {
        entry.source_row: bom_matches_by_source_row[entry.source_row]
        for entries in top_five_by_model.values()
        for entry in entries
        if entry.source_row in bom_matches_by_source_row
    }

    recommendable_entries_by_model: dict[str, tuple[_TopFiveEntry, ...]] = {}
    for model, entries in top_five_by_model.items():
        recommendable_list: list[_TopFiveEntry] = []
        selected_assemblies: set[str] = set()
        for entry in entries:
            match = assignments.get(entry.source_row)
            if match is None:
                continue
            entry_parts = {_part_key(part) for part in entry.source.parts}
            if any(part in selected_assemblies for part in entry_parts):
                continue
            if match.match_kind == "semi":
                selected_assemblies.add(match.matched_part.casefold())
            recommendable_list.append(entry)
            if len(recommendable_list) >= SUMMARY_TOP_LIMIT:
                break
        recommendable_entries = tuple(recommendable_list)
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
        if match.match_kind == "semi" and match.assembly_key is not None:
            scenario_slots = tuple(
                slot
                for slot in usage_slots
                if slot.assembly_key == match.assembly_key
            )
            material_checks = _material_checks_for_semifinished_scenario(
                match,
                scenario_slots,
                usage_slots,
                allocations_by_slot,
                semifinished_allocations_by_assembly,
                quantity,
                item_data,
            )
        else:
            scenario_slots = tuple(
                slot
                for slot in usage_slots
                if (
                    slot.model_key == match.model_key
                    and match.slot_key is not None
                    and slot.finished_good_key == match.slot_key[1]
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
        owner = entry.model_key
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


def _component_slot_key(
    slot: _BomUsageSlot,
) -> tuple[str, str, str, frozenset[str]]:
    return (
        slot.model_key,
        slot.finished_good_key,
        slot.assembly_item_key,
        slot.parts,
    )


def _find_semifinished_slot(
    part_key: str,
    anchor_slot: _BomUsageSlot,
    usage_slots: tuple[_BomUsageSlot, ...],
) -> _BomUsageSlot | None:
    candidates = [
        slot
        for slot in usage_slots
        if (
            slot.assembly_item_key == part_key
            and slot.model_key == anchor_slot.model_key
            and slot.finished_good_key == anchor_slot.finished_good_key
        )
    ]
    return min(candidates, key=lambda slot: slot.row_number) if candidates else None


def _consume_inventory_allocations(
    allocations: list[_InventoryAllocation],
    required_qty: float,
    consumed_by_source: dict[int, float],
) -> tuple[
    float,
    float,
    float,
    float,
    float,
    float | None,
    float | None,
    float | None,
]:
    """Consume available Excess quantities once, returning qty and USD totals."""
    available_qty = 0.0
    po_qty = 0.0
    consumed_qty = 0.0
    available_amounts: list[float | None] = []
    consumed_amounts: list[float | None] = []
    remaining_amounts: list[float | None] = []
    remaining_need = max(required_qty, 0.0)

    for allocation in allocations:
        metrics = _last_period_metrics(allocation.source)
        source_qty = max(metrics.stock_qty, 0)
        already_consumed = min(
            source_qty,
            max(consumed_by_source.get(allocation.source_row, 0.0), 0.0),
        )
        available_from_source = max(source_qty - already_consumed, 0.0)
        consumed_from_source = min(available_from_source, remaining_need)
        consumed_by_source[allocation.source_row] = (
            already_consumed + consumed_from_source
        )
        remaining_need -= consumed_from_source
        available_qty += available_from_source
        consumed_qty += consumed_from_source
        po_qty += max(metrics.po_qty, 0)
        available_amounts.append(
            _source_amount_for_consumed_qty(allocation.source, available_from_source)
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
                max(available_from_source - consumed_from_source, 0.0),
            )
        )

    return (
        available_qty,
        po_qty,
        consumed_qty,
        max(required_qty - available_qty, 0.0),
        max(available_qty - required_qty, 0.0),
        _sum_optional_amount(available_amounts),
        _sum_optional_amount(consumed_amounts),
        _sum_optional_amount(remaining_amounts),
    )


def _material_check_from_allocations(
    *,
    parts: tuple[str, ...],
    bom_rows: tuple[_BomRowDetail, ...],
    usage_qty: float,
    required_qty: float,
    allocations: list[_InventoryAllocation],
    item_data: ItemWorkbookData | None,
    consumed_by_source: dict[int, float],
    role: str,
    note: str | None = None,
    extra_purchase_override: float | None = None,
) -> _MaterialCheck:
    (
        available_qty,
        po_qty,
        consumed_qty,
        calculated_extra_purchase_qty,
        remaining_qty,
        available_amount,
        consumed_amount,
        remaining_amount,
    ) = _consume_inventory_allocations(
        allocations,
        required_qty,
        consumed_by_source,
    )
    extra_purchase_qty = (
        calculated_extra_purchase_qty
        if extra_purchase_override is None
        else extra_purchase_override
    )
    note_parts = [value for value in (note,) if value]
    if not allocations:
        note_parts.append("找不到此必要料號的 Excess 來源")
    if any(
        amount is None
        for amount in (available_amount, consumed_amount, remaining_amount)
    ):
        note_parts.append("缺少單價，USD 金額無法完整計算")
    price = next(
        (
            allocation.source.price
            for allocation in allocations
            if allocation.source.price is not None
        ),
        _fallback_price_for_parts(item_data, frozenset(parts)),
    )
    return _MaterialCheck(
        parts=tuple(parts),
        bom_rows=bom_rows,
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
        role=role,
    )


def _material_checks_for_semifinished_scenario(
    match: _BomMatch,
    scenario_slots: tuple[_BomUsageSlot, ...],
    usage_slots: tuple[_BomUsageSlot, ...],
    allocations_by_slot: dict[
        tuple[str, str, str, frozenset[str]],
        list[_InventoryAllocation],
    ],
    semifinished_allocations_by_assembly: dict[
        tuple[str, str, str],
        list[_InventoryAllocation],
    ],
    order_quantity: int,
    item_data: ItemWorkbookData | None,
) -> tuple[_MaterialCheck, ...]:
    """Consume a semi-finished assembly first, then recursively explode D -> H."""
    if match.assembly_key is None:
        return ()

    consumed_by_source: dict[int, float] = {}
    root_slot = min(scenario_slots, key=lambda slot: slot.row_number, default=None)
    if root_slot is None:
        return ()

    checks: list[_MaterialCheck] = []

    def expand_assembly(
        assembly_slot: _BomUsageSlot,
        required_qty: float,
        *,
        display_usage_qty: float,
        role: str,
        active_assemblies: frozenset[tuple[str, str, str]],
    ) -> None:
        assembly_key = assembly_slot.assembly_key
        allocations = semifinished_allocations_by_assembly.get(assembly_key, [])
        note_parts: list[str] = []
        if assembly_key in active_assemblies:
            note_parts.append("BOM 組裝階層出現循環，停止繼續展開")
            remaining_need = required_qty
            available_qty = 0.0
            po_qty = 0.0
            consumed_qty = 0.0
            available_amount = 0.0
            consumed_amount = 0.0
            remaining_amount = 0.0
            after_stock_qty = 0.0
        else:
            (
                available_qty,
                po_qty,
                consumed_qty,
                _extra_purchase_qty,
                after_stock_qty,
                available_amount,
                consumed_amount,
                remaining_amount,
            ) = _consume_inventory_allocations(
                allocations,
                required_qty,
                consumed_by_source,
            )
            remaining_need = max(required_qty - consumed_qty, 0.0)

        if remaining_need > 1e-9:
            note_parts.append(
                f"半成品庫存不足 {_format_number(remaining_need)} 件，已展開下階物料"
            )
        checks.append(
            _MaterialCheck(
                parts=(assembly_slot.assembly_item_name,),
                bom_rows=(
                    (
                        assembly_slot.row_number,
                        assembly_slot.assembly_item_name,
                        1.0,
                    ),
                ),
                usage_qty=display_usage_qty,
                required_qty=required_qty,
                available_qty=available_qty,
                po_qty=po_qty,
                consumed_qty=consumed_qty,
                # 半成品不足時改由下階物料供應，不在半成品列重複要求採購。
                extra_purchase_qty=0.0,
                remaining_qty=after_stock_qty,
                available_amount=available_amount,
                consumed_amount=consumed_amount,
                remaining_amount=remaining_amount,
                price=next(
                    (
                        allocation.source.price
                        for allocation in allocations
                        if allocation.source.price is not None
                    ),
                    _fallback_price_for_parts(
                        item_data,
                        frozenset({assembly_slot.assembly_item_name}),
                    ),
                ),
                status=("足夠" if remaining_need <= 1e-9 else "不足；已展開下階物料"),
                note="；".join(note_parts) or None,
                role=role,
            )
        )

        if remaining_need <= 1e-9 or assembly_key in active_assemblies:
            return

        next_active = active_assemblies | {assembly_key}
        child_slots = sorted(
            {
                _component_slot_key(slot): slot
                for slot in usage_slots
                if slot.assembly_key == assembly_key
            }.values(),
            key=lambda slot: slot.row_number,
        )
        for child_slot in child_slots:
            child_usage, child_note = _slot_usage_qty(child_slot)
            if child_usage <= 0:
                continue
            child_required = remaining_need * child_usage
            semi_children = [
                part
                for part in child_slot.parts
                if _is_semifinished_part(part)
                and _find_semifinished_slot(part, assembly_slot, usage_slots) is not None
            ]
            if len(child_slot.parts) == 1 and semi_children:
                nested_slot = _find_semifinished_slot(
                    semi_children[0],
                    assembly_slot,
                    usage_slots,
                )
                if nested_slot is not None:
                    expand_assembly(
                        nested_slot,
                        child_required,
                        display_usage_qty=child_usage,
                        role="下階半成品",
                        active_assemblies=next_active,
                    )
                    continue

            checks.append(
                _material_check_from_allocations(
                    parts=tuple(sorted(child_slot.parts)),
                    bom_rows=child_slot.bom_rows,
                    usage_qty=child_usage,
                    required_qty=child_required,
                    allocations=allocations_by_slot.get(
                        _component_slot_key(child_slot),
                        [],
                    ),
                    item_data=item_data,
                    consumed_by_source=consumed_by_source,
                    role="下階物料",
                    note=child_note,
                )
            )

    expand_assembly(
        root_slot,
        order_quantity * match.usage_qty,
        display_usage_qty=match.usage_qty,
        role="半成品",
        active_assemblies=frozenset(),
    )
    return tuple(checks)


def _material_checks_for_scenario(
    scenario_slots: tuple[_BomUsageSlot, ...],
    allocations_by_slot: dict[
        tuple[str, str, str, frozenset[str]],
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
            _component_slot_key(slot),
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
                replace(suggestion, suggestion_id=f"建議編號{counter}")
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


def _recommendation_values(
    summary: _ModelSummary,
    gating: CellRichText | None,
    suggestions: tuple[_ProductionSuggestion, ...],
) -> list[Any]:
    values = _summary_values(
        summary.model,
        summary.on_hand_qty,
        summary.opo_qty,
        summary.on_hand_amount,
        summary.opo_amount,
        summary.total_amount,
        gating,
        suggestions,
    )
    return [values[0], summary.customer, *values[1:]]


def _merge_customer_values(
    current: str | None,
    incoming: str | None,
) -> str | None:
    values: list[str] = []
    for text in (current, incoming):
        if not text:
            continue
        tokens = [value.strip() for value in str(text).split("/")]
        index = 0
        while index < len(tokens):
            value = tokens[index]
            normalized = unicodedata.normalize("NFKC", value).strip().casefold()
            # Excel-style #N/A or N/A can be split into two tokens by the
            # slash separator; discard the pair as one empty value.
            if normalized in {"#n", "n"} and index + 1 < len(tokens):
                next_normalized = (
                    unicodedata.normalize("NFKC", tokens[index + 1])
                    .strip()
                    .casefold()
                )
                if next_normalized == "a":
                    index += 2
                    continue
            if value and not _is_empty_customer_value(value) and value not in values:
                values.append(value)
            index += 1
    return "/".join(values) or None


def _is_empty_customer_value(value: str) -> bool:
    normalized = unicodedata.normalize("NFKC", value).strip().casefold()
    return normalized in {
        "#n",
        "#a",
        "#n/a",
        "n/a",
        "na",
        "none",
        "null",
        "-",
        "--",
    }


def _suggestion_lines(
    suggestions: tuple[_ProductionSuggestion, ...],
    formatter: Any,
) -> str | None:
    if not suggestions:
        return None
    return "\n\n".join(
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
    preview = "\n".join(
        f"{detail}；"
        for detail in suggestion.shortage_details[:preview_limit]
    )
    remaining = len(suggestion.shortage_details) - preview_limit
    if remaining > 0:
        return (
            f"共 {len(suggestion.shortage_details)} 筆料號／替代料群不足：\n"
            f"{preview}\n"
            f"另有 {remaining} 筆，詳見供料明細"
        )
    return (
        f"共 {len(suggestion.shortage_details)} 筆料號／替代料群不足：\n"
        f"{preview}"
    )


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
                        check.role,
                        _format_supply_bom_rows(check),
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


def _format_supply_bom_rows(check: _MaterialCheck) -> str:
    if "半成品" in check.role:
        return "、".join(
            f"第{row_number}列（D欄 L ASSEMBLY_ITEM {part}）"
            for row_number, part, _quantity in check.bom_rows
        )
    return _format_bom_rows(check.bom_rows)


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
            top_entries,
            bom_matches_by_source_row,
            item_data,
        )
        worksheet.append(
            _recommendation_values(
                summary,
                gating,
                suggestions,
            )
        )

    if selected_models is not None:
        for model in selected_models:
            if model.casefold() in displayed_source_models:
                continue
            values = [None] * len(headers)
            values[7] = f"{model}：找不到該來源 MODEL 的 Excess 摘要資料。"
            worksheet.append(values)
    if worksheet.max_row == 3:
        values = [None] * len(headers)
        values[7] = "目前沒有符合條件的資料。"
        worksheet.append(values)

    total_row = worksheet.max_row + 1
    worksheet.append(
        [
            "總計",
            None,
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
    _format_summary_sheet(worksheet, total_row, customer_column=True)


def _gating_rich_text(
    top_entries: tuple[_TopFiveEntry, ...],
    bom_matches_by_source_row: dict[int, _BomMatch],
    item_data: ItemWorkbookData | None,
) -> CellRichText | None:
    if not top_entries:
        return None

    rich_text = CellRichText()
    _append_rich_text(rich_text, "建議邏輯")
    for index, entry in enumerate(top_entries):
        _append_rich_text(rich_text, "\n\n" if index else "\n")
        part_text = entry.source.parts[0] if entry.source.parts else ""
        match = bom_matches_by_source_row.get(entry.source_row)
        _append_rich_text(
            rich_text,
            f"Top {entry.rank} 料號：{part_text} USD {entry.stock_amount:,}\n",
        )
        if match is None:
            _append_rich_text(rich_text, "未找到可辨識的 BOM MODEL")
            continue
        item_type = (
            item_data.item_type_for(match.matched_part)
            if item_data is not None
            else None
        )
        if match.match_kind == "semi":
            _append_rich_text(
                rich_text,
                f"TYPE：{item_type or '空白'}（半成品免 TYPE 篩選）\n",
            )
        elif item_type:
            _append_rich_text(rich_text, f"TYPE：{item_type}\n")
        if match.match_kind == "semi":
            _append_rich_text(
                rich_text,
                f"半成品判定：料號以 3 開頭，匹配 BOM D 欄 L ASSEMBLY_ITEM："
                f"{match.matched_part}\n",
            )
        else:
            _append_rich_text(
                rich_text,
                "匹配 BOM H 欄 COMPONENT_ITEM\n",
            )
        _append_rich_text(
            rich_text,
            f"對應 BOM MODEL：{match.target_model}\n",
        )
        _append_rich_text(
            rich_text,
            f"對應 TOP_ASSEMBLY_ITEM：{match.calculation_top_assembly}\n",
        )
        if match.match_kind == "semi":
            _append_rich_text(
                rich_text,
                f"對應 BOM 列：第{match.assembly_row or match.bom_rows[0][0]}列 "
                f"（D欄 L ASSEMBLY_ITEM {match.matched_part}）\n",
            )
            _append_rich_text(
                rich_text,
                f"半成品每台用量：{_format_number(match.usage_qty)} 件/台"
                + (
                    "（D=A 頂層組裝品，按每台 1 件計算）"
                    if abs(match.usage_qty - 1.0) < 1e-9
                    else "（依 BOM 階層用量乘算）"
                )
                + "\n",
            )
        else:
            _append_rich_text(
                rich_text,
                f"對應 BOM 列：{_format_bom_rows(match.bom_rows)}\n",
            )
            _append_rich_text(
                rich_text,
                f"EXTENDED_QT 使用量：{_format_bom_usage(match.bom_rows, match.usage_qty)}\n",
            )
        raw_quantity = entry.stock_qty / match.usage_qty
        usage_label = (
            "半成品每台用量"
            if match.match_kind == "semi"
            else "EXTENDED_QT"
        )
        _append_rich_text(
            rich_text,
            "計算：Excess stock QTY "
            f"{_format_number(entry.stock_qty)} ÷ {usage_label} "
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
            entry.model,
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


def _format_summary_sheet(
    worksheet: Any,
    total_row: int,
    *,
    customer_column: bool = False,
) -> None:
    column_count = len(SUMMARY_HEADERS) + (1 if customer_column else 0)
    _format_common_sheet(worksheet, column_count, total_row)
    worksheet.freeze_panes = "A4"
    worksheet.auto_filter.ref = f"A3:{get_column_letter(column_count)}{total_row - 1}"
    worksheet.row_dimensions[3].height = 58
    widths = [
        32,
        *([24] if customer_column else []),
        12,
        12,
        19,
        17,
        14,
        64,
        30,
        14,
        22,
        26,
        24,
        22,
        28,
        36,
        23,
    ]
    for column, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(column)].width = width

    narrative_start = 8 if customer_column else 7
    for row in range(4, total_row):
        line_count = max(
            _wrapped_line_count(
                worksheet.cell(row, column).value,
                worksheet.column_dimensions[get_column_letter(column)].width or 18,
            )
            for column in range(narrative_start, column_count + 1)
        )
        worksheet.row_dimensions[row].height = max(
            30,
            min(409.5, line_count * 18 + 8),
        )
    numeric_start = 3 if customer_column else 2
    for row in range(4, total_row + 1):
        for column in range(numeric_start, numeric_start + 5):
            worksheet.cell(row, column).number_format = "#,##0;[Red]-#,##0;0"
    for cell in worksheet[total_row]:
        cell.font = Font(name="Arial", size=9, bold=True, color="FF0070C0")


def _format_consumption_sheet(worksheet: Any, last_row: int) -> None:
    _format_common_sheet(worksheet, len(CONSUMPTION_HEADERS), last_row)
    worksheet.freeze_panes = "A4"
    worksheet.auto_filter.ref = (
        f"A3:{get_column_letter(len(CONSUMPTION_HEADERS))}{max(last_row, 3)}"
    )
    worksheet.row_dimensions[3].height = 58
    widths = [
        12, 24, 24, 28, 10, 28, 14, 20, 26, 23, 21, 12, 26, 34, 42
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
        12, 24, 28, 30, 16, 48, 12, 14, 18, 20, 14, 14, 16, 16, 15, 20, 20, 38
    ]
    for column, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(column)].width = width
    for row in range(4, last_row + 1):
        worksheet.row_dimensions[row].height = 42
        for column in (7, 8, 9, 11, 12, 13, 14):
            worksheet.cell(row, column).number_format = "#,##0.###;[Red]-#,##0.###;0"
        worksheet.cell(row, 10).number_format = "#,##0.00;[Red]-#,##0.00;0"
        worksheet.cell(row, 15).number_format = "#,##0.0000;[Red]-#,##0.0000;0"
        for column in (16, 17):
            worksheet.cell(row, column).number_format = "#,##0.00;[Red]-#,##0.00;0"
        if str(worksheet.cell(row, 18).value).startswith("不足"):
            worksheet.cell(row, 18).font = Font(name="Arial", size=9, color="FFFF0000")


def _format_detail_sheet(worksheet: Any, column_count: int, last_row: int) -> None:
    _format_common_sheet(worksheet, column_count, last_row)
    worksheet.freeze_panes = "A4"
    worksheet.auto_filter.ref = f"A3:{get_column_letter(column_count)}{max(last_row, 3)}"
    worksheet.row_dimensions[3].height = 44
    no_fill = PatternFill(fill_type=None)

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
    for row in worksheet.iter_rows(min_row=1, max_row=max(last_row, 3), max_col=column_count):
        for cell in row:
            cell.font = Font(name="Arial", size=9)
            cell.border = border
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    worksheet["A1"].font = Font(name="Arial", size=10, bold=True, color="FF704016")
    for column in range(1, column_count + 1):
        cell = worksheet.cell(3, column)
        cell.fill = PatternFill(
            "solid",
            fgColor=_header_fill_color(worksheet, column, column_count),
        )
        cell.font = Font(name="Arial", size=9, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in range(4, last_row + 1):
        for column in range(1, column_count + 1):
            cell = worksheet.cell(row, column)
            header = worksheet.cell(3, column).value
            if column == 1 or (
                worksheet.title
                in {SELECTED_RECOMMENDATIONS_SHEET, FULL_RECOMMENDATIONS_SHEET}
                and header
                in {
                    "Gating",
                    "AVTC建议加单",
                    "Raken建议加单",
                    "消耗率",
                    "建議料號可消耗\nUSD",
                    "所有相關料號原有\nExcess USD（加單前）",
                    "所有相關料號可消耗\nUSD",
                    "加單後剩餘\nExcess USD",
                    "庫存數量是否可滿足加單",
                    "需額外採購",
                    "AVTC 呆滯 预计TO-BE",
                    "Raken 呆滯 预计TO-BE",
                }
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


def _header_fill_color(worksheet: Any, column: int, column_count: int) -> str:
    """Return a consistent fill for a logical group of report columns."""
    blue = "D9EAF7"
    peach = "FCE4D6"
    yellow = "FFF2CC"
    green = "E2F0D9"
    red = "F4CCCC"
    purple = "E4DFEC"
    gray = "E7E6E6"
    title = worksheet.title

    if title in {SELECTED_RECOMMENDATIONS_SHEET, FULL_RECOMMENDATIONS_SHEET}:
        # RAKEN has Customer immediately after the first model column;
        # AVTC does not have that extra column.
        groups = (
            ((1, 7, blue), (8, 9, peach), (10, 10, yellow),
             (11, 14, green), (15, 16, red), (17, 17, gray))
            if column_count == len(SUMMARY_HEADERS) + 1
            else
            ((1, 6, blue), (7, 8, peach), (9, 9, yellow),
             (10, 13, green), (14, 15, red), (16, 16, gray))
        )
    elif title == "消耗建議":
        groups = (
            (1, 6, blue), (7, 7, peach), (8, 11, green),
            (12, 12, yellow), (13, 14, red), (15, 15, gray)
        )
    elif title == "供料明細":
        groups = (
            (1, 4, blue), (5, 6, peach), (7, 11, yellow),
            (12, 14, green), (15, 17, purple), (18, 18, gray)
        )
    elif title == "庫存金額Top5":
        header = worksheet.cell(3, column).value
        if header and (
            str(header).startswith("PartNo")
            or str(header).startswith("替代料")
        ):
            return blue
        if header in {"排名", "來源列", "MODEL"}:
            return blue
        if header == "Price（USD)":
            return yellow
        if header in {"Excess stock QTY", "Excess POQTY"}:
            return green
        if header in {
            "Excess stock Amount USD",
            "Excess POAmount USD",
            "Excess TotalAMT USD",
        }:
            return peach
        return gray
    else:
        groups = (
            (1, 3, blue), (4, 5, peach), (6, 6, yellow),
            (7, 7, gray), (8, 8, blue), (9, 9, yellow)
        )

    for start, end, color in groups:
        if start <= column <= end:
            return color
    return gray


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
