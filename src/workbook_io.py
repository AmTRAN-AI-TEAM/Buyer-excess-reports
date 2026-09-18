from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


EXCEL_EXTENSIONS = {".xlsx", ".xlsm"}


@dataclass(frozen=True)
class ShortageSource:
    path: Path
    sheet_name: str


@dataclass(frozen=True)
class ShortageColumns:
    header_row: int
    part_no: int
    description: int
    planner: int
    buyer: int
    lt: int
    moq: int
    wo_demand: int | None
    over_shortage: int | None
    hld: int | None
    overshortage_with_hld: int | None
    po_remain: int | None
    mps_columns: tuple[int, ...]


@dataclass(frozen=True)
class ShortageRecord:
    part_no: str
    description: Any
    planner: Any
    buyer: Any
    lt: Any
    moq: Any
    overshortage: float
    wo_demand: float
    open_po: float
    source_row: int


def excel_files(input_dir: Path) -> list[Path]:
    return sorted(
        p
        for p in input_dir.iterdir()
        if p.is_file()
        and p.suffix.lower() in EXCEL_EXTENSIONS
        and not p.name.startswith("~$")
    )


def find_bom_file(input_dir: Path) -> Path:
    candidates = [p for p in excel_files(input_dir) if p.name.lower() == "bom.xlsx"]
    if len(candidates) == 1:
        return candidates[0]

    if not candidates:
        raise FileNotFoundError(f"Cannot find BOM.xlsx in {input_dir}")

    names = ", ".join(p.name for p in candidates)
    raise ValueError(f"Found multiple BOM.xlsx files: {names}")


def find_shortage_source(input_dir: Path, bom_file: Path) -> ShortageSource:
    matches: list[ShortageSource] = []
    for workbook_path in excel_files(input_dir):
        if workbook_path.resolve() == bom_file.resolve():
            continue

        workbook = load_workbook(workbook_path, read_only=True, data_only=True)
        for sheet_name in workbook.sheetnames:
            if "shortage" in sheet_name.lower():
                matches.append(ShortageSource(workbook_path, sheet_name))
        workbook.close()

    if len(matches) == 1:
        return matches[0]

    if not matches:
        raise FileNotFoundError(
            f"Cannot find a worksheet with 'shortage' in its name under {input_dir}"
        )

    details = ", ".join(f"{m.path.name}:{m.sheet_name}" for m in matches)
    raise ValueError(f"Found multiple shortage worksheets; please keep only one: {details}")


def load_shortage_records(source: ShortageSource) -> list[ShortageRecord]:
    workbook = load_workbook(source.path, read_only=True, data_only=True)
    worksheet = workbook[source.sheet_name]
    columns = detect_shortage_columns(worksheet)

    records: list[ShortageRecord] = []
    for row_index, row in enumerate(
        worksheet.iter_rows(min_row=columns.header_row + 1, values_only=True),
        start=columns.header_row + 1,
    ):
        part_no = _cell(row, columns.part_no)
        if part_no in (None, ""):
            continue

        overshortage = _overshortage_value(row, columns)
        wo_demand = _wo_demand_value(row, columns)
        open_po = _number(_cell(row, columns.po_remain)) if columns.po_remain else 0.0

        records.append(
            ShortageRecord(
                part_no=str(part_no).strip(),
                description=_cell(row, columns.description),
                planner=_cell(row, columns.planner),
                buyer=_cell(row, columns.buyer),
                lt=_cell(row, columns.lt),
                moq=_cell(row, columns.moq),
                overshortage=overshortage,
                wo_demand=wo_demand,
                open_po=open_po,
                source_row=row_index,
            )
        )

    workbook.close()
    return records


def detect_shortage_columns(worksheet: Worksheet) -> ShortageColumns:
    header_row = _find_header_row(worksheet)
    headers = {
        _normalize_header(worksheet.cell(header_row, column).value): column
        for column in range(1, worksheet.max_column + 1)
    }

    def require(*names: str) -> int:
        for name in names:
            normalized = _normalize_header(name)
            if normalized in headers:
                return headers[normalized]
        raise ValueError(
            f"Cannot find required shortage column {names} on row {header_row}"
        )

    def optional(*names: str) -> int | None:
        for name in names:
            normalized = _normalize_header(name)
            if normalized in headers:
                return headers[normalized]
        return None

    mps_columns = tuple(
        column
        for column in range(1, worksheet.max_column + 1)
        if _normalize_header(worksheet.cell(header_row - 1, column).value) == "mps"
    )

    return ShortageColumns(
        header_row=header_row,
        part_no=require("Part No", "PartNo", "Part Number"),
        description=require("DESCRIPTION", "Description", "Desc"),
        planner=require("PLANNER", "Planner"),
        buyer=require("Buyer", "GSD"),
        lt=require("LT"),
        moq=require("MOQ"),
        wo_demand=1,
        over_shortage=optional("OVER_SHORTAGE", "Overshortage"),
        hld=optional("HLD"),
        overshortage_with_hld=optional("Overshortage1"),
        po_remain=optional("PO_REMAIN", "PO REMAIN", "Open PO"),
        mps_columns=mps_columns,
    )


def _find_header_row(worksheet: Worksheet) -> int:
    for row in range(1, min(worksheet.max_row, 30) + 1):
        values = [
            _normalize_header(worksheet.cell(row, column).value)
            for column in range(1, min(worksheet.max_column, 40) + 1)
        ]
        if "partno" in values and (
            "description" in values or "desc" in values
        ):
            return row
    raise ValueError("Cannot identify the shortage header row")


def _overshortage_value(row: tuple[Any, ...], columns: ShortageColumns) -> float:
    if columns.overshortage_with_hld:
        value = _cell(row, columns.overshortage_with_hld)
        if value is not None:
            return _number(value)

    over_shortage = _number(_cell(row, columns.over_shortage)) if columns.over_shortage else 0
    hld = _number(_cell(row, columns.hld)) if columns.hld else 0
    return over_shortage + hld


def _wo_demand_value(row: tuple[Any, ...], columns: ShortageColumns) -> float:
    if columns.wo_demand:
        value = _cell(row, columns.wo_demand)
        if value is not None:
            return _number(value)

    return sum(_number(_cell(row, column)) for column in columns.mps_columns)


def _cell(row: tuple[Any, ...], one_based_column: int | None) -> Any:
    if one_based_column is None:
        return None
    index = one_based_column - 1
    if index >= len(row):
        return None
    return row[index]


def _number(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except ValueError:
        return 0.0


def _normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return "".join(ch for ch in str(value).strip().lower() if ch.isalnum())
