"""Read-only source checks using the same discovery rules as report generation."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from previous_report import find_previous_excess_report
from report_errors import ReportError, explain_error
from workbook_io import (
    EXCEL_EXTENSIONS,
    find_bom_file,
    find_item_file,
    find_price_file,
    find_shortage_source,
)


@dataclass(frozen=True)
class SourceRequirement:
    key: str
    label: str
    optional: bool = False


@dataclass(frozen=True)
class SourceCheck:
    requirement: SourceRequirement
    status: str
    detail: str

    @property
    def blocks_run(self) -> bool:
        return self.status == "error" or (self.status == "missing" and not self.requirement.optional)


def source_requirements(customer: str) -> tuple[SourceRequirement, ...]:
    if customer not in ("AVTC", "RAKEN"):
        raise ValueError(f"Unknown customer: {customer}")
    requirements = [
        SourceRequirement("shortage", "Shortage"),
        SourceRequirement("bom", "BOM"),
        SourceRequirement("item", "Item" if customer == "AVTC" else "Item information"),
    ]
    if customer == "RAKEN":
        requirements.append(SourceRequirement("price", "Price"))
    requirements.append(SourceRequirement("previous", "next_time_excess_report", optional=True))
    return tuple(requirements)


def input_signature(input_dir: Path) -> tuple:
    """Notice additions, removals, saves and replacements without rereading Excel."""
    try:
        files = []
        for path in input_dir.iterdir():
            if path.is_file() and path.suffix.lower() in EXCEL_EXTENSIONS and not path.name.startswith("~$"):
                stat = path.stat()
                files.append((path.name, stat.st_size, stat.st_mtime_ns, stat.st_ctime_ns))
        return tuple(sorted(files))
    except OSError as exc:
        return (type(exc).__name__, str(exc))


def check_sources(input_dir: Path, customer: str) -> tuple[SourceCheck, ...]:
    requirements = source_requirements(customer)
    results: dict[str, SourceCheck] = {}
    bom_file = None
    # Find BOM first because the runner excludes it when looking for Shortage.
    finders = {"bom": find_bom_file, "item": find_item_file, "price": find_price_file,
               "previous": find_previous_excess_report}
    for requirement in sorted(requirements, key=lambda entry: entry.key != "bom"):
        try:
            if requirement.key == "shortage":
                source = find_shortage_source(input_dir, bom_file)
                detail = f"{source.path.name}  ／  {source.sheet_name}"
                path = source.path
            else:
                path = finders[requirement.key](input_dir)
                detail = path.name if path is not None else "未提供（選填，仍可執行）"
                if requirement.key == "bom":
                    bom_file = path
            status = "found" if path is not None else "missing"
        except (FileNotFoundError, NotADirectoryError):
            status = "missing"
            detail = "未提供（選填，仍可執行）" if requirement.optional else (
                "未找到含 shortage 的工作表" if requirement.key == "shortage" else "未找到符合檔名的 Excel 檔案"
            )
        except Exception as exc:
            status = "error"
            cause = exc.cause if isinstance(exc, ReportError) else exc
            detail = explain_error(cause)[0]
            if isinstance(exc, ReportError) and exc.path is not None:
                detail = f"{exc.path.name}：{detail}"
        results[requirement.key] = SourceCheck(requirement, status, detail)
    return tuple(results[requirement.key] for requirement in requirements)
