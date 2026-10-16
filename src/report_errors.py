"""Attach the actual file and operation to errors shown by the GUI and CLI."""

from __future__ import annotations

import ast
import re
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator
from zipfile import BadZipFile


class ReportError(Exception):
    def __init__(
        self,
        path: Path | None,
        operation: str,
        cause: Exception,
        sheet: str | None = None,
    ) -> None:
        self.path = path
        self.operation = operation
        self.cause = cause
        self.sheet = sheet
        super().__init__(self.user_message())

    def user_message(self) -> str:
        lines = [f"處理階段：{self.operation}"]
        if self.path is not None:
            lines.append(f"檔案／資料夾：{self.path}")
        if self.sheet:
            lines.append(f"工作表：{self.sheet}")
        reason, suggestion = explain_error(self.cause)
        lines.extend(("", f"原因：{reason}", f"建議：{suggestion}"))
        return "\n".join(lines)


@contextmanager
def file_context(path: Path | None, operation: str, sheet: str | None = None) -> Iterator[None]:
    try:
        yield
    except ReportError:
        raise
    except Exception as exc:
        raise ReportError(path, operation, exc, sheet) from exc


def explain_error(error: Exception) -> tuple[str, str]:
    text = str(error)
    if isinstance(error, PermissionError):
        return "無法存取檔案或資料夾。", "請關閉正在使用此檔案的 Excel，並確認資料夾可讀寫。"
    if isinstance(error, BadZipFile):
        return "Excel 檔案損壞，或內容不是有效的 .xlsx／.xlsm 檔案。", "請使用 Excel 開啟後重新另存，再執行一次。"
    missing_column = re.fullmatch(r"Cannot find required shortage column (.+) on row (\d+)", text)
    if missing_column:
        try:
            names = "／".join(ast.literal_eval(missing_column[1]))
        except (ValueError, SyntaxError, TypeError):
            names = missing_column[1]
        return f"第 {missing_column[2]} 列缺少必要欄位「{names}」。", "請確認 Shortage 的欄位名稱與檔案版本。"
    for keyword in ("BOM", "Item", "Price"):
        if text.startswith(f"Cannot find an Excel file with {keyword}"):
            return f"找不到檔名包含 {keyword} 的 Excel 檔。", "請將檔案放入所選工廠的 input 資料夾。"
        if text.startswith(f"Found multiple {keyword} Excel files:"):
            return f"找到多份 {keyword} 檔案：{text.split(':', 1)[1].strip()}", "請只保留本次要使用的一份來源檔。"
    if text.startswith("Cannot find a worksheet with 'shortage'"):
        return "找不到名稱包含 shortage 的工作表。", "請確認 Shortage 檔案已放入，且分頁名稱包含 shortage。"
    if text.startswith("Found multiple shortage worksheets"):
        return f"找到多份 Shortage 工作表：{text.split(':', 1)[1].strip()}", "請只保留本次要使用的一份 Shortage 來源。"
    if text.startswith("Found multiple previous Excess reports"):
        return f"找到多份上一版報表：{text.split(':', 1)[1].strip()}", "請只保留一份檔名包含 next_time 的上一版報表。"
    if text == "Cannot identify the shortage header row":
        return "無法辨識 Shortage 的表頭。", "請確認前 30 列內包含料號與必要欄位。"
    if text.startswith("Cannot find required sheet "):
        return f"找不到必要工作表「{text.removeprefix('Cannot find required sheet ')}」。", "請確認放入的是正確的來源檔案。"
    if text.startswith("Cannot find required ") and text.endswith(" sheet"):
        return f"找不到必要的 {text[21:-6]} 工作表或欄位。", "請確認工作表名稱、表頭與來源檔案版本。"
    if text.startswith("Cannot find required previous Excess column:"):
        return f"上一版報表缺少必要欄位：{text.split(':', 1)[1].strip()}", "請放入完整的 next_time_excess_report 報表。"
    if text.startswith("Cannot find previous Excess header row"):
        return "上一版報表找不到包含 PartNo1 的表頭。", "請確認放入的是程式輸出的 next_time_excess_report。"
    if text.startswith("Ambiguous previous Excess match at row "):
        return f"上一版報表有重複料號或替代料群組：{text.removeprefix('Ambiguous previous Excess match at row ')}", "請移除重複資料，避免承接到錯誤的人工資訊。"
    if "has an uncached formula" in text:
        return f"上一版報表的公式尚未儲存計算結果。\n{text}", "請在 Excel 重新計算並儲存上一版報表。"
    if isinstance(error, FileNotFoundError):
        return "找不到指定的檔案或資料夾。", "請確認檔案仍在原位置，且已完整解壓縮軟體資料夾。"
    return text or type(error).__name__, "請檢查上述檔案與處理階段；若仍無法排除，請提供此訊息。"
