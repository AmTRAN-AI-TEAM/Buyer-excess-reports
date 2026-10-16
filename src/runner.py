from __future__ import annotations

import argparse
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Sequence

try:
    from tqdm import tqdm
except Exception:  # pragma: no cover - optional dependency fallback
    tqdm = None

from bom_alternates import load_bom_alternates
from customer_profiles import profile_for
from excess_builder import build_excess_rows
from formatting import write_excess_workbook
from incomplete_alternates import write_incomplete_alternates_workbook
from previous_report import find_previous_excess_report, load_previous_excess_source
from workbook_io import (
    find_bom_file,
    find_item_file,
    find_price_file,
    find_shortage_source,
    load_item_workbook_data,
    load_shortage_records,
)


CUSTOMERS = ("AVTC", "RAKEN")
PROGRESS_STEPS = 11
NEXT_TIME_EXCESS_REPORT_NAME = "next_time_excess_report.xlsx"


@dataclass(frozen=True)
class RunResult:
    bom_file: Path
    item_file: Path
    price_file: Path | None
    shortage_file: Path
    shortage_sheet: str
    output_file: Path
    next_time_output_file: Path
    incomplete_alternates_file: Path
    input_part_count: int
    output_row_count: int
    part_column_count: int


class Progress:
    def __init__(self, total: int, enabled: bool = True):
        self.total = total
        self.enabled = enabled
        self.current = 0
        self.active_label: str | None = None
        self.bar = None

    def __enter__(self) -> "Progress":
        if self.enabled and tqdm is not None and self.total > 0 and sys.stdout.isatty():
            self.bar = tqdm(
                total=self.total,
                desc="Buyer Excess Reports",
                unit="step",
                dynamic_ncols=True,
                leave=True,
            )
        return self

    def step(self, label: str) -> None:
        self._complete_active_step()
        self.active_label = label
        if self.bar is not None:
            self.bar.set_description_str(label)
            return

        if self.enabled:
            print(f"[{self.current}/{self.total}] {label}", flush=True)

    def message(self, message: str) -> None:
        if self.bar is not None:
            self.bar.write(message)
        else:
            print(message)

    def warnings(self, messages: Sequence[str]) -> None:
        if not messages:
            return
        separator = "＝" * 40
        self.message("\n".join([
            "", separator,
            *(f"⚠️ {message}" for message in messages),
            separator, "",
        ]))

    def __exit__(self, exc_type, exc, tb) -> None:
        if exc_type is None:
            self.finish()
        if self.bar is not None:
            self.bar.close()

    def finish(self) -> None:
        self._complete_active_step()
        self.active_label = None
        if self.current < self.total:
            remaining = self.total - self.current
            self.current = self.total
            if self.bar is not None:
                self.bar.update(remaining)

        if self.bar is not None:
            self.bar.set_description_str("Done")
        elif self.enabled:
            print(f"[{self.current}/{self.total}] Done", flush=True)

    def _complete_active_step(self) -> None:
        if self.active_label is None or self.current >= self.total:
            return

        self.current += 1
        if self.bar is not None:
            self.bar.update(1)


def is_frozen_app() -> bool:
    return bool(getattr(sys, "frozen", False))


def project_root() -> Path:
    if is_frozen_app():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def generate_report(
    input_dir: Path,
    output_file: Path,
    *,
    customer: str | None = None,
    show_progress: bool = True,
) -> RunResult:
    input_dir = input_dir.resolve()
    output_file = output_file.resolve()
    customer_code = resolve_customer(customer, input_dir, output_file)
    profile = profile_for(customer_code)

    with Progress(PROGRESS_STEPS, enabled=show_progress) as progress:
        progress.step("Find BOM")
        bom_file = find_bom_file(input_dir)

        progress.step("Find Item workbook")
        item_file = find_item_file(input_dir)
        price_file = find_price_file(input_dir) if profile and profile.separate_price_file else None

        progress.step("Find shortage sheet")
        shortage_source = find_shortage_source(input_dir, bom_file)

        progress.step("Find previous Excess report")
        previous_report_file = find_previous_excess_report(input_dir)

        progress.message(f"BOM: {bom_file}")
        progress.message(f"Item: {item_file}")
        if price_file is not None:
            progress.message(f"Price: {price_file}")
        progress.message(f"Shortage: {shortage_source.path} [{shortage_source.sheet_name}]")
        if previous_report_file is not None:
            progress.message(f"Previous Excess: {previous_report_file}")

        progress.step("Read BOM alternates")
        alternates = load_bom_alternates(bom_file)

        progress.step("Read Item/Price data")
        item_data = load_item_workbook_data(
            item_file,
            price_path=price_file,
            price_headers=profile.price_headers if profile else ("Uprice",),
            preferred_cost_sheet=profile.preferred_cost_sheet if profile else None,
        )

        progress.step("Read shortage data")
        shortage_records = load_shortage_records(shortage_source)

        progress.step("Read previous Excess data")
        previous_source = None
        warning_messages: list[str] = []
        if previous_report_file is None:
            previous_excess_totals = None
            warning_messages.append("找不到上一版來源，不輸出improve資訊")
        else:
            previous_source = load_previous_excess_source(
                previous_report_file,
                customer=customer_code,
            )
            previous_excess_totals = previous_source.totals
            warning_messages.extend(previous_source.warnings)

        progress.step("Build Excess rows")
        include_excess_input_columns = customer_code == "RAKEN"
        result = build_excess_rows(
            shortage_records,
            alternates,
            item_data,
            include_excess_input_columns=include_excess_input_columns,
            previous_excess_totals=previous_excess_totals,
            previous_excess_inputs=previous_source.inputs if previous_source else None,
            customer=customer_code,
        )
        progress.warnings([*warning_messages, *result.warnings])

        progress.step("Write Excel files")
        write_excess_workbook(
            result.rows,
            output_file,
            result.row1_headers,
            result.row2_headers,
            result.part_column_count,
            customer=customer_code,
            customer_sheet_rows=(
                previous_source.inputs.customer_lookup.sheet_rows
                if previous_source and previous_source.inputs and previous_source.inputs.customer_lookup
                else None
            ),
        )
        next_time_output_file = output_file.with_name(NEXT_TIME_EXCESS_REPORT_NAME)
        if output_file.resolve() != next_time_output_file.resolve():
            shutil.copyfile(output_file, next_time_output_file)

        progress.step("Write incomplete alternates file")
        incomplete_alternates_file = output_file.with_name("incomplete_alternates.xlsx")
        write_incomplete_alternates_workbook(
            alternates.incomplete_alternate_groups,
            incomplete_alternates_file,
        )

        progress.message(f"Read shortage parts: {result.input_part_count}")
        progress.message(f"Generated Excess rows: {result.output_row_count}")
        progress.message(f"Part columns: {result.part_column_count}")
        progress.message(f"Incomplete alternate groups: {len(alternates.incomplete_alternate_groups)}")
        progress.message(f"Output: {output_file}")
        progress.message(f"Next-time Excess source: {next_time_output_file}")
        progress.message(f"Incomplete alternates: {incomplete_alternates_file}")

    return RunResult(
        bom_file=bom_file,
        item_file=item_file,
        price_file=price_file,
        shortage_file=shortage_source.path,
        shortage_sheet=shortage_source.sheet_name,
        output_file=output_file,
        next_time_output_file=next_time_output_file,
        incomplete_alternates_file=incomplete_alternates_file,
        input_part_count=result.input_part_count,
        output_row_count=result.output_row_count,
        part_column_count=result.part_column_count,
    )


def customer_input_dir(root: Path, customer: str) -> Path:
    return root / "input" / customer.upper()


def customer_output_file(root: Path, customer: str) -> Path:
    return root / "output" / customer.upper() / "excess_report.xlsx"


def resolve_customer(customer: str | None, *paths: Path | None) -> str | None:
    if customer:
        return customer.upper()

    matches = {
        part.upper()
        for path in paths
        if path is not None
        for part in path.parts
        if part.upper() in CUSTOMERS
    }
    if len(matches) > 1:
        names = ", ".join(sorted(matches))
        raise ValueError(f"Cannot infer one customer from paths: {names}")
    return next(iter(matches), None)


def select_customer_window() -> str | None:
    try:
        import tkinter as tk
        from tkinter import ttk
    except Exception as exc:  # pragma: no cover - depends on Windows runtime
        print(f"Cannot open customer selection window: {exc}", file=sys.stderr)
        return None

    selected = {"value": None}
    try:
        root = tk.Tk()
    except Exception as exc:  # pragma: no cover - depends on Windows runtime
        print(f"Cannot create customer selection window: {exc}", file=sys.stderr)
        return None

    root.title("Buyer Excess Reports")
    root.resizable(False, False)
    root.attributes("-topmost", True)

    frame = ttk.Frame(root, padding=20)
    frame.grid(row=0, column=0, sticky="nsew")
    ttk.Label(
        frame,
        text="請選擇本次要產出的客戶",
        font=("Microsoft JhengHei UI", 11, "bold"),
    ).grid(row=0, column=0, columnspan=2, sticky="w")
    ttk.Label(
        frame,
        text="請先把 BOM、Item 與 shortage 檔放進對應 input 資料夾。",
    ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 14))

    def choose(value: str) -> None:
        selected["value"] = value
        root.destroy()

    for col, customer in enumerate(CUSTOMERS):
        ttk.Button(
            frame,
            text=customer,
            command=lambda value=customer: choose(value),
            width=16,
        ).grid(row=2, column=col, padx=4, ipadx=6, ipady=6)

    ttk.Label(
        frame,
        text="關閉視窗或按 Esc 會取消本次執行。",
    ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(14, 0))
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.bind("<Escape>", lambda _event: root.destroy())
    _center_tk_window(root)
    root.mainloop()
    return selected["value"]


def select_customer_prompt() -> str | None:
    if not sys.stdin.isatty():
        print("未指定客戶，請加上 --customer AVTC 或 --customer RAKEN。", file=sys.stderr)
        return None

    print("請選擇本次要產出的客戶：")
    for index, customer in enumerate(CUSTOMERS, start=1):
        print(f"  {index}. {customer}")

    valid = {customer.casefold(): customer for customer in CUSTOMERS}
    valid.update({str(index): customer for index, customer in enumerate(CUSTOMERS, start=1)})

    while True:
        choice = input("輸入 1/2 或 AVTC/RAKEN，直接 Enter 取消：").strip()
        if not choice:
            return None
        selected = valid.get(choice.casefold())
        if selected:
            return selected
        print("輸入無法辨識，請重新輸入。")


def select_customer() -> str | None:
    if is_frozen_app() or sys.platform.startswith("win"):
        selected = select_customer_window()
        if selected:
            return selected
    return select_customer_prompt()


def _center_tk_window(root) -> None:
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    left = max((root.winfo_screenwidth() - width) // 2, 0)
    top = max((root.winfo_screenheight() - height) // 2, 0)
    root.geometry(f"+{left}+{top}")


def pause_for_windows_exe(enabled: bool) -> None:
    if not enabled:
        return
    try:
        input("\n按 Enter 結束...")
    except EOFError:
        return


def build_parser(root: Path) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Generate Buyer Excess report from BOM, Item/Price, and a shortage worksheet."
    )
    parser.add_argument(
        "--customer",
        choices=CUSTOMERS,
        default=None,
        help="Customer folder under input/output. If omitted, asks you to choose AVTC or RAKEN.",
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=None,
        help="Override input folder. If omitted with --customer, uses input/<customer>.",
    )
    parser.add_argument(
        "--output-file",
        type=Path,
        default=None,
        help="Override output file. If omitted with --customer, uses output/<customer>/excess_report.xlsx.",
    )
    parser.add_argument(
        "--no-pause",
        action="store_true",
        help="Do not wait for Enter after Windows exe completes.",
    )
    parser.add_argument(
        "--no-progress",
        action="store_true",
        help="Disable the tqdm progress display.",
    )
    parser.set_defaults(root=root)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    root = project_root()
    parser = build_parser(root)
    args = parser.parse_args(argv)

    pause_after_run = is_frozen_app() and not args.no_pause
    try:
        customer = args.customer.upper() if args.customer else None
        has_custom_paths = args.input_dir is not None or args.output_file is not None
        if customer is None and not has_custom_paths:
            customer = select_customer()
            if customer is None:
                print("已取消執行。")
                return 1
            customer = customer.upper()

        if args.input_dir is not None:
            input_dir = args.input_dir
        elif customer is not None:
            input_dir = customer_input_dir(root, customer)
        else:
            input_dir = root / "input"

        customer = resolve_customer(customer, input_dir, args.output_file)

        if args.output_file is not None:
            output_file = args.output_file
        elif customer is not None:
            output_file = customer_output_file(root, customer)
        else:
            output_file = root / "output" / "excess_report.xlsx"

        customer = resolve_customer(customer, input_dir, output_file)

        if customer:
            print(f"Customer: {customer}")
        generate_report(input_dir, output_file, customer=customer, show_progress=not args.no_progress)
        print("Done.")
        return 0
    except Exception as exc:  # noqa: BLE001 - show friendly exe errors
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    finally:
        pause_for_windows_exe(pause_after_run)
