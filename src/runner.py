from __future__ import annotations

import argparse
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path
from time import monotonic
from typing import Callable, Sequence

try:
    from tqdm import tqdm
except Exception:  # pragma: no cover - optional dependency fallback
    tqdm = None

from bom_alternates import load_bom_alternates, load_bom_models
from customer_profiles import profile_for
from excess_builder import build_excess_rows
from excess_summary import write_excess_summary
from formatting import write_excess_workbook
from incomplete_alternates import write_incomplete_alternates_workbook
from previous_report import find_previous_excess_report, load_previous_excess_source
from report_errors import file_context
from workbook_io import (
    find_bom_file,
    find_item_file,
    find_price_file,
    find_shortage_source,
    load_item_workbook_data,
    load_shortage_records,
)


CUSTOMERS = ("AVTC", "RAKEN")
PROGRESS_LABELS = (
    "Find BOM",
    "Find Item workbook",
    "Find shortage sheet",
    "Find previous Excess report",
    "Read BOM alternates",
    "Read Item/Price data",
    "Read shortage data",
    "Read previous Excess data",
    "Build Excess rows",
    "Write Excel files",
    "Write Excess summary",
    "Write incomplete alternates file",
)
PROGRESS_STEPS = len(PROGRESS_LABELS)
NEXT_TIME_EXCESS_REPORT_NAME = "next_time_excess_report.xlsx"


def progress_steps_for_customer(
    customer: str | None,
    *,
    include_summary: bool = True,
) -> int:
    """Return the number of real processing stages for a customer."""
    has_summary = bool(customer and customer.upper() in {"AVTC", "RAKEN"})
    return PROGRESS_STEPS - (0 if has_summary and include_summary else 1)


@dataclass(frozen=True)
class ProgressEvent:
    kind: str
    current: int
    total: int
    label: str = ""
    message: str = ""
    elapsed_seconds: float = 0.0
    stage_durations: tuple[tuple[str, float], ...] = ()


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
    summary_output_file: Path | None
    input_part_count: int
    output_row_count: int
    part_column_count: int


@dataclass(frozen=True)
class SummaryModelSelection:
    selected_models: tuple[str, ...] | None
    include_summary: bool
    cancelled: bool = False


class Progress:
    def __init__(self, total: int, enabled: bool = True, callback: Callable[[ProgressEvent], None] | None = None):
        self.total = total
        self.enabled = enabled
        self.current = 0
        self.active_label: str | None = None
        self.bar = None
        self.callback = callback
        self.started_at = monotonic()
        self.active_started_at = self.started_at
        self.stage_durations: dict[str, float] = {}

    def _event(self, kind: str, *, label: str = "", message: str = "") -> ProgressEvent:
        return ProgressEvent(
            kind, self.current, self.total, label, message,
            max(0.0, monotonic() - self.started_at), tuple(self.stage_durations.items()),
        )

    def __enter__(self) -> "Progress":
        if self.callback is None and self.enabled and tqdm is not None and self.total > 0 and sys.stdout is not None and sys.stdout.isatty():
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
        self.active_started_at = monotonic()
        if self.callback is not None:
            self.callback(self._event("progress", label=label))
            return
        if self.bar is not None:
            self.bar.set_description_str(label)
            return

        if self.enabled and sys.stdout is not None:
            print(f"[{self.current}/{self.total}] {label}", flush=True)

    def message(self, message: str) -> None:
        if self.callback is not None:
            self.callback(self._event("message", message=message))
        elif self.bar is not None:
            self.bar.write(message)
        elif sys.stdout is not None:
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

        if self.callback is not None:
            self.callback(self._event("progress", label="Done"))
        elif self.bar is not None:
            self.bar.set_description_str("Done")
        elif self.enabled and sys.stdout is not None:
            print(f"[{self.current}/{self.total}] Done", flush=True)

    def _complete_active_step(self) -> None:
        if self.active_label is None or self.current >= self.total:
            return

        self.current += 1
        self.stage_durations[self.active_label] = max(0.0, monotonic() - self.active_started_at)
        self.active_label = None
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
    on_event: Callable[[ProgressEvent], None] | None = None,
    summary_models: tuple[str, ...] | None = None,
    include_summary: bool = True,
) -> RunResult:
    input_dir = input_dir.resolve()
    output_file = output_file.resolve()
    customer_code = resolve_customer(customer, input_dir, output_file)
    profile = profile_for(customer_code)

    with Progress(
        progress_steps_for_customer(
            customer_code,
            include_summary=include_summary,
        ),
        enabled=show_progress,
        callback=on_event,
    ) as progress:
        progress.step("Find BOM")
        with file_context(input_dir, "尋找 BOM 檔案"):
            bom_file = find_bom_file(input_dir)

        progress.step("Find Item workbook")
        with file_context(input_dir, "尋找 Item 檔案"):
            item_file = find_item_file(input_dir)
        with file_context(input_dir, "尋找 Price 檔案"):
            price_file = find_price_file(input_dir) if profile and profile.separate_price_file else None

        progress.step("Find shortage sheet")
        with file_context(input_dir, "尋找 Shortage 工作表"):
            shortage_source = find_shortage_source(input_dir, bom_file)

        progress.step("Find previous Excess report")
        with file_context(input_dir, "尋找上一版報表"):
            previous_report_file = find_previous_excess_report(input_dir)

        progress.message(f"BOM: {bom_file}")
        progress.message(f"Item: {item_file}")
        if price_file is not None:
            progress.message(f"Price: {price_file}")
        progress.message(f"Shortage: {shortage_source.path} [{shortage_source.sheet_name}]")
        if previous_report_file is not None:
            progress.message(f"Previous Excess: {previous_report_file}")

        progress.step("Read BOM alternates")
        with file_context(bom_file, "讀取 BOM 替代料"):
            alternates = load_bom_alternates(
                bom_file,
                include_component_usages=customer_code in {"AVTC", "RAKEN"},
                component_usage_model_from_bom=customer_code in {"AVTC", "RAKEN"},
            )

        progress.step("Read Item/Price data")
        with file_context(item_file, "讀取 Item／Price 資料"):
            item_data = load_item_workbook_data(
                item_file,
                price_path=price_file,
                price_headers=profile.price_headers if profile else ("Uprice",),
                preferred_cost_sheet=profile.preferred_cost_sheet if profile else None,
                fallback_to_first_price_sheet=profile.fallback_to_first_price_sheet if profile else False,
            )

        progress.step("Read shortage data")
        with file_context(shortage_source.path, "讀取 Shortage 資料", shortage_source.sheet_name):
            shortage_records = load_shortage_records(shortage_source)
        progress.step("Read previous Excess data")
        previous_source = None
        warning_messages: list[str] = []
        if previous_report_file is None:
            previous_excess_totals = None
            warning_messages.append("找不到上一版來源，不輸出improve資訊")
        else:
            with file_context(previous_report_file, "讀取上一版報表"):
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
        with file_context(output_file, "寫入 Excess 報表"):
            write_excess_workbook(
                result.rows,
                output_file,
                result.row1_headers,
                result.row2_headers,
                result.part_column_count,
                customer=customer_code,
                alternate_price_rows=result.alternate_price_rows,
                customer_sheet_rows=(
                    previous_source.inputs.customer_sheet_rows
                    if previous_source and previous_source.inputs
                    else None
                ),
            )
        next_time_output_file = output_file.with_name(NEXT_TIME_EXCESS_REPORT_NAME)
        if output_file.resolve() != next_time_output_file.resolve():
            with file_context(next_time_output_file, "寫入下次使用的上一版報表"):
                shutil.copyfile(output_file, next_time_output_file)

        summary_output_file = None
        if customer_code in {"AVTC", "RAKEN"} and include_summary:
            progress.step("Write Excess summary")
            summary_output_file = output_file.with_name("excess_summary.xlsx")
            with file_context(summary_output_file, "寫入 Excess summary"):
                model_count, top_five_count = write_excess_summary(
                    result.summary_source_rows,
                    result.demand_labels[-1],
                    summary_output_file,
                    component_usages=alternates.component_usages,
                    customer=customer_code,
                    selected_models=summary_models,
                )
            progress.message(f"Excess summary models: {model_count}")
            progress.message(f"Per-model stock-amount top rows: {top_five_count}")

        progress.step("Write incomplete alternates file")
        incomplete_alternates_file = output_file.with_name("incomplete_alternates.xlsx")
        with file_context(incomplete_alternates_file, "寫入不完全替代料明細"):
            write_incomplete_alternates_workbook(
                alternates.incomplete_alternate_groups,
                incomplete_alternates_file,
            )

        progress.message(f"Read shortage parts: {result.input_part_count}")
        progress.message(f"Generated Excess rows: {result.output_row_count}")
        if customer_code in {"RAKEN", "AVTC"}:
            progress.message(f"替代料借價（紅字）: {len(result.alternate_price_rows)} 列")
        progress.message(f"Part columns: {result.part_column_count}")
        progress.message(f"Incomplete alternate groups: {len(alternates.incomplete_alternate_groups)}")
        progress.message(f"Output: {output_file}")
        progress.message(f"Next-time Excess source: {next_time_output_file}")
        if summary_output_file is not None:
            progress.message(f"Excess summary: {summary_output_file}")
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
        summary_output_file=summary_output_file,
        input_part_count=result.input_part_count,
        output_row_count=result.output_row_count,
        part_column_count=result.part_column_count,
    )


def customer_input_dir(root: Path, customer: str) -> Path:
    return root / "input" / customer.upper()


def customer_output_file(root: Path, customer: str) -> Path:
    return root / "output" / customer.upper() / "excess_report.xlsx"


def load_summary_models(input_dir: Path) -> tuple[str, ...]:
    """Read the distinct BOM MODEL values used by the summary selector."""
    bom_file = find_bom_file(input_dir)
    return load_bom_models(bom_file)


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


def select_summary_models_prompt(
    models: tuple[str, ...],
) -> tuple[tuple[str, ...] | None, bool]:
    """Prompt for summary target models; return (models, include_summary)."""
    selected = set(models)
    print("\n請選擇要列入 summary 第一張「加單建議」的 BOM MODEL：")
    print("預設全選；輸入搜尋文字篩選，編號可用逗號分隔切換勾選。")
    print("指令：A 全選、N 全取消、D 完成、S 略過 summary。")

    while True:
        query = input("搜尋 MODEL（直接 Enter 顯示全部）：").strip().casefold()
        matches = [model for model in models if query in model.casefold()]
        if matches:
            for index, model in enumerate(matches, start=1):
                mark = "x" if model in selected else " "
                print(f"  [{mark}] {index}. {model}")
        else:
            print("  沒有符合的 MODEL。")
        print(f"已選取 {len(selected)} / {len(models)} 個 MODEL。")

        action = input("切換編號 / A全選 / N全取消 / D完成 / S略過 summary：").strip()
        normalized = action.casefold()
        if normalized in {"a", "all", "全部"}:
            selected = set(models)
            continue
        if normalized in {"n", "none", "清除"}:
            selected.clear()
            continue
        if normalized in {"s", "skip"}:
            return None, False
        if normalized in {"d", "done", "完成"}:
            if selected:
                if len(selected) == len(models):
                    return None, True
                return tuple(model for model in models if model in selected), True
            print("目前沒有勾選 MODEL；請勾選項目，或輸入 S 略過 summary。")
            continue

        try:
            indexes = {
                int(value.strip())
                for value in action.split(",")
                if value.strip()
            }
        except ValueError:
            print("輸入無法辨識，請使用搜尋文字、編號或指令。")
            continue
        if not indexes or any(index < 1 or index > len(matches) for index in indexes):
            print(f"請輸入 1 到 {len(matches)} 的編號。")
            continue
        for index in indexes:
            model = matches[index - 1]
            if model in selected:
                selected.remove(model)
            else:
                selected.add(model)


def select_summary_models_window(
    models: tuple[str, ...],
) -> SummaryModelSelection | None:
    """Open a searchable MODEL picker; return None when Tk cannot open."""
    try:
        import tkinter as tk
        from tkinter import ttk
    except Exception as exc:  # pragma: no cover - depends on local Tk install
        print(f"無法載入 MODEL 選擇介面：{exc}", file=sys.stderr)
        return None

    try:
        root = tk.Tk()
    except Exception as exc:  # pragma: no cover - headless environments
        print(f"無法開啟 MODEL 選擇介面，改用終端機選單：{exc}", file=sys.stderr)
        return None

    result: list[SummaryModelSelection] = []
    background = "#fff7e8"
    panel = "#fff2db"
    ink = "#422715"
    muted = "#6b452b"

    root.title("選擇加單建議 MODEL")
    root.configure(bg=background)
    root.resizable(True, True)
    root.geometry("920x660")
    root.minsize(760, 500)
    try:
        root.attributes("-topmost", True)
        root.after(700, lambda: root.attributes("-topmost", False))
    except tk.TclError:
        pass

    body = tk.Frame(root, bg=background, padx=24, pady=20)
    body.pack(fill="both", expand=True)
    body.grid_columnconfigure(0, weight=1)
    body.grid_rowconfigure(2, weight=1)

    tk.Label(
        body,
        text="選擇要列入 summary「加單建議」的 BOM MODEL",
        bg=background,
        fg=ink,
        font=("TkDefaultFont", 17, "bold"),
        anchor="w",
    ).grid(row=0, column=0, sticky="ew")
    tk.Label(
        body,
        text="預設全選。可搜尋 MODEL 並個別勾選；略過會只跳過 summary。",
        bg=background,
        fg=muted,
        font=("TkDefaultFont", 10),
        anchor="w",
    ).grid(row=1, column=0, sticky="ew", pady=(5, 12))

    controls = tk.Frame(body, bg=background)
    controls.grid(row=2, column=0, sticky="nsew")
    controls.grid_columnconfigure(1, weight=1)
    controls.grid_rowconfigure(1, weight=1)

    search_var = tk.StringVar()
    tk.Label(
        controls,
        text="搜尋 MODEL：",
        bg=background,
        fg=ink,
        font=("TkDefaultFont", 10),
    ).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=(0, 9))
    search_entry = ttk.Entry(controls, textvariable=search_var)
    search_entry.grid(row=0, column=1, sticky="ew", padx=(0, 12), pady=(0, 9), ipady=5)
    all_var = tk.BooleanVar(value=bool(models))
    all_checkbox = tk.Checkbutton(
        controls,
        text="全選所有 MODEL",
        variable=all_var,
        bg=background,
        activebackground=background,
        fg=ink,
        font=("TkDefaultFont", 10, "bold"),
        command=lambda: toggle_all(),
    )
    all_checkbox.grid(row=0, column=2, sticky="e", pady=(0, 9))

    list_frame = tk.Frame(controls, bg=panel, highlightbackground="#b98952", highlightthickness=1)
    list_frame.grid(row=1, column=0, columnspan=3, sticky="nsew")
    list_frame.grid_columnconfigure(0, weight=1)
    list_frame.grid_rowconfigure(0, weight=1)
    canvas = tk.Canvas(list_frame, bg=panel, highlightthickness=0)
    canvas.grid(row=0, column=0, sticky="nsew")
    scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
    scrollbar.grid(row=0, column=1, sticky="ns")
    canvas.configure(yscrollcommand=scrollbar.set)
    inner = tk.Frame(canvas, bg=panel)
    inner_window = canvas.create_window((0, 0), window=inner, anchor="nw")
    inner.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", lambda event: canvas.itemconfigure(inner_window, width=event.width))
    canvas.bind("<MouseWheel>", lambda event: canvas.yview_scroll(-int(event.delta / 120), "units"))
    canvas.bind("<Button-4>", lambda _event: canvas.yview_scroll(-1, "units"))
    canvas.bind("<Button-5>", lambda _event: canvas.yview_scroll(1, "units"))

    variables: dict[str, tk.BooleanVar] = {}
    buttons: dict[str, tk.Checkbutton] = {}
    for index, model in enumerate(models):
        variable = tk.BooleanVar(value=True)
        button = tk.Checkbutton(
            inner,
            text=model,
            variable=variable,
            bg=panel,
            activebackground=panel,
            fg=ink,
            selectcolor=background,
            font=("TkDefaultFont", 10),
            anchor="w",
            justify="left",
            wraplength=370,
            command=lambda: sync_selection(),
        )
        button.grid(row=index // 2, column=index % 2, sticky="ew", padx=8, pady=3)
        variables[model] = variable
        buttons[model] = button
    inner.grid_columnconfigure(0, weight=1)
    inner.grid_columnconfigure(1, weight=1)

    footer = tk.Frame(body, bg=background)
    footer.grid(row=3, column=0, sticky="ew", pady=(11, 0))
    count_label = tk.Label(footer, bg=background, fg=muted, font=("TkDefaultFont", 10), anchor="w")
    count_label.pack(side="left", fill="x", expand=True)
    button_row = tk.Frame(footer, bg=background)
    button_row.pack(side="right")

    def close_with(selection: SummaryModelSelection) -> None:
        result.append(selection)
        root.destroy()

    def cancel() -> None:
        close_with(SummaryModelSelection(None, False, cancelled=True))

    def skip() -> None:
        close_with(SummaryModelSelection(None, False))

    def confirm() -> None:
        chosen = tuple(model for model in models if variables[model].get())
        if not chosen:
            root.bell()
            count_label.configure(text="請至少選擇一個 MODEL，或按略過 summary。", fg="#9b2727")
            return
        close_with(
            SummaryModelSelection(
                None if len(chosen) == len(models) else chosen,
                True,
            )
        )

    ttk.Button(button_row, text="取消執行", command=cancel).pack(side="left", padx=(0, 8), ipadx=8, ipady=5)
    ttk.Button(button_row, text="略過此頁，不產出summary", command=skip).pack(side="left", padx=(0, 8), ipadx=8, ipady=5)
    confirm_button = ttk.Button(button_row, text="確認選取", command=confirm)
    confirm_button.pack(side="left", ipadx=12, ipady=5)

    def filter_models(*_args) -> None:
        query = search_var.get().strip().casefold()
        visible = [model for model in models if query in model.casefold()]
        for button in buttons.values():
            button.grid_remove()
        for index, model in enumerate(visible):
            buttons[model].grid(row=index // 2, column=index % 2, sticky="ew", padx=8, pady=3)
        canvas.configure(scrollregion=canvas.bbox("all"))

    def sync_selection() -> None:
        selected_count = sum(variable.get() for variable in variables.values())
        all_var.set(bool(models) and selected_count == len(models))
        count_label.configure(text=f"已選取 {selected_count} / {len(models)} 個 MODEL", fg=muted)
        confirm_button.configure(state="normal" if selected_count else "disabled")

    def toggle_all() -> None:
        for variable in variables.values():
            variable.set(all_var.get())
        sync_selection()

    search_var.trace_add("write", filter_models)
    if not models:
        count_label.configure(text="BOM 中沒有可選的非空 MODEL。", fg="#a86312")
        all_checkbox.configure(state="disabled")
        confirm_button.configure(state="disabled")
    else:
        sync_selection()

    root.protocol("WM_DELETE_WINDOW", cancel)
    root.bind("<Escape>", lambda _event: cancel())
    _center_tk_window(root)
    root.grab_set()
    search_entry.focus_set()
    root.mainloop()
    return result[0] if result else SummaryModelSelection(None, False, cancelled=True)


def select_customer() -> str | None:
    if is_frozen_app():
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
    if not enabled or sys.stdin is None:
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

        summary_models: tuple[str, ...] | None = None
        include_summary = True
        if customer in {"AVTC", "RAKEN"} and sys.stdin is not None and sys.stdin.isatty():
            bom_models = load_summary_models(input_dir)
            selection = select_summary_models_window(bom_models)
            if selection is None:
                summary_models, include_summary = select_summary_models_prompt(bom_models)
            elif selection.cancelled:
                print("已取消執行。")
                return 1
            else:
                summary_models = selection.selected_models
                include_summary = selection.include_summary

        if customer:
            print(f"Customer: {customer}")
        generate_report(
            input_dir,
            output_file,
            customer=customer,
            show_progress=not args.no_progress,
            summary_models=summary_models,
            include_summary=include_summary,
        )
        print("Done.")
        return 0
    except Exception as exc:  # noqa: BLE001 - show friendly exe errors
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    finally:
        pause_for_windows_exe(pause_after_run)
