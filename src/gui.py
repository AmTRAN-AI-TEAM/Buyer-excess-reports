"""Desktop interface shared by the Python entry point and the Windows build."""

from __future__ import annotations

import ctypes
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from datetime import datetime
from pathlib import Path
from time import monotonic
from tkinter import font as tkfont
from tkinter import ttk
from typing import Callable

from PIL import Image, ImageTk

from report_errors import ReportError
from progress_timing import RemainingTimeEstimator, TimingHistory
from runner import (
    CUSTOMERS,
    PROGRESS_LABELS,
    PROGRESS_STEPS,
    ProgressEvent,
    RunResult,
    customer_input_dir,
    customer_output_file,
    generate_report,
    project_root,
)


APP_TITLE = "Excess報表整理器"
BACKGROUND = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1])) / "assets" / "app_background.png"
WHITE = "#f2f8ff"
MUTED = "#a7c9eb"
CYAN = "#55ddff"
PANEL = "#071b3e"
PHASE_LABELS = {
    "Find BOM": "尋找 BOM 檔案",
    "Find Item workbook": "尋找 Item／Price 檔案",
    "Find shortage sheet": "辨識 Shortage 工作表",
    "Find previous Excess report": "檢查上一版報表",
    "Read BOM alternates": "讀取 BOM 與替代料關係",
    "Read Item/Price data": "讀取 Item 與價格資料",
    "Read shortage data": "讀取 Shortage 資料",
    "Read previous Excess data": "讀取上一版報表與人工更新",
    "Build Excess rows": "計算 Excess 並整理報表資料",
    "Write Excel files": "寫入 Excess 與 next_time 報表",
    "Write incomplete alternates file": "寫入不完全替代料明細",
    "Done": "所有報表已完成",
}
REQUIRED_FILES = {
    "AVTC": ("Shortage", "BOM", "Item", "next_time_excess_report"),
    "RAKEN": ("Shortage", "BOM", "item information", "Price RKN", "next_time_excess_report"),
}


def enable_dpi_awareness() -> None:
    if sys.platform == "win32":
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except (AttributeError, OSError):
                pass


def open_folder(path: Path) -> None:
    """Use the OS file manager without creating a console window."""
    if sys.platform == "win32":
        os.startfile(str(path))
    else:
        command = "open" if sys.platform == "darwin" else "xdg-open"
        subprocess.Popen([command, str(path)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


class ActionButton(tk.Canvas):
    """A keyboard-accessible button matching the supplied blue artwork."""

    def __init__(
        self,
        parent,
        text: str,
        command: Callable[[], None],
        *,
        width: int,
        height: int,
        font: tuple,
        primary: bool = True,
        background: str = PANEL,
    ) -> None:
        super().__init__(
            parent, width=width, height=height, bg=background,
            highlightthickness=0, bd=0, takefocus=True, cursor="hand2",
        )
        self.command = command
        self.label = text
        self.primary = primary
        self.hovered = False
        self.pressed = False
        self.button_font = font
        self.button_width = width
        self.button_height = height
        self.bind("<Enter>", lambda _event: self._hover(True))
        self.bind("<Leave>", self._leave)
        self.bind("<ButtonPress-1>", self._press)
        self.bind("<ButtonRelease-1>", self._release)
        self.bind("<Return>", self._key_invoke)
        self.bind("<space>", self._key_invoke)
        self.bind("<FocusIn>", lambda _event: self._draw())
        self.bind("<FocusOut>", lambda _event: self._draw())
        self._draw()

    def _draw(self) -> None:
        self.delete("all")
        width, height = self.button_width, self.button_height
        radius = min(12, height // 4)
        fill = ("#229aff" if self.hovered else "#147fe3") if self.primary else ("#123e73" if self.hovered else "#0b2c58")
        if self.pressed:
            fill = "#0866bd" if self.primary else "#071f43"
        outline = CYAN if self.hovered or self.focus_get() is self else "#329fe7"
        points = [
            radius, 2, width - radius, 2, width - 2, 2, width - 2, radius,
            width - 2, height - radius, width - 2, height - 2,
            width - radius, height - 2, radius, height - 2, 2, height - 2,
            2, height - radius, 2, radius, 2, 2,
        ]
        self.create_polygon(points, smooth=True, fill=fill, outline=outline, width=2)
        self.create_text(width / 2, height / 2, text=self.label, fill=WHITE, font=self.button_font)

    def _hover(self, value: bool) -> None:
        self.hovered = value
        self._draw()

    def _leave(self, _event) -> None:
        self.pressed = False
        self._hover(False)

    def _press(self, _event) -> None:
        self.focus_set()
        self.pressed = True
        self._draw()

    def _release(self, event) -> None:
        activate = self.pressed and 0 <= event.x < self.button_width and 0 <= event.y < self.button_height
        self.pressed = False
        self._draw()
        if activate:
            self.invoke()

    def _key_invoke(self, _event) -> str:
        self.invoke()
        return "break"

    def invoke(self) -> None:
        self.command()


class AppDialog(tk.Toplevel):
    def __init__(self, app: "ExcessApp", title: str, *, kind: str, width: int = 650) -> None:
        super().__init__(app.root)
        self.app = app
        self.kind = kind
        self.title(title)
        self.configure(bg="#25669e")
        self.resizable(False, False)
        self.transient(app.root)
        self.body = tk.Frame(self, bg=PANEL, padx=app.px(28), pady=app.px(24))
        self.body.pack(fill="both", expand=True, padx=1, pady=1)
        self.body.grid_columnconfigure(0, weight=1)
        self.content_width = app.px(width - 60)
        self.width = app.px(width)
        self.primary_button: ActionButton | None = None
        self.secondary_button: ActionButton | None = None

    def label(self, text: str, *, size: int = 15, color: str = WHITE, bold: bool = False, pady=(0, 10)) -> tk.Label:
        widget = tk.Label(
            self.body, text=text, bg=PANEL, fg=color, font=self.app.font(size, bold),
            anchor="w", justify="left", wraplength=self.content_width,
        )
        widget.pack(fill="x", pady=tuple(self.app.px(value) for value in pady))
        return widget

    def buttons(self, primary_text: str, command: Callable[[], None], *, secondary_text: str | None = None, secondary_command: Callable[[], None] | None = None) -> None:
        row = tk.Frame(self.body, bg=PANEL)
        row.pack(fill="x", pady=(self.app.px(18), 0))
        self.primary_button = ActionButton(
            row, primary_text, command, width=self.app.px(150), height=self.app.px(46),
            font=self.app.font(15, True),
        )
        self.primary_button.pack(side="right")
        if secondary_text is not None and secondary_command is not None:
            self.secondary_button = ActionButton(
                row, secondary_text, secondary_command, width=self.app.px(300), height=self.app.px(46),
                font=self.app.font(14), primary=False,
            )
            self.secondary_button.pack(side="left", padx=(0, self.app.px(12)))

    def show(self, on_close: Callable[[], None]) -> None:
        self.protocol("WM_DELETE_WINDOW", on_close)
        self.bind("<Escape>", lambda _event: on_close())
        self.update_idletasks()
        height = self.winfo_reqheight()
        x = self.app.root.winfo_rootx() + (self.app.width - self.width) // 2
        y = self.app.root.winfo_rooty() + (self.app.height - height) // 2
        x = max(0, min(x, self.winfo_screenwidth() - self.width))
        y = max(30, min(y, self.winfo_screenheight() - height - 40))
        self.geometry(f"{self.width}x{height}+{x}+{y}")
        self.grab_set()
        if self.primary_button is not None:
            self.primary_button.focus_set()


class ExcessApp:
    def __init__(
        self,
        root: tk.Tk,
        *,
        app_root: Path | None = None,
        folder_opener: Callable[[Path], None] = open_folder,
    ) -> None:
        self.root = root
        self.app_root = app_root or project_root()
        self.folder_opener = folder_opener
        self.events: queue.Queue = queue.Queue()
        self.dialog: AppDialog | None = None
        self.widgets: list[tk.Widget] = []
        self.customer: str | None = None
        self.running = False
        self.exit_code = 0
        self.screen = "welcome"
        self.last_progress: ProgressEvent | None = None
        self.result: RunResult | None = None
        self.worker: threading.Thread | None = None
        self.time_estimator: RemainingTimeEstimator | None = None
        self.timing_history = TimingHistory(self.app_root / ".app_state" / "run_timings.json", PROGRESS_LABELS)
        self._eta_next_update = 0.0
        families = set(tkfont.families(root))
        self.font_family = next((name for name in ("Microsoft JhengHei UI", "Microsoft JhengHei", "Noto Sans CJK TC", "Arial") if name in families), "TkDefaultFont")
        self.log_font_family = "Consolas" if "Consolas" in families else self.font_family
        with Image.open(BACKGROUND) as source:
            ratio = source.height / source.width
            self.width = int(min(1200, root.winfo_screenwidth() * 0.92, root.winfo_screenheight() * 0.86 / ratio))
            self.height = round(self.width * ratio)
            self.scale = self.width / 1200
            self.background = ImageTk.PhotoImage(source.resize((self.width, self.height), Image.Resampling.LANCZOS), master=root)
        root.title(APP_TITLE)
        root.configure(bg="#03102d")
        root.resizable(False, False)
        x = max(0, (root.winfo_screenwidth() - self.width) // 2)
        y = max(30, (root.winfo_screenheight() - self.height) // 2 - 20)
        root.geometry(f"{self.width}x{self.height}+{x}+{y}")
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.report_callback_exception = self._callback_error
        self.canvas = tk.Canvas(root, width=self.width, height=self.height, highlightthickness=0, bg="#03102d")
        self.canvas.pack(fill="both", expand=True)
        style = ttk.Style(root)
        style.theme_use("clam")
        style.configure("Excess.Horizontal.TProgressbar", troughcolor="#092348", background="#32cfff", bordercolor="#173f72", lightcolor="#32cfff", darkcolor="#32cfff")
        style.configure("Excess.Vertical.TScrollbar", troughcolor=PANEL, background="#224c7a", bordercolor=PANEL, arrowcolor=MUTED)
        self.show_welcome()
        self._poll_id = root.after(60, self._poll_events)

    def px(self, value: float) -> int:
        return max(1, round(value * self.scale)) if value else 0

    def font(self, size: int, bold: bool = False) -> tuple:
        return (self.font_family, -max(10, self.px(size)), "bold" if bold else "normal")

    def _text(self, x: int, y: int, text: str, *, size: int = 16, color: str = WHITE, bold: bool = False, anchor="center") -> int:
        return self.canvas.create_text(self.px(x), self.px(y), text=text, fill=color, font=self.font(size, bold), anchor=anchor)

    def _reset(self, screen: str) -> None:
        for widget in self.widgets:
            widget.destroy()
        self.widgets.clear()
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self.background)
        self.screen = screen

    def _button(self, x: int, y: int, text: str, command: Callable[[], None], *, width=230, height=58, size=20, primary=True) -> ActionButton:
        button = ActionButton(
            self.canvas, text, command, width=self.px(width), height=self.px(height),
            font=self.font(size, True), primary=primary, background="#0843a2",
        )
        self.widgets.append(button)
        self.canvas.create_window(self.px(x), self.px(y), window=button)
        return button

    def show_welcome(self) -> None:
        self._reset("welcome")
        self._text(600, 338, "BUYER EXCESS REPORTS", size=14, color=CYAN, bold=True)
        self._text(600, 401, APP_TITLE, size=46, bold=True)
        self._text(600, 453, "將來源資料整理成清楚、完整的 Excess 報表", size=18, color=MUTED)
        self.start_button = self._button(600, 524, "Start  →", self.show_factories, width=250, height=62, size=23)
        self._text(600, 596, "選擇工廠　・　確認檔案　・　產出報表", size=13, color=MUTED)
        self.start_button.focus_set()

    def show_factories(self) -> None:
        self._reset("factory")
        self._text(600, 330, "STEP 01  /  選擇工廠", size=14, color=CYAN, bold=True)
        self._text(600, 383, "請選擇所屬工廠", size=35, bold=True)
        self._text(600, 426, "選擇後，請依提示確認本次需要的來源檔案", size=17, color=MUTED)
        self.customer_buttons = {}
        for x, customer in zip((454, 746), CUSTOMERS):
            self.customer_buttons[customer] = self._button(x, 501, customer, lambda value=customer: self.choose_customer(value), width=248, height=72, size=26)
        self._text(600, 566, "AVTC 與 RAKEN 的輸入、輸出資料夾分開存放", size=14, color=MUTED)
        self._button(147, 612, "← 返回", self.show_welcome, width=110, height=36, size=13, primary=False)

    def _dismiss_dialog(self) -> None:
        if self.dialog is not None:
            self.dialog.grab_release()
            self.dialog.destroy()
            self.dialog = None

    def choose_customer(self, customer: str) -> None:
        if self.running or self.dialog is not None:
            return
        self.customer = customer
        dialog = self.dialog = AppDialog(self, "確認來源檔案", kind="checklist")
        dialog.label(f"{customer}  /  來源檔案確認", size=13, color=CYAN, bold=True)
        dialog.label("請確認以下檔案已放入", size=25, bold=True, pady=(0, 6))
        dialog.label(f"input/{customer}/", size=16, color=CYAN, pady=(0, 16))
        for filename in REQUIRED_FILES[customer]:
            optional = filename == "next_time_excess_report"
            dialog.label(f"•  {filename}" + ("（可選）" if optional else ""), size=17, color=MUTED if optional else WHITE, pady=(0, 8))
        dialog.label("BOM／Item／Price 依檔名關鍵字辨識；Shortage 分頁名稱須包含 shortage。\n上一版檔名須包含 next_time；未提供時仍可產出本期報表。", size=12, color=MUTED, pady=(10, 0))
        self.checklist_status = dialog.label("補齊檔案後，請回到此視窗按「確認無誤」。", size=12, color=MUTED, pady=(12, 0))
        dialog.buttons("確認無誤", self.start_run, secondary_text="有缺漏，打開input資料夾", secondary_command=self.open_input_folder)
        dialog.show(self._dismiss_dialog)

    def open_input_folder(self) -> None:
        if self.customer is None:
            return
        path = customer_input_dir(self.app_root, self.customer)
        try:
            path.mkdir(parents=True, exist_ok=True)
            self.folder_opener(path)
            self.checklist_status.configure(text="已開啟 input 資料夾。補齊後，請按「確認無誤」。", fg=CYAN)
        except Exception as exc:
            self.checklist_status.configure(text=f"無法開啟資料夾：{path}\n{exc}", fg="#ffb7b7")

    def start_run(self) -> None:
        if self.running or self.customer is None:
            return
        self._dismiss_dialog()
        self.running = True
        self._show_running()
        customer = self.customer
        self.worker = threading.Thread(target=self._run_worker, args=(customer,), daemon=True, name="excess-report-worker")
        self.worker.start()

    def _show_running(self) -> None:
        self._reset("running")
        self.time_estimator = RemainingTimeEstimator(PROGRESS_LABELS, self.timing_history.baseline(self.customer))
        self._text(90, 305, "STEP 02  /  整理報表", size=13, color=CYAN, bold=True, anchor="w")
        self.title_item = self._text(90, 338, "正在整理報表", size=29, bold=True, anchor="w")
        self._text(1108, 330, f"工廠  {self.customer}", size=17, color=CYAN, bold=True, anchor="e")
        self.phase_item = self._text(90, 375, "準備讀取來源檔案…", size=16, anchor="w")
        self.progressbar = ttk.Progressbar(self.canvas, style="Excess.Horizontal.TProgressbar", maximum=PROGRESS_STEPS, mode="determinate")
        self.widgets.append(self.progressbar)
        self.canvas.create_window(self.px(90), self.px(400), anchor="nw", width=self.px(1020), height=self.px(12), window=self.progressbar)
        self.count_item = self._text(90, 430, f"已完成 0 / {PROGRESS_STEPS} 個階段", size=14, color=MUTED, anchor="w")
        self.percent_item = self._text(1110, 430, "0%  ｜  預估剩餘：估算中", size=16, color=CYAN, bold=True, anchor="e")
        frame = tk.Frame(self.canvas, bg=PANEL, highlightbackground="#2871aa", highlightthickness=1)
        self.widgets.append(frame)
        self.canvas.create_window(self.px(90), self.px(450), anchor="nw", width=self.px(1020), height=self.px(158), window=frame)
        self.log_text = tk.Text(
            frame, bg=PANEL, fg="#d1e9ff", insertbackground=WHITE, relief="flat", bd=0,
            font=(self.log_font_family, -max(12, self.px(16))), wrap="word", state="disabled",
            padx=self.px(12), pady=self.px(9), takefocus=True,
            spacing1=self.px(1), spacing3=self.px(2),
        )
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.log_text.yview, style="Excess.Vertical.TScrollbar")
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        self.footer_item = self._text(600, 625, "請保持程式開啟，完成後會自動提示取檔。", size=12, color=MUTED)
        self._append_log(f"工廠：{self.customer}　開始整理報表。")

    def _run_worker(self, customer: str) -> None:
        try:
            result = generate_report(
                customer_input_dir(self.app_root, customer), customer_output_file(self.app_root, customer),
                customer=customer, on_event=self.events.put,
            )
        except Exception as exc:
            self.events.put(("error", exc))
        else:
            self.events.put(("success", result))

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{datetime.now():%H:%M:%S}] {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _update_eta_label(self, *, force: bool = False) -> None:
        if self.screen != "running" or self.time_estimator is None:
            return
        now = monotonic()
        if not force and now < self._eta_next_update:
            return
        self._eta_next_update = now + 1.0
        current = self.last_progress.current if self.last_progress else 0
        total = self.last_progress.total if self.last_progress else PROGRESS_STEPS
        percentage = current * 100 // max(total, 1)
        self.canvas.itemconfigure(self.percent_item, text=f"{percentage}%  ｜  {self.time_estimator.snapshot().text}")

    def _handle_event(self, event) -> None:
        if isinstance(event, ProgressEvent):
            if event.kind == "message":
                self._append_log(event.message)
            else:
                self.last_progress = event
                self.progressbar["value"] = event.current
                label = PHASE_LABELS.get(event.label, event.label)
                self.canvas.itemconfigure(self.phase_item, text=label)
                self.canvas.itemconfigure(self.count_item, text=f"已完成 {event.current} / {event.total} 個階段")
                self.time_estimator.observe(event.label, event.elapsed_seconds, event.stage_durations)
                self._update_eta_label(force=True)
                self._append_log(label)
            return
        kind, payload = event
        self.running = False
        self.time_estimator.stop(success=kind == "success")
        self._update_eta_label(force=True)
        if kind == "success":
            self.result = payload
            try:
                self.timing_history.record(self.customer, self.time_estimator.durations)
            except OSError as exc:
                self._append_log(f"報表已完成；無法保存耗時紀錄，下次將重新估算：{exc}")
            self._show_success()
        else:
            self.exit_code = 1
            self.canvas.itemconfigure(self.title_item, text="報表處理失敗", fill="#ffb8bd")
            self.canvas.itemconfigure(self.footer_item, text="請查看錯誤原因，按「確認」後將關閉程式。")
            self._append_log(str(payload))
            self._show_error(payload)

    def _poll_events(self) -> None:
        try:
            # Bound each batch so even a busy worker cannot starve the UI.
            for _ in range(100):
                self._handle_event(self.events.get_nowait())
        except queue.Empty:
            pass
        self._update_eta_label()
        self._poll_id = self.root.after(60, self._poll_events)

    def _show_success(self) -> None:
        self.canvas.itemconfigure(self.title_item, text="報表整理完成")
        self.canvas.itemconfigure(self.footer_item, text="按「確認」後，將開啟 output 資料夾並關閉程式。")
        dialog = self.dialog = AppDialog(self, "報表整理完成", kind="success", width=600)
        dialog.label("COMPLETED", size=13, color=CYAN, bold=True)
        dialog.label("已完成，打開output資料夾取檔案", size=23, bold=True, pady=(0, 16))
        dialog.label(f"output/{self.customer}/", size=16, color=CYAN)
        dialog.label("excess_report.xlsx\nnext_time_excess_report.xlsx\nincomplete_alternates.xlsx", size=15, pady=(0, 12))
        dialog.label("按「確認」後開啟資料夾，並關閉此程式。", size=13, color=MUTED)
        dialog.buttons("確認", self._finish_success)
        dialog.show(self._finish_success)

    def _finish_success(self) -> None:
        path = customer_output_file(self.app_root, self.customer).parent
        try:
            self.folder_opener(path)
        except Exception as exc:
            self._dismiss_dialog()
            self.exit_code = 1
            self._show_error(ReportError(path, "報表已完成，但無法開啟 output 資料夾", exc))
        else:
            self._destroy()

    def _show_error(self, error: Exception) -> None:
        self._dismiss_dialog()
        dialog = self.dialog = AppDialog(self, "報表處理失敗", kind="error", width=680)
        dialog.label("PROCESSING STOPPED", size=13, color="#ffb8bd", bold=True)
        dialog.label("無法完成本次報表", size=25, bold=True)
        if isinstance(error, ReportError):
            message = error.user_message()
        else:
            phase = self.last_progress.label if self.last_progress else "啟動介面"
            message = ReportError(None, PHASE_LABELS.get(phase, phase), error).user_message()
        # A scrollable message keeps even long paths/errors inside the screen.
        details = tk.Text(
            dialog.body, bg=PANEL, fg=WHITE, relief="flat", bd=0,
            font=self.font(15), wrap="word", height=10, width=1,
            padx=0, pady=0, highlightthickness=0,
        )
        details.insert("1.0", message)
        details.configure(state="disabled")
        details.pack(fill="x", pady=(0, self.px(10)))
        dialog.label("按「確認」後，將關閉程式。", size=13, color=MUTED)
        dialog.buttons("確認", self._destroy)
        dialog.show(self._destroy)

    def _callback_error(self, _exc_type, value: Exception, _traceback) -> None:
        self.exit_code = 1
        self._show_error(value)

    def close(self) -> None:
        if self.running:
            self.root.bell()
            self.canvas.itemconfigure(self.footer_item, text="報表仍在處理中，請等待完成後再關閉。")
            return
        self._destroy()

    def _destroy(self) -> None:
        if self.worker is not None and not self.running:
            self.worker.join(timeout=1)
        if getattr(self, "_poll_id", None):
            self.root.after_cancel(self._poll_id)
        self._dismiss_dialog()
        self.root.destroy()


def main() -> int:
    enable_dpi_awareness()
    root = tk.Tk()
    root.withdraw()
    try:
        app = ExcessApp(root)
    except Exception as exc:
        from tkinter import messagebox
        messagebox.showerror(APP_TITLE, f"無法啟動介面：\n{exc}", parent=root)
        root.destroy()
        return 1
    root.deiconify()
    root.mainloop()
    return app.exit_code
