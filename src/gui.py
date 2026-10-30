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
from tkinter import font as tkfont
from tkinter import ttk
from typing import Callable

from PIL import Image, ImageTk

from report_errors import ReportError
from source_check import SourceCheck, check_sources, input_signature, source_requirements
from runner import (
    CUSTOMERS,
    ProgressEvent,
    RunResult,
    customer_input_dir,
    customer_output_file,
    generate_report,
    load_summary_models,
    project_root,
    progress_steps_for_customer,
)


APP_TITLE = "Excess報表整理器"
BACKGROUND = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[1])) / "assets" / "app_background.png"
INK = "#422715"
MUTED = "#6b452b"
ACCENT = "#704016"
PANEL = "#fff2db"
LIGHT = "#fff7e8"
SUCCESS = "#28613b"
ERROR = "#9b2727"
WARNING = "#a86312"
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
    "Write Excess summary": "產生 excess_summary.xlsx",
    "Write incomplete alternates file": "寫入不完全替代料明細",
    "Done": "所有報表已完成",
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
    """A keyboard-accessible button with a real disabled state."""

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
        backdrop: ImageTk.PhotoImage | None = None,
    ) -> None:
        super().__init__(
            parent, width=width, height=height, bg=background,
            highlightthickness=0, bd=0, takefocus=True, cursor="hand2",
        )
        self.command = command
        self.label = text
        self.primary = primary
        self.enabled = True
        self.backdrop = backdrop
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
        if self.backdrop is not None:
            self.create_image(0, 0, image=self.backdrop, anchor="nw")
        width, height = self.button_width, self.button_height
        radius = min(12, height // 4)
        fill = ("#865024" if self.hovered else "#683c1c") if self.primary else ("#f5d7a6" if self.hovered else "#edc58d")
        if self.pressed:
            fill = "#482610" if self.primary else "#d9ac70"
        outline = "#e2b96d" if self.hovered or self.focus_get() is self else "#a36c32"
        text_color = LIGHT if self.primary else INK
        if not self.enabled:
            fill, outline, text_color = "#b8afa0", "#958c80", "#615b53"
        points = [
            radius, 2, width - radius, 2, width - 2, 2, width - 2, radius,
            width - 2, height - radius, width - 2, height - 2,
            width - radius, height - 2, radius, height - 2, 2, height - 2,
            2, height - radius, 2, radius, 2, 2,
        ]
        self.create_polygon(points, smooth=True, fill=fill, outline=outline, width=2)
        self.create_text(width / 2, height / 2, text=self.label, fill=text_color, font=self.button_font)

    def set_state(self, text: str, *, enabled: bool) -> None:
        self.label = text
        self.enabled = enabled
        self.hovered = self.pressed = False
        self.configure(cursor="hand2" if enabled else "arrow", takefocus=enabled)
        self._draw()

    def _hover(self, value: bool) -> None:
        self.hovered = value and self.enabled
        self._draw()

    def _leave(self, _event) -> None:
        self.pressed = False
        self._hover(False)

    def _press(self, _event) -> None:
        if not self.enabled:
            return
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
        if self.enabled:
            self.command()


class AppDialog(tk.Toplevel):
    def __init__(self, app: "ExcessApp", title: str, *, kind: str, width: int = 760) -> None:
        super().__init__(app.root)
        self.app = app
        self.kind = kind
        self.title(title)
        self.configure(bg="#a36c32")
        self.resizable(False, False)
        self.transient(app.root)
        self.body = tk.Frame(self, bg=PANEL, padx=app.px(28), pady=app.px(24))
        self.body.pack(fill="both", expand=True, padx=1, pady=1)
        self.body.grid_columnconfigure(0, weight=1)
        self.content_width = app.px(width - 60)
        self.width = app.px(width)
        self.primary_button: ActionButton | None = None

    def label(self, text: str, *, size: int = 20, color: str = INK, bold: bool = False, pady=(0, 12)) -> tk.Label:
        widget = tk.Label(
            self.body, text=text, bg=PANEL, fg=color, font=self.app.font(size, bold),
            anchor="w", justify="left", wraplength=self.content_width,
        )
        widget.pack(fill="x", pady=tuple(self.app.px(value) for value in pady))
        return widget

    def buttons(self, primary_text: str, command: Callable[[], None]) -> None:
        row = tk.Frame(self.body, bg=PANEL)
        row.pack(fill="x", pady=(self.app.px(18), 0))
        self.primary_button = ActionButton(
            row, primary_text, command, width=self.app.px(240), height=self.app.px(58),
            font=self.app.font(21, True),
        )
        self.primary_button.pack()

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


class ModelSelectionDialog(tk.Toplevel):
    def __init__(
        self,
        app: "ExcessApp",
        *,
        customer: str,
        initial_models: tuple[str, ...] | None,
        on_confirm: Callable[[tuple[str, ...] | None], None],
        on_skip: Callable[[], None],
        on_back: Callable[[], None],
    ) -> None:
        super().__init__(app.root)
        self.app = app
        self.on_confirm = on_confirm
        self.on_skip = on_skip
        self.on_back = on_back
        self.initial_models = initial_models
        self.models: tuple[str, ...] = ()
        self.model_vars: dict[str, tk.BooleanVar] = {}
        self.model_buttons: dict[str, tk.Checkbutton] = {}

        self.title("選擇加單建議 MODEL")
        self.configure(bg="#a36c32")
        self.resizable(True, True)
        self.minsize(app.px(820), app.px(560))
        self.transient(app.root)
        self.protocol("WM_DELETE_WINDOW", on_back)
        self.bind("<Escape>", lambda _event: on_back())

        body = tk.Frame(self, bg=PANEL, padx=app.px(24), pady=app.px(20))
        body.pack(fill="both", expand=True, padx=1, pady=1)
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(3, weight=1)

        tk.Label(
            body,
            text=f"{customer}：選擇要產出的加單建議 MODEL",
            bg=PANEL,
            fg=INK,
            font=app.font(28, True),
            anchor="w",
        ).grid(row=0, column=0, sticky="ew")
        tk.Label(
            body,
            text="選項來自 BOM 的 MODEL 欄。新工作表只列出勾選 MODEL 的加單建議。",
            bg=PANEL,
            fg=MUTED,
            font=app.font(16),
            anchor="w",
        ).grid(row=1, column=0, sticky="ew", pady=(app.px(6), app.px(12)))

        controls = tk.Frame(body, bg=PANEL)
        controls.grid(row=2, column=0, sticky="ew", pady=(0, app.px(8)))
        controls.grid_columnconfigure(0, weight=1)
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(
            controls,
            textvariable=self.search_var,
            font=app.font(17),
            relief="solid",
            bd=1,
        )
        self.search_entry.grid(row=0, column=0, sticky="ew", padx=(0, app.px(12)), ipady=app.px(7))
        self.search_entry.configure(state="disabled")
        self.search_placeholder = "搜尋 MODEL…"
        self.search_var.set(self.search_placeholder)
        self.search_entry.bind("<FocusIn>", self._clear_search_placeholder)
        self.search_entry.bind("<FocusOut>", self._restore_search_placeholder)

        self.all_var = tk.BooleanVar(value=True)
        self.all_checkbox = tk.Checkbutton(
            controls,
            text="全選所有 MODEL",
            variable=self.all_var,
            command=self._toggle_all,
            bg=PANEL,
            activebackground=PANEL,
            fg=INK,
            selectcolor=LIGHT,
            font=app.font(16, True),
            state="disabled",
        )
        self.all_checkbox.grid(row=0, column=1, sticky="e")

        list_frame = tk.Frame(body, bg=PANEL, highlightbackground="#b98952", highlightthickness=1)
        list_frame.grid(row=3, column=0, sticky="nsew")
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)
        self.list_canvas = tk.Canvas(list_frame, bg=PANEL, highlightthickness=0)
        self.list_canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(
            list_frame,
            orient="vertical",
            command=self.list_canvas.yview,
            style="Excess.Vertical.TScrollbar",
        )
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.list_canvas.configure(yscrollcommand=scrollbar.set)
        self.list_inner = tk.Frame(self.list_canvas, bg=PANEL)
        self.list_window = self.list_canvas.create_window(
            (0, 0), window=self.list_inner, anchor="nw"
        )
        self.list_inner.bind(
            "<Configure>",
            lambda _event: self.list_canvas.configure(
                scrollregion=self.list_canvas.bbox("all")
            ),
        )
        self.list_canvas.bind(
            "<Configure>",
            lambda event: self.list_canvas.itemconfigure(
                self.list_window,
                width=event.width,
            ),
        )
        self.list_canvas.bind(
            "<MouseWheel>",
            lambda event: self.list_canvas.yview_scroll(
                -int(event.delta / 120), "units"
            ),
        )

        footer = tk.Frame(body, bg=PANEL)
        footer.grid(row=4, column=0, sticky="ew", pady=(app.px(10), 0))
        self.count_label = tk.Label(
            footer,
            text="正在讀取 BOM MODEL…",
            bg=PANEL,
            fg=MUTED,
            font=app.font(15),
            anchor="w",
        )
        self.count_label.pack(side="left", fill="x", expand=True)
        button_row = tk.Frame(footer, bg=PANEL)
        button_row.pack(side="right")
        self.back_button = ActionButton(
            button_row,
            "返回",
            on_back,
            width=app.px(105),
            height=app.px(48),
            font=app.font(16, True),
            primary=False,
        )
        self.back_button.pack(side="left", padx=(0, app.px(8)))
        self.skip_button = ActionButton(
            button_row,
            "略過此頁，不產出summary",
            on_skip,
            width=app.px(260),
            height=app.px(48),
            font=app.font(15, True),
            primary=False,
        )
        self.skip_button.pack(side="left", padx=(0, app.px(8)))
        self.confirm_button = ActionButton(
            button_row,
            "確認選取",
            self._confirm,
            width=app.px(155),
            height=app.px(48),
            font=app.font(16, True),
        )
        self.confirm_button.set_state("讀取中…", enabled=False)
        self.confirm_button.pack(side="left")

        self.search_var.trace_add("write", self._filter_models)
        self.update_idletasks()
        width, height = app.px(980), app.px(650)
        x = app.root.winfo_rootx() + (app.width - width) // 2
        y = app.root.winfo_rooty() + (app.height - height) // 2
        x = max(0, min(x, self.winfo_screenwidth() - width))
        y = max(30, min(y, self.winfo_screenheight() - height - 40))
        self.geometry(f"{width}x{height}+{x}+{y}")
        self.grab_set()

    def set_models(self, models: tuple[str, ...] | None, error: Exception | None = None) -> None:
        if error is not None:
            self.count_label.configure(
                text=f"讀取 BOM MODEL 失敗：{error}",
                fg=ERROR,
            )
            self.confirm_button.set_state("無 MODEL 清單", enabled=False)
            self.all_checkbox.configure(state="disabled")
            return
        self.models = models or ()
        if not self.models:
            self.count_label.configure(text="BOM 中沒有可選的非空 MODEL。", fg=WARNING)
            self.confirm_button.set_state("無 MODEL 清單", enabled=False)
            self.all_checkbox.configure(state="disabled")
            return

        initially_selected = (
            set(self.models)
            if self.initial_models is None
            else set(self.initial_models)
        )
        self.model_vars.clear()
        self.model_buttons.clear()
        for index, model in enumerate(self.models):
            variable = tk.BooleanVar(value=model in initially_selected)
            button = tk.Checkbutton(
                self.list_inner,
                text=model,
                variable=variable,
                command=self._sync_selection,
                bg=PANEL,
                activebackground=PANEL,
                fg=INK,
                selectcolor=LIGHT,
                font=self.app.font(15),
                anchor="w",
                justify="left",
                wraplength=self.app.px(405),
            )
            button.grid(row=index // 2, column=index % 2, sticky="ew", padx=self.app.px(8), pady=self.app.px(3))
            self.model_vars[model] = variable
            self.model_buttons[model] = button
        self.list_inner.grid_columnconfigure(0, weight=1)
        self.list_inner.grid_columnconfigure(1, weight=1)
        self.search_entry.configure(state="normal")
        self.search_entry.delete(0, "end")
        self.search_entry.insert(0, self.search_placeholder)
        self.all_checkbox.configure(state="normal")
        self._filter_models()
        self._sync_selection()
        self.search_entry.focus_set()

    def _clear_search_placeholder(self, _event=None) -> None:
        if self.search_var.get() == self.search_placeholder:
            self.search_var.set("")

    def _restore_search_placeholder(self, _event=None) -> None:
        if not self.search_var.get():
            self.search_var.set(self.search_placeholder)

    def _filter_models(self, *_args) -> None:
        query = self.search_var.get().strip().casefold()
        if query == self.search_placeholder.casefold():
            query = ""
        visible = [model for model in self.models if query in model.casefold()]
        for model, button in self.model_buttons.items():
            button.grid_remove()
        for index, model in enumerate(visible):
            self.model_buttons[model].grid(
                row=index // 2,
                column=index % 2,
                sticky="ew",
                padx=self.app.px(8),
                pady=self.app.px(3),
            )
        self.list_canvas.configure(scrollregion=self.list_canvas.bbox("all"))

    def _toggle_all(self) -> None:
        selected = self.all_var.get()
        for variable in self.model_vars.values():
            variable.set(selected)
        self._sync_selection()

    def _sync_selection(self) -> None:
        selected_count = sum(variable.get() for variable in self.model_vars.values())
        self.all_var.set(bool(self.models) and selected_count == len(self.models))
        self.count_label.configure(
            text=f"已選取 {selected_count} / {len(self.models)} 個 MODEL",
            fg=MUTED,
        )
        self.confirm_button.set_state(
            "確認選取",
            enabled=selected_count > 0,
        )

    def _confirm(self) -> None:
        selected = tuple(
            model for model in self.models if self.model_vars[model].get()
        )
        self.grab_release()
        self.destroy()
        self.on_confirm(None if len(selected) == len(self.models) else selected)


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
        self.dialog: tk.Toplevel | None = None
        self.widgets: list[tk.Widget] = []
        self.customer: str | None = None
        self.summary_models: tuple[str, ...] | None = None
        self.include_summary = True
        self.running = False
        self.exit_code = 0
        self.screen = "welcome"
        self.last_progress: ProgressEvent | None = None
        self.result: RunResult | None = None
        self.worker: threading.Thread | None = None
        self._check_token = 0
        self._check_busy = False
        self._model_token = 0
        self.source_checks: tuple[SourceCheck, ...] = ()
        families = set(tkfont.families(root))
        self.font_family = next((name for name in ("Microsoft JhengHei UI", "Microsoft JhengHei", "Noto Sans CJK TC", "Arial") if name in families), "TkDefaultFont")
        self.log_font_family = "Consolas" if "Consolas" in families else self.font_family
        with Image.open(BACKGROUND) as source:
            ratio = source.height / source.width
            self.width = int(min(1440, root.winfo_screenwidth() * 0.92, root.winfo_screenheight() * 0.86 / ratio))
            self.height = round(self.width * ratio)
            self.scale = self.width / 1200
            self.background_image = source.resize((self.width, self.height), Image.Resampling.LANCZOS)
            self.background = ImageTk.PhotoImage(self.background_image, master=root)
        root.title(APP_TITLE)
        root.configure(bg=INK)
        root.resizable(False, False)
        x = max(0, (root.winfo_screenwidth() - self.width) // 2)
        y = max(30, (root.winfo_screenheight() - self.height) // 2 - 20)
        root.geometry(f"{self.width}x{self.height}+{x}+{y}")
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.report_callback_exception = self._callback_error
        self.canvas = tk.Canvas(root, width=self.width, height=self.height, highlightthickness=0, bg=INK)
        self.canvas.pack(fill="both", expand=True)
        style = ttk.Style(root)
        style.theme_use("clam")
        style.configure("Excess.Horizontal.TProgressbar", troughcolor="#e7c795", background="#875322", bordercolor="#a36c32", lightcolor="#a5753c", darkcolor="#683c1c")
        style.configure("Excess.Vertical.TScrollbar", troughcolor=PANEL, background="#d9b77e", bordercolor=PANEL, arrowcolor=MUTED)
        root.bind("<Escape>", self._go_back)
        self.show_welcome()
        self._poll_id = root.after(60, self._poll_events)

    def px(self, value: float) -> int:
        return max(1, round(value * self.scale)) if value else 0

    def font(self, size: int, bold: bool = False) -> tuple:
        return (self.font_family, -max(10, self.px(size)), "bold" if bold else "normal")

    def _text(self, x: int, y: int, text: str, *, size: int = 20, color: str = INK, bold: bool = False, anchor="center") -> int:
        return self.canvas.create_text(self.px(x), self.px(y), text=text, fill=color, font=self.font(size, bold), anchor=anchor)

    def _reset(self, screen: str) -> None:
        self._check_token += 1
        self._check_busy = False
        for widget in self.widgets:
            widget.destroy()
        self.widgets.clear()
        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor="nw", image=self.background)
        self.screen = screen

    def _button(self, x: int, y: int, text: str, command: Callable[[], None], *, width=290, height=74, size=26, primary=True) -> ActionButton:
        button_width, button_height = self.px(width), self.px(height)
        left, top = self.px(x) - button_width // 2, self.px(y) - button_height // 2
        backdrop = ImageTk.PhotoImage(self.background_image.crop((left, top, left + button_width, top + button_height)), master=self.root)
        button = ActionButton(
            self.canvas, text, command, width=button_width, height=button_height,
            font=self.font(size, True), primary=primary, backdrop=backdrop,
        )
        self.widgets.append(button)
        self.canvas.create_window(self.px(x), self.px(y), window=button)
        return button

    def show_welcome(self) -> None:
        self._reset("welcome")
        self._text(600, 350, APP_TITLE, size=56, bold=True)
        self.start_button = self._button(600, 455, "Start  →", self.show_factories, width=310, height=78, size=29)
        self._text(600, 550, "選擇工廠　・　確認檔案　・　產出報表", size=18, color=MUTED)
        self.start_button.focus_set()

    def show_factories(self) -> None:
        self._reset("factory")
        self._text(600, 266, "STEP 01  /  選擇工廠", size=18, color=ACCENT, bold=True)
        self._text(600, 324, "請選擇所屬工廠", size=44, bold=True)
        self._text(600, 378, "選擇後，將自動檢查本次需要的來源檔案", size=21, color=MUTED)
        self.customer_buttons = {}
        for x, customer in zip((435, 765), CUSTOMERS):
            self.customer_buttons[customer] = self._button(x, 464, customer, lambda value=customer: self.choose_customer(value), width=280, height=86, size=32)
        self._text(600, 551, "AVTC 與 RAKEN 的輸入、輸出資料夾分開存放", size=18, color=MUTED)
        self._button(135, 598, "← 返回", self.show_welcome, width=124, height=42, size=17, primary=False)

    def _go_back(self, _event) -> None:
        if self.dialog is not None or self.running:
            return
        if self.screen == "checklist":
            self.show_factories()
        elif self.screen == "factory":
            self.show_welcome()

    def _dismiss_dialog(self) -> None:
        if self.dialog is not None:
            try:
                self.dialog.grab_release()
            except tk.TclError:
                pass
            try:
                self.dialog.destroy()
            except tk.TclError:
                pass
            self.dialog = None

    def _draw_source_icon(self, key: str, y: int, status: str) -> None:
        """Draw filled status circles and white symbols without emoji fonts."""
        tag = f"source-status-{key}"
        self.canvas.delete(tag)
        x, y, radius = self.px(85), self.px(y), self.px(12)
        color = {"found": SUCCESS, "error": ERROR, "warning": WARNING, "checking": "#887b6c"}[status]
        self.canvas.create_oval(x - radius, y - radius, x + radius, y + radius,
                                fill=color, outline="", tags=tag)

        def line(*points: float) -> None:
            self.canvas.create_line(*points, fill=LIGHT, width=self.px(2.5),
                                    capstyle="round", joinstyle="round", tags=tag)

        if status == "found":
            line(x - radius * .48, y, x - radius * .12, y + radius * .35,
                 x + radius * .48, y - radius * .38)
        elif status == "error":
            offset = radius * .36
            line(x - offset, y - offset, x + offset, y + offset)
            line(x - offset, y + offset, x + offset, y - offset)
        elif status == "warning":
            line(x, y - radius * .48, x, y + radius * .08)
            dot = self.px(1.5)
            center_y = y + radius * .43
            self.canvas.create_oval(x - dot, center_y - dot, x + dot, center_y + dot,
                                    fill=LIGHT, outline="", tags=tag)
        else:
            dot = self.px(1.5)
            for offset in (-radius * .4, 0, radius * .4):
                self.canvas.create_oval(x + offset - dot, y - dot, x + offset + dot, y + dot,
                                        fill=LIGHT, outline="", tags=tag)

    def choose_customer(self, customer: str) -> None:
        if self.running or self.dialog is not None:
            return
        self.customer = customer
        self.summary_models = None
        self.include_summary = True
        self._show_model_selector(customer)

    def _show_model_selector(
        self,
        customer: str,
        *,
        initial_models: tuple[str, ...] | None = None,
    ) -> None:
        self._model_token += 1
        token = self._model_token
        dialog = ModelSelectionDialog(
            self,
            customer=customer,
            initial_models=initial_models,
            on_confirm=self._summary_models_selected,
            on_skip=self._skip_summary_selection,
            on_back=self._back_from_model_selection,
        )
        self.dialog = dialog
        self._model_dialog = dialog
        input_dir = customer_input_dir(self.app_root, customer)
        events = self.events

        def load_models() -> None:
            try:
                models = load_summary_models(input_dir)
            except Exception as exc:
                events.put(("models", (token, None, exc)))
            else:
                events.put(("models", (token, models, None)))

        threading.Thread(
            target=load_models,
            daemon=True,
            name="excess-model-loader",
        ).start()

    def _receive_models(self, payload) -> None:
        token, models, error = payload
        if token != self._model_token or self.dialog is not getattr(self, "_model_dialog", None):
            return
        self._model_dialog.set_models(models, error)

    def _summary_models_selected(self, models: tuple[str, ...] | None) -> None:
        self.dialog = None  # The selector closes itself before this callback.
        self._model_token += 1
        self.summary_models = models
        self.include_summary = True
        if self.customer is not None:
            self._start_checklist(self.customer)

    def _skip_summary_selection(self) -> None:
        self._dismiss_dialog()
        self._model_token += 1
        self.summary_models = None
        self.include_summary = False
        if self.customer is not None:
            self._start_checklist(self.customer)

    def _back_from_model_selection(self) -> None:
        self._dismiss_dialog()
        self._model_token += 1
        self.customer = None
        self.show_factories()

    def _return_to_model_selection(self) -> None:
        if self.running or self.dialog is not None:
            return
        if self.customer is None:
            self.show_factories()
            return
        self._show_model_selector(self.customer, initial_models=self.summary_models)

    def _start_checklist(self, customer: str) -> None:
        self._reset("checklist")
        self.source_checks = ()
        self._text(78, 239, "來源檔案確認", size=32, bold=True, anchor="w")
        self._text(1122, 240, f"input/{customer}/", size=20, color=ACCENT, bold=True, anchor="e")
        self.checklist_rows = {}
        for index, requirement in enumerate(source_requirements(customer)):
            y = 289 + index * 46
            self._draw_source_icon(requirement.key, y, "checking")
            name = self._text(112, y, requirement.label, size=21, bold=True, anchor="w")
            detail = self._text(440, y, "正在檢查…", size=17, color=MUTED, anchor="w")
            self.canvas.create_line(self.px(78), self.px(y + 23), self.px(1122), self.px(y + 23), fill="#b98952")
            self.checklist_rows[requirement.key] = (name, detail, y)
        self._text(600, 518, "支援 .xlsx／.xlsm；BOM、Item、Price 依檔名辨識，Shortage 依分頁名稱辨識。", size=15, color=MUTED)
        self.checklist_status = self._text(600, 549, "正在檢查來源檔案，請稍候…", size=18, color=ACCENT)
        self.checklist_button = self._button(480, 596, "正在檢查來源檔案…", self.start_run, width=440, height=54, size=22)
        self.checklist_button.set_state("正在檢查來源檔案…", enabled=False)
        self.rescan_button = self._button(830, 596, "重新掃描檔案", self._refresh_checklist,
                                         width=220, height=54, size=22, primary=False)
        self._button(135, 596, "← 返回", self._return_to_model_selection, width=124, height=42, size=17, primary=False)
        self._refresh_checklist()

    def _refresh_checklist(self, *, start_when_ready: bool = False) -> None:
        if self.screen != "checklist" or self._check_busy:
            return
        path = customer_input_dir(self.app_root, self.customer)
        signature = input_signature(path)
        self._check_busy = True
        self.checklist_button.set_state("正在檢查來源檔案…", enabled=False)
        self.rescan_button.set_state("掃描中…", enabled=False)
        for key, (_name, detail, y) in self.checklist_rows.items():
            self._draw_source_icon(key, y, "checking")
            self.canvas.itemconfigure(detail, text="正在檢查…", fill=MUTED)
        self.canvas.itemconfigure(self.checklist_status, text="正在檢查來源檔案，請稍候…", fill=ACCENT)
        token, customer, events = self._check_token, self.customer, self.events

        def inspect() -> None:
            checks = check_sources(path, customer)
            # Keep Tk objects out of this closure, including during app shutdown.
            events.put(("checklist", (token, signature, checks, start_when_ready)))

        threading.Thread(target=inspect, daemon=True, name="excess-source-check").start()

    def _fit_text(self, text: str, *, size: int, width: int) -> str:
        font = tkfont.Font(self.root, font=self.font(size))
        if font.measure(text) <= self.px(width):
            return text
        # Bound the visible row without letting long filenames overlap its neighbours.
        low, high = 0, len(text)
        while low < high:
            middle = (low + high + 1) // 2
            if font.measure(text[:middle] + "…") <= self.px(width):
                low = middle
            else:
                high = middle - 1
        return text[:low] + "…"

    def _receive_checks(self, payload) -> None:
        token, signature, checks, start_when_ready = payload
        if token != self._check_token or self.screen != "checklist":
            return
        self._check_busy = False
        self.rescan_button.set_state("重新掃描檔案", enabled=True)
        if signature != input_signature(customer_input_dir(self.app_root, self.customer)):
            self.source_checks = ()
            for key, (_name, detail, y) in self.checklist_rows.items():
                self._draw_source_icon(key, y, "warning")
                self.canvas.itemconfigure(detail, text="檔案有變更，請重新掃描", fill=WARNING)
            self.checklist_button.set_state("請重新掃描檔案", enabled=False)
            self.canvas.itemconfigure(self.checklist_status,
                                      text="掃描期間檔案有變更，請按「重新掃描檔案」再確認。", fill=WARNING)
            return
        self.source_checks = checks
        for check in checks:
            key = check.requirement.key
            name, detail, y = self.checklist_rows[key]
            optional_missing = check.requirement.optional and check.status == "missing"
            color = SUCCESS if check.status == "found" else (WARNING if optional_missing else ERROR)
            self._draw_source_icon(key, y, "found" if check.status == "found" else ("warning" if optional_missing else "error"))
            self.canvas.itemconfigure(name, fill=INK)
            self.canvas.itemconfigure(detail, text=self._fit_text(check.detail, size=17, width=680), fill=color)
        blockers = [check for check in checks if check.blocks_run]
        missing = [check for check in blockers if check.status == "missing"]
        if not blockers:
            status = "必要檔案已備齊；新增或更換檔案後，請重新掃描。"
            text = "確認無誤，開始執行"
        elif missing:
            status = "缺少：" + "、".join(check.requirement.label for check in missing) + "。放入後請按「重新掃描檔案」。"
            text = "請依照提示放入缺少檔案"
        else:
            status = "來源檔案有重複或無法讀取，修正後請按「重新掃描檔案」。"
            text = "請依照提示修正來源檔案"
        self.canvas.itemconfigure(self.checklist_status, text=self._fit_text(status, size=18, width=1020), fill=ERROR if blockers else SUCCESS)
        self.checklist_button.set_state(text, enabled=not blockers)
        if start_when_ready and not blockers:
            self._begin_run()

    def start_run(self) -> None:
        if self.running or self.customer is None or self.screen != "checklist" or self._check_busy:
            return
        if not self.source_checks or any(check.blocks_run for check in self.source_checks):
            return
        # Verify again off the UI thread: a file may have changed since the last scan.
        self._refresh_checklist(start_when_ready=True)

    def _begin_run(self) -> None:
        self.running = True
        self._show_running()
        customer = self.customer
        self.worker = threading.Thread(target=self._run_worker, args=(customer,), daemon=True, name="excess-report-worker")
        self.worker.start()

    def _show_running(self) -> None:
        self._reset("running")
        self.last_progress = None
        progress_total = progress_steps_for_customer(
            self.customer,
            include_summary=self.include_summary,
        )
        self._text(78, 226, "STEP 02  /  整理報表", size=17, color=ACCENT, bold=True, anchor="w")
        self.title_item = self._text(78, 266, "正在整理報表", size=36, bold=True, anchor="w")
        self._text(1122, 265, f"工廠  {self.customer}", size=22, color=ACCENT, bold=True, anchor="e")
        self.phase_item = self._text(78, 313, "準備讀取來源檔案…", size=22, anchor="w")
        self.progressbar = ttk.Progressbar(
            self.canvas,
            style="Excess.Horizontal.TProgressbar",
            maximum=progress_total,
            mode="determinate",
        )
        self.widgets.append(self.progressbar)
        self.canvas.create_window(self.px(78), self.px(341), anchor="nw", width=self.px(1044), height=self.px(16), window=self.progressbar)
        self.count_item = self._text(78, 377, f"已完成 0 / {progress_total} 個階段", size=18, color=MUTED, anchor="w")
        self.percent_item = self._text(1122, 377, "0%", size=22, color=ACCENT, bold=True, anchor="e")
        frame = tk.Frame(self.canvas, bg=PANEL, highlightbackground="#a36c32", highlightthickness=1)
        self.widgets.append(frame)
        self.canvas.create_window(self.px(78), self.px(400), anchor="nw", width=self.px(1044), height=self.px(204), window=frame)
        self.log_text = tk.Text(
            frame, bg=PANEL, fg=INK, insertbackground=INK, relief="flat", bd=0,
            font=(self.log_font_family, -max(12, self.px(20))), wrap="word", state="disabled",
            padx=self.px(16), pady=self.px(12), takefocus=True,
            spacing1=self.px(2), spacing3=self.px(3),
        )
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.log_text.yview, style="Excess.Vertical.TScrollbar")
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        self.footer_item = self._text(600, 622, "請保持程式開啟，完成後會自動提示取檔。", size=15, color=MUTED)
        self._append_log(f"工廠：{self.customer}　開始整理報表。")

    def _run_worker(self, customer: str) -> None:
        try:
            result = generate_report(
                customer_input_dir(self.app_root, customer), customer_output_file(self.app_root, customer),
                customer=customer, on_event=self.events.put,
                summary_models=self.summary_models,
                include_summary=self.include_summary,
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
                percentage = event.current * 100 // max(event.total, 1)
                self.canvas.itemconfigure(self.percent_item, text=f"{percentage}%")
                self._append_log(label)
            return
        kind, payload = event
        if kind == "checklist":
            self._receive_checks(payload)
            return
        if kind == "models":
            self._receive_models(payload)
            return
        self.running = False
        if kind == "success":
            self.result = payload
            self._show_success()
        else:
            self.exit_code = 1
            self.canvas.itemconfigure(self.title_item, text="報表處理失敗", fill=ERROR)
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
        self._poll_id = self.root.after(60, self._poll_events)

    def _show_success(self) -> None:
        self.canvas.itemconfigure(self.title_item, text="報表整理完成")
        self.canvas.itemconfigure(self.footer_item, text="按「確認」後，將開啟 output 資料夾並關閉程式。")
        dialog = self.dialog = AppDialog(self, "報表整理完成", kind="success", width=760)
        dialog.label("COMPLETED", size=17, color=SUCCESS, bold=True)
        dialog.label("已完成，打開output資料夾取檔案", size=29, bold=True, pady=(0, 18))
        dialog.label(f"output/{self.customer}/", size=21, color=ACCENT)
        filenames = [
            "excess_report.xlsx",
            "next_time_excess_report.xlsx",
        ]
        if self.result and self.result.summary_output_file is not None:
            filenames.append("excess_summary.xlsx")
        filenames.append("incomplete_alternates.xlsx")
        dialog.label("\n".join(filenames), size=20, pady=(0, 16))
        if self.result and self.result.summary_output_file is None:
            dialog.label(
                "本次已略過 summary；若資料夾已有舊版 excess_summary.xlsx，會保留原檔。",
                size=16,
                color=MUTED,
                pady=(0, 12),
            )
        dialog.label("按「確認」後開啟資料夾，並關閉此程式。", size=18, color=MUTED)
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
        dialog = self.dialog = AppDialog(self, "報表處理失敗", kind="error", width=840)
        dialog.label("PROCESSING STOPPED", size=17, color=ERROR, bold=True)
        dialog.label("無法完成本次報表", size=32, bold=True)
        if isinstance(error, ReportError):
            message = error.user_message()
        else:
            phase = self.last_progress.label if self.last_progress else "啟動介面"
            message = ReportError(None, PHASE_LABELS.get(phase, phase), error).user_message()
        # A scrollable message keeps even long paths/errors inside the screen.
        message_frame = tk.Frame(dialog.body, bg=PANEL)
        message_frame.pack(fill="x", pady=(0, self.px(10)))
        details = tk.Text(
            message_frame, bg=PANEL, fg=INK, relief="flat", bd=0,
            font=self.font(20), wrap="word", height=9, width=1,
            padx=0, pady=0, highlightthickness=0,
        )
        details.insert("1.0", message)
        details.configure(state="disabled")
        scrollbar = ttk.Scrollbar(message_frame, orient="vertical", command=details.yview, style="Excess.Vertical.TScrollbar")
        details.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        details.pack(side="left", fill="both", expand=True)
        dialog.label("按「確認」後，將關閉程式。", size=18, color=MUTED)
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
