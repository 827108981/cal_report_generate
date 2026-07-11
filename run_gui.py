from __future__ import annotations

import os
import sys
import threading
import tkinter as tk
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from bs_section_filler.report_models import (
    MODELS,
    default_output_path,
    discover_files,
    generate_report_for_model,
    get_model,
    list_models,
)
from screenshot_pipeline.config import load_config
from screenshot_pipeline.task_store import load_task
from screenshot_pipeline.ui import ScreenshotCollectorWindow


COLORS = {
    "bg": "#F4F7F9",
    "surface": "#FFFFFF",
    "border": "#DDE5EA",
    "text": "#17212B",
    "muted": "#6B7785",
    "primary": "#16746A",
    "primary_active": "#105E57",
    "success": "#177A46",
    "warning": "#B66B12",
}


def _application_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _resource_dir() -> Path:
    candidates = [_application_dir(), Path.cwd().resolve(), _application_dir().parent]
    bundle_dir = getattr(sys, "_MEIPASS", None)
    if bundle_dir:
        candidates.insert(0, Path(bundle_dir).resolve())
    for candidate in candidates:
        if (candidate / "BS5000").exists() or (candidate / "BS2800").exists():
            return candidate
    return _application_dir()


APP_DIR = _application_dir()
RESOURCE_DIR = _resource_dir()


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("化免校准报告自动生成工具")
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        self.geometry(f"{min(1100, screen_width - 70)}x{min(680, screen_height - 110)}")
        self.minsize(940, 600)
        self.configure(background=COLORS["bg"])

        self.model_by_display = {model.display_name: model.key for model in list_models()}
        self.display_by_model = {model.key: model.display_name for model in list_models()}
        self.model_display_var = tk.StringVar()
        self.template_var = tk.StringVar()
        self.out_dir_var = tk.StringVar(value=str((APP_DIR / "result").resolve()))
        self.screenshot_task_var = tk.StringVar()
        self.status_var = tk.StringVar(value="准备就绪")
        self.screenshot_summary_var = tk.StringVar(value="尚未选择截图任务")
        self.screenshot_path_var = tk.StringVar(value="截图任务会自动归档在输出目录下")
        self.excel_vars: dict[str, tk.StringVar] = {}
        self.excel_frame: ttk.Frame | None = None
        self.start_button: ttk.Button | None = None
        self.progress: ttk.Progressbar | None = None

        self._configure_style()
        self._build()
        self._select_initial_model()

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("App.TFrame", background=COLORS["bg"])
        style.configure("Panel.TFrame", background=COLORS["surface"])
        style.configure("App.TLabel", background=COLORS["bg"], foreground=COLORS["text"], font=("Microsoft YaHei UI", 10))
        style.configure("Panel.TLabel", background=COLORS["surface"], foreground=COLORS["text"], font=("Microsoft YaHei UI", 10))
        style.configure("Muted.TLabel", background=COLORS["surface"], foreground=COLORS["muted"], font=("Microsoft YaHei UI", 9))
        style.configure("Title.TLabel", background=COLORS["bg"], foreground=COLORS["text"], font=("Microsoft YaHei UI", 21, "bold"))
        style.configure("Section.TLabel", background=COLORS["surface"], foreground=COLORS["text"], font=("Microsoft YaHei UI", 12, "bold"))
        style.configure("StatusGood.TLabel", background=COLORS["surface"], foreground=COLORS["success"], font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("StatusWarn.TLabel", background=COLORS["surface"], foreground=COLORS["warning"], font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Primary.TButton", background=COLORS["primary"], foreground="#FFFFFF", borderwidth=0, padding=(18, 10), font=("Microsoft YaHei UI", 10, "bold"))
        style.map("Primary.TButton", background=[("active", COLORS["primary_active"]), ("pressed", COLORS["primary_active"])])
        style.configure("Secondary.TButton", background="#E8F1F0", foreground=COLORS["primary"], borderwidth=0, padding=(12, 8), font=("Microsoft YaHei UI", 10))
        style.map("Secondary.TButton", background=[("active", "#D9E9E6")])
        style.configure("Thin.TButton", background="#F1F4F6", foreground=COLORS["text"], borderwidth=0, padding=(10, 7), font=("Microsoft YaHei UI", 9))
        style.map("Thin.TButton", background=[("active", "#E6ECEF")])
        style.configure("TEntry", fieldbackground="#FBFCFD", bordercolor=COLORS["border"], padding=7)
        style.configure("TCombobox", fieldbackground="#FBFCFD", padding=6)
        style.configure("TCheckbutton", background=COLORS["surface"], foreground=COLORS["text"], font=("Microsoft YaHei UI", 9))

    def _panel(self, parent, padding: int = 18):
        panel = ttk.Frame(parent, style="Panel.TFrame", padding=padding)
        panel.configure(relief="solid", borderwidth=1)
        return panel

    def _build(self) -> None:
        root = ttk.Frame(self, style="App.TFrame", padding=18)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)
        root.rowconfigure(3, weight=0)

        header = ttk.Frame(root, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="化免校准报告自动生成工具", style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="一线流程：选择机型 → 开始 / 继续采集 → 切至仪器界面按 F8/F9 → 完成采集并返回报告 → 生成报告", style="App.TLabel", foreground=COLORS["muted"]).grid(row=1, column=0, sticky="w", pady=(4, 0))
        model_control = ttk.Frame(header, style="App.TFrame")
        model_control.grid(row=0, column=1, rowspan=2, sticky="e")
        ttk.Label(model_control, text="机型", style="App.TLabel").pack(side="left", padx=(0, 8))
        self.model_box = ttk.Combobox(
            model_control,
            textvariable=self.model_display_var,
            values=[model.display_name for model in list_models()],
            state="readonly",
            width=23,
        )
        self.model_box.pack(side="left")
        self.model_box.bind("<<ComboboxSelected>>", lambda _event: self._apply_model())

        workspace = ttk.Frame(root, style="App.TFrame")
        workspace.grid(row=1, column=0, sticky="nsew")
        workspace.columnconfigure(0, weight=3)
        workspace.columnconfigure(1, weight=2)
        workspace.rowconfigure(0, weight=1)

        report_panel = self._panel(workspace)
        report_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        report_panel.columnconfigure(1, weight=1)
        ttk.Label(report_panel, text="报告文件", style="Section.TLabel").grid(row=0, column=0, columnspan=3, sticky="w")
        ttk.Label(report_panel, text="模板、校准数据和输出位置。必填项会自动识别。", style="Muted.TLabel").grid(row=1, column=0, columnspan=3, sticky="w", pady=(3, 14))
        self._path_row(report_panel, 2, "Word 空模板", self.template_var, "docx")
        self.excel_frame = ttk.Frame(report_panel, style="Panel.TFrame")
        self.excel_frame.grid(row=3, column=0, columnspan=3, sticky="ew")
        self.excel_frame.columnconfigure(1, weight=1)
        self._path_row(report_panel, 4, "输出目录", self.out_dir_var, "directory")

        screenshot_panel = self._panel(workspace)
        screenshot_panel.grid(row=0, column=1, sticky="nsew")
        screenshot_panel.columnconfigure(0, weight=1)
        ttk.Label(screenshot_panel, text="原始截图", style="Section.TLabel").grid(row=0, column=0, sticky="w")
        self.screenshot_state_label = ttk.Label(screenshot_panel, textvariable=self.screenshot_summary_var, style="StatusWarn.TLabel")
        self.screenshot_state_label.grid(row=1, column=0, sticky="w", pady=(4, 12))
        ttk.Label(screenshot_panel, text="截图任务会记录图片来源、顺序和插入位置。每个项目支持程序截屏或多张本地照片。", style="Muted.TLabel", wraplength=300, justify="left").grid(row=2, column=0, sticky="w")
        task_path = ttk.Entry(screenshot_panel, textvariable=self.screenshot_path_var, state="readonly")
        task_path.grid(row=3, column=0, sticky="ew", pady=(18, 10))
        actions = ttk.Frame(screenshot_panel, style="Panel.TFrame")
        actions.grid(row=4, column=0, sticky="ew")
        actions.columnconfigure(0, weight=1)
        actions.columnconfigure(1, weight=1)
        ttk.Button(actions, text="开始 / 继续采集", style="Primary.TButton", command=self._open_screenshot_collector).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(actions, text="选择已有截图", style="Secondary.TButton", command=self._choose_screenshot_task).grid(row=0, column=1, sticky="ew", padx=(6, 0))

        action_bar = self._panel(root, padding=14)
        action_bar.grid(row=2, column=0, sticky="ew", pady=(12, 10))
        action_bar.columnconfigure(0, weight=1)
        ttk.Label(action_bar, textvariable=self.status_var, style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(action_bar, mode="indeterminate", length=180)
        self.progress.grid(row=0, column=1, padx=(10, 16))
        ttk.Button(action_bar, text="\u6253\u5f00\u8f93\u51fa\u6587\u4ef6\u5939", style="Thin.TButton", command=self._open_output_folder).grid(row=0, column=2, padx=(0, 10))
        self.start_button = ttk.Button(action_bar, text="生成校准报告", style="Primary.TButton", command=self.start)
        self.start_button.grid(row=0, column=3, sticky="e")

        log_panel = self._panel(root, padding=10)
        log_panel.grid(row=3, column=0, sticky="nsew")
        log_panel.columnconfigure(0, weight=1)
        log_panel.rowconfigure(1, weight=1)
        ttk.Label(log_panel, text="运行记录", style="Section.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.log = tk.Text(log_panel, height=4, wrap="word", relief="flat", borderwidth=0, background="#F8FAFB", foreground=COLORS["text"], font=("Cascadia Mono", 9), padx=12, pady=8)
        self.log.grid(row=1, column=0, sticky="nsew")

    def _path_row(self, parent, row: int, label: str, variable: tk.StringVar, kind: str) -> None:
        ttk.Label(parent, text=label, style="Panel.TLabel").grid(row=row, column=0, sticky="w", padx=(0, 12), pady=6)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=6)
        ttk.Button(parent, text="浏览", style="Thin.TButton", command=lambda: self._choose_path(variable, kind)).grid(row=row, column=2, padx=(10, 0), pady=6)

    def _choose_path(self, variable: tk.StringVar, kind: str) -> None:
        if kind == "directory":
            path = filedialog.askdirectory(parent=self, initialdir=APP_DIR)
        else:
            filetypes = [("Word 文档", "*.docx"), ("所有文件", "*.*")] if kind == "docx" else [("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")]
            path = filedialog.askopenfilename(parent=self, initialdir=RESOURCE_DIR, filetypes=filetypes)
        if path:
            variable.set(path)

    def _select_initial_model(self) -> None:
        initial = "bs5000" if (RESOURCE_DIR / "BS5000").exists() else "bs2800"
        self.model_display_var.set(self.display_by_model.get(initial, next(iter(self.model_by_display))))
        self._apply_model()

    def _current_model_key(self) -> str:
        return self.model_by_display.get(self.model_display_var.get(), "bs2800")

    def _apply_model(self) -> None:
        model_key = self._current_model_key()
        model = get_model(model_key)
        self.screenshot_task_var.set("")
        discovered = discover_files(model_key, RESOURCE_DIR)
        if discovered.template_path:
            self.template_var.set(str(discovered.template_path.resolve()))
        assert self.excel_frame is not None
        for child in self.excel_frame.winfo_children():
            child.destroy()
        self.excel_vars = {}
        for row, spec in enumerate(model.excel_inputs):
            variable = tk.StringVar(value=str(discovered.excel_paths[spec.key].resolve()) if spec.key in discovered.excel_paths else "")
            self.excel_vars[spec.key] = variable
            self._path_row(self.excel_frame, row, spec.label + (" *" if spec.required else ""), variable, "xlsx")
        self._refresh_screenshot_summary()
        self.status_var.set(f"已切换为 {model.display_name}，将按正式报告规则生成")
        self.write(f"已切换机型：{model.display_name}\n")

    def _choose_screenshot_task(self) -> None:
        path = filedialog.askdirectory(parent=self, title="选择截图任务目录")
        if path:
            self.screenshot_task_var.set(path)
            self._refresh_screenshot_summary()

    def _open_output_folder(self) -> None:
        output_dir = Path(self.out_dir_var.get().strip() or (APP_DIR / "result")).expanduser()
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
            os.startfile(str(output_dir.resolve()))
            self.status_var.set("\u5df2\u6253\u5f00\u8f93\u51fa\u6587\u4ef6\u5939")
        except OSError as exc:
            messagebox.showerror("\u65e0\u6cd5\u6253\u5f00\u6587\u4ef6\u5939", str(exc), parent=self)

    def _open_screenshot_collector(self) -> None:
        parent = Path(self.out_dir_var.get().strip() or (APP_DIR / "result"))
        ScreenshotCollectorWindow(
            self,
            self._current_model_key(),
            RESOURCE_DIR,
            parent,
            self._set_screenshot_task,
        )

    def _set_screenshot_task(self, path: Path) -> None:
        self.screenshot_task_var.set(str(path.resolve()))
        self._refresh_screenshot_summary()

    def _refresh_screenshot_summary(self) -> None:
        path = self.screenshot_task_var.get().strip()
        if not path:
            self.screenshot_summary_var.set("尚未选择截图任务")
            self.screenshot_path_var.set("截图任务会自动归档在输出目录下")
            self.screenshot_state_label.configure(style="StatusWarn.TLabel")
            return
        try:
            task = load_task(path)
            if task.model_key != self._current_model_key():
                raise ValueError("截图任务机型不一致")
            config = load_config(task.model_key, RESOURCE_DIR)
            complete = sum(1 for item in config.items if len(task.assets.get(item.item_id, [])) >= item.min_count)
            total_images = sum(len(images) for images in task.assets.values())
            self.screenshot_summary_var.set(f"已完成 {complete} / {len(config.items)} 项 · 共 {total_images} 张图片")
            self.screenshot_path_var.set(str(Path(path).resolve()))
            self.screenshot_state_label.configure(style="StatusGood.TLabel" if complete == len(config.items) else "StatusWarn.TLabel")
        except Exception as exc:
            self.screenshot_summary_var.set(f"截图任务不可用：{exc}")
            self.screenshot_path_var.set(path)
            self.screenshot_state_label.configure(style="StatusWarn.TLabel")

    def _set_running(self, running: bool) -> None:
        assert self.start_button is not None and self.progress is not None
        self.start_button.configure(state="disabled" if running else "normal")
        if running:
            self.progress.start(10)
            self.status_var.set("正在生成报告，请稍候…")
        else:
            self.progress.stop()

    def write(self, message: str) -> None:
        def append() -> None:
            self.log.insert("end", message)
            self.log.see("end")
        self.after(0, append)

    def start(self) -> None:
        model_key = self._current_model_key()
        model = get_model(model_key)
        template = self.template_var.get().strip()
        out_dir = self.out_dir_var.get().strip()
        screenshot_task = self.screenshot_task_var.get().strip()
        excel_paths = {key: variable.get().strip() for key, variable in self.excel_vars.items() if variable.get().strip()}
        missing = [spec.label for spec in model.excel_inputs if spec.required and not excel_paths.get(spec.key)]
        if not template:
            missing.insert(0, "Word 空模板")
        if not out_dir:
            missing.append("输出目录")
        if missing:
            messagebox.showwarning("缺少文件", "请补充：" + "、".join(missing), parent=self)
            return
        self._set_running(True)
        self.write(f"\n开始生成：{model.display_name}\n")
        threading.Thread(
            target=self._run_generate,
            args=(model_key, template, excel_paths, default_output_path(model_key, out_dir), screenshot_task),
            daemon=True,
        ).start()

    def _run_generate(self, model_key, template, excel_paths, output_path, screenshot_task) -> None:
        try:
            result = generate_report_for_model(
                model_key=model_key,
                template_path=template,
                excel_paths=excel_paths,
                output_path=output_path,
                formula_policy="all",
                overwrite_nonblank=True,
                delete_tail_sections=False,
                screenshot_task_dir=screenshot_task or None,
                screenshot_resource_dir=RESOURCE_DIR,
            )
            self.after(0, lambda: self._finish_success(result))
        except Exception:
            detail = traceback.format_exc()
            message = detail.strip().splitlines()[-1] if detail.strip() else "未知错误"
            self.after(0, lambda: self._finish_error(message, detail))

    def _finish_success(self, result) -> None:
        self._set_running(False)
        self.status_var.set("报告已生成")
        for warning in result.warnings:
            self.write("提示：" + warning + "\n")
        self.write(f"已生成：{result.output_path}\n已填充单元格：{result.fill_count}\n表格匹配数：{result.matched_count}/{result.table_count}\n")
        messagebox.showinfo("导出成功", f"报告已生成：\n{result.output_path}", parent=self)

    def _finish_error(self, message: str, detail: str) -> None:
        self._set_running(False)
        self.status_var.set("生成失败")
        self.write("生成失败：" + message + "\n" + detail + "\n")
        messagebox.showerror("生成失败", message, parent=self)


if __name__ == "__main__":
    App().mainloop()
