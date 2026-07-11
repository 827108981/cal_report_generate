from __future__ import annotations

import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageTk

from .capture import capture_screen, select_region
from .config import load_config
from .hotkeys import GlobalHotkeyListener
from .task_store import add_file, add_image, create_task, load_task, move_asset, remove_asset, save_task


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
    "danger": "#B84848",
}


class ScreenshotCollectorWindow(tk.Toplevel):
    """A focused, operator-friendly screenshot workstation."""

    def __init__(self, parent, model_key: str, resource_dir: Path, default_parent: Path, on_selected):
        super().__init__(parent)
        self.title("原始数据截图采集")
        self.geometry("1180x720")
        self.minsize(1040, 640)
        self.configure(background=COLORS["bg"])
        self.model_key = model_key
        self.resource_dir = resource_dir
        self.on_selected = on_selected
        self.config_data = load_config(model_key, resource_dir)
        self.task = create_task(default_parent / "screenshot_tasks", model_key)
        self.item_index = 0
        self.preview_image = None
        self._capturing = False
        self.hotkeys = GlobalHotkeyListener()
        self._configure_style()
        self._build()
        self.protocol("WM_DELETE_WINDOW", self._close)
        self._start_hotkeys()
        self._refresh_all()
        self._select_item(0)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.configure("Capture.TFrame", background=COLORS["bg"])
        style.configure("CapturePanel.TFrame", background=COLORS["surface"])
        style.configure("Capture.TLabel", background=COLORS["surface"], foreground=COLORS["text"], font=("Microsoft YaHei UI", 10))
        style.configure("CaptureMuted.TLabel", background=COLORS["surface"], foreground=COLORS["muted"], font=("Microsoft YaHei UI", 9))
        style.configure("CaptureTitle.TLabel", background=COLORS["bg"], foreground=COLORS["text"], font=("Microsoft YaHei UI", 19, "bold"))
        style.configure("CaptureSection.TLabel", background=COLORS["surface"], foreground=COLORS["text"], font=("Microsoft YaHei UI", 12, "bold"))
        style.configure("CapturePrimary.TButton", background=COLORS["primary"], foreground="#FFFFFF", borderwidth=0, padding=(14, 8), font=("Microsoft YaHei UI", 10, "bold"))
        style.map("CapturePrimary.TButton", background=[("active", COLORS["primary_active"]), ("pressed", COLORS["primary_active"])])
        style.configure("CaptureSecondary.TButton", background="#E8F1F0", foreground=COLORS["primary"], borderwidth=0, padding=(12, 8))
        style.map("CaptureSecondary.TButton", background=[("active", "#D9E9E6")])
        style.configure("CaptureHotkey.TLabel", background="#E8F1F0", foreground=COLORS["primary"], font=("Microsoft YaHei UI", 9, "bold"), padding=(10, 7))
        style.configure("CaptureHotkeyWarn.TLabel", background="#FFF3E4", foreground=COLORS["warning"], font=("Microsoft YaHei UI", 9, "bold"), padding=(10, 7))
        style.configure("Capture.Treeview", background=COLORS["surface"], fieldbackground=COLORS["surface"], foreground=COLORS["text"], rowheight=31, borderwidth=0, font=("Microsoft YaHei UI", 9))
        style.configure("Capture.Treeview.Heading", background="#EEF3F5", foreground=COLORS["muted"], relief="flat", font=("Microsoft YaHei UI", 9, "bold"))
        style.map("Capture.Treeview", background=[("selected", "#DCEEEB")], foreground=[("selected", COLORS["text"])])

    def _panel(self, parent, padding: int = 16):
        frame = ttk.Frame(parent, style="CapturePanel.TFrame", padding=padding)
        frame.configure(relief="solid", borderwidth=1)
        return frame

    def _build(self) -> None:
        root = ttk.Frame(self, style="Capture.TFrame", padding=22)
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)

        header = ttk.Frame(root, style="Capture.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 16))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="原始数据截图", style="CaptureTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text=f"{self.model_key.upper()}  ·  选项目 → 切到仪器界面 → 按 F8/F9 → 回来确认图片 → 完成采集并返回报告", style="CaptureMuted.TLabel").grid(row=1, column=0, sticky="w", pady=(3, 0))
        ttk.Button(header, text="继续已有采集", style="CaptureSecondary.TButton", command=self._open_task).grid(row=0, column=1, rowspan=2, padx=(12, 8))
        ttk.Button(header, text="完成采集并返回报告", style="CapturePrimary.TButton", command=self._use_task).grid(row=0, column=2, rowspan=2)

        workspace = ttk.Frame(root, style="Capture.TFrame")
        workspace.grid(row=1, column=0, sticky="nsew")
        workspace.columnconfigure(0, weight=0, minsize=330)
        workspace.columnconfigure(1, weight=1)
        workspace.rowconfigure(0, weight=1)

        left = self._panel(workspace)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(2, weight=1)
        ttk.Label(left, text="采集清单", style="CaptureSection.TLabel").grid(row=0, column=0, sticky="w")
        self.summary_label = ttk.Label(left, style="CaptureMuted.TLabel")
        self.summary_label.grid(row=1, column=0, sticky="w", pady=(4, 12))
        self.items = ttk.Treeview(left, columns=("state", "name", "count"), show="headings", style="Capture.Treeview", selectmode="browse")
        self.items.heading("state", text="状态")
        self.items.heading("name", text="项目")
        self.items.heading("count", text="图片")
        self.items.column("state", width=62, anchor="center", stretch=False)
        self.items.column("name", width=180, anchor="w")
        self.items.column("count", width=50, anchor="center", stretch=False)
        self.items.grid(row=2, column=0, sticky="nsew")
        self.items.bind("<<TreeviewSelect>>", self._on_item_selected)

        right = self._panel(workspace)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.columnconfigure(1, weight=1)
        right.rowconfigure(4, weight=1)
        ttk.Label(right, text="当前项目", style="CaptureSection.TLabel").grid(row=0, column=0, sticky="w")
        self.item_label = ttk.Label(right, style="CaptureMuted.TLabel")
        self.item_label.grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 14))
        self.hotkey_label = ttk.Label(right, style="CaptureHotkeyWarn.TLabel", wraplength=700, justify="left")
        self.hotkey_label.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 12))

        toolbar = ttk.Frame(right, style="CapturePanel.TFrame")
        toolbar.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(0, 14))
        for column in range(5):
            toolbar.columnconfigure(column, weight=1)
        ttk.Button(toolbar, text="截取整屏后框选  F8", style="CapturePrimary.TButton", command=self._capture_region).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(toolbar, text="直接保存整屏  F9", style="CaptureSecondary.TButton", command=self._capture_fullscreen).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(toolbar, text="导入本地照片", style="CaptureSecondary.TButton", command=self._import).grid(row=0, column=2, sticky="ew", padx=6)
        ttk.Button(toolbar, text="排到前面", style="CaptureSecondary.TButton", command=lambda: self._move(-1)).grid(row=0, column=3, sticky="ew", padx=6)
        ttk.Button(toolbar, text="排到后面", style="CaptureSecondary.TButton", command=lambda: self._move(1)).grid(row=0, column=4, sticky="ew", padx=(6, 0))

        preview = ttk.Frame(right, style="CapturePanel.TFrame")
        preview.grid(row=4, column=0, sticky="nsew", padx=(0, 10))
        preview.columnconfigure(0, weight=1)
        preview.rowconfigure(1, weight=1)
        ttk.Label(preview, text="图片预览", style="CaptureMuted.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.preview_canvas = tk.Canvas(preview, background="#EEF3F5", highlightthickness=0)
        self.preview_canvas.grid(row=1, column=0, sticky="nsew")
        self.preview_canvas.bind("<Configure>", lambda _event: self._refresh_preview())

        assets_panel = ttk.Frame(right, style="CapturePanel.TFrame")
        assets_panel.grid(row=4, column=1, sticky="nsew")
        assets_panel.columnconfigure(0, weight=1)
        assets_panel.rowconfigure(1, weight=1)
        ttk.Label(assets_panel, text="图片顺序（报告将按此顺序插入）", style="CaptureMuted.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        self.assets = ttk.Treeview(assets_panel, columns=("order", "source", "size", "file"), show="headings", style="Capture.Treeview", selectmode="browse")
        self.assets.heading("order", text="序号")
        self.assets.heading("source", text="来源")
        self.assets.heading("size", text="尺寸")
        self.assets.heading("file", text="文件")
        self.assets.column("order", width=46, anchor="center", stretch=False)
        self.assets.column("source", width=54, anchor="center", stretch=False)
        self.assets.column("size", width=92, anchor="center", stretch=False)
        self.assets.column("file", width=180, anchor="w")
        self.assets.grid(row=1, column=0, sticky="nsew")
        self.assets.bind("<<TreeviewSelect>>", lambda _event: self._refresh_preview())

        footer = ttk.Frame(root, style="Capture.TFrame")
        footer.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        footer.columnconfigure(0, weight=1)
        self.status = ttk.Label(footer, style="CaptureMuted.TLabel")
        self.status.grid(row=0, column=0, sticky="w")
        ttk.Button(footer, text="删除当前图片", style="CaptureSecondary.TButton", command=self._delete).grid(row=0, column=1)

    def _current_spec(self):
        return self.config_data.items[self.item_index] if self.config_data.items else None

    def _refresh_all(self) -> None:
        self._refresh_items()
        self._refresh_assets()
        self.status.configure(text=f"任务目录：{self.task.task_dir}")

    def _start_hotkeys(self) -> None:
        available = self.hotkeys.start()
        thinkpad_hint = (
            " ThinkPad \u7b49\u7b14\u8bb0\u672c\u9ed8\u8ba4\u53ef\u80fd\u5c06 F8/F9 \u8bbe\u4e3a\u98de\u884c\u6a21\u5f0f/\u6d88\u606f\u4e2d\u5fc3\uff1a"
            "\u8bf7\u4f7f\u7528 Fn+F8\u3001Fn+F9\uff0c\u6216\u6309 Fn+Esc \u5207\u6362 F1-F12 \u6807\u51c6\u6a21\u5f0f\u3002"
        )
        if available["region"]:
            region_state = "F8 \u5df2\u542f\u7528"
        else:
            region_state = "F8 \u5df2\u88ab\u5360\u7528"
            self.bind("<F8>", lambda _event: self._capture_region())
        if available["fullscreen"]:
            fullscreen_state = "F9 \u5df2\u542f\u7528"
        else:
            fullscreen_state = "F9 \u5df2\u88ab\u5360\u7528"
            self.bind("<F9>", lambda _event: self._capture_fullscreen())
        status = f"\u5feb\u6377\u952e\u72b6\u6001\uff1a{region_state} | {fullscreen_state}\u3002"
        if available["region"] and available["fullscreen"]:
            self.hotkey_label.configure(
                text=status + "\u4fdd\u6301\u672c\u7a97\u53e3\u6253\u5f00\uff0c\u5207\u5230\u4eea\u5668\u8f6f\u4ef6\u540e\u76f4\u63a5\u6309 F8 \u6846\u9009\u6216 F9 \u4fdd\u5b58\u6574\u5c4f\u3002" + thinkpad_hint,
                style="CaptureHotkey.TLabel",
            )
        else:
            unavailable = "\u3001".join([key for key, ok in (("F8", available["region"]), ("F9", available["fullscreen"])) if not ok])
            self.hotkey_label.configure(
                text=status + f"{unavailable} \u88ab\u5176\u4ed6\u7a0b\u5e8f\u5360\u7528\uff0c\u53ea\u80fd\u5728\u672c\u7a97\u53e3\u524d\u53f0\u4f7f\u7528\u5bf9\u5e94\u6309\u94ae\u6216\u5feb\u6377\u952e\u3002" + thinkpad_hint,
                style="CaptureHotkeyWarn.TLabel",
            )
            self.after_idle(
                lambda: messagebox.showwarning(
                    "\u5feb\u6377\u952e\u4e0d\u53ef\u7528",
                    f"{unavailable} \u88ab\u5176\u4ed6\u7a0b\u5e8f\u5360\u7528\u3002\n\n"
                    "\u8bf7\u5173\u95ed\u5360\u7528\u5feb\u6377\u952e\u7684\u7a0b\u5e8f\u540e\u91cd\u65b0\u6253\u5f00\u91c7\u96c6\u7a97\u53e3\uff0c"
                    "\u6216\u5728\u672c\u7a97\u53e3\u524d\u53f0\u70b9\u51fb\u5bf9\u5e94\u622a\u56fe\u6309\u94ae\u3002",
                    parent=self,
                )
            )
        self.after(70, self._poll_hotkeys)

    def _poll_hotkeys(self) -> None:
        if not self.winfo_exists():
            return
        for event in self.hotkeys.drain():
            if self._capturing:
                continue
            if event == "region":
                self._capture_region()
            elif event == "fullscreen":
                self._capture_fullscreen()
        self.after(70, self._poll_hotkeys)

    def _stop_hotkeys(self) -> None:
        self.hotkeys.stop()

    def _close(self) -> None:
        self._stop_hotkeys()
        self.destroy()

    def _refresh_items(self) -> None:
        current = self._current_spec().item_id if self._current_spec() else ""
        for row in self.items.get_children():
            self.items.delete(row)
        completed = 0
        for item in self.config_data.items:
            count = len(self.task.assets.get(item.item_id, []))
            state = "已完成" if count >= item.min_count else "待采集"
            if state == "已完成":
                completed += 1
            self.items.insert("", "end", iid=item.item_id, values=(state, item.display_name, count))
        self.summary_label.configure(text=f"已完成 {completed} / {len(self.config_data.items)} 项")
        if current:
            self.items.selection_set(current)

    def _refresh_assets(self) -> None:
        spec = self._current_spec()
        for row in self.assets.get_children():
            self.assets.delete(row)
        if spec is None:
            self.item_label.configure(text="当前机型没有截图配置")
            self._refresh_preview()
            return
        records = self.task.assets.get(spec.item_id, [])
        required = f"至少 {spec.min_count} 张" if spec.required else "可选项目"
        self.item_label.configure(text=f"{spec.display_name}  ·  {required}  ·  当前 {len(records)} 张")
        for index, asset in enumerate(records):
            source = "截屏" if asset.source_type == "capture" else "导入"
            self.assets.insert("", "end", iid=str(index), values=(asset.order, source, f"{asset.width} × {asset.height}", asset.stored_path.name))
        if records:
            self.assets.selection_set("0")
        self._refresh_preview()

    def _refresh_preview(self) -> None:
        self.preview_canvas.delete("all")
        spec = self._current_spec()
        selected = self.assets.selection()
        if spec is None or not selected:
            self.preview_canvas.create_text(160, 80, text="选择一张图片后预览", fill=COLORS["muted"], font=("Microsoft YaHei UI", 10))
            return
        records = self.task.assets.get(spec.item_id, [])
        index = int(selected[0])
        if index >= len(records):
            return
        try:
            with Image.open(records[index].stored_path) as original:
                image = original.convert("RGB")
            width = max(100, self.preview_canvas.winfo_width())
            height = max(100, self.preview_canvas.winfo_height())
            image.thumbnail((width - 24, height - 24), Image.Resampling.LANCZOS)
            self.preview_image = ImageTk.PhotoImage(image)
            self.preview_canvas.create_image(width // 2, height // 2, image=self.preview_image, anchor="center")
        except Exception as exc:
            self.preview_canvas.create_text(160, 80, text=f"无法预览图片：{exc}", fill=COLORS["danger"], font=("Microsoft YaHei UI", 10))

    def _select_item(self, index: int) -> None:
        if not self.config_data.items:
            return
        self.item_index = max(0, min(index, len(self.config_data.items) - 1))
        item_id = self.config_data.items[self.item_index].item_id
        self.items.selection_set(item_id)
        self.items.focus(item_id)
        self.items.see(item_id)
        self._refresh_assets()

    def _on_item_selected(self, _event=None) -> None:
        selected = self.items.selection()
        if not selected:
            return
        selected_id = selected[0]
        for index, item in enumerate(self.config_data.items):
            if item.item_id == selected_id:
                self.item_index = index
                self._refresh_assets()
                break

    def _import(self) -> None:
        spec = self._current_spec()
        if spec is None or not spec.allow_local_import:
            return
        paths = filedialog.askopenfilenames(
            parent=self,
            title=f"导入 {spec.display_name} 图片（支持多选）",
            filetypes=[("图片", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff"), ("所有文件", "*.*")],
        )
        for path in paths:
            try:
                add_file(self.task, spec, path, source_type="local")
            except Exception as exc:
                messagebox.showerror("导入失败", str(exc), parent=self)
                break
        self._refresh_all()

    def _capture_region(self) -> None:
        spec = self._current_spec()
        if self._capturing or spec is None or not spec.allow_capture:
            return
        self._capturing = True
        try:
            image = self._capture_without_ui()
            cropped = select_region(self, image)
            if cropped is not None:
                add_image(self.task, spec, cropped, source_type="capture")
                self._refresh_all()
        except Exception as exc:
            messagebox.showerror("截图失败", str(exc), parent=self)
        finally:
            self._capturing = False

    def _capture_fullscreen(self) -> None:
        spec = self._current_spec()
        if self._capturing or spec is None or not spec.allow_capture:
            return
        self._capturing = True
        try:
            image = self._capture_without_ui()
            add_image(self.task, spec, image, source_type="capture")
            self._refresh_all()
            self.status.configure(text=f"已保存原始分辨率截图：{image.width} × {image.height}")
        except Exception as exc:
            messagebox.showerror("截图失败", str(exc), parent=self)
        finally:
            self._capturing = False

    def _capture_without_ui(self):
        hidden_windows = [self]
        if self.master is not self and hasattr(self.master, "withdraw"):
            hidden_windows.append(self.master)
        try:
            for window in hidden_windows:
                window.withdraw()
            self.update_idletasks()
            self.master.update_idletasks()
            time.sleep(0.28)
            return capture_screen()
        finally:
            for window in reversed(hidden_windows):
                try:
                    window.deiconify()
                    window.lift()
                except tk.TclError:
                    pass
            self.focus_force()

    def _selected_asset_index(self) -> int:
        selected = self.assets.selection()
        return int(selected[0]) if selected else -1

    def _move(self, delta: int) -> None:
        spec = self._current_spec()
        index = self._selected_asset_index()
        if spec is None or index < 0:
            return
        move_asset(self.task, spec.item_id, index, delta)
        self._refresh_assets()
        target = max(0, min(index + delta, len(self.task.assets.get(spec.item_id, [])) - 1))
        self.assets.selection_set(str(target))
        self._refresh_preview()

    def _delete(self) -> None:
        spec = self._current_spec()
        index = self._selected_asset_index()
        if spec is None or index < 0:
            return
        if not messagebox.askyesno("删除图片", "删除当前选中的截图？", parent=self):
            return
        remove_asset(self.task, spec.item_id, index)
        self._refresh_all()

    def _open_task(self) -> None:
        path = filedialog.askdirectory(parent=self, title="选择已有截图任务目录")
        if not path:
            return
        try:
            task = load_task(path)
            if task.model_key != self.model_key:
                raise ValueError("已有任务的机型与当前机型不一致")
            self.task = task
            self._refresh_all()
        except Exception as exc:
            messagebox.showerror("任务打开失败", str(exc), parent=self)

    def _use_task(self) -> None:
        save_task(self.task)
        self.on_selected(self.task.task_dir)
        self._stop_hotkeys()
        self.destroy()
