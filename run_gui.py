from __future__ import annotations

import threading
import sys
import tkinter as tk
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
from bs_section_filler.word_com import update_fields_with_word


def _base_dir() -> Path:
    first = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
    candidates = [first, Path.cwd().resolve(), first.parent]
    for candidate in candidates:
        if (candidate / 'BS5000').exists() or (candidate / 'BS2800').exists():
            return candidate
    return first


BASE_DIR = _base_dir()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('BS 校准报告生成器')
        self.geometry('940x620')
        self.minsize(860, 560)

        self.model_by_display = {m.display_name: m.key for m in list_models()}
        self.display_by_model = {m.key: m.display_name for m in list_models()}
        self.model_display_var = tk.StringVar()
        self.template_var = tk.StringVar()
        self.out_dir_var = tk.StringVar(value=str((BASE_DIR / 'result').resolve()))
        self.formula_policy_var = tk.StringVar(value='all')
        self.update_fields_var = tk.BooleanVar(value=True)
        self.overwrite_var = tk.BooleanVar(value=True)
        self.keep_tail_var = tk.BooleanVar(value=False)
        self.excel_vars: dict[str, tk.StringVar] = {}
        self.input_frame: ttk.Frame | None = None

        self._configure_style()
        self._build()
        self._select_initial_model()

    def _configure_style(self):
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('.', font=('Microsoft YaHei UI', 10))
        style.configure('Title.TLabel', font=('Microsoft YaHei UI', 18, 'bold'))
        style.configure('Hint.TLabel', foreground='#667085')
        style.configure('Primary.TButton', font=('Microsoft YaHei UI', 10, 'bold'), padding=(16, 8))
        style.configure('Card.TFrame', background='#f8fafc')
        style.configure('Card.TLabelframe', background='#f8fafc')
        style.configure('Card.TLabelframe.Label', font=('Microsoft YaHei UI', 10, 'bold'))

    def _build(self):
        root = ttk.Frame(self, padding=18)
        root.pack(fill='both', expand=True)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(3, weight=1)

        header = ttk.Frame(root)
        header.grid(row=0, column=0, sticky='ew')
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text='校准报告生成器', style='Title.TLabel').grid(row=0, column=0, sticky='w')
        ttk.Label(header, text='选择机型，载入模板和原始 Excel，生成最终 Word 报告。', style='Hint.TLabel').grid(row=1, column=0, sticky='w', pady=(4, 0))

        model_bar = ttk.Frame(root, padding=(0, 18, 0, 8))
        model_bar.grid(row=1, column=0, sticky='ew')
        model_bar.columnconfigure(1, weight=1)
        ttk.Label(model_bar, text='机型').grid(row=0, column=0, sticky='w', padx=(0, 10))
        model_box = ttk.Combobox(
            model_bar,
            textvariable=self.model_display_var,
            values=[m.display_name for m in list_models()],
            state='readonly',
            width=24,
        )
        model_box.grid(row=0, column=1, sticky='w')
        model_box.bind('<<ComboboxSelected>>', lambda _event: self._apply_model())

        form = ttk.LabelFrame(root, text='文件', padding=14, style='Card.TLabelframe')
        form.grid(row=2, column=0, sticky='ew', pady=(0, 12))
        form.columnconfigure(1, weight=1)
        self._file_row(form, 0, 'Word 空模板', self.template_var, 'docx')
        self.input_frame = ttk.Frame(form)
        self.input_frame.grid(row=1, column=0, columnspan=3, sticky='ew')
        self.input_frame.columnconfigure(1, weight=1)
        self._dir_row(form, 2, '输出目录', self.out_dir_var)

        lower = ttk.Frame(root)
        lower.grid(row=3, column=0, sticky='nsew')
        lower.columnconfigure(0, weight=1)
        lower.rowconfigure(1, weight=1)

        options = ttk.LabelFrame(lower, text='选项', padding=14, style='Card.TLabelframe')
        options.grid(row=0, column=0, sticky='ew')
        options.columnconfigure(3, weight=1)
        ttk.Label(options, text='公式策略').grid(row=0, column=0, sticky='w')
        ttk.Combobox(
            options,
            textvariable=self.formula_policy_var,
            values=['all', 'auto', 'raw'],
            state='readonly',
            width=10,
        ).grid(row=0, column=1, sticky='w', padx=(8, 22))
        ttk.Checkbutton(options, text='覆盖旧值', variable=self.overwrite_var).grid(row=0, column=2, sticky='w', padx=(0, 22))
        ttk.Checkbutton(options, text='更新 Word 域', variable=self.update_fields_var).grid(row=0, column=3, sticky='w')
        ttk.Checkbutton(options, text='保留六/七/八/九章节', variable=self.keep_tail_var).grid(row=0, column=4, sticky='e')

        action_bar = ttk.Frame(lower, padding=(0, 12, 0, 8))
        action_bar.grid(row=1, column=0, sticky='ew')
        action_bar.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(action_bar, mode='indeterminate', length=220)
        self.progress.grid(row=0, column=0, sticky='w')
        self.start_button = ttk.Button(action_bar, text='生成报告', style='Primary.TButton', command=self.start)
        self.start_button.grid(row=0, column=1, sticky='e')

        self.log = tk.Text(lower, height=12, wrap='word', relief='flat', borderwidth=1)
        self.log.grid(row=2, column=0, sticky='nsew')
        lower.rowconfigure(2, weight=1)

    def _file_row(self, parent, row: int, label: str, var: tk.StringVar, file_kind: str):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky='w', padx=(0, 10), pady=6)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky='ew', pady=6)
        ttk.Button(parent, text='选择', command=lambda: self._choose_file(var, file_kind)).grid(row=row, column=2, padx=(10, 0), pady=6)

    def _dir_row(self, parent, row: int, label: str, var: tk.StringVar):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky='w', padx=(0, 10), pady=6)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky='ew', pady=6)
        ttk.Button(parent, text='选择', command=lambda: self._choose_dir(var)).grid(row=row, column=2, padx=(10, 0), pady=6)

    def _choose_file(self, var: tk.StringVar, file_kind: str):
        if file_kind == 'docx':
            filetypes = [('Word 文档', '*.docx'), ('所有文件', '*.*')]
        else:
            filetypes = [('Excel 工作簿', '*.xlsx'), ('所有文件', '*.*')]
        path = filedialog.askopenfilename(initialdir=BASE_DIR, filetypes=filetypes)
        if path:
            var.set(path)

    def _choose_dir(self, var: tk.StringVar):
        path = filedialog.askdirectory(initialdir=BASE_DIR)
        if path:
            var.set(path)

    def _select_initial_model(self):
        initial = 'bs5000' if (BASE_DIR / 'BS5000').exists() else 'bs2800'
        if initial not in MODELS:
            initial = next(iter(MODELS))
        self.model_display_var.set(self.display_by_model[initial])
        self._apply_model()

    def _current_model_key(self) -> str:
        display = self.model_display_var.get()
        return self.model_by_display.get(display, 'bs2800')

    def _apply_model(self):
        model_key = self._current_model_key()
        model = get_model(model_key)
        self.keep_tail_var.set(not model.delete_tail_sections_default)
        discovered = discover_files(model_key, BASE_DIR)
        if discovered.template_path:
            self.template_var.set(str(discovered.template_path.resolve()))

        assert self.input_frame is not None
        for child in self.input_frame.winfo_children():
            child.destroy()
        self.excel_vars = {}
        for row, spec in enumerate(model.excel_inputs):
            var = tk.StringVar()
            if spec.key in discovered.excel_paths:
                var.set(str(discovered.excel_paths[spec.key].resolve()))
            self.excel_vars[spec.key] = var
            label = spec.label + (' *' if spec.required else '')
            self._file_row(self.input_frame, row, label, var, 'xlsx')

        self.write(f'已切换机型：{model.display_name}\n')

    def _set_running(self, running: bool):
        state = 'disabled' if running else 'normal'
        self.start_button.configure(state=state)
        if running:
            self.progress.start(10)
        else:
            self.progress.stop()

    def write(self, msg: str):
        def append():
            self.log.insert('end', msg)
            self.log.see('end')
        self.after(0, append)

    def start(self):
        model_key = self._current_model_key()
        model = get_model(model_key)
        template = self.template_var.get().strip()
        out_dir = self.out_dir_var.get().strip()
        excel_paths = {key: var.get().strip() for key, var in self.excel_vars.items() if var.get().strip()}

        missing = []
        if not template:
            missing.append('Word 空模板')
        if not out_dir:
            missing.append('输出目录')
        for spec in model.excel_inputs:
            if spec.required and not excel_paths.get(spec.key):
                missing.append(spec.label)
        if missing:
            messagebox.showwarning('缺少文件', '请补充：' + '、'.join(missing))
            return

        output_path = default_output_path(model_key, out_dir)
        self._set_running(True)
        self.write(f'\n开始生成：{model.display_name}\n')
        threading.Thread(
            target=self._run,
            args=(model_key, template, excel_paths, output_path),
            daemon=True,
        ).start()

    def _run(self, model_key: str, template: str, excel_paths: dict[str, str], output_path: Path):
        try:
            result = generate_report_for_model(
                model_key=model_key,
                template_path=template,
                excel_paths=excel_paths,
                output_path=output_path,
                formula_policy=self.formula_policy_var.get(),
                overwrite_nonblank=self.overwrite_var.get(),
                delete_tail_sections=not self.keep_tail_var.get(),
            )
            for warning in result.warnings:
                self.write('提示：' + warning + '\n')
            self.write(f'已生成：{result.output_path}\n')
            self.write(f'已填充单元格：{result.fill_count}\n')
            self.write(f'表格匹配数：{result.matched_count}/{result.table_count}\n')
            self.write(f'填充日志：{result.fill_log_path}\n')
            self.write(f'匹配日志：{result.match_log_path}\n')
            if self.update_fields_var.get():
                ok, msg = update_fields_with_word(result.output_path)
                self.write(('成功：' if ok else '提示：') + msg + '\n')
            self.after(0, lambda: messagebox.showinfo('完成', f'已生成：\n{result.output_path}'))
        except Exception as exc:
            import traceback
            msg = str(exc)
            self.write('失败：' + msg + '\n' + traceback.format_exc() + '\n')
            self.after(0, lambda: messagebox.showerror('失败', msg))
        finally:
            self.after(0, lambda: self._set_running(False))


if __name__ == '__main__':
    App().mainloop()
