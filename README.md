# BS 校准报告生成工具

支持 BS-2800M 系列和 BS-5000 系列校准报告生成。核心逻辑共用同一套 Word/Excel 表格匹配引擎，机型差异集中在 `bs_section_filler/report_models.py` 配置中。

## GUI 使用

```bat
run_gui.bat
```

界面中选择机型后，程序会尝试自动带出当前目录下的模板和原始 Excel：

- BS-2800M 系列：`BS2800` 目录下的 M1/M2 校准报告 Excel。
- BS-5000 系列：`BS5000` 目录下的校准报告 Excel 和 Word 空模板。

输出文件默认保存到 `result` 目录，同时生成：

- `*_填充日志.csv`
- `*_匹配日志.csv`

## 命令行使用

自动查找 BS-5000 目录中的模板和 Excel，输出到 `result`：

```bash
python generate_report.py --model bs5000
```

指定文件：

```bash
python generate_report.py --model bs5000 --template BS5000/模板.docx --excel BS5000/校准报告.xlsx --out result/报告.docx
```

BS-2800M 双模块：

```bash
python generate_report.py --model bs2800 --template 空模板.docx --m1 M1.xlsx --m2 M2.xlsx --out result/报告.docx
```

## 常用选项

- `--formula-policy all|auto|raw`：默认 `all`，同步 Excel 中已计算结果。
- `--no-overwrite-nonblank`：只写空白或占位单元格。
- `--update-word-fields`：调用本机 Microsoft Word 更新域、目录和公式。
- `--keep-tail-sections`：保留顶层六、七、八、九章节。
