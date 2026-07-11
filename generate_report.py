from __future__ import annotations

import argparse
from pathlib import Path

from bs_section_filler.report_models import (
    MODELS,
    default_output_path,
    discover_files,
    generate_report_for_model,
    get_model,
)
from bs_section_filler.word_com import update_fields_with_word


def main() -> None:
    parser = argparse.ArgumentParser(description="BS 系列校准报告生成工具。")
    parser.add_argument("--model", default="bs2800", choices=sorted(MODELS), help="机型，默认 bs2800")
    parser.add_argument("--template", default="", help="Word 空模板 .docx；留空时按机型自动查找")
    parser.add_argument("--excel", default="", help="单 Excel 机型使用；等价于 --m1")
    parser.add_argument("--m1", default="", help="M1/主校准报告 Excel .xlsx")
    parser.add_argument("--m2", default="", help="M2 Excel .xlsx，可选；BS2800 双模块使用")
    parser.add_argument("--out", default="", help="输出 docx；留空时输出到 result 目录")
    parser.add_argument("--formula-policy", default="all", choices=["auto", "raw", "all"])
    parser.add_argument("--no-overwrite-nonblank", action="store_true")
    parser.add_argument("--update-word-fields", action="store_true")
    parser.add_argument("--keep-tail-sections", action="store_true")
    parser.add_argument("--screenshots", default="", help="截图任务目录，包含 task.json")
    args = parser.parse_args()

    model = get_model(args.model)
    discovered = discover_files(args.model, Path.cwd())
    template = Path(args.template) if args.template else discovered.template_path
    excel_paths: dict[str, Path] = dict(discovered.excel_paths)
    if args.excel:
        excel_paths["m1"] = Path(args.excel)
    if args.m1:
        excel_paths["m1"] = Path(args.m1)
    if args.m2:
        excel_paths["m2"] = Path(args.m2)
    output = Path(args.out) if args.out else default_output_path(args.model, Path("result"))

    if template is None:
        raise SystemExit(f"{model.display_name} 未找到 Word 空模板，请使用 --template 指定。")

    result = generate_report_for_model(
        model_key=args.model,
        template_path=template,
        excel_paths=excel_paths,
        output_path=output,
        formula_policy=args.formula_policy,
        overwrite_nonblank=not args.no_overwrite_nonblank,
        delete_tail_sections=False if args.keep_tail_sections else None,
        screenshot_task_dir=args.screenshots or None,
        screenshot_resource_dir=Path.cwd(),
    )
    print(f"机型：{model.display_name}")
    print(f"已生成：{result.output_path}")
    print(f"已填充单元格：{result.fill_count}")
    print(f"表格匹配数：{result.matched_count}/{result.table_count}")
    print(f"填充日志：{result.fill_log_path}")
    print(f"匹配日志：{result.match_log_path}")
    for warning in result.warnings:
        print("提示：" + warning)
    if args.update_word_fields:
        ok, message = update_fields_with_word(result.output_path)
        print(("成功：" if ok else "提示：") + message)


if __name__ == "__main__":
    main()
