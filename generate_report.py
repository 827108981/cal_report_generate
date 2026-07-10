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


def main():
    parser = argparse.ArgumentParser(description='BS 系列校准报告生成工具。')
    parser.add_argument('--model', default='bs2800', choices=sorted(MODELS), help='机型，默认 bs2800')
    parser.add_argument('--template', default='', help='Word 空模板 .docx；留空时按机型自动查找')
    parser.add_argument('--excel', default='', help='单 Excel 机型使用；等价于 --m1')
    parser.add_argument('--m1', default='', help='M1/主校准报告 Excel .xlsx')
    parser.add_argument('--m2', default='', help='M2 Excel .xlsx，可选；BS2800 双模块报告使用')
    parser.add_argument('--out', default='', help='输出 docx 路径；留空时输出到 result 目录')
    parser.add_argument('--formula-policy', default='all', choices=['auto','raw','all'], help='默认all=同步Excel中已计算结果；raw=只填原始数据；auto=检测Word公式后决定')
    parser.add_argument('--no-overwrite-nonblank', action='store_true', help='默认允许覆盖数据区非空占位；指定该参数则只写空白/占位单元格')
    parser.add_argument('--update-word-fields', action='store_true', help='调用本机 Microsoft Word 更新域/目录/公式')
    parser.add_argument('--keep-tail-sections', action='store_true', help='强制保留顶层六、七、八、九章节；不指定时按机型默认策略处理')
    args = parser.parse_args()

    model = get_model(args.model)
    discovered = discover_files(args.model, Path.cwd())
    template = Path(args.template) if args.template else discovered.template_path
    excel_paths = {k: v for k, v in discovered.excel_paths.items()}
    if args.excel:
        excel_paths['m1'] = Path(args.excel)
    if args.m1:
        excel_paths['m1'] = Path(args.m1)
    if args.m2:
        excel_paths['m2'] = Path(args.m2)
    out = Path(args.out) if args.out else default_output_path(args.model, Path('result'))

    if template is None:
        raise SystemExit(f'{model.display_name} 未找到 Word 空模板，请使用 --template 指定。')

    result = generate_report_for_model(
        model_key=args.model,
        template_path=template,
        excel_paths=excel_paths,
        output_path=out,
        formula_policy=args.formula_policy,
        overwrite_nonblank=not args.no_overwrite_nonblank,
        delete_tail_sections=False if args.keep_tail_sections else None,
    )
    print(f'机型：{model.display_name}')
    print(f'已生成：{result.output_path}')
    print(f'已填充单元格：{result.fill_count}')
    print(f'表格匹配数：{result.matched_count}/{result.table_count}')
    print(f'填充日志：{result.fill_log_path}')
    print(f'匹配日志：{result.match_log_path}')
    for w in result.warnings:
        print('提示：' + w)
    if args.update_word_fields:
        ok, msg = update_fields_with_word(result.output_path)
        print(('成功：' if ok else '提示：') + msg)


if __name__ == '__main__':
    main()
