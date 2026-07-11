from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .hierarchical_filler import MatchLog, fill_report_tables, save_logs
from screenshot_pipeline.pipeline import insert_task_screenshots


@dataclass(frozen=True)
class ExcelInputSpec:
    key: str
    label: str
    required: bool
    patterns: tuple[str, ...]


@dataclass(frozen=True)
class ReportModel:
    key: str
    display_name: str
    output_filename: str
    template_patterns: tuple[str, ...]
    excel_inputs: tuple[ExcelInputSpec, ...]
    delete_tail_sections_default: bool = True


@dataclass(frozen=True)
class DiscoveredFiles:
    template_path: Path | None
    excel_paths: dict[str, Path]


@dataclass(frozen=True)
class ReportGenerationResult:
    output_path: Path
    fill_log_path: Path
    match_log_path: Path
    fill_count: int
    matched_count: int
    table_count: int
    warnings: tuple[str, ...]


MODELS: dict[str, ReportModel] = {
    "bs2800": ReportModel(
        key="bs2800",
        display_name="BS-2800M系列",
        output_filename="BS-2800M系列_全自动生化分析仪仪器校准报告_自动生成.docx",
        template_patterns=(
            "BS2800/*_全自动生化分析仪仪器校准报告*_V*.docx",
            "*BS-2800M系列*_全自动生化分析仪仪器校准报告*_V*.docx",
            "BS2800/*.docx",
            "*.docx",
        ),
        excel_inputs=(
            ExcelInputSpec("m1", "M1 校准报告 Excel", True, ("BS2800/*M1*.xlsx", "*M1*校准报告.xlsx")),
            ExcelInputSpec("m2", "M2 校准报告 Excel", False, ("BS2800/*M2*.xlsx", "*M2*校准报告.xlsx")),
        ),
    ),
    "bs5000": ReportModel(
        key="bs5000",
        display_name="BS-5000系列",
        output_filename="BS-5000系列_全自动生化分析仪仪器校准报告_自动生成.docx",
        template_patterns=(
            "BS5000/*_全自动生化分析仪仪器校准报告*_V*.docx",
            "BS5000/*.docx",
        ),
        excel_inputs=(
            ExcelInputSpec("m1", "校准报告 Excel", True, ("BS5000/*校准报告*.xlsx", "BS5000/*.xlsx")),
        ),
        delete_tail_sections_default=False,
    ),
}


def get_model(model_key: str) -> ReportModel:
    try:
        return MODELS[model_key]
    except KeyError as exc:
        raise ValueError(f"未知机型：{model_key}。可选：{', '.join(MODELS)}") from exc


def list_models() -> list[ReportModel]:
    return list(MODELS.values())


def _first_existing(base_dir: Path, patterns: tuple[str, ...]) -> Path | None:
    for pattern in patterns:
        matches = sorted(path for path in base_dir.glob(pattern) if path.is_file())
        if matches:
            return matches[0]
    return None


def discover_files(model_key: str, base_dir: str | Path | None = None) -> DiscoveredFiles:
    model = get_model(model_key)
    root = Path(base_dir) if base_dir is not None else Path.cwd()
    template = _first_existing(root, model.template_patterns)
    excel_paths: dict[str, Path] = {}
    for spec in model.excel_inputs:
        found = _first_existing(root, spec.patterns)
        if found is not None:
            excel_paths[spec.key] = found
    return DiscoveredFiles(template, excel_paths)


def default_output_path(model_key: str, out_dir: str | Path) -> Path:
    return Path(out_dir) / get_model(model_key).output_filename


def _match_count(matches: list[MatchLog]) -> int:
    return sum(1 for item in matches if item.note == "已匹配")


def generate_report_for_model(
    model_key: str,
    template_path: str | Path,
    excel_paths: dict[str, str | Path],
    output_path: str | Path,
    formula_policy: str = "all",
    overwrite_nonblank: bool = True,
    delete_tail_sections: bool | None = None,
    screenshot_task_dir: str | Path | None = None,
    screenshot_resource_dir: str | Path | None = None,
) -> ReportGenerationResult:
    model = get_model(model_key)
    if not str(template_path).strip():
        raise ValueError("缺少 Word 空模板。")
    template = Path(template_path)
    out = Path(output_path)
    missing = [spec.label for spec in model.excel_inputs if spec.required and not excel_paths.get(spec.key)]
    if missing:
        raise ValueError("缺少必填文件：" + "、".join(missing))
    tail_policy = model.delete_tail_sections_default if delete_tail_sections is None else delete_tail_sections
    logs, matches, warnings = fill_report_tables(
        template_path=template,
        output_path=out,
        m1_excel=excel_paths["m1"],
        m2_excel=excel_paths.get("m2") or None,
        formula_policy=formula_policy,
        overwrite_nonblank=overwrite_nonblank,
        delete_tail_sections=tail_policy,
    )
    fill_log, match_log = save_logs(logs, matches, out)
    if screenshot_task_dir:
        screenshot_warnings = insert_task_screenshots(
            output_path=out,
            task_dir=screenshot_task_dir,
            resource_dir=screenshot_resource_dir,
        )
        warnings.extend(screenshot_warnings)
    if not out.exists() or out.stat().st_size <= 0:
        raise RuntimeError(f"报告保存失败：{out}")
    return ReportGenerationResult(
        output_path=out,
        fill_log_path=fill_log,
        match_log_path=match_log,
        fill_count=len(logs),
        matched_count=_match_count(matches),
        table_count=len(matches),
        warnings=tuple(warnings),
    )
