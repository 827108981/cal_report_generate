from __future__ import annotations

from pathlib import Path

from .config import load_config
from .models import ScreenshotConfig, ScreenshotTask
from .task_store import load_task
from .validator import validate_task
from .word_inserter import insert_screenshots


def load_and_validate_task(task_dir: str | Path, resource_dir: str | Path | None = None) -> tuple[ScreenshotTask, ScreenshotConfig, list[str]]:
    task = load_task(task_dir)
    config = load_config(task.model_key, resource_dir)
    issues = validate_task(task, config)
    errors = [issue.message for issue in issues if issue.level == "error"]
    warnings = [issue.message for issue in issues if issue.level == "warning"]
    if errors:
        raise ValueError("截图校验失败：" + "；".join(errors))
    return task, config, warnings


def insert_task_screenshots(
    output_path: str | Path,
    task_dir: str | Path,
    resource_dir: str | Path | None = None,
) -> list[str]:
    task, config, warnings = load_and_validate_task(task_dir, resource_dir)
    warnings.extend(insert_screenshots(output_path, output_path, task, config))
    return warnings
