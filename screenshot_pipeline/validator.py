from __future__ import annotations

from dataclasses import dataclass

from .models import ScreenshotConfig, ScreenshotTask
from .task_store import _image_info


@dataclass(frozen=True)
class ValidationIssue:
    level: str
    item_id: str
    message: str


def validate_task(task: ScreenshotTask, config: ScreenshotConfig) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    for item in config.items:
        assets = task.assets.get(item.item_id, [])
        count = len(assets)
        if item.required and count < item.min_count:
            issues.append(ValidationIssue("error", item.item_id, f"{item.display_name}至少需要{item.min_count}张图片，当前为{count}张"))
        if item.max_count is not None and count > item.max_count:
            issues.append(ValidationIssue("error", item.item_id, f"{item.display_name}最多允许{item.max_count}张图片，当前为{count}张"))
        for asset in assets:
            try:
                width, height = _image_info(asset.stored_path)
                if width < 120 or height < 80:
                    issues.append(ValidationIssue("warning", item.item_id, f"图片尺寸较小：{asset.stored_path.name}"))
            except Exception as exc:
                issues.append(ValidationIssue("error", item.item_id, f"图片无法读取：{asset.stored_path.name}，{exc}"))
    return issues
