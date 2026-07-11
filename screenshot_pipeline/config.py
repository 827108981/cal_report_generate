from __future__ import annotations

import json
from pathlib import Path

from .models import ScreenshotConfig, ScreenshotItemSpec


def config_path(model_key: str, resource_dir: str | Path | None = None) -> Path:
    root = Path(resource_dir) if resource_dir is not None else Path(__file__).resolve().parents[1]
    return root / ("BS2800" if model_key == "bs2800" else "BS5000") / "screenshot_map.json"


def load_config(model_key: str, resource_dir: str | Path | None = None) -> ScreenshotConfig:
    path = config_path(model_key, resource_dir)
    if not path.exists():
        return ScreenshotConfig(model_key=model_key, chapter="九、原始数据", items=())
    data = json.loads(path.read_text(encoding="utf-8"))
    items = tuple(
        ScreenshotItemSpec(
            item_id=str(item["item_id"]),
            display_name=str(item["display_name"]),
            anchor=str(item.get("anchor", "")),
            required=bool(item.get("required", False)),
            multiple=bool(item.get("multiple", True)),
            min_count=int(item.get("min_count", 0)),
            max_count=item.get("max_count"),
            allow_capture=bool(item.get("allow_capture", True)),
            allow_local_import=bool(item.get("allow_local_import", True)),
            page_break_between=bool(item.get("page_break_between", False)),
        )
        for item in data.get("items", [])
    )
    return ScreenshotConfig(
        model_key=str(data.get("model", model_key)),
        chapter=str(data.get("chapter", "九、原始数据")),
        items=items,
    )
