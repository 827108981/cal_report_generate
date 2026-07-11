from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal


ImageSource = Literal["capture", "local"]


@dataclass(frozen=True)
class ScreenshotItemSpec:
    item_id: str
    display_name: str
    anchor: str
    required: bool = False
    multiple: bool = True
    min_count: int = 0
    max_count: int | None = None
    allow_capture: bool = True
    allow_local_import: bool = True
    page_break_between: bool = False


@dataclass(frozen=True)
class ScreenshotConfig:
    model_key: str
    chapter: str
    items: tuple[ScreenshotItemSpec, ...]

    def by_id(self) -> dict[str, ScreenshotItemSpec]:
        return {item.item_id: item for item in self.items}


@dataclass
class ScreenshotAsset:
    item_id: str
    order: int
    stored_path: Path
    source_type: ImageSource
    original_path: str = ""
    width: int = 0
    height: int = 0
    sha256: str = ""


@dataclass
class ScreenshotTask:
    task_id: str
    model_key: str
    task_dir: Path
    assets: dict[str, list[ScreenshotAsset]] = field(default_factory=dict)

    @property
    def screenshot_dir(self) -> Path:
        return self.task_dir / "screenshots"

    def assets_for(self, item_id: str) -> list[ScreenshotAsset]:
        return self.assets.setdefault(item_id, [])
