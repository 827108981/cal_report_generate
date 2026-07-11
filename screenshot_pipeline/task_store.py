from __future__ import annotations

import hashlib
import json
import shutil
from datetime import datetime
from pathlib import Path

from PIL import Image

from .models import ScreenshotAsset, ScreenshotConfig, ScreenshotItemSpec, ScreenshotTask


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff"}


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _image_info(path: Path) -> tuple[int, int]:
    with Image.open(path) as image:
        image.verify()
    with Image.open(path) as image:
        return image.size


def _task_json(task: ScreenshotTask) -> dict:
    return {
        "task_id": task.task_id,
        "model_key": task.model_key,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "assets": {
            item_id: [
                {
                    "order": asset.order,
                    "file": str(asset.stored_path.relative_to(task.task_dir)),
                    "source_type": asset.source_type,
                    "original_path": asset.original_path,
                    "width": asset.width,
                    "height": asset.height,
                    "sha256": asset.sha256,
                }
                for asset in assets
            ]
            for item_id, assets in task.assets.items()
        },
    }


def save_task(task: ScreenshotTask) -> None:
    task.task_dir.mkdir(parents=True, exist_ok=True)
    task.screenshot_dir.mkdir(parents=True, exist_ok=True)
    (task.task_dir / "task.json").write_text(
        json.dumps(_task_json(task), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def create_task(parent_dir: str | Path, model_key: str) -> ScreenshotTask:
    parent = Path(parent_dir)
    parent.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    task_id = f"{stamp}_{model_key.upper()}"
    task_dir = parent / task_id
    suffix = 1
    while task_dir.exists():
        suffix += 1
        task_dir = parent / f"{task_id}_{suffix:02d}"
    task = ScreenshotTask(task_id=task_dir.name, model_key=model_key, task_dir=task_dir)
    save_task(task)
    return task


def load_task(task_dir: str | Path) -> ScreenshotTask:
    root = Path(task_dir)
    manifest = root / "task.json"
    if not manifest.exists():
        raise ValueError(f"截图任务目录缺少 task.json：{root}")
    data = json.loads(manifest.read_text(encoding="utf-8"))
    task = ScreenshotTask(
        task_id=str(data["task_id"]),
        model_key=str(data["model_key"]),
        task_dir=root,
    )
    for item_id, raw_assets in data.get("assets", {}).items():
        task.assets[item_id] = []
        for raw in raw_assets:
            task.assets[item_id].append(ScreenshotAsset(
                item_id=item_id,
                order=int(raw["order"]),
                stored_path=root / raw["file"],
                source_type=raw.get("source_type", "local"),
                original_path=raw.get("original_path", ""),
                width=int(raw.get("width", 0)),
                height=int(raw.get("height", 0)),
                sha256=raw.get("sha256", ""),
            ))
    return task


def add_file(
    task: ScreenshotTask,
    item: ScreenshotItemSpec,
    source_path: str | Path,
    source_type: str = "local",
) -> ScreenshotAsset:
    source = Path(source_path)
    if source.suffix.lower() not in IMAGE_EXTENSIONS:
        raise ValueError(f"不支持的图片格式：{source.name}")
    if not source.exists():
        raise FileNotFoundError(source)
    assets = task.assets_for(item.item_id)
    if not item.multiple and assets:
        raise ValueError(f"{item.display_name}只允许一张图片")
    if item.max_count is not None and len(assets) >= int(item.max_count):
        raise ValueError(f"{item.display_name}最多允许{item.max_count}张图片")

    width, height = _image_info(source)
    order = len(assets) + 1
    target = task.screenshot_dir / (
        f"{task.model_key.upper()}_{item.item_id}_{order:02d}{source.suffix.lower()}"
    )
    task.screenshot_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, target)
    asset = ScreenshotAsset(
        item_id=item.item_id,
        order=order,
        stored_path=target,
        source_type="capture" if source_type == "capture" else "local",
        original_path="" if source_type == "capture" else str(source.resolve()),
        width=width,
        height=height,
        sha256=_sha256(target),
    )
    assets.append(asset)
    save_task(task)
    return asset


def add_image(task: ScreenshotTask, item: ScreenshotItemSpec, image, source_type: str = "capture") -> ScreenshotAsset:
    temp = task.task_dir / ".capture.png"
    image.save(temp, format="PNG")
    try:
        return add_file(task, item, temp, source_type=source_type)
    finally:
        temp.unlink(missing_ok=True)


def remove_asset(task: ScreenshotTask, item_id: str, index: int) -> None:
    assets = task.assets_for(item_id)
    if index < 0 or index >= len(assets):
        return
    asset = assets.pop(index)
    asset.stored_path.unlink(missing_ok=True)
    for order, remaining in enumerate(assets, start=1):
        remaining.order = order
    save_task(task)


def move_asset(task: ScreenshotTask, item_id: str, index: int, delta: int) -> None:
    assets = task.assets_for(item_id)
    target = index + delta
    if index < 0 or target < 0 or index >= len(assets) or target >= len(assets):
        return
    assets[index], assets[target] = assets[target], assets[index]
    for order, asset in enumerate(assets, start=1):
        asset.order = order
    save_task(task)
