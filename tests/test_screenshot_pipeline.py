from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from docx import Document
from PIL import Image

from screenshot_pipeline.config import load_config
from screenshot_pipeline.task_store import add_file, create_task, move_asset
from screenshot_pipeline.validator import validate_task
from screenshot_pipeline.word_inserter import insert_screenshots


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = next(ROOT.joinpath("BS5000").glob("*.docx"))


class ScreenshotPipelineTests(unittest.TestCase):
    def test_local_multiple_images_are_auto_named_and_reordered(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            task = create_task(root, "bs5000")
            config = load_config("bs5000", ROOT)
            item = config.items[2]
            first = root / "first.jpg"
            second = root / "second.png"
            Image.new("RGB", (320, 200), "red").save(first)
            Image.new("RGB", (320, 200), "blue").save(second)

            add_file(task, item, first)
            add_file(task, item, second)
            self.assertEqual([asset.order for asset in task.assets[item.item_id]], [1, 2])
            self.assertEqual(task.assets[item.item_id][0].stored_path.name, "BS5000_09_03_POWER_01.jpg")
            self.assertEqual(task.assets[item.item_id][1].stored_path.name, "BS5000_09_03_POWER_02.png")

            move_asset(task, item.item_id, 1, -1)
            self.assertEqual([asset.order for asset in task.assets[item.item_id]], [1, 2])
            self.assertEqual(task.assets[item.item_id][0].stored_path.name, "BS5000_09_03_POWER_02.png")

    def test_word_template_anchor_accepts_multiple_images(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            task = create_task(root, "bs5000")
            config = load_config("bs5000", ROOT)
            item = config.items[0]
            image = root / "raw.png"
            Image.new("RGB", (320, 200), "green").save(image)
            add_file(task, item, image)
            add_file(task, item, image)

            output = root / "output.docx"
            insert_screenshots(TEMPLATE, output, task, config)
            self.assertEqual(len(Document(output).inline_shapes), len(Document(TEMPLATE).inline_shapes) + 2)

    def test_empty_task_reports_missing_required_images(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            task = create_task(temp_dir, "bs5000")
            config = load_config("bs5000", ROOT)
            errors = [issue for issue in validate_task(task, config) if issue.level == "error"]
            self.assertEqual(len(errors), len(config.items))


if __name__ == "__main__":
    unittest.main()
