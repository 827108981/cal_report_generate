"""Screenshot collection and Word insertion pipeline."""

from .config import ScreenshotConfig, load_config
from .models import ScreenshotTask
from .task_store import create_task, load_task

__all__ = ["ScreenshotConfig", "ScreenshotTask", "create_task", "load_config", "load_task"]
