from __future__ import annotations

import ctypes
import sys

from PIL import Image, ImageGrab


def _enable_dpi_awareness() -> None:
    """Use physical screen pixels instead of DPI-scaled logical pixels on Windows."""
    if sys.platform != "win32":
        return
    try:
        # Per-monitor V2 keeps both ImageGrab and Tk coordinates in physical pixels.
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
        return
    except Exception:
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


_enable_dpi_awareness()


def capture_screen() -> Image.Image:
    """Capture before showing the selection UI so the tool is never in its own image."""
    try:
        return ImageGrab.grab(include_layered_windows=True)
    except TypeError:
        return ImageGrab.grab()


def select_region(parent, image: Image.Image):
    """Select from a full-resolution frozen screen image and return the native-pixel crop."""
    import tkinter as tk
    from PIL import ImageTk

    result: dict[str, Image.Image | None] = {"image": None}
    dialog = tk.Toplevel(parent)
    dialog.title("框选原始分辨率截图")
    dialog.transient(parent)
    dialog.grab_set()
    dialog.attributes("-topmost", True)
    max_width = max(800, parent.winfo_screenwidth() - 70)
    max_height = max(540, parent.winfo_screenheight() - 130)
    view_width = min(image.width, max_width)
    view_height = min(image.height, max_height)
    dialog.geometry(f"{view_width + 20}x{view_height + 62}+20+20")

    hint = tk.Label(
        dialog,
        text=f"全屏原图：{image.width} × {image.height} 像素   ·   拖动框选   ·   Enter 保存   ·   Esc 取消",
        anchor="w",
        padx=10,
        pady=8,
    )
    hint.pack(fill="x")
    body = tk.Frame(dialog)
    body.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    body.rowconfigure(0, weight=1)
    body.columnconfigure(0, weight=1)
    canvas = tk.Canvas(body, width=view_width, height=view_height, highlightthickness=0)
    x_scroll = tk.Scrollbar(body, orient="horizontal", command=canvas.xview)
    y_scroll = tk.Scrollbar(body, orient="vertical", command=canvas.yview)
    canvas.configure(xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set, scrollregion=(0, 0, image.width, image.height))
    canvas.grid(row=0, column=0, sticky="nsew")
    y_scroll.grid(row=0, column=1, sticky="ns")
    x_scroll.grid(row=1, column=0, sticky="ew")
    photo = ImageTk.PhotoImage(image)
    canvas.create_image(0, 0, image=photo, anchor="nw")
    state = {"start": None, "rect": None}

    def press(event):
        state["start"] = (canvas.canvasx(event.x), canvas.canvasy(event.y))
        if state["rect"] is not None:
            canvas.delete(state["rect"])
        state["rect"] = canvas.create_rectangle(*state["start"], *state["start"], outline="#ff3b30", width=2)

    def drag(event):
        if state["start"] and state["rect"] is not None:
            canvas.coords(state["rect"], state["start"][0], state["start"][1], canvas.canvasx(event.x), canvas.canvasy(event.y))

    def accept(_event=None):
        if not state["start"]:
            return
        x1, y1 = state["start"]
        x2, y2 = canvas.coords(state["rect"])[2:]
        left, right = sorted((max(0, int(x1)), min(image.width, int(x2))))
        top, bottom = sorted((max(0, int(y1)), min(image.height, int(y2))))
        if right - left >= 10 and bottom - top >= 10:
            result["image"] = image.crop((left, top, right, bottom))
            dialog.destroy()

    def cancel(_event=None):
        dialog.destroy()

    canvas.bind("<ButtonPress-1>", press)
    canvas.bind("<B1-Motion>", drag)
    dialog.bind("<Return>", accept)
    dialog.bind("<Escape>", cancel)
    dialog.focus_force()
    dialog.wait_window()
    return result["image"]
