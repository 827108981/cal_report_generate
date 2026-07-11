from __future__ import annotations

import platform
from pathlib import Path


def update_fields_with_word(docx_path: str | Path) -> tuple[bool, str]:
    """Use desktop Word to refresh fields. This is optional post-processing.

    The GUI calls this in a background thread after already reporting successful
    generation, so Word startup/repair prompts cannot hide the export result.
    """
    if platform.system().lower() != "windows":
        return False, "当前不是 Windows，未执行 Word COM 更新。"
    try:
        import pythoncom  # type: ignore
        import win32com.client  # type: ignore
    except Exception as exc:
        return False, f"未安装 pywin32 或无法导入 Word COM：{exc}"

    word = None
    document = None
    path = str(Path(docx_path).resolve())
    try:
        pythoncom.CoInitialize()
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(path, ConfirmConversions=False, ReadOnly=False, AddToRecentFiles=False)
        document.Fields.Update()
        for toc in document.TablesOfContents:
            toc.Update()
        for table in document.Tables:
            try:
                table.Range.Fields.Update()
            except Exception:
                pass
        document.Save()
        return True, "已调用 Microsoft Word 更新域、目录、页码和表格公式。"
    except Exception as exc:
        return False, f"Word COM 更新失败：{exc}"
    finally:
        try:
            if document is not None:
                document.Close(False)
        except Exception:
            pass
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
