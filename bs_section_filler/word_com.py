from __future__ import annotations

import platform
from pathlib import Path


def update_fields_with_word(docx_path: str | Path) -> tuple[bool, str]:
    if platform.system().lower() != 'windows':
        return False, '当前不是 Windows，未执行 Word COM 更新。'
    try:
        import win32com.client  # type: ignore
    except Exception as e:
        return False, f'未安装 pywin32 或无法导入 win32com：{e}'
    word = None
    path = str(Path(docx_path).resolve())
    try:
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        doc = word.Documents.Open(path)
        doc.Fields.Update()
        for toc in doc.TablesOfContents:
            toc.Update()
        for table in doc.Tables:
            try:
                table.Range.Fields.Update()
            except Exception:
                pass
        doc.Save()
        doc.Close(False)
        return True, '已调用 Microsoft Word 更新域、目录、页码和表格公式。'
    except Exception as e:
        return False, f'Word COM 更新失败：{e}'
    finally:
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
