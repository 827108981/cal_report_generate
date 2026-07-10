from __future__ import annotations

import re
import zipfile
from pathlib import Path


def find_word_table_formulas(docx_path: str | Path) -> list[str]:
    formulas: list[str] = []
    with zipfile.ZipFile(docx_path) as z:
        for name in z.namelist():
            if not name.startswith('word/') or not name.endswith('.xml'):
                continue
            xml = z.read(name).decode('utf-8', errors='ignore')
            instr = re.findall(r'<w:instrText[^>]*>(.*?)</w:instrText>', xml)
            instr += re.findall(r'<w:fldSimple[^>]*w:instr="([^"]+)"', xml)
            for s in instr:
                up = s.upper().strip()
                if not up:
                    continue
                if up.startswith(('TOC', 'PAGEREF', 'PAGE', 'NUMPAGES', 'HYPERLINK', 'MACROBUTTON')):
                    continue
                if up.startswith('=') or any(k in up for k in ['SUM(', 'AVERAGE(', 'PRODUCT(', 'ROUND(', 'IF(']):
                    formulas.append(s)
    return formulas
