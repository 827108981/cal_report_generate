from __future__ import annotations

import re
from decimal import Decimal
from typing import Any, Iterable
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

EXCEL_ERROR_VALUES = {'#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!', '#SPILL!', '#CALC!'}
SECTION_RE = re.compile(r'^[（(][一二三四五六七八九十百]+[）)]\s*.+')
TOP_SECTION_RE = re.compile(r'^[一二三四五六七八九十百]+[、.]\s*.+')


def is_blank_value(value: Any) -> bool:
    if value is None:
        return True
    s = str(value).strip()
    return s == '' or s in EXCEL_ERROR_VALUES


def normalize_text(text: Any) -> str:
    if text is None:
        return ''
    s = str(text)
    s = s.replace('\u3000', ' ')
    s = s.replace('：', ':')
    s = s.replace('（', '(').replace('）', ')')
    s = s.replace('，', ',').replace('；', ';')
    s = s.replace('＜', '<').replace('＞', '>')
    s = s.replace('≤', '<=').replace('≥', '>=')
    s = s.replace('μ', 'µ')
    s = s.upper()
    s = re.sub(r'\s+', '', s)
    s = re.sub(r'[\t\r\n:：,，;；/\\\-—_（）()\[\]【】<>《》]+', '', s)
    return s


def compact_texts(texts: Iterable[Any]) -> list[str]:
    result: list[str] = []
    seen = set()
    for x in texts:
        if is_blank_value(x):
            continue
        s = str(x).strip()
        if not s:
            continue
        n = normalize_text(s)
        if n in seen:
            continue
        seen.add(n)
        result.append(s)
    return result


def looks_like_section(text: Any) -> bool:
    s = str(text or '').strip()
    return bool(SECTION_RE.match(s))


def looks_like_top_section(text: Any) -> bool:
    s = str(text or '').strip()
    return bool(TOP_SECTION_RE.match(s))


def canonical_section(text: Any) -> str:
    s = str(text or '').strip()
    s = re.sub(r'^[（(][一二三四五六七八九十百]+[）)]\s*', '', s)
    return s.strip()


def _decimals_from_number_format(fmt: str) -> int | None:
    fmt = fmt or ''
    # Ignore quoted text and colors lightly.
    fmt = re.sub(r'".*?"', '', fmt)
    fmt = re.sub(r'\[[^\]]+\]', '', fmt)
    if '%' in fmt:
        m = re.search(r'0\.([0#]+)%', fmt)
        return len(m.group(1)) if m else 0
    m = re.search(r'0\.([0#]+)', fmt)
    if m:
        return len(m.group(1))
    return None


def format_number_by_excel_format(value: Any, number_format: str | None = None) -> str:
    if is_blank_value(value):
        return ''
    if isinstance(value, bool):
        return 'TRUE' if value else 'FALSE'
    if isinstance(value, int):
        return str(value)
    if isinstance(value, Decimal):
        value = float(value)
    if isinstance(value, float):
        fmt = number_format or ''
        if '%' in fmt:
            dec = _decimals_from_number_format(fmt)
            if dec is None:
                dec = 2
            return f'{value * 100:.{dec}f}%'
        dec = _decimals_from_number_format(fmt)
        if dec is None:
            # Word 表格空间有限，Excel“General”的长浮点尾巴不能原样写入；默认按报告常见精度显示。
            dec = 4
        if value.is_integer() and dec == 0:
            return str(int(value))
        text = f'{value:.{dec}f}'
        if dec > 0:
            text = text.rstrip('0').rstrip('.')
            # 小数很小但被截成 0 时，保留固定小数，避免丢掉有效量级。
            if text in {'0', '-0'} and abs(value) > 0:
                text = f'{value:.{dec}f}'
        return text
    s = str(value).strip()
    if s.upper() == 'PASS':
        return '■ Pass'
    if s.upper() == 'FAIL':
        return '■ Fail'
    return s


def format_value(value: Any) -> str:
    return format_number_by_excel_format(value)


def tc_text(tc) -> str:
    texts = []
    for el in tc.iter():
        if el.tag == qn('w:t') and el.text:
            texts.append(el.text)
    return ''.join(texts).strip()


def _remove_children_except_tcpr(tc):
    for child in list(tc):
        if child.tag != qn('w:tcPr'):
            tc.remove(child)


def _first_paragraph(tc):
    for child in tc:
        if child.tag == qn('w:p'):
            return child
    return None


def _clear_paragraph_runs(p):
    for child in list(p):
        if child.tag != qn('w:pPr'):
            p.remove(child)


def set_tc_text_keep_style(tc, value: Any, font_size_pt: float = 9.0) -> None:
    """写入单元格内容，同时尽量保留单元格属性/段落属性，并设置合适的数据字体。

    不再删除 tcPr、tblGrid、列宽等结构，避免写入数据后表格尺寸被撑开。
    """
    text = value if isinstance(value, str) else format_value(value)
    text = '' if text is None else str(text)

    # 保留第一个段落的 pPr，清掉多余段落/运行，减少因默认字号导致的撑表。
    p = _first_paragraph(tc)
    if p is None:
        _remove_children_except_tcpr(tc)
        p = OxmlElement('w:p')
        tc.append(p)
    else:
        # 删除除第一个段落和 tcPr 外的内容。
        keep = {id(p)}
        for child in list(tc):
            if child.tag == qn('w:tcPr') or id(child) in keep:
                continue
            tc.remove(child)
        _clear_paragraph_runs(p)

    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(round(font_size_pt * 2))))
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), str(int(round(font_size_pt * 2))))
    rPr.append(sz)
    rPr.append(szCs)
    r.append(rPr)
    t = OxmlElement('w:t')
    if text.startswith(' ') or text.endswith(' '):
        t.set(qn('xml:space'), 'preserve')
    t.text = text
    r.append(t)
    p.append(r)


def tc_is_writable(tc, overwrite_nonblank: bool = False) -> bool:
    txt = tc_text(tc).strip()
    if overwrite_nonblank:
        return True
    if txt == '':
        return True
    if txt in {'/', '／', '待填', '待填写', 'N/A', 'NA'}:
        return True
    if normalize_text(txt) in {'是否', 'PASSFAIL', '通过不通过', '合格不合格', '■PASSFAIL', '□PASS□FAIL'}:
        return True
    return False


def row_physical_tcs(table, row_index: int):
    return list(table._tbl.tr_lst[row_index].tc_lst)


def physical_table_matrix(table) -> list[list[str]]:
    matrix: list[list[str]] = []
    for tr in table._tbl.tr_lst:
        matrix.append([tc_text(tc) for tc in tr.tc_lst])
    return matrix


def row_nonblank_positions(row: list[str]) -> list[tuple[int, str]]:
    return [(i, str(v).strip()) for i, v in enumerate(row) if not is_blank_value(v)]


def safe_ratio(a: set[str], b: set[str]) -> float:
    if not a or not b:
        return 0.0
    return len(a & b) / max(1, len(a | b))


def set_table_fixed_layout(table) -> None:
    """固定表格布局，防止填入数据后自动扩宽超出页面。"""
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = tblPr.find(qn('w:tblLayout'))
    if layout is None:
        layout = OxmlElement('w:tblLayout')
        tblPr.append(layout)
    layout.set(qn('w:type'), 'fixed')
