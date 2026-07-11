from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterator, Union

from docx.document import Document as _Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph

from .utils import canonical_section, compact_texts, looks_like_section, normalize_text, physical_table_matrix


@dataclass
class WordTableBlock:
    index: int
    table: Table
    section: str
    title: str
    object_name: str
    header_norms: set[str]
    matrix: list[list[str]]
    module_hint: str = ""

    @property
    def key_text(self) -> str:
        return " ".join(x for x in [self.section, self.title, self.object_name] if x)


def iter_block_items(parent: _Document) -> Iterator[Union[Paragraph, Table]]:
    for child in parent.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def _extract_object_name(matrix: list[list[str]]) -> str:
    if not matrix:
        return ""
    for row in matrix[:3]:
        texts = compact_texts(row)
        joined = " ".join(texts)
        if "试剂名称" in joined:
            values = [text for text in texts if "试剂名称" not in text]
            return values[-1] if values else ""
    first = compact_texts(matrix[0])
    return first[-1] if first else ""


def _extract_title(matrix: list[list[str]], section: str) -> str:
    if not matrix:
        return section
    for row in matrix[:3]:
        texts = compact_texts(row)
        if not texts:
            continue
        if "试剂名称" in " ".join(texts):
            return section
        return texts[0]
    return section


def _header_norms(matrix: list[list[str]]) -> set[str]:
    headers: set[str] = set()
    keywords = [
        "次数", "重复次数", "测试数据", "采光周期", "检测项目", "要求", "结论",
        "吸光度", "结果", "均值", "SD", "CV", "MAX", "MIN", "指标", "温度值",
        "参数值", "校准状态", "平均值", "R²", "R", "a", "b",
    ]
    for row in matrix[:6]:
        for value in row:
            if any(keyword in str(value) for keyword in keywords):
                normalized = normalize_text(value)
                if normalized:
                    headers.add(normalized)
    return headers


def _looks_like_context_title(text: str) -> bool:
    value = (text or "").strip()
    if not value:
        return False
    if len(value) > 80 and not value.startswith(("1、", "2、", "3、", "4、", "5、")):
        return False
    if re.match(r"^[1-9]\d*[、.]\s*.+", value):
        return True
    if "波长下" in value and "测定" in value and any(
        keyword in value for keyword in ["溶液", "标准溶液", "重铬酸钾", "橙黄G", "硫酸铜"]
    ):
        return True
    return any(
        keyword in value
        for keyword in ["测试结果", "线性范围验证", "电解质准确度", "电解质精密度", "电解质稳定性", "电解质携带污染率"]
    )


def _has_reagent_name(matrix: list[list[str]]) -> bool:
    return any("试剂名称" in " ".join(compact_texts(row)) for row in matrix[:3])


def _context_should_replace_reagent(context: str, table_obj: str) -> bool:
    """Use a nearby reagent sentence only when it contradicts the copied table label.

    The BS-5000 template contains a stability table whose reagent-name cell was copied
    from the next sulfate table. The paragraph before the table contains the correct
    0.5A Orange-G name. Other tables should keep their own reagent labels.
    """

    if not context or not table_obj:
        return False
    reagent_tokens = ["橙黄G", "硫酸铜", "重铬酸钾", "亚硝酸钠"]
    context_tokens = {token for token in reagent_tokens if token in context}
    object_tokens = {token for token in reagent_tokens if token in table_obj}
    return bool(context_tokens and object_tokens and context_tokens.isdisjoint(object_tokens))


def _module_from_text(text: str) -> str | None:
    normalized = normalize_text(text)
    if "模块二" in normalized or "M2模块" in normalized or "模块2" in normalized or "ISE模块二" in normalized or "ISE模块2" in normalized:
        return "M2"
    if "模块一" in normalized or "M1模块" in normalized or "模块1" in normalized or "ISE模块一" in normalized or "ISE模块1" in normalized:
        return "M1"
    return None


def _is_generic_table_identity(title: str, obj: str, matrix: list[list[str]]) -> bool:
    normalized_title = normalize_text(title)
    normalized_obj = normalize_text(obj)
    generic = {
        "次数", "项目测试时间", "项目交叉污染", "质控品批号正常水平", "DILUTION19",
        "R22", "CL", "电解质项目的精密度选配ISE须执行",
    }
    if normalized_title in generic or normalized_obj in generic:
        return True
    first_text = " ".join(matrix[0]) if matrix else ""
    if "质控品批号" in first_text and not any(len(compact_texts(row)) > 2 for row in matrix[1:4]):
        return True
    if matrix and matrix[0] and normalize_text(matrix[0][0]) == "次数":
        return True
    return False


def get_word_table_blocks(doc) -> list[WordTableBlock]:
    blocks: list[WordTableBlock] = []
    current_section = ""
    current_context_title = ""
    current_module_hint = ""
    table_index = 0

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            module_hint = _module_from_text(text)
            if module_hint:
                current_module_hint = module_hint
            if looks_like_section(text):
                current_section = canonical_section(text)
                current_context_title = ""
            elif _looks_like_context_title(text):
                current_context_title = text
            continue

        matrix = physical_table_matrix(block)
        for row in matrix[:2]:
            texts = compact_texts(row)
            if texts and looks_like_section(texts[0]):
                current_section = canonical_section(texts[0])

        table_title = _extract_title(matrix, current_section)
        table_object = _extract_object_name(matrix)
        if current_context_title and _has_reagent_name(matrix) and _context_should_replace_reagent(current_context_title, table_object):
            title = current_context_title
            object_name = current_context_title
        elif current_context_title and _is_generic_table_identity(table_title, table_object, matrix):
            title = current_context_title
            object_name = current_context_title
        else:
            title = table_title
            object_name = table_object

        blocks.append(
            WordTableBlock(
                index=table_index,
                table=block,
                section=current_section,
                title=title,
                object_name=object_name,
                header_norms=_header_norms(matrix),
                matrix=matrix,
                module_hint=current_module_hint,
            )
        )
        table_index += 1
    return blocks
