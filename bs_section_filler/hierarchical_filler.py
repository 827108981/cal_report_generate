from __future__ import annotations

import csv
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Optional

from docx import Document

from .excel_blocks import ExcelBlock, SourceCell, read_excel_blocks
from .utils import (
    compact_texts,
    format_value,
    is_blank_value,
    normalize_text,
    looks_like_top_section,
    physical_table_matrix,
    row_nonblank_positions,
    row_physical_tcs,
    safe_ratio,
    set_tc_text_keep_style,
    set_table_fixed_layout,
    tc_is_writable,
    tc_text,
)
from .word_blocks import WordTableBlock, get_word_table_blocks
from .word_formula_check import find_word_table_formulas


@dataclass
class FillLog:
    word_table_index: int
    word_section: str
    word_object: str
    word_row: int
    word_col: int
    module: str
    excel_sheet: str
    excel_section: str
    excel_object: str
    excel_row: int
    excel_col: int
    value: str
    reason: str


@dataclass
class MatchLog:
    word_table_index: int
    word_section: str
    word_object: str
    module: str
    excel_sheet: str
    excel_start_row: int
    excel_end_row: int
    excel_section: str
    excel_object: str
    score: float
    note: str


def _contains(a: str, b: str) -> bool:
    if not a or not b:
        return False
    return a in b or b in a


def _score_block(word: WordTableBlock, excel: ExcelBlock) -> float:
    score = 0.0
    ws = normalize_text(word.section)
    es = normalize_text(excel.section)
    wo = normalize_text(word.object_name)
    eo = normalize_text(excel.object_name)
    wt = normalize_text(word.title)
    et = normalize_text(excel.title)

    if ws and es:
        if ws == es:
            score += 60
        elif _contains(ws, es):
            score += 38
    # 对象名是关键，例如 0.5A橙黄G溶液
    if wo and eo:
        if wo == eo:
            score += 70
        elif _contains(wo, eo):
            score += 48
    if wt and et:
        if wt == et:
            score += 22
        elif _contains(wt, et):
            score += 14
    # 表头相似度作为兜底
    score += safe_ratio(word.header_norms, excel.header_norms) * 30

    # 如果完全无章节或对象匹配，只靠表头不允许高分
    if not ((_contains(ws, es) and ws and es) or (_contains(wo, eo) and wo and eo) or (_contains(wt, et) and wt and et)):
        score = min(score, 20)
    return score


def _choose_block(word: WordTableBlock, blocks: list[ExcelBlock], module: str) -> tuple[Optional[ExcelBlock], float]:
    candidates = [b for b in blocks if b.module == module]
    if not candidates:
        return None, 0.0
    best = None
    best_score = -1.0
    for b in candidates:
        s = _score_block(word, b)
        if s > best_score:
            best = b
            best_score = s
    # 低于阈值不要强行填，避免填错表
    if best is None or best_score < 55:
        return None, best_score
    return best, best_score


def _source_row_texts(row: list[SourceCell]) -> list[str]:
    return [c.text for c in row if not is_blank_value(c.value)]


def _looks_like_numeric_header_source_row(row: list[SourceCell]) -> bool:
    """识别 Excel 中类似 1/2/3 的二层表头，避免 Word 行键 1 错匹配到表头。

    v4 用“数字行键必须在 Excel 第一列”来规避该问题，但这会误杀
    吸光度重复性/稳定性等表，因为这些真实数据行的次数在 Excel 第4列。
    这里改成只跳过“整行都是小整数序号”的表头行。
    """
    texts = [normalize_text(c.text) for c in row if not is_blank_value(c.value)]
    if len(texts) < 2:
        return False
    if not all(t.isdigit() for t in texts):
        return False
    nums = [int(t) for t in texts]
    return nums == list(range(nums[0], nums[0] + len(nums))) and max(nums) <= 20


HIGH_LEVEL_ROW_LABELS = {'测试数据', '准确性计算', '重复性计算', '稳定性计算'}


def _is_placeholder_text(t: str) -> bool:
    n = normalize_text(t)
    return n in {'', '/', '／', '是否', 'PASS', 'FAIL', '■PASS', '■FAIL', 'PASSFAIL', '■PASSFAIL', '□PASS□FAIL', '通过不通过', '合格不合格', 'NA', 'N/A'}

def _row_key_candidates(texts: list[str]) -> list[str]:
    """从 Word 行中抽取行键：优先数字/编号，其次 MAX/MIN/CV/结论等。"""
    clean = [t.strip() for t in texts if t and t.strip() and not _is_placeholder_text(t)]
    if not clean:
        return []
    keys: list[str] = []
    for t in clean:
        n = normalize_text(t)
        if n in {'测试数据', '准确性计算', '重复性计算', '稳定性计算'}:
            continue
        # 数字行键、MAX/MIN/CV/结论等
        if n:
            keys.append(n)
    # 同时返回组合键，便于“稳定性计算 + MAX”匹配
    combo = normalize_text(''.join(clean[:2])) if len(clean) >= 2 else ''
    if combo:
        keys.insert(0, combo)
    return keys


def _business_key(text: str) -> str:
    n = normalize_text(text).replace('ΓGT', 'GGT').replace('Γ', 'G')
    if '加样误差' in n or '加样准确度指标' in n:
        return '加样准确度指标'
    for root in ['ALT', 'GGT', 'LDH', 'TG', 'UREA', 'TP', 'TBD', 'NA', 'CL']:
        if n == root or n.startswith(root):
            return root
    if n in {'CA', 'K', 'P'}:
        return n
    return n


def _key_matches(key: str, source_norm: str) -> bool:
    if key == source_norm:
        return True
    if _business_key(key) == _business_key(source_norm) and _business_key(key):
        return True
    for root in [
        '最大吸光度', '相对偏倚', '准确性计算', '重复性计算', '稳定性计算', '结论',
        '波动百分比', '参考色素', '携带污染率', '橙黄G样本针携带污染', '判定要求',
        '判定标准CV', '判定标准SD', '判定标准XOVER',
    ]:
        nr = normalize_text(root)
        if nr and nr in key and nr in source_norm:
            return True
    return False


def _strict_first_key_matches(key: str, source_norm: str) -> bool:
    if key == source_norm:
        return True
    bk = _business_key(key)
    bs = _business_key(source_norm)
    return bk == bs and bk in {'ALT', 'GGT', 'LDH', 'TG', 'UREA', 'TP', 'TBD', 'NA', 'CL', 'CA', 'K', 'P'}


def _find_source_row(word_row_texts: list[str], excel_rows: list[list[SourceCell]], start_index: int = 0) -> tuple[Optional[list[SourceCell]], int]:
    word_nonblank = [t for t in word_row_texts if not is_blank_value(t) and not _is_placeholder_text(t)]
    if not word_nonblank:
        return None, -1
    keys = _row_key_candidates(word_nonblank)
    if not keys:
        return None, -1

    # 第一优先级：用 Word 当前行第一个非空内容作为行键，要求匹配 Excel 行第一个非空内容。
    # 这样可以避免线性表中 Word 行“1 | 100 | 空...”误匹配到 Excel 表头行“1 | 2 | 3”。
    primary = normalize_text(word_nonblank[0])
    if primary:
        for idx, srow in enumerate(excel_rows):
            if idx < start_index:
                continue
            stexts = _source_row_texts(srow)
            if len(stexts) < 2:
                continue
            first_norm = normalize_text(stexts[0])
            protected_first_labels = {'项目', '测试项目', '测定项目', '单位', 'X轴', '通道', '校准位置参数', '指标', '要求', '判定要求'}
            if first_norm in protected_first_labels and first_norm != primary:
                continue
            if _strict_first_key_matches(primary, first_norm):
                # 纯数字行键时，只跳过 Excel 的二层数字表头（如 1/2/3），
                # 不再要求行键必须在 Excel 第一列，否则吸光度重复性/稳定性等真实数据行会被误排除。
                if primary.isdigit() and _looks_like_numeric_header_source_row(srow):
                    continue
                return srow, idx
            for root in ['最大吸光度', '相对偏倚', '准确性计算', '重复性计算', '稳定性计算', '结论', '波动百分比']:
                if root in primary and root in first_norm:
                    return srow, idx

    # 第二优先级：兼容少数行键不在第一列的表格。
    for idx, srow in enumerate(excel_rows):
        if idx < start_index:
            continue
        if len(srow) < 2:
            continue
        stexts = _source_row_texts(srow)
        snorms = [normalize_text(t) for t in stexts]
        first_norm = snorms[0] if snorms else ''
        protected_first_labels = {'项目', '测试项目', '测定项目', '单位', 'X轴', '通道', '校准位置参数', '指标', '要求', '判定要求'}
        if first_norm in protected_first_labels and first_norm not in keys:
            continue
        scombo2 = normalize_text(''.join(stexts[:2])) if len(stexts) >= 2 else ''
        for key in keys:
            if not key:
                continue
            if key == scombo2:
                return srow, idx
            # 数字行键允许出现在 Excel 非第一列（例如吸光度稳定性检测第24行，Excel第一列是纵向合并标签“测试数据”，采光周期在第4列）。
            # 但必须排除真正的二层数字表头（如 1/2/3），防止线性范围表误匹配表头。
            if key.isdigit() and key in snorms and not _looks_like_numeric_header_source_row(srow):
                return srow, idx
            # 非数字键仍不做过度宽松匹配，只允许较长文本/业务根词匹配。
            for sn in snorms:
                if _key_matches(key, sn):
                    return srow, idx
                for root in ['最大吸光度', '相对偏倚', '准确性计算', '重复性计算', '稳定性计算', '结论', '波动百分比']:
                    if root in key and root in sn:
                        return srow, idx
    return None, -1


def _source_data_after_key(word_row_texts: list[str], source_row: list[SourceCell], prefer_first_key: bool = False) -> list[SourceCell]:
    """根据 Word 行键定位 Excel 行中的数据起点。"""
    word_nonblank = [t for t in word_row_texts if not is_blank_value(t) and not _is_placeholder_text(t)]
    if not word_nonblank:
        return []
    stexts = [c.text for c in source_row]
    snorms = [normalize_text(t) for t in stexts]

    # 默认匹配 Word 行最后一个非空标签，例如线性表已有靶值 100 时从 100 右侧开始填；
    # 覆盖模式下则从第一行键右侧开始，便于覆盖模板中已有占位值/旧值。
    # 覆盖模式通常从第一行键右侧写，但遇到“测试数据/重复性计算 + 次数或Mean”等
    # 合并标签行时，真正行键是最后一个非空标签，必须从它右侧取数据，避免漏填前半段。
    if prefer_first_key and normalize_text(word_nonblank[0]) in HIGH_LEVEL_ROW_LABELS and len(word_nonblank) >= 2:
        key_text = word_nonblank[-1]
    else:
        key_text = word_nonblank[0] if prefer_first_key else word_nonblank[-1]
    key_norm = normalize_text(key_text)
    if key_norm in snorms:
        pos = snorms.index(key_norm)
        return source_row[pos + 1:]
    for pos, sn in enumerate(snorms):
        if _key_matches(key_norm, sn):
            return source_row[pos + 1:]
    # 尝试前两个组合，例如 稳定性计算MAX
    if len(word_nonblank) >= 2:
        combo = normalize_text(''.join(word_nonblank[:2]))
        scombo = normalize_text(''.join(stexts[:2])) if len(stexts) >= 2 else ''
        if combo == scombo:
            return source_row[2:]
    # 退化：如果第一列是固定合并标签“测试数据”，Word 第二列是次数，Excel 第一列是次数
    if len(word_nonblank) >= 2 and normalize_text(word_nonblank[0]) in {'测试数据', '准确性计算', '重复性计算', '稳定性计算'}:
        key = normalize_text(word_nonblank[1])
        if key in snorms:
            return source_row[snorms.index(key) + 1:]
    # 再退化：跳过与 Word 非空标签数量相同的前缀
    return source_row[min(len(word_nonblank), len(source_row)):]


def _word_empty_tcs_after_key(table, row_index: int, overwrite_nonblank: bool) -> list[tuple[int, object]]:
    """返回 Word 当前行的数据单元格。

    旧逻辑从“第一个可写空单元格”开始填，遇到纵向合并标题导致的物理空格时，
    会把第一个数据误填到左侧标签列，后面的“反应盘外圈吸光度”等列就空了。
    现在改成：先找行键所在列（如 1、2、MAX、Mean、结论），数据从行键右侧开始填。
    """
    tcs = row_physical_tcs(table, row_index)
    texts = [tc_text(tc) for tc in tcs]
    nonblank = [(i, t.strip()) for i, t in enumerate(texts) if t and t.strip() and not _is_placeholder_text(t)]
    if not nonblank:
        return []

    # 非覆盖模式：从最后一个非占位标签右侧开始，避免写到左侧纵向合并标签列。
    # 覆盖模式：从第一个非占位行键右侧开始，确保已有旧值/占位值（如 PASS、携带污染率）也能被新 Excel 数据覆盖。
    if overwrite_nonblank and normalize_text(nonblank[0][1]) in HIGH_LEVEL_ROW_LABELS and len(nonblank) >= 2:
        key_col = nonblank[-1][0]
    else:
        key_col = nonblank[0][0] if overwrite_nonblank else nonblank[-1][0]
    start = key_col + 1

    if start >= len(tcs):
        return []
    return [(ci, tc) for ci, tc in enumerate(tcs[start:], start=start) if tc_is_writable(tc, overwrite_nonblank)]

def _looks_like_header_row(texts: list[str]) -> bool:
    joined = ' '.join(t for t in texts if t)
    nonblank = [t.strip() for t in texts if t and t.strip()]
    if nonblank:
        first_norm = normalize_text(nonblank[0])
        if first_norm in {'项目', '项目次数', '项目测试时间', '项目交叉污染', '测试项目', '测定项目', 'X轴', '通道', '校准位置参数'}:
            return True
    keywords = ['次数', '采光周期', '检测项目', '校准位置参数', '要求', '吸光度', '结果', '温度值']
    if sum(1 for k in keywords if k in joined) >= 2 and not any(t.strip().isdigit() for t in texts if t.strip()):
        return True
    # 线性范围表的第二层表头：前两列空白，后面是 1/2/3，不能当作数据行。
    if len(texts) >= 5 and (not texts[0].strip()) and (not texts[1].strip()) and nonblank and all(x.isdigit() for x in nonblank):
        return True
    return False




def _font_size_for_word_table(word: WordTableBlock) -> float:
    """在不改变表格布局的前提下尽量提高字号。

    v2 字号过小；v4 按列数自适应提高字号，同时仍固定表格布局，不改变列宽。
    """
    cols = max((len(r) for r in physical_table_matrix(word.table)), default=4)
    if cols <= 4:
        return 10.5
    if cols <= 8:
        return 10.0
    if cols <= 10:
        return 9.5
    if cols <= 12:
        return 9.0
    return 8.5



def _next_writable_after_label(table, row_index: int, label_col: int, next_label_col: int | None, overwrite_nonblank: bool):
    """在同一行内，找某个标签右侧、下一个标签左侧的第一个可写单元格。"""
    tcs = row_physical_tcs(table, row_index)
    end = next_label_col if next_label_col is not None else len(tcs)
    for ci in range(label_col + 1, min(end, len(tcs))):
        if tc_is_writable(tcs[ci], overwrite_nonblank):
            return ci, tcs[ci]
    return None


def _fill_label_value_pairs(word: WordTableBlock, excel: ExcelBlock, r_idx: int, row_texts: list[str], source_row: list[SourceCell], raw_only: bool, overwrite_nonblank: bool, font_size: float) -> list[FillLog]:
    """处理一行多个“标签-值”对，例如 R²/R/a/b。普通按行键连续填充会只填最后一个标签后面的值。"""
    label_roots = ['R²', 'R', 'A', 'B']
    positions = [(i, t.strip()) for i, t in enumerate(row_texts) if t and t.strip() and not _is_placeholder_text(t)]
    norm_positions = [(i, normalize_text(t)) for i, t in positions]
    # 至少包含两个标签，才进入成对填充逻辑，避免干扰普通数据行。
    pair_labels = []
    for i, n in norm_positions:
        if n in {'R²', 'R', 'A', 'B'} or n.startswith('R²'):
            pair_labels.append((i, n))
    if len(pair_labels) < 2:
        return []

    snorms = [normalize_text(c.text) for c in source_row]
    logs: list[FillLog] = []
    for idx, (label_col, label_norm) in enumerate(pair_labels):
        next_col = pair_labels[idx + 1][0] if idx + 1 < len(pair_labels) else None
        target = _next_writable_after_label(word.table, r_idx, label_col, next_col, overwrite_nonblank)
        if not target:
            continue
        # 找 Excel 同名标签，取其右侧第一个有效值。
        src_label_idx = None
        for si, sn in enumerate(snorms):
            if sn == label_norm or (label_norm == 'R²' and sn.startswith('R²')):
                src_label_idx = si
                break
        if src_label_idx is None:
            continue
        src = None
        for sj in range(src_label_idx + 1, len(source_row)):
            cand = source_row[sj]
            if raw_only and cand.is_formula:
                continue
            if not is_blank_value(cand.value):
                src = cand
                break
        if src is None:
            continue
        target_col, tc = target
        set_tc_text_keep_style(tc, src.text, font_size_pt=font_size)
        logs.append(FillLog(
            word_table_index=word.index,
            word_section=word.section,
            word_object=word.object_name,
            word_row=r_idx + 1,
            word_col=target_col + 1,
            module=excel.module,
            excel_sheet=excel.sheet,
            excel_section=excel.section,
            excel_object=excel.object_name,
            excel_row=src.row,
            excel_col=src.col,
            value=src.text,
            reason='同一行多标签配对填充',
        ))
    return logs


def _row_tracking_key(row_texts: list[str]) -> str:
    clean = [t.strip() for t in row_texts if t and t.strip() and not _is_placeholder_text(t)]
    if not clean:
        return ''
    if normalize_text(clean[0]) in HIGH_LEVEL_ROW_LABELS and len(clean) >= 2:
        return normalize_text(clean[-1])
    return normalize_text(clean[0])


def _candidate_blocks_for_fill(word: WordTableBlock, blocks: list[ExcelBlock], module: str, primary: ExcelBlock | None) -> list[ExcelBlock]:
    word_identity = normalize_text(word.section + word.title + word.object_name + ''.join(''.join(row) for row in word.matrix[:2]))
    allow_section_expansion = any(key in word_identity for key in [
        normalize_text('主机位置校准'),
        normalize_text('暗电流检测'),
        normalize_text('样本携带污染率检测'),
        normalize_text('临床测试精密度及准确性'),
    ])
    if not allow_section_expansion:
        return [primary] if primary is not None else []

    ws = normalize_text(word.section)
    candidates: list[ExcelBlock] = []
    for block in blocks:
        if block.module != module:
            continue
        es = normalize_text(block.section)
        if primary is block or (ws and es and _contains(ws, es)):
            candidates.append(block)
    if primary is not None and primary not in candidates:
        candidates.append(primary)
    if not candidates:
        return [primary] if primary is not None else []
    seen: set[tuple[str, int, int, str]] = set()
    unique: list[ExcelBlock] = []
    for block in candidates:
        key = (block.sheet, block.start_row, block.end_row, block.module)
        if key in seen:
            continue
        seen.add(key)
        unique.append(block)
    return sorted(unique, key=lambda block: _score_block(word, block), reverse=True)


def _choose_source_for_row(
    word: WordTableBlock,
    row_texts: list[str],
    candidates: list[ExcelBlock],
    used_rows: dict[tuple[int, str], int],
) -> tuple[ExcelBlock | None, list[SourceCell] | None, int]:
    tracking_key = _row_tracking_key(row_texts)
    best_block: ExcelBlock | None = None
    best_row: list[SourceCell] | None = None
    best_index = -1
    best_score = -1.0
    for order, block in enumerate(candidates):
        start_index = used_rows.get((id(block), tracking_key), -1) + 1 if tracking_key else 0
        source_row, source_index = _find_source_row(row_texts, block.rows, start_index=start_index)
        if source_row is None:
            continue
        score = _score_block(word, block) - (order * 0.01)
        if source_index >= start_index:
            score += 4
        if score > best_score:
            best_block = block
            best_row = source_row
            best_index = source_index
            best_score = score
    if best_block is not None and tracking_key and best_index >= 0:
        used_rows[(id(best_block), tracking_key)] = best_index
    return best_block, best_row, best_index


def _sync_label_from_source(word, r_idx: int, row_texts: list[str], source_row: list[SourceCell], font_size: float) -> None:
    if not row_texts or not source_row:
        return
    word_label = (row_texts[0] or '').strip()
    source_label = (source_row[0].text or '').strip()
    if not word_label or not source_label:
        return
    if normalize_text(word_label) == normalize_text(source_label):
        return
    roots = ['最大吸光度', '相对偏倚', '参考色素', '携带污染率']
    if not any(normalize_text(root) in normalize_text(word_label) and normalize_text(root) in normalize_text(source_label) for root in roots):
        return
    tcs = row_physical_tcs(word.table, r_idx)
    if tcs:
        set_tc_text_keep_style(tcs[0], source_label, font_size_pt=font_size)


def _looks_like_clinical_accuracy_table(word: WordTableBlock) -> bool:
    if normalize_text(word.section) != normalize_text('临床测试精密度及准确性'):
        return False
    joined = normalize_text(''.join(''.join(row) for row in word.matrix[:4]))
    return '测试项目' in joined and '质控靶值' in joined and '是否在控' in joined


def _clinical_source_block(candidates: list[ExcelBlock]) -> ExcelBlock | None:
    for block in candidates:
        text = normalize_text(''.join(c.text for row in block.rows for c in row))
        if '测试项目' in text and '是否在控' in text:
            return block
    return None


def _fill_clinical_accuracy_table(word: WordTableBlock, candidates: list[ExcelBlock], overwrite_nonblank: bool, font_size: float) -> list[FillLog] | None:
    if not _looks_like_clinical_accuracy_table(word):
        return None
    excel = _clinical_source_block(candidates)
    if excel is None:
        return []

    rows_by_key: dict[str, list[list[SourceCell]]] = {}
    for row in excel.rows:
        if not row:
            continue
        key = _business_key(normalize_text(row[0].text))
        if key in {'ALT', 'GGT', 'LDH', 'TG', 'UREA', 'TP', 'TBD', 'NA', 'CL', 'CA', 'K', 'P'}:
            rows_by_key.setdefault(key, []).append(row)

    used: dict[str, int] = {}
    logs: list[FillLog] = []
    source_cols = [3, 5, 8, 11, 13, 15]
    for r_idx, row_texts in enumerate(physical_table_matrix(word.table)):
        if not row_texts:
            continue
        project = row_texts[0].strip()
        key = _business_key(normalize_text(project))
        available = rows_by_key.get(key)
        if not available:
            continue
        item_index = used.get(key, 0)
        if item_index >= len(available):
            continue
        used[key] = item_index + 1
        source_row = available[item_index]
        by_col = {cell.col: cell for cell in source_row}
        tcs = row_physical_tcs(word.table, r_idx)
        for offset, source_col in enumerate(source_cols, start=1):
            if offset >= len(tcs):
                break
            source = by_col.get(source_col)
            if source is None or is_blank_value(source.value):
                continue
            if not tc_is_writable(tcs[offset], overwrite_nonblank):
                continue
            set_tc_text_keep_style(tcs[offset], source.text, font_size_pt=font_size)
            logs.append(FillLog(
                word_table_index=word.index,
                word_section=word.section,
                word_object=word.object_name,
                word_row=r_idx + 1,
                word_col=offset + 1,
                module=excel.module,
                excel_sheet=excel.sheet,
                excel_section=excel.section,
                excel_object=excel.object_name,
                excel_row=source.row,
                excel_col=source.col,
                value=source.text,
                reason='临床准确性表按原始列映射填充',
            ))
    return logs


def _fill_table_from_block(word: WordTableBlock, excel: ExcelBlock, raw_only: bool, overwrite_nonblank: bool) -> list[FillLog]:
    return _fill_table_from_blocks(word, [excel], raw_only=raw_only, overwrite_nonblank=overwrite_nonblank)


def _fill_table_from_blocks(word: WordTableBlock, candidates: list[ExcelBlock], raw_only: bool, overwrite_nonblank: bool) -> list[FillLog]:
    logs: list[FillLog] = []
    matrix = physical_table_matrix(word.table)
    font_size = _font_size_for_word_table(word)
    clinical_logs = _fill_clinical_accuracy_table(word, candidates, overwrite_nonblank, font_size)
    if clinical_logs is not None:
        return clinical_logs
    used_rows: dict[tuple[int, str], int] = {}
    for r_idx, row_texts in enumerate(matrix):
        if r_idx == 0 and len(compact_texts(row_texts)) == 1:
            continue
        if _looks_like_header_row(row_texts):
            continue
        excel, source_row, src_index = _choose_source_for_row(word, row_texts, candidates, used_rows)
        if excel is None or source_row is None:
            continue
        _sync_label_from_source(word, r_idx, row_texts, source_row, font_size)
        # 先处理 R²/R/a/b 这类一行多个标签-值对。
        pair_logs = _fill_label_value_pairs(word, excel, r_idx, row_texts, source_row, raw_only, overwrite_nonblank, font_size)
        if pair_logs:
            logs.extend(pair_logs)
            continue
        target_cells = _word_empty_tcs_after_key(word.table, r_idx, overwrite_nonblank)
        if not target_cells:
            continue
        src_values = _source_data_after_key(row_texts, source_row, prefer_first_key=overwrite_nonblank)
        if not src_values:
            continue
        fill_i = 0
        for target_col, tc in target_cells:
            # 跳过公式源单元格：Word 有公式时只填原始数据；Word 无公式时才同步 Excel 结果
            while fill_i < len(src_values) and (raw_only and src_values[fill_i].is_formula):
                fill_i += 1
            if fill_i >= len(src_values):
                break
            src = src_values[fill_i]
            fill_i += 1
            if is_blank_value(src.value):
                continue
            set_tc_text_keep_style(tc, src.text, font_size_pt=font_size)
            logs.append(FillLog(
                word_table_index=word.index,
                word_section=word.section,
                word_object=word.object_name,
                word_row=r_idx + 1,
                word_col=target_col + 1,
                module=excel.module,
                excel_sheet=excel.sheet,
                excel_section=excel.section,
                excel_object=excel.object_name,
                excel_row=src.row,
                excel_col=src.col,
                value=src.text,
                reason='章节+名称+表头匹配后按行键复制',
            ))
    return logs


def _infer_module_for_table(word: WordTableBlock, occurrence_by_key: dict[str, int], has_m2: bool) -> str:
    if getattr(word, 'module_hint', ''):
        return word.module_hint
    text = normalize_text(word.key_text)
    if 'M2' in text and 'M1' not in text:
        return 'M2'
    if 'M1' in text and 'M2' not in text:
        return 'M1'
    if not has_m2:
        return 'M1'
    # 如果模板/标准答案里同一表连续出现两次，第一次按 M1、第二次按 M2
    key = normalize_text(word.section + '|' + word.title + '|' + word.object_name)
    count = occurrence_by_key.get(key, 0)
    occurrence_by_key[key] = count + 1
    return 'M1' if count % 2 == 0 else 'M2'


def _has_explicit_module_hint(word: WordTableBlock) -> bool:
    if getattr(word, 'module_hint', ''):
        return True
    text = normalize_text(word.key_text)
    return ('M2' in text and 'M1' not in text) or ('M1' in text and 'M2' not in text)



def _delete_top_sections_from(doc, start_nums: set[str] = {'六', '七', '八', '九'}) -> int:
    """删除从“六、...”开始到文档结尾的所有正文块。

    用户要求六、七、八、九章节不要生成；这里删除顶层章节，不影响（六）（七）这类检测子章节。
    """
    body = doc.element.body
    children = list(body)
    delete_from = None
    for idx, child in enumerate(children):
        texts = []
        for t in child.iter():
            if t.tag.endswith('}t') and t.text:
                texts.append(t.text)
        text = ''.join(texts).strip()
        if not text:
            continue
        if looks_like_top_section(text):
            first = text[0]
            if first in start_nums:
                delete_from = idx
                break
    if delete_from is None:
        return 0
    removed = 0
    for child in children[delete_from:]:
        try:
            body.remove(child)
            removed += 1
        except Exception:
            pass
    return removed


def _prepare_tables_for_fixed_fill(doc) -> None:
    for table in doc.tables:
        set_table_fixed_layout(table)

def fill_report_tables(
    template_path: str | Path,
    output_path: str | Path,
    m1_excel: str | Path,
    m2_excel: str | Path | None = None,
    formula_policy: str = 'all',
    overwrite_nonblank: bool = True,
    delete_tail_sections: bool = True,
) -> tuple[list[FillLog], list[MatchLog], list[str]]:
    """按“模块 -> 章节 -> 名称 -> 表头 -> 行键”填充 Word 表格。

    formula_policy:
      - auto：若 Word 模板有表格公式，只填 Excel 原始数据；若没有公式，同步 Excel 已计算结果。
      - raw：只填 Excel 原始数据，跳过 Excel 公式单元格。
      - all：连 Excel 公式的缓存结果也填入 Word。
    """
    warnings: list[str] = []
    formulas = find_word_table_formulas(template_path)
    if formula_policy == 'auto':
        raw_only = bool(formulas)
        if formulas:
            warnings.append(f'检测到 Word 表格公式 {len(formulas)} 个：仅复制 Excel 原始数据，公式结果由 Word 更新。')
        else:
            raw_only = False
            warnings.append('未检测到 Word 表格计算公式：已同步 Excel 中的计算结果，避免报告结果为空。')
    elif formula_policy == 'raw':
        raw_only = True
    else:
        raw_only = False

    excel_blocks = read_excel_blocks(m1_excel, 'M1')
    if m2_excel:
        excel_blocks.extend(read_excel_blocks(m2_excel, 'M2'))

    doc = Document(str(template_path))
    removed_tail_blocks = _delete_top_sections_from(doc) if delete_tail_sections else 0
    if removed_tail_blocks:
        warnings.append(f'已删除顶层六、七、八、九章节相关正文块：{removed_tail_blocks} 个。')
    _prepare_tables_for_fixed_fill(doc)
    word_blocks = get_word_table_blocks(doc)
    has_m2 = any(b.module == 'M2' for b in excel_blocks)
    occurrence_by_key: dict[str, int] = {}
    fill_logs: list[FillLog] = []
    match_logs: list[MatchLog] = []

    for wb in word_blocks:
        module = _infer_module_for_table(wb, occurrence_by_key, has_m2)
        src, score = _choose_block(wb, excel_blocks, module)
        if src is None and module != 'M1' and not _has_explicit_module_hint(wb):
            fallback_src, fallback_score = _choose_block(wb, excel_blocks, 'M1')
            if fallback_src is not None:
                module = 'M1'
                src = fallback_src
                score = fallback_score
        if src is None:
            match_logs.append(MatchLog(
                word_table_index=wb.index, word_section=wb.section, word_object=wb.object_name,
                module=module, excel_sheet='', excel_start_row=0, excel_end_row=0,
                excel_section='', excel_object='', score=score, note='未找到可信匹配，未填充',
            ))
            continue
        match_logs.append(MatchLog(
            word_table_index=wb.index, word_section=wb.section, word_object=wb.object_name,
            module=module, excel_sheet=src.sheet, excel_start_row=src.start_row, excel_end_row=src.end_row,
            excel_section=src.section, excel_object=src.object_name, score=score, note='已匹配',
        ))
        fill_candidates = _candidate_blocks_for_fill(wb, excel_blocks, module, src)
        fill_logs.extend(_fill_table_from_blocks(wb, fill_candidates, raw_only=raw_only, overwrite_nonblank=overwrite_nonblank))

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return fill_logs, match_logs, warnings


def save_logs(fill_logs: list[FillLog], match_logs: list[MatchLog], out_prefix: str | Path) -> tuple[Path, Path]:
    out_prefix = Path(out_prefix)
    fill_path = out_prefix.with_name(out_prefix.stem + '_填充日志.csv')
    match_path = out_prefix.with_name(out_prefix.stem + '_匹配日志.csv')
    fill_path.parent.mkdir(parents=True, exist_ok=True)
    with fill_path.open('w', encoding='utf-8-sig', newline='') as f:
        fields = list(FillLog.__dataclass_fields__.keys())
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for item in fill_logs:
            w.writerow(asdict(item))
    with match_path.open('w', encoding='utf-8-sig', newline='') as f:
        fields = list(MatchLog.__dataclass_fields__.keys())
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for item in match_logs:
            w.writerow(asdict(item))
    return fill_path, match_path
