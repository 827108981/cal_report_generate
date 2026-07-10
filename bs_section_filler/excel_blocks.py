from __future__ import annotations

import math
import re
import statistics
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.utils.cell import coordinate_to_tuple, range_boundaries

from .utils import canonical_section, compact_texts, format_number_by_excel_format, is_blank_value, looks_like_section, normalize_text


@dataclass
class SourceCell:
    value: Any
    row: int
    col: int
    is_formula: bool = False
    number_format: str = 'General'

    @property
    def text(self) -> str:
        return format_number_by_excel_format(self.value, self.number_format)

    @property
    def norm(self) -> str:
        return normalize_text(self.text)


@dataclass
class ExcelBlock:
    module: str
    sheet: str
    section: str
    title: str
    object_name: str
    start_row: int
    end_row: int
    rows: list[list[SourceCell]]
    header_norms: set[str]

    @property
    def key_text(self) -> str:
        return ' '.join(x for x in [self.section, self.title, self.object_name] if x)


def _cell_is_formula(cell: Cell) -> bool:
    return isinstance(cell.value, str) and cell.value.startswith('=')


def _to_number(value: Any) -> float | None:
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    try:
        if text.endswith('%'):
            return float(text[:-1]) / 100
        return float(text)
    except ValueError:
        return None


def _split_args(text: str) -> list[str]:
    args: list[str] = []
    start = 0
    depth = 0
    in_quote = False
    for i, ch in enumerate(text):
        if ch == '"':
            in_quote = not in_quote
        elif not in_quote:
            if ch == '(':
                depth += 1
            elif ch == ')':
                depth -= 1
            elif ch == ',' and depth == 0:
                args.append(text[start:i].strip())
                start = i + 1
    args.append(text[start:].strip())
    return args


def _build_formula_evaluator(ws_val, ws_formula):
    cache: dict[tuple[int, int], Any] = {}
    visiting: set[tuple[int, int]] = set()

    def strip_outer_parens(expr: str) -> str:
        e = expr.strip()
        changed = True
        while changed and e.startswith('(') and e.endswith(')'):
            changed = False
            depth = 0
            in_quote = False
            for i, ch in enumerate(e):
                if ch == '"':
                    in_quote = not in_quote
                elif not in_quote:
                    if ch == '(':
                        depth += 1
                    elif ch == ')':
                        depth -= 1
                        if depth == 0 and i != len(e) - 1:
                            return e
            if depth == 0:
                e = e[1:-1].strip()
                changed = True
        return e

    def cell_value(ref: str) -> Any:
        row, col = coordinate_to_tuple(ref.replace('$', ''))
        return evaluate(row, col)

    def cell_number(ref: str) -> float | None:
        return _to_number(cell_value(ref))

    def range_numbers(ref: str) -> list[float]:
        min_col, min_row, max_col, max_row = range_boundaries(ref.replace('$', ''))
        nums: list[float] = []
        for rr in range(min_row, max_row + 1):
            for cc in range(min_col, max_col + 1):
                n = _to_number(evaluate(rr, cc))
                if n is not None:
                    nums.append(n)
        return nums

    def range_values(ref: str) -> list[Any]:
        min_col, min_row, max_col, max_row = range_boundaries(ref.replace('$', ''))
        values: list[Any] = []
        for rr in range(min_row, max_row + 1):
            for cc in range(min_col, max_col + 1):
                values.append(evaluate(rr, cc))
        return values

    def range_column_numbers(ref: str) -> list[float]:
        min_col, min_row, max_col, max_row = range_boundaries(ref.replace('$', ''))
        nums: list[float] = []
        for rr in range(min_row, max_row + 1):
            for cc in range(min_col, max_col + 1):
                n = _to_number(evaluate(rr, cc))
                if n is not None:
                    nums.append(n)
        return nums

    def average(ref: str) -> float | None:
        nums = range_numbers(ref)
        return sum(nums) / len(nums) if nums else None

    def stdev(ref: str) -> float | None:
        nums = range_numbers(ref)
        if len(nums) < 2:
            return None
        return statistics.stdev(nums)

    def slope(y_ref: str, x_ref: str) -> float | None:
        ys = range_column_numbers(y_ref)
        xs = range_column_numbers(x_ref)
        pairs = [(x, y) for x, y in zip(xs, ys) if x is not None and y is not None]
        if len(pairs) < 2:
            return None
        mean_x = sum(x for x, _ in pairs) / len(pairs)
        mean_y = sum(y for _, y in pairs) / len(pairs)
        den = sum((x - mean_x) ** 2 for x, _ in pairs)
        if den == 0:
            return None
        return sum((x - mean_x) * (y - mean_y) for x, y in pairs) / den

    def intercept(y_ref: str, x_ref: str) -> float | None:
        m = slope(y_ref, x_ref)
        if m is None:
            return None
        ys = range_column_numbers(y_ref)
        xs = range_column_numbers(x_ref)
        if not xs or not ys:
            return None
        return (sum(ys) / len(ys)) - m * (sum(xs) / len(xs))

    def correl(xs: list[float], ys: list[float]) -> float | None:
        pairs = [(x, y) for x, y in zip(xs, ys)]
        if len(pairs) < 2:
            return None
        mean_x = sum(x for x, _ in pairs) / len(pairs)
        mean_y = sum(y for _, y in pairs) / len(pairs)
        sx = math.sqrt(sum((x - mean_x) ** 2 for x, _ in pairs))
        sy = math.sqrt(sum((y - mean_y) ** 2 for _, y in pairs))
        if sx == 0 or sy == 0:
            return None
        return sum((x - mean_x) * (y - mean_y) for x, y in pairs) / (sx * sy)

    def split_top_level_operator(expr: str, ops: set[str]) -> tuple[str, str, str] | None:
        depth = 0
        in_quote = False
        for i in range(len(expr) - 1, -1, -1):
            ch = expr[i]
            if ch == '"':
                in_quote = not in_quote
                continue
            if in_quote:
                continue
            if ch == ')':
                depth += 1
                continue
            if ch == '(':
                depth -= 1
                continue
            if depth != 0 or ch not in ops:
                continue
            prev = expr[i - 1] if i > 0 else ''
            if ch in '+-' and (i == 0 or prev in '+-*/(<>=,'):
                continue
            return expr[:i], ch, expr[i + 1:]
        return None

    def split_top_level_comparison(expr: str) -> tuple[str, str, str] | None:
        depth = 0
        in_quote = False
        i = 0
        while i < len(expr):
            ch = expr[i]
            if ch == '"':
                in_quote = not in_quote
                i += 1
                continue
            if in_quote:
                i += 1
                continue
            if ch == '(':
                depth += 1
                i += 1
                continue
            if ch == ')':
                depth -= 1
                i += 1
                continue
            if depth == 0:
                two = expr[i:i + 2]
                if two in {'<=', '>='}:
                    return expr[:i], two, expr[i + 2:]
                if ch in {'=', '<', '>'}:
                    return expr[:i], ch, expr[i + 1:]
            i += 1
        return None

    def countif(range_ref: str, criterion_expr: str) -> int:
        criterion = scalar_expr(criterion_expr)
        if criterion is None:
            criterion = criterion_expr.strip().strip('"')
        if isinstance(criterion, str):
            criterion_text = criterion.strip()
        else:
            criterion_text = format_number_by_excel_format(criterion)

        op = '='
        right_text = criterion_text
        for prefix in ('<=', '>=', '<>', '<', '>', '='):
            if criterion_text.startswith(prefix):
                op = prefix
                right_text = criterion_text[len(prefix):].strip()
                break

        if right_text == '':
            return sum(1 for value in range_values(range_ref) if is_blank_value(value))

        right_num = _to_number(right_text)
        total = 0
        for value in range_values(range_ref):
            if op == '<>':
                if is_blank_value(value):
                    total += 1
                elif right_num is not None:
                    left_num = _to_number(value)
                    if left_num is None or left_num != right_num:
                        total += 1
                elif str(value).strip() != right_text:
                    total += 1
                continue

            if right_num is not None:
                left_num = _to_number(value)
                if left_num is None:
                    continue
                ok = (
                    (op == '=' and left_num == right_num)
                    or (op == '<' and left_num < right_num)
                    or (op == '>' and left_num > right_num)
                    or (op == '<=' and left_num <= right_num)
                    or (op == '>=' and left_num >= right_num)
                )
            else:
                left_text = '' if is_blank_value(value) else str(value).strip()
                ok = (op == '=' and left_text == right_text)
            if ok:
                total += 1
        return total

    def scalar_expr(expr: str) -> Any:
        e = strip_outer_parens(expr)
        if e.startswith('"') and e.endswith('"'):
            return e[1:-1]
        if re.fullmatch(r'-?\d+(?:\.\d+)?%?', e):
            return float(e[:-1]) / 100 if e.endswith('%') else float(e)
        split = split_top_level_operator(e, {'+', '-'})
        if split:
            left_expr, op, right_expr = split
            left = _to_number(scalar_expr(left_expr))
            right = _to_number(scalar_expr(right_expr))
            if left is None or right is None:
                return None
            return left + right if op == '+' else left - right
        split = split_top_level_operator(e, {'*', '/'})
        if split:
            left_expr, op, right_expr = split
            left = _to_number(scalar_expr(left_expr))
            right = _to_number(scalar_expr(right_expr))
            if left is None or right is None:
                return None
            if op == '/':
                return left / right if right != 0 else None
            return left * right
        if re.fullmatch(r'\$?[A-Z]{1,3}\$?\d+', e):
            return cell_value(e)
        m = re.fullmatch(r'ABS\((.+)\)', e, flags=re.I)
        if m:
            n = _to_number(scalar_expr(m.group(1)))
            return abs(n) if n is not None else None
        m = re.fullmatch(r'AVERAGE\(([^)]+)\)', e, flags=re.I)
        if m:
            return average(m.group(1))
        m = re.fullmatch(r'STDEV(?:\.S)?\(([^)]+)\)', e, flags=re.I)
        if m:
            return stdev(m.group(1))
        m = re.fullmatch(r'MAX\(([^)]+)\)', e, flags=re.I)
        if m:
            nums = range_numbers(m.group(1))
            return max(nums) if nums else None
        m = re.fullmatch(r'MIN\(([^)]+)\)', e, flags=re.I)
        if m:
            nums = range_numbers(m.group(1))
            return min(nums) if nums else None
        m = re.fullmatch(r'COUNTIF\(([^,]+),\s*(.+)\)', e, flags=re.I)
        if m:
            return countif(m.group(1), m.group(2))
        return None

    def compare_expr(condition: str) -> bool:
        c = strip_outer_parens(condition)
        split = split_top_level_comparison(c)
        if not split:
            return bool(scalar_expr(c))
        left = scalar_expr(split[0])
        right = scalar_expr(split[2])
        op = split[1]
        if op == '=' and (_to_number(left) is None or _to_number(right) is None):
            left_text = '' if is_blank_value(left) else str(left).strip()
            right_text = '' if is_blank_value(right) else str(right).strip()
            return left_text == right_text
        ln = _to_number(left)
        rn = _to_number(right)
        if ln is None or rn is None:
            return False
        if op == '<=':
            return ln <= rn
        if op == '>=':
            return ln >= rn
        if op == '<':
            return ln < rn
        if op == '>':
            return ln > rn
        return ln == rn

    def condition_expr(expr: str) -> bool:
        e = strip_outer_parens(expr)
        upper = e.upper()
        if upper.startswith('AND(') and e.endswith(')'):
            return all(condition_expr(arg) for arg in _split_args(e[4:-1]))
        if upper.startswith('OR(') and e.endswith(')'):
            return any(condition_expr(arg) for arg in _split_args(e[3:-1]))
        return compare_expr(e)

    def evaluate_formula(formula: str) -> Any:
        f = formula.strip()
        if f.startswith('='):
            f = f[1:].strip()

        m = re.fullmatch(r'AVERAGE\(([^)]+)\)', f, flags=re.I)
        if m:
            return average(m.group(1))
        m = re.fullmatch(r'STDEV(?:\.S)?\(([^)]+)\)', f, flags=re.I)
        if m:
            return stdev(m.group(1))
        m = re.fullmatch(r'MIN\(([^)]+)\)', f, flags=re.I)
        if m:
            nums = range_numbers(m.group(1))
            return min(nums) if nums else None
        m = re.fullmatch(r'MAX\(([^)]+)\)', f, flags=re.I)
        if m:
            nums = range_numbers(m.group(1))
            return max(nums) if nums else None
        m = re.fullmatch(r'COUNTIF\(([^,]+),\s*(.+)\)', f, flags=re.I)
        if m:
            return countif(m.group(1), m.group(2))
        m = re.fullmatch(r'POWER\(([^,]+),\s*([^)]+)\)', f, flags=re.I)
        if m:
            base = _to_number(scalar_expr(m.group(1)))
            exp = _to_number(scalar_expr(m.group(2)))
            return base ** exp if base is not None and exp is not None else None
        m = re.fullmatch(r'SLOPE\(([^,]+),\s*([^)]+)\)', f, flags=re.I)
        if m:
            return slope(m.group(1), m.group(2))
        m = re.fullmatch(r'INTERCEPT\(([^,]+),\s*([^)]+)\)', f, flags=re.I)
        if m:
            return intercept(m.group(1), m.group(2))
        m = re.fullmatch(r'(\$?[A-Z]{1,3}\$?\d+)\s*/\s*(\$?[A-Z]{1,3}\$?\d+)', f, flags=re.I)
        if m:
            left = cell_number(m.group(1))
            right = cell_number(m.group(2))
            return left / right if left is not None and right not in (None, 0) else None
        m = re.fullmatch(r'CORREL\(([^,]+),\s*\(([^)]+)\)\s*\*\s*(\$?[A-Z]{1,3}\$?\d+)\s*\+\s*(\$?[A-Z]{1,3}\$?\d+)\)', f, flags=re.I)
        if m:
            ys = range_column_numbers(m.group(1))
            xs = range_column_numbers(m.group(2))
            a = cell_number(m.group(3))
            b = cell_number(m.group(4))
            if a is None or b is None:
                return None
            predicted = [x * a + b for x in xs]
            return correl(ys, predicted)
        m = re.fullmatch(r'ABS\(\((\$?[A-Z]{1,3}\$?\d+)\s*-\s*\((\$?[A-Z]{1,3}\$?\d+)\s*\*\s*(\$?[A-Z]{1,3}\$?\d+)\s*\+\s*(\$?[A-Z]{1,3}\$?\d+)\)\)\s*/\s*\(\2\s*\*\s*\3\s*\+\s*\4\)\)', f, flags=re.I)
        if m:
            actual = cell_number(m.group(1))
            x = cell_number(m.group(2))
            a = cell_number(m.group(3))
            b = cell_number(m.group(4))
            if actual is None or x is None or a is None or b is None:
                return None
            expected = x * a + b
            return abs((actual - expected) / expected) if expected else None
        if f.upper().startswith('IF(') and f.endswith(')'):
            args = _split_args(f[3:-1])
            if len(args) >= 3:
                return scalar_expr(args[1]) if condition_expr(args[0]) else scalar_expr(args[2])
        value = scalar_expr(f)
        if value is not None:
            return value
        if '&' in f:
            parts = _split_formula_concat(f)
            if parts:
                return ''.join(format_number_by_excel_format(scalar_expr(part)) for part in parts)
        return None

    def evaluate(r: int, c: int) -> Any:
        key = (r, c)
        if key in cache:
            return cache[key]
        raw = ws_val.cell(r, c).value
        formula = ws_formula.cell(r, c).value
        if not (isinstance(formula, str) and formula.startswith('=')):
            cache[key] = raw
            return raw
        compact_formula = re.sub(r'\s+', '', formula.upper())
        if (
            isinstance(raw, str)
            and raw.strip().upper() in {'PASS', 'FAIL'}
            and compact_formula.startswith('=IF(AND(')
            and '>=' in compact_formula
            and '<=' in compact_formula
            and 'COUNTIF' not in compact_formula
            and 'MAX(' not in compact_formula
            and 'MIN(' not in compact_formula
        ):
            cache[key] = raw
            return raw
        if key in visiting:
            return raw if not is_blank_value(raw) else None
        visiting.add(key)
        try:
            value = evaluate_formula(formula)
            if value is not None:
                cache[key] = value
                return value
            cache[key] = raw
            return raw
        finally:
            visiting.discard(key)

    return evaluate


def _split_formula_concat(text: str) -> list[str]:
    parts: list[str] = []
    start = 0
    in_quote = False
    for i, ch in enumerate(text):
        if ch == '"':
            in_quote = not in_quote
        elif ch == '&' and not in_quote:
            parts.append(text[start:i].strip())
            start = i + 1
    parts.append(text[start:].strip())
    return parts


def _compress_row(ws_val, ws_formula, r: int, max_col: int, evaluator=None) -> list[SourceCell]:
    out: list[SourceCell] = []
    for c in range(1, max_col + 1):
        v = ws_val.cell(r, c).value
        f = ws_formula.cell(r, c).value
        if evaluator is not None and isinstance(f, str) and f.startswith('='):
            evaluated = evaluator(r, c)
            if evaluated is not None:
                v = evaluated
        if is_blank_value(v):
            continue
        out.append(SourceCell(value=v, row=r, col=c, is_formula=isinstance(f, str) and f.startswith('='), number_format=ws_val.cell(r, c).number_format or 'General'))
    return out


def _row_text(row: list[SourceCell]) -> str:
    return ' '.join(c.text for c in row if c.text)


def _module_from_text(text: str) -> str | None:
    s = normalize_text(text)
    if '模块二' in s or 'M2模块' in s or '模块2' in s or 'ISE模块二' in s:
        return 'M2'
    if '模块一' in s or 'M1模块' in s or '模块1' in s or 'ISE模块一' in s:
        return 'M1'
    return None


def _is_module_marker(row: list[SourceCell]) -> bool:
    texts = compact_texts(c.text for c in row)
    if len(texts) != 1:
        return False
    return _module_from_text(texts[0]) is not None


def _is_anchor(row: list[SourceCell]) -> bool:
    if not row:
        return False
    if _is_module_marker(row):
        return False
    texts = [c.text for c in row]
    first = texts[0]
    joined = _row_text(row)
    n = normalize_text(joined)

    # 明确排除：表头/注释行不应被拆成新数据块。
    if first.startswith('注') or any(x in first for x in ['测试次数', '测试项目']):
        return False

    # 明确对象行：试剂名称 / 样本针 / 试剂针 / 某个线性值标题
    if '试剂名称' in joined:
        return True
    if any(k in joined for k in [
        '测试结果', '吸光度线性值', '吸光度准确性', '暗电流测试', '真空及压力检测',
        '工作环境检测', '仪器维护保养', '仪器基本状态检查', '主机位置校准', 'SDM位置校准',
        '反应盘温度检测', '样本携带污染率检测', '电解质准确度', '电解质精密度',
        '电解质线性检测', '线性范围验证', '电解质稳定性', '电解质携带污染率',
        '项目交叉污染', '项目测试时间', '杂散光', '亚硝酸钠'
    ]):
        return True
    # 排除纯标签/表头/批号/统计结果行，避免把同一张表拆成多个块。
    # 旧逻辑中“统计：R²=...”“Mean/SD/CV”“携带污染率CLH”等单行结果被当成新表，
    # 导致 Word 后半列/后半段无法填充。现在只保留上面的明确对象行作为 anchor。
    if len(texts) == 1:
        if any(x in first for x in ['批号', '重复次数', '测试数据', '次数', '采光周期', '要求', '结论', '指标', '统计', 'R²', 'R：', 'a：', 'b：', 'Mean', 'SD', 'CV', '携带污染率CLH', '携带污染率CHL']):
            return False
        # 单行说明/对象行也作为块起点。v4 为了避免统计行拆块取消了兜底，
        # 导致“杂散光检测”等只有说明文字开头的表格匹配不到；这里在排除统计/表头后恢复兜底。
        if len(n) >= 4 and not looks_like_section(first):
            return True
    return False


def _clean_object_title(text: str) -> str:
    s = str(text or '').strip()
    # 说明性段落第一行通常是表名，如“1、电解质准确度”。
    if '\n' in s:
        s = s.split('\n', 1)[0].strip()
    return s


def _extract_object_name(row: list[SourceCell]) -> str:
    texts = compact_texts(c.text for c in row)
    if not texts:
        return ''
    joined = ' '.join(texts)
    if '试剂名称' in joined:
        vals = [t for t in texts if '试剂名称' not in t]
        return vals[-1] if vals else ''
    return _clean_object_title(texts[0])


def _extract_title(row: list[SourceCell], section: str) -> str:
    texts = compact_texts(c.text for c in row)
    if not texts:
        return section
    first = texts[0]
    if '试剂名称' in first and len(texts) > 1:
        return section
    return _clean_object_title(first)


def _header_norms(rows: list[list[SourceCell]]) -> set[str]:
    headers: set[str] = set()
    header_keywords = ['次数', '重复次数', '测试数据', '采光周期', '检测项目', '要求', '结论', '吸光度', '结果', '均值', 'SD', 'CV', 'MAX', 'MIN', '指标', '温度值', '参数值', '校准状态']
    for row in rows[:5]:
        for cell in row:
            t = cell.text
            if any(k in t for k in header_keywords):
                n = normalize_text(t)
                if n:
                    headers.add(n)
    return headers


def read_excel_blocks(path: str | Path, module: str = 'M1') -> list[ExcelBlock]:
    """按“章节 -> 对象/名称 -> 表格”提取 Excel 数据块。

    注意：这里不是全局项目名匹配；每个块都带章节和对象名，后续 Word 也按相同层级匹配。
    """
    path = str(path)
    wb_val = load_workbook(path, data_only=True)
    wb_formula = load_workbook(path, data_only=False)
    blocks: list[ExcelBlock] = []

    for ws_val in wb_val.worksheets:
        ws_formula = wb_formula[ws_val.title]
        max_col = ws_val.max_column or 1
        max_row = ws_val.max_row or 1
        evaluator = _build_formula_evaluator(ws_val, ws_formula)
        compressed: dict[int, list[SourceCell]] = {
            r: _compress_row(ws_val, ws_formula, r, max_col, evaluator)
            for r in range(1, max_row + 1)
        }

        current_section = ''
        current_module = module
        section_at_row: dict[int, str] = {}
        module_at_row: dict[int, str] = {}
        anchors: list[int] = []

        for r in range(1, max_row + 1):
            row = compressed[r]
            if not row:
                section_at_row[r] = current_section
                module_at_row[r] = current_module
                continue
            first = row[0].text
            row_module = _module_from_text(_row_text(row)) if _is_module_marker(row) else None
            if row_module:
                current_module = row_module
                section_at_row[r] = current_section
                module_at_row[r] = current_module
                continue
            if looks_like_section(first):
                current_section = canonical_section(first)
                section_at_row[r] = current_section
                module_at_row[r] = current_module
                continue
            section_at_row[r] = current_section
            module_at_row[r] = current_module
            if _is_anchor(row):
                # 避免把说明性长段落当成表格：下一行或当前行需要有至少两个有效单元，或标题属于明确表名
                anchors.append(r)

        # section row 到下一个 section row 也算边界
        section_rows = [r for r in range(1, max_row + 1) if compressed[r] and looks_like_section(compressed[r][0].text)]

        for idx, start in enumerate(anchors):
            # 结束到下一个 anchor 或下一个 section 前一行
            candidates = [x for x in anchors[idx+1:] if x > start]
            candidates += [x for x in section_rows if x > start]
            end = min(candidates) - 1 if candidates else max_row
            # 去掉尾部空白
            while end > start and not compressed[end]:
                end -= 1
            rows = [compressed[r] for r in range(start, end + 1) if compressed[r]]
            if not rows:
                continue
            section = section_at_row.get(start, '')
            title = _extract_title(rows[0], section)
            obj = _extract_object_name(rows[0])
            blocks.append(ExcelBlock(
                module=module_at_row.get(start, module),
                sheet=ws_val.title,
                section=section,
                title=title,
                object_name=obj,
                start_row=start,
                end_row=end,
                rows=rows,
                header_norms=_header_norms(rows),
            ))

    # 去重：同一 start/end 只保留一次
    uniq: list[ExcelBlock] = []
    seen = set()
    for b in blocks:
        key = (b.sheet, b.start_row, b.end_row, b.section, normalize_text(b.object_name), normalize_text(b.title))
        if key in seen:
            continue
        seen.add(key)
        uniq.append(b)
    return uniq
