"""Execute a small Python AST language without eval/exec, imports or object access.

Only assignments, literals and calls to the documented table functions are legal.
The worker additionally runs in a disposable, timed subprocess without credentials.
"""
from __future__ import annotations

import ast
import copy
import json
import re
import sys
import unicodedata
from datetime import datetime
from decimal import Decimal, InvalidOperation

MAX_ROWS = 10000
MAX_CODE = 16000


class SandboxError(ValueError):
    pass


def text(value):
    return '' if value is None else unicodedata.normalize('NFKC', str(value)).strip()


def numeric(value):
    s = text(value)
    if ',' in s:
        if not re.fullmatch(r'[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?', s):
            raise ValueError('千分位分隔符格式不明确')
        s = s.replace(',', '')
    if not re.fullmatch(r'[+-]?\d+(?:\.\d+)?', s):
        raise ValueError('不是明确的十进制数值')
    n = Decimal(s)
    if len(n.as_tuple().digits) > 15 or abs(n) > Decimal('1e12'):
        raise ValueError('超过 Excel 精确数值范围（15 位有效数字、绝对值 ≤ 1e12）')
    return n


def date_value(value):
    s = text(value).replace('年', '-').replace('月', '-').replace('日', '').replace('/', '-').replace('.', '-')
    for fmt in ('%Y-%m-%d', '%Y%m%d', '%Y-%m-%d 00:00:00', '%Y-%m-%dT00:00:00'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    raise ValueError('日期须为明确的年月日，且日期必须存在')


def validate_code(code):
    if not isinstance(code, str) or not code.strip() or len(code) > MAX_CODE:
        raise SandboxError('清洗代码须为 1 至 16000 字符')
    try:
        tree = ast.parse(code)
    except (SyntaxError, RecursionError) as e:
        raise SandboxError('清洗代码语法错误') from e
    if len(list(ast.walk(tree))) > 3000 or len(tree.body) > 100:
        raise SandboxError('代码超过复杂度上限')

    def expression(node, depth=0):
        if depth > 20:
            raise SandboxError('表达式嵌套过深')
        if isinstance(node, ast.Constant) and type(node.value) in (str, int, float, bool, type(None)):
            if isinstance(node.value, str) and len(node.value) > 2000:
                raise SandboxError('单个文本常量过长')
            return
        if isinstance(node, ast.Name) and re.fullmatch(r'[a-zA-Z][a-zA-Z0-9_]{0,40}', node.id):
            return
        if isinstance(node, (ast.List, ast.Tuple)):
            for child in node.elts:
                expression(child, depth + 1)
            return
        if isinstance(node, ast.Dict) and all(k is not None for k in node.keys):
            for child in [*node.keys, *node.values]:
                expression(child, depth + 1)
            return
        if isinstance(node, ast.UnaryOp) and isinstance(node.op, (ast.USub, ast.UAdd)) and isinstance(node.operand, ast.Constant) and type(node.operand.value) in (int, float):
            return
        if isinstance(node, ast.Call) and isinstance(node.func, ast.Name) and node.func.id in FUNCTIONS:
            for child in node.args:
                expression(child, depth + 1)
            for kw in node.keywords:
                if kw.arg is None or kw.arg.startswith('_'):
                    raise SandboxError('不支持展开或私有参数')
                expression(kw.value, depth + 1)
            return
        raise SandboxError('仅允许赋值、字面量和公开表格函数；禁止 import、属性访问、索引、循环和任意函数')

    assigned = set()
    for statement in tree.body:
        if not isinstance(statement, ast.Assign) or len(statement.targets) != 1 or not isinstance(statement.targets[0], ast.Name):
            raise SandboxError('每条语句必须为变量赋值')
        name = statement.targets[0].id
        if not re.fullmatch(r'[a-zA-Z][a-zA-Z0-9_]{0,40}', name) or name in FUNCTIONS:
            raise SandboxError('变量名非法或覆盖表格函数')
        expression(statement.value)
        for child in ast.walk(statement.value):
            if isinstance(child, ast.Name) and child.id not in FUNCTIONS and child.id not in assigned:
                raise SandboxError(f'变量 {child.id} 尚未定义')
        assigned.add(name)
    if 'result' not in assigned:
        raise SandboxError('最后须给 result 赋值')
    return tree


FUNCTIONS = {'table', 'merge', 'map_fields', 'trim', 'replace_text', 'cast', 'filter_rows', 'dedupe', 'split_address', 'fill', 'join', 'select_columns'}


class Runtime:
    def __init__(self, tables):
        self.tables = tables
        self.events = []
        self.functions = {name: getattr(self, name) for name in FUNCTIONS}

    def rows(self, rows):
        if not isinstance(rows, list) or len(rows) > MAX_ROWS or any(not isinstance(r, dict) for r in rows):
            raise SandboxError('表格须为记录列表，且总行数不得超过 10000')
        return rows

    def column(self, name):
        if not isinstance(name, str) or not name or name.startswith('_') or len(name) > 200:
            raise SandboxError('字段名非法；下划线开头的字段为内部溯源保留字段')
        return name

    def columns(self, names):
        if not isinstance(names, list) or not names or len(names) > 60:
            raise SandboxError('列清单须为 1 至 60 个字段')
        return [self.column(c) for c in names]

    def check_columns(self, rows, names):
        self.rows(rows)
        for r in rows:
            for c in names:
                if c not in r:
                    raise SandboxError(f'缺少源字段：{c}')

    def problem(self, row, field, message):
        row.setdefault('_errors', []).append({'column': field, 'error': message})

    def event(self, row, reason):
        self.events.append({'refs': row['_refs'], 'reason': reason})

    def table(self, name):
        if name not in self.tables:
            raise SandboxError(f'不存在的源表 {name}')
        rows = copy.deepcopy(self.tables[name])
        for r in rows:
            r['_before'] = {k: v for k, v in r.items() if not k.startswith('_')}
        return rows

    def merge(self, tables):
        if not isinstance(tables, list):
            raise SandboxError('merge 参数须为表格列表')
        return self.rows([r for rows in tables for r in self.rows(rows)])

    def map_fields(self, rows, mapping, keep=False):
        if not isinstance(mapping, dict) or not mapping or len(mapping) > 60:
            raise SandboxError('映射须为目标字段到源字段的字典')
        for target, source in mapping.items():
            self.column(target)
            self.column(source)
        self.check_columns(rows, mapping.values())
        return [{**(r if keep else {k: v for k, v in r.items() if k.startswith('_')}), '_before': {k: r[v] for k, v in mapping.items()}, **{k: r[v] for k, v in mapping.items()}} for r in rows]

    def trim(self, rows, columns):
        columns = self.columns(columns)
        self.check_columns(rows, columns)
        for r in rows:
            for c in columns:
                if isinstance(r[c], str):
                    r[c] = text(r[c]).replace('\u200b', '').replace('\ufeff', '')
        return rows

    def replace_text(self, rows, column, replacements):
        self.column(column)
        self.check_columns(rows, [column])
        if not isinstance(replacements, dict) or len(replacements) > 30 or any(not isinstance(k, str) or not k or not isinstance(v, str) for k, v in replacements.items()):
            raise SandboxError('替换须为非空文本到文本的字典')
        for r in rows:
            value = text(r[column])
            for before, after in replacements.items():
                value = value.replace(before, after)
                if len(value) > 10000:
                    raise SandboxError('替换结果过长')
            r[column] = value
        return rows

    def cast(self, rows, column, kind, format='%Y-%m-%d'):
        self.column(column)
        self.check_columns(rows, [column])
        if kind not in ('text', 'decimal', 'integer', 'date'):
            raise SandboxError('类型仅支持 text/decimal/integer/date')
        if format not in ('%Y-%m-%d', '%Y%m%d', '%Y/%m/%d'):
            raise SandboxError('日期格式不支持')
        for r in rows:
            try:
                value = r[column]
                if value is None or value == '':
                    continue
                if kind == 'text':
                    if not isinstance(value, str):
                        raise ValueError('数值转文本可能丢失工号前导零，请人工核对')
                    r[column] = text(value)
                elif kind == 'date':
                    r[column] = date_value(value).strftime(format)
                else:
                    n = numeric(value)
                    if kind == 'integer' and n != n.to_integral_value():
                        raise ValueError('不是整数')
                    r[column] = str(n)
                r.setdefault('_types', {})[column] = kind
            except (ValueError, InvalidOperation) as e:
                self.problem(r, column, str(e))
        return rows

    def filter_rows(self, rows, column, op, value, numeric_compare=False):
        self.column(column)
        self.check_columns(rows, [column])
        if op not in ('eq', 'ne', 'lt', 'le', 'gt', 'ge', 'in', 'not_in', 'contains', 'not_empty'):
            raise SandboxError('筛选运算符不支持')
        result = []
        for r in rows:
            try:
                a, b = r[column], value
                if numeric_compare:
                    a, b = numeric(a), numeric(b)
                comparisons = {
                    'eq': lambda: a == b, 'ne': lambda: a != b, 'lt': lambda: a < b,
                    'le': lambda: a <= b, 'gt': lambda: a > b, 'ge': lambda: a >= b,
                    'in': lambda: a in b, 'not_in': lambda: a not in b,
                    'contains': lambda: str(b) in str(a), 'not_empty': lambda: bool(text(a)),
                }
                if comparisons[op]():
                    result.append(r)
                else:
                    self.event(r, f'筛选排除：{column} {op} {value}')
            except (ValueError, TypeError, InvalidOperation):
                r.setdefault('_errors', []).append({'column': column, 'error': '无法判断筛选条件，请修正源数据后重传或跳过本行', 'skip_only': True})
                result.append(r)
        return result

    def dedupe(self, rows, keys):
        keys = self.columns(keys)
        self.check_columns(rows, keys)
        def comparable(r, k):
            if r.get('_types', {}).get(k) in ('decimal', 'integer') and text(r[k]):
                return format(numeric(r[k]).normalize(), 'f')
            return r[k]
        groups = {}
        for r in rows:
            key = json.dumps([comparable(r, k) for k in keys], ensure_ascii=False, sort_keys=True)
            groups.setdefault(key, []).append(r)
        result = []
        for group in groups.values():
            unique = {}
            for r in group:
                value = json.dumps({k: comparable(r, k) for k in r if not k.startswith('_')}, sort_keys=True, ensure_ascii=False)
                if value in unique:
                    self.event(r, '完全重复记录已去重')
                else:
                    unique[value] = r
            if len(unique) > 1:
                for r in unique.values():
                    r['_conflict_keys'] = keys
            result.extend(unique.values())
        return result

    def split_address(self, rows, column, province, city, district):
        self.check_columns(rows, [self.column(column)])
        for c in (province, city, district):
            self.column(c)
        for r in rows:
            address = text(r[column])
            # Deliberately no guessed administrative-region completion.
            m = re.match(r'^(北京市|上海市|天津市|重庆市)(.+?(?:区|县))', address)
            if m:
                r[province], r[city], r[district] = m[1], m[1], m[2]
                continue
            m = re.match(r'^(.+?(?:省|自治区))(.+?(?:市|自治州|地区|盟))(.+?(?:区|县|旗|市))', address)
            if m:
                r[province], r[city], r[district] = m.groups()
            else:
                r[province], r[city], r[district] = '', '', ''
                self.problem(r, district, '地址无法明确拆为省市区，请填写实际省市区或跳过')
        return rows

    def fill(self, rows, column, value, only_empty=True):
        self.column(column)
        self.rows(rows)
        if type(value) not in (str, int, float, bool, type(None)):
            raise SandboxError('填充值须为标量')
        for r in rows:
            if not only_empty or not text(r.get(column)):
                r[column] = value
        return rows

    def join(self, left, right, left_key, right_key, fields):
        self.check_columns(left, [self.column(left_key)])
        self.check_columns(right, [self.column(right_key)])
        if not isinstance(fields, dict) or not fields or len(fields) > 60:
            raise SandboxError('关联字段须为目标字段到右表字段的字典')
        for target, source in fields.items():
            self.column(target)
            self.column(source)
        self.check_columns(right, fields.values())
        lookup = {}
        for r in right:
            lookup.setdefault(text(r[right_key]), []).append(r)
        for r in left:
            found = lookup.get(text(r[left_key]), []) if text(r[left_key]) else []
            if len(found) != 1:
                for target in fields:
                    self.problem(r, target, '精确关联缺失或存在多个匹配，请人工核对')
                    r[target] = ''
            else:
                r.update({target: found[0][source] for target, source in fields.items()})
                r.setdefault('_before', {}).update({target: found[0][source] for target, source in fields.items()})
                r['_refs'] += copy.deepcopy(found[0]['_refs'])
        return left

    def select_columns(self, rows, columns):
        columns = self.columns(columns)
        self.check_columns(rows, columns)
        return [{**{k: v for k, v in r.items() if k.startswith('_')}, **{c: r[c] for c in columns}} for r in rows]

    def run(self, code):
        tree = validate_code(code)
        variables = {}

        def evaluate(node):
            if isinstance(node, ast.Constant):
                return node.value
            if isinstance(node, ast.Name):
                return variables[node.id]
            if isinstance(node, (ast.List, ast.Tuple)):
                return [evaluate(child) for child in node.elts]
            if isinstance(node, ast.Dict):
                return {evaluate(k): evaluate(v) for k, v in zip(node.keys, node.values)}
            if isinstance(node, ast.UnaryOp):
                return -node.operand.value if isinstance(node.op, ast.USub) else node.operand.value
            if isinstance(node, ast.Call):
                return self.functions[node.func.id](*[evaluate(a) for a in node.args], **{k.arg: evaluate(k.value) for k in node.keywords})
            raise SandboxError('表达式被拒绝')

        for statement in tree.body:
            variables[statement.targets[0].id] = evaluate(statement.value)
        result = self.rows(variables['result'])
        if any('_refs' not in r or len(r) > 90 for r in result):
            raise SandboxError('结果丢失溯源信息或列数超限')
        return {'rows': result, 'events': self.events}


def worker():
    try:
        # Windows consoles may default to a legacy code page; IPC is always UTF-8.
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding='utf-8')
        try:
            import resource
            resource.setrlimit(resource.RLIMIT_CPU, (8, 8))
            resource.setrlimit(resource.RLIMIT_AS, (768 * 1024 * 1024, 768 * 1024 * 1024))
        except (ImportError, ValueError, OSError):
            pass
        payload = json.loads(sys.stdin.buffer.read(24 * 1024 * 1024))
        output = Runtime(payload['tables']).run(payload['code'])
        chunks = []
        size = 0
        for chunk in json.JSONEncoder(ensure_ascii=False, allow_nan=False).iterencode(output):
            size += len(chunk.encode('utf-8'))
            if size > 24 * 1024 * 1024:
                raise SandboxError('输出超过 24 MB')
            chunks.append(chunk)
        sys.stdout.write(''.join(chunks))
    except Exception as e:
        sys.stdout.write(json.dumps({'error': str(e)[:1000]}, ensure_ascii=False))
        sys.exit(1)


if __name__ == '__main__':
    worker()
