# -*- coding: utf-8 -*-
# =============================================================================
#  ExcelToJsonCs.py
#  双击本文件（或打包后的 exe）-> 扫描配置的 Excel 源目录下所有 Excel 族文件
#  (csv / xls / xlsx / xlsm) -> 每个文件导出一份 JSON + 一份 C# 解析类。
#
#  Excel 约定（以 Role.xlsx 为例）：
#   第 1 行：字段英文名        (Id, Name, Cast, Ability_1 ...)
#   第 2 行：字段类型          (int / string / int[] / string[] / List<int> / List<string>)
#   第 3 行：中文注释          (用于 C# 字段后的 // 注释)
#   第 4 行起：数据
#   数组(int[]/string[]/List<int>/List<string>)：横向占「连续多列」，每列一个元素
#                         续列表头必须留空（工具靠表头为空判断数组跨列边界）。
#   注释行：某行第 1 列（最左列）首字符为 '#' -> 忽略该行
#   注释列：某列第 1 行（表头）首字符为 '#' -> 忽略该列
#   只读取第一个 Sheet（csv 只有一张表）。
#
#  JSON 文件名 = 表名(Excel文件名) + "Category"  ->  RoleCategory.json
#  C#   文件名 = 表名(Excel文件名)               ->  Role.cs
#  C# 类名    = RoleCategory（包装类） / Role（行类）
#
#  配置：exe 同目录 config.json（首次运行自动生成）
#  缓存：.excel_export_cache.json 记录各文件 hash，无变化则跳过导出
#  日志：控制台 + 同目录 export_log.txt（每次运行覆盖）
#
#  ---------------------------------------------------------------------------
#  校验与日志（本次新增）
#   分级：ERROR（阻断/跳过，避免生成错误数据或错误代码） / WARN（可疑但不阻断）
#   定位：每条提示包含 [表名] + Sheet + 行号 + 列号(字母) + 字段名 + 具体原因
#   规则：
#     - 主键列（第 1 个字段）为空 / 非整数 / 重复        -> ERROR，跳过该行
#     - 字段类型缺失、未知类型、重复字段名、字段名为非法C#标识符/保留字、
#       首个字段为数组                                   -> ERROR，跳过整个文件
#     - int 列填了非数字 / 小数(截断)                    -> WARN
#     - int 超出 C# Int32 取值范围 [-]                   -> WARN
#     - 字符串含控制字符、数据区中间出现空行、
#       表头为空但类型行有值(游离列)、数据超出表头定义列  -> WARN
#     - 引用不存在（REF_RULES 配置）                     -> ERROR
# =============================================================================

# ====================== 配置区 ======================
DEFAULT_CONFIG = {
    # Excel 源目录。相对路径相对于「本工具所在目录」，绝对路径直接使用。默认 "." 即工具同目录
    "EXCEL_SOURCE_DIR": ".",
    # JSON 输出目录：相对「本工具所在目录」
    "JSON_OUTPUT_DIR": "Json",
    # C# 输出目录：相对「本工具所在目录」
    "CS_OUTPUT_DIR": "Cs",
    # 是否为 Unity 项目生成相关代码（using UnityEngine; [SerializeField] 等）；非 Unity 改 False
    "USE_UNITY": True,
    # 是否使用 Newtonsoft.Json（using Newtonsoft.Json; [JsonIgnore] 等）；用 System.Text.Json 改 False
    "USE_NEWTONSOFT": True,
    # 是否为 Category 包装类添加 [Config] 特性（依赖你的框架；不需要改 False）
    "USE_CONFIG_ATTR": True,

    # ---- 校验相关（本次新增） ----
    # 日志文件名（相对工具目录，"" 表示不写文件，仅控制台输出）
    "LOG_FILE": "export_log.txt",
    # 是否校验 int 是否落在 C# Int32 范围内
    "INT_RANGE_CHECK": True,
    # 单文件最多输出的同类提示条数（防止刷屏），超出部分折叠为一条汇总
    "MAX_ISSUES_PER_FILE": 300,
    # 额外必填字段：{"表名": ["字段名1", "字段名2"]}，命中为空时 WARN（主键列始终必填）
    "REQUIRED_FIELDS": {},
    # 跨表引用校验规则：{"源表名.字段名": "目标表名"} 或 {"源表名.字段名": "目标表名.字段名"}
    #   未写目标字段时，默认校验目标表的「主键列」（第 1 个字段）。
    #   例：{"Role.Cast": "Trait", "Role.SepName": "Nation.Id"}
    "REF_RULES": {},
}
CONFIG_FILENAME = "config.json"
CACHE_FILENAME = ".excel_export_cache.json"
# =================================================


import os
import re
import sys
import json
import hashlib

# ------------------------------ 常量与基础工具 ------------------------------
LEVEL_ERROR = "ERROR"
LEVEL_WARN = "WARN"

# C# Int32 取值范围
INT32_MIN = -2147483648
INT32_MAX = 2147483647

# 合法 C# 标识符 / 保留字
IDENT_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")
CS_KEYWORDS = {
    "abstract", "as", "base", "bool", "break", "byte", "case", "catch", "char",
    "checked", "class", "const", "continue", "decimal", "default", "delegate",
    "do", "double", "else", "enum", "event", "explicit", "extern", "false",
    "finally", "fixed", "float", "for", "foreach", "goto", "if", "implicit",
    "in", "int", "interface", "internal", "is", "lock", "long", "namespace",
    "new", "null", "object", "operator", "out", "override", "params", "private",
    "protected", "public", "readonly", "ref", "return", "sbyte", "sealed",
    "short", "sizeof", "stackalloc", "static", "string", "struct", "switch",
    "this", "throw", "true", "try", "typeof", "uint", "ulong", "unchecked",
    "unsafe", "ushort", "using", "virtual", "void", "volatile", "while",
}

# 全部合法类型（均为「归一化」后的小写、无空格形式）
VALID_TYPES = ("int", "string", "int[]", "string[]", "list<int>", "list<string>")
VALID_TYPES_DISPLAY = "int, string, int[], string[], List<int>, List<string>"

EXCEL_EXTS = (".csv", ".xls", ".xlsx", ".xlsm")


def col_letter(n):
    """列索引（从 0 开始）转 Excel 列字母。0->A, 1->B, ..., 25->Z, 26->AA。"""
    result = ""
    n += 1
    while n > 0:
        n -= 1
        result = chr(ord('A') + n % 26) + result
        n //= 26
    return result


def cell_str(v):
    if v is None:
        return ""
    return str(v).strip()


def norm_type(raw):
    """归一化类型字符串：转小写并去掉所有空格。未知类型也返回一个值，由调用方判断合法性。"""
    return cell_str(raw).lower().replace(" ", "")


def is_array_type(t):
    """是否为「横向多列」的集合类型（数组或 List）。"""
    return t.endswith("[]") or t.startswith("list<")


def is_int_element(t):
    """集合元素是否为 int（决定单元格用整数解析还是字符串）。"""
    return t in ("int[]", "list<int>")


def csharp_type(t):
    return {
        "int": "int",
        "string": "string",
        "int[]": "int[]",
        "string[]": "string[]",
        "list<int>": "List<int>",
        "list<string>": "List<string>",
    }.get(t, "string")


def has_control_chars(s):
    """是否含 Excel/CSV 中通常不应出现的控制字符（保留制表/换行/回车）。"""
    for ch in s:
        if ord(ch) < 32 and ch not in "\t\n\r":
            return True
    return False


def strict_int(s):
    """严格解析整数：'5001'/'5001.0' -> 5001；否则 None。用于主键列。"""
    try:
        return int(s)
    except (ValueError, TypeError):
        pass
    try:
        f = float(s)
        if f.is_integer():
            return int(f)
    except (ValueError, TypeError):
        pass
    return None


def coerce_int_value(v):
    """宽容解析 int 列，返回 (整数值, 警告原因或 None)。空 -> (0, None)。"""
    if v is None:
        return 0, None
    if isinstance(v, bool):
        return (1 if v else 0), None
    if isinstance(v, int):
        return v, None
    if isinstance(v, float):
        if v != v:  # NaN
            return 0, "值为 NaN"
        if v.is_integer():
            return int(v), None
        return int(v), "值 '%s' 不是整数，已截断为 %d" % (v, int(v))
    s = cell_str(v)
    if s == "":
        return 0, None
    try:
        return int(s), None
    except ValueError:
        pass
    try:
        f = float(s)
        if f.is_integer():
            return int(f), None
        return int(f), "值 '%s' 不是整数，已截断为 %d" % (s, int(f))
    except ValueError:
        return 0, "值 '%s' 不是合法整数" % s


# ------------------------------- 日志系统 ----------------------------------
class Issue(object):
    """一条校验提示。level: ERROR/WARN；file_fatal=True 表示需跳过整个文件。"""

    __slots__ = ("level", "table", "sheet", "row", "col", "field", "reason", "file_fatal", "code")

    def __init__(self, level, table, reason, sheet=None, row=0, col="",
                 field="", file_fatal=False, code=""):
        self.level = level
        self.table = table
        self.reason = reason
        self.sheet = sheet
        self.row = row
        self.col = col
        self.field = field
        self.file_fatal = file_fatal
        self.code = code  # 归类码，用于同类折叠统计

    def format(self):
        loc = []
        if self.sheet:
            loc.append("Sheet:%s" % self.sheet)
        if self.row:
            loc.append("第%d行" % self.row)
        if self.col:
            loc.append("%s列" % self.col)
        if self.field:
            loc.append("字段'%s'" % self.field)
        locstr = (" " + " ".join(loc)) if loc else ""
        return "[%s] [%s]%s: %s" % (self.level, self.table, locstr, self.reason)


class Logger(object):
    """同时输出到控制台与日志文件；支持摘要统计。"""

    def __init__(self, log_path=None):
        self.path = log_path
        self._fh = None
        self.error_count = 0
        self.warn_count = 0
        if log_path:
            try:
                self._fh = open(log_path, "w", encoding="utf-8")
            except Exception as e:
                print("[WARN] 无法创建日志文件 %s: %s" % (log_path, e))
                self._fh = None

    def emit(self, text):
        print(text)
        if self._fh:
            try:
                self._fh.write(text + "\n")
            except Exception:
                pass

    def issue(self, iss):
        """输出一条校验提示并累计计数。"""
        if iss.level == LEVEL_ERROR:
            self.error_count += 1
        else:
            self.warn_count += 1
        self.emit(iss.format())

    def close(self):
        if self._fh:
            try:
                self._fh.flush()
                self._fh.close()
            except Exception:
                pass
            self._fh = None


# 全局日志对象（main 中初始化）
LOG = None


def log(text):
    if LOG is not None:
        LOG.emit(text)
    else:
        print(text)


def pause():
    """双击 exe/exe 运行结束后保持窗口；自动化(非 tty)环境不阻塞。"""
    try:
        if sys.stdin is not None and sys.stdin.isatty():
            input("\n按回车键关闭窗口...")
    except Exception:
        pass


# ----------------------------- 加载外部配置 -----------------------------
def load_config(base):
    """读取 base 目录下的 config.json（若存在）覆盖默认配置；不存在则写出默认配置。"""
    path = os.path.join(base, CONFIG_FILENAME)
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                user = json.load(f)
            for k in DEFAULT_CONFIG:
                if k in user:
                    cfg[k] = user[k]
        except Exception as e:
            log("[WARN] 读取 %s 失败，使用默认配置: %s" % (CONFIG_FILENAME, e))
    else:
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
            log("[INFO] 已生成配置文件 %s，可随时修改。" % CONFIG_FILENAME)
        except Exception:
            pass
    return cfg


def resolve_source_dir(base, src_dir):
    """解析 Excel 源目录。相对路径基于 base，绝对路径直接使用。"""
    if os.path.isabs(src_dir):
        return src_dir
    return os.path.normpath(os.path.join(base, src_dir))


# ---------------------------- 读取表格为二维网格 ----------------------------
def read_grid(path):
    """读取 Excel 族文件的第一张表，返回 (sheet_name, rows)。

    rows 为二维列表（每行是各单元格的值）。读取失败或格式不支持返回 (None, None)。
    """
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        return ws.title, rows
    if ext == ".xls":
        import xlrd
        book = xlrd.open_workbook(path)
        sh = book.sheet_by_index(0)
        return sh.name, [sh.row_values(i) for i in range(sh.nrows)]
    if ext == ".csv":
        import csv
        with open(path, "r", encoding="utf-8-sig", newline="") as f:
            return "CSV", [list(r) for r in csv.reader(f)]
    return None, None


def find_table_file(excel_dir, table):
    """在源目录中按「表名=文件名(不含扩展名)」查找文件路径。"""
    try:
        names = os.listdir(excel_dir)
    except Exception:
        return None
    for fname in names:
        if fname.startswith("~$"):
            continue
        base, ext = os.path.splitext(fname)
        if ext.lower() in EXCEL_EXTS and base == table:
            p = os.path.join(excel_dir, fname)
            if os.path.isfile(p):
                return p
    return None


# ------------------------------ 解析字段与数据 ------------------------------
def parse(grid, sheet, fname):
    """解析网格。

    返回 (fields, data_rows, issues, ref_values)：
      fields     : 字段定义列表
      data_rows  : 已剔除异常行的数据行列表
      issues     : Issue 列表（分级提示）
      ref_values : {字段名: [(Excel行号, 值的字符串形式), ...]}，供跨表引用校验使用
    """
    issues = []
    nrow = len(grid)
    ncol = max((len(r) for r in grid), default=0) if nrow else 0

    header = grid[0] if nrow >= 1 else []
    types = grid[1] if nrow >= 2 else []
    comments = grid[2] if nrow >= 3 else []

    # ---- 表头/类型行/注释行的整体性检查 ----
    if ncol == 0 or not any(cell_str(c) for c in header):
        issues.append(Issue(LEVEL_ERROR, fname, "表头行为空，无法解析任何字段。",
                            sheet=sheet, row=1, file_fatal=True, code="header_empty"))
        return [], [], issues, {}

    if nrow < 2:
        issues.append(Issue(LEVEL_ERROR, fname, "缺少类型行，无法确定字段类型。",
                            sheet=sheet, row=2, file_fatal=True, code="type_row_missing"))
        return [], [], issues, {}

    if nrow < 3:
        issues.append(Issue(LEVEL_WARN, fname, "缺少注释行，字段注释将回退为字段名。",
                            sheet=sheet, row=3, code="comment_row_missing"))

    # ---- 逐列构建字段 ----
    fields = []
    seen_names = {}
    c = 0
    while c < ncol:
        hstr = cell_str(header[c]) if c < len(header) else ""
        # 注释列 / 空列 -> 跳过（空列也可能是数组的续列，稍后由数组展开消费）
        if hstr == "" or hstr.startswith("#"):
            c += 1
            continue

        raw_type = types[c] if c < len(types) else None
        t = norm_type(raw_type)

        # 类型合法性检查（缺失 / 未知 均为文件级致命错误，避免生成错误 C#）
        if cell_str(raw_type) == "":
            issues.append(Issue(
                LEVEL_ERROR, fname,
                "类型行为空，缺少类型定义（支持: %s）。" % VALID_TYPES_DISPLAY,
                sheet=sheet, row=2, col=col_letter(c), field=hstr,
                file_fatal=True, code="type_missing"))
        elif t not in VALID_TYPES:
            issues.append(Issue(
                LEVEL_ERROR, fname,
                "类型 '%s' 不在支持列表中（支持: %s）。"
                % (cell_str(raw_type), VALID_TYPES_DISPLAY),
                sheet=sheet, row=2, col=col_letter(c), field=hstr,
                file_fatal=True, code="type_unknown"))

        # 字段名合法性（非法标识符 / 保留字 会导致生成的 C# 无法编译）
        if not IDENT_RE.match(hstr):
            issues.append(Issue(
                LEVEL_ERROR, fname,
                "字段名不是合法的 C# 标识符（只能由字母/数字/下划线组成，且不能以数字开头）。",
                sheet=sheet, row=1, col=col_letter(c), field=hstr,
                file_fatal=True, code="field_name_invalid"))
        elif hstr in CS_KEYWORDS:
            issues.append(Issue(
                LEVEL_ERROR, fname,
                "字段名是 C# 保留字，会导致生成的代码无法编译。",
                sheet=sheet, row=1, col=col_letter(c), field=hstr,
                file_fatal=True, code="field_name_keyword"))

        is_arr = is_array_type(t)
        if is_arr:
            end = c + 1
            while end < ncol:
                nh = cell_str(header[end] if end < len(header) else "")
                if nh != "":
                    break
                end += 1
            col_end = end
            # 数组至少占 2 列才有实际数据（1 列意味着续列被下一个表头截断了）
            if col_end <= c + 1:
                issues.append(Issue(
                    LEVEL_WARN, fname,
                    "类型为 '%s'，但仅占 1 列（紧接着 %s列的表头 '%s' 不为空），数组无法横向展开。"
                    "如确需多列数组，请将续列表头留空。"
                    % (t,
                       col_letter(c + 1) if (c + 1) < ncol else "?",
                       cell_str(header[c + 1] if (c + 1) < len(header) else "")),
                    sheet=sheet, row=1, col=col_letter(c), field=hstr, code="array_single_col"))
        else:
            col_end = c + 1

        # 重复字段名检查
        if hstr in seen_names:
            issues.append(Issue(
                LEVEL_ERROR, fname,
                "字段名与 %s列重复。" % col_letter(seen_names[hstr]),
                sheet=sheet, row=1, col=col_letter(c), field=hstr,
                file_fatal=True, code="field_dup"))
        else:
            seen_names[hstr] = c

        comment = cell_str(comments[c] if c < len(comments) else "")
        fields.append({
            "name": hstr,
            "ctype": t,
            "is_array": is_arr,
            "col_start": c,
            "col_end": col_end,
            "comment": comment,
        })
        c = col_end

    # 首字段为数组/List -> C# 字典键无效
    if fields and fields[0]["is_array"]:
        issues.append(Issue(
            LEVEL_ERROR, fname,
            "第1个字段是数组/List 类型，不能作为字典键；第一个字段必须是 int 或 string，请调整列顺序。",
            sheet=sheet, row=1, col=col_letter(fields[0]["col_start"]),
            field=fields[0]["name"], file_fatal=True, code="first_field_array"))

    # ---- 游离列 / 超出表头定义列 的检查 ----
    covered = set()
    for f in fields:
        covered.update(range(f["col_start"], f["col_end"]))
    for cc in range(ncol):
        if cc in covered:
            continue
        hv = cell_str(header[cc]) if cc < len(header) else ""
        if hv.startswith("#"):
            continue  # 注释列，允许有类型/数据
        tv = cell_str(types[cc]) if cc < len(types) else ""
        if tv != "":
            issues.append(Issue(
                LEVEL_WARN, fname,
                "类型行为 '%s' 但表头为空，该列将被忽略"
                "（数组续列应保持类型行为空；独立字段请补充表头名）。" % tv,
                sheet=sheet, row=2, col=col_letter(cc), code="orphan_type"))

    # ---- 数据行解析（从第 4 行开始，索引 3） ----
    data_rows = []
    ref_values = {f["name"]: [] for f in fields}
    seen_keys = {}
    empty_row_idx = []
    first_data_row = None
    last_data_row = None
    stray_reported = False

    if not fields:
        return [], [], issues, {}

    key_field = fields[0]
    key_col = key_field["col_start"]
    last_defined = max(f["col_end"] for f in fields)

    for ri in range(3, nrow):
        excel_row = ri + 1
        row = grid[ri]

        def get(cc, _row=row):
            return _row[cc] if cc < len(_row) else None

        # 整行全空 -> 记录（稍后判断是否为数据区中间的空行）
        if not any(cell_str(get(cc)) for cc in range(ncol)):
            empty_row_idx.append(excel_row)
            continue

        # 注释行：最左列首字符为 '#'
        if cell_str(get(0)).startswith("#"):
            continue

        # 关键修复：主键列（第 1 个字段）为空 -> ERROR 并跳过该行（不再静默丢弃）
        key_cell = cell_str(get(key_col))
        if key_cell == "":
            preview = "、".join(
                "%s列=%r" % (col_letter(cc), cell_str(get(cc)))
                for cc in range(ncol)
                if cc != key_col and cell_str(get(cc)) != ""
            )
            if len(preview) > 80:
                preview = preview[:80] + "…"
            tip = "（该行其余非空单元格: %s）" % preview if preview else "（该行仅有主键列为空）"
            issues.append(Issue(
                LEVEL_ERROR, fname,
                "主键列为空，已跳过该行%s。" % tip,
                sheet=sheet, row=excel_row, col=col_letter(key_col),
                field=key_field["name"], code="key_empty"))
            continue

        # 主键类型校验（int 主键必须是合法整数）
        if key_field["ctype"] == "int":
            key_int = strict_int(key_cell)
            if key_int is None:
                issues.append(Issue(
                    LEVEL_ERROR, fname,
                    "主键列内容 '%s' 不是合法整数，已跳过该行。" % key_cell,
                    sheet=sheet, row=excel_row, col=col_letter(key_col),
                    field=key_field["name"], code="key_not_int"))
                continue
            if _cfg_int_range() and (key_int > INT32_MAX or key_int < INT32_MIN):
                issues.append(Issue(
                    LEVEL_WARN, fname,
                    "主键值 %d 超出 C# Int32 范围 [%d, %d]。"
                    % (key_int, INT32_MIN, INT32_MAX),
                    sheet=sheet, row=excel_row, col=col_letter(key_col),
                    field=key_field["name"], code="int_range"))
            key_val = str(key_int)
        else:
            key_val = key_cell

        # 主键重复校验（C# 端 dict.Add 会抛异常）
        if key_val in seen_keys:
            issues.append(Issue(
                LEVEL_ERROR, fname,
                "主键 '%s' 与第%d行重复，已跳过该行。" % (key_cell, seen_keys[key_val]),
                sheet=sheet, row=excel_row, col=col_letter(key_col),
                field=key_field["name"], code="key_dup"))
            continue
        seen_keys[key_val] = excel_row

        if first_data_row is None:
            first_data_row = excel_row
        last_data_row = excel_row

        # ---- 逐字段取值 ----
        obj = {}
        for f in fields:
            name = f["name"]
            if f["is_array"]:
                vals = []
                for cc in range(f["col_start"], f["col_end"]):
                    v = get(cc)
                    s = cell_str(v)
                    if s == "":
                        continue
                    if is_int_element(f["ctype"]):
                        iv, warn = coerce_int_value(v)
                        if warn:
                            issues.append(Issue(
                                LEVEL_WARN, fname,
                                "%s，已按 %d 处理。" % (warn, iv),
                                sheet=sheet, row=excel_row, col=col_letter(cc),
                                field=name, code="int_invalid"))
                        elif _cfg_int_range() and not (INT32_MIN <= iv <= INT32_MAX):
                            issues.append(Issue(
                                LEVEL_WARN, fname,
                                "值 %d 超出 C# Int32 范围 [%d, %d]。" % (iv, INT32_MIN, INT32_MAX),
                                sheet=sheet, row=excel_row, col=col_letter(cc),
                                field=name, code="int_range"))
                        vals.append(iv)
                        ref_values[name].append((excel_row, str(iv)))
                    else:
                        vals.append(s)
                        ref_values[name].append((excel_row, s))
                obj[name] = vals
            elif f["ctype"] == "int":
                v = get(f["col_start"])
                iv, warn = coerce_int_value(v)
                if warn:
                    issues.append(Issue(
                        LEVEL_WARN, fname,
                        "%s，已按 %d 处理。" % (warn, iv),
                        sheet=sheet, row=excel_row, col=col_letter(f["col_start"]),
                        field=name, code="int_invalid"))
                elif _cfg_int_range() and not (INT32_MIN <= iv <= INT32_MAX):
                    issues.append(Issue(
                        LEVEL_WARN, fname,
                        "值 %d 超出 C# Int32 范围 [%d, %d]。" % (iv, INT32_MIN, INT32_MAX),
                        sheet=sheet, row=excel_row, col=col_letter(f["col_start"]),
                        field=name, code="int_range"))
                obj[name] = iv
                ref_values[name].append((excel_row, str(iv)))
            else:
                s = cell_str(get(f["col_start"]))
                if s and has_control_chars(s):
                    issues.append(Issue(
                        LEVEL_WARN, fname,
                        "字符串包含控制字符，请检查是否误粘贴了不可见字符。",
                        sheet=sheet, row=excel_row, col=col_letter(f["col_start"]),
                        field=name, code="ctrl_char"))
                obj[name] = s
                ref_values[name].append((excel_row, s))

        # 必填字段（不含主键，主键已单独处理）为空 -> WARN
        for req in _cfg_required_fields(fname):
            nm = req
            if nm in obj and _is_empty_value(obj[nm]):
                issues.append(Issue(
                    LEVEL_WARN, fname,
                    "必填字段为空（由 REQUIRED_FIELDS 配置）。",
                    sheet=sheet, row=excel_row, field=nm, code="required_empty"))

        # 超出表头定义范围的单元格（只报一次，避免刷屏）
        if not stray_reported and last_defined < ncol:
            extra = []
            for cc in range(last_defined, ncol):
                hv = cell_str(header[cc]) if cc < len(header) else ""
                if hv.startswith("#"):
                    continue
                if cell_str(get(cc)) != "":
                    extra.append(col_letter(cc))
            if extra:
                stray_reported = True
                issues.append(Issue(
                    LEVEL_WARN, fname,
                    "存在超出表头定义范围的单元格(%s列)，将被忽略。" % ",".join(extra),
                    sheet=sheet, row=excel_row, code="stray_col"))

        data_rows.append(obj)

    # 数据区中间的空行提示
    if data_rows and first_data_row and last_data_row and empty_row_idx:
        mid = [r for r in empty_row_idx if first_data_row < r < last_data_row]
        if mid:
            issues.append(Issue(
                LEVEL_WARN, fname,
                "数据区中间存在空行（第 %s 行），已跳过。"
                % ",".join(str(r) for r in mid),
                sheet=sheet, code="blank_row_mid"))

    return fields, data_rows, issues, ref_values


# 运行期使用的配置引用（避免层层传参）
_CFG = {}


def _cfg_int_range():
    return bool(_CFG.get("INT_RANGE_CHECK", True))


def _cfg_required_fields(fname):
    table = os.path.splitext(fname)[0]
    req = _CFG.get("REQUIRED_FIELDS") or {}
    return req.get(table, []) if isinstance(req, dict) else []


def _is_empty_value(v):
    if isinstance(v, list):
        return len(v) == 0
    return cell_str(v) == ""


# ------------------------------- 生成 JSON ---------------------------------
def gen_json(data_rows):
    return json.dumps({"dataList": data_rows}, ensure_ascii=False, indent=2)


# ------------------------------- 生成 C# -----------------------------------
def gen_cs(basename, fields, cfg):
    if not fields:
        return ""
    key = fields[0]
    key_type = csharp_type(key["ctype"])
    rowclass = basename
    catclass = basename + "Category"

    L = []
    L.append("using System.Collections;")
    L.append("using System.Collections.Generic;")
    L.append("using System.ComponentModel;")
    if cfg["USE_NEWTONSOFT"]:
        L.append("using Newtonsoft.Json;")
    if cfg["USE_UNITY"]:
        L.append("using UnityEngine;")
    L.append("")
    L.append("")
    L.append("/*")
    L.append(" *重点！！！！！！！！！！！！！！！！！！！！！！！")
    L.append(" *")
    L.append(" *该代码属于导表工具自动生成")
    L.append(" *切勿修改该代码 否则下次导表会被覆盖")
    L.append(" */")
    L.append("")
    L.append("")
    L.append("[System.Serializable]")
    if cfg["USE_CONFIG_ATTR"]:
        L.append("[Config]")
    L.append("public partial class %s:IConfigCategory" % catclass)
    L.append("{")
    L.append("\t[JsonIgnore]")
    L.append("\tpublic static %s Inst;" % catclass)
    L.append("")
    L.append("")
    L.append("\tprivate Dictionary<%s, %s> dict = new Dictionary<%s, %s>();"
             % (key_type, rowclass, key_type, rowclass))
    if cfg["USE_UNITY"]:
        L.append("\t[SerializeField]")
    L.append("\tpublic List<%s> dataList= new List<%s>();" % (rowclass, rowclass))
    L.append("")
    L.append("\tpublic %s()" % catclass)
    L.append("\t{")
    L.append("\t\tInst = this;")
    L.append("\t}")
    L.append("")
    L.append("")
    L.append("\tpublic override void EndInit()")
    L.append("\t{")
    L.append("\t\tforeach (%s config in dataList)" % rowclass)
    L.append("\t\t{")
    L.append("\t\t\tconfig.EndInit();")
    L.append("\t\t\tthis.dict.Add(config.%s,config);" % key["name"])
    L.append("\t\t}")
    L.append("\t\tthis.AfterEndInit(); ")
    L.append("\t}")
    L.append("\tpublic %s GetOne(%s id)" % (rowclass, key_type))
    L.append("\t{")
    L.append("\t\tif (dict.TryGetValue(id,out %s c))" % rowclass)
    L.append("\t\t{")
    L.append("\t\t\treturn c;")
    L.append("\t\t}")
    L.append("")
    L.append("\t\treturn null;")
    L.append("\t}")
    L.append("\tpublic bool isContain(%s id)" % key_type)
    L.append("\t{")
    L.append("\t\treturn dict.ContainsKey(id);")
    L.append("\t}")
    L.append("\tpublic Dictionary<%s, %s> GetDict()" % (key_type, rowclass))
    L.append("\t{")
    L.append("\t\treturn dict;")
    L.append("\t}")
    L.append("}")
    L.append("")
    L.append("")
    L.append("")
    L.append("[System.Serializable]")
    L.append("public partial class %s:IConfig" % rowclass)
    L.append("{")
    for f in fields:
        ct = csharp_type(f["ctype"])
        comment = f["comment"] if f["comment"] else f["name"]
        L.append("\tpublic %s %s; // %s" % (ct, f["name"], comment))
    L.append("}")
    L.append("")
    L.append("")
    return "\n".join(L)


# ----------------------------- 增量缓存 -----------------------------
def compute_file_hash(filepath):
    """返回 (size, mtime, md5_hex) 三元组。"""
    stat = os.stat(filepath)
    size = stat.st_size
    mtime = int(stat.st_mtime)
    md5 = hashlib.md5()
    with open(filepath, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            md5.update(chunk)
    return (size, mtime, md5.hexdigest())


def load_cache(base):
    """读取缓存文件，返回 dict: {filename: {size, mtime, md5}}。"""
    path = os.path.join(base, CACHE_FILENAME)
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_cache(base, cache):
    """写入缓存文件。"""
    path = os.path.join(base, CACHE_FILENAME)
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log("[WARN] 写入缓存失败: %s" % e)


# ------------------------------- 处理单个文件 -------------------------------
def process_file(path, fname, base, cfg):
    """处理单个 Excel 文件，导出 JSON 和 C#。

    返回 dict: {exported: bool, fatal: bool, error_count: int, warn_count: int,
               parsed: {fields, data_rows, ref_values} or None}
    """
    result = {"exported": False, "fatal": False, "error_count": 0,
              "warn_count": 0, "parsed": None}

    try:
        sheet, grid = read_grid(path)
    except Exception as e:
        log("[ERROR] [%s] 读取失败: %s" % (fname, e))
        result["error_count"] = 1
        result["fatal"] = True
        return result

    if not grid:
        log("[ERROR] [%s] 无法读取或格式不支持，已跳过。" % fname)
        result["error_count"] = 1
        result["fatal"] = True
        return result

    if len(grid) < 4:
        log("[WARN] [%s] 行数不足（需要 表头/类型/注释/至少一行数据，当前仅 %d 行），已跳过。"
            % (fname, len(grid)))
        result["warn_count"] = 1
        return result

    fields, data_rows, issues, ref_values = parse(grid, sheet, fname)

    # 输出提示（做同类折叠，避免单文件刷屏）
    cap = int(cfg.get("MAX_ISSUES_PER_FILE", 300) or 300)
    _emit_issues(fname, issues, cap)

    result["error_count"] = sum(1 for i in issues if i.level == LEVEL_ERROR)
    result["warn_count"] = sum(1 for i in issues if i.level == LEVEL_WARN)

    fatal = any(i.file_fatal for i in issues)
    if fatal or not fields:
        result["fatal"] = True
        if not fields:
            log("[SKIP] [%s] 未解析到任何有效字段，未导出。" % fname)
        else:
            log("[SKIP] [%s] 存在致命配置错误，未导出（避免生成错误数据/代码）。" % fname)
        return result

    basename = os.path.splitext(fname)[0]
    json_dir = os.path.join(base, cfg["JSON_OUTPUT_DIR"])
    cs_dir = os.path.join(base, cfg["CS_OUTPUT_DIR"])
    os.makedirs(json_dir, exist_ok=True)
    os.makedirs(cs_dir, exist_ok=True)

    json_path = os.path.join(json_dir, basename + "Category.json")
    cs_path = os.path.join(cs_dir, basename + ".cs")

    with open(json_path, "w", encoding="utf-8") as f:
        f.write(gen_json(data_rows))
    with open(cs_path, "w", encoding="utf-8") as f:
        f.write(gen_cs(basename, fields, cfg))

    skipped_rows = len([i for i in issues if i.code in
                        ("key_empty", "key_not_int", "key_dup")])
    extra = ("（跳过 %d 行异常数据）" % skipped_rows) if skipped_rows else ""
    log("[OK]   [%s] -> %s + %s  ; 数据行 %d 条%s"
        % (fname, os.path.relpath(json_path, base), os.path.relpath(cs_path, base),
           len(data_rows), extra))

    result["exported"] = True
    result["parsed"] = {"fields": fields, "data_rows": data_rows, "ref_values": ref_values}
    return result


def _emit_issues(fname, issues, cap):
    """输出提示列表；同类(code)超过阈值时折叠为汇总行。"""
    if not issues:
        return
    per_code = {}
    shown = 0
    folded = {}
    for iss in issues:
        code = iss.code or "misc"
        per_code[code] = per_code.get(code, 0) + 1
        if shown < cap:
            LOG.issue(iss)
            shown += 1
        else:
            folded[code] = folded.get(code, 0) + 1
    if folded:
        detail = "、".join("%s×%d" % (k, v) for k, v in folded.items())
        log("[WARN] [%s] 该文件提示超过 %d 条，其余已折叠：%s"
            % (fname, cap, detail))
    log("")


# ---------------------------- 跨表引用校验 ----------------------------
def _load_ref_target_values(excel_dir, table, field, cache):
    """加载目标表的「引用取值集合」，返回 (status, values)。
    status: 'ok' / 'missing_file' / 'nofield' / 'error'
    """
    key = (table, field)
    if key in cache:
        return cache[key]

    path = find_table_file(excel_dir, table)
    if not path:
        cache[key] = ("missing_file", None)
        return cache[key]
    try:
        sheet, grid = read_grid(path)
        if not grid:
            cache[key] = ("error", None)
            return cache[key]
        fields, _rows, _issues, ref_values = parse(grid, sheet, os.path.basename(path))
    except Exception:
        cache[key] = ("error", None)
        return cache[key]

    if not fields:
        cache[key] = ("nofield", None)
        return cache[key]

    target_field = field if field else fields[0]["name"]
    if target_field not in ref_values:
        cache[key] = ("nofield", None)
        return cache[key]

    vals = set(v for (_row, v) in ref_values.get(target_field, []) if v != "")
    cache[key] = ("ok", vals)
    return cache[key]


def validate_refs(excel_dir, cfg, parsed_tables):
    """跨表引用校验。parsed_tables: {表名: {"ref_values": {...}}}（本次实际解析的表）。"""
    rules = cfg.get("REF_RULES") or {}
    if not isinstance(rules, dict) or not rules:
        return
    log("")
    log("---- 跨表引用校验 ----")
    cache = {}
    for src, tgt in rules.items():
        if "." not in str(src) or not str(tgt):
            log("[WARN] [REF_RULES] 规则 '%s' -> '%s' 格式不正确（源应为 表名.字段名），已忽略。"
                % (src, tgt))
            continue
        table, field = str(src).split(".", 1)
        tgt = str(tgt)
        ttable, tfield = (tgt.split(".", 1) + [None])[:2]

        if table not in parsed_tables:
            log("[WARN] [REF_RULES] 源表 '%s' 本次未导出（可能内容未变化或未找到），跳过其引用校验。"
                % table)
            continue

        status, vals = _load_ref_target_values(excel_dir, ttable, tfield, cache)
        if status == "missing_file":
            log("[ERROR] [REF_RULES] 源表 '%s' 字段 '%s' 引用的目标表 '%s' 未找到对应 Excel 文件。"
                % (table, field, ttable))
            continue
        if status == "nofield":
            log("[ERROR] [REF_RULES] 目标表 '%s' 中不存在字段 '%s'。"
                % (ttable, tfield if tfield else "(主键列)"))
            continue
        if status == "error":
            log("[ERROR] [REF_RULES] 读取目标表 '%s' 失败，跳过其引用校验。" % ttable)
            continue

        refv = parsed_tables[table]["ref_values"]
        if field not in refv:
            log("[WARN] [REF_RULES] 源表 '%s' 中不存在字段 '%s'。" % (table, field))
            continue

        hit = 0
        for (row, val) in refv.get(field, []):
            if val == "":
                continue
            if val not in vals:
                hit += 1
                src_iss = Issue(
                    LEVEL_ERROR, table,
                    "引用的 Id '%s' 在表 '%s' 中不存在。" % (val, ttable),
                    row=row, field=field, code="ref_missing")
                LOG.issue(src_iss)
        if hit == 0:
            log("[OK]   [REF_RULES] %s.%s -> %s 引用校验通过。" % (table, field, ttable))


# --------------------------------- 主流程 -----------------------------------
def main():
    global LOG, _CFG
    # 兼容 .py 直接运行 与 PyInstaller 打包后的 exe
    if getattr(sys, "frozen", False):
        tool_dir = os.path.dirname(os.path.abspath(sys.executable))
    else:
        tool_dir = os.path.dirname(os.path.abspath(__file__))

    # 控制台中文输出（Windows 下避免 gbk 编码报错）
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

    cfg = load_config(tool_dir)
    _CFG = cfg

    # 初始化日志（控制台 + 文件）
    log_name = cfg.get("LOG_FILE", "export_log.txt")
    log_path = os.path.join(tool_dir, log_name) if log_name else None
    LOG = Logger(log_path)

    # 解析 Excel 源目录
    excel_dir = resolve_source_dir(tool_dir, cfg.get("EXCEL_SOURCE_DIR", "."))
    if not os.path.isdir(excel_dir):
        log("[ERROR] Excel 源目录不存在: %s" % excel_dir)
        log("        请检查 config.json 中的 EXCEL_SOURCE_DIR 配置。")
        LOG.close()
        pause()
        return

    log("Excel 源目录: %s" % excel_dir)
    log("JSON 输出目录: %s" % os.path.join(tool_dir, cfg["JSON_OUTPUT_DIR"]))
    log("C#   输出目录: %s" % os.path.join(tool_dir, cfg["CS_OUTPUT_DIR"]))
    if log_path:
        log("日志文件: %s" % log_path)
    log("")

    # 加载增量缓存
    cache = load_cache(tool_dir)

    processed = 0
    skipped = 0
    failed = 0
    parsed_tables = {}  # 表名 -> 解析结果（供引用校验）

    for fname in sorted(os.listdir(excel_dir)):
        if fname.startswith("~$"):
            continue
        ext = os.path.splitext(fname)[1].lower()
        if ext not in EXCEL_EXTS:
            continue
        fpath = os.path.join(excel_dir, fname)
        if os.path.isdir(fpath):
            continue

        # --- 增量检查 ---
        try:
            fsize, fmtime, fmd5 = compute_file_hash(fpath)
        except Exception as e:
            log("[ERROR] [%s] 无法读取文件: %s" % (fname, e))
            failed += 1
            continue

        cached = cache.get(fname)
        if cached and cached.get("size") == fsize and cached.get("mtime") == fmtime:
            log("[SKIP] [%s] 内容未变化，跳过导出。" % fname)
            skipped += 1
            continue
        if cached and cached.get("md5") == fmd5:
            cache[fname] = {"size": fsize, "mtime": fmtime, "md5": fmd5}
            log("[SKIP] [%s] 内容未变化（时间戳不同），跳过导出。" % fname)
            skipped += 1
            continue

        # --- 导出 ---
        try:
            r = process_file(fpath, fname, tool_dir, cfg)
        except Exception as e:
            log("[ERROR] [%s] 导出异常: %s" % (fname, e))
            failed += 1
            continue

        if r["exported"]:
            cache[fname] = {"size": fsize, "mtime": fmtime, "md5": fmd5}
            processed += 1
            table = os.path.splitext(fname)[0]
            parsed_tables[table] = r["parsed"]
        elif r["fatal"]:
            # 致命错误：不更新缓存，下次运行会重新校验
            failed += 1
        else:
            # 仅行数不足等非致命跳过：视为跳过
            skipped += 1

    # 跨表引用校验
    try:
        validate_refs(excel_dir, cfg, parsed_tables)
    except Exception as e:
        log("[WARN] 引用校验过程出错: %s" % e)

    # 保存缓存
    save_cache(tool_dir, cache)

    log("")
    log("=" * 56)
    log("处理完成: 导出 %d 个 | 跳过 %d 个 | 失败 %d 个 | 警告 %d 条 | 错误 %d 条"
        % (processed, skipped, failed, LOG.warn_count, LOG.error_count))
    if log_path:
        log("详细日志: %s" % log_path)
    log("=" * 56)

    LOG.close()
    pause()


if __name__ == "__main__":
    main()
