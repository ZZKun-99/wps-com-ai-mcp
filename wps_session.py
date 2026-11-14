# -*- coding: utf-8 -*-
"""wps_session.py —— 活动文档薄壳 + 分层加载（V3）

写给下一个接手的自己
====================
本文件是 **新增层**，不改动 server.py 的老地基（COM 挂接 / 僵尸实例识别 / 编辑态救援）。
server.py 只加三处接线：tools/list 只广播薄工具、handle_tool_call 加 wps_do 分支、
缺的东西从本模块取。老 24 个工具的定义与分发**一个字都没动**，仍可被调用。

要做的事
--------
1. **薄壳（Session 层）**：AI 以后只说「活动表 B5 = xxx」「第 3 行 = [a,b,c]」，
   不用传 file_path、不用传 mode="com"。
2. **分层加载**：默认只向 AI 暴露 2 个极简工具（wps_docs / wps_do），
   Excel 的详细参数说明**按需**用 `wps_do(app="excel", action="help")` 取。
   → 常驻 schema 从 ~10.4k 字符降到 ~0.75k 字符，具体见文末「分层加载效果」。

V3 = V2 的全部内容 + 本层（2026-09-28 合并）
--------------------------------------------
V2 那条修改路径的内容已**全部并入 V3 的 server.py**（相对 V2 只多 84 行、只换 3 行）：
`activate_sheet` / `read_row` 两个工具、`_file_read_row`、以及为 comtypes 兜底而加的
`_com_to_plain` / `_com_attr` / `_com_set_value` / `_safe_sheet_name`（`edit_cell`、
`_com_write_xlsx` 的写值处也换成了 `_com_set_value`）。所以老工具里现在有 26 个。
- 薄壳侧对应动作：`wps_do(app="excel", action="activate", args={"sheet":"Sheet2"})`
  与 `get ref="3:3"`（read_row 的等价物，返回体更瘦）。
- `activate_sheet` 已登记进 `_LEGACY_WRITE`：legacy 逃生口只放读类，写/改界面的走薄壳动作。

调用约定（给人看的提示词片段，可直接粘进人设/场景）
--------------------------------------------------
    「需要操作 WPS 时调用 wps_do。默认操作我在 WPS 里已经打开、尚未保存的那份文档（内存实时值，
      不是磁盘上的旧文件）。参数不确定时先 wps_do(app="excel", action="help") 取说明再调用。
      除非我明确要求，不要自动保存——我要保留 Ctrl+Z。」

三条硬约束（与交接文档一致，别违反）
------------------------------------
- 只操作「人正在打开的活文档」，绝不读磁盘、绝不 `Dispatch` 新建实例、绝不 `Quit`/`Close` 人的文档。
- 写操作默认**不 Save**（`WPS_MCP_AUTOSAVE=0`），保人的 Ctrl+Z。
- 批量写必须一次 `Range.Value = 二维数组`，禁止 Python 循环逐格写（COM 跨进程极慢）。

与交接文档的两处**有意偏离**（都是实测/架构原因，别改回去）
------------------------------------------------------------
1. **不用 `from server import ...`**。server.py 是以 `python server.py` 方式跑的，
   `sys.modules["__main__"]` 才是它；在子模块里 `import server` 会**再执行一遍** server.py，
   产生第二份模块级状态（`_APP_CACHE` 等各自一份）。改为**依赖注入**：
   `server.py` 调用时把自己作为 `host` 传进来，本模块只通过 `host.xxx` 访问内部函数。
   副作用是好事——测试可以传一个纯替身 host，完全不碰真 WPS。
2. **Session 只缓存「数据」，不缓存 COM 对象**。老地基的 `_APP_CACHE` 缓存 app 句柄并带探活，
   那没问题；但把 `wb`/`ws` 跨请求长期攥在手里会有两个后果：
   ① 人关掉工作簿/关掉 WPS 后变成悬空引用，调用抛裸 COM 错误；
   ② pywin32 持引用会**拖住 et.exe 不退出**，而残留 et.exe 占 ROT 正是避坑文档 B4「挂到空进程」的成因。
   所以这里每次调用都重新走 `_get_wps_app()`（实测热挂接 15–16ms，几乎免费），
   只把「已用区域二维数组」这种纯 Python 数据留在缓存里。

快照默认关闭
------------
`WPS_MCP_SNAPSHOT_TTL_MS` 默认 **0**（不快照）。原因：AI 两次工具调用之间隔着模型生成时间（常 ≥1s），
TTL 2~3s 基本必然 miss，而 miss 的代价是**整表 UsedRange**，比直接读一格贵得多。
需要连读多段时才把 TTL 打开（并受 `WPS_MCP_SNAPSHOT_MAX_CELLS` 体积闸限制）。

环境变量
--------
- `WPS_MCP_SNAPSHOT_TTL_MS`    快照存活毫秒，0=不缓存（默认）
- `WPS_MCP_SNAPSHOT_MAX_CELLS` 快照体积闸，超过就不缓存（默认 50000）
- `WPS_MCP_AUTOSAVE`           1=写完自动保存（默认 0，保 Ctrl+Z）
- `WPS_MCP_MERGE_SCAN_MAX`     合并区逐格扫描上限（默认 2000 格）
- `WPS_MCP_READ_MAX_CELLS`     单次 get 的单元格上限（默认 200000）
- `WPS_MCP_LEGACY_TOOLS`       1=tools/list 退回全量老工具（默认 0，即分层加载）
- `WPS_MCP_THIN_WRITE_WAIT`    无路径写操作等人类松手的秒数（默认沿用 BUSY_RETRY_SECONDS）
"""

import os
import re
import time

# ==================== 配置 ====================

SNAPSHOT_TTL = float(os.environ.get("WPS_MCP_SNAPSHOT_TTL_MS", "0") or 0) / 1000.0
SNAPSHOT_MAX_CELLS = int(os.environ.get("WPS_MCP_SNAPSHOT_MAX_CELLS", "50000") or 50000)
AUTOSAVE = os.environ.get("WPS_MCP_AUTOSAVE", "0") not in ("0", "false", "False", "")
MERGE_SCAN_MAX = int(os.environ.get("WPS_MCP_MERGE_SCAN_MAX", "2000") or 2000)
READ_MAX_CELLS = int(os.environ.get("WPS_MCP_READ_MAX_CELLS", "200000") or 200000)

APP_KEYS = ("excel", "word", "ppt")
_APP_TYPE = {"excel": "excel", "word": "word", "ppt": "powerpoint"}
EXCEL_WRITE_ACTIONS = {"set", "save", "activate", "width", "height", "fill", "style",
                       "merge", "unmerge", "brush", "insert", "delete", "clear", "copy",
                       "move", "replace"}   # find 是只读，故意不放进来
WORD_WRITE_ACTIONS = {"set", "save"}
HOST_API = (
    "_get_wps_app", "_release_app", "_get_open_document", "_com_call", "_com_try_attr",
    "_com_to_plain",
    "_safe_com_str", "_safe_com_bool", "_safe_sheet_name", "_caption_doc_name", "_open_count",
    "_bypass_shim", "_esc_exit_edit_mode", "_match_hint", "_com_color",
    "_underline_pair", "WpsAttachError",
    "BUSY_RETRY_SECONDS", "BUSY_RETRY_STEP", "BUSY_ESC_ENABLED",
)


def normalize_app(app):
    """excel/word/ppt 归一化（顺手收下 powerpoint / pptx / slide 这些说法）。"""
    key = (app or "").strip().lower()
    if key in ("powerpoint", "pptx", "ppt", "slide", "slides"):
        return "ppt"
    return key


# 「动作名后缀模式位」：0 = 保守（字面 / 不带格式），1 = 增强（正则 / 连格式）。
# 用户要求既要能按参数名区分、也要能按 0/1 后缀区分，所以两种都给：
#     copy0 = copy + with_format=false      copy1 = copy + with_format=true
#     find0 = find + regex=false            find1 = find + regex=true
#     replace0 = replace + regex=false      replace1 = replace + regex=true
# 规则：显式传了参数就**以显式参数为准**（后缀只是省字的写法）。
MODE_SUFFIX = {"copy": "with_format", "find": "regex", "replace": "regex"}


def split_mode_suffix(action):
    """'find1' → ('find', {'regex': True})；'copy0' → ('copy', {'with_format': False})。

    不是这种写法就原样返回 (action, None)。后缀只在基名是已登记的动作时才认，
    不会误伤别的动作名。
    """
    act = (action or "").strip().lower()
    if len(act) > 1 and act[-1] in "01":
        base = act[:-1]
        key = MODE_SUFFIX.get(base)
        if key:
            return base, {key: act[-1] == "1"}
    return act, None


def is_write(app, action):
    """本次调用是不是写操作——决定走「等松手 + 必要时 ESC」还是「只读旁路」。"""
    key = normalize_app(app)
    act, _flag = split_mode_suffix(action)
    if key == "excel":
        return act in EXCEL_WRITE_ACTIONS
    if key == "word":
        return act in WORD_WRITE_ACTIONS
    return False


# ==================== A1 ↔ 行列换算 ====================

_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d{1,7})$")
_RANGE_RE = re.compile(
    r"^\$?([A-Za-z]{1,3})?\$?(\d{1,7})?\s*:\s*\$?([A-Za-z]{1,3})?\$?(\d{1,7})?$")


def col_to_index(col):
    """列标识 → 1-based 列号：'G'/'g'/'AA'/7/'7'；None 原样返回。"""
    if col is None:
        return None
    if isinstance(col, bool):
        return None
    if isinstance(col, int):
        return int(col)
    text = str(col).strip().upper()
    if not text:
        return None
    if text.isdigit():
        return int(text)
    idx = 0
    for ch in text:
        if not ("A" <= ch <= "Z"):
            raise ValueError("非法列标识：%r" % (col,))
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def col_letter(col_index):
    """1-based 列号 → 列字母（7 → 'G'）。"""
    n = int(col_index)
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(ord("A") + rem) + out
    return out


def parse_ref(ref):
    """解析 A1 式引用。返回 dict(kind, r1, c1, r2, c2, sheet)。

    kind ∈ {"cell", "row", "col", "range"}；行/列可缺省（None）——
    `C:C` 只给列不给行、`3:3` 只给行不给列，缺的那一侧由调用方按「已用区域」补齐：
    读一列 = 该列 × 已用区域的行区间，读一行 = 该行 × 已用区域的列区间。

    支持：B5、$B$5、A2:D10、C:C、C:E、3:3、3:7、Sheet1!B5、Sheet1!A2:D10
    """
    if ref is None:
        raise ValueError("缺少 ref（单元格/行/列/区域地址，如 B5、3:3、C:C、A2:D10）")
    text = str(ref).strip()
    if not text:
        raise ValueError("ref 不能为空")
    sheet = None
    if "!" in text:
        sheet, text = text.split("!", 1)
        sheet = sheet.strip().strip("'\"") or None
        text = text.strip()
    text = text.replace(" ", "")

    m = _CELL_RE.match(text)
    if m:
        c = col_to_index(m.group(1))
        r = int(m.group(2))
        return {"kind": "cell", "r1": r, "c1": c, "r2": r, "c2": c, "sheet": sheet}

    m = _RANGE_RE.match(text)
    if m:
        l1, n1, l2, n2 = m.group(1), m.group(2), m.group(3), m.group(4)
        if not l1 and not l2 and not n1 and not n2:
            raise ValueError("无法解析 ref：%r" % (ref,))
        if l1 and l2 and n1 and n2:
            r1, c1, r2, c2 = int(n1), col_to_index(l1), int(n2), col_to_index(l2)
            kind = "range"
        elif not n1 and not n2 and (l1 or l2):
            c1 = col_to_index(l1)
            c2 = col_to_index(l2) if l2 else c1
            r1 = r2 = None
            kind = "col"
        elif not l1 and not l2 and (n1 or n2):
            r1 = int(n1)
            r2 = int(n2) if n2 else r1
            c1 = c2 = None
            kind = "row"
        else:
            raise ValueError(
                "ref 格式含糊：%r。列区间要写成 C:C / C:E，行区间要写成 3:3 / 3:7" % (ref,))
        if c1 and c2 and c2 < c1:
            c1, c2 = c2, c1
        if r1 and r2 and r2 < r1:
            r1, r2 = r2, r1
        return {"kind": kind, "r1": r1, "c1": c1, "r2": r2, "c2": c2, "sheet": sheet}

    raise ValueError(
        "无法解析 ref：%r。支持 B5 / $B$5 / A2:D10 / C:C / 3:3 / Sheet1!B5" % (ref,))


def ref_text(r1, c1, r2, c2):
    """(r1,c1,r2,c2) → 'A1' 或 'A2:D10'。"""
    if r1 == r2 and c1 == c2:
        return "%s%d" % (col_letter(c1), r1)
    return "%s%d:%s%d" % (col_letter(c1), r1, col_letter(c2), r2)


def flat_len(values):
    """值序列的「个数」：一维给长度，二维给首行宽度，标量给 None。"""
    if isinstance(values, (list, tuple)):
        if values and isinstance(values[0], (list, tuple)):
            return len(values[0]) or None
        return len(values) or None
    return None


def resolve_write_box(ref, box, values, origin_c=None, origin_r=None):
    """把写操作的 ref 收敛成确定的 (r1, c1, r2, c2)。

    整行写（"3:3"）缺列、整列写（"C:C"）缺行，缺的那一侧按
    「数据区原点 + 值的个数」补齐 —— 这样「第 3 行 = [a,b,c]」就落到数据区首列起的 3 格，
    和 get 的补齐口径一致（读整行也是按数据区的列区间）。
    """
    origin_r = origin_r or (box or {}).get("first_row", 1)
    origin_c = origin_c or (box or {}).get("first_col", 1)
    if ref["r1"] is not None and ref["c1"] is not None:
        return ref["r1"], ref["c1"], ref["r2"], ref["c2"]
    n = flat_len(values)
    if ref["r1"] is not None:
        r = ref["r1"]
        w = n or 1
        return r, origin_c, r, origin_c + w - 1
    if ref["c1"] is not None:
        c = ref["c1"]
        h = n or 1
        return origin_r, c, origin_r + h - 1, c
    raise ValueError("无法确定写入位置：ref=%r" % (ref,))


# ==================== 值 / 二维数组规范化 ====================


def coerce_grid(raw, rows, cols):
    """把 COM 读回来的东西统一成 rows×cols 的二维 list（缺的补 None）。

    COM 的脾气：1×1 返回标量；1×N 有的绑定返回 (a,b,c)，有的返回 ((a,b,c),)；
    N×1 有的返回 ((a,),(b,))。这里全部拍平对齐。
    """
    if rows <= 0 or cols <= 0:
        return []
    if raw is None:
        grid = [[None] * cols for _ in range(rows)]
    elif isinstance(raw, (list, tuple)):
        seq = list(raw)
        if seq and isinstance(seq[0], (list, tuple)) and rows > 1:
            grid = [list(row)[:cols] for row in seq]
        elif seq and isinstance(seq[0], (list, tuple)) and rows == 1:
            flat = list(seq[0])
            grid = [flat[:cols]]
        elif rows == 1 and cols > 1:
            grid = [seq[:cols]]
        elif cols == 1 and rows > 1:
            grid = [[v] for v in seq[:rows]]
        elif rows == 1 and cols == 1:
            grid = [[seq[0] if len(seq) == 1 else seq]]
        else:
            grid = [[v] for v in seq[:rows]]
    else:
        grid = [[raw]]
    for row in grid:
        while len(row) < cols:
            row.append(None)
    while len(grid) < rows:
        grid.append([None] * cols)
    return grid


def shape_values(grid):
    """二维 grid → 按形状瘦身：1×1 给标量，单行/单列给一维，其余给二维。

    返回体只回值本身——不重复输出列名/非空字典（老 read_row 把同一批值抄了三遍，
    26 列一行 ≈1.7k 字符，三分之二是重复）。
    """
    if not grid:
        return None
    if len(grid) == 1 and len(grid[0]) == 1:
        return grid[0][0]
    if len(grid) == 1:
        return list(grid[0])
    if len(grid[0]) == 1:
        return [row[0] for row in grid]
    return [list(row) for row in grid]


def to_matrix(values, rows, cols):
    """AI 传进来的 values → rows×cols 的 tuple-of-tuple（直接喂 COM）。

    规则（错了就抛带修法提示的错，不要猜）：
      1×1        → 标量或 [[v]] 都行
      1×N / N×1  → 一维数组，或 [[..]] 二维（按形状匹配）
      N×M        → 二维数组，行列数必须与 ref 完全一致
    """
    if rows == 1 and cols == 1:
        if isinstance(values, (list, tuple)):
            # 先判"单元素且该元素还是 list/tuple"（二维 [[v]]），再判单元素（一维 [v]）；
            # 顺序反了的话 [[123]] 会被收成 [123]，内层 list 原样喂给 COM 会类型错乱。
            if len(values) == 1 and isinstance(values[0], (list, tuple)):
                values = values[0][0]
            elif len(values) == 1:
                values = values[0]
            else:
                raise ValueError(
                    "目标只有 1 格（%d 格），但 values 给了 %d 个值。要么把 ref 写成区域，"
                    "要么只传一个值。" % (1, len(values)))
        return ((values,),)

    if not isinstance(values, (list, tuple)):
        raise ValueError("目标是 %d 行 %d 列区域，values 需要是二维数组，收到标量 %r" % (rows, cols, values))

    items = list(values)
    if items and not isinstance(items[0], (list, tuple)):
        if rows == 1 and len(items) == cols:
            return (tuple(items),)
        if cols == 1 and len(items) == rows:
            return tuple((v,) for v in items)
        raise ValueError(
            "values 是一维 %d 个值，与目标 %d 行 %d 列不匹配。单行请把 ref 写成 3:3 或 A3:H3，"
            "单列请写成 C:C。" % (len(items), rows, cols))

    matrix = [list(row) for row in items]
    if len(matrix) != rows or any(len(row) != cols for row in matrix):
        got = "%d 行 × %s" % (len(matrix), "/".join(str(len(r)) for r in matrix[:3]) or "?")
        raise ValueError(
            "values 形状 %s 与目标 %d 行 %d 列不一致。整片写入要求行列数严格相同"
            "（不够的格子请显式补 null）。" % (got, rows, cols))
    return tuple(tuple(row) for row in matrix)


# ==================== COM 写入（comtypes 兜底） ====================


def set_range_value(target, value):
    """给 Range 的默认属性（Value，DISPID=0）赋值。

    pywin32 下 `target.Value = value` 直接可用；comtypes lazybind 在本机 WPS 类型库上
    会对 Range.Value 抛 AttributeError，需退化为 DISPID 0 + PROPERTYPUT 调用。
    本文件自带这一份，避免依赖 v2 分支的改动（v2 是另一条修改路径）。
    """
    try:
        target.Value = value
        return
    except Exception as first_err:
        direct_err = first_err
    inner = getattr(target, "_comobj", None)
    if inner is None or not hasattr(inner, "_invoke"):
        disp = getattr(target, "disp", None)
        inner = getattr(disp, "_comobj", None)
    if inner is None or not hasattr(inner, "_invoke"):
        raise AttributeError("无法定位 COM 底层对象以写入 Value：%s" % (direct_err,))
    last_err = None
    for invkind in (8, 4):  # PROPERTYPUT / PROPERTYPUTREF：本机 WPS 只吃后者
        try:
            inner._invoke(0, invkind, 0, value)
            return
        except Exception as exc:
            last_err = exc
    raise last_err


# ==================== 快照（只缓存数据） ====================

_SNAP = {}


def _snap_key(app_type, doc_full_name, sheet):
    return "%s|%s|%s" % (app_type, (doc_full_name or "").lower(), sheet or "")


def snapshot_get(app_type, doc_full_name, sheet):
    if SNAPSHOT_TTL <= 0:
        return None
    item = _SNAP.get(_snap_key(app_type, doc_full_name, sheet))
    if not item:
        return None
    age = time.time() - item["ts"]
    if age > SNAPSHOT_TTL:
        _SNAP.pop(_snap_key(app_type, doc_full_name, sheet), None)
        return None
    item["age_ms"] = int(age * 1000)
    return item


def snapshot_put(app_type, doc_full_name, sheet, grid, box):
    if SNAPSHOT_TTL <= 0:
        return
    cells = (box["rows"] or 0) * (box["cols"] or 0)
    if cells > SNAPSHOT_MAX_CELLS:
        return
    _SNAP[_snap_key(app_type, doc_full_name, sheet)] = {
        "ts": time.time(), "grid": grid, "box": box, "age_ms": 0}


def snapshot_invalidate(app_type=None, doc_full_name=None, sheet=None):
    """任何一次写操作后立即作废（下次读重新拉）。不传参数=全清。"""
    if app_type is None:
        _SNAP.clear()
        return
    if doc_full_name is None:
        for k in [k for k in _SNAP if k.startswith(app_type + "|")]:
            _SNAP.pop(k, None)
        return
    _SNAP.pop(_snap_key(app_type, doc_full_name, sheet), None)


def snapshot_info():
    """给返回值附上「你读的是缓存还是实时」——让人/模型自己判断新不新。"""
    return {"snapshot_ttl_ms": int(SNAPSHOT_TTL * 1000), "snapshot_keys": len(_SNAP)}


# ==================== 活动文档定位（薄壳的路径兜底） ====================


def _doc_ident(host, doc):
    """拿文档名 / 路径 / 未保存标记。编辑态旁路下部分字段可能为空，全部容错。"""
    return {
        "name": host._safe_com_str(doc, "Name", "") or "",
        "path": host._safe_com_str(doc, "FullName", "") or "",
        "unsaved": host._safe_com_bool(doc, "Saved", None) is False,
    }


def _thin_write_wait(host):
    raw = os.environ.get("WPS_MCP_THIN_WRITE_WAIT", "")
    try:
        if raw != "":
            return float(raw)
    except Exception:
        pass
    return float(getattr(host, "BUSY_RETRY_SECONDS", 5.0))


def locate_active_doc(host, wps_app, app_type, path=None, write=False):
    """定位「人当前正在操作的那份文档」。返回 (doc, mode)。wps_app 由调用方传入（同一句柄）。

    mode: "path"（给了绝对路径）| "single"（只开了一本）| "bypass"（编辑态只读旁路）

    定位顺序：
      1. 有路径 → 老地基按 FullName 精确匹配（自带重试/旁路/ESC 三级救援，直接用）。
      2. 无路径 → 先试「实例里只有一本已打开文档」（老地基的 single_open_doc）。
         这一步在**健康态**下就能成功，是最常见的路径，零额外代价。
      3. 仍失败且是**读** → 用窗口标题反推文件名，直接调老地基的 `_bypass_shim`
         造无按键旁路。这一步就是本模块给「无路径」补的救援（老地基的救援分支写在
         `if target:` 内部，不传路径根本进不去），否则编辑态下会直接抛 no_open_document。
      4. 仍失败且是**写** → 复用老地基拍板的策略：等 BUSY_RETRY_SECONDS（给人松手的机会）
         → 仍不行且允许就发一次 ESC → 再试一次。写**绝不走旁路**（实测编辑态下四种写法全挂，
         上旁路只会抛难懂的裸错误，这正是 2026-09-24 那次的教训）。
    最终仍失败就把最后一次结构化错误原样抛出，让 AI 照 human_action_required 请人协助。
    """
    def _try():
        return host._get_open_document(wps_app, path if path else None, app_type)

    last_err = None
    try:
        return _try(), ("path" if path else "single")
    except host.WpsAttachError as err:
        last_err = err

    if not write:
        caption = host._caption_doc_name(wps_app) or ""
        if caption and host._open_count(wps_app, app_type) == 1:
            # 把标题当 target 传给旁路：_bypass_shim 用 basename(target) 与标题比对，
            # 走老地基自己的判断，不重写也不复制它的逻辑。
            shim = host._bypass_shim(wps_app, app_type, path or caption)
            if shim is not None:
                return shim, "bypass"

    if write:
        deadline = time.time() + max(_thin_write_wait(host), 0.0)
        while time.time() < deadline:
            time.sleep(min(getattr(host, "BUSY_RETRY_STEP", 0.5), 0.5))
            try:
                return _try(), ("path" if path else "single")
            except host.WpsAttachError as err:
                last_err = err
        if getattr(host, "BUSY_ESC_ENABLED", True) and host._esc_exit_edit_mode(wps_app):
            time.sleep(0.5)
            try:
                return _try(), ("path" if path else "single")
            except host.WpsAttachError as err:
                last_err = err

    raise _reinterpret(host, wps_app, app_type, last_err, write)


def _reinterpret(host, wps_app, app_type, err, write):
    """把「没打开文档」这类误导性报错改写成「人正在编辑态」。

    老地基的救援分支写在 `if target:` 内部，无路径这条支路走不到，于是编辑态下会拿到
    reason=no_open_document / target_not_open —— 但文件明明开着（窗口标题就在眼前）。
    这类误报 2026-09-24 已经踩过一次（误报 target_not_open），这里按同样口径纠正，
    并补上 call_mode，让 AI 这轮就知道该请人做什么，而不是以为文件没开。
    """
    payload = dict(getattr(err, "payload", {}) or {})
    caption = host._caption_doc_name(wps_app) or ""
    if caption and host._open_count(wps_app, app_type) == 1 and \
            payload.get("reason") in ("no_open_document", "target_not_open", "path_required"):
        payload["reason"] = "com_name_resolution_refused"
        payload["hint"] = ("文档其实开着（窗口标题 = %r），但 WPS 拒绝按名字解析 COM 成员。"
                           "实测这是人正处于**单元格编辑态**（编辑框抱着某个单元格）的表现，"
                           "**不是**文件没打开。" % (caption,))
        payload["human_action_required"] = (
            "请按一下 Esc 或回车退出单元格编辑态，然后让我重试（退出后同一秒即可恢复）。"
            if write else
            "读已经读到内存里的实时内容了；若要写入，请按一下 Esc 或回车退出编辑态。")
        payload["call_mode"] = "write" if write else "read"
        return host.WpsAttachError(payload)
    return err


def with_doc(host, app, path, write, fn):
    """挂接 → 定位 → 干活 → 释放。app 句柄只在本次调用内存活，绝不跨请求攥着。"""
    app_type = _APP_TYPE[app]
    wps_app, attached = host._get_wps_app(app_type)
    try:
        if not attached:
            try:
                wps_app.Visible = False
            except Exception:
                pass
        doc, mode = locate_active_doc(host, wps_app, app_type, path, write)
    except Exception:
        host._release_app(wps_app, app_type, attached)
        raise
    try:
        result = fn(wps_app, doc, mode)
        if isinstance(result, dict):
            result.setdefault("attached", bool(attached))
            if mode == "bypass":
                result["live_read_via"] = "caption_bypass"
                result["note"] = ("人正停在某个单元格的编辑态里，本次是只读旁路（读到的是内存实时值）；"
                                  "要写的话请人先按一下 Esc 或回车退出编辑态。")
        return result
    finally:
        try:
            if not attached:
                doc.Close()
        except Exception:
            pass
        host._release_app(wps_app, app_type, attached)


# ==================== Excel 底座 ====================


def _sheet_of(host, doc, args):
    """取工作表对象。缺省用当前激活表（人正盯着的那张）。"""
    name = (args or {}).get("sheet")
    if name:
        return host._com_call(doc, "Sheets", name)
    return doc.ActiveSheet


def used_box(host, ws):
    """已用区域：首/末行列 + 地址。取不到时给 None 字段而不是抛错。"""
    ok, used = host._com_try_attr(ws, "UsedRange")
    if not ok or used is None:
        return None
    r = host._com_try_attr(used, "Row")[1]
    c = host._com_try_attr(used, "Column")[1]
    rows = host._com_try_attr(host._com_try_attr(used, "Rows")[1], "Count")[1]
    cols = host._com_try_attr(host._com_try_attr(used, "Columns")[1], "Count")[1]
    try:
        r, c, rows, cols = int(r), int(c), int(rows), int(cols)
    except Exception:
        return None
    return {"first_row": r, "first_col": c, "last_row": r + rows - 1, "last_col": c + cols - 1,
            "rows": rows, "cols": cols,
            "address": ref_text(r, c, r + rows - 1, c + cols - 1)}


def _read_grid(host, ws, r1, c1, r2, c2):
    """一次 Range.Value 读一片（不循环、不逐格）。"""
    left = host._com_call(ws, "Cells", r1, c1)
    right = host._com_call(ws, "Cells", r2, c2)
    rng = host._com_call(ws, "Range", left, right)
    ok, raw = host._com_try_attr(rng, "Value")
    # Range.Value 是「带可选参数的属性」，comtypes lazybind 只会返回 NamedProperty 包装器，
    # 不解包就进了返回值 → 上游 json.dumps 报 "Object of type NamedProperty is not
    # JSON serializable"（2026-10-20 真机实测：连读空格都会炸）。
    raw = host._com_to_plain(raw) if ok else None
    return coerce_grid(raw, r2 - r1 + 1, c2 - c1 + 1)


def _read_grid_attr(host, ws, r1, c1, r2, c2, attr):
    """一次读回 Range 的某个批量属性（目前用于 Formula：一次读完，用来认出"哪些格是公式"）。

    为什么要单独读 Formula：Range.Value 读到的是**计算结果**，只看它分不清"公式格"和普通格，
    而 replace 必须跳过公式格 —— 否则会把公式换成写死的值，把人的表改坏。
    """
    left = host._com_call(ws, "Cells", r1, c1)
    right = host._com_call(ws, "Cells", r2, c2)
    rng = host._com_call(ws, "Range", left, right)
    ok, raw = host._com_try_attr(rng, attr)
    # 同 _read_grid：Formula 也是可选参数属性，必须解包 NamedProperty，否则 json 序列化必炸。
    raw = host._com_to_plain(raw) if ok else None
    return coerce_grid(raw, r2 - r1 + 1, c2 - c1 + 1)


def _read_via_snapshot(host, ws, doc, sheet_name, box, r1, c1, r2, c2):
    """快照命中就切片，否则直读。快照只在 TTL>0 时才存在。"""
    ident = _doc_ident(host, doc)
    snap = snapshot_get("excel", ident["path"] or ident["name"], sheet_name)
    if snap is None:
        grid = _read_grid(host, ws, r1, c1, r2, c2)
        # 只在真的会缓存（TTL>0）时才预读整个 UsedRange：默认 TTL=0 时 snapshot_put 直接丢弃，
        # 无条件预读会让"读一格"附带上万格的 COM 读且结果被扔掉（1000×10 实测读 1 格拉 10011 格）。
        if box is not None and SNAPSHOT_TTL > 0:
            full = _read_grid(host, ws, box["first_row"], box["first_col"], box["last_row"], box["last_col"])
            snapshot_put("excel", ident["path"] or ident["name"], sheet_name, full, box)
        return grid, None
    sb = snap["box"]
    sr, sc = r1 - sb["first_row"], c1 - sb["first_col"]
    er, ec = r2 - sb["first_row"], c2 - sb["first_col"]
    rows = [[snap["grid"][r][c] if 0 <= r < len(snap["grid"]) and 0 <= c < len(snap["grid"][r])
             else None for c in range(sc, ec + 1)] for r in range(sr, er + 1)]
    return rows, snap["age_ms"]


def list_merges(host, ws, box):
    """列出合并区及其左上角锚点（解决「AI 不知道该往哪格写」）。

    快路径：整片 UsedRange 的 MergeCells 是 False → 一次调用返回「没有合并」；
           是 True → 整片就是一个合并区。只有返回 Null（有合并但不全是）才逐格扫，
           且受 WPS_MCP_MERGE_SCAN_MAX 限制。地址里带 $，就用 parse_ref 反解行列，省掉额外 COM 调用。
    """
    ok, used = host._com_try_attr(ws, "UsedRange")
    if not ok or used is None:
        return {"count": 0, "areas": [], "note": "读不到 UsedRange"}
    ok, flag = host._com_try_attr(used, "MergeCells")
    if ok and flag is False:
        return {"count": 0, "areas": []}
    if ok and flag is True:
        return {"count": 1, "areas": [_area(box["first_row"], box["first_col"],
                                            box["last_row"], box["last_col"], whole=True)]}

    cells = box["rows"] * box["cols"]
    if cells > MERGE_SCAN_MAX:
        return {"count": None, "areas": [], "scanned": False,
                "note": ("表内有合并区但需逐格确认，已用区域 %d 格超过扫描上限 %d。"
                         "可用 args.range 缩小范围后重问。" % (cells, MERGE_SCAN_MAX))}
    found = {}
    for r in range(box["first_row"], box["last_row"] + 1):
        for c in range(box["first_col"], box["last_col"] + 1):
            cell = host._com_call(ws, "Cells", r, c)
            ok, merged = host._com_try_attr(cell, "MergeCells")
            if not ok or not merged:
                continue
            ok, area = host._com_try_attr(cell, "MergeArea")
            if not ok or area is None:
                continue
            ok, addr = host._com_try_attr(area, "Address")
            if not ok or not addr:
                continue
            key = str(addr).replace("$", "").split("!")[-1]
            if key in found:
                continue
            ar = host._com_try_attr(area, "Row")[1]
            ac = host._com_try_attr(area, "Column")[1]
            rs = host._com_try_attr(host._com_try_attr(area, "Rows")[1], "Count")[1]
            cs = host._com_try_attr(host._com_try_attr(area, "Columns")[1], "Count")[1]
            try:
                found[key] = _area(int(ar), int(ac), int(ar) + int(rs) - 1, int(ac) + int(cs) - 1)
            except Exception:
                continue
    return {"count": len(found), "areas": sorted(found.values(), key=lambda a: (a["row"], a["col"]))}


def _area(r1, c1, r2, c2, whole=False):
    return {"addr": ref_text(r1, c1, r2, c2), "row": r1, "col": c1,
            "rows": r2 - r1 + 1, "cols": c2 - c1 + 1,
            "anchor": ref_text(r1, c1, r1, c1), "whole_used_range": whole}


def _merge_area_desc(host, area):
    """把一个 MergeArea 描述成 {addr, anchor, rows, cols, end_row, end_col}；读不到返回 None。"""
    ok, addr = host._com_try_attr(area, "Address")
    if not ok or addr is None:
        return None
    try:
        key = str(host._com_to_plain(addr)).replace("$", "").split("!")[-1]
        ar = int(host._com_to_plain(host._com_try_attr(area, "Row")[1]))
        ac = int(host._com_to_plain(host._com_try_attr(area, "Column")[1]))
        rs = int(host._com_to_plain(host._com_try_attr(host._com_try_attr(area, "Rows")[1], "Count")[1]))
        cs = int(host._com_to_plain(host._com_try_attr(host._com_try_attr(area, "Columns")[1], "Count")[1]))
    except Exception:
        return None
    return {"addr": key, "anchor": ref_text(ar, ac, ar, ac), "row": ar, "col": ac,
            "end_row": ar + rs - 1, "end_col": ac + cs - 1, "rows": rs, "cols": cs}


def merge_scope(host, ws, r1, c1, r2, c2):
    """目标范围的合并区情况：是否合并 / 整块多大 / 锚点在哪 / 是否只与合并区「部分相交」。

    为什么要先问再写：合并区只有左上角锚点存值，其余格是幻影——
    往区内非锚点写会落空（WPS 静默吞掉），往「合并区的一部分」写会丢值。
    COM 的 MergeCells 是三态：False 整片无合并 / True 整片就是一个合并区 / Null 有合并但不全覆盖。
    返回 covers=True 表示请求范围被单一合并区完整覆盖（可整块写）；partial=True 表示有合并区
    只被请求范围盖住一部分（此时**不该写**，会丢值）。
    """
    out = {"range": ref_text(r1, c1, r2, c2), "merged": False,
           "areas": [], "covers": False, "partial": False}
    rng = _range_of(host, ws, r1, c1, r2, c2)
    ok, flag = host._com_try_attr(rng, "MergeCells")
    flag = host._com_to_plain(flag) if ok else None
    if flag is False:
        return out
    if flag is True:
        desc = _merge_area_desc(host, host._com_try_attr(rng, "MergeArea")[1])
        if desc is None:
            return out
        desc["contains_range"] = True
        out.update({"merged": True, "areas": [desc], "covers": True})
        return out

    # Null：范围内有合并但不全覆盖 → 逐格收集相交的合并区（受 MERGE_SCAN_MAX 限流）
    cells = (r2 - r1 + 1) * (c2 - c1 + 1)
    if cells > MERGE_SCAN_MAX:
        out["merged"] = True
        out["scanned"] = False
        out["note"] = ("范围内有合并区，但 %d 格超过逐格扫描上限 %d，无法逐一确认。"
                       "请把范围收小到合并区附近再问（如单格或单行）。" % (cells, MERGE_SCAN_MAX))
        return out
    found = {}
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = host._com_call(ws, "Cells", r, c)
            ok2, merged = host._com_try_attr(cell, "MergeCells")
            if not ok2 or not host._com_to_plain(merged):
                continue
            ok3, area = host._com_try_attr(cell, "MergeArea")
            if not ok3 or area is None:
                continue
            desc = _merge_area_desc(host, area)
            if desc and desc["addr"] not in found:
                found[desc["addr"]] = desc
    areas = sorted(found.values(), key=lambda a: (a["row"], a["col"]))
    for a in areas:
        a["contains_range"] = (a["row"] <= r1 and a["end_row"] >= r2 and
                               a["col"] <= c1 and a["end_col"] >= c2)
    out["merged"] = bool(areas)
    out["areas"] = areas
    out["covers"] = len(areas) == 1 and areas[0]["contains_range"]
    out["partial"] = any(not a["contains_range"] for a in areas)
    return out


def merge_anchor(host, ws, r, c):
    """(r,c) 若落在合并区里且不是左上角锚点，返回锚点 (r,c)；否则返回原值。

    实测要求：合并单元格只能写左上角。与其让模型自己去查，不如写的时候顺手纠正并如实回报。
    保留此函数（保持既有行为不变的调用方可用），实现统一走 merge_scope。
    """
    scope = merge_scope(host, ws, r, c, r, c)
    if not scope["merged"]:
        return r, c, False
    a = scope["areas"][0]
    if (a["row"], a["col"]) == (r, c):
        return r, c, False
    return a["row"], a["col"], True


def data_span(host, ws, box):
    """数据区的列跨度 = 表头行（已用区首行）里第一个到最后一个非空列。

    为什么不用 UsedRange 的 first_col：**它会漂**。2026-09-28 真机实测：在 C:H 有数据的表里
    往 A18 写一格，UsedRange 立刻变成 A1:H18，于是随后「整行写 20:20」原本该落到 C20:H20，
    却跑去了 A20:C20 —— 人看到的就是"写歪了"。表头行几乎不会被顺手写坏，拿它当数据原点稳得多。
    读整行也走同一个跨度，保证读写对齐。
    """
    if box is None:
        return None, None
    headers = _read_grid(host, ws, box["first_row"], box["first_col"],
                         box["first_row"], box["last_col"])[0]
    first = last = None
    for i, v in enumerate(headers):
        if v is not None and str(v).strip() != "":
            if first is None:
                first = box["first_col"] + i
            last = box["first_col"] + i
    if first is None or last is None:
        return box["first_col"], box["last_col"]
    return first, last


def next_empty_row(host, ws, box, key_col=None):
    """下一空行：在关键列（默认已用区域首列）里找最后一个非空，+1。

    不是 UsedRange.Rows.Count+1——表头下面留空行时那个算法会给出错的行号。
    """
    if box is None:
        return 1
    col = key_col or box["first_col"]
    col = col_to_index(col) or box["first_col"]
    vals = _read_grid(host, ws, box["first_row"], col, box["last_row"], col)
    last = box["first_row"] - 1
    for i, row in enumerate(vals):
        v = row[0]
        if v is not None and str(v).strip() != "":
            last = box["first_row"] + i
    return last + 1


# ==================== Excel 动作 ====================


def _excel_where(host, doc, mode, args):
    args = args or {}
    ws = _sheet_of(host, doc, args)
    box = used_box(host, ws)
    key_col = col_to_index(args.get("key_col")) or (box or {}).get("first_col")
    span = data_span(host, ws, box)
    out = {
        "app": "excel",
        "doc": _doc_ident(host, doc),
        "sheet": host._safe_com_str(ws, "Name", ""),
        "used": box,
        "data_cols": ref_text((box or {}).get("first_row", 1), span[0], (box or {}).get("first_row", 1), span[1])
        if box and span[0] else None,
        "next_row": next_empty_row(host, ws, box, key_col),
        "next_row_key_col": col_letter(key_col) if key_col else None,
        "merges": list_merges(host, ws, box) if box else {"count": 0, "areas": []},
    }
    sheets = []
    try:
        parent = ws.Parent
        cnt = int(host._com_try_attr(parent, "Sheets")[1].Count)
        sheets = [host._safe_com_str(host._com_call(parent, "Sheets", i), "Name", "")
                  for i in range(1, cnt + 1)]
    except Exception:
        pass
    if sheets:
        out["sheets"] = sheets
    out.update(host._match_hint("excel"))
    return out


def _excel_get(host, doc, mode, args):
    args = args or {}
    ws = _sheet_of(host, doc, args)
    box = used_box(host, ws)
    ref = parse_ref(args.get("ref") or args.get("addr"))
    if ref["sheet"]:
        ws = host._com_call(doc, "Sheets", ref["sheet"])
    if box is None:
        box = {"first_row": 1, "first_col": 1, "last_row": 1, "last_col": 1, "rows": 1, "cols": 1}
    # 缺省的一侧按数据区补齐：C:C = 该列 × 数据行区间；3:3 = 该行 × 数据列区间
    span = data_span(host, ws, box)
    r1 = ref["r1"] if ref["r1"] is not None else box["first_row"]
    r2 = ref["r2"] if ref["r2"] is not None else box["last_row"]
    c1 = ref["c1"] if ref["c1"] is not None else (span[0] or box["first_col"])
    c2 = ref["c2"] if ref["c2"] is not None else (span[1] or box["last_col"])

    total_rows = r2 - r1 + 1
    max_rows = args.get("max_rows")
    truncated = False
    if max_rows:
        try:
            max_rows = int(max_rows)
        except Exception:
            max_rows = None
    if max_rows and total_rows > max_rows:
        r2 = r1 + max_rows - 1
        truncated = True
    if (r2 - r1 + 1) * (c2 - c1 + 1) > READ_MAX_CELLS:
        raise ValueError(
            "这次要读 %d 格，超过单次上限 %d。请缩小 ref，或传 args.max_rows 只取前若干行。"
            % ((r2 - r1 + 1) * (c2 - c1 + 1), READ_MAX_CELLS))

    grid, age = _read_via_snapshot(host, ws, doc, host._safe_com_str(ws, "Name", ""), box, r1, c1, r2, c2)
    out = {
        "app": "excel",
        "sheet": host._safe_com_str(ws, "Name", ""),
        "ref": ref_text(r1, c1, r2, c2),
        "shape": [r2 - r1 + 1, c2 - c1 + 1],
        "values": shape_values(grid),
        "live": True,
    }
    if truncated:
        out["truncated"] = True
        out["total_rows"] = total_rows
    if box:
        out["used"] = box["address"]
    if age is not None:
        out["from_snapshot_ms"] = age
    # 读的时候就把「是不是合并单元格、整块多大」一起回：AI 不必再单独猜/试探
    # （2026-10-20：人反馈写到合并区时没被当成合并区处理，根子就是读回来的信息里没这一项）
    out["merge"] = merge_scope(host, ws, r1, c1, r2, c2)
    # 读到 =DISPIMG 就提示一句：这类格子值只有一个公式，图得另走 image 动作
    hits = _dispimg_cells(out["values"])
    if hits:
        out["images"] = ("范围里有 %d 个 WPS 嵌入图片（=DISPIMG）。要看图请用 "
                         "action=\"image\" 传同一个 ref，会落成文件并返回路径。" % hits)
    return out


def _dispimg_cells(values):
    """数一下读回来的值里有多少个 =DISPIMG（不解析 ID，只做发现性提示）。"""
    if isinstance(values, str):
        return 1 if "DISPIMG(" in values.upper() else 0
    if isinstance(values, (list, tuple)):
        return sum(_dispimg_cells(v) for v in values)
    return 0


def _excel_set(host, doc, mode, args):
    args = args or {}
    ws = _sheet_of(host, doc, args)
    ref = parse_ref(args.get("ref") or args.get("addr"))
    if ref["sheet"]:
        ws = host._com_call(doc, "Sheets", ref["sheet"])
    box = used_box(host, ws)
    if "values" not in args and "value" not in args:
        # 必须显式带 values：2026-09-28 真机踩过——漏传 values 时 args.get 返回 None，
        # 单格会被"成功地"写成空值，看着像成功其实是把人的数据抹了。
        raise ValueError(
            "set 缺少 values（要写入的值）。单格传 \"values\": \"x\"，整行传 [a,b,c]，"
            "整片传 [[a,b,c],[d,e,f]]；确实要清空某格请显式传 \"values\": null。")
    raw = args.get("values", args.get("value"))
    span = data_span(host, ws, box)
    r1, c1, r2, c2 = resolve_write_box(ref, box, raw, origin_c=span[0],
                                       origin_r=(box or {}).get("first_row"))

    requested = ref_text(r1, c1, r2, c2)
    anchored_from = None
    # 写入前先查合并区（2026-10-20 人要求 + 真实事故：往合并区的一部分写会静默丢值）：
    #   covers  → 整块就是一个合并区，照写（值落在锚点）
    #   partial → 请求范围只盖住某个合并区的一部分 → **直接拒绝**，别赌
    #   单格在区内非锚点 → 落到左上锚点并回报整块尺寸
    scope = merge_scope(host, ws, r1, c1, r2, c2)
    merge_info = None
    if scope["merged"]:
        bad = [a for a in scope["areas"] if not a.get("contains_range")]
        if bad:
            raise ValueError(
                "目标范围 %s 只覆盖到合并区 %s 的一部分，写进去会丢值（合并区只有左上锚点能存值）。"
                "请改成整块覆盖（如 %s）或只写锚点 %s；要按格写请先 unmerge。"
                % (scope["range"],
                   "、".join(a["addr"] for a in bad),
                   "、".join(a["addr"] for a in bad),
                   "、".join(a["anchor"] for a in bad)))
        area = scope["areas"][0]
        merge_info = {
            "area": area["addr"], "anchor": area["anchor"],
            "rows": area["rows"], "cols": area["cols"],
            "note": "目标在合并区 %s（%d行×%d列），值只会存在锚点 %s"
                    % (area["addr"], area["rows"], area["cols"], area["anchor"]),
        }
        if (r1, c1) != (area["row"], area["col"]):
            anchored_from = requested
            r1 = r2 = area["row"]
            c1 = c2 = area["col"]
            merge_info["note"] = ("目标 %s 在合并区 %s 内部但不是锚点，已自动改写到锚点 %s"
                                  "（%d行×%d列）"
                                  % (requested, area["addr"], area["anchor"],
                                     area["rows"], area["cols"]))

    rows, cols = r2 - r1 + 1, c2 - c1 + 1
    if raw is None and rows * cols > 1:
        # V3.2：整片清空走 ClearContents（实测本机 WPS 可用）。比写一坨 null 又快又干净，
        # 而且**保留格式与合并**（Cells.Clear 才会连格式一起清，这里刻意不用它）。
        rng0 = _range_of(host, ws, r1, c1, r2, c2)
        try:
            host._com_call(rng0, "ClearContents")
        except Exception:
            set_range_value(rng0, tuple(tuple(None for _ in range(cols)) for _ in range(rows)))
        ident = _doc_ident(host, doc)
        snapshot_invalidate("excel", ident["path"] or ident["name"],
                            host._safe_com_str(ws, "Name", ""))
        return {"app": "excel", "sheet": host._safe_com_str(ws, "Name", ""),
                "written": ref_text(r1, c1, r2, c2), "shape": [rows, cols],
                "cleared": True, "saved": False, "undoable": True,
                "note": "整片清空（保留格式与合并单元格）"}
    value = to_matrix(raw, rows, cols)

    left = host._com_call(ws, "Cells", r1, c1)
    right = host._com_call(ws, "Cells", r2, c2)
    rng = host._com_call(ws, "Range", left, right)
    set_range_value(rng, value)

    ident = _doc_ident(host, doc)
    snapshot_invalidate("excel", ident["path"] or ident["name"], host._safe_com_str(ws, "Name", ""))

    saved = None
    if AUTOSAVE:
        try:
            doc.Save()
            saved = True
        except Exception:
            saved = False

    out = {
        "app": "excel",
        "sheet": host._safe_com_str(ws, "Name", ""),
        "written": ref_text(r1, c1, r2, c2),
        "shape": [rows, cols],
        "saved": saved if AUTOSAVE else False,
        "undoable": not AUTOSAVE,
    }
    if anchored_from:
        out["anchored_from"] = anchored_from
        out["anchored_note"] = "目标格在合并区内部，已自动写到左上角锚点（合并区只能写锚点）"
    if merge_info:
        out["merge"] = merge_info
    if ref["kind"] == "row" and span[0] and span[1]:
        width = span[1] - span[0] + 1
        if cols != width:
            # 不拦、不截断，只如实告知：表头列宽与给的值个数不一致时，
            # 读整行只回到表头跨度，多写出去的那些列不会出现在 get "26:26" 里。
            out["width_mismatch"] = {"header_cols": ref_text(1, span[0], 1, span[1]) + "→%d 列" % width,
                                     "values": cols,
                                     "note": "值个数与表头列宽不同；本次按值个数写。"
                                             "要精确控制请直接给列范围，如 \"G26:I26\"。"}
    return out


# ==================== Excel 日常格式/结构动作（V3.2） ====================
# 为什么加这一批：实战测试发现薄壳只有 where/get/set/activate/save，
# 「列宽、行高、填充色、字体、合并/拆分、格式刷、行列增删」这些天天要用的动作
# 都只能走老工具，而老工具在薄模式下**不广播**、legacy 逃生口又只放读类 → 等于做不到。
# 这些动作的参数说明写在 HELP 里，**不占常驻 schema**（tools/list 仍是 2 个工具 / 755 字符）。
#
# 统一口径：一律用 ref（"B5"/"3:3"/"C:C"/"A2:D10"），与 get/set 完全一致。

def _prop(host, obj, name):
    """取 COM **属性**（Interior / Font 这类是属性，不是方法）。

    踩过的坑：`_com_call(rng, "Interior")` 会先 getattr 拿到 Interior 对象、再把它当函数**调用**，
    于是报 `找不到成员 (-2147352573)`。属性必须直接取。
    """
    ok, value = host._com_try_attr(obj, name)
    if not ok or value is None:
        raise ValueError("取不到 COM 属性 %s（对象类型 %s）" % (name, type(obj).__name__))
    return value


def _find_merge_in(host, ws, r1, c1, r2, c2, cap=2000):
    """在给定范围内找第一个合并区，返回去掉 $ 的地址；没有就 None。

    为什么要自己找：实测 `Range("A1:D1").MergeArea` 在"部分合并"（区里既有合并格又有普通格）
    时返回 **None**，所以不能靠它拿旧合并区。
    """
    if (r2 - r1 + 1) * (c2 - c1 + 1) > cap:
        return None
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = host._com_call(ws, "Cells", r, c)
            ok, merged = host._com_try_attr(cell, "MergeCells")
            if not ok or not merged:
                continue
            ok2, area = host._com_try_attr(cell, "MergeArea")
            if not ok2 or area is None:
                continue
            ok3, addr = host._com_try_attr(area, "Address")
            if ok3 and addr:
                return str(addr).replace("$", "").split("!")[-1]
    return None


def _range_of(host, ws, r1, c1, r2, c2):
    left = host._com_call(ws, "Cells", r1, c1)
    right = host._com_call(ws, "Cells", r2, c2)
    return host._com_call(ws, "Range", left, right)


def _target_box(ref, box, default_edge="used"):
    """把 ref 收敛成确定范围。缺省的一侧按数据区补齐（与 get 同口径用表头跨度）。"""
    first_row = (box or {}).get("first_row", 1)
    last_row = (box or {}).get("last_row", first_row)
    r1 = ref["r1"] if ref["r1"] is not None else first_row
    r2 = ref["r2"] if ref["r2"] is not None else last_row
    c1 = ref["c1"] if ref["c1"] is not None else (box or {}).get("first_col", 1)
    c2 = ref["c2"] if ref["c2"] is not None else (box or {}).get("last_col", c1)
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def _excel_width(host, doc, mode, args):
    """设置列宽。args: col（"A" 或 "A:C"）+ w（Excel 列宽单位，约等于字符数）。"""
    args = args or {}
    ws = _sheet_of(host, doc, args)
    spec = str(args.get("col") or args.get("cols") or "").strip()
    if not spec:
        raise ValueError('width 需要 args.col（如 "A" 或 "A:C"）与 args.w（列宽，如 20）')
    ref = parse_ref(spec if ":" in spec else spec + ":" + spec)
    if ref["c1"] is None or ref["c2"] is None:
        raise ValueError('width 的 col 只能给列（"A" 或 "A:C"），收到 %r' % (spec,))
    w = args.get("w", args.get("width"))
    if w is None:
        raise ValueError("width 需要 args.w（列宽数值，行高才用 height）")
    c1, c2 = sorted((ref["c1"], ref["c2"]))
    cols = host._com_call(ws, "Columns",
                          "%s:%s" % (col_letter(c1), col_letter(c2)) if c2 > c1 else col_letter(c1))
    old = host._com_try_attr(cols, "ColumnWidth")[1]
    _set_prop(cols, "ColumnWidth", float(w))
    return {"app": "excel", "sheet": host._safe_com_str(ws, "Name", ""),
            "columns": "%s:%s" % (col_letter(c1), col_letter(c2)), "width": float(w),
            "width_before": old}


def _set_prop(obj, name, value):
    """给 COM 对象设属性（属性名可能只认 DISPID，先名字后直调）。"""
    try:
        setattr(obj, name, value)
        return
    except Exception as first:
        last = first
    dispid = _HOST_DISPATCH.get(name)
    inner = getattr(obj, "_oleobj_", None) or getattr(obj, "_comobj", None)
    if inner is not None and dispid is not None:
        for invkind in (4, 8):
            try:
                inner._invoke(dispid, invkind, 0, value)
                return
            except Exception as err:
                last = err
    raise last


# 少数属性在某些绑定下名字查不到，给 DISPID 兜底（Excel/WPS 共用）
_HOST_DISPATCH = {"ColumnWidth": 0x0000026D, "RowHeight": 0x0000026F,
                  "NumberFormat": 0x000000C1, "Merge": 0x000000CE, "UnMerge": 0x0000023E,
                  "PasteSpecial": 0x00000788, "Copy": 0x0000022B, "Insert": 0x000000FB,
                  "Delete": 0x00000075, "CutCopyMode": 0x0000019E}


def _excel_height(host, doc, mode, args):
    """设置行高。args: row（"3" 或 "3:6"）+ h。"""
    args = args or {}
    ws = _sheet_of(host, doc, args)
    spec = str(args.get("row") or args.get("rows") or "").strip()
    if not spec:
        raise ValueError('height 需要 args.row（如 "3" 或 "3:6"）与 args.h（行高，如 30）')
    ref = parse_ref(spec if ":" in spec else spec + ":" + spec)
    if ref["r1"] is None or ref["r2"] is None:
        raise ValueError('height 的 row 只能给行（"3" 或 "3:6"），收到 %r' % (spec,))
    h = args.get("h", args.get("height"))
    if h is None:
        raise ValueError("height 需要 args.h（行高数值）")
    r1, r2 = sorted((ref["r1"], ref["r2"]))
    rows = host._com_call(ws, "Rows", "%d:%d" % (r1, r2) if r2 > r1 else str(r1))
    old = host._com_try_attr(rows, "RowHeight")[1]
    _set_prop(rows, "RowHeight", float(h))
    return {"app": "excel", "sheet": host._safe_com_str(ws, "Name", ""),
            "rows": "%d:%d" % (r1, r2), "height": float(h), "height_before": old}


def _excel_fill(host, doc, mode, args):
    """给 ref 区域填背景色。args: ref + color（"#RRGGBB"，None 表示清除填充）。"""
    args = args or {}
    ws = _sheet_of(host, doc, args)
    ref = parse_ref(args.get("ref") or args.get("addr"))
    if ref["sheet"]:
        ws = host._com_call(doc, "Sheets", ref["sheet"])
    box = used_box(host, ws)
    r1, c1, r2, c2 = _target_box(ref, box)
    rng = _range_of(host, ws, r1, c1, r2, c2)
    color = args.get("color", args.get("bg_color"))
    interior = _prop(host, rng, "Interior")
    if color in (None, "", "none", "clear"):
        _set_prop(interior, "Pattern", -4142)          # xlNone
        done = None
    else:
        _set_prop(interior, "Pattern", 1)              # xlSolid
        _set_prop(interior, "Color", host._com_color(color))
        done = color
    return {"app": "excel", "sheet": host._safe_com_str(ws, "Name", ""),
            "filled": ref_text(r1, c1, r2, c2), "color": done}


def _excel_style(host, doc, mode, args):
    """改字体/对齐/数字格式。args: ref + bold/italic/size/name/color/underline/align/align_v/format。"""
    args = args or {}
    ws = _sheet_of(host, doc, args)
    ref = parse_ref(args.get("ref") or args.get("addr"))
    if ref["sheet"]:
        ws = host._com_call(doc, "Sheets", ref["sheet"])
    box = used_box(host, ws)
    r1, c1, r2, c2 = _target_box(ref, box)
    rng = _range_of(host, ws, r1, c1, r2, c2)
    applied = {}
    font = None
    for key, prop, conv in (("bold", "Bold", bool), ("italic", "Italic", bool),
                            ("size", "Size", float), ("name", "Name", str)):
        if args.get(key) is not None:
            font = font or _prop(host, rng, "Font")
            _set_prop(font, prop, conv(args[key]))
            applied[key] = args[key]
    if args.get("color") is not None:
        font = font or _prop(host, rng, "Font")
        _set_prop(font, "Color", host._com_color(args["color"]))
        applied["color"] = args["color"]
    if args.get("underline") is not None:
        font = font or _prop(host, rng, "Font")
        _set_prop(font, "Underline", host._underline_pair(args["underline"])[0])
        applied["underline"] = args["underline"]
    amap = {"left": -4131, "center": -4108, "right": -4152, "justify": -4130}
    vmap = {"top": -4160, "center": -4108, "bottom": -4107}
    if args.get("align") or args.get("align_h"):
        key = str(args.get("align") or args.get("align_h")).lower()
        _set_prop(rng, "HorizontalAlignment", amap.get(key, -4131))
        applied["align"] = key
    if args.get("align_v"):
        key = str(args["align_v"]).lower()
        _set_prop(rng, "VerticalAlignment", vmap.get(key, -4160))
        applied["align_v"] = key
    if args.get("format") is not None:
        _set_prop(rng, "NumberFormat", str(args["format"]))
        applied["format"] = args["format"]
    if not applied:
        raise ValueError(
            "style 至少要给一项：bold/italic/size/name/color/underline/align/align_v/format"
            "（如 args={\"ref\":\"A1\",\"bold\":true,\"size\":14}）")
    return {"app": "excel", "sheet": host._safe_com_str(ws, "Name", ""),
            "styled": ref_text(r1, c1, r2, c2), "applied": applied}


def _excel_merge(host, doc, mode, args):
    """合并 ref 区域。若与既有合并区重叠会**先拆开原来的**再按新范围合并 ——
    所以「扩大合并」「缩小合并」都是同一个动作。"""
    args = args or {}
    ws = _sheet_of(host, doc, args)
    ref = parse_ref(args.get("ref") or args.get("addr"))
    if ref["sheet"]:
        ws = host._com_call(doc, "Sheets", ref["sheet"])
    if ref["r1"] is None or ref["c1"] is None or ref["r2"] is None or ref["c2"] is None:
        raise ValueError('merge 需要完整区域，如 args={"ref": "A1:C1"}')
    r1, c1, r2, c2 = _target_box(ref, None)
    if (r1, c1) == (r2, c2):
        raise ValueError("merge 至少要有两个格（单格没有可合并的对象）")

    previous = None
    probe = _range_of(host, ws, r1, c1, r2, c2)
    ok, merged = host._com_try_attr(probe, "MergeCells")
    if ok and merged is not False:
        # 先找出原来那个合并区长什么样（用于如实回报"扩大/缩小"），再整片拆掉
        previous = _find_merge_in(host, ws, r1, c1, r2, c2)
        host._com_call(probe, "UnMerge")
    host._com_call(probe, "Merge")
    ident = _doc_ident(host, doc)
    snapshot_invalidate("excel", ident["path"] or ident["name"], host._safe_com_str(ws, "Name", ""))
    return {"app": "excel", "sheet": host._safe_com_str(ws, "Name", ""),
            "merged": ref_text(r1, c1, r2, c2), "previous_merge": previous,
            "note": "合并区只能写左上角锚点；用 get 读时非锚点格会是 null"}


def _excel_unmerge(host, doc, mode, args):
    """拆分 ref 区域（若 ref 落在更大的合并区里，会拆掉那个完整的合并区）。"""
    args = args or {}
    ws = _sheet_of(host, doc, args)
    ref = parse_ref(args.get("ref") or args.get("addr"))
    if ref["sheet"]:
        ws = host._com_call(doc, "Sheets", ref["sheet"])
    if ref["r1"] is None or ref["c1"] is None:
        raise ValueError('unmerge 需要完整区域，如 args={"ref": "A1:C1"}')
    r1, c1, r2, c2 = _target_box(ref, None)
    rng = _range_of(host, ws, r1, c1, r2, c2)
    ok, merged = host._com_try_attr(rng, "MergeCells")
    if ok and merged is False:
        return {"app": "excel", "unmerged": None, "was_merged": False,
                "note": "%s 本来就不是合并区" % ref_text(r1, c1, r2, c2)}
    area = rng
    ok2, ma = host._com_try_attr(rng, "MergeArea")
    if ok2 and ma is not None:
        area = ma
    ok3, addr = host._com_try_attr(area, "Address")
    addr_txt = str(addr).replace("$", "").split("!")[-1] if ok3 and addr else         (_find_merge_in(host, ws, r1, c1, r2, c2) or ref_text(r1, c1, r2, c2))
    host._com_call(area, "UnMerge")
    ident = _doc_ident(host, doc)
    snapshot_invalidate("excel", ident["path"] or ident["name"], host._safe_com_str(ws, "Name", ""))
    return {"app": "excel", "sheet": host._safe_com_str(ws, "Name", ""),
            "unmerged": addr_txt, "was_merged": True}


def _bool_arg(value, default=False, name="参数"):
    """把 args 里的布尔写成真布尔。模型有时传字符串 "true"/"false"，别用 truthy 判断（"false" 也是真）。"""
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    if text in ("1", "true", "yes", "y", "on", "带", "带格式", "有"):
        return True
    if text in ("0", "false", "no", "n", "off", "", "不带", "不带格式", "无"):
        return False
    raise ValueError('%s 只能是 true / false，收到 %r' % (name, value))


def _spanned_box(host, ws, ref, box):
    """把 ref 缺省的一侧按「数据区」补齐 —— 与 get 同口径
    （整行走表头跨度、整列走数据行区间，不是 UsedRange 首列，那个会被散格拽跑）。"""
    if box is None:
        box = {"first_row": 1, "first_col": 1, "last_row": 1, "last_col": 1}
    span = data_span(host, ws, box)
    r1 = ref["r1"] if ref["r1"] is not None else box["first_row"]
    r2 = ref["r2"] if ref["r2"] is not None else box["last_row"]
    c1 = ref["c1"] if ref["c1"] is not None else (span[0] or box["first_col"])
    c2 = ref["c2"] if ref["c2"] is not None else (span[1] or box["last_col"])
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def _excel_copy(host, doc, mode, args):
    """把 from 区域搬到 to。args: from, to, with_format（默认 false）。

    为什么需要这个动作：没有它，「把 C3 复制到 B2」必须由模型**把值读进上下文再原样吐出去**
    （两次调用夹着一次搬运）。复制 1000 格或一段长文本时，token 翻倍、还可能抄错/截断。
    有了 copy，值只在 WPS 内部走，一次调用完成，**值永远不进模型上下文**。

    with_format=false（默认）—— 只搬**值**：
        读 src.Value → 写 dst.Value，一次赋值。不碰系统剪贴板，不动格式。
        公式会落地成「当前算出来的值」（不是公式文本）；要公式文本请用 with_format=true。
    with_format=true —— 走 COM 原生 `src.Copy(dst)`（等于人手 Ctrl+C / Ctrl+V）：
        值 + 公式 + 数字格式 + 字体/底色/边框 + 合并单元格一起搬，**相对引用会自动调整**。
        代价：占用系统剪贴板（会替换人的 Ctrl+V 内容）；合并区也按 Excel 语义整块粘过去。

    to 只给一个格也认：粘贴语义就是从左上角展开（源 A1:C3 → to="G1" 落 G1:I3）。

    跨表：from / to 各自可带表前缀（"Sheet2!A1:C3"），或分别用 from_sheet / to_sheet。
    to 不给表时**跟随源表**（同表复制最常见）；args.sheet 同时给两边打底。
    """
    args = args or {}
    raw_from = args.get("from") or args.get("src") or args.get("source")
    raw_to = args.get("to") or args.get("dst") or args.get("target")
    # 先点名缺的是谁：直接 parse_ref(None) 只会说"缺少 ref"，模型不知道该补 from 还是 to。
    if raw_from in (None, ""):
        raise ValueError('copy 缺少 from（源区域，如 "C3" 或 "A1:C3"；跨表写 "Sheet1!A1:C3"）')
    if raw_to in (None, ""):
        raise ValueError('copy 缺少 to（目标左上角，如 "B2"；要写到别的表就写 "Sheet2!B2"）')
    src_ref = parse_ref(raw_from)
    dst_ref = parse_ref(raw_to)
    if src_ref["r1"] is None and src_ref["c1"] is None:
        # 只给了一侧也算合法（"3:3" 整行 / "C:C" 整列），那一侧按数据区补齐；
        # 两侧都没有才真的没法定位。
        raise ValueError('copy 的 from 至少要有行或列（如 "C3"、"A1:C3"、"3:3"、"C:C"），收到 %r'
                         % (raw_from,))
    if dst_ref["r1"] is None or dst_ref["c1"] is None:
        raise ValueError('copy 的 to 需要明确的行列（如 "B2"；只当左上角用，不用写全区域），收到 %r'
                         % (raw_to,))

    base = args.get("sheet")
    from_name = src_ref["sheet"] or args.get("from_sheet") or base
    to_name = dst_ref["sheet"] or args.get("to_sheet") or from_name
    src_ws = host._com_call(doc, "Sheets", from_name) if from_name else doc.ActiveSheet
    dst_ws = src_ws if to_name == from_name else (
        host._com_call(doc, "Sheets", to_name) if to_name else doc.ActiveSheet)

    r1, c1, r2, c2 = _spanned_box(host, src_ws, src_ref, used_box(host, src_ws))
    rows, cols = r2 - r1 + 1, c2 - c1 + 1
    src = _range_of(host, src_ws, r1, c1, r2, c2)

    dr, dc = dst_ref["r1"], dst_ref["c1"]
    dst_tl = host._com_call(dst_ws, "Cells", dr, dc)
    with_format = _bool_arg(args.get("with_format", args.get("format")), False, "with_format")

    src_txt = ref_text(r1, c1, r2, c2)
    dst_txt = ref_text(dr, dc, dr + rows - 1, dc + cols - 1)
    src_sheet = host._safe_com_str(src_ws, "Name", "")
    dst_sheet = host._safe_com_str(dst_ws, "Name", "")

    if with_format:
        # COM 原生整块复制：最忠实（含合并、边框、相对引用调整），一步到位。
        try:
            host._com_call(src, "Copy", dst_tl)
        except Exception as err:
            # 有些版本 Copy 的 Destination 走不了位置参数 → 退化成"复制到剪贴板 + 目标粘贴"。
            try:
                host._com_call(src, "Copy")
                host._com_call(dst_tl, "PasteSpecial")
            except Exception:
                raise ValueError(
                    "带格式复制失败（%s）。可改用 with_format=false 只搬值；"
                    "若必须带格式，请人确认 WPS 没有弹窗/处于单元格编辑态后重试。"
                    % (str(err)[:120],))
        how = "com_copy"
        note = ("已连格式一起搬（值+公式+格式+合并；相对引用按 Excel 语义自动调整）。"
                "本次占用过系统剪贴板。")
    else:
        # 只搬值：读进服务端内存 → 直接赋给目标，**完全不经过模型上下文**，也不碰剪贴板。
        value = host._com_try_attr(src, "Value")[1]
        if rows * cols > 1:
            dst = _range_of(host, dst_ws, dr, dc, dr + rows - 1, dc + cols - 1)
        else:
            dst = dst_tl
        set_range_value(dst, value)
        how = "values_only"
        note = ("只搬了值，没动格式（目标原有的底色/边框/数字格式保持原样）；"
                "源里的公式落地成当前算出来的值。要连公式和格式一起搬，传 with_format=true。")

    ident = _doc_ident(host, doc)
    snapshot_invalidate("excel", ident["path"] or ident["name"], dst_sheet)
    if src_sheet != dst_sheet:
        snapshot_invalidate("excel", ident["path"] or ident["name"], src_sheet)

    return {"app": "excel", "copied": {"from": src_txt, "to": dst_txt,
                                       "from_sheet": src_sheet, "to_sheet": dst_sheet,
                                       "shape": [rows, cols]},
            "with_format": with_format, "how": how,
            "saved": False, "undoable": True, "note": note}


def _excel_clear(host, doc, mode, args):
    """清空区域。args: ref + what。

    what="contents"（默认）只清值/公式，**保留格式与合并单元格**；
    what="formats"  只清格式，保留值；
    what="all"      值和格式一起清 —— 「把这片恢复成没动过的样子」用这个。

    与 set values=null 的区别：那个固定只走 ClearContents。
    为什么需要 all：实测在空区域上打一次 `fill color=None`，UsedRange 会被撑到那片区域
    （没填充也算"已格式化"），表看起来"不干净"；要真正还原就用 clear what="all"。
    """
    args = args or {}
    ws = _sheet_of(host, doc, args)
    ref = parse_ref(args.get("ref") or args.get("addr"))
    if ref["sheet"]:
        ws = host._com_call(doc, "Sheets", ref["sheet"])
    what = str(args.get("what") or "contents").strip().lower()
    method = {"contents": "ClearContents", "formats": "ClearFormats", "all": "Clear"}.get(what)
    if method is None:
        raise ValueError('clear 的 what 只能是 "contents" / "formats" / "all"，收到 %r' % (what,))
    r1, c1, r2, c2 = _target_box(ref, used_box(host, ws))
    # 合并区只有锚点存值：ClearContents/ClearFormats 对「合并区的一部分」会硬报
    # 「不能对合并单元格作部分修改」（2026-10-21 真机踩到，传单格 C19 而整块是 C19:F19）。
    # 清空的语义本来就是「把这格清干净」，所以命中合并区就把范围扩到整块，并如实回报扩到哪。
    requested = ref_text(r1, c1, r2, c2)
    merge_info = None
    scope = merge_scope(host, ws, r1, c1, r2, c2)
    if scope["merged"] and scope["areas"]:
        areas = scope["areas"]
        nr1 = min([r1] + [a["row"] for a in areas])
        nc1 = min([c1] + [a["col"] for a in areas])
        nr2 = max([r2] + [a["end_row"] for a in areas])
        nc2 = max([c2] + [a["end_col"] for a in areas])
        if (nr1, nc1, nr2, nc2) != (r1, c1, r2, c2):
            merge_info = {
                "requested": requested,
                "expanded_to": ref_text(nr1, nc1, nr2, nc2),
                "areas": [a["addr"] for a in areas],
                "note": "目标落在合并区 %s 内，清空范围已扩到整块 %s（合并区不能只清一部分）"
                        % ("、".join(a["addr"] for a in areas), ref_text(nr1, nc1, nr2, nc2)),
            }
            r1, c1, r2, c2 = nr1, nc1, nr2, nc2
    rng = _range_of(host, ws, r1, c1, r2, c2)
    host._com_call(rng, method)
    ident = _doc_ident(host, doc)
    snapshot_invalidate("excel", ident["path"] or ident["name"], host._safe_com_str(ws, "Name", ""))
    out = {"app": "excel", "sheet": host._safe_com_str(ws, "Name", ""),
           "cleared": ref_text(r1, c1, r2, c2), "what": what,
           "saved": False, "undoable": True}
    if merge_info:
        out["merge"] = merge_info
    return out


def _cell_text(value):
    """格子值 → 用于匹配的文本。数字 10.0 也要能匹配 "10"（别因为 .0 匹配不上）。"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _restore_number(original, text):
    """替换后把"数字变文本"掰回来：原值是 int/float，且新文本可解析成数字 → 按原类型返回。

    只认纯十进制字面量（含负号、小数点），"1,000" "1e3" "0x10" 一律不动——
    宁可保持字符串，也不要猜错把用户的文本改成数字。
    """
    if not isinstance(original, (int, float)) or isinstance(original, bool):
        return text
    s = str(text).strip()
    if not s:
        return text
    try:
        num = float(s)
    except ValueError:
        return text
    if "." not in s and "e" not in s and "E" not in s:
        # 长得像整数：原值是 int 或整值 float → 转回 int
        return int(num)
    return num


def _find_pattern(what, regex, kmode, case):
    """把 what 编译成一个正则（字面模式先 re.escape，所以查找和替换共用一条路径）。

    kmode 用锚点表达，避免为四种模式各写一套匹配/替换逻辑：
      exact → \\A...\\Z    starts → \\A...    ends → ...\\Z    contains → 不锚
    好处：字面/正则 × 查找/替换 四种组合走同一份代码，行为不会分叉。
    """
    body = str(what)
    body = "(?:%s)" % body if regex else re.escape(body)
    if kmode == "exact":
        body = r"\A" + body + r"\Z"
    elif kmode == "starts":
        body = r"\A" + body
    elif kmode == "ends":
        body = body + r"\Z"
    flags = 0 if case else re.IGNORECASE
    try:
        return re.compile(body, flags)
    except re.error as err:
        raise ValueError("正则写错了：%s（pattern=%r）。要按字面匹配就不要传 regex=true"
                         % (err, what))


def _scan_box(host, doc, args, require_in=False):
    """find / replace 的扫描区域：in 给了就用它（缺的一侧按数据区补齐），不给就用整张已用区。"""
    ws = _sheet_of(host, doc, args)
    box = used_box(host, ws)
    raw = args.get("in") or args.get("ref") or args.get("range")
    if raw:
        ref = parse_ref(raw)
        if ref["sheet"]:
            ws = host._com_call(doc, "Sheets", ref["sheet"])
            box = used_box(host, ws)
        r1, c1, r2, c2 = _spanned_box(host, ws, ref, box)
    elif require_in:
        raise ValueError(
            'replace 必须显式给 in（替换范围），例如 {"in":"G2:G16","what":"美白","to":"亮白"}。'
            '整表替换一旦错了很难挽回，所以这里不给默认范围；想先看会改哪些，用 action="find"。')
    else:
        b = box or {"first_row": 1, "first_col": 1, "last_row": 1, "last_col": 1}
        r1, c1, r2, c2 = b["first_row"], b["first_col"], b["last_row"], b["last_col"]
    cells = (r2 - r1 + 1) * (c2 - c1 + 1)
    if cells > READ_MAX_CELLS:
        raise ValueError("这次要扫 %d 格，超过单次上限 %d。请把 in 缩到某几列/某几行（如 "
                         '"G2:G200"）。' % (cells, READ_MAX_CELLS))
    return ws, r1, c1, r2, c2, box


def _excel_find(host, doc, mode, args):
    """查 what 在哪，**连带把命中格的数据一起带回来**。

    args: what（要找的内容）、in（区域，缺省=整张已用区）、sheet、
          mode=contains|exact|starts|ends（默认 contains）、regex=false、case=false、limit=50。

    值只在服务端读进来筛，**只有命中的才回给模型** —— 比"整列读进上下文自己筛"省得多。
    动作名后缀同样有效：find1 = 正则开、find0 = 强制字面；显式传 regex 优先。
    """
    args = args or {}
    what = args.get("what", args.get("value", args.get("pattern")))
    if what is None or str(what) == "":
        raise ValueError('find 需要 args.what（要找的内容），例如 {"what":"美白","in":"G:G"}')
    kmode = str(args.get("mode") or "contains").strip().lower()
    if kmode not in ("contains", "exact", "starts", "ends"):
        raise ValueError("find 的 mode 只能是 contains / exact / starts / ends，收到 %r" % (kmode,))
    regex = _bool_arg(args.get("regex"), False, "regex")
    case = _bool_arg(args.get("case"), False, "case")
    limit = args.get("limit")
    limit = int(limit) if limit not in (None, "") else 50

    ws, r1, c1, r2, c2, _box = _scan_box(host, doc, args)
    rx = _find_pattern(what, regex, kmode, case)
    grid = _read_grid(host, ws, r1, c1, r2, c2)

    matches, total = [], 0
    for i, row in enumerate(grid):
        for j, value in enumerate(row):
            text = _cell_text(value)
            if not text or rx.search(text) is None:
                continue
            total += 1
            if len(matches) < limit:
                matches.append({"cell": ref_text(r1 + i, c1 + j, r1 + i, c1 + j),
                                "row": r1 + i, "col": c1 + j, "value": value})
    return {"app": "excel", "sheet": host._safe_com_str(ws, "Name", ""),
            "in": ref_text(r1, c1, r2, c2), "what": what, "regex": regex, "mode": kmode,
            "case": case, "scanned": (r2 - r1 + 1) * (c2 - c1 + 1),
            "total": total, "returned": len(matches), "truncated": total > len(matches),
            "matches": matches}


def _excel_replace(host, doc, mode, args):
    """在 in 里把 what 换成 to（to 省略 = 删除）。一次调用做完，不需要人确认。

    args: in（**必填**）、what、to（可省=删除）、regex=false、case=false、
          mode=contains|exact|starts|ends、sheet。后缀 replace1 = 正则。

    三条安全设计（实测踩出来的，别改）：
      1. **in 必填**：整表误替换最难挽回，宁可要求明确写出范围（想先预览就用 find）。
      2. **跳过含公式的格**：只替换普通值。否则读出来的是计算值、写回去会把公式变成死值。
      3. 先整片一次赋值（快）；失败（例如范围含合并区）再**只逐格写真的变了的格**兜底。
         公式格与未变格的数量都如实回报。
    """
    args = args or {}
    what = args.get("what", args.get("value", args.get("pattern")))
    if what is None or str(what) == "":
        raise ValueError('replace 需要 args.what（要被替换的内容），例如 '
                         '{"in":"G2:G16","what":"美白","to":"亮白"}')
    new = args.get("to", args.get("with", args.get("replacement")))
    new = "" if new is None else str(new)
    kmode = str(args.get("mode") or "contains").strip().lower()
    if kmode not in ("contains", "exact", "starts", "ends"):
        raise ValueError("replace 的 mode 只能是 contains / exact / starts / ends，收到 %r" % (kmode,))
    regex = _bool_arg(args.get("regex"), False, "regex")
    case = _bool_arg(args.get("case"), False, "case")

    ws, r1, c1, r2, c2, _box = _scan_box(host, doc, args, require_in=True)
    grid = _read_grid(host, ws, r1, c1, r2, c2)
    fgrid = _read_grid_attr(host, ws, r1, c1, r2, c2, "Formula")   # 一次读回公式，用来认"公式格"
    rx = _find_pattern(what, regex, kmode, case)

    out, changes, skipped_formula, hit, changed = [], [], 0, 0, 0
    for i, row in enumerate(grid):
        orow = []
        for j, value in enumerate(row):
            formula = fgrid[i][j] if i < len(fgrid) and j < len(fgrid[i]) else None
            if isinstance(formula, str) and formula.startswith("="):
                skipped_formula += 1
                orow.append(formula)          # 原样写回公式文本，保住公式
                continue
            text = _cell_text(value)
            if not text or rx.search(text) is None:
                orow.append(value)
                continue
            hit += 1
            # 字面模式用函数替换：re.sub 会把替换串里的 \U \x \g 当反向引用，
            # Windows 路径（C:\Users\test）一替换就抛 bad escape。正则模式保留原样（用户可能故意用 \1）。
            after = rx.sub(lambda m: new, text) if not regex else rx.sub(new, text)
            if after == text:
                orow.append(value)             # 命中了，但替换后内容没变
                continue
            # 原值是数字且替换结果仍是数字 → 转回数字，别把 int/float 写成文本
            # （写成文本后 SUM/AVERAGE 会忽略它，还会标"以文本形式存储的数字"）
            orow.append(_restore_number(value, after))
            changed += 1
            if len(changes) < 20:
                changes.append({"cell": ref_text(r1 + i, c1 + j, r1 + i, c1 + j),
                                "from": text[:80], "to": after[:80]})
        out.append(orow)

    base = {"app": "excel", "sheet": host._safe_com_str(ws, "Name", ""),
            "in": ref_text(r1, c1, r2, c2), "what": what, "to": new, "regex": regex,
            "mode": kmode, "case": case, "scanned": (r2 - r1 + 1) * (c2 - c1 + 1),
            "hit": hit, "changed": changed, "skipped_formula": skipped_formula,
            "changes": changes, "changes_truncated": changed > len(changes),
            "saved": False, "undoable": True}
    if changed == 0:
        base["note"] = "没有格子需要改（命中 %d 个，但替换后内容不变）" % hit
        return base

    how = "grid"
    try:
        set_range_value(_range_of(host, ws, r1, c1, r2, c2), tuple(tuple(r) for r in out))
    except Exception:
        # 兜底：只逐格写真的变了的格（例如范围里含合并单元格时，整片赋值会被拒）
        how = "cells"
        for i, row in enumerate(grid):
            for j, value in enumerate(row):
                if out[i][j] is not value:
                    set_range_value(host._com_call(ws, "Cells", r1 + i, c1 + j), out[i][j])

    ident = _doc_ident(host, doc)
    snapshot_invalidate("excel", ident["path"] or ident["name"], host._safe_com_str(ws, "Name", ""))
    base["how"] = how
    base["note"] = "公式格已跳过（%d 个），它们的公式原文原样保留" % skipped_formula
    return base


def _axis_spec(text):
    """'B' / 'B:C' / '3' / '3:5' → (kind, first, last)；kind ∈ {"col","row"}。"""
    raw = str(text or "").strip().replace("$", "").replace(" ", "")
    if not raw:
        raise ValueError('move 需要 from：列写 "B" 或 "B:C"，行写 "3" 或 "3:5"')
    a, b = raw.split(":", 1) if ":" in raw else (raw, raw)
    if a.isalpha() and b.isalpha():
        return "col", col_to_index(a), col_to_index(b)
    if a.isdigit() and b.isdigit():
        return "row", int(a), int(b)
    raise ValueError('move 的地址只能写成 "B"/"B:C"（列）或 "3"/"3:5"（行），收到 %r' % (text,))


def _excel_move(host, doc, mode, args):
    """把 from 这段列/行**剪切后插到 to 的位置**（= 人手「剪切 → 插入剪切的单元格」）。

    实测定死语义（别改成 Cut(Destination)，那是**覆盖式**、会直接把目标列的数据毁掉）：
        Columns("B").Cut() + Columns("D").Insert(xlToRight)
        A B C D E  →  A C B D E      B 插到 **D 之前**，原列整体右移
    所以 to 的含义是「插到这里」，position 默认 before；要插到它后面传 position="after"。

    实现顺序刻意选成「先插空列 → 再把源剪进空列 → 最后删掉被剪空的源列」，
    这样任一步失败的最坏结果只是**多出一段空列/空行**，绝不会丢数据。

    args: from（"B"/"B:C" 或 "3"/"3:5"）、to（单个："D" 或 "8"）、position、sheet。
    返回 landed_at（源现在的起始列号/行号）与 affected（受影响的区间），方便核对。
    """
    args = args or {}
    ws = _sheet_of(host, doc, args)
    kind, s1, s2 = _axis_spec(args.get("from"))
    tkind, t1, t2 = _axis_spec(args.get("to"))
    if tkind != kind:
        raise ValueError('move 的 from 是%s、to 是%s，两者要同类型（列配列、行配行）'
                         % ("列" if kind == "col" else "行", "列" if tkind == "col" else "行"))
    if t1 != t2:
        raise ValueError('move 的 to 只能是一个位置（"D" 或 "8"），收到 %r；要移动多列请写在 from 里'
                         % (args.get("to"),))
    pos = str(args.get("position") or "before").strip().lower()
    if pos in ("after", "后", "之后", "后面"):
        pos = "after"
    elif pos in ("before", "前", "之前", "前面"):
        pos = "before"
    else:
        raise ValueError("move 的 position 只能是 before / after，收到 %r" % (args.get("position"),))

    src_first, src_last, target = min(s1, s2), max(s1, s2), t1
    if src_first <= target <= src_last:
        raise ValueError("move 的 to 落在 from 里面了（from=%s, to=%s），这样移动没有意义"
                         % (args.get("from"), args.get("to")))
    count = src_last - src_first + 1
    insert_ref = target + 1 if pos == "after" else target
    key = "Columns" if kind == "col" else "Rows"
    unit = col_letter if kind == "col" else str
    span = lambda a, b: ("%s:%s" % (unit(a), unit(b))) if a != b else unit(a)
    landed = insert_ref - count if src_last < insert_ref else insert_ref
    affected = span(min(src_first, insert_ref), max(src_last, insert_ref + count - 1))

    try:
        # ① 在插入点多开 count 个空列/空行（右边的整体后移）
        host._com_call(host._com_call(ws, key, span(insert_ref, insert_ref + count - 1)),
                       "Insert", -4161 if kind == "col" else -4121)
        # ② 源在插入点右边的要顺移 count 位
        sf = src_first + count if src_first >= insert_ref else src_first
        sl = src_last + count if src_last >= insert_ref else src_last
        host._com_call(host._com_call(ws, key, span(sf, sl)), "Cut",
                       host._com_call(ws, key, span(insert_ref, insert_ref + count - 1)))
        # ③ 删掉被剪空的源列/行
        host._com_call(host._com_call(ws, key, span(sf, sl)), "Delete")
    except Exception as err:
        raise ValueError(
            "移动%s失败：%s。后果与处理：若表里多出空白的列/行，说明只做完了前两步，删掉那些空列/行即可；"
            "若源列/行不见了，请人按 Ctrl+Z 撤销。常见原因是 WPS 弹了模态框或正处在单元格编辑态。"
            % ("列" if kind == "col" else "行", str(err)[:140]))

    ident = _doc_ident(host, doc)
    snapshot_invalidate("excel", ident["path"] or ident["name"], host._safe_com_str(ws, "Name", ""))
    return {"app": "excel", "sheet": host._safe_com_str(ws, "Name", ""),
            "moved": {"from": span(src_first, src_last), "to": unit(target),
                      "position": pos, "kind": "cols" if kind == "col" else "rows",
                      "count": count, "inserted_at": unit(insert_ref)},
            "landed_at": unit(landed), "affected": affected,
            "how": "insert-then-cut（先插空列/行，最坏只多一段空的，不会丢数据）",
            "saved": False, "undoable": True,
            "note": "人的活动表没被切换；合并、格式、公式都跟着走（Excel 原生的剪切语义）"}


def _excel_brush(host, doc, mode, args):
    """格式刷：把 args.from 的**格式**刷到 args.to（只带格式，不带值与公式）。

    实测要点：必须走 Copy + PasteSpecial(xlPasteFormats=-4122)，逐格抄 Interior/Font
    会丢掉边框、数字格式这些；副作用是**会占用系统剪贴板**（等于替人按了 Ctrl+C）。
    """
    args = args or {}
    src = parse_ref(args.get("from") or args.get("src"))
    dst = parse_ref(args.get("to") or args.get("dst"))
    base = args.get("sheet")
    # 源/目各自取表：src 带表前缀就用它，否则用 args.sheet，再否则跟随活动表。
    # （曾把 src 表重定向后**两个 Range 都用同一个 ws**，导致跨表时格式刷回了源表、目标表纹丝不动。）
    from_name = src["sheet"] or args.get("from_sheet") or base
    to_name = dst["sheet"] or args.get("to_sheet") or base
    src_ws = host._com_call(doc, "Sheets", from_name) if from_name else _sheet_of(host, doc, args)
    dst_ws = src_ws if to_name == from_name else (
        host._com_call(doc, "Sheets", to_name) if to_name else _sheet_of(host, doc, args))
    box = used_box(host, src_ws)
    sr1, sc1, sr2, sc2 = _target_box(src, box)
    # 目标坐标若缺省，按**目标表**的数据区补齐（源表的 box 只用于源坐标）。
    dst_box = box if dst_ws is src_ws else used_box(host, dst_ws)
    dr1, dc1, dr2, dc2 = _target_box(dst, dst_box)
    s_rng = _range_of(host, src_ws, sr1, sc1, sr2, sc2)
    d_rng = _range_of(host, dst_ws, dr1, dc1, dr2, dc2)
    _com_call_disp(host, s_rng, "Copy", 0x0000022B)
    try:
        _com_call_disp(host, d_rng, "PasteSpecial", 0x00000788, -4122)
    finally:
        try:
            app = host._get_wps_app("excel")[0]
            _set_prop(app, "CutCopyMode", False)     # 收掉蚂蚁线
        except Exception:
            pass
    src_txt, dst_txt = ref_text(sr1, sc1, sr2, sc2), ref_text(dr1, dc1, dr2, dc2)
    src_sheet, dst_sheet = host._safe_com_str(src_ws, "Name", ""), host._safe_com_str(dst_ws, "Name", "")
    ident = _doc_ident(host, doc)
    snapshot_invalidate("excel", ident["path"] or ident["name"], dst_sheet)
    if src_sheet != dst_sheet:
        snapshot_invalidate("excel", ident["path"] or ident["name"], src_sheet)
    return {"app": "excel", "sheet": dst_sheet, "from_sheet": src_sheet, "to_sheet": dst_sheet,
            "from": src_txt, "to": dst_txt,
            "note": "只刷了格式（不含值与公式）；过程中占用过系统剪贴板"}


def _com_call_disp(host, obj, name, dispid, *cargs):
    """按名字调用，失败用 DISPID 直调（PasteSpecial/Copy 在个别绑定下名字查不到）。"""
    try:
        return host._com_call(obj, name, *cargs)
    except Exception:
        inner = getattr(obj, "_oleobj_", None) or getattr(obj, "_comobj", None)
        if inner is None:
            raise
        return inner._invoke(dispid, 1, 0, *cargs)


def insert_split_risk(host, ws, kind, at):
    """插入行/列时，插入点是否落在某个合并区**内部**（= 会把这块合并拆开）。

    只扫插入位置那一行/列与已用区的交集，成本是列数（或行数）量级，不扫整表。
    判定依据：合并区起点严格在插入点之前（row < at ≤ end_row）就会被拆；
    起点正好等于插入点的合并区只是整体后移，不算拆。

    为什么要在写之前判：WPS 在「插入会拆开合并单元格」时的行为未实测，
    最坏情况是弹模态确认框 —— 那一个 COM 调用会挂死，整个 stdio 服务跟着假死。
    宁可默认拦住让人确认，也不赌。
    """
    box = used_box(host, ws)
    if not box:
        return []
    if kind == "rows":
        scope = merge_scope(host, ws, at, box["first_col"], at, box["last_col"])
        key = "row"
    else:
        scope = merge_scope(host, ws, box["first_row"], at, box["last_row"], at)
        key = "col"
    if scope.get("scanned") is False:
        return [{"addr": scope.get("range"), "note": "合并区过多未逐一确认，按风险处理"}]
    risk = []
    for a in scope.get("areas") or []:
        if a.get(key) is not None and a[key] < at:
            risk.append({"addr": a["addr"], "anchor": a["anchor"],
                         "rows": a["rows"], "cols": a["cols"]})
    return risk


def _excel_insert(host, doc, mode, args):
    """插入整行或整列。args: rows="5:6" 或 cols="C:D"（二选一），可带 force=true。

    安全闸：插入点若落在合并区内部，默认拒绝（见 insert_split_risk）。
    明确要拆就带 force=true —— 此时照插，并把被拆的合并区如实回报。
    """
    args = args or {}
    ws = _sheet_of(host, doc, args)
    spec_rows = str(args.get("rows") or "").strip()
    spec_cols = str(args.get("cols") or "").strip()
    if bool(spec_rows) == bool(spec_cols):
        raise ValueError('insert 需要在 rows（如 "5:6"）或 cols（如 "C:D"）里二选一')
    force = _bool_arg(args.get("force"), False, "force")
    if spec_rows:
        ref = parse_ref(spec_rows)
        if ref["r1"] is None:
            raise ValueError("rows 要写成行区间，如 \"5:6\"")
        r1, r2 = sorted((ref["r1"], ref["r2"]))
        kind, at = "rows", r1
        target = "%d:%d" % (r1, r2)
    else:
        ref = parse_ref(spec_cols)
        if ref["c1"] is None:
            raise ValueError("cols 要写成列区间，如 \"C:D\"")
        c1, c2 = sorted((ref["c1"], ref["c2"]))
        kind, at = "cols", c1
        target = "%s:%s" % (col_letter(c1), col_letter(c2))

    risk = insert_split_risk(host, ws, kind, at)
    if risk and not force:
        raise ValueError(
            "要插的 %s 落在合并区内部，插入会把这块合并拆开：%s。\n"
            "WPS 在这种情况下可能弹「这会移除合并单元格」的模态框，那个框会把这一条命令挂死、"
            "整个服务跟着假死，所以这里默认**拒绝执行**。\n"
            "确认要拆就带 force=true 再调一次；或者先 unmerge 相关区域、插完再 merge 回去。"
            % (target, "、".join(x["addr"] or "?" for x in risk)))

    if kind == "rows":
        host._com_call(host._com_call(ws, "Rows", target), "Insert")
    else:
        host._com_call(host._com_call(ws, "Columns", target), "Insert")
    ident = _doc_ident(host, doc)
    snapshot_invalidate("excel", ident["path"] or ident["name"], host._safe_com_str(ws, "Name", ""))
    out = {"app": "excel", "sheet": host._safe_com_str(ws, "Name", ""),
           "inserted": target, "kind": kind}
    if risk:
        out["merge_split"] = risk
        out["merge_split_note"] = "已按 force=true 执行，上面这些合并区可能已被拆开，请 get 确认后再 merge 回来"
    return out


def _excel_delete(host, doc, mode, args):
    """删除整行或整列。args: rows="5:6" 或 cols="C:D"。

    安全闸：被删范围里**只要碰到合并区就拒绝** —— 那种情况下 Excel/WPS 会弹
    「这会移除合并单元格」的模态确认框，而模态框会把这一个 COM 调用**挂死**（然后整个服务假死）。
    宁可让人先手工拆掉，也不冒挂死的风险。
    """
    args = args or {}
    ws = _sheet_of(host, doc, args)
    spec_rows = str(args.get("rows") or "").strip()
    spec_cols = str(args.get("cols") or "").strip()
    if bool(spec_rows) == bool(spec_cols):
        raise ValueError('delete 需要在 rows（如 "5:6"）或 cols（如 "C:D"）里二选一')

    if spec_rows:
        ref = parse_ref(spec_rows)
        if ref["r1"] is None:
            raise ValueError("rows 要写成行区间，如 \"5:6\"")
        r1, r2 = sorted((ref["r1"], ref["r2"]))
        c1, c2 = 1, None
        target, kind = "%d:%d" % (r1, r2), "rows"
    else:
        ref = parse_ref(spec_cols)
        if ref["c1"] is None:
            raise ValueError("cols 要写成列区间，如 \"C:D\"")
        c1, c2 = sorted((ref["c1"], ref["c2"]))
        r1, r2 = 1, None
        target = "%s:%s" % (col_letter(c1), col_letter(c2))
        kind = "cols"

    box = used_box(host, ws)
    if box:
        if kind == "rows":
            probe = _range_of(host, ws, r1, box["first_col"], r2, box["last_col"])
        else:
            probe = _range_of(host, ws, box["first_row"], c1, box["last_row"], c2)
        ok, merged = host._com_try_attr(probe, "MergeCells")
        if ok and merged is not False:
            raise ValueError(
                "要删的 %s 里碰到合并单元格了。这种情况 WPS/Excel 会弹模态确认框，"
                "会把这一个 COM 调用挂死（然后整个服务假死），所以这里**拒绝执行**。\n"
                "请先 unmerge 相关区域，或让人手工确认删除。" % target)

    if kind == "rows":
        host._com_call(host._com_call(ws, "Rows", target), "Delete")
    else:
        host._com_call(host._com_call(ws, "Columns", target), "Delete")
    ident = _doc_ident(host, doc)
    snapshot_invalidate("excel", ident["path"] or ident["name"], host._safe_com_str(ws, "Name", ""))
    return {"app": "excel", "sheet": host._safe_com_str(ws, "Name", ""),
            "deleted": target, "kind": kind}


def _excel_activate(host, doc, mode, args):
    """切换（激活）工作表。等于 v2 的 activate_sheet 工具，薄壳版只传表名/序号。

    这是**会改变人所见界面**的动作（人的当前表会跟着切），所以列入写类：
    编辑态下同样拿不到 Workbook 级对象，要走「等松手 → 必要时 ESC」那条路。
    attached 时只 Activate，绝不 Save / Close / Quit，绝不新建工作表。
    """
    args = args or {}
    name = args.get("sheet") or args.get("name") or args.get("to")
    idx = args.get("index") or args.get("sheet_index")
    if not name and not idx:
        raise ValueError('activate 需要 args.sheet（表名）或 args.index（第几张，从 1 开始）')
    before = host._safe_sheet_name(doc)
    ws = host._com_call(doc, "Sheets", int(idx) if idx else name)
    target = host._safe_com_str(ws, "Name", str(name or idx))
    try:
        host._com_call(ws, "Activate")
    except Exception:
        host._com_call(ws, "Select")   # 个别环境对 Sheet.Activate 名称解析不友好，退化为 Select
    after = host._safe_sheet_name(doc)
    return {
        "app": "excel",
        "activated": target,
        "active_sheet_before": before,
        "active_sheet_after": after,
        "switched": bool(after and after != before),
        "note": "已切到该表——人眼前看到的工作表也跟着变了",
    }


def _excel_save(host, doc, mode, args):
    doc.Save()
    return {"app": "excel", "saved": True,
            "doc": _doc_ident(host, doc),
            "note": "已就地保存（等价于人按 Ctrl+S）"}


# ==================== Word / 演示 动作（薄壳化转发） ====================


def _word_where(host, doc, mode, args):
    try:
        text = doc.Content.Text
    except Exception:
        text = ""
    paras = None
    try:
        paras = int(doc.Paragraphs.Count)
    except Exception:
        pass
    out = {"app": "word", "doc": _doc_ident(host, doc),
           "chars": len(text or ""),
           "head": (text or "")[:200]}
    if paras is not None:
        out["paragraphs"] = paras
    return out


def _word_get(host, doc, mode, args):
    try:
        text = doc.Content.Text
    except Exception as exc:
        raise ValueError("读不到 Word 正文：%s" % (exc,))
    max_chars = (args or {}).get("max_chars")
    if max_chars:
        try:
            max_chars = int(max_chars)
        except Exception:
            max_chars = None
    out = {"app": "word", "chars": len(text or ""), "text": text}
    if max_chars and text and len(text) > max_chars:
        out["text"] = text[:max_chars]
        out["truncated"] = True
    return out


def _word_set(host, doc, mode, args):
    args = args or {}
    content = args.get("content")
    if content is None:
        raise ValueError("word set 需要 args.content（要写入的正文文本）")
    try:
        # 用 _bool_arg：模型可能传字符串 "false"/"0"，裸真值判断会把它们当真是，
        # 于是"本想整篇替换"变成"末尾追加"，方向弄反很危险。
        if _bool_arg(args.get("append"), False, "append"):
            doc.Content.InsertAfter(str(content))
        else:
            doc.Content.Text = str(content)
    except Exception as exc:
        raise ValueError("写 Word 正文失败：%s" % (exc,))
    saved = None
    if AUTOSAVE:
        try:
            doc.Save()
            saved = True
        except Exception:
            saved = False
    return {"app": "word", "saved": saved if AUTOSAVE else False, "undoable": not AUTOSAVE,
            "chars": len(str(content))}


def _word_save(host, doc, mode, args):
    doc.Save()
    return {"app": "word", "saved": True, "doc": _doc_ident(host, doc),
            "note": "已就地保存（等价于人按 Ctrl+S）"}


def _ppt_where(host, doc, mode, args):
    out = {"app": "ppt", "doc": _doc_ident(host, doc)}
    try:
        out["slides"] = int(doc.Slides.Count)
    except Exception:
        pass
    return out


def _ppt_get(host, doc, mode, args):
    slides = []
    try:
        cnt = int(doc.Slides.Count)
    except Exception:
        cnt = 0
    for i in range(1, cnt + 1):
        texts = []
        try:
            slide = doc.Slides(i)
            for shape in slide.Shapes:
                try:
                    if shape.HasTextFrame:
                        texts.append(shape.TextFrame.TextRange.Text)
                except Exception:
                    continue
        except Exception:
            continue
        slides.append({"index": i, "texts": texts})
    return {"app": "ppt", "count": len(slides), "slides": slides}


def _ppt_save(host, doc, mode, args):
    doc.Save()
    return {"app": "ppt", "saved": True, "doc": _doc_ident(host, doc)}


# ==================== legacy 逃生口 ====================

_LEGACY_WRITE = ("write_document", "edit_cell", "add_rows", "add_columns", "merge_cells",
                 "unmerge_cells", "set_column_width", "set_row_height", "insert_image",
                 "set_cell_style", "set_paragraph_style", "batch_write", "batch_set_style",
                 "apply_template", "activate_sheet")


def legacy_names(host):
    """老工具名清单（只报名字，不报 schema——这是分层加载省 token 的关键）。"""
    names = []
    for item in (getattr(host, "TOOLS", None) or []):
        if isinstance(item, dict) and item.get("name"):
            names.append(item["name"])
    return names


# ==================== 分层帮助文本 ====================

HELP = {
    "excel": """wps_do(app="excel", action=..., args={...}) —— 操作人当前打开的活表格（内存实时值）

【读】动作            参数                                 说明
-----               ----                                 ----
where               sheet?, key_col?                    活动表状态：表名/已用区/数据列跨度/下一空行/合并区
get                 ref, sheet?, max_rows?              读 ref 指向的值
image               ref, sheet?, dir?, download?,       看图：把图落成文件并返回路径+尺寸，图片不过上下文
                    allow_clipboard?                    自动分辨 DISPIMG 嵌入图 / 链接图 / 浮动图

【写】动作            参数                                 说明
set                 ref, values, sheet?                 写 ref（默认不保存，可 Ctrl+Z）
fill                ref, color                          给区域填背景色（color=null 清除填充）
style               ref, bold/italic/size/name/color/  字体与对齐、数字格式
                    underline/align/align_v/format
width               col, w                              列宽（col="A" 或 "A:C"）
height              row, h                              行高（row="3" 或 "3:6"）
merge               ref                                 合并区域（自带先拆后合，可扩大/缩小）
unmerge             ref                                 拆分合并区
brush               from, to                            格式刷（只刷格式，不带值/公式）
insert              rows="5:6" 或 cols="C:D"            插入整行/整列（落在合并区内部默认拒绝，可加 force=true）
delete              rows="5:6" 或 cols="C:D"            删除整行/整列（碰合并区会拒绝）
clear               ref, what="contents"                 清空区域：contents 只清内容（留格式与合并）
                                                          formats 只清格式  all 内容和格式一起清
copy                from, to, with_format=false          搬运区域：默认只搬值；true 连格式/公式/合并一起搬
                                                          from/to 可写 "Sheet2!A1:C3" 跨表；值不经模型上下文
move                from, to, position="before"          把 from 这段列/行剪切后插到 to 的位置
                                                          from="B"/"B:C"/"3"/"3:5"，to="D"/"8"；见下方例子
activate            sheet 或 index                      切换工作表（人眼前看到的表会跟着切）
save                -                                   就地保存（等于替人按 Ctrl+S）

【查】动作            参数                                 说明
find                what, in?, mode?, regex?, limit?     查找：只回命中格的地址 + 数据（值不经模型上下文）
                                                          in 缺省=整张已用区；mode=contains/exact/starts/ends
replace             in, what, to?                       查找替换（一次做完，不需要人工确认）
                                                          in **必填**（不设默认范围，防整表误替换）
                                                          含公式的格会自动跳过并如实回报

查/替换的公共参数
  regex=false       true 则 what 按**正则**解释（默认字面）
  case=false        默认不区分大小写（匹配 Excel 的默认行为）
  mode              contains（默认，包含）| exact（整格相等）| starts（开头）| ends（结尾）

动作名后缀 0/1（省字写法，显式参数优先）
  copy0 = copy + with_format=false      copy1 = copy + with_format=true
  find0 = find + regex=false            find1 = find + regex=true
  replace0 = replace + regex=false      replace1 = replace + regex=true
  例：action="replace1" 就是"正则替换"；action="copy1" 就是"连格式一起复制"

ref 写法（与 get/set 完全一致）
  "B5"        单格        "A2:D10"   区域
  "3:3"       整行        "C:C"      整列      "Sheet2!A1"  指定表

颜色                        "color": "#RRGGBB"（如 "#FF0000" 红、"#D9D9D9" 浅灰）
数字格式                    "format": "0.00" / "0%" / "yyyy-mm-dd" / "#,##0"

例子
  wps_do(app="excel", action="where")
  wps_do(app="excel", action="get",  args={"ref":"H2"})
  wps_do(app="excel", action="set",  args={"ref":"3:3", "values":["a","b","c"]})
  wps_do(app="excel", action="set",  args={"ref":"A5:C6","values":[[1,2,3],[4,5,6]]})
  wps_do(app="excel", action="fill", args={"ref":"C:C", "color":"#D9D9D9"})
  wps_do(app="excel", action="style",args={"ref":"A1:C1","bold":true,"size":14,"align":"center"})
  wps_do(app="excel", action="width",args={"col":"A:C","w":20})
  wps_do(app="excel", action="merge",args={"ref":"A1:C1"})
  wps_do(app="excel", action="brush",args={"from":"A1","to":"C3:F9"})
  wps_do(app="excel", action="copy", args={"from":"C3","to":"B2"})
  wps_do(app="excel", action="copy", args={"from":"A1:C10","to":"G1","with_format":true})
  wps_do(app="excel", action="copy", args={"from":"Sheet1!A1:C10","to":"Sheet2!A1"})
  wps_do(app="excel", action="move", args={"from":"B", "to":"D"})            # B 列挪到 D 列前面
  wps_do(app="excel", action="move", args={"from":"B", "to":"D", "position":"after"})
  wps_do(app="excel", action="move", args={"from":"5:6", "to":"2"})          # 5~6 行挪到第 2 行前
  wps_do(app="excel", action="find", args={"what":"美白","in":"G:G"})
  wps_do(app="excel", action="find1", args={"what":"\\d+位","in":"G:G"})     # 正则查找
  wps_do(app="excel", action="replace", args={"in":"G2:G16","what":"美白","to":"提亮"})
  wps_do(app="excel", action="replace1", args={"in":"G2:G16",
          "what":"(\\d+)\\s*ml","to":"\\1ml"})                              # 正则替换（后向引用）

规矩
  默认操作「人正盯着的那张表」；不传 path 时若人开着多本表会报错并列出候选，
  此时请人指认或补 path。所有写动作默认**不保存**，请人自己 Ctrl+S。

合并单元格（写前必检，get 直接回报）
  get 的返回体自带 merge：
    {"range":"C3","merged":true,"covers":true,"partial":false,
     "areas":[{"addr":"C2:F4","anchor":"C2","rows":3,"cols":4}]}
    merged=false 就是普通格；covers=true 表示这整片就是一个合并区；partial=true 表示
    请求范围只盖住某个合并区的一部分。
  set 写入前会先做同样的检查，三种处理：
    · 目标正好是合并区 → 正常写，值落在锚点，返回 merge.area/rows/cols 说明整块多大
    · 单格在合并区内但不在锚点 → 自动改写到锚点，返回 anchored_from
    · 请求范围只覆盖合并区的一部分 → **直接报错拒绝**（会丢值），请整块写或先 unmerge
  clear 命中合并区时会把范围自动扩到整块（合并区不能只清一部分，否则 WPS 硬报错），
  返回 merge.expanded_to 告诉你实际清到了哪。
  insert 的插入点落在合并区**内部**时默认拒绝（会把这块合并拆开，可能弹模态框挂死服务）：
  确认要拆就带 force=true，返回 merge_split 告诉你拆到了哪些块。
  想先确认某格多大、是不是合并区，直接 get 它，不要靠猜。

看图（action="image"）
  表里的一张「图片」其实有三种，取法不同，工具会自动分辨并在 source 里说明：
    ① WPS 嵌入单元格图片：格子是 =DISPIMG("ID_xxx",1)。
       文档存过盘 → 从 xlsx 包的 xl/cellimages.xml → rels → xl/media/* 抠**原图**（全分辨率），
       source=xlsx_media；刚插入还没存盘 → 退成剪贴板截屏，source=clipboard
       （只有屏幕分辨率、大图会被单元格裁切，并且**会覆盖人的剪贴板**）。
       这类图不是 Shape，Shapes 里数不到，Range 也没有取字节的属性 —— 别再试别的路。
    ② 链接图：=IMAGE("https://…") 或格子里直接是链接 → source=url，download=true 时已落一份。
    ③ 浮动图片（Ctrl+V 贴的那种，属于 Shape）→ 从 Shape 复制，source=clipboard。
  返回体给的是**文件路径 + 像素尺寸 + 来源**，图片本身不进模型上下文（避免几十 KB base64 烧上下文）。
    {"count":1,"images":[{"cell":"C12","kind":"wps_cellimage","id":"ID_…",
      "source":"xlsx_media","path":"C:\\…\\wps_mcp_images\\Sheet1_C12.png",
      "bytes":38836,"width":759,"height":188,"format":"png"}]}
  想换落盘位置传 dir；不想动剪贴板传 allow_clipboard=false（那就只认包里有的原图）。
  一次最多扫 400 格。
""",
    "word": """wps_do(app="word", action=..., args={...}) —— 操作人当前打开的 Word 活文档

动作            参数                    说明
-----           ----                    ----
where           -                       文档名/是否未保存/字数/段落数/开头 200 字
get             max_chars?              读全文（可用 max_chars 截断）
set             content, append?        写正文：append=true 追加，否则整篇替换
save            -                       就地保存

例子
  wps_do(app="word", action="where")
  wps_do(app="word", action="get", args={"max_chars":2000})
  wps_do(app="word", action="set", args={"content":"正文……","append":True})
""",
    "ppt": """wps_do(app="ppt", action=..., args={...}) —— 操作人当前打开的演示文稿

动作            参数      说明
-----           ----      ----
where           -         文档名/是否未保存/幻灯片页数
get             -         逐页列出文本框文字
save            -         就地保存

（需要写 PPT 时用 action="legacy"，透传老工具，见 help topic=legacy）
""",
}

LEGACY_TOPIC = """legacy 逃生口：透传任意老工具（schema 不再常驻，按需即用）

  wps_do(app="excel", action="legacy",
         args={"tool": "edit_cell", "args": {"file_path": "...", "row": 3, "column": 10,
                                            "value": "x", "mode": "com"}})

可用老工具名（NAMES_HERE）
注意：老工具默认 mode="file"（读的是磁盘旧文件），操作活文档必须显式传 "mode": "com"；
      老 edit_cell 每写一次都会 wb.Save()（会吃掉人的 Ctrl+Z），能用 action="set" 就别用它。
"""


# ==================== 看图：DISPIMG 嵌入图 / 链接图 / 浮动图 ====================
#
# 一张「表里的图片」可能是三种完全不同的东西，取法也不同：
#
#   ① WPS 嵌入单元格图片 —— 单元格里是一条 =DISPIMG("ID_xxx",1) 公式。
#      文件里它是 xl/cellimages.xml 的 <etc:cellImage>：<xdr:cNvPr name="ID_xxx">
#      配 <a:blip r:embed="rIdN">，再由 xl/_rels/cellimages.xml.rels 指到 xl/media/imageN.png。
#      → 文档存过盘就能从包里抠出**原图**（全分辨率）；没存盘（刚插入）包里没有，
#        退回剪贴板截屏（只有屏幕分辨率，且按单元格裁切）。
#      注意：这类图**不是 Shape**，ws.Shapes 里数不到，Range 也没有取字节的属性。
#
#   ② 链接图 —— =IMAGE("https://…") 或者单元格里直接写着图片链接。
#      → 直接访问链接即可（download=true 时顺手落一份到临时目录）。
#
#   ③ 浮动图片 —— Ctrl+V 贴进来的那种，是 Shape，可以 Copy 到剪贴板再抓。
#
# 落盘而不是塞 base64：图片动辄几十上百 KB，塞进返回值会烧掉模型的上下文。
# 只回路径 + 尺寸 + 来源，要看的自己去读那个文件。

_DISPIMG_RE = re.compile(r'DISPIMG\s*\(\s*"([^"]+)"', re.I)
_IMAGE_FN_RE = re.compile(r'^\s*=\s*IMAGE\s*\(\s*"([^"]+)"', re.I)
_URL_RE = re.compile(r'https?://[^\s"\'<>]+', re.I)
IMAGE_MAX_CELLS = 400
_OOXML_NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def image_out_dir(explicit=None):
    """图片落盘目录：显式 > 环境变量 WPS_MCP_IMAGE_DIR > %TEMP%/wps_mcp_images。"""
    import tempfile
    where = explicit or os.environ.get("WPS_MCP_IMAGE_DIR") or os.path.join(
        tempfile.gettempdir(), "wps_mcp_images")
    if not os.path.isdir(where):
        try:
            os.makedirs(where)
        except OSError:
            pass
    return where


def _safe_stem(text):
    return re.sub(r"[^0-9A-Za-z_\-]+", "_", str(text))[:60] or "img"


def _image_size(path):
    """读图片像素尺寸；没装 Pillow 就返回 None（不影响取图）。"""
    try:
        from PIL import Image
        with Image.open(path) as im:
            return [im.size[0], im.size[1]]
    except Exception:
        return None


def cellimages_index(path):
    """从磁盘上的 xlsx 建 {DISPIMG ID: zip 内 media 路径}。

    返回 (map, note)。note 非空说明为什么建不出来（打不开、不是 WPS 存的等）。
    只读 zip，不碰文件本身。
    """
    import zipfile
    import xml.etree.ElementTree as ET
    out = {}
    if not path or not os.path.isfile(path):
        return out, "文档没有磁盘路径（可能从未保存）"
    try:
        zf = zipfile.ZipFile(path)
    except Exception as exc:
        return out, "打不开 %s（%s）" % (path, exc)
    try:
        names = set(zf.namelist())
        if "xl/cellimages.xml" not in names:
            return out, "包里没有 xl/cellimages.xml（不是 WPS 存的、或这张表里没有嵌入图）"
        rels = {}
        if "xl/_rels/cellimages.xml.rels" in names:
            root = ET.fromstring(zf.read("xl/_rels/cellimages.xml.rels"))
            for rel in root.findall("rel:Relationship", _OOXML_NS):
                rels[rel.get("Id")] = rel.get("Target")
        root = ET.fromstring(zf.read("xl/cellimages.xml"))
        for pic in root.iter("{%s}pic" % _OOXML_NS["xdr"]):
            cnp = pic.find(".//xdr:nvPicPr/xdr:cNvPr", _OOXML_NS)
            blip = pic.find(".//a:blip", _OOXML_NS)
            if cnp is None or blip is None:
                continue
            ident = (cnp.get("name") or "").strip()
            rid = blip.get("{%s}embed" % _OOXML_NS["r"])
            target = rels.get(rid)
            if not ident or not target:
                continue
            member = "xl/" + target.lstrip("/")
            if member in names:
                out[ident] = member
        if not out:
            return out, "xl/cellimages.xml 里没解析出图片（格式可能变了）"
        return out, None
    except Exception as exc:
        return out, "解析 cellimages 失败（%s）" % exc
    finally:
        zf.close()


def _extract_cellimage(doc_path, member, out_dir, stem):
    """把包里那一份 media 原样写出来（不重编码，保住原分辨率与原格式）。"""
    import zipfile
    ext = os.path.splitext(member)[1].lower() or ".png"
    dst = os.path.join(out_dir, "%s%s" % (stem, ext))
    with zipfile.ZipFile(doc_path) as zf:
        data = zf.read(member)
    with open(dst, "wb") as fh:
        fh.write(data)
    return dst, len(data), ext.lstrip(".")


def _grab_via_clipboard(host, target, dst, is_shape=False):
    """把区域/形状当位图复制到剪贴板再抓回来。

    只在「文档没存盘、或包里查不到这张图」时兜底。精度只有屏幕分辨率，
    且按单元格裁切（大图会被切掉）—— 返回体里會如实标 source=clipboard。
    副作用：覆盖人的剪贴板。
    """
    try:
        from PIL import ImageGrab
    except Exception:
        return None, None, "没装 Pillow，走不了剪贴板兜底（pip install pillow）"
    if is_shape:
        host._com_call(target, "Copy")
    else:
        host._com_call(target, "CopyPicture", 1, 2)   # xlScreen=1, xlBitmap=2
    time.sleep(0.35)
    try:
        got = ImageGrab.grabclipboard()
    except Exception as exc:
        return None, None, "读剪贴板失败（%s）" % exc
    if got is None:
        return None, None, "剪贴板里没有位图（WPS 可能没响应这次复制）"
    if isinstance(got, list):
        return None, None, "剪贴板里是文件列表，不是位图"
    got.save(dst)
    return dst, list(got.size), None


def _download_image(url, out_dir, stem):
    """直接访问链接把图落一份（链接图的取法就是「去访问链接」）。"""
    import urllib.request
    req = urllib.request.Request(
        url, headers={"User-Agent": "Mozilla/5.0 (WPS MCP image reader)"})
    with urllib.request.urlopen(req, timeout=20) as resp:
        data = resp.read()
        ctype = (resp.headers.get("Content-Type") or "").lower()
    ext = ".png"
    low = url.lower().split("?")[0]
    for cand in (".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp", ".ico"):
        if low.endswith(cand):
            ext = cand
            break
    else:
        for key, cand in (("png", ".png"), ("jpeg", ".jpg"), ("jpg", ".jpg"),
                          ("webp", ".webp"), ("gif", ".gif"), ("bmp", ".bmp"),
                          ("icon", ".ico")):
            if key in ctype:
                ext = cand
                break
    dst = os.path.join(out_dir, "%s%s" % (stem, ext))
    with open(dst, "wb") as fh:
        fh.write(data)
    return dst, len(data), ext.lstrip(".")


def _shapes_in_box(host, ws, r1, c1, r2, c2):
    """找出左上角落在目标范围里的浮动图片（Shape）。"""
    out = []
    ok, shapes = host._com_try_attr(ws, "Shapes")
    if not ok or shapes is None:
        return out
    cnt = host._com_try_attr(shapes, "Count")[1]
    try:
        cnt = int(host._com_to_plain(cnt) or 0)
    except Exception:
        return out
    for i in range(1, cnt + 1):
        shp = host._com_call(shapes, "Item", i)
        ok2, tl = host._com_try_attr(shp, "TopLeftCell")
        if not ok2 or tl is None:
            continue
        ar = host._com_try_attr(tl, "Row")[1]
        ac = host._com_try_attr(tl, "Column")[1]
        try:
            ar, ac = int(host._com_to_plain(ar)), int(host._com_to_plain(ac))
        except Exception:
            continue
        if r1 <= ar <= r2 and c1 <= ac <= c2:
            out.append({"cell": ref_text(ar, ac, ar, ac), "kind": "shape", "shape_index": i,
                        "shape_name": host._safe_com_str(shp, "Name", "")})
    return out


def _whole_merge_box(host, ws, r1, c1, r2, c2):
    """把范围扩到它盖住的合并区整块。

    剪贴板兜底时要用：只 CopyPicture 单格会把跨格图片裁成那一格那么大。
    """
    scope = merge_scope(host, ws, r1, c1, r2, c2)
    areas = scope.get("areas") or []
    if not areas:
        return r1, c1, r2, c2
    a = areas[0]
    return (min(r1, a["row"]), min(c1, a["col"]),
            max(r2, a["end_row"]), max(c2, a["end_col"]))


def _excel_image(host, doc, mode, args):
    """看图：把 ref 里的图片取出来落成文件，返回路径（图片数据不经模型上下文）。

    args: ref（缺省单格；可给区域一次扫多格）, dir?, download?, allow_clipboard?
    三种来源自动判断，返回体里每条图都带 source：
      xlsx_media —— 从存盘文件的 xl/media 抠出的**原图**（WPS 嵌入单元格图片，全分辨率）
      clipboard  —— 剪贴板截屏兜底（文档没存盘 / 包里查不到这张图；屏幕分辨率且会被裁切）
      url        —— 单元格里是图片链接（=IMAGE(...) 或裸链接），download=true 时已落一份
      shape      —— 浮动图片，从 Shape 复制
    """
    args = args or {}
    ws = _sheet_of(host, doc, args)
    ref = parse_ref(args.get("ref") or args.get("addr") or "A1")
    if ref["sheet"]:
        ws = host._com_call(doc, "Sheets", ref["sheet"])
    box = used_box(host, ws) or {"first_row": 1, "first_col": 1, "last_row": 1, "last_col": 1}
    r1 = ref["r1"] if ref["r1"] is not None else box["first_row"]
    r2 = ref["r2"] if ref["r2"] is not None else r1
    c1 = ref["c1"] if ref["c1"] is not None else box["first_col"]
    c2 = ref["c2"] if ref["c2"] is not None else c1
    if (r2 - r1 + 1) * (c2 - c1 + 1) > IMAGE_MAX_CELLS:
        raise ValueError("ref 太大（%d 格）。看图请一格格来，或把范围缩到 %d 格以内。"
                         % ((r2 - r1 + 1) * (c2 - c1 + 1), IMAGE_MAX_CELLS))

    download = _bool_arg(args.get("download"), True, "download")
    allow_clip = _bool_arg(args.get("allow_clipboard"), True, "allow_clipboard")
    out_dir = image_out_dir(args.get("dir"))
    sheet_name = host._safe_com_str(ws, "Name", "")
    doc_path = host._safe_com_str(doc, "FullName", "") or ""

    # 1) 先把 ref 范围内的「图从哪来」扫出来（只读公式，不碰值）
    found = []
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = host._com_call(ws, "Cells", r, c)
            ok, raw = host._com_try_attr(cell, "Formula")
            text = "" if not ok or raw is None else str(host._com_to_plain(raw))
            m = _DISPIMG_RE.search(text)
            if m:
                found.append({"cell": ref_text(r, c, r, c), "kind": "wps_cellimage",
                              "id": m.group(1)})
                continue
            m = _IMAGE_FN_RE.search(text)
            if m:
                found.append({"cell": ref_text(r, c, r, c), "kind": "url_formula",
                              "url": m.group(1)})
                continue
            u = _URL_RE.search(text)
            if u:
                found.append({"cell": ref_text(r, c, r, c), "kind": "url_text",
                              "url": u.group(0)})
    found += _shapes_in_box(host, ws, r1, c1, r2, c2)

    # 2) 建「ID → 包里 media」索引（文档存过盘才有）
    idmap, zip_note = cellimages_index(doc_path)

    images = []
    for item in found:
        rec = dict(item)
        stem = _safe_stem("%s_%s" % (sheet_name, item["cell"]))
        if item["kind"] in ("url_formula", "url_text"):
            rec["source"] = "url"
            if download:
                try:
                    path, size, fmt = _download_image(item["url"], out_dir, stem)
                    rec.update({"path": path, "bytes": size, "format": fmt,
                                "width": _image_size(path)})
                except Exception as exc:
                    rec["error"] = "链接下载失败（%s）；仍可让模型直接访问该 URL" % exc
            images.append(rec)
            continue

        if item["kind"] == "shape":
            rec["source"] = "clipboard"
            if not allow_clip:
                rec["error"] = "是浮动图片，但 allow_clipboard=false"
                images.append(rec)
                continue
            ok_s, shapes = host._com_try_attr(ws, "Shapes")
            if not ok_s or shapes is None:
                rec["error"] = "取不到 Shapes 集合"
                images.append(rec)
                continue
            shp = host._com_call(shapes, "Item", item["shape_index"])
            path, size, err = _grab_via_clipboard(host, shp,
                                                  os.path.join(out_dir, stem + ".png"), is_shape=True)
            if err:
                rec["error"] = err
            else:
                rec.update({"path": path, "width": size, "format": "png"})
            images.append(rec)
            continue

        # wps_cellimage
        member = idmap.get(item["id"])
        if member:
            try:
                path, size, fmt = _extract_cellimage(doc_path, member, out_dir, stem)
                rec.update({"source": "xlsx_media", "path": path, "bytes": size, "format": fmt,
                            "width": _image_size(path), "media": member,
                            "note": "从存盘文件的 xl/media 取到原图（全分辨率）"})
            except Exception as exc:
                rec["error"] = "从包里抠图失败（%s）" % exc
        elif not allow_clip:
            rec["source"] = "none"
            rec["error"] = ("包里没有这张图（%s）。可能是刚插入还没保存；"
                            "确认要截屏兜底就传 allow_clipboard=true" % (zip_note or "未找到"))
        else:
            range_ref = _range_of(host, ws, *_whole_merge_box(host, ws, r1, c1, r2, c2))
            path, size, err = _grab_via_clipboard(host, range_ref,
                                                  os.path.join(out_dir, stem + ".png"))
            if err:
                rec.update({"source": "none", "error": err,
                            "hint": "可以先请人按 Ctrl+S 存盘，再从 xl/media 取原图"})
            else:
                rec.update({"source": "clipboard", "path": path, "width": size, "format": "png",
                            "note": ("文档还没存盘，包里没有这张图，只能用剪贴板截屏："
                                     "分辨率=屏幕、大图会被单元格裁切。会覆盖剪贴板。"
                                     "要原图请让人 Ctrl+S 后再取一次。")})
        images.append(rec)

    out = {"app": "excel", "sheet": sheet_name, "ref": ref_text(r1, c1, r2, c2),
           "count": len(images), "images": images, "live": True}
    if doc_path:
        out["doc"] = doc_path
    if zip_note:
        out["cellimages_note"] = zip_note
    if not images:
        out["hint"] = ("这个范围里没找到图片。图片只有三种可能：=DISPIMG(...) 嵌入图、"
                       "=IMAGE(...)/裸链接、浮动 Shape。若是浮动图片但不在这个范围，"
                       "用 where 看表结构后给出图片所在的格子。")
    return out


def help_text(app, topic=None, host=None):
    if topic in ("legacy", "Legacy"):
        names = legacy_names(host) if host is not None else []
        # 用 replace 而不是 format：正文里本身就有 {"tool": ...} 这种大括号，format 会炸。
        return LEGACY_TOPIC.replace("NAMES_HERE", ", ".join(names) or "(host 未提供 TOOLS)")
    return HELP.get(app, "未知 app=%r；可用 excel / word / ppt" % (app,))


# ==================== 极简工具定义（常驻 tools/list 的全部内容） ====================

THIN_TOOLS = [
    {
        "name": "wps_docs",
        "description": "列出人当前在 WPS 里打开的文档（表格/文字/演示），含未保存的。只读。",
        "inputSchema": {
            "type": "object",
            "properties": {
                "app": {"type": "string", "enum": ["excel", "word", "ppt", "all"],
                        "description": "只看某一类应用，缺省 all"},
            },
        },
    },
    {
        "name": "wps_do",
        "description": ("操作人正在打开的 WPS 活文档（读内存实时值，不是磁盘旧文件）。"
                        "参数不确定时先 action=\"help\" 取该应用的参数说明。"),
        "inputSchema": {
            "type": "object",
            "properties": {
                "app": {"type": "string", "enum": ["excel", "word", "ppt"], "description": "目标应用"},
                "action": {"type": "string",
                           "description": "动作名；先用 action=\"help\" 取该应用的完整动作表与参数"},
                "args": {"type": "object", "description": "动作参数，见 help",
                         "additionalProperties": True},
                "path": {"type": "string",
                         "description": "可选：目标文档绝对路径；缺省=人当前打开的那一本"},
            },
            "required": ["app", "action"],
        },
    },
]


# ==================== 分发 ====================


def error_payload(host, app, action, err):
    """把异常翻成「能返回给 AI 的结构化结果」。

    为什么必须翻：走 stdio 的 JSON-RPC 里，未被捕获的异常会变成裸错误
    （2026-09-24 实测过 `-32000 (-2147352567 发生意外)`，人根本看不懂，AI 也学不到东西）。
    这里统一翻成人话 + 下一步该干什么：
      · WpsAttachError → 原样返回它的结构化 payload（含 reason / human_action_required）
      · ValueError     → 参数错误 + 指路 help
      · 其它异常       → 保留原文 + 给出最可能的两种下一步
    """
    payload = getattr(err, "payload", None)
    if isinstance(payload, dict):
        return payload
    if isinstance(err, ValueError):
        return {"ok": False, "error": str(err), "app": app, "action": action,
                "hint": "参数不对。用 wps_do(app=%r, action=\"help\") 能看到正确写法与例子。" % (app,)}
    return {"ok": False, "error": str(err), "error_type": type(err).__name__,
            "app": app, "action": action,
            "hint": ("这次没成功。如果是「人正在单元格里编辑」导致的，请人按一下 Esc 或回车后重试；"
                     "否则先 action=\"where\" 看当前状态，再用 action=\"help\" 核对参数。")}


def handle(host, app, action, args=None, path=None):
    """薄壳总入口。返回可 JSON 化的 dict；错误用 ValueError（人话）或 host.WpsAttachError（结构化）。"""
    missing = [n for n in HOST_API if not hasattr(host, n)]
    if missing:
        raise RuntimeError("host 缺少老地基本模块接口：%s（server.py 接线不完整）" % ", ".join(missing))

    app = normalize_app(app)
    action, mode_flag = split_mode_suffix(action)
    args = dict(args or {})
    if mode_flag:
        for key, value in mode_flag.items():
            args.setdefault(key, value)      # 显式传的参数优先，后缀只是省字的另一种写法
    if app not in APP_KEYS:
        raise ValueError("app 必须是 excel / word / ppt，收到 %r" % (app,))
    if not action:
        raise ValueError("缺少 action；不确定就先 action=\"help\"")

    if action in ("help", "?", "说明"):
        out = {"help": help_text(app, args.get("topic"), host), "app": app}
        out.update(snapshot_info())
        return out

    if action == "legacy":
        return _dispatch_legacy(host, app, args)

    write = is_write(app, action)
    table = {
        "excel": {"where": _excel_where, "get": _excel_get, "set": _excel_set,
                  "save": _excel_save, "activate": _excel_activate,
                  "width": _excel_width, "height": _excel_height, "fill": _excel_fill,
                  "style": _excel_style, "merge": _excel_merge, "unmerge": _excel_unmerge,
                  "brush": _excel_brush, "insert": _excel_insert, "delete": _excel_delete,
                  "clear": _excel_clear, "copy": _excel_copy, "move": _excel_move,
                  "find": _excel_find, "replace": _excel_replace, "image": _excel_image},
        "word": {"where": _word_where, "get": _word_get, "set": _word_set, "save": _word_save},
        "ppt": {"where": _ppt_where, "get": _ppt_get, "save": _ppt_save},
    }[app]
    fn = table.get(action)
    if fn is None:
        raise ValueError(
            "app=%s 没有 action=%r。可用：help, %s, legacy。"
            "用 action=\"help\" 可以拿到完整参数说明。" % (app, action, ", ".join(sorted(table))))

    def _run(wps_app, doc, mode):
        return fn(host, doc, mode, args)

    return with_doc(host, app, path, write, _run)


def _dispatch_legacy(host, app, args):
    tool = args.get("tool")
    if not tool:
        raise ValueError("legacy 需要 args.tool（老工具名）+ args.args（该工具的参数）")
    if tool in _LEGACY_WRITE:
        raise ValueError(
            "legacy 不支持写类老工具（%s）——它们会 wb.Save()，会吃掉人的 Ctrl+Z。"
            "写数据请用 action=\"set\"；确实需要时请在对话里明确要求。" % (tool,))
    inner = dict(args.get("args") or {})
    inner.setdefault("mode", "com")
    import json as _json
    result = host.handle_tool_call(tool, inner)
    try:
        _json.dumps(result, ensure_ascii=False)
        return {"app": app, "legacy": tool, "result": result}
    except Exception:
        return {"app": app, "legacy": tool, "result": str(result)}


# ==================== 分层加载效果（2026-09-28 实测数字） ====================
# 老 tools/list：26 个工具定义，11,605 字符，**每一轮都付**。
# V3 tools/list：THIN_TOOLS 两个工具共 755 字符 → **省 93.5%**。
# 真要用 Excel 时再取一次 help（1,357 字符，只含 Excel；不取就不付）：
#   全程不碰 WPS / 只干 Word：  755 字符/轮
#   全程只干 Excel：            755 + 随时可选 1,357
# 想回退成老样子：环境变量 WPS_MCP_LEGACY_TOOLS=1（tools/list 立刻恢复 26 个）。
