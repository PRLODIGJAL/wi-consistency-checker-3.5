# -*- coding: utf-8 -*-
"""WI 一致性检查工具 —— 核心模块。

本文件由 8 个模块合并而成，**只做物理搬运，不改任何判定逻辑**：
    原 schema.py / pdf_schema.py / excel_audit.py / workbook_audit.py /
       outlier_scan.py / exporter.py / engine.py / batch.py

为什么要合并：交付和携带时只需一个文件，不会漏拷某个模块（这个坑真实发生过）。
代码逐行保持原样，只删掉了"指向兄弟模块的 import"（它们现在同处一个文件），
并把 pdf_schema 里那个与 engine 同名的私有函数 `_equip_set` 改名为
`_equip_set_spec`（两者实现不同、且各自只被本段使用，改名是唯一办法）。

分节索引（共 8 节；终端里 `grep -nE "^#  第 [0-9]+ 节" wi_core.py` 可直接跳转）：
    第 1 节   Excel 结构自适应       第 5 节   同规格横向一致性
    第 2 节   PDF 版式抽取           第 6 节   导出 xlsx 与历史快照
    第 3 节   单表自检（纯 Excel）   第 7 节   判定引擎
    第 4 节   覆盖面与 WI 版本链     第 8 节   批量核对

类型注解：只标「对外入口」（app.py / 回归标尺 / 批处理会直接调的那些），
内部 helper 一律不标 —— 它们传的是松散 dict，硬标收益低且容易标错。
凡标注必须与函数体实际返回一致；拿不准就不标。

★ 改完本文件必须跑零回归标尺（判定结果必须一字不变）：
    python _regress/regress.py . _regress/out_cur.json
  期望 master=5a676f58b896dfa81d7b53e9beba43960f1036324cd2b732ece6fb1bba4b1aca  lines=56800
"""

from __future__ import annotations

# 外部依赖（合并前散落在 8 个模块里各写一遍，现统一收到这里）
# 全工具只用两个三方库：pdfplumber（读 PDF）+ openpyxl（读/写 Excel）
import json
import os
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field as dc_field
from datetime import datetime
from functools import lru_cache

import openpyxl
import pdfplumber
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ── 目录四层自举（2026-11-03 目录改造）────────────────────────────────────
# 本文件在 app/engine/ 里；脚本方式启动时 sys.path[0] 是这一层。
# 第 6905 行那句 `from wic.memory.snapshot import …` 是**模块级**导入，
# 所以四层必须在那之前挂上，否则单独 import 本文件会直接 ImportError。
import sys as _sys
_APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _APP_DIR not in _sys.path:
    _sys.path.insert(0, _APP_DIR)
import paths as _paths                                   # noqa: E402
_paths.install()


# ==========================================================================
#  第 0 节  PDF 底层健壮性补丁（对所有 PDF 生效，不是任何一份文件的特例）
# ==========================================================================

"""为什么要有这一节：车间产出的 WI-PDF 由各种转换工具生成，语法并不都合规。
pdfminer（pdfplumber 的底层）对其中两类畸形会直接崩、或者静默把对象读错，
表现到用户眼里就是「这一份 WI 一个参数都识别不到」。

已证实的两类畸形（真实案例：WI-0004 rev28）：
  1. **关键字与下一个对象编号粘连**。生产工具写出 `endstream\\nendobj29 0 obj`
     —— `endobj` 后面没有换行，直接跟下一个对象的编号。pdfminer 的词法器把
     `endobj29` 当成一个整体关键字，于是 `KEYWORD_ENDOBJ` 永远不匹配，
     对象链整个错位。该 PDF 76 个 endobj 里有 10 处这样写，恰好落在
     F2~F5 四个拉丁/数字字体的 ToUnicode 上 —— 结果就是中文能读、
     `T-1011` / `550~650mW` 全变成乱码。
     修法：把 `endobj<数字>` 拆回 `endobj` 关键字 + 数字压栈。
  2. **流的 /Length 比实际少算**。上例 obj28 声明 1361、实际 1362（末尾多一个换行）。
     pdfminer 的 stream 分支本来就有「往后找 endstream」的兜底循环，能自己吸收，
     所以只要第 1 类修好、这一类不需要额外处理。

为什么用 monkeypatch 而不是改 pdfminer 源码：交付包里的 pdfminer 是 pip 现场装的，
改不了、也不该改。补丁只在这类畸形出现时才触发，对合规 PDF 是**零行为变化**
（已用 10 份语料的零回归标尺验证：判定行数与签名一字不变）。

为什么不上 OCR：能读就一定用文本层。OCR 会把 `3.0-4.0s` 读成 `3.0-4.05`、
把 `T-1011` 读成 `T-101I`；而且给 Windows 交付包加 OCR 引擎是几十 MB 的依赖负担。
只有确认是**图片扫描件**（文本层为空）时才值得动 OCR，届时再单独评估。
"""

_GLUED_KW = (b"endobj", b"endstream", b"stream", b"obj", b"xref", b"startxref", b"trailer")
_GLUED_RE = re.compile(rb"^(endobj|endstream|stream|obj|xref|startxref|trailer)(\d+)$")


def _install_pdf_robustness() -> None:
    """给 pdfminer 打一层语法容错补丁。失败则静默跳过（绝不能让工具打不开）。"""
    try:
        from pdfminer.pdfparser import PDFParser
        from pdfminer.psparser import KWD, PSKeyword
    except Exception:
        return

    if getattr(PDFParser.do_keyword, "_glued_patched", False):
        return
    orig = PDFParser.do_keyword

    def do_keyword(self, pos, token):  # noqa: ANN001 - 保留 pdfminer 原签名
        name = getattr(token, "name", None)
        if isinstance(token, PSKeyword) and isinstance(name, (bytes, bytearray)):
            name = bytes(name)
            m = _GLUED_RE.match(name)
            if m and name[: len(m.group(1))] in _GLUED_KW:
                # 例：b"endobj29" -> 先按 endobj 正常收尾，再把 29 压回栈作为下一个对象号
                orig(self, pos, KWD(m.group(1)))
                self.push((pos, int(m.group(2))))
                return
        orig(self, pos, token)

    do_keyword._glued_patched = True
    PDFParser.do_keyword = do_keyword


_install_pdf_robustness()


# ==========================================================================
#  第 1 节  Excel 结构自适应
#  按列名语义定位表头与字段，不写死 sheet/列号
# ==========================================================================

"""Excel 结构自适应层 —— 让「换个表格就得改代码」变成「换个表格照样跑」。

设计原则（重要）：
1. **不出现任何列号常量**。所有字段一律按「列名语义」定位。
2. **不依赖 sheet 名**。哪张表是采集项明细表，由列名构成自动判定。
3. **不依赖表头行位置**。表头在第 1 行还是第 2 行（前面有标题行）都能找到。
4. **认不出来就报错**（SchemaError），绝不静默返回 0 条。
5. 自检报告只描述**结构**（sheet 名 / 表头行 / 列名 / 类型统计），不含任何数据值。

对外接口：
    ts = TableSchema.load(path)            # 自动识别
    ts.usable                              # 是否可用
    ts.report()                            # 结构自检文本（可打印给用户看）
    ts.records(wi_name="WI-0008")          # 按 WI 抽取采集项
"""


class SchemaError(Exception):
    """结构无法识别，或必需字段缺失。绝不静默返回空结果。"""


# ------------------------------------------------------------------ 归一化

_FULL2HALF = str.maketrans({chr(0xFF01 + i): chr(0x21 + i) for i in range(94)})


def _key(s):
    """列名归一：全角转半角 → 去空白 → 去括号注释 → 去首尾星号冒号 → 小写。"""
    s = str(s if s is not None else "")
    s = s.translate(_FULL2HALF)
    s = re.sub(r"[\s\u3000\u00a0]+", "", s)
    s = re.sub(r"[（(][^）)]*[）)]", "", s)
    s = s.strip("*:：")
    return s.lower()


def _kind(v):
    if v is None:
        return "empty"
    if isinstance(v, bool):
        return "bool"
    if isinstance(v, (int, float)):
        return "num"
    s = str(v).strip()
    if not s:
        return "empty"
    return "num" if re.fullmatch(r"-?\d+(?:\.\d+)?", s.replace(",", "")) else "text"


def to_num(v):
    """把单元格转成 float；取不到数字返回 None。"""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r"-?\d+(?:\.\d+)?", str(v).replace(",", ""))
    return float(m.group(0)) if m else None


def num_precision(v):
    """提取一个数字原文的「小数位数」签名：3→0、3.0→1、3.5→1、0.026→3。

    这是「3.0 与 3 不能算对」的判据来源：数值相等、但小数位数不同，
    就说明两边的写法精度不一致，不能判一致。非纯数字（带单位/区间/后缀，
    如 `3s`、`95-115`、`>95℃`）返回 None —— 这类原文没有「小数位数」可言，
    不参与精度校验（此时退回到只看数值的旧口径）。

    为什么不用 Decimal：Decimal 的 `==` 是数值比较，`Decimal('3.0') ==
    Decimal('3')` 返回 True，精度信息照样丢失；必须显式读 `.as_tuple().exponent`
    才能拿到位数 —— 那不如直接从原文算，省一层对象。
    """
    if v is None or isinstance(v, bool):
        return None
    s = str(v).strip()
    # 只认「纯数字（可带一个负号、一个小数点）」；其余（单位、区间、不等号）一律 None
    m = re.fullmatch(r"-?\d+(?:\.\d+)?", s.replace(",", ""))
    if not m:
        return None
    dot = m.group(0).find(".")
    if dot < 0:
        return 0
    # 去掉末尾可能被 normalize 的精度只算「小数点后实际写了几个数字」
    return len(m.group(0)) - dot - 1


# 「括号里写的是不是单位」——只认纯量纲写法：mm / cm / MPa / rpm / mm/s / kPa /
# ℃ / °C / % / ‰ 这类。判据故意收得很死，因为这一步**只用来决定要不要把括号从
# 名字里剥出去**，剥错了代价很大（名字被削短，字面名通道、词典键全都跟着变）。
#
# 反例都实测过：`（首件）` `（见图）` `（Step1）` `（外观）` 这些括号里写的不是量纲，
# 老的半角逻辑本来就不剥（因为压根不是半角括号），这里也必须维持原样 ——
# 一旦把 `（Step1）` 当成单位，单位收窄层会拿 'Step1' 去和 PDF 的 'mm' 比，
# 整池记录被滤空，`ok` 直接掉成 `no_match`，那是**漏报**。
_UNIT_TOKEN = re.compile(
    r"^(?:[°℃℉]|[°℃℉]?[A-Za-zµμ]{1,5}|%|‰)(?:\s*/\s*[A-Za-zµμ]{1,6})?$")


def _looks_like_unit(u):
    """括号里的内容像不像单位（宁可不认，也不认错）。"""
    u = str(u or "").strip()
    return bool(u) and bool(_UNIT_TOKEN.match(u))


def _unwrap_unit(u):
    """**整串被括号包住**时把括号脱掉：「（Inch）」→「Inch」。

    ⚠️ 2026-10-30：这个函数目前**没有接进 `parse_item_name`**，只留给界面用
    （`name_lex.candidates` 判断「单位相同 ·」时用它归一，纯显示）。
    接进解析链会连带改掉「拿什么单位去收窄候选」，实测代价是一条假差异 ——
    理由与实测数据见 `_norm_unit` 的说明。
    """
    s = str(u or "").strip()
    if len(s) >= 2 and s[0] in "(（" and s[-1] in ")）":
        inner = s[1:-1].strip()
        if _looks_like_unit(inner):
            return inner
    return s


def parse_item_name(s):
    """'温度:(℃)' → ('温度','℃')；'产品上提01(mm/s)' → ('产品上提01','mm/s')。"""
    s = str(s or "").strip()
    m = re.match(r"^(.*?)[:：]\s*\(?([^()]*)\)?$", s)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    m = re.search(r"\(([^()]*)\)\s*$", s)
    if m:
        return s[:m.start()].strip(), m.group(1).strip()
    # —— 全角括号（2026-10-30）——
    # Excel 里「扩口 温度（℃）」用的是中文括号，上面两条只认半角 `(` `)`，
    # 于是整串名字连括号一起留下、`unit` 退化成空串。
    # 后果不是「少显示一个单位」这么轻：单位收窄那一层由 `_excel_unit(er)` 驱动，
    # **空单位 = 这一层整个不生效**。WI-0039 实测：PDF 同一页里既有一条
    # `温度30min`（单位 min、值 30），又有 16 条 `110~130℃` —— 单位层失效后
    # `30` 混进 ℃ 的候选池，设备层再把 16 条 ℃ 全排除（它们挂 T-1279/T-2084，
    # 本行写 T-2025/T-1278），剩它一条 → `strict` → 判「应为 30」，
    # 210 行全是这类假差异。
    # 只改进「认得出的量纲」，括号里写着「首件 / 见图 / Step1」的一律不动。
    m = re.search(r"[（(]([^（）()]*)[）)]\s*$", s)
    if m and _looks_like_unit(m.group(1)):
        return s[:m.start()].strip(), m.group(1).strip()
    return s.rstrip(":："), ""


def merge_fill_maps(ws, ncol, max_area=200000):
    """读「合并单元格」的填充关系。

    返回 (child2anchor, anchors)：
      child2anchor: {(row, col) -> (anchor_row, anchor_col)}  被合并覆盖的格子 → 左上角
      anchors:      {(anchor_row, anchor_col)}                真正带值的左上角
    行列均为 1-based，与 openpyxl 一致。面积过大的合并区域跳过，避免生成巨量映射。
    """
    child, anchors = {}, set()
    try:
        ranges = list(ws.merged_cells.ranges)
    except Exception:
        return child, anchors
    for rng in ranges:
        if rng.min_col > ncol:
            continue
        area = (rng.max_row - rng.min_row + 1) * (rng.max_col - rng.min_col + 1)
        if area > max_area:
            continue
        ar, ac = rng.min_row, rng.min_col
        anchors.add((ar, ac))
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, min(rng.max_col, ncol) + 1):
                if (r, c) != (ar, ac):
                    child[(r, c)] = (ar, ac)
    return child, anchors


# ------------------------------------------------------------- 字段同义词表
# 键 = 语义字段；值 = (中文标签, [同义词...])
# 同义词必须写成「已归一化」形式（小写、无空格）。
FIELD_SPEC = {
    "seq":      ("序号",     ["序号", "行号", "no", "no."]),
    "card":     ("参数卡",   ["工艺参数卡名称", "参数卡名称", "参数卡", "卡名称", "参数卡名"]),
    "card_ver": ("卡版本",   ["工艺参数卡版本", "参数卡版本", "卡版本", "版本号"]),
    "proc":     ("工艺/工序", ["工艺名称", "工序名称", "工序", "工艺"]),
    "group":    ("组",       ["组序号", "分组排序", "分组序号", "组号", "分组"]),
    "item":     ("采集项名称", ["数据采集项名称", "采集项名称", "数据采集项", "参数名称", "采集项", "项目"]),
    "type":     ("类型",     ["类型", "数据类型", "值类型"]),
    "wi":       ("WI 编号",  ["wi文档名称", "wi编号", "wi文档", "文件编号", "wi"]),
    "lo":       ("下限",     ["下限", "最小值", "最小", "低限"]),
    "hi":       ("上限",     ["上限", "最大值", "最大", "高限"]),
    "fixed":    ("固定值",   ["固定值", "设定值", "定值"]),
    "btrue":    ("布尔真",   ["布尔值真", "真值", "正确值"]),
    "bfalse":   ("布尔假",   ["布尔值假", "假值", "错误值"]),
    # 两个「是/否」开关列：表里本来就有，但主结构原先没认（只有第 3 节自检那套
    # 映射认）。界面要把「表里填的」和「按类型规矩该填的」并排显示，判定层就得
    # 读得到它们。不进 _score（选表打分）—— 有没有这两列不该影响选哪张表。
    "show_lim": ("显示上下限", ["是否显示上下限", "显示上下限", "显示限值",
                               "是否显示限值"]),
    "allow_over": ("允许超上下限", ["是否允许超上下限", "允许超上下限",
                                   "允许超限值", "是否允许超限"]),
    "spec":     ("适用规格", ["适用规格", "适用尺寸", "规格型号", "规格", "尺寸"]),
    "equip":    ("设备",     ["设备编号", "设备名称", "设备号", "设备"]),
}

NUMERIC_FIELDS = ("lo", "hi", "fixed")
BOOL_FIELDS = ("btrue", "bfalse")
# 「该列几乎不含数字就判认错」这道校验**只管 下限 / 上限**，不能管 固定值：
# 固定值列按填写规范本来就存成**文本**（第 2 节审计里明写「固定值 / 参数卡版本
# 需为文本格式」）。拿数字占比去筛它，等于把最正当的一列判死 ——
# DWG-0000002 表4 实测：固定值列前 25 行有 22 格文本（'0' / '2.5' / '1000'），
# 整列被忽略；而同一段里 下限/上限 列**全是空的**，反倒因样本不足逃过校验。
# 结果是 210 项里有 105 项「类型=固定值」被当成未填值，压根没进核对 —— 这是
# 最危险的一类漏判：值明明填着，工具却说没填，人工复核也不会再看第二眼。
_NUMERIC_PROFILE_GUARD = ("lo", "hi")
# ⚠️ `REQUIRED_FIELDS = ("item",)` 已删（2026-09-30 重构）：与下面的
# `CRITICAL_FIELDS` 同名同值、却是旧名，且没有任何一处读它。
CRITICAL_FIELDS = ("item",)

# 分组类字段：Excel 里经常用「合并单元格」或「留空」表示「同上」。
# 这些字段的空白格要向上继承，否则整批记录会丢掉归属（参数卡名变成 None）。
FILL_MERGE = ("card", "card_ver", "proc", "group", "spec", "equip", "wi")
FILL_CARRY = ("card", "card_ver", "proc", "group")   # 无合并声明时也允许顺延的字段

# 「是否…」开头的列是开关列，绝不是数值列（防止「是否显示上下限」被认成「上限」）
_SWITCH_PREFIX = ("是否", "is", "has", "允许", "enable")


def _syn_pairs():
    """展开成 [(归一化同义词, 字段)]，按同义词长度降序 → 长词优先，避免「名称」抢「参数卡名称」。"""
    out = []
    for f, (_label, syns) in FIELD_SPEC.items():
        for s in syns:
            out.append((s, f))
    out.sort(key=lambda x: -len(x[0]))
    return out


_SYN = _syn_pairs()


def _match_field(col_name):
    """给一个列名，返回 (字段, 是否精确匹配)。认不出返回 (None, False)。"""
    k = _key(col_name)
    if not k:
        return None, False
    is_switch = any(k.startswith(p) for p in _SWITCH_PREFIX)
    # 1) 精确匹配
    for s, f in _SYN:
        if k != s:
            continue
        if is_switch and f in NUMERIC_FIELDS + BOOL_FIELDS:
            continue
        return f, True
    # 2) 包含匹配
    for s, f in _SYN:
        if s not in k:
            continue
        if is_switch and f in NUMERIC_FIELDS + BOOL_FIELDS:
            continue
        return f, False
    return None, False


def _exact_field_cols(header, field):
    """只按「精确同义词」找列。

    关联表里常有「设备编号」和「设备组」并存，用包含匹配会把「设备组」也认成设备，
    所以这里必须精确。
    """
    syns = set(FIELD_SPEC[field][1])
    return [c for c, n in enumerate(header)
            if n is not None and _key(n) in syns]


def _usable_join_key(s, placeholder=("/", "-", "—", "\\", "N/A", "NA")):
    """关联表里被用来「无值」的占位符不算有效键。"""
    s = str(s or "").strip()
    return s if s and s not in placeholder else None


# --------------------------------------------------------------- 表头行发现

def _read_head(ws, n=25, max_col=80):
    rows = []
    for i, r in enumerate(ws.iter_rows(min_row=1, max_row=n, max_col=max_col, values_only=True), 1):
        rows.append(r)
        if i >= n:
            break
    return rows


def _header_score(rows, i):
    """第 i 行（0-based）作为表头的得分；不可能为表头返回 None。"""
    r = rows[i]
    ks = [_kind(c) for c in r]
    n = sum(1 for k in ks if k != "empty")
    # 至少 2 个非空单元格即可能是表头。原判据是 <3，于是「参数卡 | WI编号」
    # 这类两列关联表会被当成"整张表没有表头"直接跳过，按参数卡回填 WI / 设备
    # 的关联随之失效（换个工作簿就可能整批取不到采集项）。放宽到 2 列是安全的：
    # 下面还有「文本占比 ≥0.6」与「表头下方必须出现数字或更长的文本」两道判据。
    if n < 2:
        return None
    text = sum(1 for k in ks if k == "text")
    if text / n < 0.6:
        return None
    maxlen = max((len(str(c).strip()) for c in r if c is not None), default=0)
    if maxlen > 24:
        return None
    # 表头下面必须出现「数字」或「更长的文本」→ 才说明这是表头而非内容
    sig = False
    for r2 in rows[i + 1:i + 6]:
        ks2 = [_kind(c) for c in r2]
        if any(k == "num" for k in ks2):
            sig = True
            break
        ml2 = max((len(str(c).strip()) for c in r2 if c is not None), default=0)
        if ml2 > maxlen + 6:
            sig = True
            break
    if not sig:
        return None
    # 命中字段越多越像表头
    hits = sum(1 for c in r if c is not None and _match_field(c)[0])
    return 10 + hits


def find_header_row(rows):
    """返回表头行下标（0-based），找不到返回 None。"""
    best, best_s = None, -1
    for i in range(min(len(rows), 15)):
        s = _header_score(rows, i)
        if s is not None and s > best_s:
            best, best_s = i, s
            break  # 取最靠上的合格行
    return best


# ------------------------------------------------------------------ 列映射

def map_columns(header, profiles=None):
    """按列名语义把表头映射到字段。

    header   : 表头行（tuple/list）
    profiles : {col: Counter} 该列下方数据的类型分布，用于校验数值列
    返回 (mapping{field: col}, unknown[col], conflicts[note])
    """
    cands = []
    for c, name in enumerate(header):
        f, exact = _match_field(name)
        if f is None:
            continue
        k = _key(name)
        syn_len = len(k)
        cands.append((c, f, exact, syn_len, k))
    # 精确优先 → 同义词更长优先 → 列更靠前优先
    cands.sort(key=lambda x: (not x[2], -x[3], x[0]))

    mapping, used, conflicts = {}, {}, []
    for c, f, exact, _sl, k in cands:
        if f in used:
            conflicts.append("列「%s」也想当【%s】，已让位给「%s」" %
                             (k, FIELD_SPEC[f][0], _key(header[used[f]])))
            continue
        # 数值列的类型校验：整列几乎没有数字 → 大概率认错
        # （布尔列不校验：布尔值只有少数行有值，且可能是 TRUE/FALSE 文本）
        if profiles is not None and f in _NUMERIC_PROFILE_GUARD:
            p = profiles.get(c, Counter())
            tot = p.get("num", 0) + p.get("text", 0) + p.get("bool", 0)
            if tot >= 5 and p.get("num", 0) / tot < 0.2:
                conflicts.append("列「%s」认作【%s】但该列几乎不含数字，已忽略" % (k, FIELD_SPEC[f][0]))
                continue
        mapping[f] = c
        used[f] = c

    unknown = [c for c, name in enumerate(header)
               if name is not None and str(name).strip() and c not in mapping.values()]
    return mapping, unknown, conflicts


# ------------------------------------------------------------------ 结构对象

@dataclass
class SheetSchema:
    name: str
    header_row: int                 # 1-based
    header: list
    mapping: dict                   # field -> col (0-based)
    profiles: dict = dc_field(default_factory=dict)
    hidden: bool = False
    score: int = 0
    conflicts: list = dc_field(default_factory=list)
    unknown: list = dc_field(default_factory=list)

    @property
    def usable(self):
        return all(f in self.mapping for f in CRITICAL_FIELDS)

    @property
    def label(self):
        return "%s!r%d" % (self.name, self.header_row)


@dataclass
class TableSchema:
    path: str
    sheets: list = dc_field(default_factory=list)      # 全部 sheet 的 SheetSchema
    chosen: SheetSchema = None                          # 被选中的采集项表
    wi_source: str = ""                                 # column / relation:<sheet> / none
    equip_source: str = ""                              # column / relation:<sheet> / none
    card_wi: dict = dc_field(default_factory=dict)      # 关联回填用：参数卡 → WI
    card_equip: dict = dc_field(default_factory=dict)   # 关联回填用：参数卡 → 设备编号
    notes: list = dc_field(default_factory=list)
    errors: list = dc_field(default_factory=list)
    # records() 的单条记忆：(筛选条件, 结果)。同一次比对里 records() 会被调两次
    # （抽采集项一次、找「同名列挂在别的 WI 下」的线索再一次），每次都要整表读一遍。
    _rec_cache: tuple = dc_field(default=None, repr=False)

    # ---------------------------------------------------------- 识别
    @classmethod
    def load(cls, path: str, sheet: str | None = None) -> "TableSchema":
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        try:
            self = cls(path=path)
            for sn in wb.sheetnames:
                ws = wb[sn]
                hidden = getattr(ws, "sheet_state", "visible") != "visible"
                rows = _read_head(ws)
                hi = find_header_row(rows)
                if hi is None:
                    self.sheets.append(SheetSchema(sn, 0, [], {}, hidden=hidden))
                    continue
                header = list(rows[hi])
                ncol = len(header)
                after = [r for r in rows[hi + 1:hi + 31]]
                profiles = {}
                for c in range(ncol):
                    p = Counter()
                    for r in after:
                        p[_kind(r[c] if c < len(r) else None)] += 1
                    profiles[c] = p
                mapping, unknown, conflicts = map_columns(header, profiles)
                score = cls._score(mapping, hidden)
                self.sheets.append(SheetSchema(sn, hi + 1, header, mapping, profiles,
                                               hidden=hidden, score=score,
                                               conflicts=conflicts, unknown=unknown))
            self._choose(sheet)
            self._resolve_relations(wb)
            return self
        finally:
            wb.close()

    @staticmethod
    def _score(mapping, hidden):
        if hidden:
            return -1
        s = 0
        if "item" in mapping:
            s += 4
        if "type" in mapping:
            s += 3
        if "lo" in mapping or "hi" in mapping:
            s += 2
        if "fixed" in mapping:
            s += 2
        if "wi" in mapping:
            s += 3
        if "card" in mapping:
            s += 1
        if "spec" in mapping:
            s += 1
        if "equip" in mapping:
            s += 1
        if "proc" in mapping:
            s += 1
        return s

    def _choose(self, sheet=None):
        usable = [s for s in self.sheets if s.usable and not s.hidden]
        if sheet:
            hit = [s for s in self.sheets if s.name == sheet or s.name.strip() == sheet.strip()]
            if not hit:
                self.errors.append("指定的 sheet「%s」不存在" % sheet)
                return
            self.chosen = hit[0]
            return
        if not usable:
            self.errors.append(
                "没有任何 sheet 含【%s】列，无法识别采集项明细表" %
                "】【".join(FIELD_SPEC[f][0] for f in CRITICAL_FIELDS))
            return
        usable.sort(key=lambda s: (-s.score, s.name))
        self.chosen = usable[0]
        if len(usable) > 1:
            self.notes.append("检测到 %d 张表都含采集项字段，已按列名完整度选用「%s」；"
                              "候选明细：%s。如需改用其它表，可在界面上手动指定。"
                              % (len(usable), self.chosen.name,
                                 "、".join("%s(分%d, 表头r%d)" % (s.name, s.score, s.header_row)
                                           for s in usable)))

    # ------------------------------------------------- WI / 设备来源
    def _resolve_relations(self, wb):
        """补齐采集项表「自己没有」的两类归属键：WI 与 设备编号。

        很多工程参数表是主数据式的：采集项明细表只有「参数卡」，WI 和设备编号
        放在另一张关联表里（同表通常还有「参数卡1/参数卡2/设备编号」几列）。
        这里扫描这类关联表，按参数卡名把 WI 和设备编号回填进来。

        为什么设备编号这么重要：PDF 里同一台设备往往有好几套参数（按直径范围 /
        程序号区分）。拿不到设备号 → 匹配无法收窄 → 期望值会变成整份 PDF 的并集，
        把别的设备的值也显示出来，直接误导使用者。
        """
        ch = self.chosen
        if ch is None:
            return

        self.wi_source = "column" if "wi" in ch.mapping else ""
        self.equip_source = "column" if "equip" in ch.mapping else ""
        if "card" not in ch.mapping:
            if not self.wi_source:
                self.wi_source = "none"
                self.notes.append("该表既没有 WI 列也没有参数卡列，无法把行归到某个 WI")
            if not self.equip_source:
                self.equip_source = "none"
            return

        need_wi = not self.wi_source
        need_eq = not self.equip_source
        if not need_wi and not need_eq:
            return

        rel_wi, rel_eq = {}, {}
        sheet_wi, sheet_eq = "", ""
        for s in self.sheets:
            if s is ch or s.header_row == 0 or not s.header:
                continue
            if (not need_wi or sheet_wi) and (not need_eq or sheet_eq):
                break
            header = s.header
            card_cols = [c for c, n in enumerate(header)
                         if n is not None and _key(n) and re.search(r"参数卡\d*$", _key(n))]
            if not card_cols:
                continue
            wi_col = None
            if need_wi and not sheet_wi:
                wi_col = map_columns(header, None)[0].get("wi")
            eq_cols = _exact_field_cols(header, "equip") if (need_eq and not sheet_eq) else []
            if wi_col is None and not eq_cols:
                continue
            ws = wb[s.name]
            for r in ws.iter_rows(min_row=s.header_row + 1,
                                  max_col=len(header), values_only=True):
                for cc in card_cols:
                    key = _usable_join_key(r[cc] if cc < len(r) else None)
                    if not key:
                        continue
                    if wi_col is not None:
                        v = _usable_join_key(r[wi_col] if wi_col < len(r) else None)
                        if v:
                            rel_wi.setdefault(key, v)
                    for ec in eq_cols:
                        v = _usable_join_key(r[ec] if ec < len(r) else None)
                        if v:
                            rel_eq.setdefault(key, v)
            if wi_col is not None and rel_wi and not sheet_wi:
                sheet_wi = s.name
            if eq_cols and rel_eq and not sheet_eq:
                sheet_eq = s.name

        if need_wi:
            if rel_wi:
                self.card_wi = rel_wi
                self.wi_source = "relation:%s" % sheet_wi
                self.notes.append("采集项表无 WI 列，已用「%s」按参数卡名回填 WI（覆盖 %d 张参数卡）"
                                  % (sheet_wi, len(rel_wi)))
            else:
                self.wi_source = "none"
                self.notes.append("采集项表无 WI 列，且未找到可用的 WI 关联表")
        if need_eq:
            if rel_eq:
                self.card_equip = rel_eq
                self.equip_source = "relation:%s" % sheet_eq
                self.notes.append("采集项表无设备列，已用「%s」按参数卡名回填设备编号"
                                  "（覆盖 %d 张参数卡）—— 匹配据此收窄到本设备" % (sheet_eq, len(rel_eq)))
            else:
                self.equip_source = "none"
                self.notes.append("采集项表既无设备列，也没找到含「设备编号」的关联表"
                                  "—— 期望值只能给整份 PDF 的并集，判定会偏宽松")

    # ---------------------------------------------------------- 抽取
    def records(self, wi_name: str | None = None,
                sheet: str | None = None) -> list[dict]:
        """按列名抽取采集项。wi_name 为 None 时不过滤。

        同一筛选条件只读一遍工作簿（见 _rec_cache）—— 大表上一趟就是几秒、几十 MB。
        返回的列表是**共享的只读结果**：调用方请勿就地修改（compare 会自己复制一份）。
        """
        if sheet:
            self._choose(sheet)
        ch = self.chosen
        if ch is None or not ch.usable:
            raise SchemaError("；".join(self.errors) or "未能识别采集项明细表")
        ck = (wi_name, ch.name)
        if self._rec_cache and self._rec_cache[0] == ck:
            return self._rec_cache[1]
        m = ch.mapping
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        try:
            ws = wb[ch.name]
            ncol = max(m.values()) + 1
            child, anchors = merge_fill_maps(ws, ncol)
            anchor_val = {}   # 合并区左上角的真实值
            carry = {}        # 无合并声明时的「同上」顺延值
            out = []
            blank = 0
            for ri, r in enumerate(ws.iter_rows(min_row=1, max_col=ncol, values_only=True), 1):
                # 先记录合并区左上角的值（可能在表头行，也可能在数据行）
                for f in FILL_MERGE:
                    c = m.get(f)
                    if c is not None and (ri, c + 1) in anchors and c < len(r):
                        v = r[c]
                        if v is not None and str(v).strip():
                            anchor_val[(ri, c + 1)] = v
                if ri <= ch.header_row:
                    continue

                def g(f, _r=r, _ri=ri):
                    """取值：空则按「合并区左上角 → 同上」顺序继承（仅分组类字段）。"""
                    c = m.get(f)
                    if c is None or c >= len(_r):
                        return None
                    v = _r[c]
                    if v is not None and str(v).strip():
                        if f in FILL_MERGE:
                            carry[f] = v
                        return v
                    if f in FILL_MERGE:
                        a = child.get((_ri, c + 1))
                        if a is not None and a in anchor_val:
                            return anchor_val[a]
                        if f in FILL_CARRY and f in carry:
                            return carry[f]
                    return v

                item_raw = g("item")
                if item_raw is None or not str(item_raw).strip():
                    blank += 1
                    if blank >= 3:
                        carry.clear()   # 连续空行视为分段，清掉顺延值
                    if blank > 500:
                        break
                    continue
                blank = 0
                if str(item_raw).strip() == str(ch.header[m["item"]]).strip():
                    continue  # 重复表头
                cd = g("card")
                card_key = str(cd).strip() if cd is not None and str(cd).strip() else None
                wi = None
                if self.wi_source == "column":
                    w = g("wi")
                    wi = str(w).strip() if w is not None and str(w).strip() else None
                elif self.wi_source.startswith("relation:") and card_key:
                    wi = self.card_wi.get(card_key)
                if wi_name and wi != wi_name:
                    continue
                # 设备：优先用表内设备列；没有则用关联表按参数卡名回填的设备编号
                eq = None
                if self.equip_source == "column":
                    eq = g("equip")
                elif self.equip_source.startswith("relation:") and card_key:
                    eq = self.card_equip.get(card_key)
                name, unit = parse_item_name(item_raw)
                out.append({
                    # row / sheet：这条数据在 Excel 里的**行号**（界面左侧要显示，
                    # 方便按判定结果回表里定位）。1-based，与 Excel 状态栏一致。
                    "row": ri, "sheet": ch.name,
                    "card": cd, "proc": g("proc"), "group": g("group"),
                    "item": name, "unit": unit, "type": g("type"),
                    "lo": to_num(g("lo")), "hi": to_num(g("hi")),
                    # 精度签名（小数位数）：「3.0 与 3 不能算对」的判据。
                    # 只存原始字面，不存解析后的值 —— 解析后精度就没了。
                    "lo_dp": num_precision(g("lo")),
                    "hi_dp": num_precision(g("hi")),
                    # 显示上下限 / 允许超上下限：表里自带的两列开关。以前只在
                    # 自检里用过，判定层从没读 —— 界面上要拿它跟「按类型规矩
                    # 应该填什么」并排显示，填反了一眼就能看出来（2026-09-29）。
                    "show_lim": g("show_lim"), "allow_over": g("allow_over"),
                    "fixed": g("fixed"), "btrue": g("btrue"), "bfalse": g("bfalse"),
                    "fixed_dp": num_precision(g("fixed")),
                    "spec": g("spec"), "equip": eq, "wi": wi,
                })
            self._rec_cache = (ck, out)
            return out
        finally:
            wb.close()

    def wi_names(self):
        """表里出现的全部 WI（用于「找不到就报错」时给出提示）。"""
        ch = self.chosen
        if ch is None:
            return []
        if self.wi_source.startswith("relation:"):
            return sorted(set(self.card_wi.values()))
        if self.wi_source != "column":
            return []
        m = ch.mapping
        wb = openpyxl.load_workbook(self.path, read_only=True, data_only=True)
        try:
            ws = wb[ch.name]
            col = m["wi"]
            names = set()
            for r in ws.iter_rows(min_row=ch.header_row + 1, max_col=col + 1, values_only=True):
                v = r[col] if col < len(r) else None
                if v is not None and str(v).strip():
                    names.add(str(v).strip())
            return sorted(names)
        finally:
            wb.close()

    # ---------------------------------------------------------- 报告
    def report(self, verbose=True):
        L = []
        L.append("【Excel 结构自检】%s" % self.path.split("/")[-1])
        L.append("=" * 62)
        L.append("扫描到 %d 张 sheet：" % len(self.sheets))
        for s in self.sheets:
            tag = "（隐藏）" if s.hidden else ""
            if s.header_row == 0:
                L.append("  · %-14s %s未识别到表头，跳过" % (s.name, tag))
            else:
                ncol = sum(1 for c in s.header if c is not None and str(c).strip())
                L.append("  · %-14s %s表头 r%-3d 非空列 %-3d 识别字段 %d 个%s"
                         % (s.name, tag, s.header_row, ncol, len(s.mapping),
                            "  · 选为采集项表(分%d)" % s.score if s is self.chosen else ""))
        L.append("")
        ch = self.chosen
        if ch is None:
            L.append("✗ 未能识别采集项明细表")
            for e in self.errors:
                L.append("   " + e)
            return "\n".join(L)

        L.append("采集项表：%s（表头第 %d 行）" % (ch.label, ch.header_row))

        def _src_txt(src, col_label, none_label):
            if src == "none":
                return none_label
            if src.startswith("relation:"):
                return "关联表回填 → " + src.split(":", 1)[1]
            if src == "column":
                return col_label
            return "未识别"

        L.append("WI 来源：%s" % _src_txt(
            self.wi_source, "表内 WI 列",
            "无法确定（表内无 WI 列，也没找到关联表）"))
        L.append("设备来源：%s" % _src_txt(
            self.equip_source, "表内设备列",
            "无法确定（表内无设备列，也没找到含设备编号的关联表）"))
        L.append("")
        L.append("字段映射（按列名语义，非列号）：")
        for f, (_label, _syns) in FIELD_SPEC.items():
            if f in ch.mapping:
                c = ch.mapping[f]
                nm = ch.header[c]
                L.append("  %-10s ← 第 %-3d 列「%s」" % (_label, c + 1, nm))
        miss = [FIELD_SPEC[f][0] for f in FIELD_SPEC if f not in ch.mapping]
        L.append("  － 未匹配到：%s" % ("、".join(miss) or "（无）"))
        if verbose and ch.unknown:
            L.append("")
            L.append("已忽略的列（与本工具无关）：")
            L.append("  " + "、".join(str(ch.header[c]).strip() for c in ch.unknown[:24]))
        if ch.conflicts:
            L.append("")
            L.append("需留意的映射：")
            for c in ch.conflicts:
                L.append("  ! " + c)
        if self.notes:
            L.append("")
            for n in self.notes:
                L.append("· " + n)
        L.append("")
        if not ch.usable:
            L.append("✗ 缺少必需列：%s → 该表不能用于核对" %
                     "、".join(FIELD_SPEC[f][0] for f in CRITICAL_FIELDS if f not in ch.mapping))
        else:
            L.append("结构可用。可核对字段：%s" %
                     "、".join(FIELD_SPEC[f][0] for f in
                               ("lo", "hi", "fixed", "btrue", "bfalse") if f in ch.mapping))
        return "\n".join(L)


# ⚠️ 这里原本有一个 `def inspect(path, sheet=None)`（2026-09-30 重构时删除）。
# 它有三个问题：① 三处都查不到调用方（app.py / _regress / 引擎内部都没有）；
# ② 名字与标准库 `inspect` **撞名**，一旦本模块将来 `import inspect` 就会被静默
# 遮蔽，属于埋雷；③ 自检入口另有 `self_check` / `inspect_pdf` / `inspect_excel`。
# 真要用结构自检请走那三个。

# ==========================================================================
#  第 2 节  PDF 版式抽取
#  竖表 / 横表 / 程序附表三种版式 + 唯一取值解析器
# ==========================================================================

"""PDF 版式自适应层：识别作业指导书里的工艺参数表。

支持的版式：
  A 竖表   序号 | 参数名称 | 值            —— 参数名在「列」上（WI-0008 型）
  B 横表   参数 | Step 1 | Step 2 | …    —— 参数名在「行」上、Step 在列上（WI-0021 型）

B 型的关键特征（本模块识别的依据）：
  1. 表里存在一行，某几列的值是「参数」，且紧邻右侧是「Step 1 / Step 2 …」；
  2. 该行上方若干行是键值对形式的设备信息（激光机编号 / 程序号 / 焊接保护套 / 使用芯轴）；
  3. 该行下方每行 = 一个参数名 + 各 Step 的取值，值形如 `720~760mw`、`2.5mm`、`------`；
  4. 同一张 PDF 表格里可能并排两套参数子表，因此要把表格按「参数」列位置切开。

用法：
    rep = report_layout(pdf_path)        # 版式自检（可打印给用户看）
    recs = extract_stepwise(pdf_path)    # [{param,step,equip,lo,hi,unit,note,raw,…}]
"""


# --------------------------------------------------------------------------
#  版式注册表 —— 第 2 节的「总开关」，所有派发都由这张表驱动
#
#  为什么要有它：版式加到第十种之后，「谁先谁后、谁独占、出错了怎么办」全是
#  散在三个抽取函数里的 if / continue 硬编码 —— 加一种版式要改 3 处，顺序只有
#  读代码才知道。更糟的是 `extract_vertical` 用一个 try 把整页五个版式包在一起：
#  任何一个抛异常，该页剩下四个**全部静默跳过**，只留一条日志（损失被掩盖）。
#
#  有了这张表：
#    · 加一种版式 —— 写一个函数 + 往表里加一行，别处一行都不用动；
#    · 顺序       —— `order` 升序，同 channel 内生效（不再靠代码位置）；
#    · 独占       —— `exclusive=True` 认领后跳过该表后续版式（以前靠 continue）；
#    · 异常       —— 逐版式隔离，坏掉的那个单独记进 LAYOUT_ERRORS，不拖累同页；
#    · 可观测     —— LAYOUT_HITS 记命中条数，layout_registry_report() 直接给
#                    「版式自检」用，不必再另写一套硬编码 if；
#    · 灰度/回滚  —— `enabled=False` 单独停掉一种版式，不用改代码。
#
#  ⚠️ order 不是随便排的：**判据越宽，order 必须越大**。宽判据排在前面会
#     「抢走」窄判据的表。R48 实测代价：第十种的规格档一度认了纯数字，把
#     WI-0021 的步骤横表整张抢走 —— ok 18315 条里 5310 条凭空变 no_match。
#     改这张表之前先问一句：它会不会抢别人的表。
# --------------------------------------------------------------------------

LAYOUT_CHANNELS = ("stepwise", "vertical", "program")
LAYOUT_KINDS = ("table", "page")


@dataclass(frozen=True)
class LayoutSpec:
    """一种 PDF 版式的登记项。纯数据 —— 单测可以直接检查这张表本身。"""
    key: str                 # 唯一键，用于命中统计与日志
    label: str               # 中文名，给人看
    channel: str             # 归属抽取入口：stepwise / vertical / program
    kind: str                # 粒度：table（逐张表）/ page（逐页）
    order: int               # 同 channel 内的尝试顺序，升序
    fn: str                  # 解析函数名（延迟解析，这样表才能放在文件前部）
    no: int = 0              # 第 N 种版式；0 = 兜底或通道内不单独编号
    exclusive: bool = False  # 认领后跳过该表后续版式（只有 table 级有意义）
    needs: tuple = ()        # 除标准参数外还要从上下文取哪些键
    enabled: bool = True     # 开关：临时停掉某个版式做灰度 / 回滚
    hint: str = ""           # 判据说明 —— 改表时提醒「会不会抢别人的表」


LAYOUT_REGISTRY = (
    # —— stepwise 通道：逐张表，认领后可独占 ——
    LayoutSpec("enstep", "英文横表（Step 为行、参数为列）",
               "stepwise", "table", 10, "_enstep_table", no=9, exclusive=True,
               hint="要 ≥3 个 Step 行且 ≥2 种英文列名"),
    LayoutSpec("matrix", "矩阵表（参数名在首列、规格档在列头）",
               "stepwise", "table", 20, "_matrix_table", no=10, exclusive=True,
               hint="规格档必须带区间符或乘号；判据窄，排在 B 型之前"),
    LayoutSpec("stepwise", "B 型横表（参数名 | Step 列，兜底）",
               "stepwise", "table", 90, "_stepwise_fallback", needs=("pts",),
               hint="判据最宽，必须排最后，否则会抢走第九 / 十种的表"),
    # —— 曾经在这里的 `stepcol_steps`（同一种版式开第二扇门）已于 2026-11-11 删除 ——
    #
    # 它当初是为「案例6 WI-0029 有 35 条判成找不到设备」造的绕道：第 7 页三张表
    # 的设备号写在表上，步骤链却看不见。真实原因是 `_STEP_HEAD` 把「Step N」的
    # 结尾锚死了，带工序说明的列头（`Step 1 (Rotation Welding)`）一个都认不出，
    # 那三张表在横表通道里整块 0 条 —— 见 `_STEP_HEAD` 处的长注释。
    #
    # 判据修好后，同一份 PDF 的 43 条记录由横表版式正常读出，三张子表
    # （表2-i / 2-j / 2-k）各自认领自己的规格；WI-0029 的 51 条 no_match
    # 全部归零，全语料其余 19 份 WI 一条未动。
    #
    # 教训：绕道能绕过症状，也会把「为什么读不到」这个真问题埋起来 ——
    # 上次开这扇门时是「修 35 条、坏 8 条」，关掉又回到原点；这次修判据
    # 一步到位且零回归。**以后遇到「某块表整块抽不出来」，先查判据，别加通道。**
    # —— vertical 通道：逐页累加，互不独占 ——
    LayoutSpec("vertical", "竖表（参数名在列上）",
               "vertical", "page", 10, "_vertical_page", no=1),
    LayoutSpec("bare_kv", "无表头裸键值表（参数名 | 值）",
               "vertical", "page", 20, "_bare_kv_page", no=7),
    LayoutSpec("headrow_kv", "表头在首行、每行一档的横表",
               "vertical", "page", 30, "_headrow_kv_page", no=8),
    LayoutSpec("stepcol", "参数名在行 / 步骤在列",
               "vertical", "page", 40, "_stepcol_page", no=4),
    LayoutSpec("prose", "参数值写在正文叙述里",
               "vertical", "page", 50, "_prose_page", no=5, needs=("name2eq",)),
    # —— program 通道：程序参数附表 ——
    LayoutSpec("program", "程序参数附表（转置双列）",
               "program", "page", 10, "_program_page"),
)

# 命中统计与异常明细：每份 PDF 解析前清空（见 reset_pdf_failures）
LAYOUT_HITS = {}
LAYOUT_ERRORS = []


def reset_layout_stats():
    LAYOUT_HITS.clear()
    del LAYOUT_ERRORS[:]


def _layout_fn(name):
    """按名字取解析函数 —— 注册表放在文件前部（那时函数还没定义），故延迟解析。"""
    fn = globals().get(name)
    if not callable(fn):
        raise KeyError("版式注册表里的函数不存在或不可调用：%s" % name)
    return fn


def _tbox(t):
    """一张表在页面上的 bbox `(x0, top, x1, bottom)`，取不到给 None。

    只做一件事：把坐标四舍五入到 0.1pt 并转成普通元组。为什么要单独一个函数 ——
    `_layout_call` 只会给「逐张表」的版式补默认框，**逐页**的版式（竖表 / 步骤列 /
    裸键值 / 表头横表）在内部自己遍历表，得自己调这个才能填上「表格范围」。
    """
    try:
        bb = getattr(t, "bbox", None)
        if bb:
            return tuple(round(float(x), 1) for x in bb)
    except Exception:          # noqa: BLE001
        pass
    return None


def _layout_call(spec, ctx):
    """跑一个版式：异常隔离 + 命中计数 + 返回值归一。

    这三件事以前是散在各抽取函数里的重复代码，而且做得很糙：
      · 异常 —— `extract_vertical` 一个 try 包住整页五个版式，坏一个全页跳过；
      · 归一 —— 解析函数万一返回 None，`recs.extend(None)` 会直接炸，这里兜成 []；
      · 计数 —— 以前根本没有，版式命中情况只能靠猜。
    """
    if not spec.enabled:
        return []
    pno = ctx.get("pno")
    try:
        fn = _layout_fn(spec.fn)
        kw = {k: ctx.get(k) for k in spec.needs}
        if spec.kind == "table":
            out = fn(ctx.get("t"), ctx.get("rows"), pno, ctx.get("tno"), **kw)
        else:
            out = fn(ctx.get("page"), pno, **kw)
    except Exception as e:
        msg = "%s: %s" % (type(e).__name__, e)
        LAYOUT_ERRORS.append((spec.key, pno, msg))
        # 同时记进老通道，保证「单页失败不拖垮整份 PDF」这条老行为不变
        LAST_FAILED_PAGES.append(
            (pno, "版式「%s」抽取失败 %s" % (spec.key, msg)))
        return []
    if not out:
        return []
    LAYOUT_HITS[spec.key] = LAYOUT_HITS.get(spec.key, 0) + len(out)
    # 每条记录记住自己是从**哪种版式**来的。出处追溯以前只能追到「第几页第几张表」，
    # 追不到「按哪套判据读出来的」—— 人工复核时遇到可疑值，最想问的恰恰是这个。
    #
    # —— 来源坐标（2026-11-10 第一步）——
    # 同时给每条记录补上「这个值是从页面哪个格子读出来的」。**只加字段，不改判定**：
    # 判定链里没有任何一处读 `_box/_row/_col/_vcol/_blk`（可用 grep 复核）。
    #
    # 为什么非要有：案例6 那 35 条「找不到设备」的记录，人只能看到「第7页第N张表」，
    # 看不到它在表里的哪一行。而一张物理表里可能并排放着三张小表（WI-0029 p7），
    # 表外的标题到底该管哪一行 / 哪张小表，只能靠位置判断 —— 没有坐标就只能按
    # 「整行最近的标题」去猜，那正是窄验证失败的根子（35 条对上 23，却弄坏 8 条）。
    #
    # 字段（抽取函数自己写真实值；认不出的留 None，绝不编）：
    #   _box   这张表的 bbox (x0, top, x1, bottom)
    #   _row   表内行号（0 起，与 grid() / cell_boxes() 的 rows 同一套口径）
    #   _col   参数名所在列号
    #   _vcol  值所在列号（一行多个值格拼成 ' ; ' 串时为 None）
    #   _blk   同一张物理表内第几个逻辑块（0 起）—— 并排两张子表时靠它分开，
    #          只靠 (页, 表号) 分不开
    box = _tbox(ctx.get("t"))
    for r in out:
        if isinstance(r, dict):
            r.setdefault("_layout", spec.key)
            r.setdefault("_box", box)
            r.setdefault("_row", None)
            r.setdefault("_col", None)
            r.setdefault("_vcol", None)
            r.setdefault("_blk", None)
    return out


def _layouts_for(channel, ctx, kind=None):
    """按注册表顺序跑一个通道的全部版式；exclusive 的认领后不再往下试。

    kind：只要某一粒度（"table" / "page"）。一个通道里可能同时挂着两种粒度
    （2026-11-11：stepwise 通道里加了一个**页级**版式），不定向的话
    表级函数会拿不到 `t`/`rows`、页级函数会拿不到 `page`，两边都抛异常
    被 `_layout_call` 兜住返回空 —— 表面能跑，实则每个通道都在报假异常。
    """
    recs = []
    for spec in sorted((s for s in LAYOUT_REGISTRY if s.channel == channel),
                       key=lambda s: s.order):
        if kind and spec.kind != kind:
            continue
        got = _layout_call(spec, ctx)
        if got:
            recs.extend(got)
            if spec.exclusive:
                break
    return recs


def _layouts_for_table(t, rows, pno, tno, pts=None):
    """stepwise 通道（逐张表）。"""
    return _layouts_for("stepwise", {"t": t, "rows": rows, "pno": pno,
                                     "tno": tno, "pts": pts}, kind="table")


def _layouts_for_page(page, pno, channel="vertical", name2eq=None):
    """vertical / program 通道（逐页累加，互不独占）。"""
    return _layouts_for(channel, {"page": page, "pno": pno, "name2eq": name2eq},
                        kind="page")


def layout_registry_report():
    """版式自检：每种版式命中多少条、有没有抛异常、有没有被停用。"""
    L = ["【版式注册表自检】共登记 %d 种版式" % len(LAYOUT_REGISTRY)]
    for spec in sorted(LAYOUT_REGISTRY, key=lambda s: (s.channel, s.order)):
        n = LAYOUT_HITS.get(spec.key, 0)
        tag = "命中 %d 条" % n if n else "未命中"
        if not spec.enabled:
            tag = "已停用"
        errs = [e for e in LAYOUT_ERRORS if e[0] == spec.key]
        if errs:
            tag += "，异常 %d 次：%s" % (len(errs), errs[0][2][:40])
        L.append("  %-11s %-9s %-6s %-26s %s"
                 % (spec.key, spec.channel, spec.kind, spec.label[:26], tag))
    return L


# --------------------------------------------------------- pdfminer 兼容补丁

def _patch_pdfminer():
    """修 pdfminer.six 的一个已知崩溃点（会让整页、进而整个 WI 读不出东西）。

    有些 PDF 的 CIDSystemInfo 把 Registry/Ordering 写成了 PDF 关键字（如 /Adobe），
    pdfminer 却对它调 .decode()，而 PSKeyword 没有这个方法 →
        PdfminerException: 'PSKeyword' object has no attribute 'decode'
    WI-0017 第 3 页正是如此：一个坏页导致整个 WI 一条参数都抽不出来。
    这里给 PSKeyword 补一个 decode()（返回关键字名），不动用户的 venv。
    """
    try:
        from pdfminer.psparser import PSKeyword
    except Exception:
        return False
    if hasattr(PSKeyword, "decode"):
        return True

    def decode(self, *args, **kwargs):
        enc = args[0] if args else "utf-8"
        raw = str(getattr(self, "name", self))
        try:
            return raw.encode(enc, "ignore")
        except Exception:
            return raw.encode("utf-8", "ignore")

    PSKeyword.decode = decode
    return True


PATCHED_PDFMINER = _patch_pdfminer()

# 最近一次解析中失败的页码 [(页码, 原因)]：单页坏掉时如实报告，不假装整份都读到了
LAST_FAILED_PAGES = []


# --------------------------------------------------------------- 共享页面对象
#
# 一次比对要对同一份 PDF 做十来趟提取（设备号 / 横表 / 竖表 / 附表 / 元信息 /
# 文字质量…），每趟都 pdfplumber.open() 就会把每页的版面重新解析一遍。
# 实测这份 2.1MB / 6 页的 WI-0021：九趟提取 2.86s，其中绝大部分是重复解析；
# 而且每趟的临时对象把峰值内存推到 91MB。页面对象自己会缓存
# chars / lines / rects（pdfplumber 的 cached_properties），
# 复用同一批页面后：九趟 2.86s → 0.21s，峰值 91MB → 54MB。
#
# 只保留「当前这一份」，切换文件时把上一份关掉 —— 不把整批 PDF 攒在内存里。

_PAGE_CACHE = {"key": None, "pdf": None, "pages": None}


class _SharedPDF:
    """替代 pdfplumber 的 PDF 对象：只暴露调用方真正用到的 .pages。

    调用点清一色是 `with pdfplumber.open(p) as pdf: for page in pdf.pages`，
    所以这里只需要 .pages + 上下文管理器协议。
    """

    def __init__(self, pages):
        self.pages = pages

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False        # 真身由 _PAGE_CACHE 持有，这里不关

    def close(self):
        pass


def open_pdf(pdf_path):
    """按「路径 + 修改时间」复用同一份 PDF 的页面对象。

    注意：这是**进程内**的复用，页面对象会一直持有到换文件或调用
    release_pdf_cache()。长时间常驻的场景（如界面里跑完一次）应主动释放。
    """
    try:
        key = (os.path.abspath(pdf_path), os.path.getmtime(pdf_path))
    except OSError:
        key = (os.path.abspath(str(pdf_path)), 0)
    if _PAGE_CACHE["key"] == key and _PAGE_CACHE["pages"] is not None:
        return _SharedPDF(_PAGE_CACHE["pages"])
    release_pdf_cache()                      # 换文件：先把上一份关掉
    real = pdfplumber.open(pdf_path)
    try:
        pages = list(real.pages)
    except Exception:
        try:
            real.close()
        except Exception:
            pass
        raise
    _PAGE_CACHE.update(key=key, pdf=real, pages=pages)
    return _SharedPDF(pages)


def release_pdf_cache():
    """丢掉缓存的页面对象并关闭底层文件句柄。"""
    pdf, _PAGE_CACHE["pages"] = _PAGE_CACHE["pdf"], None
    _PAGE_CACHE["key"] = None
    _PAGE_CACHE["pdf"] = None
    if pdf is not None:
        try:
            pdf.close()
        except Exception:
            pass

# 「Step N」列头。**不能锚定结尾**（2026-11-11 修）：
# 实际表格的列头常常带工序说明 —— 「Step 1 (Rotation Welding)」
# 「Step 2 (Dynamic Welding)」，规范化后是 `step1(rotationwelding)`，
# 旧写法 `^step\s*(\d+)$` 要求数字后就是结尾，一个都匹配不上。
#
# 后果不是少抽一两个值，而是整块消失：`_step_head_cols` 认为这一行没有表头，
# `_stepwise_fallback` 于是整张子表 0 条 —— WI-0029 p5 右下（表2-d，
# T-1114/T-1281，2.0-6.0mm）就这样一条都没读，Excel 里该设备的 8 条采集项
# 全判「PDF 里没有这台设备的参数」；p6 右下（表2-h）同样
# （那台的 Ø1.5-4.0 档因此不存在，把整组规格门禁顶成了「对不上」）。
#
# `(?![0-9])` 替代 `$`：放宽尾部说明，但「Step 1」不会吃成「Step 10」。
_STEP_HEAD = re.compile(r"^step\s*(\d+)(?![0-9])", re.I)
_EQUIP = re.compile(r"T-\d+")
_META_KEYS = ("激光机编号", "程序号", "焊接保护套", "使用芯轴", "设备编号")

# 设备信息行（在子表标题与表头之间）绝不能当成参数行：
# 它的「值」是设备号/程序号，读进来就会污染期望值集合（曾把 T-1175 读成 1175）。
_META_NORM = ("激光机型号", "激光机编号", "程序号", "焊接保护套", "使用芯轴", "设备编号",
              "设备型号", "机台编号", "设备/工具", "设备工具",
              # 「机器编号」是最常见的一种写法（WI-0026 p5/p6 全部子表都用它），
              # 以前不在这张表里 → 它那一行会被当成参数行往下读。
              "机器编号", "机器型号")
_META_NORM_SET = frozenset(_META_NORM)

# 哪些「键: 值」是在声明**设备号**。
# 三条老判据（"激光机" / "设备编号" / "设备"）都漏了「机器编号」——
# WI-0026 p5 的 表1-a/1-b/1-d/1-e 因此设备号全空，整张表只是"被看见"、
# 却挂不到任何一台机器上 → 10800 条 no_equip_match。
# 同一份 PDF 里 表1-c 写的是「机器编号 T-1114（14）/T-1281（84）」，
# 表1-f 写的是「激光机编号 T-1362/…」—— 一份文件里混着两种写法，
# 所以必须按「键的语义」认，不能按某一处的字面。
_EQUIP_KEY_BAD = ("芯轴", "保护套", "物料", "程序号", "图号", "工装",
                  "夹具", "模具", "产品编号", "模板")
_EQUIP_KEY_RE = re.compile(r"(激光机|机台|机器|设备|设备/工具|设备工具)")


def _is_equip_meta(key, val):
    """这一行「键: 值」是不是在声明设备号。

    必须**同时**满足三条：
      ① 值里真的有 T-xxxx（"程序号 8" 这种没有设备号的，不可能是设备行）；
      ② 键是设备类（机器 / 机台 / 激光机 / 设备…）；
      ③ 键**不是**工装类（芯轴 / 保护套 / 物料 / 程序号 / 图号…）。
    第 ③ 条是硬要求：「焊接芯轴 T-2010-04」和「机器编号 T-1175」格式一模一样，
    只能靠标签区分 —— 第 28 轮已经在同一个坑上栽过一次（T-2015 芯轴被当设备）。
    """
    if not val or not _EQUIP.search(str(val)):
        return False
    kk = _n(key)
    if any(bad in kk for bad in _EQUIP_KEY_BAD):
        return False
    return bool(_EQUIP_KEY_RE.search(kk))

# 「T-1281：-1.3mm」这种「设备号：值」的复合单元格
_PREFIX_DEV = re.compile(r"^\s*(T-\d+(?:\s*/\s*T-\d+)*)\s*[（(]?\s*\d*\s*[）)]?\s*[:：]\s*")
# PDF 里折行的说明文字会有一部分落进取值格：`'Welding)\n480-600mw'`
# （WI-0023 表1-12 实测，上一行的 "(Rotation Welding)" 折成了两行）。
# 剥掉「开头一段不含数字、且以右括号收尾」的残留。要求必须以 `)` 结尾是有意的：
# 这样 `------`、`(Rotation`、「位置: 105mm; 速度: …」都不会被误剥。
_LEAD_JUNK = re.compile(r"^[^0-9]{1,60}?\)[\s\u3000\u00a0]*")


def _show_cell(cell):
    """单元格的**显示**原文：剥掉挤进来的上一行文字残留。

    只影响界面/导出里显示的原文，不参与判定 —— 剥干净了若是空的，
    就退回原样（`------`、`(Rotation` 这类本来就没有右括号，不受影响）。
    """
    s = _c(cell)
    return _LEAD_JUNK.sub("", s).strip() or s
# 子表标题：表1-5 设备/工具编号：T-1114/T-1281激光焊接参数表(1.5-5.5mm)
# 子表号的后缀**不一定是数字**：WI-0025 p5 写的是「表2-a / 表2-b / 表2-c / 表2-d」。
# 只认数字时这四张表的 sub 全是空，于是它们和同页其它表挤进 step_tables 的
# 同一个桶（键是 (sub, dia)）→ 步骤清单被并集、直径档串档。
_TITLE_SUB = re.compile(r"表\s*\d+\s*[-－]\s*[0-9A-Za-z]+")
_DIA_RANGE = re.compile(r"[（(]\s*[φΦ⌀Øø]?\s*(\d+(?:\.\d+)?)\s*[-~～]\s*(\d+(?:\.\d+)?)\s*mm\s*[）)]", re.I)
_DIA_SINGLE = re.compile(r"[（(]\s*[φΦ⌀Øø]?\s*(\d+(?:\.\d+)?)\s*mm\s*[）)]", re.I)


def _n(s):
    return re.sub(r"[\s\u3000\u00a0]+", "", str(s or "")).lower()


# ══════════════════════════════════════════════════════════════════════
# 符号登记表（第十一节新增，R47）
#
# 为什么要有这张表：同一个符号在 PDF 里有 N 种写法 ——
#     ±       U+00B1（Unicode）      WI-0020 / 2139 / 2347 / 2518
#     \uf0b1  Symbol 字体私有区      WI-0034 / WI-0036
#     ≥ ≥ ＞>  各种液晶 drones
# 以前这些知识**散落在八处正则里各自硬编码**，漏一种写法就悄悄变性状而不报错：
#   · `70\uf0b110s` 被读成单值 70，±10 蒸发 → Excel 填的 `60~80` 被判「应为 70」
#     （用户 2026-09-29 抓到：「你的标点符号是否又遗漏了？」）
#     —— 不是遗漏，是同一个符号换了种字体我就认不出来了。
#
# 表结构： (规范字形, 语义类别, [全部已知变体…，按出现概率排序])
#   · 规范字形：内部统一用这个，**不要再写裸字符**
#   · 语义类别：决定它参与什么运算（见下面的 `SYMBOL_OPS`）
#   · 全部变体：字体私有区 / 全角半角 / 中英文别名，全都列进来
#
# 加一种新写法 = 往这张表里加一个字符串，不用改任何正则。
# ══════════════════════════════════════════════════════════════════════
SYMBOL_SPECS = [
    # —— 公差 / 对称偏差 ——
    # \uf0b1 = Symbol 字体的 ±（WI-0034 / WI-0036 用的就是它）
    ("±", "tolerance", ["∓", "\uf0b1", "＋／－", "+/-"]),
    # —— 区间 ——
    # ⚠️ ASCII `-`（U+002D）一身三职：区间（`16s-20s`）、负号（`-1.5mm`）、
    # 连字符（设备号 `T-128`）。**登记在表内**（值解析要靠它认区间），但列入
    # `_SYM_NO_REPLACE` —— 不做全局替换。否则 `T-128` 变 `T~128`，`for T-128`
    # 剥不掉，`-1.5 mm for T-128` 就被读成 1.5~128（实测踩到过）。
    ("~", "range",     ["～", "〜", "-", "—", "–", "－"]),
    # —— 单边不等式 ——
    (">=", "ge",       ["≥", "≧", "\uf0b2", "＞＝"]),
    ("<=", "le",       ["≤", "≦", "\uf0b3", "\uf0a3", "＜＝"]),
    # ⚠️ 中文/英文的**词**（大于、以上、不低于…）不在这张表里 —— 它们是词不是符号，
    # 全局替换会把 `90以上` 改成 `90>`，把人能读懂的原文改坏。词只由 `_OP_RE`
    # 那一支**识别**（不改写），见下面的 `SYMBOL_WORDS`。
    (">",  "gt",       ["＞"]),
    ("<",  "lt",       ["＜"]),
    # —— 等于 / 约等 ——
    ("=",  "eq",       ["＝"]),
    ("≈",  "approx",   ["≒", "≃"]),
    # —— 单位符号 ——
    ("°",  "degree",   ["˚", "º", "\uf0b0"]),
    ("μ",  "micro",    ["µ", "\uf06d"]),
    # ⚠️ Æ/æ（U+00C6/U+00E6）看着是字母，实为某些字体把 Ø 编错的结果 ——
    # WI-0036 表1 标题写 `Æ4.0~Æ5.0`，不认它那张表的直径档就一直是 None。
    #
    # ⚠️ \uf0c6 = Symbol 字体的 Ø（Symbol 编码 0xC6），与 R47 用户点名的
    # `\uf0b1`（Symbol 的 ±）是同一类坑。原表只登了 `\uf0a6`（那是 Symbol 的
    # 另一个字符，不是 Ø）—— 差一位，于是 WI-0036 表1 标题
    # `直径为\uf0c64.0~\uf0c65.0球囊成型参数设定` 的直径档**一直解析不出来**：
    # dia=None → 该表拿不到规格分 → φ4.0-5.0 的卡被拿 φ6.0 档（表2）的值当期望值，
    # 凭空造出 10 条「不该填 N/A」。
    ("Ø",  "diam",     ["ø", "⌀", "∅", "Φ", "φ", "\uf0a6", "\uf0c6", "Æ", "æ"]),
    ("²",  "square",   ["^2"]),
    ("³",  "cube",     ["^3"]),
    # —— 乘除 ——
    ("×",  "times",    ["✕", "\uf0b4", "＊"]),
    ("÷",  "divide",   ["∕", "⁄"]),
]


# 登记在 `SYMBOL_SPECS` 里、**但不做全局替换**的写法。
# 只登记这一类：该字符在文本里还有别的身份，全局替换会改坏原文（见 range 的 `-`）。
_SYM_NO_REPLACE = {"-"}

# 只**识别不改写**的词：它们是自然语言，不是符号（见上面 gt / lt 那条警告）。
# 格式：{词: 语义类别}。只喂给 `_OP_RE`（判「这一格有没有比较约束」）。
SYMBOL_WORDS = {
    "大于": "gt", "超过": "gt", "高于": "gt", "以上": "gt", "不低于": "gt",
    "小于": "lt", "低于": "lt", "以下": "lt", "以内": "lt", "不高于": "lt",
    "大于等于": "ge", "不小于": "ge",
    "小于等于": "le", "不大于": "le",
}


def _build_symbol_trans():
    """由 `SYMBOL_SPECS` 生成两张表：单字符 → translate；多字符串 → 一次正则。

    单字符变体（`\\uf0b1`、`～`、`－`）走 `str.translate`，因为全引擎每格都在过，
    必须快。多字符变体（`+/-`、`大于`、`不低于`）translate 管不了，另给一条正则。

    ⚠️ 多字符串必须**长的先替换**：'+/-' 要在 '-' 之前处理，否则被切成 '+/'。
    """
    single, multi = {}, []
    for canon, _kind, variants in SYMBOL_SPECS:
        for v in variants:
            if not v or v == canon or v in _SYM_NO_REPLACE:
                continue
            if len(v) == 1:
                single.setdefault(ord(v), canon)
            else:
                multi.append((v, canon))
    multi.sort(key=lambda kv: -len(kv[0]))
    pattern = None
    if multi:
        pattern = (re.compile("|".join(re.escape(v) for v, _c in multi)),
                   dict(multi))
    return single, pattern


_SYMBOL_TRANS, _SYMBOL_MULTI = _build_symbol_trans()

# 「这段文本里到底有没有需要归一的符号」—— 由 SYMBOL_SPECS 自动生成，
# 不手写。表加了新变体，快速路径自动跟上，不会出现「登记了却不生效」。
#
# ⚠️ 单字符变体收进字符类、多字符变体（'+/-'、'＋／－'）只能走 alternation ——
# 只收单字符的话，'+/-' 这类因为三个字符都不在类里而被快速路径**整个跳过**，
# 多字符替换永远不执行（实测 `10+/-1` 归不出 `±`）。
def _build_sym_any():
    one, many = [], []
    for _canon, _k, vs in SYMBOL_SPECS:
        for v in vs:
            if not v or v in _SYM_NO_REPLACE:
                continue
            (one if len(v) == 1 else many).append(v)
    parts = (["[%s]" % "".join(re.escape(v) for v in one)] if one else []) \
        + [re.escape(v) for v in many]
    return re.compile("|".join(parts)) if parts else re.compile(r"(?!x)x")


_SYM_ANY = _build_sym_any()


def _sym(s):
    """把符号的各种写法折成**规范字形**（`\\uf0b110s` → `±10s`）。

    必须做在**最底层**：说不清这些字符会从哪一格冒出来（温度单位、公差、
    直径符号都有）。做在 `_c` / `_norm` / `_clean` / `_flat` 这四个全引擎都在用的
    归一化入口上，一处生效、到处受益 —— 不用在每个解析点各补一遍。

    ⚠️ 曾经这里有两份 `_sym`（R47 加多字符变体时另起了一份，旧的没删），
    后定义的那个静默生效、先定义的那个变成死代码 —— 改错了地方就会以为
    「表改了没反应」。2026-09-29 清理，只留这一份。
    """
    if not isinstance(s, str):
        return s
    # 快速路径：绝大多数串里没有需要归一的符号，别为它付正则 + translate 的钱
    # （全引擎每格都过这里，全量跑批 6 万行，能省就省）。
    if not _SYM_ANY.search(s):
        return s
    # ⚠️ 顺序不能反：先换**多字符**变体（`+/-`、`大于`），再换单字符。
    # 反过来的话 `+/-` 里的 `-` 会先被折成 `~`，整条就再也匹配不上了
    # （实测 `10+/-1` → `10+/~1`，凭空造出个 ~）。
    if _SYMBOL_MULTI:
        pat, m = _SYMBOL_MULTI
        s = pat.sub(lambda mo: m[mo.group(0)], s)
    if _SYMBOL_TRANS:
        s = s.translate(_SYMBOL_TRANS)
    return s



def symbol_class(s):
    """这段文本里出现的符号语义类别集合（{'tolerance', 'gt', …}）。

    ⚠️ 规范字形自己也要查 —— 只查变体的话 `symbol_class("70±10s")` 会查不到
    （`±` 是规范字形，不在变体列表里）。
    """
    out = set()
    s = _sym(s or "")
    for canon, kind, variants in SYMBOL_SPECS:
        for v in [canon] + list(variants):
            if v and v in s:
                out.add(kind)
                break
    for w, kind in SYMBOL_WORDS.items():
        if w in s:
            out.add(kind)
    return out


def _symbol_class_re(*kinds):
    """按语义类别取正则字符类 —— 写正则时用它，别再手抄字符集。

    ⚠️ 只收**单字符**变体（多字符串的走 `_sym` 那层已折成规范字形），
    每个字符单独 `re.escape`，否则 `-` 在字符类里会变成范围、`]` 会提前闭合。
    """
    seen, uniq = set(), []
    for canon, kind, variants in SYMBOL_SPECS:
        if kind not in kinds:
            continue
        for v in [canon] + list(variants):
            if v and len(v) == 1 and v not in seen:
                seen.add(v)
                uniq.append(re.escape(v))
    return "[" + "".join(uniq) + "]" if uniq else "(?!)"


# 由符号表生成的两个字符类 —— `parse_step_value` 与 `_apply_symbols` 用。
_SYM_TOL_RE = _symbol_class_re("tolerance")
_SYM_RANGE_RE = _symbol_class_re("range")


# 符号 → 运算：规范字形背后那件事到底怎么算。
# 以前这段运算躲在 `parse_step_value` 的一串 if 里（"1) 对称公差 / 2) 非对称
# 公差 / 3) 只给上偏差 / …"），加一种符号要同时改正则和改 if —— 两边不同步就出错。
# 现在把「符号语义」和「怎么展开成区间」并列声明：
SYMBOL_OPS = {
    # 语义类别: {"shape": 这个符号推出来的值是什么形态, "desc": 人话解释}
    "tolerance": {"shape": "range",     "desc": "对称公差 ±b → 中心值 ± b 展开成区间"},
    "range":     {"shape": "range",     "desc": "a~b → 区间 [a, b]"},
    "ge":        {"shape": "one_sided", "desc": "≥a → 只有下限，上限不限"},
    "le":        {"shape": "one_sided", "desc": "≤a → 只有上限，下限不限"},
    "gt":        {"shape": "one_sided", "desc": ">a → 大于，只有下限"},
    "lt":        {"shape": "one_sided", "desc": "<a → 小于，只有上限"},
    "eq":        {"shape": "single",    "desc": "=a → 单值"},
    "approx":    {"shape": "single",    "desc": "≈a → 约等于，按单值核"},
}


# 夹在**两个汉字之间**的空白 = PDF 折行产物，不是词边界（见 `_c`）。
_CJK_GLUE_RE = re.compile(r"(?<=[一-鿿])[ \t\r\n\u3000\u00a0]+(?=[一-鿿])")


def _c(s):
    """清洗一个单元格的原文。

    除了把空白压成单空格，还要把**汉字之间**的空白接回去：单元格里的中文被排版
    折行时会从词中间断开（「持续时⏎间」→「持续时 间」），而这个空白并不是词边界
    —— 中文不用空格分词。不接回去，「持续时间」就永远匹配不上，那条记录从
    vindex 里整个消失。WI-0004 表1-b 实测代价：它自己的 3.0s 因此落选，
    候选里只剩旁边表1-c 的 3.0-4.0，一对正确的数据被报成「不一致」。

    英文单词之间的空格必须保留（"Welding Time" 拆开就没法按词匹配），
    所以限定两侧都必须是汉字。

    ⚠️ 这里是**全引擎最底层的单元格文本入口**，`_sym`（Symbol 字体 → Unicode）
    必须做在这一层 —— 做在上一层就会漏：`parse_step_value` 走的是这里，
    实测 `70\\uf0b110s` 在 `±` 归一化之前就已被吃成一个 70。
    """
    return _CJK_GLUE_RE.sub("", re.sub(r"\s+", " ", _sym(str(s or ""))).strip())


def _is_blank(s):
    return not s or set(s) <= set("-–—~/　 \t")


# 元信息格子里常把「直径范围：」也一起写进来（如「φ2.0-5.5mm： SPC2000019-01」），
# 这不是我们想要的值，剥掉。
_META_PREFIX = re.compile(
    r"^[φΦ⌀Øø]?\s*\d+(?:\.\d+)?(?:\s*[-~～]\s*\d+(?:\.\d+)?)?\s*mm\s*[:：]\s*")


def _clean_meta(v):
    return _META_PREFIX.sub("", _c(v)).strip()


def parse_step_value(text):
    """解析精确路径单元格的值（**精确路径唯一的取值解析器**）。

    注意：竖表 / 程序附表走的是 `_extract_all_values`（候选值解析器），
    二者**粒度不同、不要合并** —— 实测合并会让 WI-3852 出现 2 条假不一致，
    理由见 `_extract_all_values` 的 docstring。

    '720~760mw'          → (720.0, 760.0, 'mw', '')
    '------'             → None（无值）
    '1.8mm'              → (1.8, 1.8, 'mm', '')
    '-1.5 mm for T-128'  → (-1.5, -1.5, 'mm', 'T-128')
    'T-1281：-1.3mm'     → (-1.3, -1.3, 'mm', 'T-1281')

    最后一种必须特别小心：一个单元格里写着「设备号：值」，若直接取第一个数字，
    会把设备号里的 1281 当成值（并且因为前面的负号变成 -1281）。所以先剥前缀。

    另外三条防线：
      · 剥掉所有 `T-xxxx` 之后若一个数字都不剩（如整格就是 `T-1175`、`T-1362/T-10161`），
        说明这格是设备编号而不是参数值 → 返回 None；
      · `1200±50` 这类公差写法按 `(1150, 1250)` 展开，避免只取到 1200 而漏判；
      · `16s-20s` 这类「单位夹在两个数字中间」的区间要认出来 —— 只取到 16 就会把
        合法的 18 判成不一致（假不一致）。

    2026-09-21 补齐（用户提的三类「识别不准」）：
      · **负号**：`-0.8mm` 必须是 -0.8。以前另一套解析器把负号吃掉变成 0.8，
        而「移动距离/位置」这类参数经常是负值 —— 符号丢了等于把对的说成错。
      · **非对称公差**：`14+0.5/-0 atm`（WI-0010 测试压力的真实写法）→ (14.0, 14.5)；
        `5+0.2-0.1` → (4.9, 5.2)。这叫「识出了符号和数字却没做加减法」。
      · **上一行文字挤进本格**：PDF 里折行的 `(Rotation Welding)` 会有一部分落进
        取值格，抽出来是 `'Welding)\\n480-600mw'`（WI-0023 表1-12 实测）。
        数字虽还能读对，但显示出来的原文是错的，先剥掉这段残留。
    """
    s = _c(text)
    if _is_blank(s):
        return None
    note = ""
    m = _PREFIX_DEV.match(s)
    if m:
        note = m.group(1)
        s = s[m.end():].strip()
    m = re.search(r"(?:for|适用|对应)\s*([A-Za-z0-9\-（）()/、,.]+)", s)
    if m:
        note = (note + " " + m.group(1)).strip()
        s = (s[:m.start()] + s[m.end():]).strip()
    if _is_blank(s):
        return None
    s = _LEAD_JUNK.sub("", s).strip()        # 剥掉挤进来的上一行文字
    if _is_blank(s):
        return None
    if not re.search(r"\d", re.sub(r"T-\d+", "", s)):
        return None                      # 整格只有设备编号，不是参数值
    unit = ""
    # 单位取「第一个以字母开头的连续串」。必须显式以字母/℃/°/百分号开头是有意的：
    #   · 旧写法 `[a-zA-Z%°/²]+` 会把 `14+0.5/-0 atm` 里的 `/` 当成单位（读到 '/'），
    #     于是压掉 atm；
    #   · 而 ℃ 根本不在字母类里，`240-250℃` 的单位一直是空的 —— 精确路径自己
    #     另写了一套（按结尾取单位）才没暴露，两套一统一就会露出来。
    um = re.search(r"[a-zA-Z℃°%][a-zA-Z%°/²³℃]*", s)
    if um:
        unit = um.group(0)
    # ⚠️ 从这里往下全部走**统一的符号运算表**（`_apply_symbols`），不再自己硬编码
    # 字符集。以前 `±` 只写了 U+00B1 一种，Symbol 字体的 `\uf0b1` 就整性状丢失
    # （`70\uf0b110s` → 70，±10 蒸发）；`~` 只写了 `~～` 没写全角 `－`。
    # 符号的**字形**归 `SYMBOL_SPECS` 管，**运算**归 `SYMBOL_OPS` 管，这里只查表。
    r = _apply_symbols(s)
    if r is not None:
        lo, hi = r
        return (lo, hi, unit, note)
    # 兜底：单值（保留正负号）
    m = re.search(r"(?<![A-Za-z0-9])(-?\d+(?:\.\d+)?)", s)
    if not m:
        return None
    lo = hi = float(m.group(1))
    return (lo, hi, unit, note)


def _apply_symbols(s):
    """按 `SYMBOL_SPECS` / `SYMBOL_OPS` 把一段文本展开成 (lo, hi)；认不出返回 None。

    每一条分支都对应 `SYMBOL_OPS` 里的一个语义类别 —— 加一种符号只要往
    `SYMBOL_SPECS` 里加变体，运算这块不用动。
    """
    _N = r"(-?\d+(?:\.\d+)?)"
    # 1) 对称公差（tolerance）：1200±50 → 1150~1250；裸 ±0.05 → -0.05~0.05
    m = re.search(_N + r"\s*" + _SYM_TOL_RE + r"\s*(\d+(?:\.\d+)?)", s)
    if m:
        base, t = float(m.group(1)), float(m.group(2))
        return (base - t, base + t)
    m = re.match(r"\s*" + _SYM_TOL_RE + r"\s*(\d+(?:\.\d+)?)", s)
    if m:
        t = float(m.group(1))
        return (-t, t)
    # 2) 非对称公差：base +上/-下 与 base +上-下
    m = re.search(_N + r"\s*\+\s*(\d+(?:\.\d+)?)\s*/\s*-\s*(\d+(?:\.\d+)?)", s)
    if m:
        base, up, low = float(m.group(1)), float(m.group(2)), float(m.group(3))
        return (base - low, base + up)
    m = re.search(_N + r"\s*\+\s*(\d+(?:\.\d+)?)\s*-\s*(\d+(?:\.\d+)?)", s)
    if m:
        base, up, low = float(m.group(1)), float(m.group(2)), float(m.group(3))
        return (base - low, base + up)
    # 3) 只给上偏差：10+0.5 → 10~10.5
    m = re.match(r"\s*" + _N + r"\s*\+\s*(\d+(?:\.\d+)?)\s*[a-zA-Z%°/²]*\s*$", s)
    if m:
        base, up = float(m.group(1)), float(m.group(2))
        return (base, base + up)
    # 4) 区间（range，容忍单位夹在中间：16s-20s）
    m = re.search(_N + r"\s*[a-zA-Z%°/²]*\s*" + _SYM_RANGE_RE + r"\s*(-?\d+(?:\.\d+)?)", s)
    if m:
        lo, hi = float(m.group(1)), float(m.group(2))
        if lo > hi:
            lo, hi = hi, lo
        return (lo, hi)
    return None


def parse_sub_title(title):
    """从子表标题里取出 (子表编号, 直径范围)。

    '表1-5设备/工具编号：T-1114/T-1281激光焊接参数表(1.5-5.5mm)'
        → ('表1-5', (1.5, 5.5))
    '表1-2，设备/工具编号：T-1175激光焊接参数表(6.0mm)'  → ('表1-2', (6.0, 6.0))

    直径范围是与 Excel「适用规格」对齐的钥匙：同一台设备有多套参数时，
    区分它们的正是这个范围（而不是设备号）。
    """
    flat = re.sub(r"[\s\u3000\u00a0]+", "", str(title or ""))
    sub = ""
    m = _TITLE_SUB.search(flat)
    if m:
        sub = m.group(0)
    m = _DIA_RANGE.search(flat)
    if m:
        return sub, (float(m.group(1)), float(m.group(2)))
    m = _DIA_SINGLE.search(flat)
    if m:
        d = float(m.group(1))
        return sub, (d, d)
    # 标题里的规格**常常不带 mm**：WI-0025 p5 右半张写的是
    # 「表2-b，设备/工具编号：T-1175激光焊接参数表(4.5-6.0)」，
    # 而左半张带 mm（「…参数表(1.5-4.0mm)」）。只认带 mm 的那一种，右半张的
    # 直径档就是 None —— 后果有两层：
    #   · 判定层：4.5~6.0 这台机器被当成「只有 1.5~4.0 一档」，
    #     于是 Excel 里 4.5/5.0/6.0 的规格全部对不上，整组参数被判「无对应」；
    #   · 索引层：几张 dia=None 的表挤进同一个桶，步骤清单被并集。
    # 这里用**标题专用**的宽松口径（_title_dias）兜底，它只在子表标题这种
    # 上下文里用，不会把数据行里的数字当成直径档。
    d2 = _title_dias(flat)
    if d2:
        return sub, tuple(d2[0])
    # 单值档同样可能不带 mm，而且是**小数**才敢认：
    # 「表2-h…激光焊接参数表(6.0)」「表2-j…(4.5)」。整数括号（程序号之类）
    # 一律不认，避免把「(8)」当成直径。
    # 少了这一档，PDF 里 6.0 / 4.5 那张表等于不存在 —— Excel 侧 6.0、4.5 的
    # 适用规格会被判成「对不上任何一档」，整组参数被门禁拦下（那是假漏判）。
    m = re.search(r"[（(]\s*[φΦ⌀Øø]?\s*(\d+\.\d+)\s*[）)]", flat)
    if m:
        d = float(m.group(1))
        if 0.05 <= d <= 60:
            return sub, (d, d)
    return sub, None


def _last_sub_title(text):
    """从「一个 cell 里塞了上下相邻好几张子表」的文本里，取**最后一张**子表的
    (子表编号, 直径范围)。

    为什么需要：pdfplumber 有时把上下相邻的两张表（表7-a Ø2.0-3.0mm /
    表7-b Ø3.25-5.0mm）并成一个 table，且把两张表的标题+数据全挤进同一个 cell。
    这时 `parse_sub_title` 只认第一个标题（表7-a），而当前 table 的数据其实
    是**最后一张**（表7-b，其数据行紧跟在这个 cell 之后）。取最后一个标题
    才能把直径档（3.25-5.0）还给这批数据 —— 否则表7-b 的值会错误沿用
    表7-a 的 (2.0,3.0)，与 Excel 的 3.25-5.0 档对不上（WI-0006 p11/p12 实测）。

    判据：文本里出现 ≥2 个「表N-x」标题。少于 2 个不接管（返回 (None, None)），
    调用方照旧走 parse_sub_title / scope_of 的老路，不改变任何现有行为。
    """
    flat = re.sub(r"[\s\u3000\u00a0]+", "", str(text or ""))
    titles = _TITLE_SUB.findall(flat)
    if len(titles) < 2:
        return None, None
    last = titles[-1]
    # 从最后一个标题处截到文末，用 parse_sub_title 解析这段的编号与直径
    pos = flat.rfind(last)
    seg = flat[pos:] if pos >= 0 else flat
    return parse_sub_title(seg)


def _step_head_cols(row):
    """该行里所有「参数表头」的列索引（值含「参数」且右侧紧邻 Step N）。"""
    out = []
    for c, v in enumerate(row or []):
        k = _n(v)
        if not k or "参数" not in k or k in ("参数卡", "参数值"):
            continue
        if c + 1 < len(row) and _STEP_HEAD.match(_n(row[c + 1])):
            out.append(c)
    return out


def _step_width(row, c):
    """子表宽度 = 1（参数名列） + 连续 Step 列数。"""
    w = 1
    while c + w < len(row) and _STEP_HEAD.match(_n(row[c + w])):
        w += 1
    return w


def _has_step_head_in(row, c0, c1):
    """只在自己这一格子表范围内判断，避免被并排的邻表打断。"""
    return any(c0 <= cc < c1 for cc in _step_head_cols(row))


def _is_title_row(row, c0, c1):
    """该行是不是「另一套子表的标题行」。

    一张 PDF 表格里可能上下堆着好几套子表（p5 就有 表1-1/1-2 与 表1-3/1-4 两段）。
    扫描数据行时必须在这里停住，否则上面那套会把下面那套的
    「激光机型号 T-1175」「程序号 46」当成自己的参数值读进来（曾因此把 1175 当成激光强度）。
    """
    txt = " ".join(_c(row[k]) for k in range(c0, min(c1, len(row))))
    if "设备/工具" in txt:
        return True
    return bool(_TITLE_SUB.search(re.sub(r"[\s\u3000\u00a0]+", "", txt)))


def _equip_set_spec(s):
    return sorted(set(_EQUIP.findall(str(s or ""))))


# 后缀不一定是数字（「表2-a」），括号里的规格也不一定带 mm（「(4.5-6.0)」「(6.0)」）。
# 只认「数字后缀 + 带 mm」时，页面文本这一路（_page_titles）会整条漏掉 ——
# 而表格行那一路常常丢字（p8「表2-m…参数表(4.5」后面的 `-6.0)` 落在表格外，
# 拼出来的标题是半截），两条路都断就只剩「没有规格档」这一种结果。
_TITLE_PAT = (r"表\s*\d+\s*[-－]\s*[0-9A-Za-z]+[^\n]*?参数表\s*[（(]\s*[φΦ⌀Øø]?\s*\d+(?:\.\d+)?"
              r"\s*(?:[-~～]\s*\d+(?:\.\d+)?)?\s*(?:mm)?\s*[）)]")


def _page_titles(page):
    """在页面文本里找出所有子表标题（带坐标）。

    为什么不能只读表格行：有些子表的标题行在表格抽取时会丢字
    （例如 p8 的「表1-13，T-1362/T-1」被吃掉，只剩后半截），
    而页面文本始终是完整的，且带位置，可以按横坐标认回是哪一套子表。
    """
    out = []
    try:
        out = page.search(_TITLE_PAT, regex=True) or []
    except Exception:
        out = []
    return out


def _cell_at(row, c):
    """取该行第 c 列的格子坐标。

    合并单元格在该行的 cells 里是一串 None，所以不能「先过滤 None 再按位置取」——
    那样索引会整体错位（WI-0021 p5 的右半边子表就是这样丢掉了直径范围）。
    这里按列号直接索引；遇到 None 再找最近的可用格。
    """
    try:
        cells = list(row.cells)
    except Exception:
        return None
    if not cells:
        return None
    if 0 <= c < len(cells) and cells[c]:
        return cells[c]
    for d in range(1, len(cells)):
        for j in (c + d, c - d):
            if 0 <= j < len(cells) and cells[j]:
                return cells[j]
    return None


def _title_at(page_titles, table, i, c, nxt):
    """按子表在表头行的横向范围，从页面标题里认回本子表的标题。

    x 范围 = 本子表「参数」列的左边界 → 右邻子表「参数」列的左边界
    （没有右邻子表时取整行最右边界）。标题中心落在这个区间里的才算自己的。
    """
    try:
        row = table.rows[i]
    except Exception:
        return ""
    b0 = _cell_at(row, c)
    if not b0:
        return ""
    x0, ytop = b0[0], b0[1]
    b1 = _cell_at(row, nxt) if nxt is not None else None
    if b1:
        x1 = b1[0]
    else:
        xs = [x[2] for x in (row.cells or []) if x]
        x1 = max(xs) if xs else x0 + 200
    best = None
    for m in page_titles:
        cx = (m.get("x0", 0) + m.get("x1", 0)) / 2.0
        if x0 - 8 <= cx <= x1 + 8 and m.get("bottom", 0) <= ytop + 3:
            if best is None or m.get("bottom", 0) > best.get("bottom", 0):
                best = m
    return best.get("text", "") if best else ""


def extract_stepwise(pdf_path: str) -> list[dict]:
    """抽取全部「参数 | Step…」横表，返回记录列表。

    每条记录：page / table / sub(子表编号) / dia(直径范围) / equip(设备集合) / equip_raw
              / prog(程序号) / sleeve(焊接保护套) / meta / param / step
              / lo / hi / unit / note(单元格里指定的设备) / raw
    """
    recs = []
    global LAST_FAILED_PAGES
    LAST_FAILED_PAGES = []
    with open_pdf(pdf_path) as pdf:
        for pno, page in enumerate(pdf.pages, 1):
            try:
                pts = _page_titles(page)
                tables = page.find_tables() or []
            except Exception as e:                      # 单页解析失败不拖垮整份 PDF
                LAST_FAILED_PAGES.append((pno, "%s: %s" % (type(e).__name__, e)))
                continue
            for tno, t in enumerate(tables):
                try:
                    rows = t.extract() or []
                except Exception as e:
                    LAST_FAILED_PAGES.append(
                        (pno, "表格提取失败 %s: %s" % (type(e).__name__, e)))
                    continue
                # 版式派发：按 `LAYOUT_REGISTRY` 里 stepwise 通道的顺序认
                # （第九种 → 第十种 → B 型兜底）。order 小的先试；认领且
                # exclusive 的那个不再往下试。顺序与独占以前靠循环里的
                # continue 硬编码，现在由注册表说了算。
                recs.extend(_layouts_for_table(t, rows, pno, tno, pts))
    return recs



def _stepwise_fallback(t, rows, pno, tno, pts=None):
    """B 型横表兜底：参数名在行、Step 在列（WI-0021 型）。

    ⚠️ 判据最宽 —— 只要某一列写着「参数」、右边紧跟 Step 就算命中，
    所以必须排在 `LAYOUT_REGISTRY` 里 stepwise 通道的**最后**；排到前面
    会把第九 / 十种版式的表抢走（R48 实测过代价：ok 18315 条里 5310 条
    凭空变 no_match）。以前这个顺序是靠循环里的 `continue` 表达的，
    现在由注册表的 `order` 说了算。
    """
    recs = []
    pts = pts or []
    for i, row in enumerate(rows):
        heads = _step_head_cols(row)
        for c in heads:
            w = _step_width(row, c)
            steps = []
            for k in range(1, w):
                m = _STEP_HEAD.match(_n(row[c + k]))
                steps.append(int(m.group(1)) if m else k)
            nxt = next((h for h in heads if h > c), None)
            win = max(c + w, nxt) if nxt is not None else len(row)

            # 子表标题（子表编号 / 直径范围）：先按位置从页面文本取，
            # 取不到再退回「表头行上方那一行」
            title, fallback = "", ""
            for j in range(i - 1, -1, -1):
                r2 = rows[j]
                if _has_step_head_in(r2, c, c + w):
                    break
                txt = " ".join(_c(r2[k]) for k in range(c, min(win, len(r2))))
                if not txt.strip():
                    continue
                flat = re.sub(r"[\s\u3000\u00a0]+", "", txt)
                # 子表标题的判据：带「设备/工具编号」，**或**写成
                # 「表1-5，T-1362/T-10058/T-10161尖端成型参数表」这种
                # 「表号 … 参数表」的样子。
                # WI-0024 的表1-1/1-2/1-4/1-5 既没有「设备/工具编号」，
                # 标题里也没有直径档 —— 旧判据两条都不沾，整张子表认不出
                # 编号（sub 为空），后果有两层：
                #   · 界面上的出处只能写「未标号」，看不出是哪张表；
                #   · 多张表挤进 step_tables 的同一个桶（键是 (sub, dia)），
                #     步骤清单被并集 —— 表1-5 只有 Step1，却因为与有 Step2
                #     的表同桶，被判成「有这一步」，Excel 的 N/A 报不出来。
                if "设备/工具" in txt or ("参数表" in txt
                                          and _TITLE_SUB.search(flat)):
                    title = txt
                    break
                if not fallback and "参数表" in txt and (
                        _DIA_RANGE.search(flat) or _DIA_SINGLE.search(flat)):
                    fallback = txt
            cand_pos = _title_at(pts, t, i, c, nxt)
            sub, dia, used = "", None, ""
            for cand in (cand_pos, title, fallback):
                if not cand:
                    continue
                s2, d2 = parse_sub_title(cand)
                if d2:
                    sub, dia, used = s2, d2, cand
                    break
                if not used:
                    sub, dia, used = s2, d2, cand

            # 设备信息：表头行上方、同一列区的键值对
            meta, equip_raw = {}, ""
            for j in range(i - 1, -1, -1):
                r2 = rows[j]
                if c >= len(r2):
                    break
                if _has_step_head_in(r2, c, c + w):
                    break
                if _is_title_row(r2, c, c + w):
                    break       # 到本套子表的标题为止，再往上就不是自己的信息了
                k = _c(r2[c])
                if not k:
                    continue
                v = _clean_meta(r2[c + 1]) if c + 1 < len(r2) else ""
                if v and k not in meta:
                    meta[k] = v
                    if k in _META_KEYS and "激光机" in k:
                        equip_raw = v
            if not equip_raw:
                # 按**语义**认领设备行（见 `_is_equip_meta`）。
                # 从下往上扫、取最近的一条：表头正上方那一行才是本子表的
                # 设备声明，再往上是程序号/保护套/芯轴。
                equip_raw = next((v for kk, v in meta.items()
                                  if _is_equip_meta(kk, v)), "")
            sub_equip = _equip_set_spec(equip_raw)

            last_param = None
            for j in range(i + 1, len(rows)):
                r2 = rows[j]
                if c >= len(r2):
                    break
                if _has_step_head_in(r2, c, c + w):
                    break
                # 往下扫也可能撞见「另一套子表」的标题行：到这里必须停。
                # p5 一张表里上下堆了 表1-1/1-2 与 表1-3/1-4 两段，不拦就会把下段的
                # 「激光机型号 T-1175」「程序号 46」当成上段的参数值读进来。
                # 这里用整行判断是安全的：表头行以下不会再有并排兄弟表的标题。
                if _is_title_row(r2, 0, len(r2)):
                    break
                pname = _c(r2[c])
                if _STEP_HEAD.match(_n(pname)) or _n(pname).startswith("("):
                    continue          # 「(Dynamic/Welding…)」这类 Step 说明行
                if _n(pname) in _META_NORM_SET:
                    continue          # 设备信息行，不是参数
                if pname:
                    last_param = pname
                if not last_param:
                    continue
                # 整行空（参数名格 + 所有 Step 格都是空的）是表格里的空白行 /
                # 版式留白，不是「明确声明此处无值」—— 不能记，否则凭空多一堆
                # 「该参数该步骤为空」的假证据。
                if _is_blank(_c(r2[c])) and all(
                        _is_blank(_c(r2[c + k]))
                        for k in range(1, w) if c + k < len(r2)):
                    continue
                for k in range(1, w):
                    if c + k >= len(r2):
                        continue
                    cell = r2[c + k]
                    pv = parse_step_value(cell)
                    if not pv:
                        # 明确写不适用（N/A）→ 与 `------` 同属「表格在声明此处无值」，
                        # 但依据要说对话（不能把 N/A 说成「表格里写的是 ------」），
                        # 所以单独打标记 —— 判据在第 2.7 节 `pdf_cell_is_na`。
                        _na = pdf_cell_is_na(cell)
                        # 有内容但解析不出值（整格只有设备编号等）→ 跳过。
                        if not _na and not _is_blank(_c(cell)):
                            continue
                        # 明确写空（`------` / `—` / 空白）→ **必须记下来**。
                        # 它不是「没这个格子」，而是表格在声明「此处无值」。
                        # 丢了它，比对层就永远分不清两件完全不同的事：
                        #   · PDF 该步骤是空的  → Excel 填 N/A、或干脆不填，都对
                        #   · PDF 该步骤有值    → Excel 不填就是**漏填**，必须报错
                        # （WI-0021 表1-1「移动距离」就是 `-1.6mm | ------ | 1.6mm`，
                        #   旧版把中间那格整个丢掉，于是 N/A 只能落进「未填值」。）
                        recs.append({
                            "page": pno, "table": tno,
                            "sub": sub, "dia": dia, "title": used,
                            "equip": sub_equip, "equip_raw": equip_raw,
                            "sub_equip": sub_equip,
                            "prog": _c(meta.get("程序号", "")),
                            "sleeve": _c(meta.get("焊接保护套", "")),
                            "meta": meta, "param": last_param,
                            "step": steps[k - 1],
                            "lo": None, "hi": None, "unit": "", "note": "",
                            "raw": _show_cell(cell), "empty": True, "na": _na,
                            # 来源坐标：第 j 行「参数名在 c 列、值在 c+k 列」
                            "_row": j, "_col": c, "_vcol": c + k,
                        })
                        continue
                    lo, hi, unit, note = pv
                    # 单元格里写了「T-1281：-1.3mm」→ 该条只属于 T-1281
                    own = sorted(set(_EQUIP.findall(note))) if note else []
                    recs.append({
                        "page": pno, "table": tno,
                        "sub": sub, "dia": dia, "title": used,
                        "equip": own or sub_equip, "equip_raw": equip_raw,
                        "sub_equip": sub_equip,
                        "prog": _c(meta.get("程序号", "")),
                        "sleeve": _c(meta.get("焊接保护套", "")),
                        "meta": meta, "param": last_param,
                        "step": steps[k - 1],
                        "lo": lo, "hi": hi, "unit": unit, "note": note,
                        "raw": _show_cell(cell), "empty": False,
                        # 来源坐标：第 j 行「参数名在 c 列、值在 c+k 列」
                        "_row": j, "_col": c, "_vcol": c + k,
                    })
    return recs

# ==================================================== 第九种版式：英文横表（Step 为行、参数为列）
#
# 长这样（WI-0020 rev12 p12）：::
#
#     ['Step #', 'Setting Parameters', '', '', '', '', '', 'Exit Conditions']
#     ['', 'Temp. (℃)', 'Pres. (Bar)', 'Position (mm)', '', 'Water Cool', 'Purge', '']
#     ['', '', '', 'Left', 'Right', '', '', '']
#     ['Step 1', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A', 'N/A']
#     ['Step 5', '115（ref）', '30', '0', '0', 'N/A', 'N/A', '50s （ref）']
#
# 与 B 型（参数名在行、Step 在列）正好互为转置。B 型的判据是「某列写着『参数』、
# 右边紧跟 Step 1/Step 2」，这里 col0 才是 Step、参数名在**列头**上 —— B 型一条
# 都认不出，于是 228 条记录的 param 全变成 'Step 3'/'Step 4'…（《项目笔记 §三十二》
# 实测：486 行、一致 0 / 无法匹配 191）。**漏的是版式，不是字** —— 这份 PDF 的
# 文字层是完整的，pdfplumber 一次就把三张表的网格抽全了。
#
# 判据必须窄，否则会抢走别的版式：
#   1. 连续 ≥3 行的 col0 是 'Step N'；
#   2. 表头块（最多往上 3 行）里能认出 ≥2 个不同的英文列头。
# 实测 12 份语料 **0 页**误触发（改完当场扫过一遍）。
#
# 产出的记录与 extract_stepwise 同构，另带三个字段：
#   no_equip  True —— 这类表**根本没写设备号**，收窄链不能因为对不上设备号就判
#                     no_equip_match（那是把「PDF 没写」当成「PDF 没有」）
#   cond      分支生效条件（'2.0-3.5 &4.0*80-150'），一格多分支时用来挑
#   layout    'enstep'，便于自检与界面显示出处
#

# 列头 → 参数名。顺序有意义：先认更具体的，最后才轮到 position。
# （'Setting Parameters' 这种组头文字里没有具体关键词，不会误命中。）
_EN_COL_KINDS = (
    ("water", r"water\s*cool|水冷|冷却"),
    ("purge", r"purge|排气"),
    ("temp", r"temp|温度"),
    ("pres", r"pres|压力"),
    ("exit", r"exit|退出"),
    ("pos", r"position|位置"),
)

_EN_COL_UNIT_FALLBACK = {"temp": "℃", "pres": "Bar", "pos": "mm"}


def _en_col_kind(text):
    """列头文字（组头行+单位行+左右行拼起来）→ (类别, 单位)。认不出返回 (None, '')。"""
    s = str(text or "")
    unit = ""
    m = re.search(r"[（(]\s*([^（）()]{1,10}?)\s*[）)]", s)
    if m:
        unit = m.group(1).strip()
    low = s.lower()
    for kind, rx in _EN_COL_KINDS:
        if re.search(rx, low):
            return kind, (unit or _EN_COL_UNIT_FALLBACK.get(kind, ""))
    return None, ""


def _en_lr(block, c):
    """列头块里这一列写的是 Left 还是 Right（'Position (mm)' 拆成左右两列时用）。"""
    for r in block:
        if c < len(r):
            v = _n(r[c]).lower()
            if v == "left":
                return "L"
            if v == "right":
                return "R"
    return ""


def _en_param_of(kind, lr, raw, unit):
    """(列类别, 左右, 原值, 列单位) → (参数名, 单位)。

    Exit Conditions 是**复合列**：同一格里有时写 '3s'、有时写 '>90℃'，
    Excel 那边分成了「退出条件时间(sec)」和「退出条件加热中间模(℃)」两项 ——
    只能按值里带的单位把它拆开，否则两项会撞在同一个键上。
    """
    if kind == "water":
        return "水冷", ""
    if kind == "purge":
        return "排气中间模", ""
    if kind == "temp":
        return "加热中间模", unit or "℃"
    if kind == "pres":
        return "压力", unit or "Bar"
    if kind == "pos":
        if lr == "L":
            return "左拉伸距离", unit or "mm"
        if lr == "R":
            return "右拉伸距离", unit or "mm"
        return "位置", unit or "mm"
    if kind == "exit":
        if re.search(r"℃|度", str(raw or "")):
            return "退出条件加热中间模", "℃"
        return "退出条件时间", "sec"
    return None, ""


# --------------------------------------------------------------------------
#  英文横表的「一格多量」（第 9.5 节）
#
#  PDF 一格常把同一步的**几个部位**并排写：
#      「中间模：Same as step5 / 左端模：110-140 / 右端模：110-140」
#      「中间模：120±10 / 端模：110-140」
#  而 Excel 给每个部位各建了一个采集项（加热中间模 / 加热左端模 / 加热右端模）。
#
#  不拆的后果（WI-0036 实测，两条）：
#    · 表4!119「加热中间模/6」—— 整格被当成中间模的值，于是**左/右端模的
#      110-140 混进应为值**，Excel 填的 100~110 被判成不一致（**假错**）；
#    · 「加热左端模 / 加热右端模」在 PDF 里找不到自己的格子 → 20 条 no_match，
#      依据写「参数名未收录」。
#
#  竖表早就有这套（`_vertical_expand` 的 `_V_ATTR_RE`），英文横表这边一直没做 ——
#  同一个病根，所以这里同样做成一张表，别再散成 if。
# --------------------------------------------------------------------------

# 部位标签表：(PDF 里的写法, 规范部位名)。加部位只改这一张表。
EN_PART_LABELS = (
    ("中间模", "中间模"),
    ("左端模", "左端模"),
    ("右端模", "右端模"),
    ("端模", "端模"),          # 「端模：110-140」= 左右两端同值
    ("上夹钳", "上夹钳"),
    ("下夹钳", "下夹钳"),
)

# ⚠️ 「端模」是「左端模」「右端模」的子串 —— 不加 `(?<![左右])` 的话，
#    「左端模：」会被拆成「左」+「端模：」，部位名当场错乱。
_EN_PART_RE = re.compile(
    r"(?<![左右])(" + "|".join(w for w, _c in EN_PART_LABELS) + r")\s*[:：]")


def _enstep_split_parts(raw):
    """「中间模：X / 左端模：Y / 右端模：Z」→ [('中间模', 'X'), ('左端模', 'Y'), …]。

    认不出返回 []（调用方保持原样，绝不猜）。三种情况不拆，与竖表口径一致：
      · 只出现一个部位（那本来就是单值）
      · 同一部位出现两次（那是别的结构，不是一格多量）
      · 某个标签后面没有值
    """
    s = str(raw or "")
    hits = list(_EN_PART_RE.finditer(s))
    if len(hits) < 2:
        return []
    names = [h.group(1) for h in hits]
    if len(set(names)) != len(names):
        return []
    out = []
    for i, h in enumerate(hits):
        end = hits[i + 1].start() if i + 1 < len(hits) else len(s)
        seg = s[h.end():end].strip().strip("/／；;，,、 ")
        if not seg:
            return []
        out.append((h.group(1), seg))
    return out


def _en_part_param(pname, label):
    """把参数名里的部位词换成这个标签（加热中间模 + 左端模 → 加热左端模）。

    参数名里认不出任何部位词就返回 None —— 认不出就别乱改名，按原样收。
    """
    s = str(pname or "")
    for w, _c in EN_PART_LABELS:
        if w in s:
            return s.replace(w, label)
    return None


# `(?<![a-z])for(?![a-z])` 而不是 `\bfor\b`：PDF 里常写成 '24bar for4.0*10-60'
# （for **后面**紧跟数字），`\b` 在 r 与 4 之间不成立 → 整格切不开，24 被当成 10-60。
# 反过来 **前面**也可能紧跟数字：`30for 2.75-5.0mm`（WI-0033 POT）里 `\b` 在 0 与 f
# 之间同样不成立 → 整格当单值，一格里的两套规格值永远分不开。
# 两侧的 `\b` 一律换成「不是字母」——`Performance` 这类仍被两道闸门挡住。
# ⚠️ `for` 前后都可能是数字（'30for 2.75-5.0mm'），也可能**紧跟单位**
# （'22barfor3.0'、'15sfor4.0' = bar/s + for + 直径）—— 那时前面是字母，
# `(?<![a-z])` 会把它挡掉（R47 实测 WI-0034 Step15：整条分支条件丢失）。
# 所以补一支：`for` 后面**直接是数字**也算 —— 'before' 后面不会跟数字。
_EN_FOR_RE = re.compile(r"(?<![a-z])for(?![a-z])|(?<=[a-z])for(?=\d)", re.I)


# 公差写法：`160+5/-0`（上偏差 +5、下偏差 -0）。斜杠在这里**不是**分支分隔符，
# 切之前先整段保护起来（换成占位符，切完还原）—— 逐字符判「斜杠前后是 +/-」不够：
# 这里的斜杠前面是数字 5、后面是 -0，两头都够不着。
_EN_TOL_RE = re.compile(r"([+-]\s*\d+(?:\.\d+)?)\s*/\s*([+-]\s*\d+(?:\.\d+)?)")


def _enstep_split(s):
    """按 `/` 把一格切成多段，但**不切**公差与 N/A 里的那个斜杠。

    WI-0033（POT —— 一张表里按规格分叉，没有表1/表2/表3）实测三个坑：
      · `160+5/-0`  —— 公差里的 `/` 被当分隔符，整格切成 `150` / `160+5` / `-0 for…`，
                       于是这一支的值变成 `-0`（应为 `160+5/-0`，条件 `2.25-3.0mm`）。
                       Excel 那行填的是 150~155（正是 `150+5/-0`），凭空被判不一致。
      · `N/A / 26bar for…` —— `N/A` 被切成 `N`、`A` 两段垃圾（条件还被填成了 `A`）；
      · `30for 2.75-5.0mm / 15for 2.25-2.5mm` —— 这个 `/` 才该切。
    """
    s = _EN_TOL_RE.sub(lambda m: m.group(1) + "\x00" + m.group(2), s)
    out, buf = [], ""
    for i, ch in enumerate(s):
        if ch == "/":
            prev = s[i - 1] if i else ""
            nxt = s[i + 1] if i + 1 < len(s) else ""
            if prev.upper() == "N" and nxt.upper() == "A":
                buf += ch                      # 不适用：N/A
                continue
            out.append(buf)
            buf = ""
            continue
        buf += ch
    out.append(buf)
    out = [x.replace("\x00", "/").strip() for x in out]
    return [x for x in out if x] or [s]


def _enstep_branches(raw):
    """把一格里的多分支切成 [(值文本, 生效条件)]。

    '24bar for4.0*10-60 / 22bar for 2.0-3.5 / &4.0*80-150'
        → [('24bar', '4.0*10-60'), ('22bar', '2.0-3.5 &4.0*80-150')]
    没有 'for' 就是单值，原样返回。
    """
    s = re.sub(r"\s+", " ", str(raw or "")).strip()
    if not _EN_FOR_RE.search(s):
        return [(s, "")]
    out = []
    for p in _enstep_split(s):
        m = _EN_FOR_RE.search(p)
        if m:
            out.append((p[:m.start()].strip(), p[m.end():].strip()))
        elif out:
            out[-1] = (out[-1][0], (out[-1][1] + " " + p).strip())
        else:
            out.append((p, ""))
    return out or [(s, "")]


# ── 「等同于第 N 步」的写法（2026-09-29 第 44 轮补）────────────────────
# PDF 里同一个意思有四种写法，都要认：
#     Same as step 4  /  Same step 4  /  Samestep4  /  Step 4
# ⚠️ `Samestep4` 这种**连写**是真实存在的、不是笔误：pdfplumber 抽文字时词间
# 空格会丢。实测 WI-0032（Atropos 球囊成型）P9/P10 三张表的 Pres. 列整列写成
# `Samestep4` / `Samestep6` / `Samestep7`，而旧正则只认带 `as` 的那种 →
# 压力列一个值都解不出来 → Excel 的「压力/5、7、8、9、10」全被打成「无法匹配」。
#
# 用户口径（2026-09-29 原话）：「这个 same 其实可以视为忽略……same 可有可无」
# —— 即 `Same` 与 `as` 都是**可省前缀**，数字还允许写成英文单词（`four`）。
# ⚠️ 正则**必须锚在开头**（`match` 而非 `search`）：出口条件那一列的
# `>Step5-5℃` 是「大于 Step5 温度减 5℃」，与「等于第 5 步」完全不是一回事 ——
# 锚定后它仍走老路（parse_step_value），不会被这条新规则吃掉。
_EN_NUM_WORD = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
    "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11,
    "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15,
}
_ENSTEP_REF_RE = re.compile(
    r"^(?:same)?\s*(?:as)?\s*step\s*#?\s*"
    r"(\d+|one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve"
    r"|thirteen|fourteen|fifteen)\b",
    re.I)
# 引用后面还可能跟偏移：`Step5+5`（WI-0032 表1 的 Step4 温度就是这么写的）
_ENSTEP_OFFSET_RE = re.compile(r"^\s*([+\-＋－])\s*(\d+(?:\.\d+)?)")


def _enstep_resolve_all(raw, cells, pname, depth=0):
    """把一格解析成 [(lo, hi, unit, note, cond), ...] —— 一格可能解出**好几条**。

    ⚠️ 引用式（`Same step4`）必须把**被引步骤的每一条分支都带出来**，各自继承
    自己的生效条件（R47）。旧口径只取第一条，于是 WI-0033 Step5 压力
    （`Same step4`）永远得到 Step4 的第一支 15bar —— 而 Step4 那一格其实是
    `15for 2.25-2.5mm / 30for 2.75-5.0mm` 两支，φ2.75 以上的四张卡明明填 30
    却被判成应为 15。
    """
    if depth > 3:
        v = parse_step_value(raw)
        return [(v[0], v[1], v[2], v[3], "")] if v else []
    s = str(raw or "").strip()
    m = _ENSTEP_REF_RE.match(s)
    if m:
        tok = m.group(1).lower()
        n = int(tok) if tok.isdigit() else _EN_NUM_WORD.get(tok)
        if n is None:
            return []
        mo = _ENSTEP_OFFSET_RE.match(s[m.end():])
        d = 0.0
        if mo:
            d = float(mo.group(2)) * (1 if mo.group(1) in "+＋" else -1)
        out = []
        for praw, _punit, pcond in (cells.get((pname, n)) or []):
            for lo, hi, u, note, c in _enstep_resolve_all(
                    praw, cells, pname, depth + 1):
                out.append(((None if lo is None else lo + d),
                            (None if hi is None else hi + d), u, note,
                            c or pcond))
        return out
    # 'SL1+30ref' / 'SR1+30 ref' 这类：引用的是**另一个我们不知道的量**（SL1）。
    # parse_step_value 会把 '30' 抠出来当值 —— 那是凭空造「应为」，比不认更糟。
    # 认不出来就如实返回 None（该格不参与判定），绝不猜。
    if re.match(r"^[A-Za-z]+\s*\d*\s*[+\-＋－]\s*\d", s):
        return []
    v = parse_step_value(s)
    return [(v[0], v[1], v[2], v[3], "")] if v else []


def _enstep_ref_via(raw, cells, pname, depth=0):
    """引用式（`Step5+5` / `Same as step5`）顺着引用链看：**被引用的那一格**有没有 ref。

    返回 True / False（认不出、或链条断了 → False）。

    用户 2026-09-30 定的口径（原话）：
      「**WI-0020 的 step5 内有 ref，所以 step5+5 也属于是 ref**，所以是浮点是是，
       要空着；**WI-0038 内 step5 没有 ref 字眼，所以 step5+5 就是正常浮点
       是否，要填写**」

    即：引用式**自己不表态**，它的性质**完全由被引用的那一格**决定：
      · 被引用格写了 ref  → 本格也是 ref（值列空着、两个开关都「是」）
      · 被引用格没写 ref  → 本格按解出来的值算（普通浮点，要填值）

    实测对照（两份球囊、同一个 `Step5+5`，填法相反，靠这一条才分得开）：
      WI-0020 表1  Step5 = `115（ref）`  → `Step5+5` 是 ref  → Excel 空着 ✓
      WI-0038 表5  Step5 = `85-95`       → `Step5+5` 是普通值 → Excel 填 90~100 ✓

    ⚠️ 与 `_inherit_ref`（「二次拉伸」继承上一步）是**两条独立的规矩**：
      那条看的是 Purpose 列写明「二次拉伸」，这条看的是格子里写了引用式。
      两条都为 True 时结论一致（都是 ref），不冲突。

    ⚠️ 链式引用要一路追下去（`Same as step5` 而 step5 自己也是 `Same as step4`），
      但与 `_enstep_resolve_all` 同一口径：`depth > 3` 收手防环。
    """
    if depth > 3:
        return False
    s = str(raw or "").strip()
    # 本格自己就写了 ref → 不用再往下追
    if _has_ref(s):
        return True
    m = _ENSTEP_REF_RE.match(s)
    if not m:
        return False                        # 不是引用式 → 本格自己说了算
    tok = m.group(1).lower()
    n = int(tok) if tok.isdigit() else _EN_NUM_WORD.get(tok)
    if n is None:
        return False
    for praw, punit, _pcond in (cells.get((pname, n)) or []):
        if _has_ref(praw) or _has_ref(punit):
            return True
        if _enstep_ref_via(praw, cells, pname, depth + 1):
            return True                     # 链式引用：继续往上游追
    return False


def _enstep_resolve(raw, cells, pname, depth=0):
    """把一格的值解析成 (lo, hi, unit, note) —— 只要第一条。

    PDF 里常见两种**引用写法**，直接交给 parse_step_value 会读出荒唐的值：
      'Same as step 7' → (7, 7, 'Same')      —— 当成压力 7 Bar
      'Step5+5'        → (5, 5, 'Step')      —— 当成温度 5 ℃
    必须在同表、同参数上把引用解出来（限递归深度，防死循环）。

    引用写法见上面 `_ENSTEP_REF_RE`：`Same` / `as` 可省、可连写、数字可英文。
    要拿全部分支请用 `_enstep_resolve_all`。
    """
    v = _enstep_resolve_all(raw, cells, pname, depth)
    return v[0][:4] if v else None


_EN_COND_RANGE_RE = re.compile(
    r"^\s*(\d+(?:\.\d+)?)\s*[-~～]\s*(\d+(?:\.\d+)?)\s*(?:mm|cm)?\s*$")


def _enstep_cond_score(cond, spec):
    """分支条件与 Excel「适用规格」的贴合度 0~1。

    条件写的是**直径区间**（'2.25-2.5mm' / '3.25-5.0mm'）时按**区间覆盖**打分：
    规格整段落进条件区间 = 1.0，一点不沾 = 0.0。

    ⚠️ 旧口径是「条件里的数字字面有没有出现在规格里」—— 'φ3.25-4.0' 对
    '2.75-5.0mm' 得 0 分（3.25、4.0 都不等于 2.75、5.0）。两臂同分 → 下游取并集
    → Excel 明明按规格填对了也报不一致。WI-0033 实测：Step10 加热中间模
    （160~165 for 2.25-3.0 / 150~155 for 3.25-5.0）与 Step14 压力
    （26bar for 2.25-4.0 / 24bar for 4.5-5.0）四张卡全被判错，改完全对。

    认不出区间（条件里数字不止两个，如 '2.0-3.5 &4.0*80-150'）就退回旧口径，
    老语料不受影响。
    """
    s = str(cond or "")
    nums = re.findall(r"\d+(?:\.\d+)?", s)
    sr = _spec_range(spec)
    if sr and nums:
        slo, shi = sr
        if len(nums) == 2:
            m = _EN_COND_RANGE_RE.match(s)
            if m:
                lo, hi = float(m.group(1)), float(m.group(2))
                if hi < lo:
                    lo, hi = hi, lo
                if shi > slo:
                    return max(0.0, min(hi, shi) - max(lo, slo)) / (shi - slo)
                return 1.0 if lo - 1e-9 <= slo <= hi + 1e-9 else 0.0
        elif len(nums) == 1:
            v = float(nums[0])
            return 1.0 if slo - 1e-9 <= v <= shi + 1e-9 else 0.0
    cn = nums
    if not cn:
        return 0.0
    sn = set(re.findall(r"\d+(?:\.\d+)?", str(spec or "")))
    if not sn:
        return 0.0
    return sum(1 for x in cn if x in sn) / float(len(cn))


def _enstep_detect(rows):
    """这张表是不是第九种版式？是就返回 (表头块起始, 首个数据行, [数据行号])，否则 None。"""
    if not rows:
        return None
    step_rows = [i for i, r in enumerate(rows)
                 if r and _STEP_HEAD.match(_n(r[0]))]
    if len(step_rows) < 3:
        return None
    d0 = step_rows[0]
    if d0 == 0:
        return None
    h0 = max(0, d0 - 3)
    block = rows[h0:d0]
    ncol = max((len(r) for r in rows), default=0)
    kinds = set()
    heads = []
    for c in range(ncol):
        txt = " ".join(_c(r[c]) for r in block if c < len(r))
        kind, unit = _en_col_kind(txt)
        if kind:
            kinds.add(kind)
            heads.append((c, kind, unit, _en_lr(block, c)))
        else:
            heads.append(None)
    if len(kinds) < 2:
        return None
    return h0, d0, step_rows, heads


def _enstep_title(t):
    """表上方最近的『表N…』标题（表号与规格档都写在它里面）。"""
    try:
        lines = t.page.extract_text_lines() or []
        y0 = t.bbox[1]
    except Exception:
        return ""
    return _title_above(lines, y0, None)


def _enstep_notes(t, steps, pno, tno, sub, dia, title):
    """表**下方备注**里的取值。

    WI-0020 三张表下面都有一行：
      「备注：…拉伸速度依据球囊长度及拉伸距离而定…；充气速率：一般为0.2-15之间。」
    Excel 为**每一步**都建了「充气速度/N」采集项（填 0.2~15），而表里根本没有这一列
    —— 不做正文取值，这 72 条就永远是「PDF 里找不到」。

    备注是对**整张表**说的，所以按这张表出现的每一步各记一条，并在 note 里写明
    出处是表下方备注 —— 不冒充表里印着的值。
    """
    out = []
    try:
        lines = t.page.extract_text_lines() or []
        y1 = t.bbox[3]
    except Exception:
        return out
    for ln in lines:
        top = ln.get("top")
        txt = ln.get("text") or ""
        if top is None or top < y1 - 6 or top - y1 > 140:
            continue
        if "备注" not in txt and "充气" not in txt:
            continue
        m = re.search(r"充气(?:速率|速度)[^0-9]{0,10}(\d+(?:\.\d+)?)\s*[-~～到]\s*"
                      r"(\d+(?:\.\d+)?)", txt)
        if not m:
            continue
        lo, hi = float(m.group(1)), float(m.group(2))
        if lo > hi:
            lo, hi = hi, lo
        for st in steps:
            out.append({
                "page": pno, "table": tno, "sub": sub, "dia": dia, "title": title,
                # ⚠️ **故意不带设备号**：备注写的是「充气速率：一般为 0.2-15」，
                # 那是对**整张表**说的通用口径，不是某一批设备的专属参数。
                # 第 53 轮试过把标题里的设备号挂上去，结果 card 0012（T-1177）
                # 在设备对齐那一层被剔掉，从 `ok/strict`（唯一确定）降级到
                # `ok/fb_kw`（兜底匹配）—— 判定强度反而变弱（实测 12 条）。
                # `no_equip=True` 的正确含义正是「这条不参与设备过滤」。
                "equip": [], "equip_raw": "",
                "sub_equip": [], "prog": "", "sleeve": "",
                "meta": {}, "param": "充气速度", "step": st,
                "lo": lo, "hi": hi, "unit": "",
                "note": "表下方备注：充气速率一般为 %g-%g" % (lo, hi),
                "raw": txt.strip(), "empty": False,
                "no_equip": True, "layout": "enstep_note", "cond": "",
            })
    return out


# ── 「二次拉伸」= 上一步工艺的延续（2026-09-29 第 46 轮，用户口径）─────────
#
# 用户原话：「正常来说是不一致，应该为浮点 是 否 1-30，但是这里特殊，这个属于
# 二次拉伸……所以出现二次拉伸，就默认是上一步延展，上一步有 ref 就是 ref。」
#
# 实测 WI-0032 表3（φ7.0~10.0）第 10 页：
#     Step6 125 32  [50-165 (REF) | 50-165 (REF)]  >110℃(ref)  Blowballoon.吹制球囊
#     Step7 125 Samestep6 [1-30 | 1-30]            >115℃(ref)  Thesecondstretch. 二次拉伸。
# 坐标核实过：`(REF)` 落在 Step6 行内，Step7 行的 Position 就是干净的 `1-30`
# —— **PDF 在这里漏标了 ref**。但 Purpose 列写明本步是「二次拉伸」，
# 即上一步（Step6，其 Position 标了 REF）工艺的延续 → 本步同一列也按 ref 算。
#
# 这条规矩**只加不减**：不认它，Excel 那两行（浮点 + 值列空 + 两个开关都是）
# 就会被判成不一致，而它其实是对的；认了它，机制上仍然是从 **PDF** 得出的结论
# （不是「Excel 填得自洽就算对」—— 那正是用户点名批评的旧机制）。
_EN_CONT_PURPOSE_RE = re.compile(r"二次\s*拉伸|再次\s*拉伸|second\s*stretch", re.I)


def _en_purpose_col(block):
    """表头块里 Purpose 列的下标（认不出返回 None）。"""
    for r in block:
        for c, v in enumerate(r or ()):
            if re.search(r"purpose|目\s*的", _n(v), re.I):
                return c
    return None


# --------------------------------------------------------------------------
#  子表标题里的设备号清单（第 53 轮）
#
#  第九种版式（英文横表）一直把 `equip` 写成死值 `[]` + `no_equip=True`，
#  于是收窄链整整少一级：**只能按直径分档**。而 WI-0036 的表2、表3 直径都是
#  Ø6.0，光看直径永远分不开 —— 表2（T-1177）的 `SL1+30 ref` 会串到
#  表3（其余 6 台）的 `5-180` 上，`dominant_shape` 又「有 ref 就按 ref」，
#  于是 card 0015 的「左/右拉伸距离/9」被判成「允许超应为是」（假错）。
#
#  标题里 T-xxxx 出现的**位置**决定它是不是设备号 —— 做成一张表，别散成 if：
#
#   写法                                       出处                  算设备号？
#   ──────────────────────────────────────────────────────────────────────
#   ① 「…参数设定：(T-1177、T-10706…)」        WI-0036 表1/2/3      对：括号里
#   ② 「…设备/工具编号：T-1114/T-1281激光…」   WI-0024              对：标签后
#   ③ 「表1-13，T-1362/T-1…尖端成型参数表」    WI-0024              对：表号后紧跟
#   ④ 「使用芯轴 T-2015-06」「焊接保护套…」    非标题行             错：工装编号
#
#  ⚠️ 这条**只在子表标题这一路用**（射程最大就是标题那一行）。数据行里的
#     T-xxxx 必须靠标签区分（第 28 轮已经在这个坑上栽过一次：芯轴 T-2015 被
#     当设备，把 T-1175 那一整套参数全挂错组）。
# --------------------------------------------------------------------------

# ④ 的排除词：标题里出现这些词，说明紧跟的 T-xxxx 是工装不是设备。
# 子表标题里几乎不会出现，写上是为了防止 `_title_above` 抓回来的标题带上了
# 正文残余（页面文本切行时确实会）。
_TITLE_EQUIP_BLOCK = re.compile(
    r"芯轴|保护套|模具|工装|夹具|刀具|量具|热缩管|套管")

# ② 的标签（标签之后直到第一个非「T-/数字/斜杠/顿号/逗号」的字符为止，全收）
_TITLE_EQUIP_LABEL = re.compile(
    r"(?:设备\s*/\s*工具\s*编号|设备编号|机器编号|机台编号|设备|机器|机台)\s*[:：]")

# 一串 T-清单：`T-1177、T-10706、…` / `T-1114/T-1281` / `T-10455/10456`
# （最后一种是「只有第一台带前缀」的省略写法，见 `_equip_set` 的注释）
_TITLE_EQUIP_RUN = re.compile(
    r"T\s*-\s*\d+(?:\s*[、,，/]\s*(?:T\s*-\s*)?\d+)*")


def _title_equip_set(title):
    """从**子表标题**里取出设备号集合；认不出返回空 frozenset。

    三级依次退让，先中最稳的那一档就收手（与 `_equip_set` 口径一致，
    返回的是 `T-<数字>` 形式的集合）：
      ① 括号里的 T-清单（`(T-1177、T-10706…)`）
      ② 「设备/工具编号：」标签之后的 T-清单
      ③ 「表N，」之后紧跟的 T-清单

    为什么必须①优先：标题里**只有**括号那一段是表格在明确声明设备范围。
    ②③ 是给「标题里没写括号」的老版式留的退路（WI-0024 那一批）。
    """
    flat = _c(str(title or ""))
    if not flat or "T" not in flat.upper():
        return frozenset()

    def _scan(text):
        """在一段文字里收 T-清单，**遇到工装排除词就停**。

        三级退让都要走它：① 是逐段括号，②③ 是标题的一段后缀。
        不截断的话，③ 那条退路会把「表1（使用芯轴 T-2015-06）」里的芯轴
        编号当成设备号收进来 —— 那正是第 28 轮栽过的坑。
        """
        m = _TITLE_EQUIP_BLOCK.search(text)
        if m:
            text = text[:m.start()]
        return _TITLE_EQUIP_RUN.findall(text)

    cands = []
    # ① 括号里的
    for m in re.finditer(r"[（(]([^（）()]*)[）)]?", flat):
        cands.extend(_scan(m.group(1)))
    # ② 标签之后的
    if not cands:
        for m in _TITLE_EQUIP_LABEL.finditer(flat):
            cands.extend(_scan(flat[m.end():m.end() + 120]))
    # ③ 表号之后紧跟的（这一段到第一个汉字/参数表字样为止）
    if not cands:
        m = re.search(r"表\s*\d+\s*[-－]?\s*[0-9A-Za-z]*\s*[,，、]?\s*", flat)
        if m:
            cands.extend(_scan(flat[m.end():m.end() + 120]))
    out = set()
    for seg in cands:
        out |= set(_equip_set(seg))
    return frozenset(out)


def _enstep_table(t, rows, pno, tno):
    """第九种版式：英文横表（Step 为行、参数为列）→ 记录列表。

    认不出来返回 [] —— 空列表表示「不接管」，交给 B 型那套逻辑照旧处理。
    """
    got = _enstep_detect(rows)
    if not got:
        return []
    _h0, _d0, step_rows, heads = got
    block = rows[_h0:_d0]
    # Purpose 列（「二次拉伸」这类延续说明就写在这儿；见 `_EN_CONT_PURPOSE_RE`）
    pcol = _en_purpose_col(block)

    # 合并单元格：'Position (mm)' 横跨 Left / Right 两列，pdfplumber 只在第一列
    # 给文字、第二列是 None —— 不把组头继承过去，右边那一列就整个认不出来
    # （实测：右拉伸距离 0 条，于是 Excel 每一步的右拉伸距离全判「无法匹配」）。
    try:
        gcells = list(t.rows[_d0 - 2].cells) if _d0 >= 2 else []
    except Exception:
        gcells = []
    last = None
    for c, h in enumerate(heads):
        if h:
            last = h
            continue
        if last and c < len(gcells) and gcells[c] is None:
            heads[c] = (c, last[1], last[2], _en_lr(block, c))

    title = _enstep_title(t)
    sub, dia = parse_sub_title(title)
    if not sub:
        m = re.search(r"表\s*(\d+)", title)
        sub = m.group(0) if m else ""
    # 标题括号里的设备号清单（`…参数设定：(T-1177)`）—— 见 `_title_equip_set`。
    # 这是本版式**唯一**的设备归属信号：表里不印「设备/工具编号」，
    # 而同直径档常常并排好几张表（WI-0036 表2/表3 都是 Ø6.0），
    # 少了它，收窄链只能按直径分档，隔壁设备的写法就会串过来。
    eqset = _title_equip_set(title)

    # 先收「原值」，引用（Same as step N）要等整张表收完才能解
    cells = {}
    # 来源坐标旁路：{(参数名, 步骤): (行号, 列号)}。
    # `cells` 是按「参数名 + 步骤」归并的 —— 同一个 Step 的续行会并到一条 key 上，
    # 行号在归并时就丢了。要补行号有两条路：改 `cells` 的结构（它的解包点散在
    # `_enstep_resolve_all` 等四五处，动一下就得同步改，风险大），或者另挂一张
    # 旁路表（这里选的）。旁路表**只给坐标**，解析一行都不读它。
    cells_at = {}
    steps = []
    purposes = {}
    # ⚠️ 续行（R47）：一张表里同一 Step 常常占**好几行**，续行的 Step# 那格是空的，
    # 只有参数格里有字 —— 那是「同一个步骤、另一档规格的值」。
    # WI-0033 第 10 页表1 就是这样：Step 4 压力写 `15for 2.25-2.5mm`，
    # 下一行同一列写 `30for 2.75-5.0mm`（Step# 空）。用户点名的正是这个版式：
    # 「里面明明只有一个表格，虽然表格内有不同规格指向不同信息」。
    # 旧代码遇到认不出 Step 的行直接 continue，那一整臂的值就凭空消失了 ——
    # 只剩 `15` 一支，于是 φ2.75 / φ3.0 / φ3.25-4.0 / φ4.5-5.0 四张卡全被判错。
    # 所以：Step# 空白的行归到上一个 Step。表头块（在任何 Step 之前）不受影响，
    # 因为那时 cur_step 还是 None。
    cur_step = None
    for i in range(step_rows[0], len(rows)):
        r = rows[i]
        if not r:
            continue
        m = _STEP_HEAD.match(_n(r[0]))
        if m:
            cur_step = int(m.group(1))
            steps.append(cur_step)
            if pcol is not None and pcol < len(r):
                purposes[cur_step] = _n(r[pcol])
        elif _n(r[0]) or cur_step is None:
            continue          # 表头 / 表尾文字 / 第一个 Step 之前 —— 都不是数据行
        step = cur_step
        for ci, h in enumerate(heads):
            if not h:
                # 这一列我认不出是什么参数，但格子里有字 → 记进对账袋。
                # 不记的话，整列的值人间蒸发，下游只会说「PDF 里没有它」。
                if ci and ci < len(r) and str(r[ci] or "").strip():
                    LAST_UNPARSED.append({
                        "page": pno, "table": tno, "sub": sub, "title": title,
                        "param": "", "step": step, "col": ci,
                        "raw": _show_cell(str(r[ci]).replace("\n", " / ")),
                        "reason": "no_head"})
                continue
            c, kind, unit, lr = h
            if c >= len(r):
                continue
            raw0 = str(r[c] or "").replace("\n", " / ")
            if not raw0.strip():
                continue
            for raw, cond in _enstep_branches(raw0):
                pname, punit = _en_param_of(kind, lr, raw, unit)
                if not pname:
                    LAST_UNPARSED.append({
                        "page": pno, "table": tno, "sub": sub, "title": title,
                        "param": "", "step": step, "col": c,
                        "raw": _show_cell(raw), "reason": "no_param"})
                    continue
                # 一格多量：整格里并排写了几个部位（中间模 / 左端模 / 右端模）
                # → 各归各的参数名。拆不动就按原样收，绝不因为拆不动而丢值。
                parts = _enstep_split_parts(raw)
                if parts:
                    _got = False
                    for label, seg in parts:
                        pn2 = _en_part_param(pname, label)
                        if not pn2:
                            continue
                        cells.setdefault((pn2, step), []).append((seg, punit, cond))
                        cells_at.setdefault((pn2, step), (i, c))
                        _got = True
                    if _got:
                        continue
                cells.setdefault((pname, step), []).append((raw, punit, cond))
                cells_at.setdefault((pname, step), (i, c))
    if not cells:
        return []

    # 「二次拉伸」的步骤号 —— 这些步按上一步工艺的延续处理（继承上一步同列的 ref）
    cont_steps = {s for s, p in purposes.items()
                  if _EN_CONT_PURPOSE_RE.search(str(p or ""))}

    def _inherit_ref(pname, step):
        """本步是「二次拉伸」时，上一步同一列标了 ref → 本格也算 ref。

        查的是 `cells` 里的**原值**（引用式也算——`_has_ref` 看的是字面有没有 ref
        字眼）。上一步那一列自己若是引用（如 `Samestep6`），字面没有 ref 就不继承，
        与「上一步有 ref 就是 ref」一致。
        """
        if step not in cont_steps:
            return False
        for praw, punit, _pc in (cells.get((pname, step - 1)) or []):
            if _has_ref(praw) or _has_ref(punit):
                return True
        return False

    recs = []
    for (pname, step), lst in cells.items():
        _at = cells_at.get((pname, step)) or (None, None)
        for raw, punit, cond in lst:
            base = {
                "page": pno, "table": tno, "sub": sub, "dia": dia, "title": title,
                "equip": sorted(eqset), "equip_raw": "/".join(sorted(eqset)),
                "sub_equip": [], "prog": "", "sleeve": "",
                "meta": {}, "param": pname, "step": step,
                "lo": None, "hi": None, "unit": punit, "note": "",
                "raw": _show_cell(raw), "empty": False,
                # 标题里认得出设备号就**不再**标 no_equip —— 那等于告诉收窄链
                # 「按设备过滤无从谈起」，白白放弃唯一剩下的信号。
                "no_equip": not eqset, "layout": "enstep", "cond": cond,
                "ref_inherit": _inherit_ref(pname, step),
                # 引用式（`Step5+5`）：本格不表态，性质由**被引用格**决定
                # （见 `_enstep_ref_via`）。与 `ref_inherit` 是两条独立规矩。
                "ref_via": _enstep_ref_via(raw, cells, pname),
                # 来源坐标：第 _at[0] 行、第 _at[1] 列（`cells_at` 旁路表）
                "_row": _at[0], "_col": _at[1], "_vcol": _at[1],
            }
            vals = _enstep_resolve_all(raw, cells, pname)
            if not vals:
                # 解析不出值：只有**明确写 N/A** 才记下来 —— 那是表格在声明
                # 「此步此参数不适用」，与「压根没有这一格」是两回事，丢了它
                # 就回答不了「Excel 在这里填了值是不是错的」。
                if pdf_cell_is_na(raw):
                    base["empty"] = True
                    base["na"] = True
                    recs.append(base)
                    continue
                # ⚠️ 解不出数值、但**格子里明明白白标了 ref**（`SL1+30 ref`、
                # `SR1+30 ref`、`>Step5-5℃ (ref)`）—— ref 本身就是**完整的答
                # 案**：按类型规矩它不必填值、且「允许超上下限」应填「是」
                # （见 FLOAT_SHAPE_RULES["ref"]）。
                # 以前只在「解不出值」这一层处理，整条记录被丢掉，于是这一格
                # 人间蒸发、只能被别的规格档的写法顶替：WI-0036 表1（Ø4.0~5.0）
                # Step9 左右位写的是 `SL1+30 ref`，工具看不见，就拿
                # 表3（Ø6.0）同一格的 `5-180` 当依据，反过来判 φ4.0-5.0 的卡
                # 「允许超应为否」—— 而用户填的恰恰是对的（2026-09-30 点名）。
                # **「读不出数值」≠「这一格不存在」**：ref 是一个确定的事实，
                # lo/hi 留空、`raw` 原样保留，让 pdf_value_shape 判成 ref。
                if _has_ref(raw) or _has_ref(punit):
                    recs.append(base)
                    continue
                # 认得出参数名、格子里也有字，就是解不出数值（`SL1+3 ref`、
                # `Vacuum`、`ON`、`Close mold`……）—— 这是**我没读懂**，不是
                # 「PDF 里没有」。以前静默丢掉，下游依据写成「这份 PDF 的参数表里
                # 没有它」，等于把工具的读不懂说成文件里没有。现在记进对账袋。
                LAST_UNPARSED.append({
                    "page": pno, "table": tno, "sub": sub, "title": title,
                    "param": pname, "step": step, "col": 0,
                    "raw": _show_cell(raw), "reason": "no_value"})
                continue
            for lo, hi, _u, note, rcond in vals:
                b = dict(base)
                b["lo"], b["hi"], b["note"] = lo, hi, note or ""
                if rcond:
                    b["cond"] = rcond      # 引用解出来的分支，各自带自己的生效条件
                recs.append(b)

    recs.extend(_enstep_notes(t, steps, pno, tno, sub, dia, title))
    return recs


# ==================================================== 第十种版式：矩阵表（参数名在首列、规格档在列头）
#
# WI-0035 / WI-0037「球囊料管拉伸参数」型：
#
#     ['参数设置',         '球囊直径',  '',              '']
#     ['',                '1.25-2.0',  '2.5-3.0×20-150', '5.0-6.0']
#     ['夹钳温度',         '不需拉伸',  '150-160℃',      …]
#     ['拉伸距离（参考）',  '左边160mm 右边160mm', …]
#
# 与第九种（英文横表）互为镜像：那边 **Step 为行、参数为列**；这边**参数名在首列**、
# **规格档在列头**。两边都不印设备号、都没有步骤号 → `no_equip=True`、`step=None`，
# 设备收窄那层照样跳过（理由同第九种：「没写」不等于「没有」）。
#
# 为什么前九种一条都不认：
#   · 竖表要「参数名称」这类表头词 —— 这里首格写的是「参数设置」；
#   · 第七种（裸键值表）要求值紧挨在 c1 —— 这里值按规格散在 c1..c6；
#   · 第八种（表头在首行）要求参数名在**顶行** —— 这里参数名在**首列**；
#   · B 型横表要「Step」表头 —— 这里压根没有步骤。
#
# 于是整张表静默消失。危险处不在「判不出」，而在**照样判得出**：工具一个值都没读到，
# 却给出「PDF 里没有这一项」的依据 —— 把「我读不懂」说成「文件里没有」（R47）。
# 所以这一版式除了产出记录，读不懂的格子一律进 `LAST_UNPARSED` 对账袋。

# 表尾截止词：扫到这些行就停（后面是质量要求和注意事项，不是参数）
_MATRIX_STOPWORDS = ("质量要求", "作业注意事项", "注意事项", "备注", "注意", "批准", "编制")

# 规格档格子的字符白名单。刻意**不含字母**：`160mm` 是值、`4.0-6.0` 才是规格 ——
# 放进 m/c 会让每个带单位的值都被当成规格档，整张表的列语义立刻颠倒。
_MATRIX_SPEC_OK = re.compile(r"^[\d.\s×xX✕*~～\-－—–/ØøÆæΦφ∅、，,;；()（）]+$")

# 方位词（长词优先，`左边` 要先于 `左` 命中，否则 `边` 会残留在值里）
_MATRIX_SIDE_RE = re.compile(
    r"(左端|左边|右端|右边|上端|上边|下端|下边|远端|近端|内侧|外侧|左|右|上|下)")
# 同一个方位的几种叫法统一到一个字（Excel 写「拉伸速度左端（远端）」→ 左）
_MATRIX_SIDE_ALIAS = {"左端": "左", "左边": "左", "远端": "左",
                      "右端": "右", "右边": "右", "近端": "右",
                      "上端": "上", "上边": "上", "上": "上",
                      "下端": "下", "下边": "下", "下": "下",
                      "内侧": "内", "外侧": "外"}


def _matrix_spec_cell(s):
    """这一格是不是一个「规格档」写法（1.25-2.0 / 4.0-6.0 / 2.5-3.0×20-150）。

    不是就返回 None。判据刻意收紧：只认「数字 + 区间符 + 直径/乘号」，
    带单位的值（160mm / 150-160℃）一律不算。
    """
    t = _c(str(s or "").replace("\n", " ")).strip()
    if not t or len(t) > 40:
        return None
    if not _MATRIX_SPEC_OK.match(t):
        return None
    if not re.search(r"\d", t):
        return None
    # 必须带「区间」或「乘号」—— 只有**分了档**的表才是矩阵表。
    # 只认数字是灾难性的（R48 实测，两条都踩过）：
    #   · WI-0035 第 3 页页眉那格 `3/3` 全是数字 → 物料明细表被当成参数区；
    #   · WI-0021 第 5 页「程序号」那一行的 `39` / `28` → 整张步骤横表被
    #     第十种版式抢走（下面 `continue` 掉 B 型逻辑），18315 条 ok 里
    #     5310 条凭空变成 no_match。
    # 「只有一个规格、没写区间」的表本来就不该走这条路径 —— 那种是裸键值表
    # （第七种）的活。宁可不接管，也不能抢。
    if not re.search(r"[-~～－—–×✕xX]", t):
        return None
    return t


def _matrix_sides(text):
    """「左边160mm 右边160mm」→ [('左', '160mm'), ('右', '160mm')] —— 一格两量。

    认不出就返回 []（调用方保持原样，绝不猜）。三种情况不拆：
      · 只出现一个方位（那本来就是单值）
      · 同一方位出现两次（那是别的结构，不是左右分列）
      · 某个方位词后面没有值
    """
    t = str(text or "").replace("\n", " ")
    hits = list(_MATRIX_SIDE_RE.finditer(t))
    if len(hits) < 2:
        return []
    names = [_MATRIX_SIDE_ALIAS.get(h.group(1), h.group(1)) for h in hits]
    if len(set(names)) != len(names):
        return []
    out = []
    for i, h in enumerate(hits):
        end = hits[i + 1].start() if i + 1 < len(hits) else len(t)
        seg = t[h.end():end].strip().strip("：:；;，,、（）() ")
        if not seg:
            return []
        out.append((names[i], seg))
    return out


def _matrix_detect(rows):
    """这张表是不是第十种版式？

    是就返回 (表头行号, 规格行号, {列号: 规格串}, [(行号, 参数名, [各列原值])])。
    判据有三条，缺一不可：
      ① 有一行首格是**不含数字**的列名（「参数设置」），
      ② 紧接着（同一行或下一行）有**一列以上是规格档**，
      ③ 再往下有 **≥2 行**首格是参数名、右边带值。
    光靠 ①③ 会把物料表 / 设备表也认进来 —— ② 是把它限定成「按规格分档的参数表」。
    """
    if not rows or len(rows) < 4:
        return None
    ncol = max((len(r or []) for r in rows), default=0)
    if not (2 <= ncol <= 24):
        return None
    # 取**参数行最多**的那一组，不是第一个命中的：一份 PDF 里往往还有物料表、
    # 设备表这些长得像的表（首格也是不含数字的列名）。只靠「第一个」会认错表。
    best = None
    for i in range(len(rows) - 2):
        r = rows[i] or []
        head = _c(r[0] if r else "")
        if not head or re.search(r"\d", head):
            continue                                   # 首格空 / 含数字 → 不是列名行
        for j in (i, i + 1):
            sr = rows[j] or []
            if j != i and _c(sr[0] if sr else ""):
                continue                               # 规格行的首格必须是空的
            specs = {}
            for c in range(1, min(ncol, len(sr))):
                v = _matrix_spec_cell(sr[c])
                if v:
                    specs[c] = v
            if not specs:
                continue
            body = []
            cur = None
            for k in range(j + 1, len(rows)):
                rr = rows[k] or []
                nm = _c(rr[0] if rr else "")
                vals = [_c(rr[c]) if c < len(rr) else "" for c in range(1, ncol)]
                if not any(vals):
                    if nm and any(w in nm for w in _MATRIX_STOPWORDS):
                        break
                    continue
                if nm:
                    if any(w in nm for w in _MATRIX_STOPWORDS):
                        break
                    if _matrix_spec_cell(nm):          # 又一行规格 → 已经出了数据区
                        break
                    cur = nm
                elif cur is None:
                    continue                           # 数据区之前的杂行
                # 续行（R47 同一病根）：首格空、右边有值 → 归到上一个参数名
                body.append((k, cur, vals))
            if len(body) < 2:
                continue
            # 值必须真的写在规格列上。否则「列名行 + 下面两行说明文字」也会被
            # 认成矩阵表（说明文字里照样有数字）。
            hit = sum(1 for (_k, _nm, vs) in body
                      if any((c - 1) < len(vs) and vs[c - 1] for c in specs))
            if hit < 2:
                continue
            if best is None or hit > best[0]:
                best = (hit, i, j, specs, body)
    if not best:
        return None
    return best[1], best[2], best[3], best[4]


def _matrix_table(t, rows, pno, tno):
    """第十种版式：矩阵表 → 记录列表。认不出来返回 []（交给后面的版式照旧处理）。"""
    got = _matrix_detect(rows)
    if not got:
        return []
    h_i, _s_i, specs, body = got
    title = _enstep_title(t)
    if not title:
        for j in range(h_i - 1, -1, -1):
            txt = _c((rows[j] or [""])[0] if rows[j] else "")
            if txt:
                title = txt
                break
    sub, dia = parse_sub_title(title)

    recs = []
    for (_k, nm, vals) in body:
        for c, spec in sorted(specs.items()):
            raw0 = str(vals[c - 1] if (c - 1) < len(vals) else "" or "")
            raw0 = raw0.replace("\n", " / ").strip()
            if not raw0:
                continue
            for side, seg in (_matrix_sides(raw0) or [("", raw0)]):
                pname = (nm + side) if side else nm
                base = {
                    "page": pno, "table": tno, "sub": sub, "dia": dia, "title": title,
                    "equip": [], "equip_raw": "", "sub_equip": [], "prog": "", "sleeve": "",
                    "meta": {}, "param": pname, "step": None,
                    "lo": None, "hi": None, "unit": "", "note": "",
                    "raw": _show_cell(seg), "empty": False,
                    # PDF 这张表没印设备号 —— 是「没写」，不是「没有这台设备」
                    "no_equip": True, "layout": "matrix", "cond": spec,
                    # 来源坐标：第 _k 行（参数名在首列），值在规格档那一列 c
                    "_row": _k, "_col": 0, "_vcol": c,
                }
                if pdf_cell_is_na(seg):
                    b = dict(base)
                    b["empty"] = True
                    b["na"] = True
                    recs.append(b)
                    continue
                v = parse_step_value(seg)
                if not v:
                    # 认得出参数名、格子里也有字，就是解不出数值 ——
                    # 「不需拉伸」「料泡长度+10（参考）」这类。这是**我没读懂**，
                    # 不是「PDF 里没有」，记进对账袋让依据能如实说出来。
                    LAST_UNPARSED.append({
                        "page": pno, "table": tno, "sub": sub, "title": title,
                        "param": pname, "step": None, "col": c,
                        "raw": _show_cell(seg), "reason": "no_value"})
                    continue
                b = dict(base)
                b["lo"], b["hi"], b["unit"] = v[0], v[1], (v[2] or "")
                b["note"] = v[3] or ""
                recs.append(b)
    return recs


# ==================================================== A 型：竖表（参数名在列上）
#
# 版式特征：表头含「参数名称」「项目」，参数名在各行的某一列上，值在同一行右侧。
# 一张表里常常上下堆着好几段「适用范围 + 表头 + 数据行」——
# 例如 WI-0008 p5 一张表里堆了 4 段（不同产品/直径，值互不相同）。
#
# 适用范围出现在三处，缺一不可：
#   ① 独立一行     「设备/工具编号：T-2057/T-1254… 设备/工具说明：热焊接机」
#   ② 表头单元格内 「参数名称 (产品线A/Aperi SC Φ1.5-2.5mm)」← 适用规格藏在这里
#   ③ 块标题行     「表1-2: 机器设置参数表参数设定表：T-10528/ T-10529」
#
# 为什么必须带适用范围：丢了它，同名参数就会跨设备/跨规格并成一个大集合。
# WI-0007 的「光强」会把 T-1018 组的 1.5-2.0 与 T-1540 组的 3-5 混在一起，
# 于是把值填成隔壁设备那一套也判「一致」——值域过宽等于没核对。
#
# 本函数只负责「这段文字属于哪个块」，数值语义仍由 engine 统一解析，避免两处实现漂移。

_V_SCOPE_EQUIP = re.compile(r"设备\s*[／/]?\s*工\s*具\s*编\s*号|设备\s*编\s*号")
_V_SCOPE_DESC = re.compile(r"设备\s*[／/]?\s*工\s*具\s*说\s*明")
# 「设备/工具说明：热焊接机 T-2069/…」—— 有的表压根不写「编号」二字，
# 整页设备清单全靠这一句（WI-0019 p5）。老口径只认「编号」，结果**整页设备号
# 一个都取不到**：表3（五台）和表4（两台）分不开，温度混出两套候选。
_V_SCOPE_EQUIP_ANY = re.compile(
    r"设备\s*[／/]?\s*工\s*具\s*(?:编\s*号|说\s*明)|设备\s*(?:编\s*号|说\s*明)")
# 「使用芯轴 T-2015-06」「焊接保护套 DWG…」里的 T-xxxx 是**工装编号，不是设备**。
# 它们和设备号格式一模一样，只能靠标签区分（WI-0025 p5 实测：不挡掉的话
# T-1175 会被芯轴 T-2015 顶掉，整套参数挂到错误的设备组上）。
_NOT_EQUIP_ROW = re.compile(r"芯\s*轴|保护套|物\s*料|程序号|图\s*号|工装|夹具|模\s*具|产品编号")
_V_SUB_RE = re.compile(r"(设备参数\s*(?:设置)?\s*附表\s*[0-9][0-9A-Za-z\-－]*"
                       r"|参数附表\s*[0-9]+"
                       r"|表\s*[0-9]+\s*[-－]\s*[0-9A-Za-z]+"
                       r"|表\s*[0-9]+\s*[,，])")
_V_PROG_RE = re.compile(r"程序号\s*(?:[/／]\s*名称)?\s*[：:]\s*([0-9A-Za-z\-/、]+)")
_V_TITLE_HINT = ("参数表", "附表", "设定表", "参数设定")
_V_END = ("物料明细表", "工具设备明细表", "质量要求", "作业注意事项", "注意事项",
          "变更记录", "修订记录")
# 直径：Φ1.5-2.5 / 1.5-2.5mm / φ2.0 / 0.8-4.0mm（φ 可有可无，mm 可有可无）
#
# ⚠️ 区间后半段**也要允许再带一个直径符**：WI-0036 表1 标题写的是
# `Ø4.0~Ø5.0`（每个数字各自带符号）。旧写法只认 `Ø4.0~5.0`，于是
# `Ø4.0~Ø5.0` 被切成两个单值 (4.0,4.0) 与 (5.0,5.0) —— 直径档变成 4.0 而不是
# 4.0-5.0，按规格收窄时与 Excel 的「φ4.0-5.0」对不上，整张表的定位就偏了。
_DIA_PHI = re.compile(
    r"[φΦ⌀∅Øø]\s*(\d+(?:\.\d+)?)"
    r"(?:\s*[-~～]\s*[φΦ⌀∅Øø]?\s*(\d+(?:\.\d+)?))?")
_DIA_MM = re.compile(r"(\d+(?:\.\d+)?)\s*(?:[-~～]\s*(\d+(?:\.\d+)?))?\s*mm")
_V_CODE = re.compile(r"^(DWG|SPC|T-|P/N)")
_V_BLOB_KEY = ("参数名称", "参数值", "项目")
# 表头行里「参数名」那一列：'参数名称' / '参 数名 称 (…)'(归一后仍以其开头) / '项目'
# ⚠️ `_V_HEADER_KEYS = ("参数名称", "项目", "参数")` 已删（2026-09-30 重构）：
# 竖表表头的认法后来统一走 `SYMBOL_SPECS` + `_v_header_row`，这张旧表没人读。


def _flat(s):
    """去掉所有空白（含全角空格 / 不换行空格），再做符号归一。

    为什么这里也要走 `_sym`：R47 把符号归一做在 `_c` / `_norm` / `_clean` 上，
    漏了 `_flat` —— 而 `_flat` 正是 `dias_from_text` / `families_of` / `scope_of`
    这几个**读标题**的入口。结果 WI-0036 表1 标题
    `直径为\uf0c64.0~\uf0c65.0`（Symbol 字体的 Ø）在 `_c` 里看着是 `Ø4.0~Ø5.0`，
    一到 `_flat` 就退回私有区字符，直径档解析不出来 —— 同一段文字两个入口
    两种看法，最伤的那种 bug。
    """
    t = re.sub(r"[\s\u3000\u00a0]+", "", str(s or ""))
    # 快速路径：绝大多数串里没有需要归一的符号，别为它付 translate 的钱
    return _sym(t) if _SYM_ANY.search(t) else t


# 产品家族名：同名参数按「产品线」分成好几套附表时（WI-0011 表2-a 是 产品线A 线、
# 表2-b 是 产品线C 线，值不同），这是唯一能区分它们的线索。
_FAMILY_TOKENS = ("产品线B NC", "Aperi SC", "Aperi NC", "Origin SC", "cNovate",
                  "产品线B", "产品线A", "产品线C", "产品线D", "产品线E", "产品线F", "POT")


def families_of(text):
    """从一段文字里取产品家族名（如 产品线A / 产品线C）。

    含包含关系的取更具体的那个：「产品线B NC」命中时不再记「产品线B」。
    """
    t = _flat(text).lower()
    hit = {f for f in _FAMILY_TOKENS if _flat(f).lower() in t}
    return {f for f in hit
            if not any(f != g and _flat(f).lower() in _flat(g).lower() for g in hit)}


def equip_from_text(text):
    """从一段文字里取全部设备号。

    容忍 PDF 换行造成的断裂写法：`T-2057/T-1254/T-1534/T-/1535/T-1699`
    里的 `T-/1535` 是「T-1535」被折行切开，不修就整整丢掉两台设备。

    还要展开「只有第一台带前缀」的省略写法：`T-10455/10456/10457` 是三台设备，
    老写法 `T-(\\d+)` 只认出第一台 —— WI-0008 p8 两张表的设备号都这么写，
    Excel 那头又只填了其中两台（首台还多个空格），两边各丢一半、交集一空，
    就报「PDF 里没有该设备对应的参数」（假漏判）。
    """
    t = _flat(text).replace("T-/", "T-").replace("T-／", "T-")
    out = set()
    for m in re.finditer(r"T-(\d+(?:/\d+)*)", t):
        for n in m.group(1).split("/"):
            if n:
                out.add("T-" + n)
    return sorted(out)


def _title_dias(text):
    """子表标题里的适用规格 —— 比 `dias_from_text` 宽松，认**不带 mm** 的括号区间。

    只在「子表标题」这一种上下文里用，不得拿去扫数据行：数据行里的 (4.5-6.0)
    可能是公差、程序号、步号，扫过去等于凭空造规格。

    WI-0025 p5：表2-a 写「…激光焊接参数表(1.5-4.0mm)」、表2-b 写「…(4.5-6.0)」——
    后者没有 mm，老口径一条都取不到，于是表2-b 整套参数沿用了表2-a 的 (1.5-4.0)。
    两张表设备号又都是 T-1175，规格一错就再没有任何东西能把它们分开。
    """
    flat = _flat(text)
    out = list(dias_from_text(flat))
    if out:
        return out
    for a, b in re.findall(r"[（(]\s*(\d+(?:\.\d+)?)\s*[-~～]\s*(\d+(?:\.\d+)?)\s*[)）]", flat):
        x, y = float(a), float(b)
        if x > y:
            x, y = y, x
        # 只认「像球囊/导管直径」的量级：0.05~60mm，别把 (8-10) 这种程序号当规格
        if 0.05 <= x <= 60 and (x, y) not in out:
            out.append((x, y))
    return out


def dias_from_text(text):
    """从一段文字里取全部直径区间 [(lo,hi)]（Φ 可有可无、mm 可有可无）。"""
    out = []
    flat = _flat(text)
    for rx in (_DIA_PHI, _DIA_MM):
        for a, b in rx.findall(flat):
            x = float(a)
            y = float(b) if b else x
            if x > y:
                x, y = y, x
            if (x, y) not in out:
                out.append((x, y))
    return out


def scope_of(text):
    """从一段文字解析适用范围：{sub, equip_raw, equip, dias, prog, families}。"""
    flat = _flat(text)
    out = {"sub": "", "equip_raw": "", "equip": [], "dias": [], "prog": "", "families": []}
    m = _V_SUB_RE.search(flat)
    if m:
        out["sub"] = m.group(1).rstrip("，,")
    m = _V_SCOPE_EQUIP.search(flat)
    if m:
        tail = flat[m.end():].lstrip("：:")
        d = _V_SCOPE_DESC.search(tail)
        if d:
            tail = tail[:d.start()]
        out["equip_raw"] = tail.strip("：:，,、/（）()")
    else:
        # 没有「编号」二字时的退路：认「设备/工具说明：」（WI-0019 p5）。
        # 一行里可能并排两张子表各写一句，所以取到**下一句同类标签**为止，
        # 否则表4 的两台会被并进表3 的清单里，两块又分不开了。
        m2 = _V_SCOPE_EQUIP_ANY.search(flat)
        if m2:
            tail = flat[m2.end():].lstrip("：:")
            d = _V_SCOPE_EQUIP_ANY.search(tail)
            if d:
                tail = tail[:d.start()]
            out["equip_raw"] = tail.strip("：:，,、/（）()")
    m = _V_PROG_RE.search(flat)
    if m:
        out["prog"] = m.group(1).strip("、，,/")
    out["dias"] = dias_from_text(flat)
    out["equip"] = equip_from_text(out["equip_raw"]) if out["equip_raw"] else []
    out["families"] = sorted(families_of(flat))
    return out


def _v_header_cols(row):
    """表头行里的「参数名」列 → {列号: 单元原文}。"""
    out = {}
    for c, v in enumerate(row or []):
        k = _n(v)
        if not k:
            continue
        if k.startswith("参数名称") or k in ("项目", "参数"):
            out[c] = _c(v)
    return out


def _is_blob(row):
    """整行塌成一格的大段文字（页眉那种「合并单元格」）：

    WI-0010 p4 有一张 2 行的表，第 2 行把整页参数挤在一个格子里（含换行）。
    当成表头会把一整段文字误判成参数名，直接跳过。
    """
    filled = [_c(c) for c in (row or []) if _c(c)]
    if len(filled) != 1:
        return False
    txt = filled[0]
    return ("\n" in txt and len(txt) > 120) or sum(k in txt for k in _V_BLOB_KEY) >= 2


def _v_data_row(row, name_col, stop):
    """按「参数名在第 name_col 列」读一行数据，返回 (参数名, [值原文…]) 或 None。"""
    if name_col >= len(row or []):
        return None
    raw = _c(row[name_col])
    if not raw:
        return None
    if _v_header_cols(row):                    # 又是表头行
        return None
    if name_col > 0:                           # 序号列必须真的是序号
        n0 = _n(row[0]) if row[0] else ""
        if n0 and not re.match(r"^\d+(?:/\d+)*$", n0):
            return None
    if _V_CODE.match(raw) or "**" in raw or _n(raw) in ("n/a", "na", "无"):
        return None
    vals = []
    for j in range(name_col + 1, min(stop, len(row))):
        t = _c(row[j])
        if not t:
            continue
        if re.match(r"^\d+(?:/\d+)*$", _n(t)):
            continue                           # 序号 / 步骤号
        vals.append(t)
    if not vals:
        return None
    return raw, vals


def _row_is_not_equip(row, name_col=0):
    """这一行是不是「工装/物料清单」行 —— 是的话里面的 T-xxxx 不是设备号。

    WI-0025 p5：「使用芯轴 T-2015-06」和「焊接保护套 DWG-0000007-01」跟设备号
    写在同一个位置，格式也是 `T-数字`。不挡掉的话，逐列往上找设备时会先撞上
    芯轴那一行，把 T-2015 当成设备 —— 于是 T-1175 那一整套参数全挂错组。
    判据只看**行首标签**（参数名那一列及其左边），不因为格子里出现数字就误伤。
    """
    if not row:
        return False
    lab = " ".join(_c(x) for x in row[:max(1, name_col + 1)] if _c(x))
    return bool(_NOT_EQUIP_ROW.search(_flat(lab)))


def _v_implicit_row(row):
    """无表头的裸竖表行（WI-0007：「序号 | 参数名 | … | 值」，表头四个字被省了）。

    返回 (name_col, 参数名, [值原文…]) 或 None。
    绝不能把「物料明细表」「工具设备明细表」的行吃进来 ——
    那两段的第一列也是序号，靠「名字像不像物料/设备编号」把它们挡掉。
    """
    if not row or len(row) < 3:
        return None
    if not re.match(r"^\d{1,2}$", _n(row[0])):
        return None
    for c in range(1, len(row)):
        raw = _c(row[c])
        if not raw:
            continue
        if _V_CODE.match(raw) or "**" in raw or not re.search(r"[一-鿿A-Za-z]", raw):
            return None                        # 物料编号 / 设备编号行
        if len(raw) > 40:
            return None
        vals = [t for t in (_c(x) for x in row[c + 1:]) if t and not re.match(r"^\d+(?:/\d+)*$", _n(t))]
        if not vals or not any(re.search(r"\d", v) for v in vals):
            return None
        return c, raw, vals
    return None


# 参数名列**左边那一格**的常见写法：并排子表时它是每张子表的左边界锚点。
_V_LEAD_COL = ("序号", "项目", "参数", "no", "no.", "编号")


def _v_data_row_cols(row, name_col, stop, col_equip, left=None):
    """「每个值列各挂一组设备」时按列拆数据行（WI-0016 型）。

    表头上面一行写着「项目 | 拉伸机：T-A | 拉伸机：T-B」，同一行的两个值列
    **属于不同的设备组**。塞进一条记录（整块一个设备集）就分不清哪个值归哪台机，
    于是 T-1279 的 110-130℃ 会跟 T-1064 组的 90-120℃ 混成一套候选。

    返回 [(列号, 值原文, 该列设备列表 or None)]；本行不属于这块时返回 []。
    """
    if name_col >= len(row or []):
        return []
    raw = _c(row[name_col])
    if not raw:
        return []
    if _v_header_cols(row):                    # 又是表头行
        return []
    if left is not None and left < len(row or []):
        n0 = _n(row[left]) if row[left] else ""
        if n0 and not re.match(r"^\d+(?:/\d+)*$", n0):
            return []
    if _V_CODE.match(raw) or "**" in raw or _n(raw) in ("n/a", "na", "无"):
        return []
    out = []
    for j in range(name_col + 1, min(stop, len(row or []))):
        t = _c(row[j])
        if not t:
            continue
        if re.match(r"^\d+(?:/\d+)*$", _n(t)):
            continue                           # 序号 / 步骤号
        out.append((j, t, col_equip.get(j)))
    return out


# ------------------------------------------------------------------ 第七种版式
# **无表头的裸键值表**：整张表从头到尾就是「参数名 | 值…」。
#
# 第一种版式（竖表）要见到「参数名称 / 项目 / 参数」表头行才开块；裸行兜底
# `_v_implicit_row` 又要求 row[0] 是 1-2 位数字序号。而 WI-0004 rev28 的
# 相关参数附表 1 长这样 —— 表头被整个省掉了，第一列直接就是参数名：:
#
#     ['程序号',                 'Procedure 19/程序19']
#     ['激光焊接机',              'T-1011']
#     ['参数步骤',                'Step 1: 圆周焊']
#     ['Laser Power激光能量',     '550~650mW']
#     ['Spot size光斑尺寸',       '2.5mm']
#     ['Rotation Speed转速',      '1000 rpm']
#     ['Welding Time持续时间',    '3.0-4.0s']
#
# 于是竖表路径 0 条、裸行兜底 0 条、程序附表 0 条 —— 六个参数表一张都不认，
# 整份 WI 一个参数点都出不来。这类「只有两列、没有表头」的写法在车间 WI 里
# 反而最常见（附表越改越简），所以单独立一种版式，不塞进竖表里凑条件。
_DIA_MAX_UP = 220.0      # 子表标题离它自己的表最多这么远（pt）；再远就不是它的标题了

_BARE_MAX_COL = 12            # 一台机器一列时列数可以到十几，超过就当别的东西
_BARE_NAME_MAXLEN = 44        # 参数名不可能这么长，超了多半是塌成一格的整段文字
_BARE_MAT_KEY = ("物料编号", "物料名称", "使用量", "材质")     # 物料明细表：整表排除
_BARE_HDR_KEY = ("工序名称", "文件编号", "页码", "版本")       # 页眉/标题：跳过该行
# 这几种词出现在**任意一格**里就说明这不是参数表（整表排除，不只看首行）。
# 第一版只看了首行，结果 WI-3852 的「操作流程说明表」（物料明细表 + 序号/具体操作/
# 属性/注意事项）和 WI-0011 的球囊直径横表全被当成裸键值表收进来 —— 老语料 13 张
# 「疑似表」里 **13 张都是误判**。它们没造成错判只是因为那些参数名不在词典里
# （落 name_unknown），属于**运气**，不是判据对。
_BARE_NOT_PARAM = ("物料明细", "工具设备明细", "操作流程", "注意事项",
                   "具体操作", "质量要", "作业准备")
# 「激光焊接机」这类设备定义行的名字特征。注意「激光焊接机」里并没有「机器」二字
# （是「…接机」），所以除了「机器/设备/机台」还得收「以机结尾」——漏了这一条，
# 六张参数表的设备号一个都认不出来，全表的参数都成了「无设备归属」。
_EQUIP_NAME_RE = re.compile(r"(机器|设备|机台|仪器|机号)|机\s*$")
_BARE_PROG_KEY = re.compile(r"^\s*(程序号|程序)\s*$")


def _bare_kv_rows(rows):
    """rows 是不是「无表头裸键值表」。是则返回 (设备定义行号 or None, 参数行号列表)。

    判据一律是**结构**上的（有几列、第一列像不像参数名、值里有没有数字），
    不认任何具体参数名 —— 认名字就成了给某一款产品打补丁，换一份 WI 又漏。
    """
    if not rows or len(rows) < 3:
        return None
    ncol = max((len(r or []) for r in rows), default=0)
    if not (2 <= ncol <= _BARE_MAX_COL):
        return None
    head = _flat(" ".join(_c(x) for x in (rows[0] or []) if _c(x)))
    if any(k in head for k in _BARE_MAT_KEY):
        return None                                  # 物料明细表
    for r in rows:
        j = _flat(" ".join(_c(x) for x in (r or []) if _c(x)))
        if any(k in j for k in _BARE_NOT_PARAM):
            return None                              # 流程说明表 / 明细表：整表排除
    eq_row, params = None, []
    for i, r in enumerate(rows):
        if not r or _is_blob(r):
            continue
        c0 = _c(r[0])
        if not c0:
            continue
        joined = _flat(" ".join(_c(x) for x in r if _c(x)))
        if any(k in joined for k in _BARE_HDR_KEY):
            continue                                 # 页眉行
        if _v_header_cols(r):
            return None                              # 有表头 → 归竖表路径，别出两份
        if len(c0) > _BARE_NAME_MAXLEN or _V_CODE.match(c0) or "**" in c0:
            continue
        if not re.search(r"[一-鿿A-Za-z]", c0):
            continue                                 # 序号行 / 纯编号行
        # 裸键值表的定义就是「参数名 | 值」——**值必须紧挨在参数名右边那一列**。
        # 不要求这一条，横表/多列表就会被误收：WI-0011 p8 的球囊直径表
        # 「Balloon Diameter 球囊直径 | (空) | Setting Value 设定值」，
        # 值落在第 2 列而不是第 1 列，它是横表，不该走这条路径。
        if len(r) < 2 or not _c(r[1]):
            continue
        vals = [v for v in (_c(x) for x in r[1:])
                if v and not re.match(r"^\d+(?:/\d+)*$", _n(v))]
        if not vals:
            continue
        if eq_row is None and _EQUIP_NAME_RE.search(c0) and \
                any(_flat(v).upper().startswith("T") for v in vals):
            eq_row = i                               # 「激光焊接机 | T-1011 | …」
        params.append(i)
    # 至少两行「值里带数字」才当参数表：光有名字没有值的是说明表
    numish = sum(1 for i in params
                 if any(re.search(r"\d", v)
                        for v in (_c(x) for x in (rows[i] or [])[1:]) if v))
    if numish < 2 or not params:
        return None
    return eq_row, params


def _bare_pn_above(page, y0, xr, max_up=_DIA_MAX_UP):
    """表格**头顶**、且横向落在表格范围内的图号。

    为什么不用现成的 `_title_above`：它按「整行文本」认领，而并排子表的两个标题
    常常印在同一行文本里（「表1-c…DWG-0000010 表1-e…DWG-0000004和DWG-0000010」），
    按行认领会让左右两张子表拿到同一份图号 —— 而图号正是区分它们的唯一线索。
    所以这里按 **word 的 x 中点**逐个筛，左表只收左半的词、右表只收右半的词。
    """
    if y0 is None:
        return None
    try:
        words = page.extract_words() or []
    except Exception:
        return None
    buf = []
    for w in words:
        top = w.get("top")
        if top is None or top > y0 or y0 - top > max_up:
            continue
        if xr:
            mid = ((w.get("x0") or 0) + (w.get("x1") or 0)) / 2.0
            if not (xr[0] - 2 <= mid <= xr[1] + 2):
                continue
        txt = w.get("text") or ""
        if not _pn_cores(txt):
            continue
        buf.append((top, txt))
    if not buf:
        return None
    top1 = max(t for (t, _) in buf)
    # 只取离表格最近的那一行（同一行的 word，top 相差不到 4pt）
    line = " ".join(txt for (t, txt) in buf if abs(t - top1) < 4)
    return _pn_cores(line) or None


def _bare_kv_page(page, pgno):
    """第七种版式：无表头的「参数名 | 值…」裸键值表（WI-0004 型）。"""
    recs = []
    try:
        tables = page.find_tables() or []
    except Exception:
        return recs
    for tno, t in enumerate(tables):
        try:
            rows = t.extract() or []
        except Exception:
            continue
        got = _bare_kv_rows(rows)
        if not got:
            continue
        eq_row, params = got
        try:
            y0, x0, x1 = t.bbox[1], t.bbox[0], t.bbox[2]
        except Exception:
            y0, x0, x1 = _row_top(t, 0), None, None
        # 设备号：一台机器一列时逐列挂；值只写一遍则所有机器共用这套参数
        col_equip, blk_equip = {}, []
        if eq_row is not None:
            r = rows[eq_row] or []
            for cc in range(1, len(r)):
                eq = equip_from_text(_flat(_c(r[cc]) or ""))
                if eq:
                    col_equip[cc] = list(eq)
                    blk_equip.extend(eq)
        prog = ""
        for r in rows:
            c0 = _c(r[0]) if r else ""
            if _BARE_PROG_KEY.match(_n(c0)) and len(r) > 1:
                prog = _c(r[1]) or ""
                break
        # 图号：标题行有时被 pdfplumber 划进表内（表1-d 就是这样），先扫表内，
        # 表内没有再往表格头顶、横向落在表宽范围内的文字里找（见 `_bare_pn_above`）。
        pn = None
        for r in rows:
            for cell in (r or []):
                p = _pn_cores(_c(cell) or "")
                if p:
                    pn = p
                    break
            if pn:
                break
        if not pn:
            pn = _bare_pn_above(page, y0, (x0, x1) if x0 is not None else None)
        for i in params:
            r = rows[i] or []
            name = _c(r[0])
            items = [(cc, _c(r[cc])) for cc in range(1, len(r)) if _c(r[cc])]
            if not items:
                continue
            if len(items) == 1 and len(col_equip) > 1:
                # 「激光焊接机 | T-A | T-B | …」十台机器一列，参数值却只在第一列
                # 写了一遍 —— 这是 WI 的惯例：**同一套参数适用于这一行所有机器**。
                # 按列拆开就只剩 T-A 一台，其余九台全被报成「PDF 里没有」。
                eqs = [blk_equip]
            else:
                eqs = [col_equip.get(cc) or blk_equip for (cc, _) in items]
            for (cc, val), eq in zip(items, eqs):
                recs.append({
                    "page": pgno, "table": tno,
                    "sub": "", "equip": list(eq or []),
                    "equip_raw": "/".join(eq or []),
                    "dia": None, "dias": [],
                    "prog": prog,
                    "families": [],
                    "pn": pn,
                    "param": name, "text": val, "raw": name,
                    # 来源坐标：裸键值表，参数名固定在第 0 列、值在 cc 列
                    "_row": i, "_col": 0, "_vcol": cc, "_box": _tbox(t),
                })
    return recs


def _headrow_kv_rows(rows):
    """第八种版式：**表头在首行**、下面每行一档的横表（WI-3264 p10 型）。

        ['切割长度（mm）', '产品规格',      '切割后定长（mm）']
        ['1.5',          '产品线G/Castor…', '31±2']
        ['2.5',          'Minerva…',      '30±2']
        ['3.5',          'Minerva…',      '29±2']

    参数名在**顶行**、值在下方，与第七种（参数名在**左列**）正好是转置。
    `_v_header_cols` 只认「参数名称/项目/参数」三个词，这里一个都没有 → 竖表路径 0 条；
    `_bare_kv_rows` 又要求「参数名在 c0、值紧挨在 c1」→ 也 0 条。整张表静默消失，
    Excel 侧 419 条「切割长度」全部落 `name_unknown`。

    怎么和第七种区分（两者都是两列以上的小表）：
    **看首行含不含数字**。裸键值表的首行右边那格就是**值**
    （`程序号 | Procedure 19/程序19`，含 19）；表头行的每一格都是**列名**
    （`切割长度（mm）`，数字只可能出现在单位括号里）。
    """
    if not rows or len(rows) < 3:
        return None
    ncol = max((len(r or []) for r in rows), default=0)
    if not (2 <= ncol <= _BARE_MAX_COL):
        return None
    for r in rows:                                   # 非参数表：整表排除
        j = _flat(" ".join(_c(x) for x in (r or []) if _c(x)))
        if any(k in j for k in _BARE_NOT_PARAM):
            return None
    top = [_c(x) for x in (rows[0] or [])]
    cols = [c for c, v in enumerate(top) if v]
    if len(cols) < 2:
        return None
    for c in cols:
        # 列名里不该有数字；单位括号（（mm）/（℃）/（s））里的字母不算数字
        if re.search(r"\d", re.sub(r"[（(][^）)]*[）)]", "", top[c])):
            return None
    body = []
    for i, r in enumerate(rows[1:], 1):
        if not r:
            continue
        vals = [_c(r[c]) if c < len(r) else "" for c in cols]
        if not any(vals):
            continue
        body.append((i, vals))
    if len(body) < 2:
        return None
    if sum(1 for (_, vs) in body if any(re.search(r"\d", v) for v in vs)) < 2:
        return None                                  # 光有名字没有值的说明表
    return cols, body


def _headrow_kv_page(page, pgno):
    """第八种版式：表头在首行、下面每行一档的横表。"""
    recs = []
    try:
        tables = page.find_tables() or []
    except Exception:
        return recs
    for tno, t in enumerate(tables):
        try:
            rows = t.extract() or []
        except Exception:
            continue
        got = _headrow_kv_rows(rows)
        if not got:
            continue
        cols, body = got
        try:
            y0, x0, x1 = t.bbox[1], t.bbox[0], t.bbox[2]
        except Exception:
            y0, x0, x1 = _row_top(t, 0), None, None
        pn = None
        for r in rows:
            for cell in (r or []):
                p = _pn_cores(_c(cell) or "")
                if p:
                    pn = p
                    break
            if pn:
                break
        if not pn:
            pn = _bare_pn_above(page, y0, (x0, x1) if x0 is not None else None)
        # —— 行级图号（2026-10-29）——
        # 表级那一个 pn 是「整表一个图号」的假设，对单档表成立；但「一行一个
        # P/N」的表（WI-0044 p6 表3：22 行 22 个 P/N）会让整表 22 行粘成同一档，
        # 收窄链拿到四个不同的 Excel 规格（DWG-0000007-01/-02/-03/-04）也无从下手，
        # 只能把 22 个长度值并成一个池子 —— 多套对应值就是这么来的。
        # 所以逐行先看**本行**有没有图号，有就盖掉表级的。
        row_pn = {}
        for _i, _r in enumerate(rows):
            _got = set()
            for _cell in (_r or []):
                _got |= _pn_cell_key(_cell)
            if _got:
                row_pn[_i] = frozenset(_got)
        top = [_c(x) for x in (rows[0] or [])]
        for (i, vals) in body:
            for c, v in zip(cols, vals):
                if not v:
                    continue
                # 这一行的「规格数字」= **除本格以外**各格里的数字区间。
                # 中间列「产品规格」才是规格所在；把本格排除掉，是为了避免
                # 「值本身就是区间」时自己跟自己比（例如某档的值写成 3.0-4.0，
                # 那么 Excel 规格 φ(3.0-4.0) 会凭空命中这一行）。
                ranges = set()
                for j, vj in enumerate(vals):
                    if j != c and vj:
                        ranges |= _spec_ranges(vj)
                recs.append({
                    "page": pgno, "table": tno,
                    "sub": "", "equip": [], "equip_raw": "",
                    "dia": None, "dias": [],
                    "spec_ranges": frozenset(ranges),
                    "prog": "", "families": [],
                    "pn": pn,
                    "pn_full": row_pn.get(i) or frozenset(),
                    "param": top[c], "text": v, "raw": top[c],
                    # 来源坐标：表头在首行，第 i 行第 c 列既是参数名也是值
                    "_row": i, "_col": c, "_vcol": c, "_box": _tbox(t),
                })
    return recs


def _vertical_page(page, pgno):
    """解析一页里的所有竖表参数块。"""
    recs = []
    try:
        tables = page.find_tables() or []
    except Exception:
        return recs
    # 词级文本行 —— 表**框外**的标题只存在页面文本里（2026-11-11 补）。
    # 懒取：只有真遇到「表内没写设备/子表号」的块才去解析，绝大多数页不付这个成本。
    _plines = None
    for tno, t in enumerate(tables):
        try:
            rows = t.extract() or []
        except Exception:
            continue
        if len(rows) < 3:
            continue
        pending = None            # 最近一次读到的适用范围，等下一个表头来认领
        pend_titles = []          # [(列号, scope)] 一行里并排好几张子表标题时逐个记下
        blocks = {}               # name_col -> 块信息
        for ri, row in enumerate(rows):
            cells = [_c(c) for c in (row or [])]
            joined = " ".join(c for c in cells if c)
            fj = _flat(joined)
            if not fj:
                continue
            if any(k in fj for k in _V_END):
                pending, pend_titles, blocks = None, [], {}     # 离开参数段落
                continue
            if _is_blob(row):
                continue
            sc = scope_of(joined)
            is_title = any(k in fj for k in _V_TITLE_HINT)
            # 上下相邻两张子表被 pdfplumber 并进同一个 cell（表7-a…表7-b…）时，
            # scope_of 只取第一个标题（表7-a），但本 table 的数据行对应的是
            # **最后一张**（表7-b）。用最后一个标题把 sub / 直径档更正过来，
            # 否则表7-b 的值会错误沿用表7-a 的直径档（WI-0006 p11/p12）。
            _lsub, _ldia = _last_sub_title(joined)
            if _lsub:
                sc["sub"] = _lsub
                if _ldia:
                    sc["dias"] = [_ldia]
            if sc["equip_raw"] or sc["sub"] or is_title:
                # 块标题行（如「表1, 机器设置参数表（T-1065/…）」）里只有表名和设备清单，
                # 没有「设备/工具编号：」这个标签，得直接抓 T-xxxx。
                if is_title and not sc["equip"]:
                    sc["equip"] = equip_from_text(fj)
                if sc["equip"]:
                    sc["equip_raw"] = sc["equip_raw"] or "/".join(sc["equip"])
                pending = sc
                # —— 一行里**并排摆了好几张子表的标题**（WI-0016 p4：r1 左边
                #    「表1，DWG-0000006…参数表」、右边「表2，DWG-0000001…参数表」）。
                #    `scope_of(整行)` 只认得出第一个，右半那张子表连编号都没有，
                #    后面按子表/设备收窄就全串到左半去了。所以逐个单元格记下列号，
                #    建块时按**列区间**各领各的。
                pend_titles = []
                for tc, cell in enumerate(row or []):
                    tx = _c(cell)
                    if not tx:
                        continue
                    s2 = scope_of(tx)
                    tj = _flat(tx)
                    t_title = any(k in tj for k in _V_TITLE_HINT)
                    if t_title and not s2["equip"]:
                        s2["equip"] = equip_from_text(tj)
                    if s2["sub"] or t_title:
                        # 图号也按格记：子表标题里写着「DWG-0000001」，它是
                        # 「同一设备两套参数」唯一的区分线索（见 `_pn_cores`）
                        pend_titles.append((tc, s2, _pn_cores(tx) or None, tx))
                blocks = {}                    # 新的一段：等它的表头
            head = _v_header_cols(row)
            if head:
                keys = sorted(head)
                # 表头行**重新定义**这一段的块：上一张子表留下的块必须清掉。
                # 不清的话（WI-0016 p4 实测）「项目」行建的 name_col=0/5 会一直挂在
                # blocks 里，数据行先撞上它们，抽出来的是 param='1'（序号）这种垃圾，
                # 真正的「参数名称」列反而被顶掉 —— 整页参数名全是序号。
                blocks = {}
                # 每个块的**横向区间**：左边界取「序号 / 项目」那一列（参数名列左边那格），
                # 右边界取下一个块的左边界。区间是后面「按列认领标题 / 按列挂设备」的尺子。
                edges = []
                for c in keys:
                    lt = c - 1
                    if lt >= 0 and _n(row[lt] if lt < len(row) else "") in _V_LEAD_COL:
                        edges.append(lt)
                    else:
                        edges.append(c)
                # 「一块一格」的设备声明行：表头上三行里，找**每个非空格都带设备号**
                # 且格数 == 块数的那一行，第 k 格对第 k 块（见下面 b["equip"] 处的说明）。
                pend_block_equip = []
                for j in range(max(0, ri - 3), ri):
                    rj = rows[j] or []
                    filled = [_c(x) for x in rj if _c(x)]
                    if len(filled) < 2:
                        continue
                    if any(_NOT_EQUIP_ROW.search(_flat(tx)) for tx in filled):
                        continue
                    eqs = [equip_from_text(_flat(tx)) for tx in filled]
                    if eqs and all(eqs):
                        pend_block_equip = [sorted(e) for e in eqs]   # 越近表头越可信
                for i, c in enumerate(keys):
                    left = edges[i]
                    stop = edges[i + 1] if i + 1 < len(edges) else len(row)
                    b = dict(pending or {"sub": "", "equip_raw": "", "equip": [],
                                         "dias": [], "prog": "", "families": []})
                    # 并排子表：按列区间领**自己那张**子表的标题，别用整行的第一个
                    if pend_titles:
                        mine = [it for it in pend_titles if left <= it[0] < stop]
                        if mine:
                            s2 = mine[0][1]
                            b["pn"] = mine[0][2]
                            if s2.get("sub"):
                                b["sub"] = s2["sub"]
                            if s2.get("families"):
                                b["families"] = s2["families"]
                            if s2.get("equip"):
                                b["equip"] = s2["equip"]
                                b["equip_raw"] = s2.get("equip_raw") or "/".join(s2["equip"])
                            if s2.get("dias"):
                                b["dias"] = s2["dias"]
                            # ——标题被 PDF 拆碎时，规格也碎了——
                            # WI-0025 p5：表2-a 的标题被拆成 4 格
                            # （「表2-a，设备/工具」「编号：T-1175激光」「焊接参数表(1.5」「-4.0mm)」），
                            # 逐格 scope_of 谁都认不出「(1.5-4.0mm)」。这里把**本列区间内
                            # 所有标题格拼起来**再取一次，规格才归到它自己那张表。
                            joint = " ".join(it[3] for it in mine if len(it) > 3)
                            if joint:
                                dm = _title_dias(joint)
                                if dm:
                                    b["dias"] = dm
                    # 表头单元格里可能写着适用规格：「参数名称 (… Φ1.5-2.5mm)」
                    d = dias_from_text(head[c])
                    if d:
                        b["dias"] = d
                    s2 = scope_of(head[c])
                    if s2["sub"]:
                        b["sub"] = s2["sub"]
                    if s2["families"]:
                        b["families"] = s2["families"]
                    # ——「每个值列各挂一组设备」（WI-0016 型）——
                    # 表头上面一行常是「项目 | 拉伸机：T-A | 拉伸机：T-B」，**一列一组设备**；
                    # 而 `scope_of` 只认「设备/工具编号：」这种标签，写「拉伸机：」就一个都抓不到。
                    # 逐列往上看三行，哪列自己声明了设备就记下来；建记录时按列拆开各挂各的。
                    ce = {}
                    for j in range(max(0, ri - 3), ri):
                        rj = rows[j] or []
                        if _row_is_not_equip(rj, c):
                            continue                  # 芯轴/保护套/物料行：不是设备
                        for cc in range(c + 1, min(stop, len(rj))):
                            eq = equip_from_text(_flat(_c(rj[cc]) or ""))
                            if eq:
                                ce[cc] = list(eq)      # 越靠近表头越可信 → 后写覆盖前写
                    b["col_equip"] = ce
                    # ——「一块一格」的设备声明行（WI-0019 p5 型）——
                    # 表头上面那行只有 2 格：「设备/工具说明：热焊接机T-2069/…(五台)」
                    # 「设备/工具说明：热焊接机T-2069/T-2057(两台)」，而下面的表头有 6 列
                    # （两块各 3 列）—— 按列号认领会两张表的设备全落到左块上。
                    # 判据：这一行的**每个非空格都带设备号**，且格数 == 下面的块数
                    # → 第 k 格对第 k 块。这是「每块各声明一组设备」的版式。
                    if pend_block_equip and len(pend_block_equip) == len(keys):
                        b["equip"] = list(pend_block_equip[keys.index(c)])
                        b["equip_raw"] = "/".join(b["equip"])
                    b["left"] = left
                    b["stop"] = stop
                    # 逻辑块编号（0 起）：一张物理表里并排放两张子表时（WI-0016 p4
                    # 左右各一张），两条记录的 (页, 表号) 完全一样，只有块号能分开。
                    # 它同时也是「表外标题该归哪张小表」的抓手 —— 见 `_layout_call`。
                    b["blk"] = i
                    # —— 表框外的标题（2026-11-11）——
                    # 设备号 / 子表号 / 直径档经常只写在**表框外面**那一行标题里：
                    # WI-0029 第 7 页三张表，表内一格都没写设备，全靠表上方那行
                    # 「表2-i，设备/工具编号：T-10058/T-1362/T-10161 激光焊接参数表(Ø4.5-6.0)」。
                    # 只读表内 → 这三个块的 equip/sub/dias 全空，后面只能判「找不到设备」，
                    # 而值明明就在格子里躺着。
                    #
                    # 为什么是「补」不是「特例」：`_stepcol_page` 早就在按坐标往上找标题
                    # （`_title_above`），竖表这条链一直没接上。同一份 PDF 里，
                    # 步骤列表能认出设备号、竖表认不出 —— 这不是文档的差别，是两条链的差别。
                    #
                    # 只在**表内确实没写**的维度上补，绝不覆盖表内已有的信息：
                    # 表内写了就以表内为准（那是这张表自己的声明，比头顶的标题更近）。
                    if not (b.get("equip") and b.get("sub") and b.get("dias")):
                        if _plines is None:
                            _plines = _page_word_lines(page)
                        tline = _title_seg_above(_plines, _row_top(t, ri),
                                                 _span_xrange(t, ri, left, stop))
                        if tline:
                            s3 = scope_of(tline)
                            if not b.get("sub") and s3.get("sub"):
                                b["sub"] = s3["sub"]
                            if not b.get("dias"):
                                d3 = _title_dias(tline) or (s3.get("dias") or [])
                                if d3:
                                    b["dias"] = list(d3)
                            if not b.get("equip"):
                                e3 = (equip_from_text(_flat(tline))
                                      or (s3.get("equip") or []))
                                if e3:
                                    b["equip"] = list(e3)
                                    b["equip_raw"] = "/".join(b["equip"])
                            if not b.get("families") and s3.get("families"):
                                b["families"] = s3["families"]
                            b["title_out"] = tline      # 出处：这句话是从表外认来的
                    blocks[c] = b
                continue
            got = False
            for c, b in blocks.items():
                ce = b.get("col_equip") or {}
                if ce:
                    # 每个值列各挂一组设备 → 按列拆成多条记录，一条一列
                    items = _v_data_row_cols(row, c, b.get("stop", len(row)),
                                             ce, b.get("left"))
                    for (col, val, eq) in items:
                        recs.append({
                            "page": pgno, "table": tno,
                            "sub": b.get("sub", ""), "equip": eq or b.get("equip") or [],
                            "equip_raw": ("/".join(eq) if eq else b.get("equip_raw", "")),
                            "dia": (b.get("dias") or [None])[0],
                            "dias": b.get("dias") or [],
                            "prog": b.get("prog", ""),
                            "families": b.get("families") or [],
                            "pn": b.get("pn"),
                            "param": _c(row[c]) if c < len(row) else "",
                            "text": val, "raw": _c(row[c]) if c < len(row) else "",
                            # 来源坐标：第 ri 行，参数名在 c 列，值在 col 列
                            "_row": ri, "_col": c, "_vcol": col,
                            "_blk": b.get("blk"), "_box": _tbox(t),
                        })
                    if items:
                        got = True
                    continue
                one = _v_data_row(row, c, b.get("stop", len(row)))
                if one:
                    name, vals = one
                    recs.append({
                        "page": pgno, "table": tno,
                        "sub": b.get("sub", ""), "equip": b.get("equip") or [],
                        "equip_raw": b.get("equip_raw", ""),
                        "dia": (b.get("dias") or [None])[0],
                        "dias": b.get("dias") or [],
                        "prog": b.get("prog", ""),
                        "families": b.get("families") or [],
                        "pn": b.get("pn"),
                        "param": name, "text": " ; ".join(vals), "raw": _c(row[c]),
                        # 来源坐标：第 ri 行，参数名在 c 列。
                        # `_vcol` 留空 —— 这一行的值可能散在 c 右边好几个格子里，
                        # 由 `_v_data_row` 拼成 ' ; ' 串，单个列号说不清。
                        "_row": ri, "_col": c, "_vcol": None,
                        "_blk": b.get("blk"), "_box": _tbox(t),
                    })
                    got = True
                    # ⚠️ 这里原来是 break —— 并排摆两张子表时（WI-0016 p4 左右各一张），
                    # 每一行只喂给第一个块，右半那张子表**一条记录都出不来**，
                    # 于是它独有的参数（DWG-0000004 的 70-90℃/10-12mm/s）被报成「PDF 里没有」。
                    # 一行属于几个块就该出几条，改成 continue 全块都走。
            if got:
                continue
            # 没有表头也没拿到的裸竖表行：用最近读到的适用范围兜底（WI-0007 型）
            imp = _v_implicit_row(row)
            if imp and pending:
                c, name, vals = imp
                recs.append({
                    "page": pgno, "table": tno,
                    "sub": pending.get("sub", ""), "equip": pending.get("equip") or [],
                    "equip_raw": pending.get("equip_raw", ""),
                    "dia": (pending.get("dias") or [None])[0],
                    "dias": pending.get("dias") or [],
                    "prog": pending.get("prog", ""),
                    "families": pending.get("families") or [],
                    "param": name, "text": " ; ".join(vals), "raw": _c(row[c]),
                    # 来源坐标：裸竖表兜底行 —— 没有表头，所以没有块号；行/列照记。
                    "_row": ri, "_col": c, "_vcol": None, "_blk": None,
                    "_box": _tbox(t),
                })
    return recs


def extract_vertical(pdf_path):
    """抽取全部竖表参数，并带上「适用范围」。

    每条记录：page / table / sub(附表编号) / equip{设备集合} / equip_raw / dia(直径)
              / dias / prog(程序号) / param / text(值原文) / raw

    与 extract_stepwise / extract_program_tables 的产物是同一套字段语义 ——
    三个版式走同一条收窄链，才不会「修了一个 WI、换一个 WI 又漏」。
    """
    recs = []
    global LAST_FAILED_PAGES
    LAST_FAILED_PAGES = []
    # 正文里的参数值要跟设备清单的「类别名 → 编号」配合，先扫一遍清单
    try:
        name2eq = _equip_name_map(pdf_path)
    except Exception:
        name2eq = []
    with open_pdf(pdf_path) as pdf:
        for pno, page in enumerate(pdf.pages, 1):
            try:
                # 版式派发：按 LAYOUT_REGISTRY 里 vertical 通道的顺序依次跑
                # （竖表 → 裸键值 → 表头横表 → 步骤列 → 正文叙述）。
                # 它们互不独占、逐条累加；某一个抛异常只跳过它自己
                # （以前是一个 try 包住全部五个，坏一个该页全丢）。
                recs.extend(_layouts_for_page(page, pno, "vertical", name2eq))
            except Exception as e:                  # 单页失败不拖垮整份 PDF
                LAST_FAILED_PAGES.append((pno, "%s: %s" % (type(e).__name__, e)))
    return recs


_STEPCOL_RE = re.compile(r"[Ss]tep\s*(\d+)")
# 认「步骤列头」专用，比 `_STEPCOL_RE` 宽：中文写法「步骤1：动态焊」也算。
#
# 为什么必须分成两个正则（不能直接把中文加进 `_STEPCOL_RE`）：
#   `_STEPCOL_RE` 有两处用途，语义相反 ——
#     ① 认表头的步骤列（需要认中文，因为 WI-0013 的跨页续表用「步骤1：动态焊」，
#        而同一份 PDF 前面的表用「Step 1:动态焊」，只认英文就把整张表跳过，
#        连带 Excel 判成「PDF 没有这一步」，正确值被报成「多填」= 假差异）；
#     ② 扫描数据行时判断「是不是撞到下一张子表的步骤行」（**不能**认中文）——
#        参数名写成「步骤1的激光能量」时，含中文的写法会把这一行当成新表头，
#        整块数据提前结束。实测直接把 WI-0012 的一致数从 46 打到 34。
_STEPCOL_HEAD_RE = re.compile(r"(?:[Ss]tep|步骤|工步|工序)\s*(\d+)")
# 子表「参数名列」的锚点：一行里出现几次，就说明这一行并排放了几张子表。
# `参数步骤` 是 WI-0002 的写法，`步骤` 是 WI-0006 的写法。**必须排除值格里
# 恰好含「步骤」二字的情况** —— 所以要求整格以它开头、且不含 Step 数字。
_STEP_ANCHOR_RE = re.compile(r"^\s*(?:参数步骤|参数|步骤|工步)\s*(?:[:：].*)?$")
# 值里写着的适用设备组：WI-0006 p5 的「焊接距离」一格塞了两段
# 「(0.9~1.3)mm(forT-1203/T-1281/T-1381/T-1011)」「(1.2-1.6)mm forT-1485/…」，
# 一张表七台设备共用同一组 Step 列，只能靠这个尾巴把值归到正确的设备组。
_FOR_EQ_RE = re.compile(r"for\s*([Tt]\s*[-–]?\s*[\d]{2,6}(?:\s*[/／、,，]\s*[Tt]?\s*[-–]?\s*[\d]{2,6})*)")


def _row_top(table, i):
    """表格第 i 行的上边 y（拿不到就返回 None，退回老的行序逻辑）。"""
    try:
        return table.rows[i].bbox[1]
    except Exception:
        return None


def _line_in_x(ln, xr):
    """这一行的**横向中点**是否落在 xr 区间内。

    并排双列（WI-0002 p4：左 表1-a / 右 表1-b）时两张子表 **y 完全相同**，
    只按「头顶最近的一行」认领直径必然串台 —— 实测卡0001(Ø2.0-3.5) 拿到的
    却是右边 表1-b 的 3.75-5.00mm。所以认领必须同时带上 x。
    """
    if not xr:
        return True
    x0 = ln.get("x0")
    x1 = ln.get("x1")
    if x0 is None or x1 is None:
        return True
    cx = (float(x0) + float(x1)) / 2.0
    return (float(xr[0]) - 2.0) <= cx <= (float(xr[1]) + 2.0)


def _span_xrange(table, i, c0, c1):
    """第 i 行第 c0..c1-1 列覆盖的横向区间 (x_left, x_right)。

    拿不到单元格坐标就返回 None —— 此时退回「不按 x 过滤」的老行为，
    宁可拿到旧的（可能不精确的）答案，也不要因为几何缺失变成「无对应」。
    """
    try:
        cells = table.rows[i].cells or []
    except Exception:
        return None
    xs = []
    for c in range(c0, min(c1, len(cells))):
        cell = cells[c]
        if not cell or len(cell) < 3:
            continue
        if cell[0] is not None:
            xs.append(float(cell[0]))
        if cell[2] is not None:
            xs.append(float(cell[2]))
    if len(xs) < 2:
        return None
    return (min(xs), max(xs))


def _title_above(lines, y0, xr=None, max_up=_DIA_MAX_UP):
    """y0 上方最近的一行**子表标题**（含「表1-c」或直径档的那种行）。

    子表编号跟直径档是同一个病：都写在标题行上，而标题行与内容行在
    pdfplumber 的检测里常常错开。既然已经按 y 找直径，编号就顺便一起领走。
    xr 不为空时只认横向落在这一段里的标题（并排双列必须，见 `_line_in_x`）。
    """
    if y0 is None or not lines:
        return ""
    best, best_y = "", None
    for ln in lines:
        top = ln.get("top")
        if top is None or top > y0 or y0 - top > max_up:
            continue
        t = ln.get("text") or ""
        if not re.search(r"表\s*\d|附表\s*\d", t):
            continue
        if not _line_in_x(ln, xr):
            continue
        if best_y is None or top > best_y:
            best, best_y = t, top
    if best or not xr:
        return best
    return _title_above(lines, y0, None, max_up)      # x 过滤后一条都没有 → 退回不按 x


def _page_word_lines(page):
    """这一页的词级文本行（每行带 top / x0 / x1 和每个词的坐标）。

    与 `page.extract_text_lines()` 的差别：那个把**同一高度上并排的几段**
    拼成一整行，认标题时就没法知道「哪一段管哪一栏」。这里借 `pdf_tables._lines`
    （它已经按 y 聚行、保留词级 x），才能往下切段。
    """
    try:
        import pdf_tables
        return pdf_tables._lines(page) or []
    except Exception:          # noqa: BLE001 —— 拿不到就当没有表外标题
        return []


def _title_seg_above(plines, y0, xr, max_up=90.0):
    """y0 上方最近那行标题里，**横向落在 xr 内**的那一段文字。

    为什么不能直接用 `_title_above`（2026-11-11 实测 WI-0029 p7）：

    那一页左边「表2-i，设备/工具编号：T-10058/T-1362/T-10161 …(Ø4.5-6.0)」
    和右边「表2-j，设备/工具编号：T-10058/T-1362 …(Ø7.0-8.0)」
    **y 完全相同**（87.8），`extract_text_lines()` 把两段拼成一行。
    整行丢给 `scope_of` 只会认第一个 —— 结果右表拿到了左表的
    `T-10161` 和 `Ø4.5-6.0`。串台比认不出更糟：它造出来的是一条
    看着很像真的错误依据。

    所以按词级 x 把这一行切段（复用 `pdf_tables._segments`，它已经处理过
    「表号被 PDF 拆碎」「词间距断句」这些坑），只取与本块横向重叠最大的那段。

    `max_up` 取 90pt：标题通常紧贴表框上方十几点。`_title_above` 默认 220
    是为「表格错位好几行」留的余量，这里不需要那么宽 —— 捞太远会把
    上一张表的标题认成自己的。
    """
    if y0 is None or not plines:
        return ""
    try:
        import pdf_tables
    except Exception:          # noqa: BLE001
        return ""
    best, best_y = "", None
    for ln in plines:
        top = ln.get("top")
        if top is None or top > y0 or y0 - top > max_up:
            continue
        txt = ln.get("text") or ""
        if not re.search(r"表\s*\d|附表\s*\d", txt):
            continue
        lx0, lx1 = ln.get("x0"), ln.get("x1")
        if xr and lx0 is not None and lx1 is not None:
            # 横向一点不重叠就不是这一栏的标题（跨两栏的整行仍然算，下面切段再挑）
            if min(float(lx1), float(xr[1])) - max(float(lx0), float(xr[0])) <= 0:
                continue
        segs = pdf_tables._segments(ln) or []
        if not segs:
            continue
        if xr:
            mine = [(min(s["x1"], xr[1]) - max(s["x0"], xr[0]), s) for s in segs]
            mine = [m for m in mine if m[0] > 0]
        else:
            mine = [(1.0, s) for s in segs]
        if not mine:
            continue
        if best_y is None or top > best_y:       # 越靠近表头越好
            best_y = top
            best = max(mine, key=lambda m: m[0])[1]["text"]
    return best


def _block_title(plines, table, row_i, xr, max_up=90.0):
    """本块（一张表 / 并排两表里的一栏）**表框外**的标题文字。

    为什么必须按「表框 top + x 区间」两级定位（2026-11-11 WI-0029 p7 实测）：
    t1(表2-i, Ø4.5-6.0) 与 t2(表2-j, Ø7.0-8.0) 的标题在**同一行、同一 y**，
    `extract_text_lines()` 把它拼成「表2-i…(Ø4.5-6.0mm) 表2-j…(Ø7.0-8.0mm)」。
    整行丢给 `dias_from_text` 会得到 [(4.5,6.0),(7.0,8.0)]，而记录只存
    `dia=(dias or [None])[0]` —— **取第一个** → 右表被标成 Ø4.5-6.0。

    后果不只是规格写错：t1/t2 的设备号完全一样（都是 T-1362/T-10058/T-10161），
    这时**规格是唯一能把「Excel 这一行该看哪张表」定死的信号**。它一错，
    两张表的值就被并进同一个期望值集合 —— 激光能量的 480-560 与 0 同时出现，
    判定只能报 multi 或凭空造出「不一致」。所以这不是显示问题，是判定链的第二层。

    两级定位缺一不可：只按行 → 拿到整行（串台）；只按段 → 会捞到上一张表的标题。
    """
    if not plines:
        return ""
    y0 = None
    try:
        bb = getattr(table, "bbox", None)
        if bb:
            y0 = float(bb[1])
    except Exception:          # noqa: BLE001 —— 拿不到几何就不写，不影响抽取
        y0 = None
    if y0 is None:
        y0 = _row_top(table, row_i)
    txt = _title_seg_above(plines, y0, xr, max_up=max_up)
    if txt:
        return txt
    # 表框顶没找到（个别表标题离得远）→ 退回按「表头行上方」再找一次
    y1 = _row_top(table, row_i)
    if y1 is not None and y1 != y0:
        return _title_seg_above(plines, y1, xr, max_up=max_up)
    return ""


def _dia_above(lines, y0, xr=None, max_up=_DIA_MAX_UP):
    """在 y0 **上方最近**的一行里找直径档（Ø2.0-3.0mm / φ3.25-5.0 …）。

    为什么不能用「向上数几行」：pdfplumber 的检测顺序跟视觉顺序不是一回事，
    一页里上下堆的两张子表常被划进同一个 table，标题行与内容行会错开好几行。
    谁的标题在它头顶上、离得最近，就归谁 —— 这是唯一靠得住的判据。
    并排双列时还得再叠一层 x 区间，否则左右两表 y 相同必然串台。
    """
    if y0 is None or not lines:
        return []
    best, best_y = None, None
    for ln in lines:
        top = ln.get("top")
        if top is None or top > y0 or y0 - top > max_up:
            continue
        d = dias_from_text(ln.get("text") or "")
        if not d:
            continue
        if not _line_in_x(ln, xr):
            continue
        if best_y is None or top > best_y:      # 越靠近越好
            best, best_y = d, top
    if best or not xr:
        return list(best or [])
    return _dia_above(lines, y0, None, max_up)        # 同上：过滤空了就退回


def _split_for_groups(val):
    """把「(0.9~1.3)mm(forT-1203/T-1281) (1.2-1.6)mm forT-1485/…」切成若干段，
    每段自带它适用的设备组；一段都没有就原样返回（设备组沿用整张表的）。"""
    ms = list(_FOR_EQ_RE.finditer(val))
    if not ms:
        return [val]
    out, start = [], 0
    for m in ms:
        piece = val[start:m.end()]
        if _c(piece):
            out.append(piece)
        start = m.end()
    tail = val[start:]
    if _c(tail):
        out.append(tail)
    return out or [val]


def _stepcol_page(page, pgno):
    """第四种版式：**参数名在行、步骤在列**（WI-0006 型）。

    pdfplumber 抽出来长这样::

        ['激光焊/接机', 'T-10010', '', '', '']
        ['程序号',      '13',      '', '', '']
        ['步骤',        '', 'Step0', 'Step1', 'Step2']
        ['激光功率',     '', '0 mw',  '350-460mw', '340-450mw']
        ['转速',        '', '1000 rpm', …]

    它**既不是竖表也不是程序附表**：
      · `_vertical_page` 要见到「参数名称 / 项目 / 参数」表头行才开块，这里没有；
        裸行兜底 `_v_implicit_row` 又要求 row[0] 是 1-2 位数字序号，而这里 row[0]
        是参数名 → 竖表路径 0 条。
      · `_program_page` 要的是**步骤在行、程序在列**（正好是这里的转置），且数据行
        row[0] 必须是步骤号 → 程序附表路径 0 条。
    于是三个抽取器齐刷刷返回 0，整份 WI-0006 一个参数点都出不来 —— 那份 PDF 的
    文本层是完整的、表格也是方的，**漏的是版式，不是字**。
    """
    recs = []
    try:
        tables = page.find_tables() or []
    except Exception:
        return recs
    try:                                   # 视觉顺序的文本行，用来按位置认领直径档
        lines = page.extract_text_lines() or []
    except Exception:
        lines = []
    # 词级文本行（带每个词的 x）：并排两表的标题在**同一行同一高度**时，
    # `lines` 会把左右两段拼成一整行，只有词级坐标才切得开（见 `_block_title`）。
    plines = _page_word_lines(page)
    for tno, t in enumerate(tables):
        try:
            rows = t.extract() or []
        except Exception:
            continue
        for i, row in enumerate(rows):
            # —— 先把「左右并排」的双列拆开（WI-0002 型）——
            # 一张 table 里横向摆两张子表，行看起来是这样的：
            #   ['参数步骤','Step 1','Step 2','Step 3','', '参数步骤','Step 1',…]
            #   ['激光焊接机','T-1112','','','',            '激光焊接机','T-1112',…]
            # 不拆开的话：左右两表的步骤列会被当成同一张表的 6 步，右表的直径/程序号
            # 会串到左表上 —— 实测 WI-0002 卡0001(φ2.0-3.5) 拿到的却是 3.75-5.0 的表。
            # 判据：一行里出现**多个**「参数步骤 / 步骤」锚点，每个锚点起一张子表。
            anchors = [c for c in range(len(row or []))
                       if _STEP_ANCHOR_RE.match(_c(row[c]))]
            spans = []
            if anchors:
                for k, a in enumerate(anchors):
                    b = anchors[k + 1] if k + 1 < len(anchors) else len(row or [])
                    spans.append((a, b))
            else:
                # 老样子：整行一张表（WI-0006 等），参数名在第 0 列
                head0 = _n(row[0]) if row and row[0] else ""
                if "步骤" in head0 or "工步" in head0:
                    spans.append((0, len(row or [])))
                else:
                    # WI-0013 表1-s / 表1-t 这类：步骤列头行的**左边两格是空的**，
                    # 长这样 ['', '', '步骤1：动态焊', '步骤2：圆周焊', …]。
                    # 以前只认 row[0] 里写着「步骤/工步」，这一行第一个字是空串就
                    # 整张 continue 掉 —— 连带 Excel 判「PDF 里没有这一步」，
                    # 正确值反被报成「多填」= 假差异。
                    # 改判据：row[0] 认不出时，看**整行**里有几个步骤列头，≥2 个算数。
                    hit = sum(1 for c in range(len(row or []))
                              if _STEPCOL_HEAD_RE.search(_c(row[c])))
                    if hit < 2:
                        continue
                    spans.append((0, len(row or [])))
            for (c0, c1) in spans:
                steps = {}
                for c in range(c0, min(c1, len(row or []))):
                    # 认列头用**宽**正则（认中文「步骤1：动态焊」）；
                    # 下面扫数据行判「撞没撞到新表头」仍用窄的 `_STEPCOL_RE`，
                    # 否则参数名叫「步骤1的激光能量」的行会被当成新表头，提前收摊。
                    m = _STEPCOL_HEAD_RE.search(_c(row[c]))
                    if m:
                        steps[c] = int(m.group(1))
                if len(steps) < 2:
                    continue
                # —— 参数名在哪一列 ——
                # 默认第 0 列（WI-0006 型：row[0] 本身就是参数名）。但 WI-0013
                # 表1-s 型前面多了一列「序号」，参数名被挤到第 2 列：数据行看起来是
                # ['1', '激光能量', '0mw', '550-600mw', '0mw', '620-660mw']。
                # 若还按第 0 列读，参数名会变成序号「1」，四条值全挂在一个叫「1」
                # 的参数上 —— 等于白读。所以往上找那行「序号 | 参数名称 | 参数值或范围」
                # 的表头，取它在步骤列左边的那一列作为参数名列。
                pcol = c0
                _c_first_step = min(steps)
                for j in range(i - 1, max(-1, i - 4), -1):
                    _cand = [c for c in _v_header_cols(rows[j] or [])
                             if c0 <= c < _c_first_step]
                    if _cand:
                        pcol = min(_cand)
                        break
                # 子表表头**声明了哪几步**。必须单独记下来：某一步的值是 NA / 空时
                # 建点阶段就被丢掉了，只看幸存记录的步骤号会把「0 起算表」误判成
                # 「1 起算表」（WI-0002 p6 表1-h：焊接时间只有 Step1 有值，
                # 幸存步骤集 = {1} → min≠0 → 不挪位 → Excel 步骤2 找不到 → 报
                # 「该表里没有步骤2」，其实它在 Step1 上好好写着 3.0-3.5s）。
                _steps_all = tuple(sorted(set(steps.values())))
                # —— 认领适用范围：设备号 / 程序号 / 直径档 / 子表编号
                equip_raw, prog, sub, dias = "", "", "", []
                # 本子表**横向**占的那一段（并排双列时左/右各一段）。
                xr = _span_xrange(t, i, c0, c1)
                # ① 先读**子表自己的行**。标题行常常就躺在表格里 —— WI-0002 p4 的
                #    r1 一边是「表1-a…Ø2.0-3.5mm」、另一边是「表1-b…Ø3.75-5.0mm」，
                #    按列区间切开后左右各拿各的，这是最可信的一路。
                #    ⚠️ 早先的写法是**先**按 y 到页面文本行里猜、猜到了就不再看行，
                #    而并排两表 y 相同 → 永远猜中右边那张，左边整张规格全错。
                #    规格号是三层匹配（设备→规格→步骤）的第二层，错了后面全错。
                # ② 表格里没写（标题在表格外的正文行）→ 再按 y **加 x 区间**去页面上找。
                # 上看 6 行（原为 3 行）：并排子表的标题常被推到离表头第 4~6 行
                # （WI-0025 p5：表头在 r7，表2-a/表2-b 的标题在 r2），只看 3 行的话
                # 右半张表既认不到子表号、也认不到它自己的规格档。
                for j in range(max(0, i - 6), i):
                    rj = rows[j] or []
                    seg = rj[c0:min(c1, len(rj))]           # 只取本子表的列区间
                    joined = " ".join(_c(x) for x in seg if _c(x))
                    if not joined:
                        continue
                    # 上下相邻两张子表被 pdfplumber 并进同一个 cell（表7-a…表7-b…）
                    # 时，scope_of 只取第一个标题，但本 table 的数据对应**最后一张**。
                    # 用最后一个标题把 sub / 直径档更正过来，否则表7-b 的值会错误
                    # 沿用表7-a 的直径档（WI-0006 p11/p12：540-620 被当成 Ø2.0-3.0）。
                    _lsub, _ldia = _last_sub_title(joined)
                    if not sub:
                        if _lsub:
                            sub = _lsub
                        else:
                            s2 = scope_of(joined)
                            if s2.get("sub"):
                                sub = s2["sub"]
                    if not dias:
                        if _ldia:
                            dias = [_ldia]
                        else:
                            # 只看**像标题**的行：上看 6 行会扫到上一张子表的数据行，
                            # 「移动速度 0.5mm/s」里的 0.5mm 会被当成直径档（实测表2-c
                            # 因此拿到 (0.5,0.5) 而不是 (1.5,4.0)，规格层直接错档）。
                            tj = _flat(joined)
                            if (_V_SUB_RE.search(tj) or "设备" in tj
                                    or any(k in tj for k in _V_TITLE_HINT)):
                                # 标题里的规格可能不带 mm（「…参数表(4.5-6.0)」），
                                # 用标题专用的宽松口径，否则右半张表沿用左半张的直径档。
                                d2 = _title_dias(joined)
                                if d2:
                                    dias = list(d2)
                    if not prog and seg and "程序号" in _n(seg[0]):
                        prog = " ".join(_c(x) for x in seg[1:] if _c(x)).strip()
                # 设备单独往上多看几行：有的表把设备写在离表头第 4~6 行的地方
                # （WI-0025 p5：「激光机型号 T-1175」/ 子表标题「设备/工具编号：T-1175…
                # 参数表(1.5-4.0mm)」都在表头上面 4~5 行，只看 3 行就一台都认不到）。
                # 仍然**只取本子表的列区间**，所以不会抓到并排那张表的设备。
                for j in range(max(0, i - 6), i):
                    rj = rows[j] or []
                    seg = rj[c0:min(c1, len(rj))]
                    joined = " ".join(_c(x) for x in seg if _c(x))
                    if not joined or _row_is_not_equip(seg, 0):
                        # 「使用芯轴 T-2015-06」里的 T-2015 是**芯轴编号，不是设备**。
                        # 不挡掉就会先撞上这一行，整套参数挂到 T-2015 名下。
                        continue
                    eq = equip_from_text(joined)
                    if eq:
                        equip_raw = "/".join(eq)     # 后写覆盖 → 取离表头最近的一行
                # ② 表格行里没认领到编号/规格 → 到**表框外**的标题里按几何位置找。
                #
                # ⚠️ 不能再用整行文本（原写法 `_title_above` / `_dia_above`）：
                # 并排两表的标题在同一行、同一 y（WI-0029 p7 的 表2-i / 表2-j
                # 都在 y=87.8），`extract_text_lines()` 把两段拼成一行 →
                # 两张表都解析出两个规格，而 `dia` 只取第一个 →
                # 右表(Ø7.0-8.0)被标成 Ø4.5-6.0，与左表设备号又完全相同，
                # 于是两表的值被并进同一个期望值集合（480-560 与 0 同时成为候选）。
                # `_block_title` 按表框 x 区间切段，只取与本块重叠最大的那段。
                if not sub or not dias:
                    _bt = _block_title(plines, t, i, xr)
                    if _bt:
                        if not sub:
                            s2 = scope_of(_bt)
                            if s2.get("sub"):
                                sub = s2["sub"]
                        if not dias:
                            d2 = _title_dias(_bt)
                            if d2:
                                dias = list(d2)
                for j in range(i + 1, len(rows)):
                    rj = rows[j] or []
                    name = _c(rj[pcol]) if pcol < len(rj) else ""
                    if not name:
                        continue
                    nn = _n(name)
                    # 撞上下一张子表的表头 / 又一个步骤行 → 本块结束
                    if (nn.startswith("激光焊") or "程序号" in nn or nn in ("步骤", "工步")
                            or _STEPCOL_RE.search(name) or re.match(r"^表\s*\d", name)
                            or _v_header_cols(rj)):
                        break
                    for c, s in sorted(steps.items()):
                        if c >= len(rj):
                            continue
                        val = _c(rj[c])
                        if not val:
                            continue
                        # 一格可能塞了**好几段**「值 + forT-设备组」（WI-0006 p5 的焊接
                        # 距离：「(0.9~1.3)mm(forT-1203/T-1281/T-1381/T-1011)
                        # (1.2-1.6)mm forT-1485/…」）。只取第一段的话，第二组设备的
                        # 期望值就永远查不到 → 判成 no_equip_match「PDF 里没有该设备
                        # 对应的参数」，那是**假漏判**：值明明在格子里躺着，
                        # 只是被前一段的正则吃掉了。按每个 for 段把整格切开，
                        # 每段各自认领自己的设备组。
                        for piece in _split_for_groups(val):
                            # 把 forT-… 尾巴从「值」里剥掉：它只是适用范围，不是参数值。
                            # 留着会被数值解析器读成区间「1203-1281」（设备号被当成参数），
                            # 混进期望值集合 —— 那是**假一致**的温床：Excel 真填 1203 也能过。
                            m = _FOR_EQ_RE.search(piece)
                            body = (piece[:m.start()] if m else piece).strip()
                            if not body or not re.search(r"\d", body):
                                continue                # 只剩括号/标点的碎屑
                            eq2 = equip_raw
                            g = equip_from_text(piece)
                            if g and m:
                                eq2 = "/".join(g)
                            # 值解析成数值：步骤链要的是 (lo, hi, unit, note) 这四样，
                            # 而这条链原来只带 `text` 原文。不补的话它能进 vertical 通道、
                            # 进不了步骤链 —— 记录形状不一致，归一那一步就得靠猜
                            # （上次窄验证就是在这里把 35 条全变成「多填」的）。
                            _pv = parse_step_value(body) or (None, None, "", "")
                            recs.append({
                                "page": pgno, "table": tno, "sub": sub,
                                "equip": _equip_set_spec(eq2), "equip_raw": eq2,
                                "dia": (dias or [None])[0], "dias": dias or [],
                                "prog": prog, "families": families_of(body) or [],
                                "step": s, "param": name,
                                "steps_all": _steps_all,
                                "text": body, "raw": val,
                                "lo": _pv[0], "hi": _pv[1],
                                "unit": _pv[2] or "", "note": _pv[3] or "",
                                # 来源坐标：第 j 行，参数名在 pcol 列，值在步骤列 c
                                "_row": j, "_col": pcol, "_vcol": c,
                                "_box": _tbox(t),
                            })
    return recs


# ⚠️ 这里原本有一个 `_vertical_summary(pdf_path, max_items=10)`（2026-09-30 重构删除）。
# 它按 (子表, 设备, 直径) 分组打印「竖表参数块摘要」，但三处都查不到调用方 ——
# 竖表自检实际上走 `self_check` / `pdf_read_audit`，那份报告是按版式统计的，
# 比这份分组清单更全。留着只会让人以为自检有两条路。


# ==================================================== C 型：程序参数附表（转置双列）
#
# 版式特征（WI-0009 亲水涂层 / WI-3852 亲水自动线 等）：
#
#     工序名称 | 设备参数设置附表8 | 文件编号 | WI-0009 | …        ← 附表编号在页眉表
#     设备/工具编号：T-10013、T10014 | 设备/工具说明：亲水涂层机
#     程序号/名称 | 8
#     步骤 | 参数名称 | 参数值或范围
#          |          | 程序 A (底层溶液) | 程序 B (表层溶液)     ← 两列 = 两套期望值
#     7    | 提拉上升01 | 230mm，19mm/s  | 230mm，19mm/s
#
# 为什么必须把这张表单独解析：它的「期望值」被三个维度切成很多份 ——
# 附表编号（设备/程序号/海波管型号）× 程序列（底层/表层溶液）× 步骤行。
# 只按参数名抽、把同名值并成一个大集合的话：
#   · WI-3852「表层溶液 固化时间」的候选里会混进底层溶液的 40~60s；
#   · WI-0009「产品上提01」会出现 13 个候选（跨全部 14 张附表）；
# 于是「填错到隔壁列/隔壁附表」也会被judged成一致 —— 这种假一致最危险。
# 所以这里把 (附表, 程序列, 步骤, 参数, 子标签) 全部如实带出来，交给上层去收窄。

_SUBTABLE_NAME = re.compile(r"设备参数\s*(?:设置)?\s*附表\s*([0-9]+(?:\s*[-－]\s*[0-9A-Za-z]{1,3})?)")
_EQUIP_LABEL = re.compile(r"设备\s*/\s*工具编号\s*[：:]\s*(.+?)(?:\s*设备\s*/\s*工具说明|\s*$)")
_PROG_NO_LABEL = re.compile(r"^程序号")
_DWG_RE = re.compile(r"DWG\s*\d+")
_PROG_COL_RE = re.compile(r"程序\s*([A-Za-zＡ-Ｚａ-ｚ]{1,2})?\s*[（(]\s*([^）)]{0,16}?)\s*[）)]")
_SOLUTION_RE = re.compile(r"(底层溶液|表层溶液)")
_HEAD_STEP = re.compile(r"^(?:步骤|step)$", re.I)
_HEAD_PARAM = re.compile(r"^参数\s*名称")
# 「位置: 357.3mm」「速度：10-15mm/s」——短标签 + 值
_SEG_LABEL = re.compile(r"^([\u4e00-\u9fff]{1,5}|[A-Za-z]{2,10})\s*[:：]\s*(.+)$", re.S)


def split_segments(text):
    """把一个值单元格拆成 [(子标签, 值原文)]。

    '位置: 357.3mm; 速度: 10-15mm/s;' → [('位置','357.3mm'), ('速度','10-15mm/s')]
    '230mm，19mm/s'                   → [('','230mm'), ('','19mm/s')]
    '（100-120s）x2'                  → [('','（100-120s）x2')]

    带标签的段（位置/速度）必须保留标签：同一行里「位置 408mm」和「速度 20mm/s」
    是两个不同的参数，混在一起正是「位置被当成速度」的老毛病。
    """
    s = _c(text)
    if not s or _is_blank(s):
        return []
    out = []
    for piece in re.split(r"[;；\n]+", s):
        piece = piece.strip()
        if not piece:
            continue
        # 没有冒号的段（如「365mm，20mm/s」）按逗号再拆，让单位和值一一对应
        subs = [piece]
        if ":" not in piece and "：" not in piece:
            subs = [q.strip() for q in re.split(r"[，,]+", piece) if q.strip()]
        for q in subs:
            m = _SEG_LABEL.match(q)
            out.append((m.group(1), m.group(2).strip()) if m else ("", q))
    return out


def _program_cols(row):
    """识别「程序 A (底层溶液) | 程序 B (表层溶液)」列标题行。

    合并单元格会让右半列读成 None → 向前继承上一列的标签（否则右半子表丢列）。
    """
    cols, last = {}, None
    for c in range(len(row)):
        cell = _c(row[c])
        if not cell:
            if last:
                cols[c] = last
            continue
        m = _PROG_COL_RE.search(cell) or re.search(r"程序\s*([A-Za-z])", cell)
        if not m:
            last = None                    # 出现别的列内容，停止继承
            continue
        letter = (m.group(1) or "").upper() or None
        sol = ""
        sm = _SOLUTION_RE.search(cell)
        if sm:
            sol = sm.group(1)
        last = (letter, sol, cell)
        cols[c] = last
    return cols


def _program_meta(rows, head_i):
    """读表头行上方的元信息：设备编号 / 程序号 / 海波管 DWG。"""
    equip_raw, prog_raw, dwg = "", "", ()
    for j in range(head_i):
        row = rows[j] or []
        joined = " ".join(_c(c) for c in row if c)
        if not joined:
            continue
        m = _EQUIP_LABEL.search(joined)
        if m and not equip_raw:
            # 「T-10013、T10014」这种顿号分隔也要能拆成两台设备
            equip_raw = _c(m.group(1)).replace("、", "/").replace("，", "/")
        if _PROG_NO_LABEL.match(_n(row[0]) if row and row[0] else ""):
            vals = [_c(c) for c in row[1:] if _c(c)]
            prog_raw = prog_raw or " ".join(vals)
        if not prog_raw:
            # 「程序号：5」还可能**挤在同一格的末尾**（WI-0009 p14 附表9：
            # 「设备/工具编号：T-10012 设备/工具说明：亲水涂层机 程序号：5」）。
            # 上面那条只看 row[0] 是不是以「程序号」开头，够不着这种写法 →
            # 整张附表的程序号丢失；而 p15（For 产品线F NC）写成独立一行
            # 「程序号/名称 7」能抽到。于是 T-10012 只剩 7，Excel 填的 5
            # 被判「应改为 7」—— 假不一致。
            m2 = re.search(r"程序号\s*[/／]?\s*(?:名称|No\.?)?\s*[：:]\s*([^\s，,；;（(]*)", joined)
            if m2 and _c(m2.group(1)):
                prog_raw = _c(m2.group(1))
        if not dwg:
            dwg = tuple(dict.fromkeys(_DWG_RE.findall(joined)))
    return equip_raw, prog_raw, dwg


def _program_page(page, pgno):
    """解析一页里的所有「程序参数附表」。"""
    recs = []
    try:
        tables = page.find_tables() or []
    except Exception:
        return recs
    try:                                   # 视觉顺序的文本行，用来认领「For XX」产品线
        lines = page.extract_text_lines() or []
    except Exception:
        lines = []

    sub_raw = ""
    for t in tables:                        # 附表编号写在页眉那一行小表里
        try:
            rows = t.extract() or []
        except Exception:
            continue
        for row in rows[:2]:
            for cell in row or []:
                m = _SUBTABLE_NAME.search(_c(cell)) if cell else None
                if m:
                    sub_raw = "附表" + re.sub(r"\s+", "", m.group(1))
                    break
            if sub_raw:
                break
        if sub_raw:
            break

    for t in tables:
        try:
            rows = t.extract() or []
        except Exception:
            continue
        if len(rows) < 4:
            continue
        head_i = None
        for i, row in enumerate(rows):
            cells = [_n(c) for c in row if c]
            if any(_HEAD_STEP.match(c) for c in cells) and any(_HEAD_PARAM.match(c) for c in cells):
                head_i = i
                break
        if head_i is None:
            continue
        equip_raw, prog_raw, dwg = _program_meta(rows, head_i)
        # 产品线（「For 产品线F NC」「For PTCA」这类写在附表标题里的限定）：
        # 同一台设备可以有多张附表，分别属于不同产品线，程序号/参数值都不同。
        # 不带出来就只能靠设备号收窄 → 两套候选一起留着，Excel 填哪一套都可能
        # 被另一套判不一致。取**头顶最近**的标题行（与 `_dia_above` 同一套路）。
        fams = []
        head_y = None
        try:
            head_y = t.rows[head_i].bbox[1]
        except Exception:
            head_y = None
        best_ty = None
        for ln in lines:
            tx = ln.get("text") or ""
            if not re.search(r"参数表|（\s*For|For\s+\w", tx):
                continue
            ty = ln.get("top")
            if ty is None:
                continue
            if head_y is not None and ty > head_y + 5:
                continue
            if best_ty is None or ty > best_ty:
                fams, best_ty = families_of(tx), ty
        if not dwg:
            # WI-3852 型：海波管型号写在表头那一行（「参数：使用海波管DWG1000015/…」），
            # 它是区分同名附表的关键线索，必须带出来。
            head_joined = " ".join(_c(c) for c in (rows[head_i] or []) if c)
            dwg = tuple(dict.fromkeys(_DWG_RE.findall(head_joined)))
        cols = {}
        for j in range(head_i + 1, len(rows)):
            cand = _program_cols(rows[j] or [])
            if cand:
                cols = cand
                break
        if not cols:
            continue
        for j in range(head_i + 1, len(rows)):
            row = rows[j] or []
            if any(_PROG_COL_RE.search(_c(c)) for c in row if c):
                continue                    # 程序列标题行本身
            n0 = _n(row[0]) if row and row[0] else ""
            if not re.match(r"^\d+(?:/\d+)*$", n0):
                continue
            param = _c(row[1]) if len(row) > 1 else ""
            if not param or _HEAD_PARAM.match(_n(param)):
                continue
            steps = tuple(int(x) for x in n0.split("/"))
            for c, (letter, sol, _raw) in cols.items():
                if c >= len(row):
                    continue
                for label, text in split_segments(row[c]):
                    recs.append({
                        "page": pgno, "sub": sub_raw, "equip_raw": equip_raw,
                        "equip": _equip_set_spec(equip_raw), "prog_no_raw": prog_raw,
                        "families": fams,
                        "dwg": dwg, "program": letter, "solution": sol,
                        "step_raw": n0, "steps": steps, "step": min(steps),
                        "param": param, "label": label, "text": text,
                        "raw": _c(row[c]),
                        # 来源坐标：第 j 行，参数名固定在第 1 列，值在 c 列。
                        # 刻意**不补 `table`** —— 这条链原本就不带表号，补了会让
                        # 表目录的「抽到几条」计数挪窝，属于本步不该有的变化。
                        "_row": j, "_col": 1, "_vcol": c, "_box": _tbox(t),
                    })
    return recs


def extract_program_tables(pdf_path):
    """抽取全部「设备参数(设置)附表N」转置双列附表。

    每条记录：page / sub(附表编号) / equip{eq} / equip_raw / prog_no_raw / dwg
              / program(A/B) / solution(底层溶液/表层溶液) / step{,_raw} / steps
              / param / label(位置/速度/…) / text(值原文) / raw(整格原文)

    值本身不在这里解析 —— 版式层只负责「这段文字属于哪张附表、哪一列、哪一步」，
    数值语义统一由 engine 解析，避免两处实现漂移。
    """
    recs = []
    global LAST_FAILED_PAGES
    LAST_FAILED_PAGES = []
    with open_pdf(pdf_path) as pdf:
        for pno, page in enumerate(pdf.pages, 1):
            try:
                # 版式派发：program 通道（现在只有程序附表一种，加新附表版式
                # 只改 LAYOUT_REGISTRY，不用动这里）
                recs.extend(_layouts_for_page(page, pno, "program"))
            except Exception as e:                  # 单页失败不拖垮整份 PDF
                LAST_FAILED_PAGES.append((pno, "%s: %s" % (type(e).__name__, e)))
    return recs


def _program_summary(pdf_path, max_items=14):
    """程序附表（WI-0009 / WI-3852 型）的自检摘要：只列识别到的附表与列。"""
    try:
        recs = extract_program_tables(pdf_path)
    except Exception as e:
        return ["", "程序附表自检失败：%s" % e]
    if not recs:
        return []
    groups = {}
    for r in recs:
        k = (r["sub"], r["equip_raw"], r["prog_no_raw"], r["dwg"])
        g = groups.setdefault(k, {"n": 0, "cols": {}, "params": []})
        g["n"] += 1
        g["cols"][(r["program"], r["solution"])] = True
        if r["param"] not in g["params"]:
            g["params"].append(r["param"])
    L = ["", "识别到 %d 张「设备参数附表」（程序 A/B 双列转置表）、共 %d 个取值点。"
         % (len(groups), len(recs))]
    for (sub, equip, pno, dwg), g in groups.items():
        cols = "｜".join("程序%s%s" % (a or "?", b) for (a, b) in g["cols"])
        bits = [sub or "未编号"]
        if equip:
            bits.append("设备 %s" % equip)
        if pno:
            bits.append("程序号 %s" % pno)
        if dwg:
            bits.append("/".join(dwg))
        L.append("  %-46s %-28s 取值点 %d" % ("· " + "  ".join(bits), cols, g["n"]))
        L.append("        参数：%s" % "、".join(g["params"][:max_items])
                 + ("…" if len(g["params"]) > max_items else ""))
    return L


def report_layout(pdf_path: str, max_items: int = 14) -> str:
    """版式自检报告：只列出「识别到了什么」，不含判定细节。"""
    L = []
    L.append("【PDF 版式自检】%s" % pdf_path.split("/")[-1].replace("\xa0", " "))
    L.append("=" * 62)
    L.append("pdfminer 兼容补丁：%s" % ("已启用（修复 PSKeyword.decode 缺失导致的整页解析失败）"
                                    if PATCHED_PDFMINER else "未启用（当前 pdfminer 无需补丁）"))
    with open_pdf(pdf_path) as pdf:
        L.append("总页数：%d" % len(pdf.pages))
        found_pages = []
        for pno, page in enumerate(pdf.pages, 1):
            n_tab = len(page.find_tables() or [])
            marks = []
            if "参数名称" in (page.extract_text() or ""):
                marks.append("竖表(参数名称)")
            for t in page.find_tables() or []:
                for row in t.extract() or []:
                    if _step_head_cols(row):
                        marks.append("横表(参数/Step)")
                        break
            if n_tab or marks:
                found_pages.append(pno)
                L.append("  p%-3d 表格数 %-2d %s" % (pno, n_tab, " | ".join(marks) or "—"))
    reset_layout_stats()               # 版式自检要的是「这一份 PDF」的命中情况
    recs = extract_stepwise(pdf_path)
    L.append("")
    L.extend(layout_registry_report())
    L.append("")
    bad = mojibake_report(pdf_path)
    if bad:
        L.append("字体映射失真（文字/数字可能错位，涉及页的值请人工核对）：")
        for pno, n, s in bad:
            L.append("    p%-3d %d 处乱码片段，如：%s" % (pno, n, s))
        L.append("")
    if LAST_FAILED_PAGES:
        L.append("有 %d 处解析失败（已跳过，其余页正常）：" % len(LAST_FAILED_PAGES))
        for pno, why in LAST_FAILED_PAGES:
            L.append("    p%-3d %s" % (pno, why[:96]))
        L.append("")
    if not recs:
        L.append("✗ 未识别到「参数 / Step」型横向参数表。")
        L.append("  若该 PDF 用的是竖表（表头含「参数名称」），请走竖表抽取路径。")
        L.extend(_program_summary(pdf_path, max_items))
        return "\n".join(L)

    tables = {}
    for r in recs:
        tables.setdefault((r["page"], r["table"], r.get("sub", ""), r["equip_raw"]), []).append(r)
    L.append("识别到 %d 套参数子表、共 %d 个参数取值点。" % (len(tables), len(recs)))
    L.append("")
    for (pno, tno, sub, equip), rs in tables.items():
        steps = sorted(set(r["step"] for r in rs))
        params = []
        for r in rs:
            if r["param"] not in params:
                params.append(r["param"])
        dia = next((r.get("dia") for r in rs if r.get("dia")), None)
        head = "  p%-3d 表%-2d %-8s 设备：%-24s" % (pno, tno, sub or "—", equip or "（未标）")
        L.append(head)
        extra = []
        extra.append("直径 %s" % ("%.1f-%.1fmm" % dia if dia and dia[0] != dia[1]
                                  else ("%.1fmm" % dia[0] if dia else "未标注")))
        prog = next((r.get("prog") for r in rs if r.get("prog")), "")
        if prog:
            extra.append("程序号 %s" % prog)
        sleeve = next((r.get("sleeve") for r in rs if r.get("sleeve")), "")
        if sleeve:
            extra.append("保护套 %s" % sleeve)
        L.append("        %s" % "  |  ".join(extra))
        L.append("        Step：%s" % "、".join("Step %d" % s for s in steps))
        L.append("        参数：%s" % "、".join(params))
    allp = []
    for r in recs:
        if r["param"] not in allp:
            allp.append(r["param"])
    L.append("")
    L.append("全部参数名（去重 %d 个）：%s" % (len(allp), "、".join(allp[:max_items])
                                        + ("…" if len(allp) > max_items else "")))
    L.extend(_program_summary(pdf_path, max_items))
    return "\n".join(L)




# --------------------------------------------------- 字体映射失真（乱码）检测

# 缺 ToUnicode 映射的 PDF，pdfminer 会把文字映射到希腊字母/拉丁扩展区，
# 出现「ƌƚŝŵĞƐ」（其实是 产品线A）、「ϭϵ͘ϱ」（其实是 19.5）这类串。
# 中文正常、数字却错位 —— 最危险的情况：看着有值，值其实是错的。
_MOJI_RUN = re.compile(r"[\u0370-\u03ff\u0100-\u024f\u1e00-\u1eff\u0180-\u024f]{3,}")


def mojibake_report(pdf_path):
    """按页检测「字体映射失真」，返回 [(页码, 乱码片段数, 样例)]。

    只做检测、不做自动还原：这类串的映射不是固定偏移，硬猜会造出错误的数值，
    对核对工具来说「给出错值」比「说不出值」更糟。检测出来就如实提示人工核对。
    """
    out = []
    with open_pdf(pdf_path) as pdf:
        for pno, page in enumerate(pdf.pages, 1):
            try:
                text = page.extract_text() or ""
            except Exception:
                continue
            runs = _MOJI_RUN.findall(text)
            if len(runs) >= 3:
                out.append((pno, len(runs), "、".join(runs[:3])))
    return out


def text_quality_warning(pdf_path):
    """给界面/报告用的一句话提示；没问题时返回空串。"""
    try:
        bad = mojibake_report(pdf_path)
    except Exception:
        return ""
    if not bad:
        return ""
    parts = ["第 %d 页(%d 处，如 %s)" % (p, n, s) for p, n, s in bad[:4]]
    return ("该 PDF 有 %d 页文字疑似「字体映射失真」：%s。"
            "这类页的数字也可能是错位的（如 ϭϵ͘ϱ 实为 19.5），"
            "本工具不做自动还原，涉及这些页的数值请人工核对原文。"
            % (len(bad), "；".join(parts)))

# ==========================================================================
# ==========================================================================
#  第 2.6 节  类型填写规范 + 缺失场景判定（声明式规则表）
#
#  为什么单独立一节：这两套规矩以前是**散落的 if** —— 类型规矩只在第 3 节自检里
#  有几段，判定层（第 7 节）又另写一套；两边各写一遍必然漂移：自检说「该填上下限」、
#  判定却说「空着不算错」（或反过来），界面上两句话互相打架。
#  现在**判定层与界面共用同一张表**（compare 算出来的 want_show / want_over 就是
#  查这张表得来的），改规矩只改一处。
#  ⚠ 待办：第 3 节「单表自检」还是纯 Excel 侧的老规则，没接这张表。实测语料里
#    有 32 行「类型与两个是/否开关对不上」（类型=固定值却显示上下限=是 31 行、
#    类型=浮点却显示上下限=否 1 行），界面上现在看得到，自检还没报 —— 等用户
#    定了口径再接（接了会改变自检问题条数，不能悄悄改）。
#
#  用户 2026-09-29 定的类型规矩：
#    固定值 → 数值填在「固定值」列；显示上下限=否；允许超上下限=否
#    字符串 → 所有值列都不填；显示上下限=否；允许超上下限=否
#    浮点   → 显示上下限=是，再按 **PDF 那一格的值形态** 分三类：
#              ① PDF 是范围（8~10）      → 上下限**都**要填；允许超上下限=否
#              ② PDF 是 > / < / ≥ / ≤（≥90）→ 只填对应那一侧；允许超上下限=否
#              ③ PDF 那一格**带 ref 标记** → 显示上下限=是、允许超上下限=**是**，
#                                            值一列都不填（上下限也不用填）
#
#  ③ 的判定看的是 **ref 这个字眼本身**，不是「值是不是单一值」（2026-09-29 更正）：
#    `58(ref)` 是 ref，`8~10(ref)` 同样是 ref —— 一律两个都填「是」、值全都不填。
#    反过来说，PDF 只给一个普普通通的数（没标 ref）**不等于** ref，
#    那种按最保守处理（见 FLOAT_SHAPE_RULES["single"] 的注释）。
#
#  本节的职责边界：
#    · `TYPE_RULES` / `FLOAT_SHAPE_RULES`  只描述「该怎么填」，不含任何取值逻辑
#    · `pdf_value_shape` / `_has_ref`      只判断「PDF 这一格是什么形态」
#    · `type_fill_rule`                    只做查表
#    · `decide_blank_cell`                 只回答「Excel 空着对不对」
#  四件事分开，改哪一条都不会牵动别的。
# ==========================================================================

# 类型 → 值该填在哪一列 / 显示上下限 / 允许超上下限 / 是否必须填值。
# `must_fill` / `allow_over` 写 None = 「由 PDF 值形态决定」（只有浮点类这样）。
TYPE_RULES = {
    "固定值": {"cols": ("fixed",), "show_lim": False, "allow_over": False,
               "must_fill": True},
    "字符串": {"cols": (), "show_lim": False, "allow_over": False,
               "must_fill": False},
    "布尔值": {"cols": ("btrue", "bfalse"), "show_lim": False, "allow_over": False,
               "must_fill": True},
    "浮点": {"cols": ("lo", "hi"), "show_lim": True, "allow_over": None,
             "must_fill": None},
    "整数": {"cols": ("lo", "hi"), "show_lim": True, "allow_over": None,
             "must_fill": None},
    "十进制": {"cols": ("lo", "hi"), "show_lim": True, "allow_over": None,
               "must_fill": None},
}

# PDF 值形态 → 浮点类该怎么填
FLOAT_SHAPE_RULES = {
    "range": {"must_fill": True, "allow_over": False,
              "why": "PDF 给的是区间（两侧都有限）→ 上下限都要填"},
    "one_sided": {"must_fill": True, "allow_over": False,
                  "why": "PDF 给的是 > / < / ≥ / ≤（单边约束）→ 只填对应那一侧"},
    # ⚠ 这条规矩是**双向**的，两处都要实现，只做一半就会漏判：
    #   空着 → 对（decide_blank_cell 的 type_exempt）
    #   填了 → 错（compare 里的 ref_filled；见 WI-0020 表4!164 的教训：
    #          填了下限 115、PDF 是 `115(ref)`，只做前一半时判成了「一致」）
    "ref": {"must_fill": False, "allow_over": True,
            "why": "PDF 这一格标了 (ref)（不管是一个数还是 8~10 这种区间）"
                   " → 特殊浮点：上下限都不填，其余空位也不用填；填了反而是错"},
    # PDF 只给一个普普通通的数、又没标 ref：不在用户定的三类里，规矩没覆盖。
    # 这里按**最保守**处理（不必填、也不允许超）—— 用户没说这种情况该报漏填，
    # 宁可放过也别凭空造出一批「不一致」。
    "single": {"must_fill": False, "allow_over": False,
               "why": "PDF 只给一个值、且没标 (ref) → 规矩未覆盖，按不必填处理"},
    "none": {"must_fill": False, "allow_over": False, "why": "PDF 这一格没有值"},
}

# 形态从宽到严：多张候选表形态不一致时取最严的那个（宁可让你多看一眼，
# 也不把该填的轻轻放过）。ref 例外：它走 dominant_shape 里的短路，不看这里的序。
_SHAPE_RANK = {"range": 3, "one_sided": 2, "ref": 1, "single": 1, "none": 0}

# `>90℃` 这类原文里的比较符。必须**先看原文再看数值**：取值解析器把 `>90℃`
# 读成 (90, 90)，光看上下限会被误判成「单一值」—— 那该填的就不用填了。
#
# ⚠️ 字符集由 `SYMBOL_SPECS` 生成（R47），不再手抄 —— 手抄只写了 `<>≥≤＞＜`，
# Symbol 字体画出来的比较符（`\uf0b3` 之类）就认不出来，`≤90` 会被当普通单值。
# 中英文别名（大于 / 以上 / 不低于 …）也在表里，一并收进来。
_OP_RE = re.compile(
    _symbol_class_re("gt", "lt", "ge", "le") + r"|"
    + "|".join(re.escape(w) for w in sorted(SYMBOL_WORDS, key=len, reverse=True)))

# ref 标记。三种写法都要认，且**不能误伤**含 ref 的其它词：
#   `0.2MPa（ref）` / `0.0215" (ref)`  ← PDF 侧（全角、半角括号都有实测）
#   `压力/4(Bar):ref`                  ← Excel 侧（解析后落在 unit 上，见下）
#   `58 ref`                           ← 独立成词
# 所以要求 ref 前后有括号 / 冒号 / 空白 / 串首尾 —— "Reference"、"preform"
# 这类不会被认成 ref。
_REF_RE = re.compile(r"[\(（]\s*ref\s*[\)）]|[:：]\s*ref\s*$|"
                     r"(?:^|\s)ref(?:\s|$|[）\)])|^\s*ref\s*$", re.I)


def _has_ref(text):
    """文本里有没有 ref 标记。（PDF 原文 / Excel 采集项都走这一个入口。）"""
    return bool(_REF_RE.search(str(text or "")))


def pdf_value_shape(rec):
    """PDF 这一格的值形态 → ref / one_sided / range / single / none。"""
    rec = rec or {}
    # 「二次拉伸」这一步继承来的 ref：PDF 在这里漏标了，但右侧 Purpose 写明了它是
    # 上一步工艺的延续（见 `_enstep_table` 的 `_inherit_ref`）—— 比数值推断可靠，
    # 与显式 ref 同等优先。
    if rec.get("ref_inherit"):
        return "ref"
    # 引用式（`Step5+5` / `Same as step5`）：本格没写 ref，但**被引用的那一格**
    # 写了 —— 按用户 2026-09-30 的口径，本格同样按 ref 算（见 `_enstep_ref_via`）。
    # 放在「本格显式 ref」之前：两者结论一致，放这里是让依据里能说清来源。
    if rec.get("ref_via"):
        return "ref"
    # ref 次之：它是工艺文件里**显式写的标记**，比从数值推断出来的形态可靠。
    # `8~10(ref)` 光看上下限是区间，但按规矩它就是 ref（两个都填「是」、值不填）。
    if _has_ref(rec.get("raw")) or _has_ref(rec.get("unit")):
        return "ref"
    if _OP_RE.search(str(rec.get("raw") or "")):
        return "one_sided"
    lo, hi = rec.get("lo"), rec.get("hi")
    if lo is None and hi is None:
        return "none"
    if lo is None or hi is None:
        return "one_sided"
    try:
        if abs(float(lo) - float(hi)) > 1e-9:
            return "range"
    except (TypeError, ValueError):
        return "single"
    return "single"


def _pdf_value_shape_noref(rec):
    """忽略 ref 标记、只按**数值形态**判断这一格（非球囊 WI 用）。

    球囊 WI 里 PDF 标了 (ref) 的格子 = 不填（走 pdf_value_shape 的 ref 分支）；
    非球囊 WI 里 ref 不享受特殊待遇，照常填写、照常数值比对 —— 所以要看
    「去掉 ref 字眼后，这个数值本身是什么形态」（区间 / 单边 / 单值 / 空）。
    """
    rec = rec or {}
    if _OP_RE.search(str(rec.get("raw") or "")):
        return "one_sided"
    lo, hi = rec.get("lo"), rec.get("hi")
    if lo is None and hi is None:
        return "none"
    if lo is None or hi is None:
        return "one_sided"
    try:
        if abs(float(lo) - float(hi)) > 1e-9:
            return "range"
    except (TypeError, ValueError):
        return "single"
    return "single"


def _demote_ref_shape(er, step_index):
    """非球囊 WI 时，把「这一格是 ref」降级成数值形态。

    只处理 PDF 侧有候选格的场景（沿用 pdf_shapes_for 的收窄链）；PDF 侧
    一条候选格都没有（`not shapes`，那本来就是靠 Excel 的人工 ref 标注凑的）
    时，这里没有数值可看，返回 "none" —— 交给下游按「没有值」兜底。
    """
    step, base = parse_step_item(er.get("item"))
    if step is None or not step_index:
        return "none"
    ck = _canonical_param(base)
    if not ck:
        return "none"
    recs = _narrow_recs(er, step_index.get((ck, step)) or [])
    shapes = {_pdf_value_shape_noref(r) for r in recs if not r.get("empty")}
    if not shapes:
        return "none"
    return dominant_shape(shapes)


def _equip_keep(er, recs):
    """设备对齐的**唯一判据** —— 全项目只此一份，改这里等于改所有链路。

    留下两类记录：
      · 设备号与 Excel 有交集的（这台机器确实印在这张表上）；
      · `no_equip` 的（这份 PDF 的表里**根本没印设备号**）。
    后者必须留：那是「PDF 没写」，不是「PDF 没有这台设备」。按设备收窄时把它
    剔掉，等于把「没写」当成「没有」，凭空造一批 no_equip_match。

    ⚠️ 为什么必须单独抽出来：2026-09-30 重构前，这条判据在 `_step_expected`
    （严格链）和本函数（形态链）里各写了一遍。两份一旦漂移，同一行会因为它
    走哪条链路而得到不同答案 —— 那正是「同一个问题换个文件又冒出来」的根源。
    """
    eq = set(_equip_set(er.get("equip")))
    return [r for r in recs
            if (set(r.get("equip") or ()) & eq) or r.get("no_equip")]


def _narrow_recs_by_equip(er, recs):
    """设备对齐 —— **保守版**（形态链用）：收不动就原样返回。

    ⚠️ 规格相同 ≠ 同一套参数。WI-0036 的表2、表3 直径**都是 Ø6.0**，但
    表2 只服务 T-1177、表3 服务其余 6 台 —— 两张表的写法并不一致：
      · 表2 Step9 左/右拉伸距离写 `SL1+30 / ref`（ref 形态，不填值）
      · 表3 Step9 同一格写 `5-180`（普通区间，要填 5~180）
    按直径收窄分不开，两种形态一并集 → `dominant_shape`「有 ref 就按 ref」
    → card 0015（那 6 台）被判「允许超上下限应为『是』」，而它填的 5~180 /
    「否」恰恰是对的。用户 2026-09-30 点名：「原文档中好像没有写它是哪里是
    ref…… 这个就是完全误判」—— PDF 里确实没有，是隔壁设备的 ref 串过来了。
    """
    eq = set(_equip_set(er.get("equip")))
    if not eq or not recs:
        return recs
    if not any(r.get("equip") for r in recs):
        return recs                      # PDF 侧一张表都没印设备号 → 无从收窄
    return _equip_keep(er, recs) or recs


def _equip_gate(er, recs):
    """设备对齐 —— **严格版**（严格链用）：收不动 = 这台设备在 PDF 里查无此机。

    返回 `None` 表示「锁不上」，由调用方判 `no_equip_match`；否则返回收窄后的
    候选。与 `_narrow_recs_by_equip` 共用 `_equip_keep` 这一份判据，
    差别**只在收不动时怎么办**：保守版退回原候选，严格版直接定案。

    2026-09-30 已统一：补上与保守版同款守卫「PDF 侧一张表都没印设备号 → 无从
    收窄，原样返回」。实测全语料 21 份 62200 行，影响 **0 行** —— 因为凡是
    没印设备号的表，记录都带着 `no_equip=True`，`_equip_keep` 本来就会把
    它们留下，两版殊途同归。统一只是把「纸面上该一致」的两份代码对齐，
    杜绝将来有人改其中一份、导致又冒出行为分叉。
    """
    if not _equip_set(er.get("equip")):
        return recs
    if not any(r.get("equip") for r in recs):
        return recs                      # PDF 侧一张表都没印设备号 → 无从收窄
    return _equip_keep(er, recs) or None


def _narrow_recs(er, recs):
    """形态/原文收窄的**统一入口**：先按设备，再按规格。

    两级都要：只按规格，同直径的多张表分不开（见 `_narrow_recs_by_equip`）；
    只按设备，同一台设备的多个规格档又会互相串（见 `_narrow_recs_by_spec`）。
    与 `_step_expected` 的收窄链同构 —— 那里是「设备对齐 → 规格级 → 规格」，
    这里是这两级的保守版（任一级收不窄就退回，绝不因为收窄而丢候选）。
    """
    return _narrow_recs_by_spec(er, _narrow_recs_by_equip(er, recs))


# 标题里「直径 × 球囊长度」的一支：`Æ2.0~4.0mm x 20-150mm` / `Æ1.5mm x 15-120mm`。
# 一支标题可以有好几支（用 ；/，/、 分隔），**每支各带各的直径**。
_QLEN_PAIR_RE = re.compile(
    r"(\d+(?:\.\d+)?)\s*(?:[~～\-－到至]\s*(\d+(?:\.\d+)?))?\s*mm"
    r"(?:\s*[x×*X]\s*(\d+(?:\.\d+)?)"
    r"\s*(?:[~～\-－到至]\s*(\d+(?:\.\d+)?))?(?:\s*mm)?)?",
    re.I)


def _title_dia_len_pairs(title):
    """子表标题 → [(直径区间, 长度区间|None), ...]（**一支一支拆开**）。

    ⚠️ 长度段**必须跟着它那一支的直径**，绝不能从整个标题里捞区间 —— 这是
    第 53 轮修掉的病根。WI-0038 表3 的标题是
        「直径为 Æ2.0~4.0mm x 20-150mm；**Æ2.0mm** x 200-300mm…」
    那个 200-300 只管 **Æ2.0 单一直径**的球囊。旧写法 `_spec_ranges(title)`
    把 {(2.0,4.0), (20,150), (200,300)} 一股脑当成本表长度，于是规格
    `φ(2.5-4.0)xQL(200-300)mm` 的卡（card 0065）误以为表3 也管 200-300，
    与表5 并列 → 两张表的写法一并集 → 被判成 ref。
    拆开成对之后：表3 那一支的直径 (2.0,2.0) 压根盖不住 φ2.5-4.0 → 不算数。
    """
    out = []
    for a, b, c, d in _QLEN_PAIR_RE.findall(str(title or "")):
        lo = float(a)
        hi = float(b) if b else lo
        if hi < lo:
            lo, hi = hi, lo
        if c:
            l2 = float(c)
            h2 = float(d) if d else l2
            if h2 < l2:
                l2, h2 = h2, l2
            out.append(((lo, hi), (l2, h2)))
        else:
            out.append(((lo, hi), None))
    return out


def _qlen_top(recs, top, er):
    """直径分不出胜负时，用标题里的**球囊长度**再收窄一次（第 53 轮）。

    为什么需要第二级：球囊 WI 的参数表常常按「直径 **×** 球囊长度」分档，
    而 `parse_sub_title` 只取了直径那一段。WI-0038 实测：
        表3  「直径为 Æ2.0~4.0mm x 20-150mm；Æ2.0mm x 200-300mm…」
        表5  「Æ2.5~4.0mm x 200-300mm…」
    而 Excel 的规格是 `φ(2.5-4.0)xQL(20-150)mm` 与 `φ(2.5-4.0)xQL(200-300)mm`
    —— 直径段几乎一样，只有长度段能把 `QL(20-150)` 和 `QL(200-300)` 分开。
    分不开的代价：两张表的写法一并集，`dominant_shape` 又「有 ref 就按 ref」，
    于是本该按普通浮点填值的卡被判成 ref（不该填值）。

    ⚠️ 长度段必须**跟着它那一支的直径**（见 `_title_dia_len_pairs`）：
    一支标题里写好几档时，各档的直径是不同的，混着比就是把别档的长度
    算到本档头上。

    与 `_dia_score` 的分工：它是**第一级**、这个函数是并列时的**第二级**。
    只在「两边都写得有长度段、且至少有一张对得上」时才收窄 ——
    任一边没有长度信息就返回空集（调用方保持原样），绝不因为收窄而丢候选。

    ⚠️ 这是**唯一**的长度收窄实现：类型路径（`_narrow_recs_by_spec`）与取值
    路径（`_step_expected`）共用它。第 53 轮就吃过亏 —— 只在类型路径里加了
    长度收窄，取值路径自己另有一份「直径并列就全留」，于是 card 0089 的
    「加热中间模/5」类型判对了（不再说它是 ref），期望值却还是 表4(95~105)
    与 表6(85~95) 的并集 —— 两条路径给出两个互相打架的答案。
    """
    return _qlen_keys(recs, top, *_qlen_need(er.get("spec")))


def _qlen_need(spec_text):
    """Excel「适用规格」→ (直径区间, 其余区间集合)。后者通常就是球囊长度段。

    拆开的理由：直径段要拿去跟子表的直径档比，长度段要拿去跟标题里的「× 长度」
    比，两者是**两个不同的维度**，混在一起就会拿长度去比直径（见 `_spec_range`
    里 WI-0020 三张卡全军覆没那次）。
    """
    spec = _spec_range(spec_text)
    need = {x for x in _spec_ranges(spec_text)
            if spec is None or x != tuple(spec)}
    return spec, need


def _qlen_keys(recs, top, spec, need):
    """直径并列时用标题里的**球囊长度**再分一级 → 收窄后的 key 集合。

    收不窄（任一边没有长度信息）返回**空集**，由调用方决定怎么办 ——
    类型路径是「保持原样」，取值路径是「保持原样」。绝不因为收窄而丢候选。
    """
    if not spec or not need:
        return set()
    by_key = {}
    for r in recs or ():
        if not r.get("dia") or not r.get("title"):
            continue
        k = (r.get("sub") or "", tuple(r["dia"]))
        if k not in top:
            continue
        # 只收**直径盖得住本行规格**那一支的长度段
        have = {ln for pdia, ln in _title_dia_len_pairs(r.get("title"))
                if ln and _dia_score(spec, pdia) >= 1}
        by_key.setdefault(k, set()).update(have)
    if not any(by_key.values()):
        return set()
    scored = {k: len(need & v) for k, v in by_key.items()}
    best = max(scored.values())
    if best <= 0:
        return set()
    return {k for k, s in scored.items() if s == best}


def _narrow_recs_by_spec(er, recs):
    """把候选格按 Excel「适用规格」收窄到**同档的那张子表**。

    为什么必须有这一层：`step_index[(规范名, 步骤号)]` 装的是**跨全部规格**的候选。
    取「值形态」时不收窄，隔壁档的写法就会串过来 —— 实测 WI-0032 表4!156
    「退出条件时间/12」：本行适用规格 φ3.0-6.0（对应表2，PDF 写 `3-15s`，是**区间**），
    可表3（φ7.0-10.0）同一步写的是 `3-15s(ref)`。三张表的形态一并集就成了
    {range, ref}，`dominant_shape` 又「有 ref 就按 ref」，于是本行被判成
    「两个开关都该填『是』」→ 报 Excel 把「允许超」填反了。**这是误报**：
    收窄到表2 之后只剩 range，应为「是 + 否」，与 Excel 填的完全一致。

    与 `_step_expected` 的规格级收窄**同一口径**（`_dia_score` 取最高分、并列全留），
    但**更保守**：收不窄（没写规格 / 候选都没标直径 / 最高分为 0）就**原样返回**，
    绝不因为收窄失败而丢掉候选 —— 少收一级只是退回旧行为，不会凭空造结论。
    """
    spec = _spec_range(er.get("spec"))
    if not spec or not recs:
        return recs
    keys = {(r.get("sub") or "", tuple(r["dia"])) for r in recs if r.get("dia")}
    if not keys:
        return recs
    scored = {k: _dia_score(spec, k[1]) for k in keys}
    best = max(scored.values())
    if best < 1:
        return recs
    top = {k for k, s in scored.items() if s == best}
    # 直径并列（如 Æ2.0~4.0 与 Æ2.5~4.0 都被 φ2.5-4.0 压住）→ 用球囊长度再分
    if len(top) > 1:
        top = _qlen_top(recs, top, er) or top
    keep = [r for r in recs
            if not r.get("dia") or (r.get("sub") or "", tuple(r["dia"])) in top]
    return keep or recs


def _pdf_direct_shapes(recs):
    """PDF 候选格里**直接写出来的**值形态 —— 引用式（`Step N` / `SamestepN`）不算。

    为什么要单拎出来：引用式那一格（WI-0032 的 `Step5+5`）本身**不表态**，
    它只是把值指到别处去。`_enstep_resolve` 会把值解出来（`Step5+5` → 110、110），
    可解析器只留下了 lo/hi ——「被引用的那一格是不是 ref」这层信息在半路丢了。
    于是 `pdf_value_shape` 看到 lo==hi 就报 `single`，拿它当「PDF 给了字面值」的
    证据会得出错误结论（详见 `pdf_shapes_for` 里 加热中间模/4 的实测）。
    """
    out = set()
    for r in recs or ():
        if r.get("empty"):
            continue
        if _ENSTEP_REF_RE.match(str(r.get("raw") or "").strip()):
            continue
        out.add(pdf_value_shape(r))
    return out


def pdf_shapes_for(er, step_index):
    """这一行在 PDF 里全部候选格的形态集合（**已按适用规格收窄**）。

    ref 有两个来源，但**分量不一样**：
      · PDF 那一格直接写 `58(ref)` —— 这是判据本身，最高优先；
      · Excel 采集项名写 `压力/4:ref`（解析后 ref 落在 unit 上）—— 这只是**人工
        标注**，用户 2026-09-29（第 46 轮）实测它并不总是对。

    ⚠️ 人工标注只在 PDF **没给出直接字面值**时才采纳（第 46 轮改）。
    实测 WI-0032 表4!185/186「左/右拉伸距离/7」：PDF 那一格写的是 `1-30`
    （干净区间、一个 ref 字眼都没有），Excel 名字里却带着 `:ref`。旧口径
    「Excel 标了就算 ref」→ dominant 短路成 ref → 判「两个开关都该是」，
    于是**这行看着「一致」，真正该改的（补 1~30、允许超改否）被整条放过**。
    用户原话：「可能你看到左拉伸距离后面有个 ref……但是这个不一定对，
    **以表格为基础参考**。表格当中没有 ref 那就应该在上下限里面填上 1 到 30。」

    ⚠️⚠️ **第 53 轮收紧**：以前引用式（`Step5+5`）那一格被 `_pdf_direct_shapes`
    排除后 `direct` 为空，于是整行退化成「沿用 Excel 的人工 ref 标注」——
    而引用式的性质**根本查得出来**：顺着引用链看被引用格（见 `_enstep_ref_via`）。
    旧口径下同一个 `Step5+5`，只要 Excel 单位列写了 `ref`，不论被引用格有没有
    ref 一律判 ref，于是把「该填值」的行判成「多填」。
    用户 2026-09-30 点名并定了口径：
      「**WI-0020 的 step5 内有 ref，所以 step5+5 也属于是 ref**，要空着；
       **WI-0038 内 step5 没有 ref 字眼，所以 step5+5 就是正常浮点是否，要填写**」
    所以：引用式的形态由 `ref_via` 给出，**人工标注不再顶替它**。

    人工标注现在的射程只剩一条：**PDF 侧一条候选格都找不到**时（`not shapes`），
    才退回人工标注 —— 那种情况下确实没有别的东西可依据。

    ⚠️ 收窄这一步不能省：不收窄会把隔壁规格档的写法当成本行的形态（见
    `_narrow_recs_by_spec` 里 WI-0032 表4!156 的实测）。
    """
    shapes = set()
    step, base = parse_step_item(er.get("item"))
    if step is not None and step_index:
        ck = _canonical_param(base)
        if ck:
            recs = _narrow_recs(er, step_index.get((ck, step)) or [])
            shapes = {pdf_value_shape(r) for r in recs if not r.get("empty")}
    if not shapes and (_has_ref(er.get("unit")) or _has_ref(er.get("item"))):
        shapes.add("ref")
    return shapes


def dominant_shape(shapes):
    """多个候选表形态不一致时取最严的；**有 ref 就按 ref 算**。

    ref 是显式标记，比从数值推断出来的形态可靠：一张表写 `3-15s(ref)`、
    另一张写 `3-15s` 时，按 ref 处理（不填），而不是按区间报漏填。
    """
    shapes = set(shapes or ())
    if not shapes:
        return "none"
    if "ref" in shapes:
        return "ref"
    return max(shapes, key=lambda s: _SHAPE_RANK.get(s, 0))


# --------------------------------------------------------------------------
#  球囊类 WI 的**特别规矩**（用户 2026-09-29 新增，只在球囊 WI 上生效）
#
#   PDF 那一格出现 `>` / `<`（≥ / ≤ 同）→ Excel 的「类型」**不是浮点**，是**字符串**，
#   「显示上下限」「允许超上下限」两个都填「否」。
#
#  为什么必须单独一条：通用规矩里 PDF 给 `>90℃` 属于**单边约束**，按浮点处理、
#  要填满对应一侧的限值；球囊这条产品线偏偏不这么填 —— 它把整格当字符串
#  （值列全空、两个开关全否）。同一份 PDF 形态在两条产品线里对应两种填法，
#  只能按产品线分叉。用户明确选了「只在球囊的 WI 生效」，别外扩到别的线。
#
#  ⚠ 认的是**原文里那个比较符本身**，不是「单边值」这种推断出来的形态：
#    `pdf_value_shape` 里 lo/hi 只给了一侧也会得到 `one_sided`，那种没有比较符，
#    不属于这条规矩的射程。所以下面走 `pdf_op_raw` 单独查原文，不复用 shape。
#
#  边界：只在 Excel 类型本来是浮点类（浮点 / 整数 / 十进制）时才要求改成字符串。
#    类型本来就是字符串的 → 通用规矩（TYPE_RULES["字符串"]）已经要求两个都「否」；
#    固定值 / 布尔值 → 不在这条规矩射程内，保持原判，绝不顺手扩大打击面。
# --------------------------------------------------------------------------
_BALLOON_RE = re.compile("球囊")

# 会被这条规矩要求改成「字符串」的类型
_FLOATISH = ("浮点", "整数", "十进制")


def doc_is_balloon(*texts):
    """这些文本里有没有「球囊」→ 这份 WI 是不是球囊类。"""
    for t in texts:
        if t and _BALLOON_RE.search(str(t)):
            return True
    return False


def detect_balloon(pdf_steps=None, excel_records=None, pdf_path=None):
    """这次比对是不是球囊类 WI（球囊专属规矩的总开关）。

    判定按「整份文档是不是球囊类」来（用户 2026-10-14 定）：只要文件名 / 正文 /
    子表标题 / Excel 工艺名里任何一处出现「球囊」，就算球囊类。

    为什么不能只认子表标题：像 WI-0006 / WI-0023 这种「球囊**远端焊接**」类 WI，
    正文和文件名都写着球囊，但子表标题是「…激光焊接参数表」—— 光看标题会漏判，
    ref「不填」这条规矩就对它们失效（实测 WI-0006 的移动距离被误判成多填）。
    所以优先取**最宽的口径**（文件名 + 正文），子表标题 / 工艺名只作兜底。
    """
    if pdf_path:
        # 文件名 + 前 3 页正文：最可靠（文件名可能被改，正文不会）
        if doc_is_balloon(os.path.basename(pdf_path)):
            return True
        try:
            with open_pdf(pdf_path) as _p:
                for _i in range(min(3, len(_p.pages))):
                    if doc_is_balloon(_p.pages[_i].extract_text() or ""):
                        return True
        except Exception:
            pass
    for r in (pdf_steps or ()):
        if doc_is_balloon(r.get("title")):
            return True
    if not pdf_steps:
        for r in (excel_records or ()):
            if doc_is_balloon(r.get("proc")):
                return True
    return False


def pdf_op_raw(er, step_index):
    """这一行在 PDF 里那一格的**原文**有没有 `>` / `<`（≥ ≤ 同）。

    返回匹配到的原文（如 `>90℃`）；没有则返回 ""。
    单独一个函数是因为：`pdf_value_shape` 把「只给了一侧值」也算 one_sided，
    而那条规矩认的是**比较符本身** —— 不能拿推断出来的形态顶替。

    ⚠️ 同样要**先按适用规格收窄**：不收窄会把隔壁档的原文当成本行的原文
    （实测 WI-0032「退出条件加热中间模/6(℃)」本行该读表2 的 `>110℃(ref)`，
    旧口径读出来的是表1 的 `>115℃(ref)`）。
    """
    step, base = parse_step_item(er.get("item"))
    if step is None or not step_index:
        return ""
    ck = _canonical_param(base)
    if not ck:
        return ""
    for r in _narrow_recs(er, step_index.get((ck, step)) or []):
        raw = str(r.get("raw") or "")
        if _OP_RE.search(raw):
            return raw
    return ""


def expected_type_for(er, op_raw, balloon=False):
    """这一行 Excel 的「类型」列**该**填什么。

    规矩没在类型层面表态时返回 None（→ 只判两个开关，沿用现在的类型）。
    目前只有球囊那条会在类型层面表态。
    """
    typ = str(er.get("type") or "").strip()
    if balloon and op_raw and typ in _FLOATISH:
        return "字符串"
    return None


def type_fill_rule(typ, shape=None):
    """(类型, PDF 值形态) → {must_fill, allow_over, show_lim, why}。

    认不得的类型返回全 None —— 调用方必须自己决定怎么兜底，别在这里替它猜。
    """
    r = TYPE_RULES.get(str(typ or "").strip())
    if not r:
        return {"must_fill": None, "allow_over": None, "show_lim": None, "why": ""}
    out = {"must_fill": r["must_fill"], "allow_over": r["allow_over"],
           "show_lim": r["show_lim"], "why": ""}
    if r["must_fill"] is None or r["allow_over"] is None:
        sub = FLOAT_SHAPE_RULES.get(shape or "none", FLOAT_SHAPE_RULES["none"])
        out["must_fill"] = sub["must_fill"]
        out["allow_over"] = sub["allow_over"]
        out["why"] = sub["why"]
    return out


def _fmt_vals(vals, ref=False):
    """把候选值排成给人看的一句话；ref=True 时每个值后面带 (ref)。

    带 ref 的值如果不标出来，「PDF 有值 10、Excel 空着」看着就像漏填 ——
    而按类型规矩它恰恰**不该填**。标了 (ref) 才不会让人以为判错了。
    """
    def _vkey(v):
        return tuple((1, 0.0) if x is None else (0, float(x)) for x in v)

    # ⚠️ 两道防线，缺一不可：
    #   ① 先**滤掉**没有数字的元组 —— 读不出数值的格子（如 `SL1+30 ref`）
    #      它不该在「应为」里印成 `None` 让人以为是 0 或者漏项；
    #   ② 排序键也不能裸写 float —— 万一上游漏了过滤，Python 会在比较
    #      None 与 float 时抛 TypeError，整个比对当场崩掉（2026-09-30 实测）。
    #      期望值只有这一个出口，这里必须自己扛住。
    vals = [v for v in (vals or ()) if v[0] is not None or v[1] is not None]
    out = []
    for v in sorted(vals, key=_vkey):
        lo, hi = v[0], v[1]
        s = (_fmt_num(lo) if (hi is None or abs(float(lo) - float(hi)) < 1e-9)
             else "%s~%s" % (_fmt_num(lo), _fmt_num(hi)))
        out.append(s + "(ref)" if ref else s)
    return "、".join(out)


# ==========================================================================
#  第 2.7 节  缺失场景判定表（PDF 侧状态 × Excel 侧状态 → 判定）
#
#  用户 2026-09-30 定的口径（原话）：
#    「PDF内有值，excel没有那就是漏填，
#      PDF没这一步或者填写NA或者没值，excel则可以是空或者NA」
#
#  这条规矩**两侧都要判**：PDF 侧是「这一格到底是什么情况」，Excel 侧是
#  「这一格填了什么」。以前这两件事写成两条**互相抄**的 if/elif 链 ——
#  `decide_blank_cell` 只管 Excel 空，`compare` 里 `na_declared` 那段只管
#  Excel 填 N/A。同一句规矩改一遍漏一遍，实测「PDF 有值但按类型规矩不必填」
#  在 Excel 空时判「不填正确」、在 Excel 填 N/A 时却判「不该填 N/A（漏填）」，
#  一条规矩两个结论。
#
#  本节把两条链合成**一张表**：
#     · 加规矩 / 改判定 → 改 `MISSING_RULES`（一张 8×2 的表）
#     · 新增 PDF 侧状态 → 改 `pdf_cell_state`（顺序即优先级）
#     · 改措辞         → 改 `decide_missing_cell` 里的模板
#  三件事互不牵动，不会再出现「改了这一半忘了那一半」。
#
#  ⚠️ 表只覆盖 Excel 侧**没值**的两种形态（blank / na）。
#     Excel 填了值的行走数值比对，不进这张表；但「填了不该填的」
#     （PDF 空 / N/A / 没这一步 / 是 ref）在第 7 节按**同一套 PDF 侧状态**判
#     （pdf_blank / pdf_na / no_step / ref_filled 那几支），口径一致。
# ==========================================================================

# PDF 侧「这一格明确写了不适用」。只收**字面就是 N/A** 的写法 ——
# 「无」「不适用」「/」在工艺文件里也可能是别的意思，收进来会误伤真值格。
PDF_NA_RE = re.compile(r"\An\s*\.?\s*/?\s*\.?\s*a\.?\Z", re.I)


def pdf_cell_is_na(text):
    """PDF 这一格是不是**明确写着** N/A（表格在声明「此步不适用」）。

    ⚠️ 以前这个判据在两种版式的解析函数里各抄了一遍（同一条「n 斜杠 a」正则），
    第三种版式（B 型横表）则**根本没抄** —— 于是 PDF 里明明白白写着 N/A
    的格子被当成「解析失败」整格丢掉，比对层永远看不到「这一格不适用」。
    判据只有这一份，新增版式直接调它。
    """
    return bool(PDF_NA_RE.match(str(_c(text) or "").strip()))


def excel_says_na(er):
    """Excel 这一格是不是在**声明**「此处无值」（N/A）。

    与 PDF 侧共用 `pdf_cell_is_na`，不再各写一套。以前这里写的是
    `kind == "fixed" and 值 in ("N/A","NA")` —— 只认固定值列，
    落在别处（或大小写 / 标点不同）的 N/A 就不算，规矩因此长出第二套口径。
    """
    for k in ("fixed", "lo", "hi", "btrue", "bfalse"):
        s = str(er.get(k) or "").strip()
        if s and (s.lower() in ("n/a", "na", "n.a.", "n/a.") or pdf_cell_is_na(s)):
            return True
    return False


def excel_cell_state(er, kind, actual):
    """Excel 这一格的状态 → value / na / blank / unreadable。

    ⚠️ 第四种（`unreadable`）以前被当成 `blank` 混进缺失场景：值列明明写着
    `T-2054-49`（「中间模编号」这类文本编号），却因为读不出数字被当成「没填」。
    那等于把「填了、但我读不懂」说成「没填」—— 万一 PDF 侧恰好有值，
    缺失场景就会判成**漏填**，是凭空造出来的假错。
    """
    if actual is not None:
        return "value"
    if excel_says_na(er):
        return "na"
    if any(str(er.get(k) or "").strip()
           for k in ("lo", "hi", "fixed", "btrue", "bfalse")):
        return "unreadable"
    return "blank"


# PDF 侧状态（穷举）。加状态必须同时补 `MISSING_RULES` 与 `pdf_cell_state`。
PDF_CELL_STATES = (
    "value",      # 该格有值，且按类型规矩**该填**
    "exempt",     # 该格有值，但按类型规矩**不必填**（ref / 单一值 …）
    "ref_only",   # 只标了 (ref)、没给数值（`SL1+30 ref`）
    "na",         # 该格明确写 N/A
    "blank",      # 该格写空（`------`）
    "no_step",    # 该档压根没有这一步
    "no_param",   # 该设备该档的表里**根本没印这个参数**（比 no_step 更靠前，第 53 轮）
    "no_spec",    # 规格套不上任何一档 —— 无从核对
    "unknown",    # 其余（参数表里没这一项 / 设备对不上 …）
)

#                    Excel 空                          Excel 填 N/A
MISSING_RULES = {
    # PDF 有值、且按类型规矩该填 → 不填 / 填 N/A 都是**漏填**
    "value":    {"blank": ("mismatch", "missing"),
                 "na":    ("mismatch", "na_wrong")},
    # PDF 有值、但按类型规矩不必填 → 空着、填 N/A **都对**
    "exempt":   {"blank": ("ok", "type_exempt"), "na": ("ok", "type_exempt")},
    # PDF 只标了 (ref) → 同上；依据要把原文摆出来，不借别的档的数值
    "ref_only": {"blank": ("ok", "ref_cell"),    "na": ("ok", "ref_cell")},
    # 用户说的三种「PDF 侧没有」：写 N/A / 写空 / 没这一步 → 空着、填 N/A 都对
    "na":       {"blank": ("ok", "pdf_na"),      "na": ("ok", "pdf_na")},
    "blank":    {"blank": ("ok", "pdf_blank"),   "na": ("ok", "pdf_blank")},
    "no_step":  {"blank": ("ok", "no_step"),     "na": ("ok", "no_step")},
    # 比 no_step 更靠前：这张表里连这个参数都没印（不是「没有这一步」，
    # 是**整张表没有这一列**）。规格必须**精确套上某一档**才敢下这个结论
    # （见 `_param_gap`）—— 答得出「看哪张表」才能说「那张表里没有」。
    "no_param": {"blank": ("ok", "no_param"),    "na": ("ok", "no_param")},
    # ⚠️ 规格套不上档 = 「我不知道 PDF 有没有这一步」—— 用户列的三种
    #    （没这一步 / 写 NA / 没值）**都不包含它**。判「对」就是假一致，
    #    比漏报更伤，所以如实报「无对应」交人核（依据里把话说全，见 _nospec_basis）。
    "no_spec":  {"blank": ("no_match", "no_spec_match"),
                 "na":    ("no_match", "no_spec_match")},
    "unknown":  {"blank": ("novalue", ""),       "na": ("novalue", "")},
}


def pdf_cell_state(res, er, shapes):
    """PDF 这一格到底是什么情况 → `MISSING_RULES` 的行键。

    判定顺序**本身就是规矩**（先硬证据、后推断）：
      ref 标记 / 写 N/A / 写空 都是 PDF 里**白纸黑字**的声明，最可靠；
      no_spec / no_step 是**结构性**结论；最后才轮到「有值 → 该不该填」，
      而这一问交给第 2.6 节的类型规矩表，本节不另写一套。
    """
    if not res:
        return "unknown"
    scope = res.get("scope")
    if scope == "ref_cell":
        return "ref_only"
    if res.get("na"):                       # 明确写 N/A（`------` 之外的另一种写法）
        return "na"
    if res.get("blank"):
        return "blank"
    if scope == "no_spec_match":
        return "no_spec"
    if scope == "no_step":
        return "no_step"
    if scope == "no_param":
        return "no_param"
    if scope in ("no_equip_match", "none"):
        return "unknown"
    vals = res.get("vals") or set()
    if not vals:
        return "unknown"
    rule = type_fill_rule(er.get("type"), dominant_shape(shapes))
    return "value" if rule["must_fill"] else "exempt"


def decide_missing_cell(er, res, shapes, excel_state="blank"):
    """缺失场景的唯一入口：Excel 这一格**没值**（空 / 填 N/A），对不对？

    返回 (status, scope, basis, expected)。
    """
    pstate = pdf_cell_state(res, er, shapes)
    status, scope = MISSING_RULES[pstate][excel_state]
    res = res or {}
    sub = res.get("sub") or "该档参数表"
    vals = res.get("vals") or set()
    shape = dominant_shape(shapes)
    is_na = (excel_state == "na")
    tail = "填 N/A 正确" if is_na else "不填正确"

    if pstate == "value":
        return (status, scope,
                "PDF %s 该步骤有值（%s），此栏%s —— 漏填"
                % (sub, _fmt_vals(vals), "却填了 N/A" if is_na else "是空白"),
                sorted(vals))
    if pstate == "exempt":
        rule = type_fill_rule(er.get("type"), shape)
        return (status, scope,
                "PDF %s 该步骤有值（%s），但按类型规矩这一项不必填（%s）—— %s"
                % (sub, _fmt_vals(vals, shape == "ref"), rule["why"], tail),
                sorted(vals))
    if pstate == "ref_only":
        return (status, scope,
                "PDF %s 这一格只标了 (ref)、没给数值（原文「%s」）—— 按类型规矩"
                "不必填值，也不限制超上下限 —— %s"
                % (sub, res.get("ref_raw") or "—", tail), [])
    if pstate == "na":
        return (status, scope,
                "PDF %s 这一格写的是「%s」（表格在声明此步不适用） —— %s"
                % (sub, _show_cell(res.get("raw") or "N/A"), tail), [])
    if pstate == "blank":
        return (status, scope,
                "PDF %s 里这一格本身就是空值（表格里写的是「%s」） —— %s"
                % (sub, _show_cell(res.get("raw") or "------"), tail), [])
    if pstate == "no_param":
        # 与 no_step 的差别要写进依据：一个是「没有这一步」，一个是「整张表
        # 没有这一列」。并列出这张表实际印了哪些参数 —— 那是「不是我漏读」的证据。
        return status, scope, _noparam_basis(res, er, " —— " + tail), []
    if pstate == "no_step":
        return status, scope, _nostep_basis(res, " —— " + tail), []
    if pstate == "no_spec":
        return status, scope, _nospec_basis(res, er), []
    return status, scope, "", []


def flag_contradiction(er, shape=None, op_raw="", balloon=False):
    """类型与「显示上下限 / 允许超上下限」两列**对不对得上**。

    返回 None（对得上、或无从判断）或一句中文说明（对不上）。

    为什么单独一个函数：类型规矩说「该填什么」，这两列说「实际填了什么」，
    两者对不上是**配置层的错** —— 比数值更靠前：类型写错、开关填反，
    下面数值比得再准也是白比（2026-09-29 用户点名：表4!311/369 是字符串，
    「显示上下限」却填了「是」，一直没被筛出来，因为字符串整类被当成
    「采集项不核对」跳过了）。

    两条**类型层面**的规矩也在这里判（`expected_type_for`）：球囊 WI 里 PDF 带
    `>` / `<` 的格子类型必须是「字符串」。`op_raw` 是 PDF 那一格的原文（带比较符
    才非空），`balloon` 是不是球囊 WI —— 两个都由 `compare()` 算好传进来。

    三条不判的边界（宁可放过，也不凭空造错）：
      · 表里那一格是空的 → 不判（没填 ≠ 填错）
      · 认不得的类型     → 不判
      · 浮点类**且 PDF 没给值** → 「允许超」该填什么说不准，只判「显示上下限」
    """
    typ = str(er.get("type") or "").strip()
    if typ not in TYPE_RULES:
        return None
    want_typ = expected_type_for(er, op_raw, balloon)
    # 规矩在类型层面表了态 → 按**应有的类型**去查两个开关（不是按表里填的类型）
    base = want_typ or typ
    r = type_fill_rule(base, shape)
    pairs = [("显示上下限", r["show_lim"], _yn(er.get("show_lim")))]
    # 浮点类的「允许超」要看 PDF 值形态；形态没定（PDF 没这一格）时不判
    if not (TYPE_RULES[base]["must_fill"] is None and shape in (None, "", "none")):
        pairs.append(("允许超上下限", r["allow_over"], _yn(er.get("allow_over"))))
    bad = []
    for cn, want, got in pairs:
        if want is None or got is None or got == want:
            continue
        bad.append("%s 应为「%s」，表里填的是「%s」" % (cn, "是" if want else "否",
                                                     "是" if got else "否"))
    if not bad:
        return None
    if want_typ and typ != want_typ:
        # 球囊专属：类型本身就不对，把 PDF 那一格原文摆出来，让人一眼看到是 `>90℃`
        return ("类型应为「%s」，表里填的是「%s」—— 球囊类 WI 里 PDF 那一格写着"
                "「%s」，带 > / < 的一律按「字符串」处理（值列全空、两个开关都否）"
                "%s" % (want_typ, typ, op_raw,
                        "；另：" + "；".join(bad) if bad else ""))
    return "类型=「%s」，但 %s —— 按类型规矩填反了" % (typ, "；".join(bad))


# ==========================================================================
#  第 3 节  单表自检（纯 Excel）
#  参数卡配置矛盾、必填、6 条填写规则
# ==========================================================================

# -*- coding: utf-8 -*-
"""表格自检（Excel 侧，完全不依赖 PDF）。

把工程参数表里**自带却从未被用过的列**变成可执行校验：类型 / 是否必填 /
是否显示上下限 / 是否允许超上下限 / 是否填写 N/A / 固定值 / 布尔值真值 /
参数卡版本 / 组序号 / 设备号。这些列本身就是「工艺参数卡填写规范」，
只是以前没人拿它去校验自己。

三条原则：
1. **结构自适应** —— 列按名字语义找（复用第 1 节的 _key 归一化），不写死列号与
   sheet 名；表头在第 1 行还是第 2 行都能认（rev03 的表头就在第 2 行）。
2. **不碰 PDF 链路** —— 本模块独立于 compare，改动风险为零。
3. **只报能被证伪的** —— 每条结论都带「工作表!行号 + 参数卡名 + 采集项名」定位。
   宁可少报也不误报：口径上凡是「全表统一的做法」一律不当问题（见下面的
   两处误报教训），只报**自相矛盾**与**同类不一致**。

对外接口：
    rs = audit_xlsx(path, sheet=None)   # sheet=None 时把每张明细表都自检一遍
    r.issues / r.stats / r.summary() / r.text()
"""


# ---------------------------------------------------------------- 审计列字典
# 只收「自检要用」的列。刻意与第 1 节的 FIELD_SPEC 分开，避免互相干扰映射。
AUDIT_FIELDS = {
    "card":       ["工艺参数卡名称", "参数卡名称", "参数卡名", "参数卡"],
    "card_ver":   ["工艺参数卡版本", "参数卡版本", "卡版本"],
    "txn":        ["事务", "事务类型"],
    "proc":       ["工艺名称", "工序名称", "工序"],
    "group":      ["组序号", "分组排序", "分组序号"],
    "item":       ["数据采集项名称", "采集项名称", "数据采集项"],
    "type":       ["类型", "数据类型"],
    "wi":         ["wi文档名称", "wi编号", "wi文档"],
    "wi_ver":     ["wi文档版本", "wi版本"],
    "show_lim":   ["是否显示上下限"],
    "allow_over": ["是否允许超上下限"],
    "required":   ["是否必填"],
    "allow_na":   ["是否填写n/a", "是否填写na", "允许填写n/a"],
    "lo":         ["下限", "最小值"],
    "hi":         ["上限", "最大值"],
    "fixed":      ["固定值"],
    "btrue":      ["布尔值真"],
    "bfalse":     ["布尔值假"],
    "spec":       ["适用规格", "适用尺寸", "规格"],
    "equip":      ["设备编号", "设备名称", "设备号"],
    "pn":         ["产品编码", "产品料号", "料号"],
}

# 这些字段在 Excel 里常是「合并单元格」或留空表示「同上」，必须向下继承
_FILL = ("card", "card_ver", "proc", "txn", "spec")
_CARRY = ("card", "card_ver", "proc", "txn")

# 「值需要预先设定」的类型。字符串/时间戳类型的值是人填的，不参与「有没有值」判定
# —— 曾经把这 5328 个字符串格子报成「必填未填」，实际它们本来就该是空的。
_VALUE_TYPES = ("浮点", "整数", "十进制", "固定值", "布尔值")

# 类型家族。同一采集项在 A 卡用「固定值」、B 卡用「浮点」是**正常业务表达**
# （能定值就定值、有区间就写区间）——同类内不同不算问题；跨家族才是真矛盾
# （「程序号」怎么会既是布尔值又是固定值？）。
_TYPE_FAMILY = {"浮点": "数值", "整数": "数值", "十进制": "数值", "固定值": "数值",
                "布尔值": "布尔", "字符串": "文本", "时间戳": "时间"}

# 需要绑定设备的类型 / 事务（填写规范第 3 条要排除它们）
_NO_EQUIP_TYPE = ("布尔值",)
_NO_EQUIP_TXN = ("清场",)

_YES = ("是", "y", "yes", "true", "1")
_NO = ("否", "n", "no", "false", "0", "")


def _yn(v):
    """把「是/否」类单元格归一成 True / False / None。"""
    if v is None:
        return None
    s = str(v).strip().lower()
    if s in _YES:
        return True
    if s in _NO:
        return False
    return None


# 同义词展开：长词优先且长词先匹配，避免「规格」抢走「适用规格」这类包含关系
_CAND = []
for _f, _syns in AUDIT_FIELDS.items():
    for _s in _syns:
        _CAND.append((_key(_s), _f))
_CAND.sort(key=lambda x: -len(x[0]))


def match_col(name):
    """列名 → 审计字段；认不出返回 None。精确匹配优先，再退到包含匹配。"""
    k = _key(name)
    if not k:
        return None
    for a, f in _CAND:
        if k == a:
            return f
    for a, f in _CAND:
        if a in k:
            return f
    return None


def _map_header(header):
    """表头行 → {字段: 列下标}。同一字段多个候选时取最靠左的。"""
    m, used = {}, set()
    for c, nm in enumerate(header):
        if nm is None or not str(nm).strip():
            continue
        f = match_col(nm)
        if f and f not in used:
            m[f] = c
            used.add(f)
    return m


_WEIGHT = {"item": 4, "type": 3, "required": 3, "lo": 2, "hi": 2, "fixed": 2,
           "show_lim": 2, "card": 2, "group": 1, "wi": 1, "equip": 1}


def _score(m):
    return sum(v for k, v in _WEIGHT.items() if k in m)


def _find_header(rows, limit=15):
    """在开头若干行里找「命中审计列最多」的一行，返回 0-based 下标。

    不用第 1 节的 find_header_row：那张判据不认识「是否必填」这类开关列，
    rev03 这种「第 1 行是说明、第 2 行才是表头」的表会被低估。
    """
    best, bi = 0, None
    for i in range(min(len(rows), limit)):
        m = _map_header(rows[i])
        s = len(m) + _score(m)
        if s > best:
            best, bi = s, i
    return bi


@dataclass
class Issue:
    cat: str          # 分类（配置矛盾 / 填写规范 / 跨卡不一致 / 空值缺口）
    level: str        # "bad" 必须改 / "warn" 需确认
    loc: str          # 工作表!行号
    what: str         # 问题描述
    got: str          # 当前实际值

    def key(self):
        return (self.cat, self.what)


@dataclass
class AuditResult:
    path: str = ""
    sheet: str = ""
    header_row: int = 0
    mapping: dict = dc_field(default_factory=dict)
    rows: int = 0
    issues: list = dc_field(default_factory=list)
    stats: dict = dc_field(default_factory=dict)
    notes: list = dc_field(default_factory=list)
    candidates: list = dc_field(default_factory=list)

    def by_kind(self):
        g = defaultdict(list)
        for it in self.issues:
            g[(it.level, it.what)].append(it)
        return g

    def counts(self):
        c = Counter()
        for it in self.issues:
            c[it.level] += 1
        return c

    # ------------------------------------------------------------ 渲染
    def summary(self, with_title=True, with_notes=True):
        """给界面用的 [(文本, 标签)]；标签沿用 app.py 已有的 t/h/ok/bad/warn/dim/note。

        with_notes=False 用于「一次自检多张表」：说明文字只该讲一遍，
        挂在每张表后面会重复刷屏。
        """
        out = []
        c = self.counts()
        if with_title:
            out.append(("工作表「%s」｜%d 行 ｜ 表头第 %d 行 ｜ 认出自检列 %d 个"
                        % (self.sheet, self.rows, self.header_row, len(self.mapping)), "t"))
            out.append(("  可校验列：%s"
                        % ("、".join(self.stats.get("col_labels", [])) or "（无）"), "dim"))
        if not self.issues:
            out.append(("  未发现配置矛盾与同类不一致。", "ok"))
        else:
            out.append(("  发现 %d 条问题（%d 条必须改 / %d 条需确认）"
                        % (len(self.issues), c.get("bad", 0), c.get("warn", 0)), "h"))

        g = self.by_kind()
        order = sorted(g.items(), key=lambda kv: (kv[0][0] != "bad", -len(kv[1])))
        for (level, what), items in order:
            tag = "bad" if level == "bad" else "warn"
            out.append(("  %s %-32s %5d 条"
                        % ("!" if level == "bad" else "·", what, len(items)), tag))
            for it in items[:3]:
                out.append(("      %-15s %s" % (it.loc, it.got[:50]), "note"))
            if len(items) > 3:
                out.append(("      …… 另有 %d 条" % (len(items) - 3), "dim"))
        for line in self.stats.get("lines", []):
            out.append(("  " + line, "note"))
        if with_notes:
            for n in self.notes:
                out.append(("  · " + n, "note"))
        return out

    def text(self, with_title=True):
        L = []
        if with_title:
            L.append("工作表「%s」· 数据 %d 行 · 表头第 %d 行"
                     % (self.sheet, self.rows, self.header_row))
            L.append("-" * 62)
        c = self.counts()
        L.append("发现 %d 条问题（必须改 %d / 需确认 %d）"
                 % (len(self.issues), c.get("bad", 0), c.get("warn", 0)))
        g = self.by_kind()
        order = sorted(g.items(), key=lambda kv: (kv[0][0] != "bad", -len(kv[1])))
        for (level, what), items in order:
            L.append("  [%s] %s —— %d 条" % ("必改" if level == "bad" else "确认", what, len(items)))
            for it in items[:20]:
                L.append("      %-15s %s" % (it.loc, it.got))
            if len(items) > 20:
                L.append("      …… 另有 %d 条" % (len(items) - 20))
        L.append("")
        for line in self.stats.get("lines", []):
            L.append(line)
        if self.notes:
            L.append("")
            for n in self.notes:
                L.append("· " + n)
        return "\n".join(L)


# ------------------------------------------------------------------ 单表自检

def _audit_sheet(ws, sn, hrow, m, path, candidates, user_notes):
    """对一张已定位表头的明细表跑全部自检。"""
    r = AuditResult(path=path, sheet=sn, header_row=hrow, mapping=m, candidates=candidates)
    ncol = max(m.values()) + 1
    child, anchors = merge_fill_maps(ws, ncol)
    anchor_val, carry = {}, {}
    kind_count = Counter()
    MAX_PER_KIND = 3000

    rows = 0
    cards, items, wis, pns, specs, equips = set(), set(), set(), set(), set(), set()
    type_cnt, txn_cnt = Counter(), Counter()
    req_tot = req_expected = req_filled = 0
    req_blank = []
    grp_last = {}
    # 跨卡一致性用：同名采集项 → Counter(类型)
    item_type = defaultdict(Counter)
    item_type_loc = {}
    # 格式混用用：固定值 / 参数卡版本 的数字格式 vs 文本格式
    fmt = {"fixed": {"num": [], "text": []}, "card_ver": {"num": [], "text": []}}

    for ri, raw in enumerate(ws.iter_rows(min_row=1, max_col=ncol, values_only=True), 1):
        for f in _FILL:
            c = m.get(f)
            if c is not None and (ri, c + 1) in anchors and c < len(raw):
                v = raw[c]
                if v is not None and str(v).strip():
                    anchor_val[(ri, c + 1)] = v
        if ri <= hrow:
            continue

        def g(f, _r=raw, _ri=ri):
            c = m.get(f)
            if c is None or c >= len(_r):
                return None
            v = _r[c]
            if v is not None and str(v).strip():
                if f in _FILL:
                    carry[f] = v
                return v
            if f in _FILL:
                a = child.get((_ri, c + 1))
                if a is not None and a in anchor_val:
                    return anchor_val[a]
                if f in _CARRY and f in carry:
                    return carry[f]
            return v

        item_raw = g("item")
        if item_raw is None or not str(item_raw).strip():
            continue
        rows += 1

        def S(f):
            v = g(f)
            return "" if v is None else str(v).strip()

        card = S("card") or "(未命名卡)"
        itm = S("item")
        cards.add(card)
        items.add(itm)
        for f, box in (("wi", wis), ("pn", pns), ("spec", specs), ("equip", equips)):
            if S(f):
                box.add(S(f))
        typ = S("type")
        txn = S("txn")
        type_cnt[typ] += 1
        if txn:
            txn_cnt[txn] += 1
        item_type[itm][typ] += 1
        item_type_loc.setdefault((itm, typ), "%s!r%d" % (sn, ri))

        loc = "%s!r%d" % (sn, ri)
        tag = "%s / %s" % (card[:26], itm[:26])
        lo, hi = to_num(g("lo")), to_num(g("hi"))
        lo_s, hi_s = S("lo"), S("hi")
        fx, bt, bf = S("fixed"), S("btrue"), S("bfalse")
        req, na = _yn(g("required")), _yn(g("allow_na"))
        show = _yn(g("show_lim"))
        over = _yn(g("allow_over"))
        eq = S("equip")

        def add(cat, level, what, got):
            if kind_count[what] >= MAX_PER_KIND:
                return
            kind_count[what] += 1
            r.issues.append(Issue(cat, level, loc, what, got))

        # -------- 一、配置自洽性（每行自己跟自己是否说得通）
        # 「类型」与「值填在哪一列」必须配套（用户 2026-09-24 定的规矩）：
        #   浮点 / 整数 / 十进制 → 值填在 下限 / 上限 两列（区间两侧都填，单边只填一侧）
        #   固定值            → 值填在 固定值 列
        # 填错列时判定层是**照读**的（容错，见 compare 的「哪一列真有值」），
        # 所以只有自检能提示。
        # ⚠ 别把这一类降成 warn：大表 rev03 实测有 31 行是
        #   「类型=固定值、固定值列空、值却填在下限/上限列」（如
        #   「步骤3的激光能量(mW)：」下限=440 上限=480）—— 那是真问题
        #   （要么类型该写浮点、要么值该挪到固定值列），必须保持 bad。
        if typ == "固定值" and not fx:
            if lo is not None or hi is not None:
                # 值在，只是列放错了 —— 比"列为空"更能说明该动哪里
                add("配置矛盾", "bad",
                    "类型=固定值，但值填在「下限/上限」列（固定值应填「固定值」列）",
                    "%s 下限=%s 上限=%s" % (tag, lo_s, hi_s))
            else:
                add("配置矛盾", "bad", "类型=固定值，但「固定值」列为空", tag)
        if lo is not None and hi is not None and lo > hi:
            add("配置矛盾", "bad", "下限 > 上限（写反）",
                "%s 下限=%s 上限=%s" % (tag, lo_s, hi_s))
        if typ == "布尔值" and not bt:
            add("配置矛盾", "bad", "类型=布尔值，但未定义「布尔值真」", tag)
        if req is True and na is True:
            add("配置矛盾", "warn", "必填=是，却又允许填 N/A（互相矛盾）", tag)
        if show is True and lo is None and hi is None:
            add("配置矛盾", "bad", "显示上下限=是，但上下限都为空", tag)
        if show is False and (lo is not None or hi is not None):
            add("配置矛盾", "warn", "显示上下限=否，却填了上下限",
                "%s 下限=%s 上限=%s" % (tag, lo_s, hi_s))
        # 「允许超上下限」这列原先**从没被任何规则用过**（schema 里收了，自检里没管）。
        # 它=「是」时上下限不生效：声明了上下限、又说允许超，等于白填 ——
        # 采集时任何值都能过，这一项实际上没有任何约束。
        # 2026-09-24 用户的实际案例：WI-0039 的「压力（Mpa）」，上限列填 0.52、
        # 下限留空、允许超上下限=是。而 PDF 写的是 `> 0.52Mpa`（**下限**约束）——
        # 正确填法应是「下限 0.52、上限留空、允许超上下限=否」，
        # 现在这样既填错了列、又把限值关掉了，两头都不对。
        # 全语料 17.5 万行实测：这一列 100% 是「否」，只有上面那 1 行是「是」
        # → 这个组合几乎可以确定是填错，但保留 warn（「需确认」）而不是 bad，
        # 因为「某参数确实不设限」也可能是有意的业务设置。
        if over is True and (lo is not None or hi is not None):
            add("配置矛盾", "warn", "允许超上下限=是，却填了上下限（限值将不生效）",
                "%s 下限=%s 上限=%s" % (tag, lo_s, hi_s))
        # 原先两句并成一句「上下限不完整」，会把两种完全不同的情况说成同一件事：
        #   · 两侧都没填 —— 大概率是真漏填；
        #   · 只填了一侧 —— 这**可能是有意的单边约束**（PDF 里 `> 0.52Mpa` 就是
        #     只有下限、没有上限）。若不分家，用户按规矩改成单边之后仍会看到
        #     一句像在报错的「不完整」，反而以为没改对（2026-09-24 实际对话）。
        # 反向：数值类型的值跑到「固定值」列，同样是填错列（与上面刚加的那条对称）。
        if typ in ("浮点", "整数", "十进制") and fx and lo is None and hi is None:
            add("配置矛盾", "warn",
                "类型是数值，但值填在「固定值」列（数值类型应填上下限）",
                "%s 固定值=%s" % (tag, fx))
        if typ in ("浮点", "整数", "十进制") and not fx:
            if lo is None and hi is None:
                add("配置矛盾", "warn", "类型是数值，但上下限都没填", tag)
            elif lo is None or hi is None:
                add("配置矛盾", "warn",
                    "类型是数值，但上下限只填了一侧（单边约束可忽略，请确认）", tag)
        if typ == "字符串" and (lo is not None or hi is not None):
            add("配置矛盾", "warn", "类型=字符串，却给了数值上下限", tag)

        # -------- 二、填写规范（变更记录里的 6 条注意事项）
        # 第 3 条：设备对应的参数采集项需要绑定设备号
        if (not eq or eq == "/") and typ not in _NO_EQUIP_TYPE and txn not in _NO_EQUIP_TXN:
            add("填写规范", "warn", "未绑定设备号（规范第3条）", tag)
        # 第 4 条：序号排序：相同步骤需要一致
        if m.get("group") is not None:
            gv = to_num(g("group"))
            if gv is not None:
                prev = grp_last.get(card)
                if prev is not None and gv <= prev:
                    add("填写规范", "warn", "同一参数卡内组序号未递增（规范第4条）",
                        "%s 上一序号=%s 本序号=%s" % (tag, prev, gv))
                grp_last[card] = gv
        # 第 5 条：字符串可填数字字母符号；浮点只能填数字
        if typ in ("浮点", "整数", "十进制"):
            for fld, val in (("下限", lo_s), ("上限", hi_s)):
                if val and not re.fullmatch(r"-?\d+(?:\.\d+)?", val.replace(",", "")):
                    add("填写规范", "bad", "类型是数值，但%s不是数字（规范第5条）" % fld,
                        "%s %s=%s" % (tag, fld, val))
        # 第 6 条：固定值 / 参数卡版本 需为文本格式 → 只在「整列格式混用」时才是问题
        for fld, nm in (("fixed", "固定值"), ("card_ver", "参数卡版本")):
            rawv = g(fld)
            if rawv is None or not str(rawv).strip():
                continue
            k = "num" if (isinstance(rawv, (int, float)) and not isinstance(rawv, bool)) else "text"
            bucket = fmt[fld][k]
            if len(bucket) < 50:
                bucket.append((loc, tag, str(rawv)))

        # -------- 三、必填项填空（只对「值需预设」的类型判，字符串不算空缺）
        if req is True:
            req_tot += 1
            if typ in _VALUE_TYPES:
                req_expected += 1
                if lo_s or hi_s or fx or bt or bf:
                    req_filled += 1
                elif len(req_blank) < 500:
                    req_blank.append((loc, "%s（类型=%s）" % (tag, typ)))

    r.rows = rows

    # -------- 四、跨卡不一致：同名采集项在不同参数卡里类型不同
    for itm, cnt in item_type.items():
        fams = {t: _TYPE_FAMILY.get(t, "其它") for t in cnt}
        if len(set(fams.values())) > 1:      # 跨家族才算矛盾
            detail = "、".join("%s×%d" % (k or "(空)", v) for k, v in cnt.most_common())
            top_k = cnt.most_common(1)[0][0]
            r.issues.append(Issue(
                "跨卡不一致", "bad", item_type_loc.get((itm, top_k), ""),
                "同名采集项在不同参数卡里类型跨类不一致",
                "%s → %s" % (itm[:30], detail)))

    # -------- 五、格式混用：同一列既有数字格式又有文本格式
    for fld, nm in (("fixed", "固定值"), ("card_ver", "参数卡版本")):
        nums, texts = fmt[fld]["num"], fmt[fld]["text"]
        if len(nums) >= 2 and len(texts) >= 2:
            for loc, tag, val in nums[:6]:
                r.issues.append(Issue(
                    "跨卡不一致", "warn", loc,
                    "%s列数字格式与文本格式混用（规范第6条）" % nm,
                    "%s %s=%s（数字格式；另 %d 处为文本格式）" % (tag, nm, val, len(texts))))

    # -------- 统计
    rate = (100.0 * req_filled / req_expected) if req_expected else 0.0
    lines = []
    lines.append("数据行 %d ｜ 参数卡 %d 张 ｜ 采集项名 %d 个" % (rows, len(cards), len(items)))
    if wis:
        lines.append("WI %d 个：%s" % (len(wis), "、".join(sorted(wis)[:8])))
    if pns:
        lines.append("产品编码 %d 个 ｜ 适用规格 %d 种 ｜ 设备 %d 种"
                     % (len(pns), len(specs), len(equips)))
    if type_cnt:
        lines.append("类型分布：%s" % "、".join("%s×%d" % (k or "(空)", v)
                                             for k, v in type_cnt.most_common(8)))
    if txn_cnt:
        lines.append("事务分布：%s" % "、".join("%s×%d" % (k, v)
                                             for k, v in txn_cnt.most_common(6)))
    if req_tot:
        lines.append("必填 %d 格：其中 %d 格需要预设值，已填 %d 格（%.2f%%）"
                     % (req_tot, req_expected, req_filled, rate))
    r.stats = {
        "col_labels": [AUDIT_FIELDS[f][0] for f in
                       ("type", "required", "show_lim", "allow_over", "allow_na",
                        "fixed", "btrue", "group", "equip") if f in m],
        "lines": lines,
        "required_total": req_tot,
        "required_expected": req_expected,
        "required_filled": req_filled,
        "required_rate": rate,
        "required_blank_sample": req_blank,
    }
    return r


# ------------------------------------------------------------------ 主入口

def locate_sheets(wb):
    """扫描已打开的工作簿，返回 [(sheet名, 表头行, 列映射, 总行数)]，只含明细表。

    「明细表」判据：含「采集项名称」列（item）。抽成独立函数是因为覆盖看板
    与版本链校验也要按同样的办法找表 —— 两处各写一套迟早会漂移。
    """
    cands = []
    for sn in wb.sheetnames:
        if getattr(wb[sn], "sheet_state", "visible") != "visible":
            continue
        ws = wb[sn]
        head = []
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=True), 1):
            head.append(row)
            if i >= 15:
                break
        hi = _find_header(head)
        if hi is None:
            continue
        m = _map_header(head[hi])
        if "item" not in m:
            continue                      # 没有采集项列 → 不是明细表
        cands.append((sn, hi + 1, m, ws.max_row or 0))
    return cands


def audit_xlsx(path: str, sheet: str | None = None) -> "AuditResult":
    """自检一份工程参数表。

    sheet=None 时会把**所有**含「采集项名称」列的明细表都自检一遍
    （大表里「rev03 录入表」和「表4 配置表」各有一类问题，只看一张会漏）；
    指定 sheet 则只自检那一张。返回 [AuditResult]。
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        cands = locate_sheets(wb)

        if not cands:
            r = AuditResult(path=path)
            r.notes.append("没有任何工作表含「采集项名称」列，无法自检。")
            return [r]

        cand_meta = [{"sheet": s, "cols": len(m), "score": _score(m), "rows": n}
                     for s, _h, m, n in cands]
        order = sorted(cands, key=lambda c: (-_score(c[2]), -c[3], c[0]))

        if sheet:
            hit = [c for c in order if c[0] == sheet or c[0].strip() == sheet.strip()]
            if hit:
                order = hit
            else:
                r = AuditResult(path=path)
                r.notes.append("指定的工作表「%s」不是采集项明细表；"
                               "可自检的是：%s" % (sheet, "、".join(c[0] for c in cands)))
                return [r]

        notes = []
        if len(order) > 1:
            notes.append("本次自检了 %d 张表：%s。每张表的问题类型不同，"
                         "（如「表4」是配置层、「rev03」是录入层），建议都看一遍。"
                         % (len(order), "、".join("%s(%d行)" % (c[0], c[3]) for c in order)))
        notes.append("单表自检只看表格自身是否自洽（不读 PDF 内容）；"
                     "「需确认」类可能是有意的业务设置，请按实际情况判断。")

        out = []
        for sn, hrow, m, _nr in order:
            out.append(_audit_sheet(wb[sn], sn, hrow, m, path, cand_meta, notes))
        for r in out:
            r.notes = list(notes)
        return out
    finally:
        wb.close()

# ==========================================================================
#  第 4 节  覆盖面与 WI 版本链
#  录了数据却没 PDF 可核；版本号三方对账
# ==========================================================================

# -*- coding: utf-8 -*-
"""工作簿级校验：覆盖面看板 + WI 版本链。

单表自检（第 3 节）回答「这张表自己说不说得通」；
本模块回答「这些文件凑在一起，有没有整块压根没被核过」：

  · 录入了数据的 WI，手上有没有对应的作业指导书 PDF？
  · 表里登记为「运行」的 WI，文件在不在？
  · 每行声明的 WI 版本，与登记版本、与 PDF 文件名版本，是否对得上？

这三问以前没人问。后果是实打实的：rev03 里有 59274 行数据，其中
WI-0018 / WI-0019 / WI-0022 共 2664 行**从来没有 PDF 可核**，但工具
只会逐行报「无对应」，不会说「我根本没这份文件」—— 从界面上完全看不出
这是「数据不对」还是「缺文件」。覆盖面看板就是为了把这两件事分开。

对外接口：
    r = audit_workbook(xlsx, pdf_paths=[...])
    r.summary() / r.text()
"""


WI_RE = re.compile(r"WI[-\s]?(\d{3,4})", re.I)
REV_RE = re.compile(r"rev\s*(\d+)", re.I)

# 登记表里表示「这版正在用」的状态词
_ACTIVE = ("运行", "启用", "生效", "使用中", "有效")


def wi_of(name):
    """从任意文本里抽 WI 编号，如 'WI-0024, rev06 尖端成型…' → 'WI-0024'。"""
    m = WI_RE.search(str(name or ""))
    return ("WI-%s" % m.group(1)) if m else ""


def rev_of(name):
    """从任意文本里抽版本号，统一写成 'revNN'；抽不到返回 ''。"""
    m = REV_RE.search(str(name or ""))
    return ("rev%s" % m.group(1)) if m else ""


def _norm_rev(v):
    """把 'rev06' / '06' / '6' / 'rev6' 都归一到 'rev6'，便于比较。"""
    s = str(v or "").strip()
    m = re.search(r"(\d+)", s)
    return ("rev%s" % int(m.group(1))) if m else ""


@dataclass
class CoverageReport:
    xlsx: str = ""
    registry_sheet: str = ""                       # 版本登记表所在 sheet
    scanned_sheets: list = dc_field(default_factory=list)
    data: dict = dc_field(default_factory=dict)     # {WI: {"rows","vers","sheets"}}
    registry: dict = dc_field(default_factory=dict)  # {WI: {"ver","status"}}
    pdfs: dict = dc_field(default_factory=dict)      # {WI: {"ver","file"}}
    pdf_unknown: list = dc_field(default_factory=list)   # 文件名里认不出 WI 编号的 PDF
    issues: list = dc_field(default_factory=list)
    notes: list = dc_field(default_factory=list)

    # ------------------------------------------------------------ 统计
    def data_rows(self):
        return sum(v["rows"] for v in self.data.values())

    def uncovered_rows(self):
        """有录入数据、却没有 PDF 可核的行数 —— 这是最该被看见的数字。"""
        n = 0
        for wi, v in self.data.items():
            if wi not in self.pdfs:
                n += v["rows"]
        return n

    def counts(self):
        c = Counter()
        for it in self.issues:
            c[it.level] += 1
        return c

    # ------------------------------------------------------------ 渲染
    def summary(self):
        out = []
        c = self.counts()
        tot, unc = self.data_rows(), self.uncovered_rows()
        out.append(("覆盖面", "t"))
        out.append(("  本表录入 %d 行，覆盖 %d 个 WI ｜ 手上 PDF %d 份（%d 个 WI）"
                    % (tot, len(self.data), len(self.pdfs) + len(self.pdf_unknown), len(self.pdfs)),
                    "dim"))
        if unc:
            out.append(("  其中 %d 行没有对应 PDF，无法核对（占 %.1f%%）"
                        % (unc, 100.0 * unc / tot if tot else 0), "bad"))
        else:
            out.append(("  所有录入数据的 WI 都有 PDF 可核。", "ok"))
        out.append(("", None))

        # 三类缺口清单
        no_pdf = [w for w in self.data if w not in self.pdfs]
        no_data = [w for w in self.pdfs if w not in self.data]
        reg_no_pdf = [w for w, v in self.registry.items()
                      if w not in self.pdfs and any(k in v["status"] for k in _ACTIVE)]
        if no_pdf:
            out.append(("  有录入数据、但缺 PDF（%d 个）" % len(no_pdf), "h"))
            for wi in sorted(no_pdf):
                v = self.data[wi]
                vers = "、".join(k for k, _ in v["vers"].most_common(3)) or "未标注版本"
                out.append(("      %-9s %6d 行  声明版本 %-18s 来源 %s"
                            % (wi, v["rows"], vers, "、".join(sorted(v["sheets"]))), "bad"))
        if reg_no_pdf:
            out.append(("  登记为「运行」、但缺 PDF（%d 个）" % len(reg_no_pdf), "h"))
            for wi in sorted(reg_no_pdf):
                v = self.registry[wi]
                out.append(("      %-9s 登记版本 %s" % (wi, v["ver"] or "?"), "warn"))
        if no_data:
            out.append(("  有 PDF、但本表没有录入数据（%d 个）" % len(no_data), "h"))
            for wi in sorted(no_data):
                out.append(("      %-9s %s" % (wi, self.pdfs[wi]["file"][:54]), "note"))
            out.append(("      这些 WI 的录入数据通常在另一份参数表里 —— "
                        "两份表的 WI 一般不重叠，属正常。", "dim"))
        if self.pdf_unknown:
            out.append(("  文件名里认不出 WI 编号的 PDF（%d 份）" % len(self.pdf_unknown), "h"))
            for f in self.pdf_unknown[:6]:
                out.append(("      " + f[:62], "note"))
        if not (no_pdf or reg_no_pdf or no_data):
            out.append(("  覆盖面完整：数据、登记、文件三方 WI 对得上。", "ok"))
        out.append(("", None))

        # 版本链
        vers = [i for i in self.issues if i.cat == "版本链"]
        if vers:
            out.append(("版本链（%d 条）" % len(vers), "h"))
            for it in vers:
                out.append(("  %s %-9s %s" % ("!" if it.level == "bad" else "·",
                                              it.loc, it.what), 
                            "bad" if it.level == "bad" else "warn"))
                if it.got:
                    out.append(("      " + it.got[:64], "note"))
        else:
            out.append(("版本链", "h"))
            out.append(("  录入版本、登记版本与 PDF 文件名版本一致。", "ok"))
        out.append(("", None))
        for n in self.notes:
            out.append(("  · " + n, "dim"))
        return out

    def text(self):
        L = ["【覆盖面与版本链】%s" % self.xlsx]
        L.append("=" * 62)
        tot, unc = self.data_rows(), self.uncovered_rows()
        L.append("录入 %d 行 / %d 个 WI ｜ 手上 PDF %d 份 ｜ 无 PDF 可核 %d 行"
                 % (tot, len(self.data), len(self.pdfs), unc))
        L.append("")
        L.append("逐 WI 对账：")
        L.append("  %-9s %8s %-18s %-12s %-12s %s" % ("WI", "录入行数", "声明版本", "登记版本", "PDF版本", "判定"))
        for wi in sorted(set(self.data) | set(self.registry) | set(self.pdfs)):
            d = self.data.get(wi)
            g = self.registry.get(wi, {})
            p = self.pdfs.get(wi, {})
            dv = "、".join(k for k, _ in d["vers"].most_common(3)) if d else "—"
            gv = g.get("ver") or "—"
            pv = p.get("ver") or "—"
            bad = []
            if d and wi not in self.pdfs:
                bad.append("缺PDF")
            if d and gv != "—" and d["vers"] and len(d["vers"]) == 1 and _norm_rev(dv) != _norm_rev(gv):
                bad.append("版本≠登记")
            if d and pv != "—" and d["vers"] and len(d["vers"]) == 1 and _norm_rev(dv) != _norm_rev(pv):
                bad.append("版本≠PDF")
            if p and gv != "—" and pv != "—" and _norm_rev(pv) != _norm_rev(gv):
                bad.append("PDF≠登记")
            if not d and p:
                verdict = "仅PDF，本表无录入数据"
            elif bad:
                verdict = "、".join(bad)
            else:
                verdict = "OK"
            L.append("  %-9s %8s %-18s %-12s %-12s %s"
                     % (wi, d["rows"] if d else "—", dv, gv, pv, verdict))
        L.append("")
        if self.issues:
            L.append("问题清单：")
            for it in self.issues:
                L.append("  [%s] %-9s %s" % ("必改" if it.level == "bad" else "确认", it.loc, it.what))
                if it.got:
                    L.append("        " + it.got)
        if self.notes:
            L.append("")
            for n in self.notes:
                L.append("· " + n)
        return "\n".join(L)


# ------------------------------------------------------------------ 主流程

def _read_registry(wb):
    """找「WI 版本登记表」：含 WI + 版本列的小表。返回 (sheet, {WI: {...}})。

    必须排除明细表 —— rev03 这类录入表同样有「WI文档名称 / WI文档版本」两列，
    若不排除，它会以「WI 数不少」抢走登记表的位置，把版本链校验变成自己跟自己比。
    """
    best = (None, {}, -1)
    for sn in wb.sheetnames:
        if sn.startswith("hidden_") or getattr(wb[sn], "sheet_state", "visible") != "visible":
            continue
        ws = wb[sn]
        head = []
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), 1):
            head.append(row)
            if i >= 10:
                break
        hi = _find_header(head)
        if hi is None:
            continue
        m = _map_header(head[hi])
        if "wi" not in m or "wi_ver" not in m:
            continue                      # 要同时有 WI 与版本列才算登记表
        if "item" in m:
            continue                      # 明细表（录入表）不是登记表
        status_col = None
        for c, nm in enumerate(head[hi]):
            if nm and "状态" in str(nm):
                status_col = c
                break
        reg = {}
        for r in ws.iter_rows(min_row=hi + 2, values_only=True):
            if m["wi"] >= len(r):
                continue
            wi = wi_of(r[m["wi"]])
            if not wi:
                continue
            ver = r[m["wi_ver"]] if m["wi_ver"] < len(r) else None
            status = ""
            if status_col is not None and status_col < len(r):
                status = str(r[status_col] or "").strip()
            reg[wi] = {"ver": str(ver or "").strip(), "status": status}
        score = (100 if status_col is not None else 0) + len(reg)
        if score > best[2]:
            best = (sn, reg, score)
    return best[0], best[1]


def audit_workbook(xlsx_path: str, pdf_paths: tuple = (),
                   scan_dir: str | None = None) -> "CoverageReport":
    """覆盖面 + 版本链校验。

    xlsx_path : 工程参数表
    pdf_paths : 已知的 PDF 路径（一般是界面上选的那份 + 同目录其它 PDF）
    scan_dir  : 若给出，则把该目录下所有 *.pdf 一并纳入（用于「手上到底有几份」）
    """
    r = CoverageReport(xlsx=xlsx_path)
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        # ---- 1) 明细表里的 WI 与声明版本
        cands = locate_sheets(wb)
        r.scanned_sheets = [c[0] for c in cands]
        for sn, hrow, m, _nr in cands:
            if "wi" not in m:
                r.notes.append("工作表「%s」没有 WI 列，未计入覆盖面统计"
                               "（它可能要靠参数卡关联表才能推到 WI）。" % sn)
                continue
            ws = wb[sn]
            ncol = max(m.values()) + 1
            for row in ws.iter_rows(min_row=hrow + 1, max_col=ncol, values_only=True):
                if m["wi"] >= len(row):
                    continue
                wi = wi_of(row[m["wi"]])
                if not wi:
                    continue
                cell = r.data.setdefault(wi, {"rows": 0, "vers": Counter(), "sheets": set()})
                cell["rows"] += 1
                cell["sheets"].add(sn)
                if "wi_ver" in m and m["wi_ver"] < len(row):
                    v = str(row[m["wi_ver"]] or "").strip()
                    if v:
                        cell["vers"][v] += 1

        # ---- 2) 登记表
        rsheet, reg = _read_registry(wb)
        r.registry_sheet = rsheet or ""
        r.registry = reg
    finally:
        wb.close()

    # ---- 3) 手上的 PDF（文件名是唯一的版本来源，所以文件名必须规范）
    paths = list(pdf_paths)
    if scan_dir and os.path.isdir(scan_dir):
        for fn in sorted(os.listdir(scan_dir)):
            if fn.lower().endswith(".pdf"):
                p = os.path.join(scan_dir, fn)
                if p not in paths:
                    paths.append(p)
    for p in paths:
        fn = os.path.basename(p)
        wi = wi_of(fn)
        if not wi:
            r.pdf_unknown.append(fn)
            continue
        ver = rev_of(fn)
        old = r.pdfs.get(wi)
        if old is None:
            r.pdfs[wi] = {"ver": ver, "file": fn}
        else:
            # 同一 WI 有多份 PDF：保留版本号更大的那份，并注明另有副本
            if _norm_rev(ver) > _norm_rev(old["ver"]):
                r.pdfs[wi] = {"ver": ver, "file": fn}
                r.notes.append("同一个 %s 有多个 PDF，已按较大版本号采用「%s」；另有「%s」"
                               % (wi, fn, old["file"]))
            else:
                r.notes.append("同一个 %s 有多个 PDF：「%s」与「%s」，已采用版本号较大的那份。"
                               % (wi, fn, old["file"]))

    # ---- 4) 版本链校验
    for wi, d in sorted(r.data.items()):
        loc = wi
        # 4a 本表内出现多个版本声明
        if len(d["vers"]) > 1:
            r.issues.append(Issue(
                "版本链", "bad", loc, "同一 WI 在本表内出现多个版本声明",
                "、".join("%s×%d 行" % (k, v) for k, v in d["vers"].most_common())))
        dv = d["vers"].most_common(1)[0][0] if d["vers"] else ""
        dvn = _norm_rev(dv)
        # 4b 与登记版本不符
        g = r.registry.get(wi)
        if dv and g and g["ver"] and dvn and _norm_rev(g["ver"]) != dvn:
            r.issues.append(Issue(
                "版本链", "bad", loc, "录入数据声明的 WI 版本与登记表不一致",
                "本表 %s，登记表「%s」%s；本 WI 在 %s 里共 %d 行"
                % (dv, r.registry_sheet, g["ver"], "、".join(sorted(d["sheets"])), d["rows"])))
        # 4c 与 PDF 文件名版本不符
        p = r.pdfs.get(wi)
        if dv and p and p["ver"] and dvn and _norm_rev(p["ver"]) != dvn:
            r.issues.append(Issue(
                "版本链", "bad", loc, "录入数据依据的 WI 版本与手上 PDF 不一致",
                "本表 %s，PDF 文件名 %s（%s）—— 数据可能还停留在旧版 WI 上"
                % (dv, p["ver"], p["file"][:44])))
        # 4d PDF 与登记表不符
        if p and g and p["ver"] and g["ver"] and _norm_rev(p["ver"]) != _norm_rev(g["ver"]):
            r.issues.append(Issue(
                "版本链", "warn", loc, "PDF 文件名版本与登记表不一致",
                "PDF %s，登记表 %s" % (p["ver"], g["ver"])))

    # ---- 5) 覆盖面结论（缺文件的严重性远高于「有文件没数据」）
    for wi, d in sorted(r.data.items()):
        if wi not in r.pdfs:
            r.issues.append(Issue(
                "覆盖面", "bad", wi, "有录入数据，但没有对应的 PDF，整块无法核对",
                "%d 行（来源 %s）｜声明版本 %s"
                % (d["rows"], "、".join(sorted(d["sheets"])),
                   "、".join(k for k, _ in d["vers"].most_common()) or "未标注")))
    for wi, g in sorted(r.registry.items()):
        if wi not in r.pdfs and any(k in g["status"] for k in _ACTIVE):
            r.issues.append(Issue(
                "覆盖面", "warn", wi, "登记为「%s」但没有 PDF 文件" % g["status"],
                "登记版本 %s" % (g["ver"] or "?")))

    if r.scanned_sheets:
        r.notes.append("参与统计的明细表：%s。" % "、".join(r.scanned_sheets))
    if r.registry_sheet:
        r.notes.append("版本登记表取自「%s」（含 WI 与版本两列的那张）。" % r.registry_sheet)
    else:
        r.notes.append("没找到「WI 版本登记表」（需同时含 WI 列与版本列），"
                       "版本链只能与 PDF 文件名对账。")
    return r

# ==========================================================================
#  第 5 节  同规格横向一致性
#  同规格组内找离群格子
# ==========================================================================

# -*- coding: utf-8 -*-
"""同规格横向一致性扫描：把「绝大多数产品都这么写、个别产品不一样」的格子捞出来。

为什么需要它：rev03 是 333 个产品编码 × 1998 张参数卡 = 59274 行，逐行复制粘贴
出来的。人工检查不可能，工具逐行核 PDF 也只能发现「与规定不符」，发现不了
「这一格和其它 332 个同类产品写得不一样」—— 而后者恰恰是漏改、错填最典型的样子。

分组键必须是「同一套工艺参数真的应该相同」的那一层：
    WI + 工序 + 采集项 + 直径 + 括注
不含直径会大面积误报 —— 2.00 与 2.50 的焊接参数本来就不同，那是设计差异，
不是错误。含直径后剩下的差异才是「同类里冒出来的一个」。

所以本模块只报**离群**：组内存在明显多数（≥75%），少数派才被列出来。
两边势均力敌时不下结论 —— 那更可能是两套并行工艺，不是错。

对外接口：
    r = scan_outliers(xlsx)
    r.summary() / r.text()
"""


# 规格形如 2.00X10(150) / 2.5*100（70） ；抓 直径 / 长度 / 括注
_SPEC_RE = re.compile(
    r"^\s*[φΦ]?\s*(\d+(?:\.\d+)?)\s*[xX*×]\s*(\d+(?:\.\d+)?)\s*"
    r"(?:[（(]\s*([^)）]*?)\s*[)）])?\s*$")


def spec_group(spec):
    """把规格归到「同类」：返回分组键字符串。认不出就用原值（保守）。"""
    s = str(spec or "").strip()
    if not s:
        return "(未填规格)"
    m = _SPEC_RE.match(s)
    if not m:
        return s                      # 如 'ALL' / 'φ1.0-1.25' → 原样作组
    dia, length, paren = m.group(1), m.group(2), (m.group(3) or "")
    return "%s(φ%s)" % (paren or "-", dia)


def _numtxt(v):
    """数值显示：整数不带小数点（2.0 → 2），便于「2 与 2.0」视为同一个值。"""
    if v is None:
        return ""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        f = float(v)
        return str(int(f)) if f == int(f) else ("%g" % f)
    return str(v).strip()


# 「不是值」的写法。rev03 里 固定值='N/A' 出现 14659 次 —— 那是「这一步不存在」
# 的业务标记。若把它当值参与比较，它会成了组内多数派，反而把真正的工艺值
# 报成离群（实测过：36 处 N/A 碾压 9 处 460~560）。这类格子一律不参与比较。
_NA_TOKENS = ("n/a", "na", "none", "null", "-", "--", "—", "/", "无", "不适用", "不涉及")


def _is_na(s):
    return str(s or "").strip().lower() in _NA_TOKENS


def value_of(lo, hi, fixed, btrue):
    """把一个采集项格子的「值」归一成可比较的字符串；没有可比较值时返回 ''。"""
    if _is_na(lo) or _is_na(hi) or _is_na(fixed):
        return ""                       # 「不适用」不是值
    has_lo = lo is not None and str(lo).strip() != ""
    has_hi = hi is not None and str(hi).strip() != ""
    if has_lo or has_hi:
        a, b = _numtxt(to_num(lo)) if has_lo else "", _numtxt(to_num(hi)) if has_hi else ""
        if a and b and a != b:
            return "%s~%s" % (a, b)
        return a or b
    fx = str(fixed or "").strip()
    if fx:
        return fx
    bt = str(btrue or "").strip()
    if bt:
        return bt
    return ""


@dataclass
class Outlier:
    wi: str
    proc: str
    item: str
    grp: str
    loc: str
    card: str
    pn: str
    val: str
    major: str
    major_n: int
    total: int
    level: str = "bad"    # bad=零星（漏改典型） / warn=成组（可能是有意的另一套工艺）

    def kind(self):
        return (self.wi, self.proc, self.item)


@dataclass
class OutlierReport:
    path: str = ""
    sheet: str = ""
    rows: int = 0
    scanned: int = 0                 # 有效分组数
    with_minor: int = 0              # 有离群值的分组数
    outliers: list = dc_field(default_factory=list)
    notes: list = dc_field(default_factory=list)
    min_group: int = 4
    minor_ratio: float = 0.25

    def by_item(self, level=None):
        g = defaultdict(list)
        for o in self.outliers:
            if level and o.level != level:
                continue
            g[(o.wi, o.proc, o.item)].append(o)
        return g

    def by_product(self):
        """按产品编码聚合：{产品编码: {采集项种类}}。

        实测里最有信息量的一屏就是这个 —— 426 处离群其实只涉及 9 个产品编码，
        说明那不是零散笔误，而是「这一批产品和其余 36 个不一样」。
        按采集项列出来看不出这件事，按产品列出来一眼就明白。
        """
        g = defaultdict(set)
        for o in self.outliers:
            g[o.pn].add(o.kind())
        return g

    def summary(self):
        out = []
        out.append(("横向一致性", "t"))
        out.append(("  扫描 %d 行 ｜ 有效分组 %d 个（同一 WI+工序+采集项+直径）｜ 其中 %d 个组存在离群值"
                    % (self.rows, self.scanned, self.with_minor), "dim"))
        if not self.outliers:
            out.append(("  未发现「同类里冒出来」的格子。", "ok"))
            out.append(("", None))
            for n in self.notes:
                out.append(("  · " + n, "dim"))
            return out

        for lv, title, tag, brief in (
                ("bad", "零星离群（个别产品与其余不同，最像漏改）", "bad", False),
                ("warn", "成组差异（一小批产品与其余不同，更像另一套工艺）", "warn", True)):
            items_by = self.by_item(lv)
            n = sum(len(v) for v in items_by.values())
            if not n:
                continue
            out.append(("  %s —— %d 处 / %d 个采集项" % (title, n, len(items_by)), "h"))
            for (wi, proc, item), items in sorted(items_by.items(),
                                                  key=lambda kv: -len(kv[1]))[:6 if brief else 10]:
                o0 = items[0]
                out.append(("     %-9s %-8s %-22s %4d 处  组内多数 %s（%d/%d）"
                            % (wi, proc[:8], item[:22], len(items),
                               o0.major, o0.major_n, o0.total), tag))
                # 成组差异动辄几百处，逐条列出来只会刷屏 —— 这种「一整批产品
                # 都和其余不同」的情形，看它涉及哪些产品比看它落在哪些行更有效。
                for o in items[:(2 if brief else 3)]:
                    out.append(("         %-14s %-11s 值 %-9s（组内 %d 处）"
                                % (o.loc, o.pn, o.val, o.total), "note"))
                if len(items) > (2 if brief else 3):
                    out.append(("         …… 另有 %d 处" % (len(items) - (2 if brief else 3)), "dim"))
            if len(items_by) > (6 if brief else 10):
                out.append(("     …… 另有 %d 个采集项" % (len(items_by) - (6 if brief else 10)), "dim"))
            out.append(("", None))

        bp = self.by_product()
        if bp:
            top = sorted(bp.items(), key=lambda kv: (-len(kv[1]), kv[0]))
            kmax = len(top[0][1])
            shown = [p for p, k in top if len(k) == kmax]
            out.append(("  据以判定的产品（共 %d 个产品编码出现在离群里）" % len(bp), "h"))
            out.append(("     其中最集中：%s%s（各命中 %d 个采集项）"
                        % ("、".join(shown[:10]),
                           " 等 %d 个" % len(shown) if len(shown) > 10 else "",
                           kmax), "note"))
            out.append(("", None))

        for n in self.notes:
            out.append(("  · " + n, "dim"))
        return out

    def text(self):
        L = ["【横向一致性】%s" % self.path]
        L.append("工作表「%s」· 扫描 %d 行 · 有效分组 %d 个 · 有离群值的组 %d 个"
                 % (self.sheet, self.rows, self.scanned, self.with_minor))
        L.append("=" * 62)
        L.append("判定口径：同一 (WI + 工序 + 采集项 + 直径) 下，若 ≥75% 的行取值相同，")
        L.append("          则其余少数行的取值列为离群（疑似漏改/错填）。")
        L.append("          「N/A / 无 / -」这类「这一格不适用」的标记不参与比较。")
        L.append("")
        if not self.outliers:
            L.append("未发现离群值。")
        for (wi, proc, item), items in sorted(self.by_item().items(), key=lambda kv: -len(kv[1])):
            o0 = items[0]
            L.append("  %s ｜ %s ｜ %s —— %d 处；组内多数 %s（%d 处）"
                     % (wi, proc, item, len(items), o0.major, o0.major_n))
            for o in items[:20]:
                L.append("      %-15s %-14s 值 %-10s 卡 %s（该组 %d 处）"
                         % (o.loc, o.pn, o.val, o.card[:24], o.total))
            if len(items) > 20:
                L.append("      …… 另有 %d 处" % (len(items) - 20))
        if self.notes:
            L.append("")
            for n in self.notes:
                L.append("· " + n)
        return "\n".join(L)


# ------------------------------------------------------------------ 主流程

def scan_outliers(xlsx_path: str, sheet: str | None = None, min_group: int = 4,
                  minor_ratio: float = 0.25,
                  max_outliers: int = 500) -> "OutlierReport":
    """扫描同规格横向离群值。

    min_group   : 组内至少多少行才判（小样本不下结论）
    minor_ratio : 少数派上限比例；同时要求多数派 ≥ 1-minor_ratio
    """
    r = OutlierReport(path=xlsx_path, min_group=min_group, minor_ratio=minor_ratio)
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    try:
        cands = locate_sheets(wb)
        if not cands:
            r.notes.append("没有找到含「采集项名称」列的明细表。")
            return r
        order = sorted(cands, key=lambda c: (-_score(c[2]), -c[3], c[0]))
        if sheet:
            hit = [c for c in order if c[0] == sheet or c[0].strip() == sheet.strip()]
            order = hit or order
        sn, hrow, m, _nr = order[0]
        r.sheet = sn

        missing = [k for k in ("item", "spec", "wi") if k not in m]
        if missing:
            r.notes.append("该表缺少 %s 列，无法按「WI+工序+采集项+直径」分组，"
                           "横向一致性跳过。" % "、".join(missing))
            return r

        ws = wb[sn]
        ncol = max(m.values()) + 1
        groups = defaultdict(list)      # 组键 -> [(loc, card, pn, val)]
        rows = 0
        for ri, row in enumerate(ws.iter_rows(min_row=hrow + 1, max_col=ncol, values_only=True),
                                 hrow + 1):

            def S(f):
                c = m.get(f)
                if c is None or c >= len(row):
                    return ""
                return "" if row[c] is None else str(row[c]).strip()

            item = S("item")
            if not item:
                continue
            val = value_of(row[m["lo"]] if "lo" in m and m["lo"] < len(row) else None,
                           row[m["hi"]] if "hi" in m and m["hi"] < len(row) else None,
                           S("fixed"), S("btrue"))
            if not val:
                continue                # 字符串类/空值不参与横向比较
            rows += 1
            key = (S("wi"), S("proc"), item, spec_group(S("spec")))
            groups[key].append(("%s!r%d" % (sn, ri), S("card"), S("pn"), val))

        r.rows = rows
        # ---- 逐组找离群
        kind_count = Counter()
        for key, itemsv in groups.items():
            if len(itemsv) < min_group:
                continue
            r.scanned += 1
            cnt = Counter(v for _l, _c, _p, v in itemsv)
            if len(cnt) < 2:
                continue
            total = len(itemsv)
            major, major_n = cnt.most_common(1)[0]
            if major_n < total * (1 - minor_ratio):
                continue                # 没有明显多数 → 可能是两套并行工艺，不下结论
            r.with_minor += 1
            cap = max(1, int(total * minor_ratio))
            for loc, card, pn, v in itemsv:
                if v == major or cnt[v] > cap:
                    continue
                if kind_count[key[2]] >= max_outliers:
                    continue
                kind_count[key[2]] += 1
                # 零星几个不同 = 漏改最典型的样子；一小批产品成组不同
                # 更可能是「这批按另一套工艺做」，所以分开提示，别混在一起吓人。
                lv = "bad" if cnt[v] <= 3 else "warn"
                r.outliers.append(Outlier(
                    wi=key[0], proc=key[1], item=key[2], grp=key[3],
                    loc=loc, card=card, pn=pn, val=v,
                    major=major, major_n=major_n, total=total, level=lv))
    finally:
        wb.close()

    r.notes.append("口径：同 (WI+工序+采集项+直径) 内 ≥75% 一致，其余列为离群；"
                   "势均力敌（无多数派）的组不列入 —— 那更可能是两套并行工艺。")
    r.notes.append("「N/A / 无 / -」等「这一格不适用」的标记不参与比较，"
                   "否则它们会形成假多数派，把真正的工艺值报成离群。")
    r.notes.append("离群不等于错误，但「绝大多数产品都那么写、只有它不一样」"
                   "是漏改最常见的形态，值得逐个回看。")
    return r

# ==========================================================================
#  第 6 节  导出 xlsx 与历史快照
#  带筛选的 xlsx；跨版本看变化
# ==========================================================================

# -*- coding: utf-8 -*-
"""导出（Excel）与历史快照对比。

两块职责：

1. **export_sheets** —— 把结果写成带筛选、冻结首行的 .xlsx。
   只有纯文本 .txt 时，几千条明细没法筛选、没法排序、没法按档位归类，
   实际用起来还是要人肉再抄一遍。

2. **历史快照** —— 每次比对存一份「判定签名」，下次比对时自动对比，回答
   「这次改版，修好了几条、又新引入几条、哪些换了依据」。
   用户一直被「改了很多遍看着没变化」困扰，根因之一正是没有任何跨版本
   对比依据 —— 只能靠记忆，而记忆在几千行面前是无效的。

签名键用 (卡片, 采集项, 同类序号)：同一张卡的同名采集项可能有多行
（不同设备/规格），所以必须加序号，否则会互相覆盖 —— 这个坑之前
在界面上踩过一次（16650 行里 13320 行被覆盖显示错）。
"""


_HDR_FILL = PatternFill("solid", fgColor="EEF2F7")
# ⚠️ `_BAD_FONT` / `_OK_FONT`（不一致标红、一致标绿）已删（2026-09-30 重构）：
# `export_sheets` 里只用了表头填充，状态从来没标过色 —— 是个没接上的半成品。
# 哪天要做「不一致高亮」，在这里建常量、**不要**在导出函数里内联颜色。

# 状态 -> 中文，全项目唯一定义。
# 以前这段映射在 app._export 内联、app._FLAG 类常量、batch.run_batch 内联里各写一遍，
# 改一处必然漂移；现在统一引用这里，任何一处要改都只改这一行。
# 键与取值必须与历史字符串逐字一致（回归对拍保证）。
STATUS_CN = {
    "ok": "一致",
    "mismatch": "不一致",
    "multi": "多套候选",
    "no_match": "无对应",
    "no_scope": "无法判定(缺设备线索)",
    "collection": "采集项(不核对)",
    "novalue": "未填值",
}


def status_cn_of(chk):
    """单行的判定中文。与 STATUS_CN 只差 novalue 这一类，其余原样透传。

    novalue 里其实混着两件完全不同的事：
      · 格子真的空着                        → 「未填值」
      · 格子**填了**，但填的是文字/代码
        （`定位标记模板：T-2046-56` 这种），
        取不出可比较的数，所以没参与核对   → 「未参与核对」

    2026-10-29 用户点名：「表格明明填着，为什么说未填值」。WI-0003 一张卡
    九十几行全印「未填值」，读的人只会得一个结论 —— 这表是空的；而事实是
    我们没核。措辞不改，等于把工具的盲区记在文件的账上。

    分类计数仍走 STATUS_CN（汇总里「未填值 N 条」的口径不变），
    只在**行级**把话说准 —— 两个通道各管一件事，互不牵连。
    """
    st = (chk or {}).get("status") or ""
    if st == "novalue" and (chk or {}).get("unreadable"):
        return "未参与核对（填的是文字/代码）"
    return STATUS_CN.get(st, st)


def ensure_snap_dir(base):
    d = os.path.join(base, SNAP_DIRNAME)
    os.makedirs(d, exist_ok=True)
    return d


# ------------------------------------------------------------------ 导出 xlsx

def export_sheets(sheets, out_path):
    """sheets: [(sheet名, [列名...], [[值...], ...]), ...]

    每张表都加自动筛选、冻结首行、按内容估算列宽 —— 打开就能筛，不用再整理。
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, columns, rows in sheets:
        ws = wb.create_sheet(title=str(name)[:31] or "Sheet")
        ws.append([str(c) for c in columns])
        for c in range(1, len(columns) + 1):
            cell = ws.cell(row=1, column=c)
            cell.font = Font(bold=True)
            cell.fill = _HDR_FILL
            cell.alignment = Alignment(vertical="center")
        for r in rows:
            ws.append(["" if v is None else v for v in r])
        ws.freeze_panes = "A2"
        if columns:
            ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(columns)),
                                              max(2, ws.max_row))
        # 列宽：按前 200 行内容估，中文按 2 倍宽算，限制在 8~46
        for i, col in enumerate(columns, 1):
            w = 0
            for r in list(rows)[:200]:
                if i <= len(r) and r[i - 1] is not None:
                    s = str(r[i - 1])
                    w = max(w, sum(2 if ord(ch) > 127 else 1 for ch in s))
            ws.column_dimensions[get_column_letter(i)].width = \
                min(46, max(9, len(str(col)) * 2 + 2, min(w + 2, 46)))
    wb.save(out_path)
    return out_path


# ------------------------------------------------------------------ 历史快照
# 2026-10-30 拆分第一步：整块搬到 wic/memory/snapshot.py（含 SNAP_DIRNAME）。
# 这里只做转出口，下面的代码仍然直接调 build_signature / save_snapshot …
# 界面走的是 engine.xxx，所以 app.py 一行都不用改。
from wic.memory.snapshot import (           # noqa: E402
    SNAP_DIRNAME, ensure_snap_dir, build_signature, save_snapshot,
    latest_snapshot, diff_snapshot)


# ==========================================================================
#  第 7 节  判定引擎
#  compare 分层收窄；收不窄就如实标注，不猜
# ==========================================================================

"""WI 一致性检查核心引擎：通用 PDF 参数抽取 + 类型分派核对。

设计（重构后）：
- 通用抽取：扫描所有页，找任何含「参数名称」的表，抽 {参数名: 值集合}。
- 名称规范化：中英文关键词，把 Excel 名和 PDF 名归一到标准键。
- 类型分派：浮点/整数/十进制/固定值 → 核对规格；字符串/布尔值 → 采集项跳过。

Excel 侧（2026-09-17 起）：
- 完全交给第 1 节的「结构自适应层」处理——不写死 sheet 名、不写死列号，
  按列名语义定位字段；认不出来直接抛错，不静默返回空。
"""


# 采集项名称的「名称+单位」拆分统一由第 1 节提供，避免两处实现漂移
_parse_item_name = parse_item_name   # （原 from schema import parse_item_name）
# 表格自检（纯 Excel 侧）：把「类型 / 是否必填 / 是否显示上下限 / 是否填写 N/A /
# 固定值 / 布尔值真 / 组序号 / 设备号」这些表里自带、却从未被任何校验用过的列
# 变成可执行规则。独立模块，不碰 PDF 抽取与比对链路 —— 改动风险为零。
_audit_xlsx = audit_xlsx   # （原 from excel_audit import audit_xlsx）
# 覆盖面与版本链：回答「这些文件凑一起有没有整块没被核过」——录入了数据的 WI
# 手上有没有对应 PDF、登记为「运行」的文件在不在、每行声明的 WI 版本对不对得上。
_audit_workbook = audit_workbook   # （原 from workbook_audit import audit_workbook）
# 同规格横向一致性：把「绝大多数产品都这么写、个别产品不一样」的格子捞出来 ——
# 逐行核 PDF 只能发现「与规定不符」，发现不了「和同类产品不一致」，而那正是漏改的样子。
_scan_outliers = scan_outliers   # （原 from outlier_scan import scan_outliers）
# PDF 版式自适应：竖表（参数名在列上）/ 横表（「参数 | Step N」型）/ 程序附表（转置双列）
# open_pdf：同一份 PDF 的十来趟提取共用一批页面对象（省时间也省内存，见第 2 节）


# ---------------------------------------------------------------- 基础工具

# 最近一次解析中失败的页码 [(页码, 原因)]：单页坏掉时如实报告，不假装整份都读到了
LAST_PDF_FAILED = []


def _note_pdf_failure(pgno, exc):
    LAST_PDF_FAILED.append((pgno, "%s: %s" % (type(exc).__name__, exc)))


def reset_pdf_failures():
    del LAST_PDF_FAILED[:]
    del LAST_UNPARSED[:]
    reset_layout_stats()      # 版式命中数 / 异常明细也是「每份 PDF 一份」


# ── 「未认领格」对账袋（R47 补）────────────────────────────────────────
# 这是**整个工具最要命的一个洞**的补丁，值得说清楚它补的是什么。
#
# 解析 PDF 时，「表里写着字、但我没能把它变成一个值」的格子，以前一律
# `continue` 掉 —— 于是「我读不懂」和「PDF 里压根没有这个参数」在下游是
# **同一件事**：都是「没有记录」。后果是判定层会理直气壮地写依据：
#       「参数名已收录，但这份 PDF 的参数表里没有它 —— 请人工核一眼」
# 而 PDF 里明明白纸黑字写着 `SL1+3 ref` / `中间模：135`。工具把「我没读懂」
# 说成了「文件里没有」—— 这是在**替人下结论**，而且下的是错的结论。
# 用户 2026-09-29 的原话：「你的判定机制去哪里了」。
#
# 有了这个袋子，两类事情就分开了：
#   · 表里没有这一格          → 真的没有，依据照旧
#   · 表里有字、我没认出来     → 进袋，依据要改成「PDF 写了 X，我没解出值」
# 以后再来什么千奇百怪的模板，只要表里出现我看不懂的字，就会进这个袋子、
# 被人看见，而不会被静默吞掉、伪装成「文件里没有」。
#
# 袋子里每条长这样（页面/表/子表/规格都带上，好让人一眼找到原文在哪）：
#   {"page":, "table":, "sub":, "title":, "param":, "step":, "col":,
#    "raw":, "reason":}
# reason 三种：`no_head`（这一列我认不出是什么参数）/ `no_param`（认不出
# 参数名）/ `no_value`（认得出参数名，但那格子里解不出数值）。
# ⚠️ `N/A` 不进袋 —— 那是表格在声明「此步不适用」，已有 `empty` 机制处理。
LAST_UNPARSED = []


def _unparsed_hits(er, index):
    """这一行 Excel 匹配不上时，PDF 侧有没有「写了字但我没读出值」的**同名**格子。

    ⚠️ 只按**规范名**精确命中，不做模糊匹配 —— 试过「同步骤 + 名字有公共片段」，
    实测 WI-0038 把「退出条件加热中间模/11」匹配到了水冷列的 `ON`（公共子串是
    「中间模」这种量纲词），**凭空造出一条假提示**。挂在单行上的依据一旦说错，
    人会照着去改一个本来就对的 Excel —— 那比不提示更糟。
    宽松的发现不挂单行，走 `pdf_read_audit()` 那份整体体检清单。
    """
    key = _canonical_param(er.get("item"))
    return (index.get(key) or []) if key else []


def _norm(s):
    return re.sub(r"\s+", "", _sym(s) or "")


def _clean(s):
    return re.sub(r"\s+", " ", _sym(s) or "").strip()


# ----------------------------------------------------------------- 值解析

# 「Table 1」「表2」这类是**指向另一张表的引用**，不是参数值。
# 不剥掉的话『Pleat Head Open Stop Micrometer Setting / Table 1』会被解析成 1.0，
# 凭空给「千分尺设定值」添一个 1.0 的候选。
_REF_TEXT = re.compile(r"(?i)\b(?:table|tab|附录|附表)\s*[0-9]+(?:[-\u2010-\u2015][0-9A-Za-z])?")


def _strip_ref(text):
    return _REF_TEXT.sub(" ", str(text or ""))


def _bound_side(text):
    """值文本里**不等号的方向**：'>' 系 → 'lo'（下限约束），'<' 系 → 'hi'，认不出 → None。

    2026-09-24 起因（用户点破的真缺陷）：WI-0039 的 `> 0.52Mpa` 是**下限**约束
    （必须大于 0.52，填 0.6 / 0.7 / 1.0 都行）。原先 `>` 被解析器丢掉、
    读成精确值 (0.52, 0.52) —— 于是「Excel 把 0.52 填在**上限**列、还勾了允许超」
    这种填错，也一路被判成「一致」。把方向留下来，判定层才能分清
    「该填下限却填了上限」。

    判据：符号或「大于/小于」后面**必须紧跟数字**才认 —— 说明文字里
    孤零零的「大于」不算。中文说法先归一（「不小于」含着「小于」、
    「不大于」含着「大于」，直接搜会两个都中，所以先替换成符号再统一判）。
    全语料实测：10 份 PDF 的取值格里**没有一条**带不等号 ——
    这条只影响 WI-0039 这类新样本，零回归哈希不动。
    """
    t = str(text or "")
    t = re.sub(r"大于等于|≥|≧", ">", t)
    t = re.sub(r"小于等于|≤|≦", "<", t)
    t = re.sub(r"不小于|不低于|不少于", ">", t)
    t = re.sub(r"不大于|不高于|不超过|不多于", "<", t)
    gt = re.search(r"(?:[>＞]|大于)\s*[0-9.]", t)
    lt = re.search(r"(?:[<＜]|小于)\s*[0-9.]", t)
    if gt and not lt:
        return "lo"
    if lt and not gt:
        return "hi"
    return None


def _raw_num_dps(raw):
    """从 PDF 取值格原文里，按出现顺序提取每个数字的 (数值, 小数位数)。

    用于「3.0 与 3 不能算对」的精度守卫：期望值记录里 `lo`/`hi` 已转成 float
    （精度丢失），但 `raw` 原文还在 —— 从这里把每个数的精度找回来。

    例：`'95-115'` → [(95,0),(115,0)]；`'3.0~5.0'` → [(3.0,1),(5.0,1)]；
        `'0.2-15'` → [(0.2,1),(15,0)]；`'>95℃'` → [(95,0)]；`'3s'` → [(3,0)]。
    只认「前面不是字母/数字」的数值串（与 _extract_all_values 的负号锁定同源），
    避免把设备号 `T-1281` 里的 1281 当成参数值精度。
    """
    out = []
    for m in re.finditer(r"(?<![A-Za-z0-9])-?\d+(?:\.\d+)?", str(raw or "")):
        tok = m.group(0)
        dot = tok.find(".")
        dp = 0 if dot < 0 else len(tok) - dot - 1
        out.append((float(tok), dp))
    return out


def _rec_dp(r):
    """从一条 PDF 记录提取其取值精度签名 (lo_dp, hi_dp)。

    按「数值相等」把 raw 里的数字精度配对到 lo / hi 上；**只要有一端配对不上
    就整条返回 None**（该值精度未知 → 不参与精度校验，退回只看数值的旧口径，
    绝不因此误判）。

    为什么必须「全配齐才算数」：公差展开（`160+5/-0` → (160,165)）时，hi=165
    在 raw 里**根本不存在**（是 160+5 算出来的），只配到 lo 就返回 (0, None)，
    会让「160~165 vs 160~165」这种完全一致的写法被误判成精度不一致。
    """
    lo, hi = r.get("lo"), r.get("hi")
    if lo is None:
        return None
    dps = _raw_num_dps(r.get("raw"))
    lo_dp = None
    for val, dp in dps:
        if abs(val - lo) < 1e-9:
            lo_dp = dp
            break
    if lo_dp is None:
        return None
    if hi is None or abs(hi - lo) < 1e-9:
        return (lo_dp, lo_dp)
    hi_dp = None
    for val, dp in dps:
        if abs(val - hi) < 1e-9:
            hi_dp = dp
            break
    if hi_dp is None:
        return None
    return (lo_dp, hi_dp)


def _rec_dps(r):
    """竖表记录的精度签名：{(lo, hi): (lo_dp, hi_dp)}，一条记录可能有多个值。

    ⚠️ 不能直接用 `_rec_dp`：它读的是记录上的 `lo` / `hi` 字段，那是**横表**
    记录才有的；竖表把值放在 `vals` 集合里（一格可能装着「位置:105mm；速度:50mm/s」
    好几个量），记录级根本没有 lo/hi。拿 `_rec_dp` 去量竖表，结果恒为 None ——
    于是「3.0 与 3 不能算对」这条规矩在竖表路径上从来没生效过
    （用户 2026-10-13 反馈「小数点后的数值不能忽略，需要书写一致进行比对」）。
    """
    out = {}
    raw = r.get("raw")
    dps = _raw_num_dps(raw)
    if not dps:
        return out
    for (lo, hi) in (r.get("vals") or ()):
        if lo is None:
            continue
        lo_dp = next((dp for val, dp in dps if abs(val - lo) < 1e-9), None)
        if lo_dp is None:
            continue
        if hi is None or abs(hi - lo) < 1e-9:
            out[(lo, hi)] = (lo_dp, lo_dp)
            continue
        hi_dp = next((dp for val, dp in dps if abs(val - hi) < 1e-9), None)
        if hi_dp is not None:
            out[(lo, hi)] = (lo_dp, hi_dp)
    return out


def _extract_all_values(text):
    """从值文本抽**所有** (lo, hi, unit) —— 竖表 / 程序附表的「候选值」解析器。

    ⚠️ 与 `parse_step_value`（精确路径的单值解析器）**不是重复实现，
    不要合并**。两者粒度不同：本函数把一格里的每个数都收进候选集（一格可能同时
    写着 `位置:105mm; 速度:50mm/s`，或 `100-120s` 与 `200-240s` 两个档），随后由
    「设备号 / 程序号 / 步骤号」这类适用范围去收窄；parse_step_value 则直接给出该格
    的唯一值。

    2026-09-22 实测（只读分歧审计 + 全量回归模拟）：
      · 10 份 PDF、266 个不同取值格，142 处两套结果不同，**全部**是"本函数多收
        候选、单值那套给出的值必在候选里"——没有任何一处数字读错 / 符号读反 /
        一边有值一边无值。所以二者是粒度差异，不是对错差异。
      · 把本函数改成"只取单值"再跑全量：9 个 WI 判定不变，WI-3852 有 2 条
        （表4 亲水和亲水后正压测漏-0004 / -0005「表层溶液步骤19固化时间」）
        从 ok 变成 mismatch —— 期望值由正确的 200-240 退化成格内第一个数 100-120。
      结论：合并会**引入假不一致**，两套并存是必需的。

    支持范围(240-250℃、16s-20s)、单值(70℃)、
    复合(位置:Xmm;速度:Ymm/s)、乘号(「100-120s）x2」→ 200-240)、
    公差(「14+0.5/-0 atm」→ 14~14.5)。

    2026-09-21：补上**正负号**。原先三个正则都不认符号，`-0.8mm` 会被读成 0.8 ——
    而「移动距离 / 位置 / 上提下移」这类参数经常是负值，认错符号比认错数字更严重
    （对的说成错）。符号用「前面不是字母数字」锁定，这样 `T-1362` 里的连字符
    不会被当成负号（设备号仍按 1362 处理，与原先一致）。
    """
    text = _strip_ref(text)
    result = []
    num = r"(?<![A-Za-z0-9])-?\d+(?:\.\d+)?"
    # -1) 负号作用于整个括号区间：`-(0.9~1.3)mm` 意思是 **-1.3 ~ -0.9** ——
    #     负号吃掉区间、两端**取负并交换次序**。以前这条没写，落到下面「范围」规则
    #     时负号被当成连字符丢掉，读成 0.9~1.3 —— **符号整个反过来**。
    #     这是继 `> 0.52`（第九轮）之后第二个「符号读错」：数字看着对、判定全是反的，
    #     比读不出数更危险，且界面上完全看不出来。
    #     右括号写成可选：PDF 换行会把 `-(0.9~1.3)mm(forT-1112/T-1114)` 撕成两半，
    #     到这儿只剩 `-(0.9~1.3)mm(`，右括号没了。
    #     先量后改：现有 10 份语料里这种写法出现 **0 次**，所以是纯增量。
    for m in re.finditer(r"-\s*[（(]\s*(\d+(?:\.\d+)?)\s*[~～\-]\s*(\d+(?:\.\d+)?)\s*[)）]?\s*([a-zA-Z/%²°℃]*)", text):
        a, b = float(m.group(1)), float(m.group(2))
        result.append((-b, -a, m.group(3) or ""))
    text = re.sub(r"-\s*[（(]\s*\d+(?:\.\d+)?\s*[~～\-]\s*\d+(?:\.\d+)?\s*[)）]?\s*[a-zA-Z/%²°℃]*", " ", text)
    # 0.5) 正负号公差：`3.0±0.5` 表示 2.5~3.5，`3.0+/-0.5` 同理。
    #      ⚠️ 必须排在「范围」「单值」之前，并把匹配到的片段从 text 里剔掉。
    #      不这么做的话，单值规则会把 3.0 和 0.5 各收一条，期望值集合里就多出
    #      两个「单点」—— Excel 填 3.0 判一致（假一致）、填 2.8 判不一致（假差异），
    #      而 PDF 真正表达的是「2.5 到 3.5 都能接受」。
    #      用户 2026-10-14 反馈「无法识别 WI 中特殊符号±参数进行换算比对」。
    #      `±0.5` 单独出现（不带基准）时按 -0.5~0.5 理解，与 `parse_step_value` 同口径。
    for m in re.finditer(r"(" + num + r")\s*±\s*(\d+(?:\.\d+)?)\s*([a-zA-Z/%²°℃]*)", text):
        base, tol = float(m.group(1)), float(m.group(2))
        result.append((base - tol, base + tol, m.group(3) or ""))
    text = re.sub(r"(?<![A-Za-z0-9])-?\d+(?:\.\d+)?\s*±\s*\d+(?:\.\d+)?\s*[a-zA-Z/%²°℃]*", " ", text)
    for m in re.finditer(r"(" + num + r")\s*\+\s*/\s*-\s*(\d+(?:\.\d+)?)\s*([a-zA-Z/%²°℃]*)", text):
        base, tol = float(m.group(1)), float(m.group(2))
        result.append((base - tol, base + tol, m.group(3) or ""))
    text = re.sub(r"(?<![A-Za-z0-9])-?\d+(?:\.\d+)?\s*\+\s*/\s*-\s*\d+(?:\.\d+)?\s*[a-zA-Z/%²°℃]*", " ", text)
    for m in re.finditer(r"±\s*(\d+(?:\.\d+)?)\s*([a-zA-Z/%²°℃]*)", text):
        tol = float(m.group(1))
        result.append((-tol, tol, m.group(2) or ""))
    text = re.sub(r"±\s*\d+(?:\.\d+)?\s*[a-zA-Z/%²°℃]*", " ", text)
    # 0) 公差：基准 +上偏差/-下偏差，如 "14+0.5/-0" → (14-0, 14+0.5)
    for m in re.finditer(r"(\d+(?:\.\d+)?)\s*\+\s*(\d+(?:\.\d+)?)\s*/\s*-\s*(\d+(?:\.\d+)?)\s*([a-zA-Z/%²°℃]*)", text):
        base, up, low = float(m.group(1)), float(m.group(2)), float(m.group(3))
        result.append((base - low, base + up, m.group(4) or ""))
    text = re.sub(r"\d+(?:\.\d+)?\s*\+\s*\d+(?:\.\d+)?\s*/\s*-\s*\d+(?:\.\d+)?\s*[a-zA-Z/%²°℃]*", " ", text)
    # 1) 乘号：范围 × 倍数
    for m in re.finditer(r"(\d+(?:\.\d+)?)\s*[-~～]\s*(\d+(?:\.\d+)?)\s*([a-zA-Z/%²°℃]*)[^\d]*?[x×X]\s*(\d+(?:\.\d+)?)", text):
        mult = float(m.group(4))
        result.append((float(m.group(1)) * mult, float(m.group(2)) * mult, m.group(3) or ""))
    text = re.sub(r"\d+(?:\.\d+)?\s*[-~～]\s*\d+(?:\.\d+)?\s*[a-zA-Z/%²°℃]*[^\d]*?[x×X]\s*\d+(?:\.\d+)?", " ", text)
    # 2) 范围（容忍单位夹在中间，如 "16s-20s"，也容忍负号）
    for m in re.finditer(r"(" + num + r")\s*([a-zA-Z/%²°℃]*)\s*[-~～]\s*(" + num + r")\s*([a-zA-Z/%²°℃]*)", text):
        unit = m.group(4) or m.group(2) or ""
        result.append((float(m.group(1)), float(m.group(3)), unit))
    text = re.sub(r"(" + num + r")\s*[a-zA-Z/%²°℃]*\s*[-~～]\s*(" + num + r")\s*[a-zA-Z/%²°℃]*", " ", text)
    # 3) 剩余单值（同样保留符号）
    for m in re.finditer(r"(" + num + r")\s*([a-zA-Z/%²°℃]*)", text):
        result.append((float(m.group(1)), float(m.group(1)), m.group(2) or ""))
    # ⚠️ 出口统一规范化成「下限 ≤ 上限」。PDF 会把**负值区间从大到小写**：
    #    `-0.8~-1.2mm`（本意是 -0.8 到 -1.2 这一段）。上面的范围规则按「先出现的是
    #    下限」取，于是存成 (-0.8, -1.2) —— 下限比上限还大。而 Excel 侧的
    #    「下限 / 上限」两列永远小在前，是 (-1.2, -0.8)。
    #    两边形状一样、两个数一个不差，却因为顺序不同在集合里对不上 →
    #    判成「不一致」，而界面上「实际值 -0.8~-1.2」和「期望值 -1.2~-0.8」
    #    在**人眼里完全相同**，复核的人根本没法判断该改哪一边。
    #    实测（WI-0013/2025 近端/远端焊接「移动距离」）：65 条数值差异里 49 条是这一类。
    #    这不是某个 WI 的特例 ——「区间的小端写在前面」是区间本身的定义，
    #    与怎么排版无关，所以归一化放在这个唯一出口，而不是逐处打补丁。
    #
    # ⚠️ 同时削掉二进制浮点尾巴（2026-10-29）：
    #    `1.35±0.15` 算出来是 1.2000000000000002，而 Excel 写的是 1.2 ——
    #    元组直接比永远不相等，**把对的说成错的**（假差异）。
    #    WI-3140 的 20 条「速度 mm/s」全栽在这上面：PDF 写 1.35±0.15、
    #    Excel 填 1.2~1.5，两边语义完全一致，却因为末尾那个 2e-16 判成不一致。
    #    工程参数用不到第 15 位小数，round 到 9 位既吃掉尾巴、又不碰真实精度。
    return [(round(min(a, b), 9), round(max(a, b), 9), u) for (a, b, u) in result]


def _norm_unit(u):
    """单位归一化：统一大小写、上标、常见别名（预设 Excel 里出现的全部单位）。

    ⚠️ 2026-10-30 试过、又撤掉的一改：在这里把「整串被括号包住」的单位脱掉括号
    （`（Inch）` → `inch`），让单位收窄层真正生效。实测**不能做** ——
    WI-0014 的 PDF 文字层把 `430-520mW` 印成了 `430- 520m w`，单位被切成 `m`；
    Excel 那边写 `（mW）`，脱掉括号后单位层判它「单位不符」把这条**正确**的记录
    筛掉，`一致` 当场变成 `不一致`（DWG-0000005 实测 1 条）。
    括号还在的时候单位层恒不相等、等于失效，那条错版记录侥幸留下来，判定是对的。
    大样本 14490 行脱不脱括号**完全一致**，收益几乎为零 —— 为了零收益去换一条
    假差异不值得。真正该修的是 PDF 那段 `m w` 的单位切分（见 `使用说明` 待办）。
    """
    u = (u or "").strip().lower().replace("²", "2").replace("³", "3")
    if u in ("s", "sec", "secs", "second", "seconds", "秒"):
        return "s"
    # `oc` / `ºc` 是 PDF 文本层里最常见的「摄氏度」写法 —— 排版时把度符号 °
    # 打成了字母 o（WI-0016 正文「温度75oC」）或阳性序数符 º（参数表「90~120 ºC」）。
    # 不归一的话，单位层会把这条**唯一标了单位**的记录判成「单位不符」而丢掉，
    # 于是正文里明明白白写着的 75℃ 变成「PDF 里没有该设备对应的参数」。
    if u in ("℃", "°c", "oc", "ºc", "c", "°"):
        return "℃"
    if u in ("psi", "psig", "psia"):
        return "psi"
    if u in ("inch", "in", "”", '"', "英寸"):
        return "inch"
    if u == "w/cm2":
        return "w/cm2"
    if u in ("l/min", "lpm"):
        return "l/min"
    return u  # mm/s、mm、atm、mpa 等本身已规范，直接返回


def _excel_unit(er):
    """Excel 这一行**拿去对齐的单位** —— `ref` 这类标记不算单位。

    采集项名写 `左拉伸距离/7(mm):ref` 时，`parse_item_name` 会把尾巴上的 `ref`
    当成单位（旧口径如此，`pdf_value_shape` 也靠 `_has_ref(er["unit"])` 认 ref）。
    但拿这个 'ref' 去跟 PDF 的 `mm` 对齐时，PDF 里那两条**有值**记录会被整片滤掉
    —— WI-0032 表4!185 实测：`_step_expected` 的单位层只剩 `blank_index` 里
    表1 Step7 的那条 `N/A`，于是 scope 变成 `pdf_blank`、依据写成
    「PDF 这一格本身就是空值（表格里写的是「------」）」，而 PDF 明明写着 `1-30`。
    **判定看着是「一致」，理由却是假的** —— 与用户点名的「没跟 PDF 匹配」同一个病。

    所以：ref 的**识别**仍走 `_has_ref(er["unit"])`（那里要的就是字面），
    这里只回答「拿什么单位去对齐」——标记一律当作没有单位。
    """
    u = _norm_unit(er.get("unit"))
    return "" if _has_ref(u) else u


# ------------------------------------------- 参数表抽取（按表头认表，不看页码）
#
# 这一段原来是「WI-0008 专用」：页码写死 PARAM_PAGES = [5,6,7,8,9]、取值列写死
# row[2]。它只对 WI-0008 本身有效 —— 换一份文档就会抽到完全无关的东西：
#   · WI-0007 的「序号|物料编号|物料说明|使用量」物料明细表被当成参数表，
#     抽出 7 条，参数名就是 DWG-0000008 这类物料编号；
#   · WI-0009 抽出 12 条，参数名是「移动产品到设定位置」这类操作步骤文字；
#   · WI-0023 抽出 7 条「焊接时间」（来自第 5~8 页的正文表）。
# 这些垃圾此前只是恰好被「P/N 为空」挡在判定之外；换个文档一旦 P/N 对上，
# 就会给出错误的期望值 —— 这正是「对老样本调好了、换个新文件就错」的隐患。
#
# 现在改为靠**表头语义**认表，彻底不看页码：
#   参数表的表头是「序号 | 参数名称 | 参数值或范围」，其中「参数值或范围」
#   是参数表独有的列名 —— 物料明细表是「物料编号/物料说明/使用量」、
#   流程页是「操作/说明」，都过不了这一关。取值也按该表头所在列定位，
#   不再写死第 3 列，所以列数不同、竖排横排的参数表都能读。

# 原 `UNIT_MAP`（把 °C/°c 归一成 ℃ 的旧写法）已于 2026-09-30 删除 ——
# 全项目零引用，单位归一的唯一入口是第 7.5 节的 `UNIT_TYPES` / `_norm_unit_key`。
# 留一份被取代的老常量在文件里，只会让人以为还能改它，遂删。
PRECISE_VALUE_HEADERS = ("参数值或范围", "参数值", "设定值", "设置值")
PARAM_KEYS = ("温度", "时间", "风量", "气压")


def _parse_value(v):
    """'240-250℃' → (240.0, 250.0, '℃', '')；'0.2MPa（ref）' → (0.2,0.2,'MPa','ref')。

    2026-09-21 改：直接调 `parse_step_value` —— 全工具只留**一套**取值解析。
    原先这里自己实现，只会「取前两个数字」，实测有三类真错误（都是用户点出来的
    「数字识别不准 / 符号识别不准 / 没做加减法」）：

      · `-0.8mm` —— 负号被吃掉读成 0.8。而「移动距离 / 位置」这类参数经常是负值，
        符号丢了等于把对的说成错。
      · `T-1281：-1.3mm` —— 把设备号读成值（得到 1281）。
      · `14+0.5/-0 atm` —— 只取到 14 和 0.5，不做加减法（正确是 14.0~14.5）。

    横表那一路早就会处理这几种。两套解析器并存，等于「同一格在不同链路上结论不同」，
    所以统一收口到能力更强的那一套，而不是两边各修各的。
    """
    pv = parse_step_value(v)
    if pv is None:
        return None, None, "", ""
    return pv


def _parse_spec(s):
    s = _clean(s)
    s = re.sub(r"^参\s*数\s*名\s*称", "", s)
    return s.strip("():：（） ")


def _parse_equip(s):
    s = _norm(s)
    s = re.sub(r"^.*?编号", "", s)
    s = re.split(r"设备|说明", s)[0]
    return s.lstrip("：:")


def _page_pn_core(page_text):
    cores = set()
    for pn in re.findall(r"DWG[-–]?\d+[-/]\d+[,\d]*", page_text or ""):
        m = re.search(r"(\d+)", pn)
        if m:
            cores.add(m.group(1))
    return cores.pop() if len(cores) == 1 else None


def extract_pdf_params_precise(pdf_path):
    """抽取「序号 | 参数名称 | 参数值或范围」参数表。

    返回 [{param,lo,hi,unit,equip,spec,pn,page,tno,row,col}]。
    页码不写死：逐页找「有参数表表头」的表格，找不到就不抽（见上方说明）。

    2026-11-09 加 tno/row/col 并按物理格去重：
      一行参数在 PDF 里的**身份**就是「第几页、第几张表、第几行、第几列」。
      以前不记这四个数，下游只能拿 (param, lo, hi) 认人 —— 而同一页里
      「风量 6-8 L/min」本来就会在好几个小节重复出现（那是不同规格档，
      **不是重复抽取**），用值认人会把真信息当成噪声。
      记坐标之后，「同一个物理格被两张嵌套表各读一遍」才认得出、才敢删。
    """
    records = []
    seen = set()
    with open_pdf(pdf_path) as pdf:
        for pgno, page in enumerate(pdf.pages, 1):
            got = []
            try:
                _extract_precise_page(page, pgno, got)
            except Exception as e:          # 单页失败不影响其它页
                _note_pdf_failure(pgno, e)
                continue
            for r in got:
                # ⚠️ 键里必须有 page：只写 (tno,row,col) 会把「第 6 页第 1 张表
                # 第 2 行第 2 列」当成「第 5 页同一个位置」，整页整页地误删
                # （第一次实现漏了它，WI-0008 的 88 条直接掉到 38 条）。
                k = (r.get("page"), r.get("tno"), r.get("row"), r.get("col"))
                if None not in k:
                    if k in seen:
                        continue
                    seen.add(k)
                records.append(r)
    return records


def _extract_precise_page(page, pg, records):
    """只处理真正的参数表：表里必须出现「参数值/设定值」这类表头列。

    找不到该表头列 → 整表跳过。物料明细表（物料编号/物料说明/使用量）与
    流程页表格就是这样被挡掉的，不需要知道它在第几页。

    每条记录带上 (tno,row,col) —— 那一格的坐标，供快照对拍、界面高亮用。
    """
    page_pn = _page_pn_core(page.extract_text())
    for tno, t in enumerate(page.find_tables()):
        rows = t.extract()
        if not rows:
            continue
        vcol = None
        for r in rows:                      # 定位「参数值」所在列
            for ci, c in enumerate(r):
                if any(h in _norm(c) for h in PRECISE_VALUE_HEADERS):
                    vcol = ci
                    break
            if vcol is not None:
                break
        if vcol is None:                    # 不是参数表
            continue
        equip = None
        spec = None
        for ri, row in enumerate(rows):
            if not row:
                continue
            c0 = _clean(row[0]) if row[0] else ""
            n0 = _norm(c0)
            joined = " ".join(_clean(c) for c in row if c)
            if any(k in joined for k in ("工序名称", "文件编号", "页码")):
                continue
            if n0.startswith("设备"):
                equip = _parse_equip(c0)
                continue
            if n0.startswith("序号") or n0.startswith("参数名称"):
                spec = _parse_spec(c0 if n0.startswith("参数名称") else (row[1] if len(row) > 1 else ""))
                continue
            if re.match(r"^[1-9]$", n0) and len(row) > vcol:
                name_raw, value_raw = _clean(row[vcol - 1]), row[vcol]
            elif len(n0) < 25 and any(k in n0 for k in PARAM_KEYS) and len(row) > vcol:
                name_raw, value_raw = _clean(row[0]), row[vcol]
            else:
                continue
            zh = re.search(r"[一-鿿]+", name_raw)
            param = zh.group(0) if zh else _clean(name_raw)
            if not param:               # 参数名为空的行不是参数（如 WI-0011 正文表）
                continue
            lo, hi, unit, note = _parse_value(value_raw)
            if lo is None:
                continue
            records.append({"page": pg, "equip": equip, "spec": spec,
                            "param": param, "lo": lo, "hi": hi, "unit": unit, "pn": page_pn,
                            "tno": tno, "row": ri, "col": vcol})


def _pn_core(equip_str):
    """'T-2068-DWG-0000007-03,04/...' → '2000022'；无 P/N 返回 None。"""
    cores = set()
    for pn in re.findall(r"T-\d+-(DWG[-\d,./]+)", equip_str or ""):
        m = re.search(r"(\d+)", pn)
        if m:
            cores.add(m.group(1))
    return cores.pop() if len(cores) == 1 else None


def _size(spec):
    m = re.search(r"(\d+(?:\.\d+)?)\s*[-~～]\s*(\d+(?:\.\d+)?)", spec or "")
    return (float(m.group(1)), float(m.group(2))) if m else None


def _size_set(spec):
    return {(float(a), float(b)) for a, b in
            re.findall(r"(\d+(?:\.\d+)?)\s*[-~～]\s*(\d+(?:\.\d+)?)\s*mm", spec or "")}


def _equip_set(equip_str):
    out = set()
    # 「T- 10455」——T- 和数字之间夹个空格（DWG-0000003 卡0006 实测），老正则
    # 要求 T- 紧跟数字，于是**整组里的第一台**被无声丢掉，剩下两台的交集
    # 往往就空了。先把空格吃掉再解析。
    s = re.sub(r"(T-)\s+", r"\1", str(equip_str or ""))
    for grp in re.findall(r"T-(\d+(?:/\d+)*)", s):
        for num in grp.split("/"):
            out.add(f"T-{num}")
    return frozenset(out)


def _expected_values(excel_rec, pdf_precise):
    """WI-0008 精确匹配：按 P/N + 尺寸 + 设备对齐，返回期望值集合；无对应返回 None。"""
    pn = _pn_core(excel_rec["equip"])
    size = _size(excel_rec["spec"])
    equip = _equip_set(excel_rec["equip"])
    candidates = []
    for p in pdf_precise:
        if p["param"] != excel_rec["item"]:
            continue
        if pn is not None:
            if p.get("pn") != pn:
                continue
            if size is not None and p.get("sizes") and size not in p["sizes"]:
                continue
            if not (_equip_set(p["equip"]) & equip):
                continue
        else:
            if _equip_set(p["equip"]) != equip:
                continue
        candidates.append(p)
    if not candidates:
        return None
    return {(p["lo"], p["hi"]) for p in candidates if p["lo"] is not None}


# ---------------------------------------------------------------- 通用抽取

# --------------------------------------------------------------------------
#  竖表补维度（2026-09-29 第 42 轮）：让竖表也带上横表那套「可收窄的坐标」
#
#  横表（步骤式）每条记录天生带着 sub / dia / equip / step，所以能一层层收窄到唯一值；
#  竖表老路径只带 param 与 equip —— 于是同一台机下**所有并列子表**的值被并成一个
#  大集合。实测 WI-0031：应为值 24~33 个（7 个程序 × 两侧 × 位置/速度），
#  WI-0009 / WI-3852 同样是这个版式（83 / 20 条复合格）。**假一致就藏在这里**：
#  位置那格填成速度的值，照样判一致。
#
#  补三样：
#    · attr   一格装了两种量（「位置: 1400 mm；速度: 50 mm/s」）→ 拆成两条
#    · ord    同表内同参数同属性出现多次 = 并列的几列（底层 / 表层）→ 标序号
#    · prog   表头元信息（设备编号 / 涂层程序名称 / 适用导丝规格）下发给表内每一行
#
#  ⚠️ 三条「认不出就不动」的规矩（宁可放过，也不凭空造结构）：
#    · 只有一种属性 → 不拆（那本来就是单值）
#    · 同一属性出现两次 → 不拆（那是别的结构，不是「一格两量」）
#    · 表里没有元信息行 → 不下发（老版式零影响）
# --------------------------------------------------------------------------

# 「位置 / 速度」这类属性标签。只收**成对出现**时才有意义的那些，别把
# 「浸没时间 10Sec」这种单值也卷进来。
_V_ATTR_RE = re.compile(r"(位置|速度|距离|转速|次数)\s*[:：]")

# 表头元信息行：参数名 → (下发到哪个字段, 要不要原样留着当参数)
_V_META_ROWS = {
    "涂层程序名称": "prog",
    "程序名称": "prog",
    "程序号": "prog",
    "涂层程序": "prog",
    "设备编号": "equip",
    "机器编号": "equip",
}


def _cycle_len(names):
    """命中序列是不是「一整块属性周期重复」；是则返回块长，否则 None。

    为什么需要它（2026-10-29，WI-0042）：PDF 一张附表把**程序A / 程序B 两列**
    并排写在同一格里，文本于是长成
        `位置:950; 速度: 20mm/s; ; 位置:945; 速度: 20mm/s;`
    属性序列是 `位置,速度,位置,速度`。老规则「同一属性出现两次 → 不拆」把这种
    情况也拒了，整格退化成一条记录，位置 950、945、速度 20 全混进同一个值集合 ——
    于是「位置填成速度的值」照样判一致（假一致），而 Excel 侧按侧标一收窄就
    变成「多套候选」。

    判据取最严的一条：**整块周期重复，且块内属性互不相同**。
    `位置,位置`（块长 1）仍然拒绝 —— 那确实不是「一格两量」。
    """
    n = len(names)
    for k in range(2, n // 2 + 1):
        if n % k:
            continue
        if len(set(names[:k])) != k:
            continue
        if all(names[i] == names[i % k] for i in range(n)):
            return k
    return None


def _split_vertical_attrs(text):
    """把「位置: 1400 mm；速度: 50 mm/s」这种**一格装了两种量**的文本切开。

    返回 [(属性名, 该属性的原文片段)]；认不出就返回 []（调用方保持原样，绝不猜）。

    并列两列（程序A/程序B）时切出来是 4 段：位置/速度/位置/速度 ——
    `_vertical_expand` 的「并列列序」按 (参数,属性) 出现顺序编号，
    自然得到 位置 ord0=程序A、ord1=程序B，与 Excel 的「底层/表层」对得上。
    """
    t = str(text or "")
    hits = list(_V_ATTR_RE.finditer(t))
    if len(hits) < 2:
        return []
    names = [h.group(1) for h in hits]
    if len(set(names)) != len(names) and _cycle_len(names) is None:
        return []                      # 同一属性出现两次、又不是整块周期 → 不是「一格两量」
    out = []
    for i, h in enumerate(hits):
        end = hits[i + 1].start() if i + 1 < len(hits) else len(t)
        seg = t[h.start():end].strip().rstrip("；;,，、 ")
        if seg:
            out.append((h.group(1), seg))
    return out


def _split_value_cells(text):
    """把 `_v_data_row` 拼起来的一行（多个值格用 ' ; ' 连接）拆回每一格。

    为什么可以按 ' ; ' 拆：那个分隔符**只出现在格与格之间** —— 它是
    `" ; ".join(vals)` 插进去的，格内自己的分号后面紧跟内容、前面不会有空格。
    所以「按 ' ; ' 切」等于还原格边界，不是在猜。

    只有一格时原样返回（绝大多数竖表都是这样，行为与改动前逐字一致）。
    """
    t = str(text or "")
    if " ; " not in t:
        return [t]
    parts = [p.strip().strip(";；").strip() for p in t.split(" ; ")]
    parts = [p for p in parts if p]
    return parts or [t]


def _vertical_expand(recs):
    """补维度：拆值格 → 拆格内两量 → 标并列列序 → 下发表头元信息。"""
    # ① 拆「一格两量」
    #
    # 2026-10-29 改：先按值格切开，再拆格内的两种量。
    # 之前一步到位拆属性，遇到「程序A / 程序B 并排」就整块放弃（同一属性出现
    # 两次），整行退化成一条记录 —— 位置 950、945、速度 20 全并进一个集合。
    #
    # ⚠️ 切开本身**不改变结论**：没有并列侧标时，各格的值仍会并进同一个候选集。
    #    只有当 Excel 那边认出了「底涂 / 表涂」这类侧标，ord 才会真正用来收窄 ——
    #    而那正是以前会「假一致」的地方（拿程序B的值去核程序A）。
    out = []
    for r in (recs or []):
        for col, seg0 in enumerate(_split_value_cells(r.get("text"))):
            parts = _split_vertical_attrs(seg0)
            if not parts:
                n = dict(r)
                n["text"] = seg0
                n["raw"] = seg0
                n["col"] = col
                out.append(n)
                continue
            for attr, seg in parts:
                n = dict(r)
                n["attr"] = attr
                n["text"] = seg
                n["raw"] = seg
                n["col"] = col
                out.append(n)

    # ② 并列列序（底层 / 表层…）：同表内同参数同属性，按出现先后编号
    seen = {}
    for r in out:
        k = (r.get("page"), r.get("table"), r.get("param"), r.get("attr"))
        n = seen.get(k, 0)
        seen[k] = n + 1
        r["ord"] = n

    # ③ 表头元信息下发：同一张表（page+table）内的「涂层程序名称 / 设备编号」
    #    只写在某一行上，表内其它参数行得靠它才能收窄 —— 不下发就只剩设备号一个维度。
    meta = {}
    for r in out:
        fld = _V_META_ROWS.get(str(r.get("param") or "").strip())
        if not fld:
            continue
        v = str(r.get("text") or "").strip()
        if not v:
            continue
        meta.setdefault((r.get("page"), r.get("table")), {}).setdefault(fld, set()).add(v)
    for r in out:
        m = meta.get((r.get("page"), r.get("table")))
        if not m:
            continue
        if "prog" in m and not r.get("prog"):
            # 一张表一个程序名：取唯一那个，多值才并（并了收窄链照样认得）
            ps = sorted(m["prog"])
            r["prog"] = ps[0] if len(ps) == 1 else "|".join(ps)
        if "equip" in m and not (r.get("equip") or []):
            eq = sorted({x for g in m["equip"] for x in equip_from_text(_flat(g))})
            if eq:
                r["equip"] = eq
                r["equip_raw"] = "/".join(eq)
    return out


def extract_pdf_vertical(pdf_path):
    """竖表参数 + 各自的适用范围（设备 / 规格 / 程序号 / 产品线）。

    返回 [{param, vals{(lo,hi,unit)}, equip, dia, prog, families, sub, basis, page}]

    这是「底层」那一环：横表和程序附表本来就带着适用范围（所以 WI-0021/2105/2013/3852
    能收到唯一值），只有竖表这条老路径把范围丢了、只留 {参数名: 全 PDF 同名值并集}。
    这里补齐后，所有版式都由同一条收窄链判定，不用再一个 WI 一个 WI 地打补丁。
    """
    out = []
    for r in _vertical_expand(_vertical_all(pdf_path)):
        vals, units = set(), set()
        for (lo, hi, u) in _extract_all_values(_strip_ref(r["text"])):
            vals.add((lo, hi))          # 判定只比 (下限, 上限)，单位单独收窄
            if u:
                units.add(_norm_unit(u))
        if not vals:
            continue
        r = dict(r)
        r["vals"] = vals
        r["units"] = units
        # 不等号方向（`> 0.52Mpa` 是下限约束）。**只在所有取值都是单点时才认定** ——
        # 若同一格里还抽出了区间，方向语义说不清，宁可不判、按老口径走。
        r["side"] = (_bound_side(_strip_ref(r["text"]))
                     if all(a == b for (a, b) in vals) else None)
        r["basis"] = _scope_label(r)
        out.append(r)
    return _attach_tnames(out, pdf_path)


def _attach_tnames(recs, pdf_path):
    """给每条竖表记录补上「它所在那张表印的名字」（`tname`）。

    为什么必须补（2026-11-13 用户原话「为什么人类的眼睛就可以识别，为什么你们
    就不能识别呢？规则不也就是表头或者说表外面附近的那些文字信息吗」）：
      表名**早就在** `pdf_tables.catalog()` 里读出来了 —— 第 4 页那五张表分别叫
      「标记环焊接参数表」「铂环焊接参数表」，可这条链只把 `sub="表6"` 传了下来，
      判定层手里只剩「第几张表 + 设备号」。于是 WI-0028 只能拿设备集合去猜：
      表6 是 T-1857/T-1922/T-2101，**严格包含**本卡的 T-2101/T-1857，一锁就锁错了。
      人眼不会犯这个错，不是因为眼睛更强，是因为人眼**先看名字**。

    没印表名的表（`name` 为空）就不写这个键 —— 下游一律当「不知道它叫什么」，
    与「没标设备一律保留」同一条规矩，绝不因为缺信息被排掉。
    """
    if not recs:
        return recs
    try:
        import pdf_tables
        cat = pdf_tables.catalog(pdf_path)
    except Exception:
        return recs
    by = {}
    for c in cat or ():
        if c.get("kind") == "header":
            continue
        nm = str(c.get("name") or "").strip()
        if nm:
            by[(c.get("page"), c.get("tno"))] = nm
    if not by:
        return recs
    for r in recs:
        nm = by.get((r.get("page"), r.get("table")))
        if nm:
            r["tname"] = nm
    return recs


def _scope_label(r):
    """"表1（T-1065/T-1084…）" 这样的出处标签，用来在界面上说清「这条值凭什么」。"""
    bits = []
    if r.get("sub"):
        bits.append(str(r["sub"]))
    if r.get("step") is not None:
        bits.append("步骤%d" % r["step"])
    eq = r.get("equip") or []
    if eq:
        bits.append(eq[0] + ("等%d台" % len(eq) if len(eq) > 1 else ""))
    if r.get("dia"):
        a, b = r["dia"]
        bits.append("%.2fmm" % a if a == b else "%.2f-%.2fmm" % (a, b))
    if r.get("prog"):
        bits.append("程序号%s" % r["prog"])
    if r.get("families"):
        bits.append("/".join(r["families"]))
    return "·".join(bits) if bits else "p%s" % r.get("page")


def extract_pdf_params(pdf_path: str) -> dict:
    """兼容视图：{参数名: {(lo,hi,unit)}}（不带范围的旧口径，仅供统计/展示）。"""
    by_name = {}
    for r in extract_pdf_vertical(pdf_path):
        by_name.setdefault(r["param"], set()).update(r["vals"])
    return by_name


# ------------------------------------------------- 产品家族（跨同名附表的唯一线索）

def doc_family(xlsx_path: str) -> set:
    """从工作簿判断产品家族（如 产品线A）。

    为什么需要：WI-0011 的「Fold 头闭口针规」在 表2-a（产品线A 线）是 0.026″、
    在 表2-b（产品线C 线）是 0.029″，而两段的设备清单互相重叠 ——
    不认产品线就只能在两个值之间含糊其辞。工作簿文件名与表内文字都写着产品线。
    只取每个表的前若干行，绝不整表扫描（几十万行的表不能为了这点信息全读一遍）。
    """
    fam = families_of(os.path.basename(xlsx_path or ""))
    if fam:
        return fam          # 文件名里已写明产品线（如「…装入盘管组件 产品线A,DHR-0001…」）
    try:
        wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    except Exception:
        return fam
    try:
        for sn in wb.sheetnames[:5]:
            fam |= families_of(sn)
            ws = wb[sn]
            for i, row in enumerate(ws.iter_rows(max_row=60, max_col=14, values_only=True), 1):
                if i > 60:
                    break
                for c in row:
                    if isinstance(c, str) and c:
                        fam |= families_of(c)
    finally:
        try:
            wb.close()
        except Exception:
            pass
    return fam


# ---------------------------------------------------------------- 名称规范化

# 英文「复合专用名」：必须两个关键词**同时**命中。
# 单独立成表是为了让「双语名回退」（见 `_canonical_param` 末尾）能复用同一份判据，
# 免得两处各写一遍、日后改一边漏一边。
_PROPER_EN_PAIRS = (
    ("vacuum", "delay", "vacuum_delay"),
    ("pleat", "temperature", "pleat_temperature"),
    ("fold", "temperature", "fold_temperature"),
    ("balloon", "pressure", "balloon_pressure"),
    ("pleat", "dwell", "pleat_dwell"),
    ("fold", "dwell", "fold_dwell"),
    ("fold", "pressure", "fold_pressure"),
)

# ⚠️ 这里**曾经**试过在 `_canonical_param` 里加一条「动宾短语 + 两位序号 → 用
# 短语+序号作键」的规则（下移浸没01 / 提拉上升01 / 产品上提01），实测**有害**，已撤：
# 老版式（WI-0009 程序附表、WI-0008、WI-0004）靠的正是「上提/提拉 → 速度」这条
# **动词归类桥**来接 Excel 与 PDF 的不同叫法。把键改成「产品上提01」之后，
# 程序附表那一层直接对不上、整行掉到竖表兜底路径（prog_strict → multi，
# 唯一值变成 14 个候选），WI-0008 的「时间」还从 ok 变成 mismatch。
# 键是**跨版式共享**的，不能只为了竖表去改它 ——
# 竖表要区分 01/02/03，改用收窄后的「参数名确认层」（见 `_vertical_expected`）。


# --------------------------------------------------------------------------
#  通用量纲 → 「大类 | 限定词」键（第 54 轮，2026-10-29）
#
#  病根：词典把「通用量纲词」直接当成键，于是任何带限定词的参数都被它吞掉。
#  WI-0003 实测（DWG-0000011 / WI-0003 rev22）：
#    · 「锻打压力」被收成 `压力`，与同一页 表2 的 `Pressure/压力 20psi`
#      （那是压模机的模具压力，设备 T-1708）**静默同桶**。后面按设备/图号收窄时，
#      带着设备的真记录反而被图号（DWG-0000004/2000321 ≠ 本卡 2000323）筛掉，
#      只剩那条没有设备号、没有图号的 `压力 20psi` 幸存 → 期望值 20~20，
#      与 Excel 填的 20~60 判「不一致」。**假不一致**，比漏报更伤。
#    · 「锻打频率」根本没进词典 → `name_unknown` → 界面上就是「无对应」。
#      PDF 里分明写着 `Forging frequency/锻打频率 15-35Hz`，两条全漏。
#
#  规则：通用量纲词只提供「大类」；参数名里**除它以外的其余实词**是限定词。
#      键 = `大类`，或 `大类|限定词`。
#  两侧（Excel 采集项 / PDF 参数名）都走同一个函数，所以只要两边字面一致就自动成对，
#  不需要为「锻打压力」单独写一条 if。
#
#  ⚠️ 只对**量纲词是后缀、且两边照抄同一个中文术语**的几类生效（压力/频率/距离）。
#     速度/温度/时间/位置**故意不动**：那几类里词典已经靠同义词合并过
#     （上提 / 提拉 / 上升 / 移动 → 速度），一刀切加限定词会把
#     「上提速度」与「提升速度」拆成两个键 —— 那是凭空造漏报。
# --------------------------------------------------------------------------
_QUAL_FILLER = ("的", "之", "值", "参数", "设定", "要求", "范围", "上下限", "标准")
_QUAL_ZH_RE = re.compile(r"[\u4e00-\u9fff]+")


def _qual_key(n, cls, words):
    """`锻打压力` → `压力|锻打`；`压力` → `压力`；`模具 锻打频率` → `频率|模具锻打`。"""
    s = str(n or "")
    for w in words:
        s = s.replace(w, " ")
    for w in _QUAL_FILLER:
        s = s.replace(w, " ")
    frags = [f for f in _QUAL_ZH_RE.findall(s) if len(f) >= 2]
    q = "".join(frags)
    return (cls + "|" + q) if q else cls


def _ck_alts(ck):
    """规范键的查找顺序：**先精确（大类|限定），再退回大类本身**。

    必须留退路：PDF 若只写「压力」而 Excel 写「锻打压力」，精确桶查不到时
    仍要能命中大类桶 —— 否则这一改就把原来的「能核对」变成「无对应」，
    属于拿新漏报换旧假错。宁可用宽一档的候选，也不能让一条本来核得了的行掉出去。
    """
    if not ck:
        return ()
    if "|" in ck:
        return (ck, ck.split("|", 1)[0])
    return (ck,)


# 字面键在 vindex 里的命名空间前缀。用同一个 dict 存，免得再给
# `_vertical_expected` 加一个参数（它被 compare 和测试多处调用）。
_LIT_PREFIX = "名:"


def _literal_key(name):
    """参数名的**字面键**：中文侧、去掉单位括号与标点、统一小写。

    与 `_canonical_param` 的分工，是这一轮定下来的：
      · 词典键回答「这是哪一类参数」—— 粗，会并（锻打压力↔压力）、也会漏（锻打频率）；
      · 字面键回答「两边写的是不是同一个词」—— 不认字就不认，但相同就是铁证。

    WI-0003 rev22 的教训：Excel 与 PDF **两边都写着「锻打频率」**，只因为词典
    里没有这一条，整项被判成「参数名未收录」，界面上就是「无对应」。
    参数名本来就是照同一份 WI 抄的 —— 字面相同是最硬的证据，不该先问词典。
    """
    s = str(name or "").strip()
    parts = [p.strip() for p in re.split(r"[/／]", s) if p.strip()]
    zh = [p for p in parts if re.search(r"[一-鿿]", p)]
    if len(parts) > 1 and zh:
        s = zh[-1]
    s = re.sub(r"[\(（\[【][^\)）\]】]*[\)）\]】]", " ", s)     # 单位括号
    s = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "", s)
    return s.lower()


@lru_cache(maxsize=8192)
def _canonical_param(name):
    """把参数名归一到标准键（中英文关键词）。识别不到返回 None。"""
    raw = str(name or "")
    n = raw.lower()
    # 双语参数名「Test Pressure/测试压力」：与 Excel 对齐的是中文一侧。
    # 若拿整串去匹配英文关键词，会把「Alarm Pressure Drop/报警线」「Test Pressure
    # hold time/保压时间」也认成「压力」，于是测试压力的期望值集合里混进了报警线
    # 的 0.2 —— 「假一致」就是这么来的。取中文一侧后：测试压力→压力，报警线→None。
    parts = [p.strip() for p in re.split(r"[/／]", raw) if p.strip()]
    zh_parts = [p for p in parts if re.search(r"[一-鿿]", p)]
    if len(parts) > 1 and zh_parts:
        n = zh_parts[-1].lower()
    # —— 设备名不是参数名，必须在物理量关键词**之前**挡掉 ——
    # 「激光焊接机」「拉伸机」「亲水涂层机」这类名字里含「激光」「拉伸」等物理量
    # 关键词，不先挡掉就会被当成参数收进 vindex。WI-0004 实测代价：
    # 「激光焊接机 | T-1011」被 `"激光" in n` 归到「激光强度」，于是 Excel
    # 「激光能量」的候选里混进了设备号里的数字（T-1011 → 1011），
    # 判成 scope=multi（两套值），一对好端端的数据被标成"需人工确认"。
    if _EQUIP_NAME_RE.search(n):
        return None
    # —— 加热 + 部位（中间模 / 左端模 / 右端模 / 端模）——
    # WI-0036 的 PDF 把三个部位并排写在**同一格**里，Excel 却各建了一个采集项
    # （加热中间模 / 加热左端模 / 加热右端模）。第 9.5 节把格拆开之后名字与 Excel
    # 一致，但词典里原本只有「加热中间模」，两个端模全判「参数名未收录」
    # （实测 20 条 no_match）。部位词取自 `EN_PART_LABELS`，加部位只改那张表。
    # ⚠️ 必须排除「退出条件加热…」—— 那是另一项，与「加热…」撞键就是假一致。
    if "加热" in n and "退出" not in n:
        for w, _c in EN_PART_LABELS:
            if w in n:
                return "加热" + w
    # 步骤式参数（WI-0021 型横表）：PDF 叫法 ↔ Excel 叫法的桥
    if "激光" in n or "laser" in n:
        return "激光强度"
    # Excel 写「步骤1焊接能量(mw)」、PDF 那一栏叫「激光强度」—— 同一台激光焊接机、
    # 同一个 mw 单位，只是叫法不同（WI-0027 实测）。不认这一条，整项会判成
    # 「参数名未收录」而完全不参与核对（2160 条）。
    if "焊接" in n and "能量" in n:
        return "激光强度"
    if "光斑" in n or "spot" in n:
        return "光斑大小"
    if "转动" in n or "旋转" in n or "转速" in n or "rotation" in n:
        return "转动速度"
    if "移动" in n and "距离" in n or "位移" in n:
        return "移动距离"
    if ("移动" in n and "速度" in n) or "move speed" in n:
        return "移动速度"
    if "焊接" in n and "时间" in n:
        return "焊接时间"
    # 「带限定的温度」与「温度」不是同一个东西，绝不能并成一个键。
    # 本节下面只用 `"温度" in n` 兜底，于是「开始温度（℃）」「结束温度（℃）」也被
    # 收成 `温度`，与拉伸机的「温度」撞键。WI-0039 实测：Excel 那两项的设备号是
    # T-1091（焗炉机），而 PDF 的「Temperature 温度 110~130」挂在 T-2084/T-1279 上，
    # 靠后面「按设备收窄」那一层才没误判 —— **一旦设备号相同**，就会拿 110~130 去比
    # 75，报一个假的不一致。假不一致比漏判更伤：工具的全部可信度都建立在
    # 「报出来的不一致都是真的」上面。
    # 必须放在下面「英文（专用名优先）」之前：`Start Temperature` 得先归到「开始温度」，
    # 否则会被英文的 `"temperature" in n` 抢走、又并回 `温度`。
    if "开始温度" in n or "起始温度" in n or ("start" in n and "temp" in n):
        return "开始温度"
    if "结束温度" in n or "终止温度" in n or ("end" in n and "temp" in n):
        return "结束温度"
    # 英文（WI-0011 等，专用名优先）
    # 两个关键词**同时**命中才算，单靠一个（如只有 pressure）走不到这里 ——
    # 这条性质很重要，见函数末尾双语名回退时的说明。
    for a, b, key in _PROPER_EN_PAIRS:
        if a in n and b in n:
            return key
    # 一般英文（兜底，覆盖纯英文参数名）
    if "temperature" in n or "temp" in n:
        return "温度"
    if "pressure" in n:
        return "压力"
    if "speed" in n or "velocity" in n:
        return "速度"
    if "intensity" in n:
        return "光强"
    if "flow" in n:
        return "风量"
    if "cure" in n or "curing" in n:
        return "固化时间"
    if "dwell" in n:
        return "停留时间"
    # WI-0006 型附表：PDF 写「持续时间」「焊接距离」，Excel 写「步骤N的焊接时间」
    # 「步骤N的移动距离」——叫法不同，指的是同一个格子。2026-09-27 补。
    # 补之前 `_canonical_param("焊接距离")` 返回 None，整份 WI-0006（198 个参数点）
    # 全部落到 scope=name_unknown，界面上跟「PDF 里根本没有」长得一模一样。
    #
    # 位置必须在英文专用名（pleat_dwell / fold_dwell）**之后**：WI-0011 的
    # 「Pleat Dwell Pleat 持续时间」里也含「持续时间」四个字，抢在前面会把
    # pleat_dwell 拆掉 —— 实测退化 6 条 ok → no_match，是零回归当场拦下来的。
    if "持续时间" in n or "保持时间" in n:
        return "焊接时间"
    if "焊接距离" in n:
        return "移动距离"
    if "position" in n:
        return "位置"
    if "program" in n:
        return "程序号"
    if "time" in n:
        return "时间"
    # 中文（通用）
    # —— 以下四条是 2026-09-24 补的。此前它们**一个都不在表里**，`_canonical_param`
    # 对它们返回 None，而 `_vertical_expected` 拿到 None 就直接放弃 → 这些参数
    # 永远判「无对应」。WI-0039 的「速率（mm/sec）」就是这么被埋掉的：PDF 里
    # 明明写着「Digital Rate mode 速率 10-12 mm/sec」、设备 T-2084/T-1279
    # 与 Excel 完全对得上，只因为「速率」这个词没被收录而静默跳过。
    # 这也是词典式方案的固有短板：词典是按已见过的文档攒的，**每来一份新 WI 都可能
    # 再撞**。所以配套做了第 2 件事：认不出来时不再静默，见 `_vertical_expected`。
    if "速率" in n:
        # 只认中文「速率」，**绝不能**拿英文 "Rate" 当关键词：
        # WI-0008 的「Flow Rate 风量」含 Rate，那是风量，必须留给上面的 "flow" 规则。
        return "速度"
    # —— 亲水涂层「溶液次数」（2026-09-29 第 42 轮补）——
    # PDF 那一栏叫「溶液亲水次数」，值写「使用底层溶液1次 / 使表层溶液2次」；
    # Excel 拆成两行「使用底层溶液次数 (次)」「使表层溶液次数  (次)」。
    # 两边都含「溶液」+「次数」，但**名字对不上**（PDF 是「溶液亲水次数」），
    # 词典原本一条都没收录 → WI-0031 84 行落进 `name_unknown`，
    # 界面上跟「PDF 里真没有」长得一模一样（这是最容易被误判成识别故障的一档）。
    # 判据用两个词**同时**出现，不能只认「溶液次数」连写。
    if "溶液" in n and "次数" in n:
        return "溶液次数"
    # —— 第九种版式（英文横表）的中文叫法，2026-09-29 补 ——
    # 这些名字**必须**排在通用键（温度/压力/速度/时间/位置/拉伸距离）之前：
    # 「退出条件时间」里也有「时间」、「左拉伸距离」里也有「拉伸距离」，
    # 让通用键先命中就会把它们并成同一个键 —— 那正是 WI-0039「开始温度/温度」
    # 撞键的老毛病（假一致就那么来的）。
    if "退出条件" in n and ("加热" in n or "中间模" in n or "温度" in n):
        return "退出条件温度"
    if "退出条件" in n and ("时间" in n or "秒" in n or "sec" in n):
        return "退出条件时间"
    if "加热中间模" in n:
        return "加热中间模"
    if "冷却中间模" in n:
        return "冷却中间模"
    if "排气中间模" in n or "排气" in n:
        return "排气中间模"
    if "抽真空" in n:
        return "抽真空"
    if "降压" in n:
        return "降压"
    if "充气速度" in n or "充气速率" in n:
        return "充气速度"
    if "左拉伸距离" in n:
        return "左拉伸距离"
    if "右拉伸距离" in n:
        return "右拉伸距离"
    if "拉伸距离" in n or "tensile distance" in n:
        return "拉伸距离"
    if "模具" in n or re.match(r"die(?![a-z])", n):
        # `die(?![a-z])` 而非 `"die" in n`：后者会把 "diesel" 之流也算进来。
        # 这里不能用 `\b`——py3 下汉字算 \w，「die模具」的 e 后面没有词边界。
        return "模具"
    if "芯轴" in n or "mandrel" in n:
        return "芯轴"
    if "固化" in n:
        return "固化时间"
    if "固化时间" in n:
        return "固化时间"
    if "停留时间" in n:
        return "停留时间"
    if "温度" in n:
        return "温度"
    # —— 通用量纲（第 54 轮）：只论「大类 + 限定词」，不再把限定词吞掉 ——
    # 这三条必须留在「具体键」之后、「速度/时间/位置」之前：
    # 上面的 模具/芯轴/移动距离/拉伸距离/开始温度 等已经把固化好的具体键抢走了，
    # 到这里剩下的才是真正需要区分限定的名字。
    # ⚠️ 别把 速度/温度/时间/位置 也改成这一套 —— 那几类靠同义词合并
    #    （上提/提拉/上升 → 速度），加限定词会把同一个参数拆成两个键，凭空造漏报。
    if "频率" in n:
        return _qual_key(n, "频率", ("频率",))
    if "气压" in n or "压力" in n:
        return _qual_key(n, "压力", ("压力", "气压"))
    if "距离" in n:
        return _qual_key(n, "距离", ("距离",))
    if "风量" in n:
        return "风量"
    if "程序号" in n:
        return "程序号"
    if "光强" in n:
        return "光强"
    if "速度" in n or "上提" in n or "移动" in n or "提拉" in n or "上升" in n or "下移" in n \
            or "下沉" in n:
        # 「下沉」与「下移」是同一个动作的两种写法：WI-0042 的 Excel 写
        # 「底涂 下沉浸没1」，PDF 附表里那一行却叫「下移浸没01」。
        # 不并成一类，字面永远对不上（词典键是这个粗粒度的「运动类」桶，
        # 真正把 01/02 分开的是后面的「运动方向 + 序号」那一层）。
        return "速度"
    # 「时长」是「时间」的同义写法（「焗炉时长」），词典原本只认「时间」二字，
    # 于是 `焗炉时长` 落进 name_unknown。补在「时间」之前，因为「时长」本身就是
    # 更具体的持续时间义，且现有语料里「时长」只出现在「焗炉时长」一个参数名上，
    # 不会撞键、不抢任何已有判定（2026-10-13 补）。
    if "时长" in n:
        return "时间"
    if "时间" in n:
        return "时间"
    if "位置" in n:
        return "位置"
    # 「切割长度」（Excel）↔「切割长度（mm）」（PDF 表头）：2026-09-28 WI-3264。
    # 放在中文段**最末**，前面所有更具体的键（温度/速度/时间/位置/拉伸距离…）都优先，
    # 这样加一条通用键不会抢走任何已有判定。
    # 「定长」必须比「长度」先判 —— WI-3264 同一张表里「切割长度（mm）」是**规格档**
    # （1.5/2.5/3.5/7.5/8.5）、「切割后定长（mm）」才是**值**（31±2 / 30±2…），
    # 两者并成一个键就会拿档位去比数值，那是最典型的假不一致。
    if "定长" in n:
        return "定长"
    # 英文 `Cut Length` 必须跟中文「长度」同一个键（2026-10-29）——
    # WI-0044 p6 里表1 的中文列叫「长度设定及接收标准」、表3 的英文列叫
    # `Cut Length`，两张表各写各的长度档，其实都是「长度」。不并键的话，
    # Excel「长度设定」只能看见表1 那 9 个与它无关的值，永远看不到表3 里
    # 真正对应的那一条 —— 而表3 恰恰是唯一带 P/N 的（能分档的那张）。
    if "长度" in n or "length" in n:
        return "长度"
    # —— 双语名：中文一侧没收录时，再试一次**英文复合专用名** ——
    # 「Vacuum Delay/真空延迟(s)」（DWG-0000003 卡0002）就栽在这：PDF p6 上
    # 「Vacuum Delay 10 seconds」好端端抽到了，而 Excel 侧按上面的规矩只取了
    # 中文一侧「真空延迟(s)」，词典里又没有这条中文 → 判「参数名未收录」。
    #
    # ⚠️ 只能退回 `_PROPER_EN_PAIRS` 这一小段（两个关键词同时命中才算），
    # **绝不能**退回通用兜底：「Alarm Pressure Drop/报警线」的中文侧「报警线」
    # 本来就该是"不认识"，一旦退回通用规则，会被 `"pressure" in n` 收成「压力」，
    # 与「Test Pressure/测试压力」撞键 —— 那正是当初改成「只取中文一侧」要防的事。
    if n != raw.lower():
        en = raw.lower()
        for a, b, key in _PROPER_EN_PAIRS:
            if a in en and b in en:
                return key
    return None


# --------------------------------------------------------------------------
#  参数名「统称 → 具体」展开表（第 53 轮）
#
#  PDF 一张表里「加热端模」往往只有**一列**（左右两端同值），Excel 却给左、右
#  各建了一个采集项（加热左端模 / 加热右端模）。两个规范名对不上，于是：
#
#    · Excel 的「加热左端模/10」在自己该看的那张表里找不到格子，只能去借
#      **隔壁规格档、另一批设备**的同名参数当期望值 —— WI-0036 card 0015
#      （T-10706…那 6 台，Ø6.0）就是这样被拿表2（T-1177，Ø6.0）的 110~140
#      判成「不一致」，而它自己那张表3 的第 10 步分明写着 110~145。
#    · 更糟的是「看起来判出来了」：期望值来自另一台设备，是**假一致**的近亲。
#
#  展开方向只有一条：**统称 → 具体**，且**只加不减** —— 具体名之间绝不互相展开。
#
#   统称（PDF 侧写法）      展开为（Excel 侧采集项）
#   ───────────────────────────────────────────────
#   端模                   左端模、右端模
#
#  ⚠️ 判据是**完整参数名里的部位词，且该部位词前面没有「左/右」** —— 与
#     `_EN_PART_RE` 的 `(?<![左右])` 同一口径：「左端模」里的「端模」是子串，
#     不能当统称（否则 `加热左端模` 会展开成 `加热左左端模`，当场错乱）。
#
#  ⚠️ 只在**建索引**时展开（`_param_keys`），PDF 记录本身**不复制**：
#     复制会让记录数虚增，`step_tables` 的步骤集合、`by_equip` 明细全都被带偏。
# --------------------------------------------------------------------------
_PARAM_ALIAS_RUNS = (
    ("端模", ("左端模", "右端模")),
)

# `(?<![左右])`：统称前面是「左/右」时不认（那是具体名里的子串）
_PARAM_ALIAS_RE = re.compile(
    "(?<![左右])(" + "|".join(w for w, _s in _PARAM_ALIAS_RUNS) + ")")


def _param_keys(name):
    """一个 PDF 参数名 → 它该挂到哪些规范名上（主名 ∪ 统称展开出的具体名）。

    返回值是**规范名集合**，供 `step_index` / `step_tables` / `sub_params`
    建索引用。主名一定在集合里；只有参数名里含「统称部位」时才多出展开名。

    认不出（主名为 None）返回空集合 —— 调用方按「这条记录不进索引」处理，
    与 `_canonical_param` 原来返回 None 时的行为一致。
    """
    main = _canonical_param(name)
    if not main:
        return set()
    out = {main}
    src = str(name or "")
    for generic, specifics in _PARAM_ALIAS_RUNS:
        m = _PARAM_ALIAS_RE.search(src)
        if m and m.group(1) == generic:
            for sp in specifics:
                alt = _canonical_param(src.replace(m.group(1), sp, 1))
                if alt:
                    out.add(alt)
    return out

_WI_NO_RE = re.compile(r"WI\s*[-–—]\s*([A-Za-z]{0,2}\d{3,6})")
# 文件名里的编号连字符可能被省掉（「WI2012,rev…」），故此处连字符可选。
_WI_NO_FN_RE = re.compile(r"WI\s*[-–—]?\s*([A-Za-z]{0,2}\d{3,6})")


def _norm_wi_no(s):
    """把各种写法的 WI 编号归一成 `WI-编号`：去掉连字符两侧的空格与非断行空格。

    为什么必须归一：Excel「表4」里那一列是按**字符串精确比对**筛采集项的
    （见 `TableSchema.records` 的 `wi != wi_name`），差一个空格就一条都取不到，
    而界面上只会说「所有参数表里都没有这个 WI 的采集项」—— 看起来像参数表的问题。
    PDF 文件名里恰恰常用不换行空格（`WI-M2471,\\xa0rev13…`）。
    """
    m = re.match(r"\s*WI\s*[-–—]?\s*([A-Za-z]{0,2}\d{3,6})\s*$", str(s or ""), re.I)
    return "WI-" + m.group(1).upper() if m else str(s or "").strip()


def detect_wi(pdf_path: str) -> str | None:
    """从 PDF 识别文件编号 WI-XXXX；识别不到返回 None。

    两级尝试（能识别的先识别，2026-09-27）：
      ① 文字层：首页/第二页找「文件编号 WI-XXXX」或裸 WI-XXXX（老口径，不变）；
      ② 文件名：文字层没有时，认文件名里的 WI-XXXX。
         扫描件 PDF（每页一张图、0 字符）文字层必然空手而归，但文件名往往就是
         「WI-0001, rev18 …」—— 放着不认，界面就会退回内置的 WI-0008，
         把「识别失败」演成「识别成了别的 WI」（2026-09-27 实际发生）。
         文件名是**兜底**不是首选：文件可以被随手改名，文字层里的编号才权威。

    ⚠️ 编号有两种写法：纯数字 `WI-0013`，和**带字母前缀**的 `WI-M2471`、
    `WI-M20017`（OEM 与设备操作类，现行目录里 313 份）。旧正则只认 `\\d+`，
    这些 WI 连编号都识别不出来 —— 界面上表现为「所有参数表里都没有这个 WI 的
    采集项」，看着像参数表缺数据，实际是编号没认出来（用户 2026-10-14 反馈
    「工具无法识别 WI 带 M 的」）。字母前缀限 1~2 位、数字 3~6 位，
    够覆盖现有全部写法，又不至于把正文里的普通词误认成编号。
    """
    try:
        with open_pdf(pdf_path) as pdf:
            for i in range(min(3, len(pdf.pages))):
                text = pdf.pages[i].extract_text() or ""
                m = re.search(r"文件编号\s*(WI\s*[-–—]\s*[A-Za-z]{0,2}\d{3,6})", text)
                if m:
                    return _norm_wi_no(m.group(1))
                m = _WI_NO_RE.search(text)
                if m:
                    return "WI-" + m.group(1).upper()
    except Exception:
        # PDF 本身坏了（`No /Root object!` 这类：传输中断、下载不全）。
        # 以前这里直接往外抛，整个批量核对就此中断 —— 一份坏文件拖垮一整批。
        # 读不了文字层不等于不知道它是谁：文件名还是可信的线索，交给下面的兜底。
        pass
    m = _WI_NO_FN_RE.search(os.path.basename(pdf_path))
    if m:
        return "WI-" + m.group(1).upper()
    return None


# ---------------------------------------------- 「同一份 PDF 只抽一遍」的小缓存
#
# 一次比对里，同一个抽取会被走两遍：顶层函数走一遍（extract_pdf_steps /
# extract_pdf_vertical / extract_pdf_programs），extract_pdf_meta 汇总表头元信息时
# 又走一遍 —— 而它们扫的是同一批页面。共享页面（本文件 open_pdf）之后每次扫描
# 已经很便宜，但「按页面重新切表 + 重新拼记录」这层纯 Python 活儿还是会重做。
#
# 每个抽取只留「当前这一份」的缓存，不把整批 PDF 的结果攒在内存里。

_STEPWISE_CACHE = {}
_VERTICAL_CACHE = {}
_PROGRAM_CACHE = {}
_IMGBLOCK_CACHE = {}


def _cached_by_file(cache, pdf_path, produce):
    """按「路径 + 修改时间」缓存某份 PDF 的一个抽取结果，只留当前这一份。"""
    try:
        key = (os.path.abspath(pdf_path), os.path.getmtime(pdf_path))
    except OSError:
        key = (str(pdf_path), 0)
    hit = cache.get(key)
    if hit is None:
        cache.clear()                   # 换文件了：旧的丢出去，别攒着
        hit = produce(pdf_path)
        cache[key] = hit
    return hit


def _stepwise_all(pdf_path):
    """extract_stepwise 的全量结果（含有值格 + 明确空值格），按路径+mtime 缓存。

    「有值」和「空格」要分两路用，但 PDF 只能解析一次 —— 每份作业指导书解析
    要好几秒，取两次结果会白白翻倍。
    """
    return _cached_by_file(_STEPWISE_CACHE, pdf_path, extract_stepwise)


def _vertical_all(pdf_path):
    """extract_vertical 的缓存（顶层抽取 + 元信息汇总各走一遍）。"""
    return _cached_by_file(_VERTICAL_CACHE, pdf_path, extract_vertical)


# ------------------------------------------------------------ 图片块（2026-11-02）
#
# 起因（用户实案）：WI-0039 第 4 页的「表7 球囊保护套切定长、扩口」**整块是图片**，
# 文字层里一个字都没有。8 条「切定长长度」因此判成「PDF 里没有这一项」，
# 而表7 上明明印着 35/60/100/140/170/220/270/330 mm —— 和 Excel 的
# 30~40 / 55~65 / … / 325~335 一一对应，全部一致。
# 用户自己看出来了：「这个部分看着像是图片类型的PDF啊，所以才会读取不到吧」。
#
# 工具的旧说法是「参数名已收录，但这份 PDF 的参数表里没有它」—— 技术上没说谎
# （文字层里确实没有），但把它说成「PDF 没有」是错的，而且**完全没提图片**。
# 这一节就是让工具自己把这件事说出来：哪几页印着图、图占多大。
#
# 判据用「渲染后的版面尺寸」，不用像素：同样一张 657×91px 的图，可能印成
# 一整个图章，也可能印成半页大的表格截图 —— 只有版面尺寸才说明它能不能装下一张表。
_IMG_MIN_RATIO = 0.02      # 占整页面积 2% 以上才算「一块图」
_IMG_MAX_RATIO = 0.85      # 顶到 85% 以上的当整页图（封面 / 整页扫描）
_IMG_MIN_SIDE = 60.0       # 最短边不足 60pt 的当图标/装饰（页眉 logo 实测 82×25pt）


def _image_blocks(pdf_path):
    """每页里「大到能装下一张表、且文字层没覆盖」的图片块。

    返回 {页号: [(x0, top, x1, bottom, 占比), …]}。

    只收**够大的**：2% 面积 + 60pt 最短边。页眉那条 82×25pt 的 logo（0.4%）
    和小图标进不来，否则每一份 WI 都会被说成「有图片」。
    顶到 85% 以上的整页图也不收：那是封面或整页扫描，另有 pdf_scan_warning
    在整份文档这一层拦，混进来只会让每一份 WI 的第一行都报「第1页占页100%」。
    """
    out = {}
    with open_pdf(pdf_path) as pdf:
        for pno, page in enumerate(pdf.pages, 1):
            area = float(page.width) * float(page.height) or 1.0
            boxes = []
            for im in (page.images or []):
                try:
                    x0, x1 = float(im["x0"]), float(im["x1"])
                    y0, y1 = float(im["top"]), float(im["bottom"])
                except (KeyError, TypeError, ValueError):
                    continue
                w, h = abs(x1 - x0), abs(y1 - y0)
                if min(w, h) < _IMG_MIN_SIDE:
                    continue
                ratio = (w * h) / area
                if not (_IMG_MIN_RATIO <= ratio < _IMG_MAX_RATIO):
                    continue
                boxes.append((round(x0, 1), round(y0, 1), round(x1, 1),
                              round(y1, 1), round(ratio, 4)))
            if boxes:
                boxes.sort(key=lambda b: -b[4])
                out[pno] = boxes
    return out


def pdf_image_blocks(pdf_path):
    """`_image_blocks` 的按文件缓存版（与其它抽取同一条缓存规矩）。"""
    if not pdf_path:
        return {}
    return _cached_by_file(_IMGBLOCK_CACHE, pdf_path, _image_blocks)


def image_blocks_note(blocks, max_pages=4):
    """把图片块翻成一句人话：「第4页 2 处（占页 5% / 10%）、第7页 3 处」。"""
    if not blocks:
        return ""
    bits = []
    for p in sorted(blocks)[:max_pages]:
        bs = blocks[p]
        pct = " / ".join("%.0f%%" % (b[4] * 100) for b in bs[:2])
        bits.append("第%d页 %d 处（占页 %s）" % (p, len(bs), pct))
    more = len(blocks) - max_pages
    if more > 0:
        bits.append("另有 %d 页" % more)
    return "、".join(bits)


def _programs_all(pdf_path):
    """extract_program_tables 的缓存（顶层抽取 + 元信息汇总各走一遍）。"""
    return _cached_by_file(_PROGRAM_CACHE, pdf_path, extract_program_tables)


def release_caches():
    """放掉这一趟留下的所有缓存（页面对象 + 各抽取结果），把内存还给系统。

    调用时机：**关窗时**，不是每跑完一次比对。原先每跑完就清，代价是下一次
    点击要重新解析整份 PDF —— 实测 1.03s 对 0.19s，慢 5 倍，而省下的内存只有
    几十 MB（页面对象常驻约 30～60MB）。界面是「换个 WI 再点一次」的用法，
    这个买卖不划算。

    不清也不会越攒越多：每份缓存的键都是「路径 + 修改时间」，换文件时
    _cached_by_file 会先 clear 再放新的；open_pdf 也会在换文件时关掉上一份。
    也就是说缓存上限始终是「当前这一份 PDF」。
    """
    for c in (_STEPWISE_CACHE, _VERTICAL_CACHE, _PROGRAM_CACHE, _IMGBLOCK_CACHE):
        c.clear()
    release_pdf_cache()


def extract_pdf_steps(pdf_path):
    """抽取「参数 | Step 1 | Step 2 …」横向参数表（WI-0021 型版式）。

    只返回**有值**的格子 —— 下游的期望值候选只能用真值。
    空格子另走 extract_pdf_step_blanks()。
    """
    return [r for r in _stepwise_all(pdf_path) if not r.get("empty")]


def extract_pdf_step_blanks(pdf_path):
    """横表里「格子存在、但明确写空」的记录（`------` / `—` / 空白）。

    这是 PDF 在声明「此处无值」，与「压根没有这个格子 / 这一步」是两回事。
    有了它，比对层才能给出正确结论：
      · PDF 空 + Excel 填 N/A   → 一致（N/A 正确）
      · PDF 空 + Excel 留空      → 一致（不填正确）
      · PDF 空 + Excel 填了值    → 不一致（PDF 本来就没值）
      · PDF 有值 + Excel 留空    → 不一致（**漏填**，以前只会含糊报「未填值」）
    """
    return [r for r in _stepwise_all(pdf_path) if r.get("empty")]


def extract_pdf_unparsed(pdf_path):
    """这份 PDF 里「有字、但我没能变成一个值」的格子清单（见 `LAST_UNPARSED`）。

    这是**读表对账**的出口：让「我没读懂」和「文件里没有」成为两件事。
    判定层拿它改写 no_match 的依据；界面/冒烟脚本拿它显示「这份文件我读懂了多少」。
    袋子在 `reset_pdf_failures()` 时清空，由 `extract_pdf_steps` 那一路顺带填充 ——
    所以**必须在 compare 之前跑过抽取**（界面与所有脚本都是这个顺序）。
    """
    if not LAST_UNPARSED:
        _stepwise_all(pdf_path)
    return list(LAST_UNPARSED)


def pdf_read_audit(pdf_path):
    """这份 PDF 的**读表体检**：一共看到多少格、认出多少、哪些没认出来。

    返回 {"cells": n, "values": n, "unparsed": [...], "rate": 0~1}。
    `cells` = 参数表里「有字的格子」总数（不含 Step# 那一列）；
    `values` = 真正变成值的记录数；`rate` = values / cells。

    用法：任何一份新语料进来，先跑这个。**rate 明显偏低 = 这个版式工具还没吃透**，
    该去补版式，而不是等着用户拿判定结果来反问「为什么又错了」。
    """
    vals = _stepwise_all(pdf_path)
    un = extract_pdf_unparsed(pdf_path)
    cells = len(vals) + len(un)
    return {"cells": cells, "values": len(vals), "unparsed": un,
            "rate": (float(len(vals)) / cells) if cells else 1.0}


def extract_pdf_programs(pdf_path):
    """抽取「设备参数(设置)附表N」转置双列附表（WI-0009 / WI-3852 型版式）。

    这类表把期望值按「附表 × 程序列(底层/表层溶液) × 步骤」切了很多份，
    必须逐份带出处抽出来，否则期望值只能是整份 PDF 的并集（等于没核对）。
    """
    return _programs_all(pdf_path)


# ----------------------------------------------------------------- 图号 / 正文取值

# 图号：同一份 PDF 常按图号分出几套参数（WI-0016 p4：表1/表3 是 DWG-0000006、
# 表2/表4 是 DWG-0000004），设备号还可能重复 —— 图号是唯一能分开它们的线索。
# `DWG-0000006-01`（带后缀）、`DWG-0000001`（数字前带横杠）都要认。
_PN_RE = re.compile(r"DWG[-\s–—]*(\d{4,8})")


def _pn_cores(text):
    """文本里写着的图号集合（frozenset）；一个都没有就返回空集。"""
    s = re.sub(r"\s+", " ", str(text or ""))
    return frozenset(_PN_RE.findall(s))


# 完整图号（含后缀档位）。`_pn_cores` 只取数字主干 —— 对「一份 PDF 一个图号」
# 够用，但对「一张表一行一个图号」就不够了：DWG-0000007-01/-02/-03/-04 会被
# 全压成 2000022，四行永远分不开，收窄链只能把四套值并成一个池子。
# 2026-10-29 WI-0044 p6 实测：表3 22 行各一个 P/N，Excel 侧 spec 也写着
# DWG-0000007-01，两边信息都全，就是**后缀被吃掉了**，于是 4 条采集项各自
# 对着 9 个长度值判，全判「不一致」（假差异）。
_PN_FULL_RE = re.compile(r"DWG[-\s–—]*(\d{4,8})(?:\s*[-–—]\s*(\d{1,3}))?", re.I)


def _pn_full(text):
    """文本里的**完整**图号集合，主干与带后缀两级都放进去。

    `DWG-0000007-01` → {'2000022-01', '2000022'}

    为什么两级都要：PDF 侧常只写主干（表头一句 `DWG-0000001`），Excel 侧
    常带后缀（`DWG-0000007-01`）。只比完整串会把「两边其实是同一个」判成对不上；
    只比主干又分不开同主干的四个后缀。两级并存，收窄时按「精确优先、主干兜底」
    两段走 —— 精确能分档，兜底不漏。
    """
    s = re.sub(r"\s+", " ", str(text or ""))
    out = set()
    for core, suf in _PN_FULL_RE.findall(s):
        out.add(core)
        if suf:
            out.add("%s-%s" % (core, suf))
    return frozenset(out)


def _pn_cell_key(cell):
    """**整格就是一个图号**时返回它的完整图号集，否则空集。

    为什么要这道「整格」判定：正文句子里常顺带提到别的图号
    （「…参见 DWG-0000006-01 的说明」），那不是本行的档位；照抓会把行绑错档，
    比不抓更糟 —— 错档会给出一个**错得很像**的期望值。
    """
    s = _c(cell) or ""
    if "DWG" not in s.upper():
        return frozenset()
    rest = _PN_FULL_RE.sub("", s)
    rest = re.sub(r"[\s/、,，;；()（）\-–—_]+", "", rest)
    # 图号之外最多剩几个字符（P/N 标记、括号说明），再长就不是「一格一个图号」了
    if len(rest) > 6 or re.search(r"[一-鿿]", rest):
        return frozenset()
    return _pn_full(s)


def _equip_name_map(pdf_path):
    """设备清单里的「设备类别名 → 编号组」，如 焗炉机 → [T-1030, T-1079, …]。

    正文步骤里常常**只写设备类别、不写编号**：「…与芯轴一起进行焗炉处理，温度75oC」。
    不把类别名接上，这条值就挂不到任何设备上，Excel 的「内管焗炉 开始温度」
    永远查无此值 —— 而值明明就印在 PDF 上。
    """
    out = []
    try:
        with open_pdf(pdf_path) as pdf:
            for page in pdf.pages:
                for t in (page.find_tables() or []):
                    rows = t.extract() or []
                    hidx, nc = None, None
                    for hi, hr in enumerate(rows):
                        for ci, c in enumerate(hr or []):
                            cj = str(c or "")
                            if "编号" in cj and ("设备" in cj or "工具" in cj):
                                hidx, nc = hi, ci
                                break
                        if hidx is not None:
                            break
                    if hidx is None:
                        continue
                    for row in rows[hidx + 1:]:
                        if nc is None or nc + 1 >= len(row or []):
                            continue
                        s = re.sub(r"(T-)\s+", r"\1", _flat(str(row[nc] or "")))
                        devs = equip_from_text(s)
                        if not devs:
                            continue
                        # 名称格里常是「Mandrel, …|聚对二甲苯涂层芯轴」—— 只取中文那一段
                        nm = ""
                        for p in re.split(r"[\n|]", str(row[nc + 1] or "")):
                            if re.search(r"[\u4e00-\u9fff]", p):
                                nm = p.strip()
                        if len(nm) < 2 or len(nm) > 12:
                            continue
                        kws = {nm}
                        core = re.sub(r"(机|器|仪|设备|装置)$", "", nm)
                        if len(core) >= 2:
                            kws.add(core)
                        out.append((tuple(sorted(kws, key=len, reverse=True)), devs))
    except Exception:
        pass
    return out


# 正文叙述里的「参数词 + 数值 + 单位」：温度75oC / 温度设定50oC / 时间30分钟
_PROSE_WORD = "温度|时间|时长|速度|速率|压力"
_PROSE_RE = re.compile(
    r"(?P<name>" + _PROSE_WORD + r")"
    r"\s*(?:设定|约为|为|达到|至|到|：|:)?\s*"
    r"(?P<num>\d+(?:\.\d+)?)"
    r"\s*(?P<unit>℃|oC|ºC|°C|°c|度|分钟|秒钟|小时|min|sec|s|mm/s|mm/sec|mm|MPa|Mpa|psi|%)")


def _prose_page(page, pgno, name2eq):
    """从**正文叙述**里捞参数值（不在任何参数表里的那种）。

    WI-0016 p3 第 9 步：「…与芯轴一起进行焗炉处理, 温度75oC, 时间30分钟」——
    焗炉的温度/时间压根没进参数表，Excel 却有「开始温度 / 结束温度」两行。
    这类值只能从正文捞，参数表路径再多也够不着。

    设备归属用**就近认领**：同一句里可能提到好几类设备
    （「拉拉伸过的内管」有「拉伸」、「芯轴」、「焗炉处理」），
    取**离这个数值最近**的那一个。实测「焗炉处理, 温度75oC」里 焗炉 距 6 字、
    芯轴 距 12 字、拉伸 距 19 字 → 归焗炉机，正解。
    """
    try:
        txt = page.extract_text() or ""
    except Exception:
        return []
    # 中文换行只是任意断行，会把「温\n度设定50oC」拆坏 —— 只在中日韩之间吃掉换行
    txt = re.sub(r"(?<=[\u4e00-\u9fff\u3000-\u303f])\s*\n\s*"
                 r"(?=[\u4e00-\u9fff\u3000-\u303f])", "", txt)
    out = []
    for sent in re.split(r"[。；;\n]", txt):
        if not sent.strip():
            continue
        pns = _pn_cores(sent)
        for m in _PROSE_RE.finditer(sent):
            vpos = m.start()
            best, bestd = None, None
            for kws, devs in name2eq:
                for kw in kws:
                    p = sent.find(kw)
                    while p != -1:
                        d = abs(p - vpos)
                        if bestd is None or d < bestd or (d == bestd and len(kw) > len(best)):
                            best, bestd = kw, d
                        p = sent.find(kw, p + 1)
            eq = []
            if best is not None:
                for kws, devs in name2eq:
                    if best in kws:
                        eq = list(devs)
                        break
            out.append({
                "page": pgno, "table": -1, "sub": "", "equip": eq,
                "equip_raw": "/".join(eq), "dia": None, "dias": [], "prog": "",
                "families": [], "param": m.group("name"), "text": m.group(0),
                "raw": m.group(0), "pn": pns or None, "prose": True,
            })
    return out


# ------------------------------------------- PDF 表头元信息（当成「参数」来核对）
#
# 为什么单列这条路：参数表的表头有一段键值块（激光机编号 / 程序号 / 焊接保护套 /
# 使用芯轴 / 设备·工具编号）。版式层其实早就把它抽到了 —— extract_stepwise 每条记录
# 都带 meta / prog / sleeve，extract_program_tables 每条带 prog_no_raw。
#
# 但比对层从没把它当成「参数」：它只跟着「激光强度」「光斑大小」那类数值走，
# 自己从来不作为一项被核对。后果是 Excel 里【每个 WI 都有】的
# 「XX 机器编号」「程序号」「焊接保护套」「使用芯轴」采集项，全部落进
# 「采集项(不核对)」或「无对应」—— 一条都没核过。
#
# 而这几项恰恰是最容易填错的：机器编号决定用哪台机的参数，程序号决定调哪个程序，
# 填错了整批产品的工艺参数全错，却恰好是唯一没有覆盖的一类。
#
# 所以这里把「已经在数据里」的元信息聚合成记录，交给与其它参数同一条收窄链。

# PDF 表头键名 → 规范化类别（用 _flat 归一后查）
_PDF_META_KEY = {
    "激光机编号": "机器编号",
    "激光机型号": "机器编号",
    "激光机台编号": "机器编号",
    "机台编号": "机器编号",
    "设备编号": "机器编号",
    "设备/工具编号": "机器编号",
    "设备工具编号": "机器编号",
    # 「机器编号」是最基础的写法（WI-0026 p5/p6 七张子表全用它），
    # 以前字典里只有「激光机… / 设备…」这一类 → 那一台机器在 PDF 里
    # "表头没有这一栏"，剩下的只能靠设备清单表兜底判（meta_list）。
    "机器编号": "机器编号",
    "机器型号": "机器编号",
    "设备型号": "机器编号",
    "程序号": "程序号",
    "程序号/名称": "程序号",
    "程序名称": "程序号",
    "焊接保护套": "保护套",
    "保护套": "保护套",
    "焊接保护套/pn": "保护套",
    "使用芯轴": "芯轴",
    "芯轴": "芯轴",
}
_PDF_META_KEY_FLAT = {_flat(k): v for k, v in _PDF_META_KEY.items()}

# Excel 采集项名 → 元信息类别
_META_ITEM_KINDS = (
    ("机器编号", "机器编号"),
    ("机台编号", "机器编号"),
    ("设备编号", "机器编号"),
    ("机器型号", "机器编号"),
    ("程序号", "程序号"),
    ("程序单号", "程序号"),
    ("保护套", "保护套"),
    ("芯轴", "芯轴"),
)
# 采集项以「机器 / 设备 / 机台」**收尾**也按「机器编号」核对（WI-3264）。
# 两条边界：
#   · 前面不能紧挨着数字或字母 —— 「1.5mm机器」是规格不是设备项；
#   · 「机器」必须是名字的最后一个词，开头出现不算（「机器 Atropos」）。
_META_EQUIP_TAIL_RE = re.compile(r"(?<![0-9A-Za-z])\s*(?:机器|设备|机台)[：:\s]*$")

# 「保护套 / 芯轴」是**光秃秃的名词**，可以出现在动词短语中间。上面那张表用的是
# 纯子串 `in`，于是 WI-0039 表4 的采集项
#     使用模板将球囊**保护套**切定长，长度      （期望 30~40mm，浮点）
# 被当成「表头元信息的保护套 P/N」，8 条一起被元信息专线抢走，去 pdf_meta 里找
# kind=保护套 —— 那 11 条元信息全是机器编号，一条都对不上，于是逐条报
# 「PDF 表头里没有「保护套」这一项」。用户看到的是 9 条全错，根因就是这里。
# 边界：这两个词必须**收尾**（可跟括号注释 / 标点），或后面紧跟「编号 / P/N /
# 号 / 型号 / 规格 / 料号」才算元信息；夹在句子中间的「球囊保护套切定长」不算。
_META_BARE_KWS = ("保护套", "芯轴")
_META_BARE_TAIL_RE = re.compile(
    r"(?:保护套|芯轴)(?:[（(][^）)]*[）)]|[：:、,，.。/／\-—_]|[\s\u3000\u00a0])*$")
_META_BARE_NEXT_RE = re.compile(
    r"(?:保护套|芯轴)(?:[（(][^）)]*[）)])?\s*[:：]?\s*"
    r"(?:编号|料号|物料|图号|型号|规格|p/?n|pn|no\.?|号)", re.IGNORECASE)

def _meta_bare_ok(s, kw):
    """光名词（保护套 / 芯轴）在名字里的这个位置，算不算「就是这一项」。"""
    start = 0
    while True:
        i = s.find(kw, start)
        if i < 0:
            return False
        if _META_BARE_TAIL_RE.fullmatch(s[i:]) or _META_BARE_NEXT_RE.match(s[i:]):
            return True
        start = i + 1


def _meta_kind_of(item):
    """Excel 采集项名 → 元信息类别；不是元信息返回 None。

    '球囊近端焊接 机器编号' → '机器编号'
    '程序号'                → '程序号'
    '焊机温度(℃)：'          → None（普通参数，走原先的路径）
    '应力管移印 机器：'       → '机器编号'（见下：名字里没有「编号」二字）
    '使用模板将球囊保护套切定长，长度' → None（「保护套」夹在句子里，不是元信息）
    """
    s = _flat(item)
    if not s:
        return None
    for kw, kind in _META_ITEM_KINDS:
        k = _flat(kw)
        if k not in s:
            continue
        # 光名词（保护套 / 芯轴）要么收尾、要么紧跟编号类字样，见 _META_BARE_KWS
        if kw in _META_BARE_KWS and not _meta_bare_ok(s, k):
            continue
        return kind
    # WI-3264 的两项叫「应力管移印 机器 / 应力管切割 机器」——**没有「编号」二字**，
    # 上表的关键词一个都不命中，于是被当成普通参数；而它的值是 `T-1653`（设备号、
    # 非数字），普通路径取不出数值 → 458×2 = 916 条全落 novalue「Excel 没填值」。
    # 事实是 Excel 明明填着 T-1653、PDF p3~p9 的设备清单里也写着 T-1653 —— 纯冤案。
    # 只认**收尾**的「机器/设备/机台」：真参数要么带单位括号（已于上一行剥掉后再判）、
    # 要么收在数字/物理量上，不会以这几个词结尾。反过来，单字「机」结尾的
    # 「焗炉机 / 喷码机 / 焊接机」**不吃** —— 它们是设备本身的名字，语义比这窄。
    s2 = re.sub(r"[（(][^）)]*[）)]\s*$", "", s)
    if _META_EQUIP_TAIL_RE.search(s2):
        return "机器编号"
    return None


def _pdf_meta_kind(key, val=None):
    """PDF 表头里的键名 → 元信息类别。

    先查字典，查不到再按**语义**兜底：键看着像设备类、值里确实是设备号
    → 按「机器编号」收。只补字典治标不治本 —— WI-0026 用的是「机器编号」，
    换一份文件可能又写「焊机编号」「机台」…，每一次漏都是整台机器查无此项。
    """
    k = _PDF_META_KEY_FLAT.get(_flat(key))
    if k:
        return k
    if val is not None and _is_equip_meta(key, val):
        return "机器编号"
    return None


def _meta_basis(sub, dia, who):
    """元信息的出处文字：附表N（设备）（直径档）。"""
    parts = []
    if sub:
        parts.append(sub)
    if who:
        parts.append(who)
    if dia:
        parts.append("φ%s-%smm" % (dia[0], dia[1]) if dia[0] != dia[1]
                     else "φ%smm" % dia[0])
    return "·".join(parts) if parts else "PDF 表头"


def _equip_list_from_pdf(pdf_path):
    """从「设备/工具编号」设备清单表里抽设备号（WI-0011 型版式）。

    WI-0011 的参数表是 A 型竖表 + 千分尺表，表头里没有设备号，
    设备单独列在一张「NO | 设备/工具编号 | 设备/工具名称」清单里。
    不读这张清单，它的「折叠球囊 机器编号」就永远无从核对。
    """
    out = []
    try:
        with open_pdf(pdf_path) as pdf:
            for page in pdf.pages:
                try:
                    tables = page.find_tables() or []
                except Exception:
                    continue
                for t in tables:
                    try:
                        rows = t.extract() or []
                    except Exception:
                        continue
                    # 表头行**未必在表格开头三行**：WI-0016 p3 那张「工具设备明细表」
                    # 前面垫了 5 行物料明细表，只看 rows[:3] 会把它整张漏掉 ——
                    # 焗炉机 T-1030/T-1079/T-1601/T-10527 一条都进不来。
                    # 判据收紧成「**某一格**里同时有设备/工具 与 编号」：
                    # 「设备参数附表1」+「文件编号」分处两格，参数附表就不会再被误当清单。
                    hidx = None
                    for hi, hr in enumerate(rows):
                        for c in (hr or []):
                            cj = str(c or "")
                            if "编号" in cj and ("设备" in cj or "工具" in cj):
                                hidx = hi
                                break
                        if hidx is not None:
                            break
                    if hidx is None:
                        continue
                    for row in rows[hidx:]:
                        for c in (row or []):
                            # 「T- 1601」（T- 和数字夹空格）与「T-10455/10456」
                            # （后几台省略前缀）都要接住，跟 `_equip_set` 同一套规矩
                            s = re.sub(r"(T-)\s+", r"\1", _flat(str(c or "")))
                            for e in equip_from_text(s):
                                if e not in out:
                                    out.append(e)
    except Exception:
        pass
    return out


def extract_pdf_meta(pdf_path):
    """抽取参数表表头的元信息块，作为可核对的「参数」。

    来源两处，都已在版式层抽好，这里只做「聚合成项 + 去重 + 带适用范围」：
      ① 横表（B 型）表头的键值块：激光机编号 / 程序号 / 焊接保护套 / 使用芯轴
      ② 程序附表（C 型）表头的「程序号/名称 N」

    返回 [{kind, value, sub, dia, equip, equip_raw, families, page, basis}]
      kind  规范化类别（机器编号 / 程序号 / 保护套 / 芯轴）
      value 原文（可能是 'T-1114（65）/T-1281（50）' 这种合并值，留给比对层拆）
      dia   该子表的直径档（机器编号在不同直径档下不同，必须带上才能收窄）
    """
    out, seen = [], set()

    def add(kind, value, sub="", dia=None, equip=None, title="", page=0, who="",
            fams=None):
        value = _clean_meta(value)
        if not kind or not value:
            return
        eq = list(equip or [])
        key = (kind, _flat(value), sub or "", tuple(dia or ()), tuple(eq))
        if key in seen:
            return
        seen.add(key)
        out.append({
            "kind": kind, "value": value,
            "sub": sub or "", "dia": tuple(dia) if dia else None,
            "equip": eq, "equip_raw": who or value,
            "families": (set(fams) if fams is not None
                         else families_of(" ".join([sub or "", title or ""]))),
            "page": page,
            "basis": _meta_basis(sub, dia, who or value),
        })

    # ① 横表表头键值块（版式层已按子表解析好，这里聚合去重）
    for r in _stepwise_all(pdf_path):
        for k, v in (r.get("meta") or {}).items():
            kind = _pdf_meta_kind(k, v)
            if kind:
                add(kind, v, r.get("sub") or "", r.get("dia"), r.get("equip"),
                    r.get("title") or "", r.get("page") or 0, r.get("equip_raw") or "")

    # ② 程序附表的「程序号/名称 N」（附表没有直径档，但有设备与海波管 DWG 可收窄）
    prog_recs = _programs_all(pdf_path)
    for r in prog_recs:
        pn = r.get("prog_no_raw")
        if pn:
            # fams 必须透传：附表标题里的「For 产品线F NC」是**同一台设备多张附表**
            # 唯一的区分线索，不带出来就两套程序号并存（见 `_program_page`）。
            add("程序号", pn, r.get("sub") or "", None, r.get("equip"),
                "", r.get("page") or 0, r.get("equip_raw") or "",
                fams=r.get("families"))

    # ③ 兜底：竖表 / 程序附表里的「设备/工具编号」→ 机器编号
    #
    # 横表型参数表（WI-0021/2105）的表头写着「激光机编号」，还带 65/89 这种机器序号，
    # 那是最强的核对依据，所以只在「本份 PDF 一条机器编号都没抽到」时才打这条兜底。
    #
    # 竖表型（WI-0007/2012/2015/2016）与附表型（WI-3852）没有那一行，
    # 它们的设备号写在每个参数块前面的「设备/工具编号：T-2057/T-1254」里。
    # 不兜底的话，这几个 WI 的「XX 机器编号」永远是「PDF 表头里没有该栏」——
    # 等于换了个说法的「不核对」。
    #
    # 这类记录的核对语义 = 「Excel 填的设备号是否真的出现在 PDF 里」，
    # 所以收窄链在这里是合适的：er.equip 收窄后若为空，就说明这台设备在 PDF 里查无此机。
    if not any(r["kind"] == "机器编号" for r in out):
        seen_eq = set()
        for r in _vertical_all(pdf_path):
            eqs = tuple(r.get("equip") or ())
            if not eqs or eqs in seen_eq:
                continue
            seen_eq.add(eqs)
            for e in eqs:
                add("机器编号", e, r.get("sub") or "", r.get("dia"), eqs,
                    "", r.get("page") or 0, "/".join(eqs))
        seen_sub = set()
        for r in prog_recs:
            raw = r.get("equip_raw")
            if not raw or raw in seen_sub:
                continue
            seen_sub.add(raw)
            add("机器编号", raw, r.get("sub") or "", None, r.get("equip"),
                "", r.get("page") or 0, raw)

    # ④ 「设备/工具编号」设备清单表（WI-0011 型）。
    #    它的参数表既没有横表表头块、也没有带设备号的竖表标题，
    #    设备只列在一张独立清单里。这一层只能做「存在性」核对
    #    （这台机在不在 PDF 的设备清单里），所以标 weak，界面会如实降档。
    #
    #    原来是「一条机器编号都没抽到才兜底」。WI-0016 揭了这个条件的短：
    #    p4 参数表里带设备号的是**拉伸机**，于是 ③ 已经出了一批机器编号，
    #    兜底被整段跳过 —— 而 p3 清单里的**焗炉机** T-1030/T-1079/T-1601/T-10527
    #    明明在 PDF 里，Excel 却报「PDF 表头里没有该设备」（假漏判）。
    #    改成**补齐式**：只补前面几层没覆盖到的设备号，不重复、不动已有结论。
    have_eq = set()
    for r in out:
        if r["kind"] == "机器编号":
            have_eq.update(r.get("equip") or ())
    for e in _equip_list_from_pdf(pdf_path):
        if e in have_eq:
            continue
        n = len(out)
        add("机器编号", e, "", None, [e], "", 0, "PDF 设备清单表")
        if len(out) > n:
            out[-1]["weak"] = True
            have_eq.add(e)
    return out


# ------------------------------------------------- 抽取端自检（换新 PDF 时先看这个）

_SUB_WS = re.compile(r"[\s\u3000\u00a0]+")


def _norm_sub(s):
    """子表编号归一：`附表1-a` / `表 1－a` / `表1-a` → `表1-A`。

    三个抽取路径对同一张子表的叫法并不一致（竖表给「附表1-a」、
    横表给「表1-1」），不归一就会把同一张表报成「文本里有、却没抽到」。
    """
    t = _SUB_WS.sub("", str(s or ""))
    for dash in ("－", "—", "–", "﹣"):
        t = t.replace(dash, "-")
    t = re.sub(r"^附表", "表", t)
    return t.upper()


# 子表编号：允许编号内部有少量空格（PDF 里常见「表 1 - a」），但**不能跨行**。
# 早期写法是先去掉整页空白再匹配，于是「表1-1 1 SPC…」被粘成「表1-11」，
# 凭空多出一个不存在的子表编号（WI-0009 的「表1-11」「表1-113」就是这么来的）。
_SUB_MENTION = re.compile(r"表[ \t\u00a0]*(\d+)[ \t\u00a0]*[-－—–﹣][ \t\u00a0]*([0-9A-Za-z]+)")


def _subs_mentioned(pdf_path):
    """页面文本里「提到过」的子表编号 → {归一编号: {页…}}。

    只认带后缀的 `表N-x`：表格正文里「见表1参数要求」这类泛指不算子表编号，
    「设备参数附表2」这类**章节名**也不是子表编号 —— 混进来只会天天误报。
    """
    out = {}
    with open_pdf(pdf_path) as pdf:
        for pno, page in enumerate(pdf.pages, 1):
            t = page.extract_text() or ""
            for a, b in _SUB_MENTION.findall(t):
                out.setdefault(_norm_sub("表%s-%s" % (a, b)), set()).add(pno)
    return out


def _subs_extracted(pdf_path):
    """四条抽取路径实际给出的子表编号（归一后）。"""
    out = set()
    for rs in (_stepwise_all(pdf_path), _vertical_all(pdf_path), _programs_all(pdf_path)):
        for r in rs:
            if r.get("sub"):
                out.add(_norm_sub(r["sub"]))
    try:
        mic = extract_pdf_micrometer(pdf_path)
    except Exception:
        mic = {}
    for k in ("open", "pass", "nopass"):
        for r in (mic.get(k) or []):
            if r.get("sub"):
                out.add(_norm_sub(r["sub"]))
    return out


def self_check(pdf_path):
    """抽取端自检：只报「结构性疑点」，不判对错、不猜值。

    为什么需要它：抽取结果偏小时，界面只会显示「无对应」—— 到底是 PDF 里
    真没有，还是我们没读出来？这两种情况后果完全不同，不能都让人自己猜。

    判据的口径很关键：**「没抽到」指那张表的行没读到，而不是标签没记上**。
    竖表路径经常拿不到子表编号（表名与数据行被 PDF 拆成不同的表格对象），
    只是标签缺了、数据全在，那不算漏。

    复用已缓存的三路抽取结果，不额外解析 PDF。
    返回 {"lines": [行…], "suspect": 疑点数}
    """
    L, suspect = [], 0
    try:
        sw = _stepwise_all(pdf_path)
        vt = _vertical_all(pdf_path)
        pg = _programs_all(pdf_path)
        try:
            mic = extract_pdf_micrometer(pdf_path)
        except Exception:
            mic = {}
        n_mic = sum(len(mic.get(k) or []) for k in ("open", "pass", "nopass"))
        with open_pdf(pdf_path) as pdf:
            n_pages = len(pdf.pages)
    except Exception as e:
        return {"lines": ["抽取自检失败：%s: %s" % (type(e).__name__, e)],
                "suspect": 1, "hollow_subs": [], "gap_subs": [],
                "missed_pages": [], "pages_total": 0, "pages_with_points": 0}

    L.append("参数点：横表 %d / 竖表 %d / 程序附表 %d / 千分尺·针规 %d"
             % (len(sw), len(vt), len(pg), n_mic))

    # 每页实际产出的记录数（四条路径合并 —— 漏掉千分尺那一路，
    # WI-0011 的对照表页会被误判成「有数字却无参数点」）
    bypage = {}
    for rs in (sw, vt, pg):
        for r in rs:
            p = r.get("page")
            if p:
                bypage[p] = bypage.get(p, 0) + 1
    for k in ("open", "pass", "nopass"):
        for r in (mic.get(k) or []):
            p = r.get("page")
            if p:
                bypage[p] = bypage.get(p, 0) + 1

    # ① 子表编号对账 —— 只报「编号所在的页一条记录都没有」的
    mention = _subs_mentioned(pdf_path)
    got = _subs_extracted(pdf_path)
    # 只比「本文档自己用的那种写法」：WI-0009 的子表叫「附表1…附表14」（无后缀），
    # 正文里「见设备参数设置附表1-11」是**区段引用**（附表1 到 11），不是子表编号；
    # 用带后缀的写法去对账就会凭空多出一个不存在的子表。
    dashed = any("-" in s for s in got)
    hollow = []
    for sub, pages in sorted(mention.items()):
        if sub in got:
            continue
        if got and not dashed:
            continue                      # 本文档的编号不带后缀 → 带后缀的引用不算
        if any(bypage.get(p, 0) for p in pages):
            continue                      # 页里有别的记录：标签没记上，但数据读到了
        hollow.append((sub, sorted(pages)))
    if mention and not hollow and (got or not mention):
        L.append("子表编号对账：文本提到 %d 个，所在页都有参数点" % len(mention))
    elif hollow:
        suspect += 1
        L.append("有 %d 个子表所在页没有任何参数点：%s"
                 % (len(hollow), "、".join("%s(p%s)" % (s, "/".join(map(str, ps)))
                                          for s, ps in hollow[:6])))

    # ② 横表：同一子表内步骤号必须连续 —— 缺号就是整列没读出来
    bysub = {}
    for r in sw:
        bysub.setdefault((r.get("sub") or "—", r.get("equip_raw") or "—"), []).append(r)
    gaps = []
    for (sub, eq), rs in bysub.items():
        steps = sorted({r["step"] for r in rs if r.get("step") is not None})
        if len(steps) >= 2 and steps != list(range(steps[0], steps[-1] + 1)):
            gaps.append((sub, eq, steps))
    if gaps:
        suspect += 1
        for sub, eq, st in gaps[:6]:
            L.append("步骤号不连续：%s（%s）→ %s" % (sub, eq, st))
    elif bysub:
        L.append("横表步骤号连续性：%d 组全部连续" % len(bysub))

    # ③ 有「参数表表头」却一条记录都没有的表 —— 真漏读
    try:
        missed = _missed_param_tables(pdf_path, bypage)
    except Exception:
        missed = []
    if missed:
        suspect += 1
        for pno, hits, head in missed[:8]:
            L.append("p%d 有参数表表头（%s）却无参数点：%s"
                     % (pno, "、".join(hits), head[:44]))

    # ④ 图片块 —— **不算疑点**，但必须当场说出来（2026-11-02 用户实案）
    #
    # WI-0039 第 4 页的表7（球囊保护套切定长、扩口）整块是图片，文字层一个字
    # 都没有，8 条采集项因此判「PDF 里没有这一项」，而表7 上明明印着
    # 35/60/100/140/170/220/270/330mm，和 Excel 完全一致。
    # 这类「没读出来」跟「文件真没有」后果完全不同，不能让人自己去猜。
    #
    # 为什么不进 suspect：图片块本身不是抽取故障 —— 工程示意图（WI-0039 第 7 页
    # 那三张 L1/L2/L3 拉伸示意）也是图片，它们不含参数。把它算成疑点，每份
    # 带插图的 WI 都会常亮一行警告，警告就不值钱了。这里只作**事实陈述**。
    try:
        img_blocks = pdf_image_blocks(pdf_path)
    except Exception:
        img_blocks = {}
    img_note = image_blocks_note(img_blocks)
    if img_blocks:
        L.append("图片块：%s —— 图里的字没有文字层，工具读不到"
                 "（若某条「找不到对应」的答案是印在图里的，只能人工看图）"
                 % img_note)

    # 判词放第一行：原来整段【抽取自检】压在报告中段，而人只会看最上面一句，
    # 等于没有。这句话回答的是「下面那张表值不值得信」，比一致/不一致的数字
    # 更该先看到。
    #
    # 刻意**不写**「覆盖 N/M 页」：这些 WI 的操作流程页、质量要求页本来就没有
    # 参数点，拿页数当覆盖率会天天喊狼来了。真正该保证、也真能保证的是
    # 「凡是有参数表表头的页，都读到了东西」—— 所以判词只承诺这一句。
    if suspect:
        bits = []
        if hollow:
            bits.append("子表缺口 %d 处" % len(hollow))
        if gaps:
            bits.append("步骤断号 %d 组" % len(gaps))
        if missed:
            bits.append("整页漏读 %d 页" % len(missed))
        verdict = ("抽取自检：本文档有 %d 处疑点（%s）—— 下面「无对应」里"
                   "可能混有「没读出来」的项，先按本段去 PDF 里核这几处"
                   % (suspect, " ｜ ".join(bits)))
    else:
        verdict = "抽取自检：有参数表表头的页都读到了，未发现结构性疑点"
    if img_blocks:
        # 「✓ 未发现疑点」这句话在文字层这一层是对的，但用户看到的是「9 条全对不上」。
        # 图里读不到的时候必须补这半句，否则那句 ✓ 会让人以为 PDF 里真没有。
        verdict += ("；但%s印着图片块，图里的字读不到" % img_note)
    L.insert(0, verdict)
    return {"lines": L, "suspect": suspect,
            # 结构化疑点：顶部那句交叉提醒要点名「可疑的是哪几处」，
            # 光靠拼字符串没法可靠地取出这些名字。
            "hollow_subs": [s for s, _ps in hollow],
            "gap_subs": [s for s, _eq, _st in gaps],
            "image_pages": sorted(img_blocks),
            "image_note": img_note,
            "missed_pages": [p for p, _h, _d in missed],
            "pages_total": n_pages,
            "pages_with_points": len(bypage)}


def _missed_param_tables(pdf_path, bypage):
    """页里有「参数表表头」却一条记录都没产出的表 —— 真漏读的唯一可靠信号。

    不能用「含多少个数字格」来判：WI-0007 的操作流程页里有十几处温度/时间数字，
    但那是正文，不是参数表；拿数字格数当判据会天天误报。
    也不能按「整页关键词」排除非参数页：这些 WI 每页往往只有一整张大表，
    参数段和「质量要求」段在同一张表里，按关键词排除会把参数表一并免检 ——
    WI-0011 p6「折叠参数设定」整块漏读就再也发现不了了。

    所以认表头：参数名称 / 参数·Step / 设定值。有表头却没产出，才是漏。
    """
    out = []
    with open_pdf(pdf_path) as pdf:
        for pno, page in enumerate(pdf.pages, 1):
            if bypage.get(pno, 0):
                continue
            try:
                tables = page.find_tables() or []
            except Exception:
                continue
            for t in tables:
                try:
                    rows = t.extract() or []
                except Exception:
                    continue
                hits = []
                for row in rows:
                    if _v_header_cols(row):
                        hits.append("参数名称")
                    if _step_head_cols(row):
                        hits.append("参数/Step")
                    # 千分尺/针规表的表头：必须**认单元格**且是短标签。
                    # 全表搜字符串会被正文骗到 —— WI-0011 p5「校准Fold Head」的
                    # 说明里写着「将一个相应直径的 Pin gauge 插入」，那页并没有
                    # 对照表，按全表搜就会误报成「有表头却无参数点」。
                    for c in (row or []):
                        k = _flat(c).lower()
                        if not k or len(k) > 24:
                            continue
                        if "settingvalue" in k or k.startswith("设定值"):
                            hits.append("设定值")
                        elif "pingage" in k or k.startswith("针规"):
                            hits.append("针规")
                if not hits:
                    continue
                head = ""
                for r in rows[:4]:
                    s = " ".join(str(c) for c in (r or []) if c)
                    if s.strip():
                        head = re.sub(r"\s+", " ", s)
                        break
                out.append((pno, sorted(set(hits)), head))
                break                      # 一页报一处就够
    return out


def inspect_pdf(pdf_path):
    """PDF 版式自检：返回 (横表记录, 报告文本)。"""
    return extract_stepwise(pdf_path), report_layout(pdf_path)


def pdf_text_warning(pdf_path):
    """PDF 文字质量提示（字体映射失真/解析失败）；没问题返回空串。"""
    msgs = []
    try:
        w = text_quality_warning(pdf_path)
        if w:
            msgs.append(w)
    except Exception:
        pass
    if LAST_PDF_FAILED:
        msgs.append("有 %d 页解析失败（已跳过）：%s"
                    % (len(LAST_PDF_FAILED),
                       "；".join("p%d %s" % (p, str(e)[:40]) for p, e in LAST_PDF_FAILED[:4])))
    return "  ".join(msgs)


def pdf_scan_warning(pdf_path):
    """扫描件检测：每页只有图片、几乎无文字 → 一句话提示；正常 PDF 返回空串。

    2026-09-27 起因：用户导入了 WI-0001/2002 的扫描件 PDF（每页一整张图、0 字符），
    detect_wi 读不出 WI 号 → 界面静默退回内置的 WI-0008，用户看着像程序出错。
    本工具从设计上**不做 OCR**（只读 PDF 自带文字层）—— 扫描件必须明说，
    绝不能让用户在「读不出来」和「文件没有」之间猜。
    判据：前 5 页文字合计 < 20 字符、且每页都有图片。
    放在引擎侧（而不是 app 里）是因为批量核对将来也要用同一句判词。
    """
    try:
        with open_pdf(pdf_path) as pdf:
            n = min(5, len(pdf.pages))
            chars, imgs = 0, 0
            for i in range(n):
                pg = pdf.pages[i]
                chars += len((pg.extract_text() or "").strip())
                imgs += len(pg.images)
        if n and chars < 20 and imgs >= n:
            return ("这份 PDF 没有文字层（前 %d 页每页只有图片、读不出任何字符）—— "
                    "是扫描件或拍照导出的 PDF。本工具不做 OCR，参数一个都读不出来，"
                    "核对无法进行。请换带文字层的电子版 PDF（导出为 PDF 而非打印扫描）；"
                    "若只有扫描件，需先经 OCR 转成带文字层的 PDF 再核对。" % n)
    except Exception:
        pass
    return ""


# --------------------------------------------------------------- Excel 抽取

def list_sheets(xlsx_path):
    """列出工作簿里的全部 sheet 名（供界面下拉框使用）。"""
    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def inspect_excel(xlsx_path, sheet=None):
    """Excel 结构自检：返回 (TableSchema, 报告文本)。

    只描述结构（sheet / 表头行 / 列名 / 类型统计），不含任何数据值，可安全外发。
    """
    ts = TableSchema.load(xlsx_path, sheet=sheet)
    return ts, ts.report()


def audit_excel(xlsx_path, sheet=None):
    """表格自检（不依赖 PDF）：返回 [AuditResult] 问题清单。

    与 inspect_excel 的分工：那个回答「这张表我能不能读」，这个回答
    「这张表自己说的话前后对不对得上」。检查项见第 3 节开头。
    sheet=None 时把所有含采集项名称的明细表都自检一遍。
    """
    return _audit_xlsx(xlsx_path, sheet=sheet)


def audit_coverage(xlsx_path, pdf_paths=(), scan_dir=None):
    """覆盖面 + WI 版本链（跨文件）：返回 CoverageReport。

    与 audit_excel 的区别：那个只看一份 Excel 内部自不自洽，这个看
    「数据和文件配不配得上」——尤其能把「有录入数据但没有 PDF 可核」
    和「数据本身不对」这两件事分开报，这两件事以前在界面上长得一模一样。
    """
    return _audit_workbook(xlsx_path, pdf_paths=pdf_paths, scan_dir=scan_dir)


def scan_excel_outliers(xlsx_path, sheet=None, min_group=4, minor_ratio=0.25):
    """同规格横向一致性扫描：返回 OutlierReport。

    只报「组内存在明显多数」的差异 —— 势均力敌的两组更像两套并行工艺，不下结论。
    「N/A / 无 / -」这类「不适用」标记不参与比较。
    """
    return _scan_outliers(xlsx_path, sheet=sheet,
                          min_group=min_group, minor_ratio=minor_ratio)


def extract_excel_params(xlsx_path: str, wi_name: str = "WI-0008",
                         sheet: str | None = None,
                         with_report: bool = False
                         ) -> list[dict] | tuple[list[dict], TableSchema]:
    """抽取采集项，返回 [{card,proc,group,item,unit,type,lo,hi,fixed,btrue,bfalse,spec,equip,wi}]。

    结构自适应：用哪张 sheet、表头在第几行、每个字段在第几列，全部由第 1 节
    按【列名语义】推断 —— 换一份全新的表也不用改代码。
    sheet=None 自动挑选采集项表；指定 sheet 名可强制使用某张表。
    with_report=True 时返回 (records, TableSchema)，便于界面显示结构自检信息。
    """
    ts = TableSchema.load(xlsx_path, sheet=sheet)
    records = ts.records(wi_name)
    return (records, ts) if with_report else records


# --------------------------------------------------------------- 千分尺查找表

def _num(s):
    s = (s or "").replace('"', "").replace("”", "")
    m = re.search(r"\d+(?:\.\d+)?", s)
    return float(m.group(0)) if m else None


# 「直径列」一格可能有几段（`3.0mm；3.5x200-300mm`、`0.8mm&1.0mm`）。
# 逗号故意**不**当分隔符：`1.0,1.25mm` 拆开后会多记一条 1.25，反而重复。
_DIA_SEG_SPLIT = re.compile(r"[；;、&\n\r]+")
# 直径单元：开头（可带 φ），可跟 mm，可再跟「x 长度[-长度]mm」。
# 末尾的 mm 可有可无 —— PDF 里 `5.0x20-150mm` 和 `5.0x20-150` 都见过。
_DIA_ONE = re.compile(
    r"^[φΦ]?\s*(\d+(?:\.\d+)?)\s*(?:mm)?\s*"
    r"(?:[xX×*＊]\s*(\d+(?:\.\d+)?)\s*(?:[-~—–]\s*(\d+(?:\.\d+)?))?\s*mm?)?")


def _dia_cells(cell):
    """把「直径列」的一格拆成 [(直径, 长度下限, 长度上限), …]。

    为什么要拆、要认长度（2026-10-30，用户报的实案）：

    旧写法是 `re.compile(r"^(\\d+(?:\\.\\d+)?)mm")` —— 数字后面必须**紧跟 mm**。
    WI-0043 第 6 页针规表的第一列长这样：

        1.25mm / 1.5mm / 2.0mm / 2.5mm / 3.0mm；3.5x200-300mm / 3.5x20-150mm
        4.0mm / 5.0x20-150mm / 5.0x200-300mm / 6.0x20-150mm / 6.0x200-300mm

    十二行里，凡是写成「直径x长度」的六行**一个都没认出来**（`5.0x20-150mm`
    的数字后面是 `x` 不是 `mm`）。于是 5.0 的卡在表里找不到 φ5.0 档，
    走了「按不大于该直径的最大档推算」，拿到 **φ4.0** 的值：

        Excel 填 0.035（5.00 X 20）→ 工具说「应为 0.03」   ← 0.030 是 4.0mm 的「过」
        Excel 填 0.036（5.00 X 20）→ 工具说「应为 0.031」  ← 0.031 是 4.0mm 的「不过」

    这不是读数错，是**查错了行**。用户原话：「0.035，它变成了它说成 0.01」
    「0.031」——三条都是这一处。

    顺带两件事：
      · 一格塞好几段（`3.0mm；3.5x200-300mm`）要拆开，否则 3.5 那一档永远缺；
      · `直径x20-150mm` 里的「20-150」是**长度档**，5.0 有 20-150 和 200-300
        两段、值还不一样（0.035 / 0.036），只按直径选行会把两段混在一起，
        所以长度要一起带出来，交给 `_micrometer_expected` 收窄。
    """
    s = str(cell or "")
    if not s.strip():
        return []
    out = []
    for seg in _DIA_SEG_SPLIT.split(s):
        seg = seg.strip()
        if not seg:
            continue
        m = _DIA_ONE.match(_norm(seg))
        if not m:
            continue
        dia = m.group(1)
        lo = m.group(2)
        hi = m.group(3)
        if lo is None:
            out.append((dia, None, None))
        else:
            lo_f = _as_float(lo)
            hi_f = _as_float(hi) if hi is not None else lo_f
            if lo_f is None:
                out.append((dia, None, None))
            else:
                out.append((dia, min(lo_f, hi_f), max(lo_f, hi_f)))
    return out


def _is_col_head(cell):
    """这一格像不像「列名」—— 短标签，而不是一整段正文。

    针规表的真表头是 `Pass/过` / `No Pass/不过`（≤10 字）。
    正文段落里也会出现「过」（「折叠过程中无夹伤」），那种一定是长句、
    带句读号。用它把两种情况分开，比维护一张词表稳。
    """
    t = _norm(cell)
    if not t or len(t) > 16:
        return False
    return not any(ch in t for ch in "。，、；;！!？?：")


def extract_pdf_micrometer(pdf_path):
    """抽取千分尺/针规查找表，**每条都带适用范围**。

    返回 {"open": [rec…], "pass": [rec…], "nopass": [rec…], "manual": set()}
    rec = {"dia","val","equip","sub","families","page","len"}
    `len` 是长度档 (lo, hi) 或 (None, None) —— 同一根径按长度分两档时值不一样，
    见 `_len_fit` 上面那段实案说明。

    为什么必须带适用范围：WI-0011 p9 一张表里并排 + 上下堆着 表2-a/2-b/2-c/2-d
    四段，同样 2.0mm 的「过」在 表2-a（产品线A 线）是 0.026″、
    在 表2-b（产品线C 线）是 0.029″ —— 只按直径合并就会给出两个候选，
    把值填成隔壁产品线的也判「一致」。老实现还有一处更隐蔽的错误：
    本行没有值时沿用上一行的值，于是 2.25mm 被安上了 2.0mm 的 0.026″。

    ⚠️ 原来这里有个 `DIA = re.compile(r"^(\\d+(?:\\.\\d+)?)mm")`，要求数字后紧贴 mm，
    于是针规表里写成「直径x长度」的六行（3.5x20-150mm、5.0x200-300mm …）一个都
    认不出来 —— φ5.0 收不到档，退到 φ4 推算，Excel 填的 0.035″ 就被判成
    「应为 0.03″」。2026-10-30 换成 `_dia_cells()`，那个常量随之删除。
    """
    open_recs, pass_recs, nopass_recs, manual = [], [], [], set()
    with open_pdf(pdf_path) as pdf:
        for pgno, page in enumerate(pdf.pages, 1):
            try:
                text = page.extract_text() or ""
                if "Setting Value" not in text and "针规" not in text and "Pin Gage" not in text:
                    continue
                # ⚠️ 上面的「有没有针规」是**页**级的，只用来快速跳过无关页；
                # 到底这张表是针规表还是「设定值」表，必须逐表判断（2026-10-30）：
                # WI-0043 第 6 页同时摆着「Setting Value/设定值」表（表1）和
                # 「Pin Gage/针规」表（表2），页级 is_gage=True 会把设定值那张
                # 也按针规处理 —— 它没有「过/不过」列，于是整张表一条都抽不出来，
                # 「Pleat 头开口时千分尺设定值」永远报「没有这类查找表」。
                for t in page.find_tables():
                    rows = t.extract()
                    width = max((len(r) for r in rows), default=0)
                    if width < 2:          # 2 列也接：WI-0043 表1 就是「直径 | 设定值」
                        continue
                    # 只有 ≥4 列才可能是「左右并排两张表」，才切段。
                    # 3 列的针规表（直径|过|不过）切段会把自己切坏：
                    # mid=1 会把「过」那一列的 0.024” 当成新的直径行，
                    # 于是凭空多出 dia='0.024' 的假记录（2026-10-30 修）。
                    mid = width // 2 if width >= 4 else 0
                    st_list = (0, mid) if mid else (0,)
                    sides = {s: {} for s in st_list}  # 左半 / 右半各有一套适用范围
                    pcols, ncols, vcols = [], [], []
                    for row in rows:                  # 先只定位列（表头列位置是全表通用的）
                        for c, cell in enumerate(_clean(c) for c in row):
                            k = _norm(cell).lower()
                            # ⚠️「过/不过」这一支**必须再确认这一格是列名**，不能纯子串。
                            # WI-0043 第 3 页质量要求写着「…在折叠过程中无夹伤。…」，
                            # 纯 `"过" in k` 命中 → 整张物料明细表被当成针规表，
                            # 于是 `NO` 列（1/2/3/4）既当直径又当值，凭空造出
                            # dia=4 / val=4.0 这种记录，混进 φ4.0 的候选里，
                            # 界面上「应为」就成了「0.03、4」（2026-11-02 实案）。
                            if "pass" in k or "过" in k:
                                if not _is_col_head(cell):
                                    continue
                                (ncols if ("不过" in k or "nopass" in k) else pcols).append(c)
                            elif "settingvalue" in k or "设定值" in k:
                                vcols.append(c)
                    # 表级判断：这张表自己有没有「过/不过」列。
                    # 有 → 针规表；没有但有「设定值」列 → 那把尺的设定值表。
                    # 两者都没有 → 这张表不参与，跳过（不猜）。
                    is_gage = bool(pcols or ncols)
                    if not is_gage and not vcols:
                        continue
                    # 取值必须与「读适用范围」在同一遍里做：一张表里堆了
                    # 表2-a/2-b 再 表2-c/2-d 两段，先把范围全算完再取值的话，
                    # 后半段的附表编号会盖到前半段的数据上（0.026″ 被记成 表2-d）。
                    for row in rows:
                        cells = [_clean(c) for c in row]
                        for start in st_list:             # 左右两段各自的适用范围
                            stop = mid if (mid and start == 0) else width
                            seg = cells[start:stop]
                            txt = _flat(" ".join(c for c in seg if c))
                            if not txt:
                                continue
                            if "设备编号" in txt or "设备/工具编号" in txt:
                                sides[start]["equip"] = equip_from_text(txt)
                            m = re.search(r"表\s*(\d+\s*[-－]\s*[0-9A-Za-z])", txt)
                            if m:
                                sides[start]["sub"] = "表" + re.sub(r"\s+", "", m.group(1))
                            fam = families_of(txt)
                            if fam:
                                sides[start]["families"] = sorted(fam)
                        for start in st_list:
                            if start >= len(cells):
                                continue
                            # 一格可能写着好几段（`3.0mm；3.5x200-300mm`），
                            # 也可能带长度档（`5.0x20-150mm`）。逐段各出各的记录，
                            # 值都一样 —— 见 `_dia_cells` 上面那段实案说明。
                            dias = _dia_cells(cells[start])
                            if not dias:
                                continue
                            st = sides[start]
                            pc = nc = vc = None
                            if is_gage:
                                pc = next((c for c in sorted(pcols) if c >= start), None)
                                nc = next((c for c in sorted(ncols) if c >= start), None)
                                p = _num(cells[pc]) if pc is not None and pc < len(cells) else None
                                np_ = _num(cells[nc]) if nc is not None and nc < len(cells) else None
                                if p is None and np_ is None:
                                    continue      # 本行空着 = 该直径不适用，绝不沿用上一行
                            else:
                                vc = next((c for c in sorted(vcols) if c >= start), None)
                                raw = cells[vc] if vc is not None and vc < len(cells) else ""
                                if not raw or "manual" in _norm(raw).lower() or "手工" in raw:
                                    if raw:
                                        for (d0, _a, _b) in dias:
                                            manual.add(d0)
                                    continue
                                v = _num(raw)
                                if v is None:
                                    continue
                            for (dia, len_lo, len_hi) in dias:
                                if is_gage:
                                    base = {"dia": dia, "equip": st.get("equip") or [],
                                            "sub": st.get("sub", ""),
                                            "families": st.get("families") or [], "page": pgno,
                                            "len": (len_lo, len_hi)}
                                    if p is not None:
                                        pass_recs.append(dict(base, val=p))
                                    if np_ is not None:
                                        nopass_recs.append(dict(base, val=np_))
                                else:
                                    open_recs.append({"dia": dia, "val": v,
                                                      "equip": st.get("equip") or [],
                                                      "sub": st.get("sub", ""),
                                                      "families": st.get("families") or [],
                                                      "page": pgno,
                                                      "len": (len_lo, len_hi)})
            except Exception as e:      # 单页解析失败不拖垮整份 PDF
                _note_pdf_failure(pgno, e)
    return {"open": open_recs, "pass": pass_recs, "nopass": nopass_recs, "manual": manual}


def _dp_hint(val_dp, actual):
    """取该值在 PDF 里的精度签名，交给界面把「2.0」如实显示成 2.0。

    2026-10-30 用户点名（截图：光强 1/2 判「小数位数与 PDF 不符」）：
    「PDF 内是 1.5--2.0，但是这个显示方面出错了」——判定没错，是界面把
    应为值 1.5~2.0 显示成了 1.5~2，跟 Excel 填的 1.5~2 长得一模一样，
    看的人根本看不出差在哪，只能信那句「小数位数不符」。
    """
    dps = (val_dp or {}).get(actual) or set()
    for p in dps:
        if p and any(x is not None for x in p):
            return p
    return None


def _micrometer_kind(item):
    if "开口" in (item or ""):
        return "open"
    if "不过" in (item or ""):
        return "nopass"
    if "过" in (item or ""):
        return "pass"
    return None


def _diameter_from_spec(spec):
    m = re.search(r"(\d+(?:\.\d+)?)", spec or "")
    return m.group(1) if m else None


# 参数卡的「规格」列写成「直径 X 长度 压力cm」三段，例：
#   '5.00 X 20 150cm' / '5.00 X 200 150cm' / '1.50 X 40 90cm'
# 第一个数是球囊直径（现行 `_diameter_from_spec` 取的就是它），
# 第二个数是长度 —— 针规表里同一根径按长度分两档，值不一样（见下），
# 所以第二个数必须单独取出来，否则 φ5.0 会同时拿到 0.035″ 与 0.036″。
_SPEC_LEN_RE = re.compile(
    r"^\s*[φΦ]?\s*\d+(?:\.\d+)?\s*(?:mm)?\s*[xX×*＊]\s*(\d+(?:\.\d+)?)")

# 新口径的球囊长度：`QL(20-150)` / `QL（200-300）` / `QL200`。
# 括号可有可无、全角半角都认、区间符号认 - ~ － — 至（PDF 抽出来的写法不统一）。
_QL_LEN_RE = re.compile(
    r"QL\s*[（(]?\s*(\d+(?:\.\d+)?)\s*(?:[-~－—至]\s*(\d+(?:\.\d+)?))?", re.I)


def _length_from_spec(spec):
    """参数卡的「适用规格」→ 球囊长度段 `(下界, 上界)`；认不出就 None（不猜）。

    两种写法都要认（2026-11-02 补第二种）：

        '5.00 X 20 150cm'      → (20, 20)     老口径：第二个数是长度（单点）
        'φ5.0xQL(20-150)'      → (20, 150)    新口径：QL 里本来就是区间
        'φ5.0xQL200'           → (200, 200)   新口径单点

    ⚠️ 为什么必须认 QL：WI-0043 的针规表同一根径按长度分两档
    （φ5.0 + 20-150 → 0.035″，φ5.0 + 200-300 → 0.036″）。参数卡写的是
    `φ5.0xQL(200-300)`，老正则只认 `X 数字`，遇到 `QL` 直接返回 None，
    于是**长度档收窄整段跳过**，两档都留在候选里 —— 界面上「应为」就出现
    两个值（0.035、0.036），Excel 填哪个都算对。这就是 11-02 用户看到的
    「下面的应为值有两个」。

    返回区间后由 `_len_span_overlap` 按**重叠**判，不再要求落在档内某一点。
    """
    s = str(spec or "")
    qls = set()
    for m in _QL_LEN_RE.finditer(s):
        lo = float(m.group(1))
        hi = float(m.group(2)) if m.group(2) else lo
        qls.add((lo, hi))
    if len(qls) == 1:
        return next(iter(qls))
    if len(qls) > 1:
        # 一张卡写了两段不同的长度（例如「φ3.5xQL(20-150),φ3.5xQL(200-300)」）
        # —— 到底按哪段查，工具不知道，弃权比猜安全。
        return None
    m = _SPEC_LEN_RE.match(s)
    if m:
        v = float(m.group(1))
        return (v, v)
    return None


def _span_text(span):
    """长度段 (lo, hi) → 人话：`20-150` / `20`（上下界相同就只说一个数）。"""
    if not span:
        return "?"
    lo, hi = span
    if lo is None and hi is None:
        return "?"
    if lo is None or hi is None:
        return str(lo if hi is None else hi)
    if abs(float(lo) - float(hi)) < 1e-9:
        return _fmt_num(lo)
    return "%s-%s" % (_fmt_num(lo), _fmt_num(hi))


def _len_span_overlap(span, want):
    """记录的长度档 `span` 与规格给的长度段 `want` 是否**重叠**。三态返回。

    与 `_len_fit` 的区别只有一个：`want` 现在是区间（QL 写法），
    所以判「重叠」而不是「这个点落在档里」。单点 `(20, 20)` 自然退化成老行为。

      True  —— 有重叠；
      False —— 明确不重叠，可以排掉；
      None  —— 有一侧没有长度信息，**不能**当成不匹配丢掉。
    """
    if not span or not want:
        return None
    lo, hi = span
    a, b = want
    if None in (lo, hi, a, b):
        return None
    return not (hi < a - 1e-9 or lo > b + 1e-9)


def _len_fit(span, L):
    """记录带的长度档 `span=(lo, hi)` 套不套得住长度 L。

    三态返回，调用方必须按三态处理：
      True  —— 该长度落在这条记录的长度档里；
      False —— 明确不在这条记录的档里；
      None  —— 这条记录没有长度档（整张表只有一种长度，如 φ1.25mm 那几行），
               此时**不能**当成不匹配丢掉，否则 1.25/1.5/2.0/2.5/3.0/4.0
               六档全部丢失。
    """
    if not span:
        return None
    lo, hi = span
    f = _as_float(L)
    if lo is None or hi is None or f is None:
        return None
    return (lo - 1e-9) <= f <= (hi + 1e-9)


# ------------------------------------------------- 步骤式横表匹配（WI-0021 型）

_STEP_ITEM_RE = re.compile(r"^\s*(?:步骤|step)\s*(\d+)\s*[的]?\s*", re.I)


# 第二种写法：**参数名/步骤号(单位)** —— WI-0020（DWG-0000009 rev08 球囊成型）那套
# Excel 就是这么写的：'压力/4(Bar)'、'加热中间模/6(℃):ref'、'充气速度/15'。
# 只认「斜杠后面紧跟数字、再可选一个单位括号与 :ref 后缀」这一种形状 ——
# 双语名 'Test Pressure/测试压力' 斜杠后面不是数字，不会被误伤。
#
# ⚠️ 单位括号必须**半角全角都认**（R47）：WI-0033 / 2463 / 2490 / 2518 这四个新
# 语料的 Excel 采集项名一律写全角 —— '加热中间模/10（℃）'、'压力/14（Bar）'。
# 只认半角的话 step 全是 None → 步骤链不启动 → 整批掉进「按参数名兜底」，
# 应为值退化成整列并集（实测 WI-0036 101 条、WI-0038 137 条无法匹配）。
_STEP_ITEM_BRACKET = r"(?:\([^()（）]*\)|（[^()（）]*）)"
_STEP_ITEM_SLASH_RE = re.compile(
    r"^(.*?)\s*/\s*(\d+)\s*" + _STEP_ITEM_BRACKET + r"?\s*"
    r"(?:[:：]\s*ref)?\s*$", re.I)


def parse_step_item(name):
    """'步骤2的激光能量' → (2, '激光能量')；'压力/4(Bar)' → (4, '压力')。

    Excel 侧用「步骤N的XX」描述，PDF 横表用「参数 | Step N」描述，
    这个拆分是把两边对上的钥匙。

    ⚠️ 第二种写法（参数名/步骤号）是 2026-09-29 补的。缺了它，WI-0020 的 486 行
    里 191 行的 step 全是 None —— `_step_expected` 第一步就返回 None，
    判定链压根没启动（界面上一律「无法匹配」，而 PDF 其实读得好好的）。
    """
    s = str(name or "")
    m = _STEP_ITEM_RE.match(s)
    if not m:
        m2 = _STEP_ITEM_SLASH_RE.match(s)
        if m2 and m2.group(1).strip():
            return int(m2.group(2)), m2.group(1).strip()
        return None, s
    return int(m.group(1)), s[m.end():]


def _spec_range(spec):
    """Excel「适用规格」→ (lo, hi)。'7.0' → (7.0, 7.0)；'2.0-4.0' → (2.0, 4.0)。

    ⚠️ 规格里常常**同时**写着直径和长度：`φ4.0xQL(10-60)mm` 是「直径 4.0 ×
    长度 10-60mm」。旧口径取「第一个区间」，拿到的是**长度** (10,60) —— 拿它去
    比直径档 (2.0,4.0) 得分 0，整张卡被判 no_spec_match。
    WI-0020 的 0124/0128/0129 三张卡就是这么全军覆没的（2026-09-29 实测）。
    所以**先认直径那一段**：φ/Ø 之后，遇到 `x`/`×`/`*`/`QL`/`,` 为止。
    认不出来（规格里没写 φ）就完全退回旧口径 —— 语料里那 12 份的行不受影响。
    """
    s = str(spec or "")
    m = re.search(r"[φΦ⌀Øø]\s*([0-9][0-9.\-~～,]*)", s)
    if m:
        seg = re.split(r"[x×*]", m.group(1))[0]
        seg = re.split(r"[,，]", seg)[0]
        m2 = re.search(r"(\d+(?:\.\d+)?)\s*[-~～]\s*(\d+(?:\.\d+)?)", seg)
        if m2:
            a, b = float(m2.group(1)), float(m2.group(2))
            return (a, b) if a <= b else (b, a)
        m2 = re.search(r"(\d+(?:\.\d+)?)", seg)
        if m2:
            v = float(m2.group(1))
            return (v, v)
    m = re.search(r"(\d+(?:\.\d+)?)\s*[-~～]\s*(\d+(?:\.\d+)?)", str(spec or ""))
    if m:
        a, b = float(m.group(1)), float(m.group(2))
        return (a, b) if a <= b else (b, a)
    m = re.search(r"(\d+(?:\.\d+)?)", str(spec or ""))
    if not m:
        return None
    v = float(m.group(1))
    return (v, v)


def _dia_score(spec, dia):
    """规格区间与子表直径范围的贴合度。

    2 = 规格完全落在该子表的直径范围内（最可靠）
    1 = 有重叠但没被完全包含（唯一命中时只能算「近似」）
    0 = 无关
    """
    if not spec or not dia:
        return 0
    a, b = spec
    c, d = dia
    if c - 1e-9 <= a and b <= d + 1e-9:
        return 2
    if min(b, d) - max(a, c) >= -1e-9:
        return 1
    return 0


def _spec_ranges(s):
    """Excel「适用规格」里的**全部**数字区间。

    `_spec_range` 只取第一个区间，够用于"这台设备的直径档位"；
    但 WI-3264 的规格是**复合**的 —— `φ(4.5-5.0)xQL(5-60)` 同时含直径区间
    (4.5,5.0) 和长度区间 (5,60)，只认第一个就会把 `xQL(5-60)` 和 `xQL(80-200)`
    混成同一档（两档的切割长度是 2.5 和 3.5，判错就是漏检）。
    """
    out = set()
    for a, b in re.findall(r"(\d+(?:\.\d+)?)\s*[-~～]\s*(\d+(?:\.\d+)?)", str(s or "")):
        a, b = float(a), float(b)
        out.add((a, b) if a <= b else (b, a))
    return out


def _range_match_score(need, have):
    """「Excel 规格里的每个数字区间都能在这一行里找到」才算对得上。

    这是 WI-3264 p10 表2 的判定口径（用户 2026-09-28 亲口定的）：
    中间列「产品规格」写着 `Minerva\\AchillesØ1.5-4.0和Ø4.5-5.0×5-60，产品线G\\CastorØ3.5-8.0`
    这类天书，**字母和标点认不全没关系，只看数字区间对不对得上**。
    一行里往往有**好几组**区间（不同产品线各一组），所以不能用单区间比。

    2 = Excel 的每个区间都能在这一行里找到（完全对上）
    1 = 只有部分区间对上（近似）
    0 = 一个都对不上
    """
    if not need or not have:
        return 0
    need, have = set(need), set(have)
    if need <= have:
        return 2
    for a, b in need:
        for c, d in have:
            if a - 1e-9 <= d and c - 1e-9 <= b:
                return 1
    return 0


def _fmt_sub(key):
    """(子表编号, (lo, hi)) → '表1-1(2.0-5.5mm)' 这样的可读标签。"""
    sub, dia = key
    name = sub or "未标号"
    if not dia:
        return name
    a, b = dia
    if a == b:
        return "%s(%.1fmm)" % (name, a)
    return "%s(%.1f-%.1fmm)" % (name, a, b)


def _nostep_basis(res, tail):
    """no_step 的说明文字。规格可能同时对应多张子表，措辞要随张数变。"""
    if res.get("many"):
        head = ("PDF %s 这几张参数表里都没有这一步（这些档只有步骤 %s）"
                % (res.get("sub") or "该设备", "、".join(str(s) for s in res.get("steps") or ())))
    else:
        head = ("PDF %s 的参数表里没有这一步（该档只有步骤 %s）"
                % (res.get("sub") or "该设备", "、".join(str(s) for s in res.get("steps") or ())))
    return head + tail


def _noparam_basis(res, er, tail):
    """no_param 的说明：该档的参数表里**没有这个采集项**。

    与 `_nostep_basis` 的区别要说清楚，否则看的人会以为是两种「没有」：
      · no_step   = 有这个参数，只是没有这一步
      · no_param  = 连这个参数都没印（比前者更靠前）

    依据里把**这张表实际印了哪些参数**列出来：那一串就是「不是我漏读，
    是表里确实没有这一列」的证据（用户 2026-09-30：「加热左端膜和右端膜
    这两个词是没有的，在 PDF 里面没有出现」）。
    """
    subs = res.get("sub") or "该档参数表"
    have = [p for p in (res.get("params_in") or ())]
    item = (er or {}).get("item") or "该采集项"
    if res.get("many"):
        head = "PDF %s 这几张参数表里都没有「%s」这一项" % (subs, item)
    else:
        head = "PDF %s 的参数表里没有「%s」这一项" % (subs, item)
    if have:
        head += "（这些表里印的参数是：%s）" % "、".join(str(p) for p in have[:12])
        if len(have) > 12:
            head += " 等 %d 项" % len(have)
    return head + tail


def _nospec_basis(res, er):
    """no_spec_match 的说明：规格在这台设备下对不上任何一档。"""
    # 末尾那句不能省：规格套不上档时，工具**答不出**「PDF 有没有这一步」，
    # 于是既不判对也不判错。但看的人需要知道「这种情况下什么算对」——
    # 否则他会以为 Excel 填 N/A 有问题，白白去改（用户 2026-09-30 的规矩：
    # PDF 没这一步 → Excel 空着或填 N/A 都对）。
    return ("Excel 的「适用规格 %s」在这台设备的参数表里对不上任何一档"
            "（该设备在 PDF 里只有 %s）—— 期望值收不窄，请核对规格是否写错。"
            "若确认该规格在 PDF 里本就没有这一档（该档没有这一步），"
            "则 Excel 留空或填 N/A 都是对的"
            % (er.get("spec") or "?", res.get("sub") or "未知档位"))


def _ordinal_target(res):
    """序位近似最终对到了哪张表（可读标签）。"""
    sub, dia = res.get("sub"), res.get("dia")
    if sub and dia:
        return _fmt_sub((sub, tuple(dia)))
    return sub or "该档参数表"


def _mark_ordinal(chk, res):
    """序位近似命中的行：判定强度单列一档，依据里写明是按什么对过去、对到哪张表。

    为什么要单列一档：这一行的「应为」不是按规格算出来的，是按**位次**推的。
    混进「硬判定」里会让「一致性有多可信」这个数虚高 —— 用户要的是
    「用作判定」，不是「假装它也是精确的」。
    """
    if not (res or {}).get("ordinal"):
        return
    chk["scope"] = "ordinal"
    chk["sub"] = _ordinal_target(res)
    note = ("Excel 规格列与 PDF 分档整体对不上（该设备两边档数相同、而多数规格连"
            "沾边的档都没有），按「从小到大第 N 个规格 ↔ 第 N 档」近似对到 %s —— "
            % _ordinal_target(res))
    chk["basis"] = note + (chk.get("basis") or "")


def _ordinal_maps(excel_records, pdf_steps, pdf_step_blanks):
    """序位近似的兜底映射：{设备: {规格区间: (子表编号, 直径)}}。

    Excel 的「适用规格」和 PDF 的分档是两套口径：
      Excel 写的是产品尺寸族（2.0-4.0 / 4.5-6.0 / 7.0 / 8.0）
      PDF  分的是设备参数档（2.0-5.5 / 6.0 / 7.0 / 8.0）
    正常情况能一一套上（规格落进某一档的区间）。但整份表格偶尔会**大面积**对不上
    （换了口径、或写的是别的产品线）。这时逐行按区间去套，行行都套不上，
    收窄链就只剩「该设备所有同名值的并集」——把隔壁档的参数当成期望值，
    凭空造出一个「应为」。

    兜底规则（用户 2026-09-20 定）：两边都按从小到大排，**第 N 个规格配第 N 档**。
    只在「大面积对不上」时启用；只有个别几行套不上，那是人为写错，
    宁可如实报出来，也不能用位次蒙混过去。

    启用条件（三条同时满足）：
      ① 该设备在两边都有梯队，且**档数相同**（不然谈不上一位对一位）；
      ② 档数 ≥ 2（一档时「位次」没有意义）；
      ③ 至少三分之二的规格在 PDF 里连沾边的档都没有（_dia_score 全 0）。
    """
    dias = {}
    for r in list(pdf_steps or []) + list(pdf_step_blanks or []):
        if not r.get("dia"):
            continue
        k = (r.get("sub") or "", tuple(r["dia"]))
        for t in (r.get("equip") or ()):
            dias.setdefault(t, set()).add(k)

    specs = {}
    for er in excel_records or ():
        eq = _equip_set(er.get("equip"))
        sp = _spec_range(er.get("spec"))
        if not eq or not sp:
            continue
        for t in eq:
            specs.setdefault(t, set()).add(sp)

    out = {}
    for t, sp_set in specs.items():
        d_set = dias.get(t)
        if not d_set:
            continue
        sp_l = sorted(sp_set)
        d_l = sorted(d_set, key=lambda k: (k[1][0], k[1][1]))
        if len(sp_l) < 2 or len(sp_l) != len(d_l):
            continue
        hit = sum(1 for s in sp_l if any(_dia_score(s, k[1]) >= 1 for k in d_l))
        if hit * 3 > len(sp_l):        # 多数仍能精确套上 → 不启用，保持精确为主
            continue
        out[t] = {s: k for s, k in zip(sp_l, d_l)}
    return out


def _ordinal_for(ordinal, eq, spec):
    """按行里挂的设备逐个查兜底映射，命中返回 (子表编号, 直径)，否则 None。"""
    if not ordinal or not eq or not spec:
        return None
    hit = None
    for t in sorted(eq):
        m = (ordinal.get(t) or {}).get(spec)
        if m is None:
            continue
        if hit is not None and hit[1] != m[1]:
            return None        # 一行挂了多台设备、映射结果还打架 → 不用，宁可不判
        hit = m
    return hit


def _param_gap(sub_params, ck, eq, spec):
    """「这台设备这个规格档的参数表里，到底有没有这个采集项？」—— 第 53 轮新增。

    返回 None            = 问不出来（没有子表信息 / 没设备 / 规格套不上任何一档 /
                           至少有一张表印了这个参数）
    返回 (top, params_in) = 该设备该档对应的子表**都没有**这个参数；
                           params_in 是这些子表实际印了哪些参数（写依据用）

    ⚠️ 这是 `_step_gap`（「有没有这一步」）的**上位问题**：连参数都没有，
    更谈不上步骤。必须**先问它** —— 只问步骤会漏掉一整类情况：

      WI-0036 card 0008（Ø4.0~5.0，设备 T-1177 那 7 台）的「加热左/右端模/6」：
      「加热端模」这个参数**只印在 Ø6.0 的表2/表3 上**（那是另一批设备的档），
      Ø4.0~5.0 的表1 里三个部位只写了「加热中间模」，压根没有端模这一列。
      Excel 在这一栏填 N/A 是**对的**（用户 2026-09-30 点名：「加热左边左端膜和
      右端膜这两个词是没有的，在 PDF 里面没有出现，所以就应该去匹配 NA」）。

      可改之前工具在 `no_spec_match`（规格套不上档）就停手了 —— 那个档位是
      「我不知道该看哪张表」，不是「PDF 里没有」。于是这条被挂成「无法匹配」，
      界面上看着像待办，而它本来是对的。

    判据比 `_step_gap` 更严：**必须规格精确套上某一档**（`_dia_score >= 1`）才
    敢下结论。套不上就是「不知道该看哪张表」，维持 `no_spec_match` 不表态 ——
    「我不知道」绝不冒充「PDF 里没有」（这条铁律见 `_nospec_basis`）。
    """
    if not sub_params or not ck or not eq or not spec:
        return None
    _any_equip = any(v.get("equip") for v in sub_params.values())
    owned = ([(k, v) for k, v in sub_params.items() if (v["equip"] & eq)]
             if _any_equip else list(sub_params.items()))
    if not owned:
        return None                     # 这台设备在 PDF 的参数表里根本没出现
    scored = [(_dia_score(spec, k[1]) if k[1] else 1, k, v) for k, v in owned]
    best = max(s for s, _k, _v in scored)
    if best < 1:
        return None                     # 规格套不上任何一档 → 不知道该看哪张表
    top = [(k, v) for s, k, v in scored if s == best]
    if any(ck in v["params"] for _k, v in top):
        return None                     # 有任意一张印了这个参数 → 不能说「没有」
    return (top,
            sorted({p for _k, v in top for p in v["params"]}))


def _step_gap(step_tables, ck, step, eq, spec):
    """「这台设备的这张参数表，有没有这一步？」—— 步骤数的**结构性问题**。

    返回 None           = 问不出来（没有子表信息 / 没设备 / 至少有一张子表有这一步）
    返回 (top, steps_in) = 该设备（该规格）对应的子表**都没有**这一步；
                          steps_in 是这些子表实际有的步骤号（写说明文字用）

    为什么要单独拎出来：`_step_expected` 收窄链的第一步是「设备对齐」，而它用的
    pool 只装了**有这一步**的记录。设备若恰好属于「没有这一步」的那张子表，
    在 pool 里永远找不到它，于是先撞出 no_equip_match —— 设备明明在 PDF 里，
    只是这张表没有这一步，却被报成「PDF 里没有这台设备」，语义完全错了。

    打分口径与 `_dia_score` 一致，只多一条：标题里不带 (φ…mm) 的子表
    （如 WI-0024「表1-3 设备/工具编号：T-1281尖端成型参数表」）没有直径档可收窄，
    设备既已对上就按命中计 —— 否则这类表整段被跳过，步骤数永远判不出来。

    **规格栏不是数字时（`ALL` / 空 / 「全部」）不能在门口就放弃。**
    Excel 里大批卡片的「适用规格」写的是 `ALL`（小样本表尤其如此）。
    旧代码把「没规格」当成「问不出来」直接 return None，于是这类行一律滑到
    下游的 `novalue（未填值）`：WI-0024 表4 的尖端成型-02（T-1281，表1-3 只有
    Step 1）与尖端成型-04（T-1362/T-10058/T-10161，表1-5 只有 Step 1），
    步骤2 五格全被挂成「待办」，而 Excel 填 N/A 本来就是对的。
    没有规格能收窄时，退到**设备这一层**问同一个问题：
    这台设备名下的子表全都没有这一步 → 结论依然唯一。
    只要有一张写了这一步就不表态（保守），交给下游按规格/设备继续收窄。
    """
    if not step_tables or not ck or step is None or not eq:
        return None
    tables = step_tables.get(ck) or {}
    # ⚠️ 这份 PDF 的子表**一张都没印设备号**时，按设备过滤这个动作本身无从谈起
    # —— 与 `_step_expected` 里规格收窄同一口径（R50）：全部子表都当候选。
    # 不做这一步，WI-0036 表2（Ø6.0，只写了步骤 6~10 的加热左/右端模）就会因为
    # 「表里没印设备号」而整段跳过，于是「有没有步骤 4/5」这个问题永远答不出来，
    # Excel 填 N/A 的 12 条只能挂成「未填值」。
    _any_equip = any(v.get("equip") for v in tables.values())
    owned = ([(k, v) for k, v in tables.items() if (v["equip"] & eq)]
             if _any_equip else list(tables.items()))
    if not owned:
        return None                     # 这台设备在本参数的子表里根本没出现
    if spec:
        scored = [(_dia_score(spec, k[1]) if k[1] else 1, k, v) for k, v in owned]
        best = max(s for s, _k, _v in scored)
        # ⚠️ 规格套不上任何一档时**不要直接放弃**：那只说明「我不知道该看哪张表」，
        # 而「这一步存不存在」有时跟看哪张表无关 —— 该参数在 PDF 里**一共就一张
        # 子表、且那张表压根没这一步**，结论就是唯一的（WI-0036 card 0008 的
        # 「加热左/右端模/4、5」：这个参数只出现在表2(Ø6.0) 的步骤 6~10，
        # φ4.0-5.0 那张卡要步骤 4，PDF 里确实没有 —— Excel 填 N/A 是对的）。
        # 所以降级到设备这一层再看一遍，与下面「没规格」那条路同一口径。
        top = ([(k, v) for s, k, v in scored if s == best] if best >= 1
               else owned)
    else:
        # 规格无从收窄（ALL / 空）→ 停在设备这一层，看这台设备名下全部子表
        top = owned
    if any(step in v["steps"] for _k, v in top):
        return None                     # 有任意一张写了这一步 → 不能说「没有」
    return top, {s for _k, v in top for s in v["steps"]}


def _no_step_res(gap):
    """把 `_step_gap` 的答案包成 `_step_expected` 的返回结构。"""
    top, steps_in = gap
    return {"vals": set(), "scope": "no_step", "blank": False,
            "sub": "、".join(_fmt_sub(k) for k, _v in top),
            "dia": None, "many": len(top) > 1, "ordinal": False,
            "steps": sorted(steps_in), "by_equip": {}}


def _no_param_res(gap):
    """把 `_param_gap` 的答案包成 `_step_expected` 的返回结构。

    `params_in` 一并带出去：依据里要能说出「这张表里印的是哪些参数」，
    让人一眼看出「不是我漏读，是表里确实没有这一列」。
    """
    top, params_in = gap
    return {"vals": set(), "scope": "no_param", "blank": False,
            "sub": "、".join(_fmt_sub(k) for k, _v in top),
            "dia": None, "many": len(top) > 1, "ordinal": False,
            "steps": [], "params_in": params_in, "by_equip": {}}


# ==========================================================================
#  第 2.8 节  收窄链（Narrowing Chain）—— 一张级表 + 三级各自的顺序
#
#  一条链 = 「按 Excel 这一行的适用范围，逐级把 PDF 候选格收窄到那一档」。
#
#  为什么要有这一节：2026-09-30 重构前，这条管道在引擎里**写了三套** ——
#    · 严格链：内联在 `_step_expected`（350 行），主判定路径；
#    · 形态链：`_narrow_recs`，只用来定「PDF 这一格是什么形态」；
#    · 兜底链：`_fallback_cand`，严格链判不出来时才介入。
#  三套的**级顺序与判据各不相同**，于是同一行可能因为走哪套而得到不同答案
#  —— 这正是「同一个问题换个文件又冒出来」的结构性根因。
#
#  这一节做的是：① 把每一级抽成独立函数（判据全项目只有一份）；
#  ② 用下面这张表声明「每条链用哪几级、按什么顺序」。
#
#  ⚠️ 铁律（每一级都必须遵守）：**收得动才收**。
#     收不动就原样返回候选 —— 少收一级只是退回上一版行为，
#     绝不因为收窄失败而丢掉候选，更不许凭空造结论。
#
#  ⚠️ 表里的顺序**不是**随便排的：判据越宽，位置必须越靠后。
#     宽判据排在前面会「抢走」窄判据的表（第 48 轮为这个付出过 5310 条
#     凭空变 no_match 的代价）。改这张表之前先问：它会不会抢别人的表。
# ==========================================================================

# 级名 → (中文名, 实现函数)。函数签名一律 (er, recs) -> recs
# （严格链的「设备对齐」是个例外：它收不动时要能定案，走 `_equip_gate`）。
NARROW_STAGES = {
    "unit":    ("单位对齐（mm/s 不与 mm 混）", "_stage_unit"),
    "cond":    ("一格多分支：按适用规格挑最贴合的那一支", "_stage_cond"),
    "equip":   ("设备号", "_equip_gate / _narrow_recs_by_equip"),
    "diam":    ("直径档", "_dia_score"),
    "qlen":    ("球囊长度（直径并列时才用）", "_qlen_keys"),
    "ordinal": ("序位近似（规格套不上档时的兜底）", "_ordinal_for"),
    "keyword": ("关键词 / 设备号", "_apply_fallback 内联"),
    "range":   ("整组区间匹配", "_range_match_score"),
}

# 每条链用哪几级、按什么顺序。**顺序即执行顺序**。
NARROW_CHAINS = {
    "strict":   ("unit", "cond", "equip", "diam", "qlen", "ordinal"),
    "shape":    ("equip", "diam", "qlen"),
    "fallback": ("keyword", "equip", "diam", "range"),
}


def _stage_unit(er, recs):
    """单位对齐：mm/s 不与 mm 混 —— 收窄链的**第一级**（只有严格链有）。

    ⚠️ 空格子必须留：它没有单位，任何单位下都算对得上。不留的话，
    「按这条 Excel 行的适用范围，PDF 对应的那一格是不是空的」就问不出来
    （空值格被单位层先剔掉 → 下游误判成「PDF 没有这一格」）。
    ⚠️ PDF 侧没写单位的记录也留 —— 那是「PDF 没写」，不是「对不上」。
    """
    unit = _excel_unit(er)
    if not unit:
        return recs
    keep = [r for r in recs
            if r.get("empty") or not r.get("unit") or _norm_unit(r["unit"]) == unit]
    return keep or recs


def _stage_cond(er, recs):
    """一格多分支：按 Excel「适用规格」挑条件最贴合的那一支。

    PDF 里同一格常常写两套值、各自带生效条件：
    `24bar for 4.0*10-60 / 22bar for 2.0-3.5 &4.0*80-150`。
    不挑支的话两支一并集，Excel 填对也报不一致。

    ⚠️ **一支都贴不上就都留着**，让下游标成 union 交人定 —— 绝不替人选。
    ⚠️ 没写条件的记录一律留：它们不受分支影响。
    """
    conds = [r for r in recs if r.get("cond")]
    if not conds:
        return recs
    spec_txt = str(er.get("spec") or "")
    scored = [(_enstep_cond_score(r["cond"], spec_txt), r) for r in conds]
    best = max(s for s, _r in scored)
    if best <= 0:
        return recs
    keep = [r for r in recs
            if not r.get("cond")
            or _enstep_cond_score(r["cond"], spec_txt) >= best - 1e-9]
    return keep or recs


def narrow_chain_report():
    """把三条链用哪几级打出来 —— 给「版式自检」看，别再另写一套硬编码。"""
    out = []
    for chain in ("strict", "shape", "fallback"):
        out.append("  %-9s %s" % (
            chain, " → ".join("%s(%s)" % (NARROW_STAGES[k][0], k)
                              for k in NARROW_CHAINS[chain])))
    return out


def _selfcheck_declarations():
    """两张声明表（NARROW_STAGES / NARROW_CHAINS / MATCH_TIER_RULES）的自检。

    表驱动重构的最大风险是「表写对了、可它指的那份实现没跟上」—— 改名、删函数、
    拼错，Python 不会报错，只会在运行到那一步时静默走样。所以这里把表里每一条
    声明回查一次：级名→实现函数在不在、链名→级名在不在、每条匹配规则的三件套
    （hit/score/why）是不是都能调。

    返回 (错误列表, 警告列表)。空错误 = 表与实现对齐。
    """
    import inspect
    errs, warns = [], []

    # ① NARROW_STAGES：每个级标注的实现函数必须真实存在
    for stage, (label, impl) in NARROW_STAGES.items():
        # 实现名可能是「_a / _b」双写，或带「内联」说明，拆出纯函数名逐个查
        fns = re.findall(r"_[\w]+", impl)
        if not fns:
            errs.append("NARROW_STAGES[%r] 的实现串 %r 里抽不出函数名" % (stage, impl))
            continue
        for fn in fns:
            if not callable(globals().get(fn)):
                errs.append("NARROW_STAGES[%r] 标注的函数 %r 不存在" % (stage, fn))

    # ② NARROW_CHAINS：每条链引用的级名必须在 NARROW_STAGES 里
    for chain, stages in NARROW_CHAINS.items():
        for s in stages:
            if s not in NARROW_STAGES:
                errs.append("NARROW_CHAINS[%r] 引用了未定义的级 %r" % (chain, s))

    # ③ MATCH_TIER_RULES：每条规则 hit/score/why 都必须是可调用的三件套
    for i, r in enumerate(MATCH_TIER_RULES):
        for attr in ("hit", "score", "why"):
            if not callable(getattr(r, attr, None)):
                errs.append("MATCH_TIER_RULES[%d]（%s）的 %s 不是可调用对象"
                            % (i, r.name, attr))
        if not isinstance(getattr(r, "tier", None), int):
            errs.append("MATCH_TIER_RULES[%d]（%s）缺 tier 层号" % (i, r.name))
        # 层号必须落在 _FB_NAME 里，否则 match_tier_report 会 KeyError
        if r.tier not in _FB_NAME:
            errs.append("MATCH_TIER_RULES[%d]（%s）的层号 %r 不在 _FB_NAME 里"
                        % (i, r.name, r.tier))

    # ④ 每条链至少有一级，且 fallback 链以 keyword 开头（分层匹配是后置 pass）
    if any(not v for v in NARROW_CHAINS.values()):
        warns.append("存在空的收窄链：%s" % [k for k, v in NARROW_CHAINS.items() if not v])

    return errs, warns


def _step_expected(er, step_index, step_tables=None, blank_index=None,
                   ordinal=None):
    """按「参数名 + 步骤号」在横表里找期望值，并逐层收窄。

    返回 dict（找不到参数返回 None）：
      vals     期望值集合（scope 不是 strict/approx/union*/device 时为空集）
      blank    PDF 该格是**明确写空**（`------`）—— 与「没这一步」不同：
               空值是表格在声明「此处无值」，Excel 填 N/A 或不填都对，
               填了值反而错。
      scope    收窄到的层级
               strict  唯一确定（规格精确对齐，或该设备下只有一套值）
               approx  唯一子表，但规格与直径只是「有重叠」，属近似
               union   规格区间横跨多张子表，且**各表取值有分歧** —— 期望值
                       只能取并集，属宽松判定（候选真的可能不止一个）
               union_same  同样横跨多张子表，但这几张这一格写的是同一个值
                       —— 用哪一张答案都一样，期望值事实上唯一，算**硬判定**。
                       与 union 必须分开：否则「现值 480~520 / 候选 480~520」
                       这种自我重复的行会混进「宽松判定明细」，纯噪音。
               device  只能收到设备这一层，该设备下仍有多套参数（宽松）
               no_step 该规格对应的子表里根本没有这一步（如 表1-1(2.0-5.5mm)
                       只有 Step1/2，Excel 却有「步骤3的…」）—— 此栏应填 N/A。
                       规格可能同时压住多张子表（Excel 的适用规格与 PDF 的分档
                       边界对不齐），此时看的是「这几张里有没有一张写了这一步」，
                       一张都没有才算 no_step。
               pdf_blank 该规格对应的子表里**有**这一步，但格子里写的是 `------`
               none    Excel 没给设备号，收不窄 —— 此时 vals 故意为空
               no_equip_match   PDF 里没有这台设备的参数
               no_spec_match    Excel 的「适用规格」在这台设备下对不上任何一档
                               （规格写错口径、或写的不是这一页的参数）——
                               期望值收不窄，**不再**退化成「该设备所有同名值并集」
      ordinal  这一行的档位是**按位次**对出来的（序位近似，见 _ordinal_maps），
               而不是按规格算出来的 —— 判定强度要单列一档，不能混进硬判定
      by_equip 按设备分组的候选值，用于「说不清时如实展示」

    为什么 scope=none 时必须返回空集：以前会把整份 PDF 的并集当成期望值显示，
    于是 T-1175 的期望值里冒出了 T-1114 的 460-530，直接误导人。

    为什么需要 no_step：横表是「按直径分档」的，不同档的步骤数不一样
    （T-1175：2.0-5.5mm 档只有 2 步，7.0mm 档有 3 步，8.0mm 档有 4 步）。
    只看「有步骤N记录的表」的话，2.0mm 规格的候选里只会剩 7.0/8.0 档的
    步骤3值 —— 等于把隔壁档的值当成了期望值。必须先看「该规格落在哪张表」，
    再问「那张表有没有这一步」。
    """
    step, base = parse_step_item(er.get("item"))
    if step is None:
        return None
    ck = _canonical_param(base)
    if not ck:
        return None

    cand = step_index.get((ck, step)) or []
    blanks = (blank_index or {}).get((ck, step)) or []
    if not cand and not blanks:
        # PDF 里**任何一张表**都没有「该参数 + 该步骤」这一格。
        #
        # ⚠️ 这不是「无从核对」，而是一条**硬证据** —— 必须先问「这台设备（该档）
        # 的参数表到底有没有这一步」（`_step_gap`），答得出来就是「PDF 没这一步」，
        # 按用户 2026-09-30 的规矩 Excel 留空 / 填 N/A **都对**。
        # 旧版在这里直接 `return None`，于是这 12 条（WI-0036 card 0008/0012/0015
        # 的「加热左/右端模/4、5」：表2 只有步骤 6~10）一律落进「未填值」，
        # 界面上看着像「这里漏了什么」，而 Excel 填 N/A 本来就是对的。
        gap = _step_gap(step_tables, ck, step,
                        _equip_set(er.get("equip")), _spec_range(er.get("spec")))
        return _no_step_res(gap) if gap is not None else None
    # 空格子（`------`）也必须进收窄链：只有让它和有值格一起按设备/规格过滤，
    # 才能回答「按这条 Excel 行的适用范围，PDF 对应的那一格是不是空的」。
    # 走的是**同一条**收窄链，不另写一套规则 —— 两套收窄必然发散。
    pool = cand + blanks

    # ---- 先问「结构性」的问题：这台设备的参数表里，到底有没有这一步？----
    #
    # 必须放在「设备对齐」的硬失败**之前**。原因：pool 只装了**有这一步**的记录，
    # 若这台设备恰好属于「没有这一步」的那张子表，pool 里永远找不到它，
    # 下面会先返回 no_equip_match（假阴性：设备明明在 PDF 里，只是没有这一步），
    # no_step 根本没有机会表态。
    #
    # 实测（WI-0024）：表1-3「设备/工具编号：T-1281尖端成型参数表」只有 Step 1，
    # 而 Excel 每个卡片都有「步骤2的激光能量(mW)」，该栏填 N/A —— 这是**对的**。
    # 旧代码在这里判出 no_equip_match，界面把它归进「类型可核对但未填值」，
    # 等于把一条正确的 N/A 当成待办事项挂在那儿（333 卡 × 5 项 × 2 组设备 = 3330 行）。
    #
    # 只在「有设备、有规格、且设备真的能在本参数的子表里找到」时才敢下结论；
    # 规格套不上任何一档时不在这里表态，交给下面的序位兜底逻辑。
    gap = _step_gap(step_tables, ck, step,
                    _equip_set(er.get("equip")), _spec_range(er.get("spec")))
    if gap is not None:
        return _no_step_res(gap)

    # —— 收窄链第 1 级：单位对齐（见第 2.8 节 NARROW_CHAINS["strict"]）——
    pool = _stage_unit(er, pool)

    # —— 收窄链第 2 级：一格多分支，按适用规格挑最贴合的那一支 ——
    pool = _stage_cond(er, pool)

    # 设备对齐（判据走 `_equip_gate` —— 与形态链共用 `_equip_keep` 这一份）
    # `no_equip`：这份 PDF 的**表里根本没印设备号**（第九种版式就是这样）。
    # 那是「PDF 没写」，不是「PDF 没有这台设备」—— 按设备收窄会把它整片
    # 判成 no_equip_match，等于把「没写」当成「没有」，凭空造漏判。
    eq = _equip_set(er.get("equip"))     # 后面规格级 / 序位近似还要用，不能省
    narrowed = _equip_gate(er, pool)
    if narrowed is None:
        return {"vals": set(), "scope": "no_equip_match",
                "blank": False, "by_equip": {}}
    pool = narrowed
    cand = [r for r in pool if not r.get("empty")]      # 收窄后的「有值」记录
    # ⚠️ 「有值」不等于「有**数值**」：`SL1+30 ref` 这类格子我读不出数字，
    # 但它是表格里实实在在的一格（标了 ref）。它必须留在 `cand` 里 —— 形态要
    # 靠它传给类型规矩 —— 但**不能往期望值集合里塞 (None, None)**：
    # 那会让期望值里混进一个没有数字的元组，下游一 `sorted` 就炸
    # （TypeError: '<' not supported between NoneType and float，2026-09-30 实测）。
    # 「读不出数值」与「没有这一格」是两件事：这里分开的第一步是让它不参与取值。
    numish = [r for r in cand if r.get("lo") is not None or r.get("hi") is not None]
    only_blank = not numish                             # 只剩空格 → PDF 在此声明无值

    by_equip = {}
    # 精度签名：{数值元组: {(lo_dp, hi_dp), ...}} —— 与 by_equip 平行收集，
    # 供判定层做「3.0 与 3 不能算对」的精度守卫。只记 raw 里真能配到精度的；
    # 配不到（raw 是 `SL1+30 ref` 这类读不出数的）就留空，守卫时跳过。
    val_dp = {}
    for r in numish:
        key = "/".join(r.get("equip") or ()) or (r.get("equip_raw") or "（未标设备）")
        tup = (r["lo"], r["hi"])
        by_equip.setdefault(key, set()).add(tup)
        dp = _rec_dp(r)
        if dp is not None and dp[0] is not None:
            val_dp.setdefault(tup, set()).add(dp)

    # 规格级：Excel「适用规格」 ↔ 子表标题里的直径范围
    spec = _spec_range(er.get("spec"))
    omap = _ordinal_for(ordinal, eq, spec)      # 序位近似的兜底目标（可能为 None）
    approx = False
    if spec and eq and step_tables:
        # 先在「该参数的全部子表」里找规格落在哪张 —— 包括没有当前这一步的表。
        # 只看 cand（有这一步的表）的话，规格明明属于 表1-1，却只能在
        # 表1-3/1-4 里挑，挑不出来就退化成设备级并集 —— 正是要避免的。
        tables = step_tables.get(ck) or {}
        # ⚠️ 这份 PDF 的子表**一张都没印设备号**时（设备号只写在标题括号里、
        # 现行的 `设备/工具编号：` 标签认不出来，如 WI-0036 表1/2/3），
        # 「按设备过滤」这个动作本身就无从谈起 —— 不能因为过滤不出东西，
        # 就把**规格收窄也一起跳过**（原写法 `owned` 为空 → 整段 if 不进）。
        # 那等于放弃唯一还剩的信号，一路滑到「该设备所有同名值的并集」：
        # WI-0036 表1（Ø4.0~5.0 档）根本没有「加热左/右端模」这一行，
        # 却被拿 表2（Ø6.0 档）的 110~140 当成了 φ4.0-5.0 卡的期望值，
        # 凭空造出 10 条「PDF 里有这一步、不该填 N/A」。
        # 一张都没印设备号 → 全部子表都当候选，交给规格去收窄。
        _any_equip = any(v.get("equip") for v in tables.values())
        owned = ([(k, v) for k, v in tables.items() if (v["equip"] & eq)]
                 if _any_equip else list(tables.items()))
        if owned:
            # 有直径档的子表按规格打分；**没有直径档的**（标题里不带 (φ…mm)，
            # 如 WI-0024 的「表1-3 设备/工具编号：T-1281尖端成型参数表」）
            # 规格无从收窄，只能收到设备这一层 —— 既然设备已经对上，就按命中计
            # （等价于 approx）。以前这类表被 k[1] 挡掉，导致「该表有没有这一步」
            # 问不出来：表1-3 只有 Step1，Excel「步骤2」填了 N/A 却报不出 no_step。
            scored = [(_dia_score(spec, k[1]) if k[1] else 1, k, v)
                      for k, v in owned]
            best = max(s for s, _k, _v in scored)
            top = [(k, v) for s, k, v in scored if s == best]
            if best == 0 and omap and omap in tables:
                # 精确套不上任何一档，而这张表整体属于「序位近似」的设备
                # （见 _ordinal_maps）：按「第 N 个规格 ↔ 第 N 档」把它对过去。
                # 只在 best==0 时才走这条路 —— 精确能套上的行，永远走精确路径。
                top, best, approx = [(omap, tables[omap])], 1, True
            if best >= 1:
                # 「规格落在哪张子表」经常不是唯一的：Excel 的「适用规格」是一段
                # **产品尺寸区间**（2.0-4.0 / 4.5-6.0），而 PDF 的分档是按设备参数
                # 划的（2.0-5.5 / 6.0 / 7.0 / 8.0），边界根本对不齐 —— 「4.5-6.0」
                # 就同时压在 表1-1(2.0-5.5) 与 表1-2(6.0) 上。
                #
                # 旧版这时要求「只命中一张表」才敢下 no_step 的结论，平手就直接
                # 跳过去，收窄链于是一路滑到末尾的「该设备下所有同名值并集」，
                # 把 7.0mm 档的 440~480 当成了 4.5-6.0 档的期望值 —— 凭空造出一个
                # 「应为」，而那一档在 PDF 里压根没有这一步。（WI-0023 远端焊接-02/07/08）
                #
                # 正确的问法跟张数无关：**这些候选档里有没有哪一张写了这一步**？
                #   · 一张都没有 → 结论唯一：Excel 该填 N/A（no_step）
                #   · 有任意一张有 → 不能说「没有这一步」，继续往下做取值收窄
                #     （平手时下游会在候选档之间取并集，并标成 union）
                if not any(step in v["steps"] for _k, v in top):
                    return {"vals": set(), "scope": "no_step", "blank": False,
                            "sub": "、".join(_fmt_sub(k) for k, _v in top),
                            "dia": None, "many": len(top) > 1, "ordinal": approx,
                            "steps": sorted({s for _k, v in top for s in v["steps"]}),
                            "by_equip": by_equip}
            elif not approx:
                # 精确套不上、也没有序位兜底 —— 关键：**不要**滑到
                # 「该设备所有同名值的并集」。那不是「这个规格的期望值」，
                # 而是把别的档的参数借过来用，凭空造出一个「应为」。
                # 如实说「这个规格在这台设备下没有对应的档」，交给人核对。
                return {"vals": set(), "scope": "no_spec_match", "blank": False,
                        "dia": None, "by_equip": by_equip,
                        "sub": "、".join(sorted({_fmt_sub(k) for k, _v in owned}))}

    if spec:
        # 序位近似命中时，用**映射到的那一档**去收窄，而不是那个套不上的规格。
        spec_eff = omap[1] if (approx and omap) else spec
        subs, bsubs, bsamp = {}, {}, {}
        for r in pool:
            if not r.get("dia"):
                continue
            k = (r.get("sub") or "", tuple(r["dia"]))
            if r.get("empty"):
                bsubs[k] = bsubs.get(k, 0) + 1     # 该档该格的「明确空值」次数
                bsamp.setdefault(k, r)             # 留一条样本：依据要能说出原文
            elif r.get("lo") is None and r.get("hi") is None:
                # 读不出数值的格子（如 `SL1+30 ref`）：**占住这一档的位子**，
                # 但不往集合里塞 (None, None) —— 那会让期望值里混进没有数字的
                # 元组，下游一 `sorted` 就崩（TypeError: NoneType < float，
                # 2026-09-30 实测）。占住位子才能让下面的规格收窄落到这一档，
                # 而不是因为「这一档没值」滑到设备级并集、把隔壁档的数值借来。
                subs.setdefault(k, set())
            else:
                tup = (r["lo"], r["hi"])
                subs.setdefault(k, set()).add(tup)
                dp = _rec_dp(r)
                if dp is not None and dp[0] is not None:
                    val_dp.setdefault(tup, set()).add(dp)
        if subs or bsubs:
            # ⚠️ 这里原来是 `set(subs) | set(bsubs)` —— 集合里装的是**字符串元组**，
            # 而 Python 的字符串哈希**跨进程随机**（PYTHONHASHSEED），迭代顺序不固定。
            # 后果：依据里「PDF 表3(2.0-4.0mm)、表5(2.5-4.0mm)」会随机翻成
            # 「表5(2.5-4.0mm)、表3(2.0-4.0mm)」。R48 把依据文案纳进零回归标尺后，
            # 这等于让 master 哈希本身带随机噪声 —— 同一份代码同一批语料，
            # 连跑三次能跑出两种结果（2026-09-29 实测），标尺因此失去意义。
            # 改成按 PDF 里子表出现的先后顺序去重：既确定，又贴近原文次序。
            keys = list(subs) + [k for k in bsubs if k not in subs]
            scored = [(_dia_score(spec_eff, k[1]), k) for k in keys]
            best = max(s for s, _k in scored)
            top = [k for s, k in scored if s == best]
            # 直径并列（表4 与 表6 都是 Æ5.0~6.0）→ 用**球囊长度**再分一级。
            # 与类型路径共用 `_qlen_keys`，否则两条路径会给出两个打架的答案
            # （实测 card 0089「加热中间模/5」：类型已判对，期望值却还是两张表并集）。
            if len(top) > 1:
                _q = _qlen_keys(pool, set(top), *_qlen_need(er.get("spec")))
                if _q:
                    top = [k for k in top if k in _q]
            if best >= 1 and top:
                vals = set()
                for k in top:
                    vals |= subs.get(k, set())
                if not vals:
                    # ⚠️ 「没有值」有两种，必须分开说 —— 混在一起就会把 ref
                    # 那一格的值写成 `------`，或者更糟：因为这一档没数值而滑到
                    # 设备级并集，把**隔壁规格档**的数值借来当「应为」。
                    # WI-0036 表1（Ø4.0~5.0）Step9 左右位原文是 `SL1+30 ref`，
                    # 借档后依据里出现表3（Ø6.0）的 `5~180`，用户看得一头雾水。
                    _refonly, _raws = [], []
                    for k in top:
                        for r in pool:
                            if (r.get("sub") or "", tuple(r["dia"] or ())) != k:
                                continue
                            if r.get("empty") or pdf_value_shape(r) != "ref":
                                continue
                            _refonly.append(r)
                            _raws.append(str(r.get("raw") or "").strip())
                    if _refonly:
                        # 这一格**只标了 (ref)、没给数值**（引用一个我读不出的量）。
                        # 结论与空值格一样（都不该填），但依据必须说对话。
                        return {"vals": set(), "scope": "ref_cell", "blank": False,
                                "sub": "、".join(_fmt_sub(k) for k in top),
                                "dia": None, "ordinal": approx,
                                "ref_raw": " / ".join(x for x in _raws if x)[:60],
                                "by_equip": by_equip}
                    # 「没有值」的第三种：PDF 这一格**明确写着 N/A**。
                    # 结论与 `------` 一样（都不该填），但依据必须说对话 ——
                    # 把 N/A 说成「表格里写的是 ------」是依据在说谎。
                    _naonly, _naraws = [], []
                    for k in top:
                        for r in pool:
                            if (r.get("sub") or "", tuple(r.get("dia") or ())) != k:
                                continue
                            if not r.get("empty") or not r.get("na"):
                                continue
                            _naonly.append(r)
                            _naraws.append(str(r.get("raw") or "").strip())
                    if _naonly:
                        return {"vals": set(), "scope": "pdf_na", "blank": True,
                                "na": True, "dia": None, "ordinal": approx,
                                "sub": "、".join(_fmt_sub(k) for k in top),
                                "raw": " / ".join(x for x in _naraws if x)[:60],
                                "by_equip": by_equip}
                    # 规格定位到了这张子表，而这一格在 PDF 里没有值（空白 / `------`）：
                    # 表里**有**这一行这一列，只是没有值。
                    return {"vals": set(), "scope": "pdf_blank", "blank": True,
                            "sub": "、".join(_fmt_sub(k) for k in top), "dia": None,
                            "ordinal": approx, "by_equip": by_equip,
                            # 多张子表并列时原文常常一样（都是 `------`），
                            # 拼出来就成了「------ / ------」—— 去重再拼。
                            "raw": " / ".join(dict.fromkeys(
                                _show_cell(bsamp[k].get("raw") or "")
                                for k in top if k in bsamp))[:60]}
                if len(top) == 1:
                    (sub, dia) = top[0]
                    return {"vals": vals, "scope": "strict" if best == 2 else "approx",
                            "blank": False, "by_equip": by_equip, "val_dp": val_dp,
                            "sub": sub, "dia": tuple(dia), "ordinal": approx}
                # 平手：Excel 的「适用规格」是一个区间，横跨了多张子表
                # （如 4.5-6.0 同时落在 表1-1(2.0-5.5) 与 表1-2(6.0) 上）。
                # 该规格下的产品用哪一套取决于实际直径，两边都可能对上，所以取并集，
                # 并如实标成 union —— 比退回「全设备并集」精确得多。
                #
                # 但「并集」不等于「没收窄到唯一」：这几张子表这一格写的值只要**一样**，
                # 用哪一套答案都相同，期望值事实上就是唯一的。旧版一律标 union，
                # 于是 WI-0023 里冒出 24 条「现值 480~520 / 候选 480~520」被列进
                # 「宽松判定明细（需人工确认）」—— 两条一模一样的行并排，纯噪音。
                # 所以只有真出现分歧（值不同，或某一张写的是 `------`）才算宽松：
                # 后者意味着「按这张表该留空、按那张表该填值」，那才是真没定论。
                _blank_top = [k for k in top if k not in subs]
                return {"vals": vals,
                        "scope": ("union" if (len(vals) > 1 or _blank_top)
                                  else "union_same"),
                        "blank": False,
                        "by_equip": by_equip, "val_dp": val_dp, "ordinal": approx,
                        "sub": "、".join(_fmt_sub(k) for k in top), "dia": None,
                        "subs": [(_fmt_sub(k), tuple(k[1])) for k in top]}

    if not eq:
        return {"vals": set(), "scope": "none", "blank": False, "by_equip": by_equip}

    if only_blank:
        # 只能收到设备这一层，而该设备下这一格在 PDF 里就是空值。
        # 空值有两种写法（空白 / `------` 与 明确写 N/A），依据要说对话。
        _na = [r for r in pool if r.get("empty") and r.get("na")]
        if _na:
            return {"vals": set(), "scope": "pdf_na", "blank": True, "na": True,
                    "by_equip": by_equip,
                    "raw": " / ".join(_show_cell(r.get("raw") or "") for r in _na)[:60]}
        _bl = [r for r in pool if r.get("empty")]
        return {"vals": set(), "scope": "pdf_blank", "blank": True,
                "by_equip": by_equip,
                "raw": " / ".join(dict.fromkeys(_show_cell(r.get("raw") or "")
                                                for r in _bl))[:60]}

    vals = set()
    for v in by_equip.values():
        vals |= v
    if not vals:
        return {"vals": set(), "scope": "no_equip_match",
                "blank": False, "by_equip": by_equip}
    if len(vals) == 1:
        return {"vals": vals, "scope": "strict", "blank": False, "by_equip": by_equip,
                "val_dp": val_dp}
    return {"vals": vals, "scope": "device", "blank": False, "by_equip": by_equip,
            "val_dp": val_dp}


# ------------------------------------------------- 程序附表匹配（WI-0009 / WI-3852 型）
#
# 这一版解决的正是「预设给太多」的老毛病：
# 以前按参数名把整份 PDF 的同名值并成一个集合，于是
#   · WI-3852「表层溶液 固化时间」的候选里混着底层溶液的 40~60s；
#   · WI-0009「产品上提01」有 13 个候选（横跨 14 张附表）；
# 结果「填到隔壁列 / 隔壁附表」也会被判成一致 —— 这种「假一致」比漏判更危险。
#
# 现在的收窄顺序（能收到多细就多细，收不到就如实标出「多套候选 + 依据」）：
#   1. 参数卡里的设备号        → 只留该设备的附表
#   2. 参数卡里的「程序号：N」  → 只留程序号=N 的附表
#   3. 卡内各行投票            → 留下能同时解释最多行的附表
#   4. 行内：程序列(底层/表层) + 步骤号 + 单位 + 参数名 + 序号
# 于是 WI-3852 的 16 条从「9 个候选」收到 1 个；跨程序列对调会被抓出来。

_PROG_IN_ITEM = re.compile(r"程序\s*([A-Za-zＡ-Ｚａ-ｚ])\s*[（(]\s*([^）)]{0,12}?)\s*[）)]")
_SOL_IN_ITEM = re.compile(r"(底层溶液|表层溶液)")
_STEP_IN_ITEM = re.compile(r"步骤\s*(\d+)")
_TAIL_ORD = re.compile(r"(\d{1,2})\s*$")
_UP_WORDS = ("上提", "上升", "提拉", "上移", "上行")
_DOWN_WORDS = ("下移", "下降", "下行", "浸没")
_UNIT_LABEL = {"mm/s": "速度", "mm": "位置"}
_PROG_NO_ITEM = re.compile(r"程序号")


def _direction(s):
    n = str(s or "")
    for w in _UP_WORDS:
        if w in n:
            return "up"
    for w in _DOWN_WORDS:
        if w in n:
            return "down"
    return ""


def _motion_coord(name):
    """运动组名 → (方向, 序号)；不是运动组名返回 (None, None)。

    「下沉浸没1」与「下移浸没01」是同一件事的两种写法：表格里 沉/移 混用，
    序号补不补零也不统一。名字对不上，但**方向和第几次对得上** ——
    这两个坐标才决定「说的是哪一格」，名字里的动词只是叫法。

    与 `_prog_match` 里 `ctx["direction"] / ctx["ordinal"]` 同一套语义：
    程序附表那条路早就在用，竖表这条路一直没有，于是只能靠字面相等。
    """
    n = str(name or "")
    d = _direction(n)
    if not d:
        return None, None
    m = _TAIL_ORD.search(n)
    return d, (int(m.group(1)) if m else None)


# 参数名同义词：Excel 侧叫法 → PDF 侧叫法。
# 只在「词典认得这个名字、但这份 PDF 里查不到同键记录」时救一次，
# 因此 PDF 自己写的就是这个名字时不会触发（不会凭空把两类并成一类）。
#   · 停留时间 ↔ 浸没时间：WI-0042 的亲水涂层附表，第 8/28 步 PDF 写
#     「浸没时间 10s」，Excel 采集项叫「底涂 停留时间」10 —— 位置、数值、
#     单位三项都对得上，只差一个叫法。
_PARAM_ALIAS = {
    "停留时间": "浸没时间",
}


# --------------------------------------------------------------------------
#  参数名词库（学习闭环的注入点）—— 2026-10-30
#
#  词库只回答一件事：**Excel 里这个名字，在 PDF 里写成什么**（Excel 原名 → PDF 原名）。
#  它是 `_PARAM_ALIAS` 的数据化版本，语义完全照抄后者：
#      · 只在「按词典键、大类桶、同义词、字面名都查不到」时才试一次；
#      · 因此 PDF 自己写着原名时不会触发 → **不会把本来并存的两类并成一类**
#        （那是假一致的来源，也是这条判断唯一不能退让的地方）。
#
#  ⚠️ 为什么不做成 `_canonical_param` 里的一层：
#     那个函数带 `@lru_cache(8192)` 且是 PDF 建索引与 Excel 查表的同一漏斗。
#     往里塞动态表有两个坑：改词库后忘了 cache_clear（表现为「确认了不生效」），
#     以及把两个名字并进同一个规范键（会真的造成假一致）。这里走「查不到时兜一层」，
#     既不进缓存，也不改主键。
#
#  ⚠️ 默认空字典 = **行为与没有这个机制时逐条一致**，这是硬门槛的要求。
# --------------------------------------------------------------------------
_LEX_ALIAS = {}          # {Excel 字面键: PDF 原名}
_LEX_NON = set()         # 已确认「不是参数」的字面键（只影响扫描提示，不参与判定）


def set_lexicon(alias=None, non_params=None):
    """注入词库。传空即回到「无词库」行为（app/测试由 name_lex 调用）。"""
    global _LEX_ALIAS, _LEX_NON
    _LEX_ALIAS = dict(alias or {})
    _LEX_NON = set(non_params or ())


# --------------------------------------------------------------------------
#  锚点通道（2026-11-02，见 anchors.py）
#
#  与上面所有通道的根本区别：**不看名字，看位置**。
#
#  其它每一条通道的入口都是「拿 Excel 的采集项名去 PDF 里找一个长得像的名字」，
#  名字一旦对不上就整行掉进「找不到对应」。可名字恰恰是这份文件里最脆弱的东西。
#  锚点是人在界面上指着 PDF 的原表格点出来的「就是这一格」，工具照着坐标读值 ——
#  名字长得像不像从此与这一行无关。
#
#  三条边界：
#    · 空锚点 = 与没有这个机制时逐条一致（第一行就 return）；
#    · 只碰**仍是无对应**的行，已经判出来的行一格不碰；
#    · 锚点指的那张表在这份 PDF 里认不出来（指纹对不上）→ 不生效，如实说明，
#      绝不拿形状相近的表顶上（那会造出假一致）。
# --------------------------------------------------------------------------
_ANCHOR_BINDS = {}


def set_anchors(binds=None):
    """注入锚点绑定 {Excel字面键: {page,tno,row,col,...}}（由 app / anchors.resolve 给）。"""
    global _ANCHOR_BINDS
    _ANCHOR_BINDS = dict(binds or {})


def _anchor_cell(pdf_path, b):
    """按锚点坐标把那一格的字读出来（读不到返回空串，绝不返回猜的值）。"""
    if not pdf_path:
        return ""
    try:
        import pdf_tables
        g = pdf_tables.grid(pdf_path, b.get("page"), b.get("tno"))
    except Exception:
        return ""
    r, c = int(b.get("row") or 0), int(b.get("col") or 0)
    if not (0 <= r < len(g)):
        return ""
    row = g[r] or []
    if not (0 <= c < len(row)):
        return ""
    return (row[c] or "").strip()


def _anchor_pos(b):
    """一格锚点坐标 → 人能读的一句话：`第4页 表2 第3行第2列`。

    有表标签（PDF 上印的那个「表2」）就用标签：位置序号（本页第 5 张表）和
    PDF 上的编号**经常不是一回事**，人拿着依据去翻 PDF 会翻错地方。
    """
    _rc = "第%d行第%d列" % (int(b.get("row") or 0), int(b.get("col") or 0))
    _lab = str(b.get("label") or "").strip()
    if _lab:
        return "%s %s" % (_lab, _rc)
    _tn = (" 第%d张表" % (int(b["tno"]) + 1)) if b.get("tno") is not None else ""
    return "%s%s %s" % (("第%d页" % b["page"]) if b.get("page") else "?", _tn, _rc)


def _apply_anchors(checks, pdf_path):
    """把配了锚点的行，按人指定的那一格重判一遍。

    改哪些行，分两种来源（这是 2026-11-02 加「校正」时分的）：

      · **学来的**（`from_rule`，由表级规则推的坐标）—— 只改仍是
        no_match / no_scope 的行。规则是推的，不该去推翻别的通道已经判出来的结论，
        否则一条规则错了会连坐一整片。
      · **人钉的**（锚点，`from_rule` 为假）—— 任何行都可以改，只要不是
        collection / novalue（那两类本来就不比数值）。
        这是「校正」机制：用户原话「40 条一下子学完了，但是其中有 3 条还是没有
        对上……实际上他认为是对上了的，但是他又抽取错误了」—— 规则推错列时，
        那几行已经有结论（甚至判成「一致」），不让人覆盖就永远改不回来。
        人明确指了一格，那就是最强证据，胜过任何启发式。

    放在 `_apply_spec_gate` 之后、`_apply_fallback` 之前：
      · 在兜底之前 —— 人指的位置比任何启发式都准，不该被「分层匹配」抢走；
      · 在规格门之后 —— 规格门管的是「这台机这一档该不该有值」，与位置无关。
    """
    stat = {"hit": 0, "ok": 0, "diff": 0, "bad": 0, "stale": 0, "fix": 0}
    if not _ANCHOR_BINDS or not pdf_path:
        return checks, stat
    for chk in checks:
        st0 = chk.get("status")
        # 采集项 / 未填值：钉了位置也没有可比的东西，一律跳过
        if st0 in ("collection", "novalue"):
            continue
        k = _literal_key(chk.get("item"))
        if not k:
            continue
        # 先按 (名字, Excel行号) 查，再退回 (名字, 全适用)。
        # 为什么要行号：一张参数卡里同名采集项常有七八行（WI-0039 表4 的
        # 「使用模板将球囊保护套切定长，长度」就是 8 行），只按名字钉会一次
        # 指中 8 行、凭空造出 8 条假差异（2026-11-02 端到端实测踩过）。
        b = _ANCHOR_BINDS.get((k, chk.get("row")))
        if b is None:
            b = _ANCHOR_BINDS.get((k, None))
        if not b:
            continue
        _manual = not b.get("from_rule")
        if not _manual and st0 not in ("no_match", "no_scope"):
            continue        # 规则只捡还没结论的；人钉的才能覆盖
        # 位置必须带上「第几张表」：一页常有好几张表（WI-0030 第 4 页就有 5 张），
        # 只说「第4页 第2行第2列」根本分不清是哪一张 —— 而这次修的正是「串表」。
        where = _anchor_pos(b)
        if b.get("stale"):
            stat["stale"] += 1
            old = (chk.get("basis") or "").strip()
            s = ("这一行你钉过位置（%s），但这份 PDF 里认不出那张表 —— 锚点这次没生效，"
                 "结论保持原样" % (where,))
            chk["basis"] = (old + "。" + s) if old and s not in old else s
            chk["anchor_stale"] = True
            continue
        act = chk.get("_actual")
        # ---- 读格子 ----
        # 多选（Ctrl / Command）时人点的不止一格：**每一格都读**，取值并集。
        # 为什么是并集而不是「必须全中」：同一个采集项在 PDF 里分写两张表时，两张
        # 表各自只覆盖一部分设备/规格 —— 2026-11-11 用户原话「它这里对应了两个表，
        # 它没有分开写。也就是说这两个表的话，取一起的话都是 157 到 181」。
        # 人把两格都点出来，意思正是「这两处都算」，任何一处对上就算对上。
        _cells_rec = list(b.get("cells") or ())
        _multi = len(_cells_rec) > 1
        if _multi:
            _alive = [c for c in _cells_rec if not c.get("stale")]
            _stale_c = [c for c in _cells_rec if c.get("stale")]
        else:
            _alive, _stale_c = [b], []      # 老锚点 / 单格：就是主格那一格
        _cm = _card_model(chk)
        _use, _dead = [], []
        for _c in _alive:
            _t = _anchor_cell(pdf_path, _c)
            if _t:
                _use.append((_c, _t))
            else:
                _dead.append(_c)
        if not _use:
            stat["bad"] += 1
            old = (chk.get("basis") or "").strip()
            s = ("按你钉的位置（%s）去读，那一格现在是空的 —— 表还是那张表，"
                 "但位置变了或这格被合并了，请重新点一次" % (where,))
            chk["basis"] = (old + "。" + s) if old and s not in old else s
            chk["anchor_stale"] = True
            continue
        # 主格 = 最后点的那一格（`cells` 存的就是点选顺序）。底下「读到的是什么」
        # 显示它；其余格子的原文进依据里那句话，供人对照是不是这两处。
        cell = _use[-1][1]
        vals = set()
        _mnote = ""
        _raws = []
        for _c, _t in _use:
            # 一格里有「175-181 for 014PTA / 181-185 for 018PTA」两档时，先按本卡
            # 型号挑那一档再解析 —— 否则永远取到第一档，0.014 的卡和 0.018 的卡
            # 会得出同一个期望值（那等于没核）。
            _seg = _pick_by_model(_t, _cm)
            if _seg != _t and not _mnote:
                _mnote = ("（这一格按型号写了两档，取的是本卡 %s 那一档）"
                          % _model_text(_cm))
            _lo, _hi, _unit, _tag = _parse_value(_seg)
            _raws.append((_c, _t, _lo, _hi))
            if _lo is not None:
                vals.add((_lo, _hi if _hi is not None else _lo))
        # 多格时把每一格的去向交代清楚：读到什么、空的、还是这张表没认出来 ——
        # 否则「并集」是个看不见的动作，人不知道自己的第二格到底算没算进去。
        _mnote2 = ""
        if _multi:
            _bits = []
            for _c, _t, _lo2, _hi2 in _raws:
                _bits.append("%s「%s」%s" % (
                    _anchor_pos(_c), _show_cell(_t),
                    "" if _lo2 is not None else "（解不出数值，没算进去）"))
            for _c in _dead:
                _bits.append("%s（这一格是空的）" % _anchor_pos(_c))
            for _c in _stale_c:
                _bits.append("%s（这张表这份文件里没认出来）" % _anchor_pos(_c))
            _mnote2 = "（你钉的 %d 格取并集：%s）" % (len(_cells_rec),
                                                     "、".join(_bits))
        if not vals or act is None:
            # 读到了字但解不出数值，或者 Excel 这格本来就没值 —— 都不改判定，
            # 只把原文摆出来（「我读到了、但没读懂」绝不能说成「文件里没有」）。
            stat["bad"] += 1
            old = (chk.get("basis") or "").strip()
            s = ("按你钉的位置（%s）读到的是「%s」%s"
                 % (where, _show_cell(cell),
                    "；Excel 这一格没填值，无可比"
                    if act is None else
                    "，但这一格解不出可比较的数值 —— 请人工看一眼"))
            chk["basis"] = (old + "。" + s) if old and s not in old else s
            chk["anchor_raw"] = cell
            continue
        chk["expected"] = sorted(vals)
        chk["scope"] = "anchor"
        _from_rule = bool(b.get("from_rule"))
        chk["anchor"] = {"page": b.get("page"), "label": where,
                         # tno 也要带上：同一页常有好几张表（WI-0028 第 4 页有 6 张），
                         # 面板「双击跳到那一页」只按 page 找会落到本页第一张表去。
                         "tno": b.get("tno"),
                         "row": int(b.get("row") or 0),
                         "col": int(b.get("col") or 0),
                         "raw": cell, "name": b.get("pdf_name") or "",
                         "from_rule": _from_rule,
                         "rule_label": b.get("rule_label") or ""}
        if _multi:
            # 多选的每一格都留下来，界面（和复核的人）才知道这一行看的是哪几处
            chk["anchor"]["cells"] = [
                {"page": _c.get("page"), "tno": _c.get("tno"),
                 "row": int(_c.get("row") or 0), "col": int(_c.get("col") or 0),
                 "raw": _t} for _c, _t, _l, _h in _raws]
        _os = chk.get("_one_sided")
        if _os:
            # 单侧填法：不能按「元组完全相等」比（见 _one_sided_verdict 的长注释）
            side, lim = _os
            vd = _one_sided_verdict(act, vals, side)
            if vd == "out":
                chk["status"] = "mismatch"
                _tail = "；Excel 只填了%s，PDF 这一格落在界外" % (
                    "上限" if side == "hi" else "下限")
            else:
                chk["status"] = "ok"
                chk["scope"] = "anchor_loose" if vd == "within" else "anchor"
                _tail = ("；Excel 只填了%s，PDF 这一格%s" % (
                    "上限" if side == "hi" else "下限",
                    "与之相同" if vd == "same" else "未超出该侧"))
        else:
            chk["status"] = "ok" if act in vals else "mismatch"
            _tail = ""
        stat["hit"] += 1
        stat["ok" if chk["status"] == "ok" else "diff"] += 1
        # 把「原来为什么报找不到」那段摘掉 —— 那段说的是「这份 PDF 的参数表里
        # 没有它」，而现在已经按人指的位置取到值了，两句摆在一起会自相矛盾
        # （实测端到端跑出来就是：「锚点通道：…读到 820±10mm」。紧接着
        #  「参数名已收录，但这份 PDF 的参数表里没有它」，看的人一头雾水）。
        # 取值提示（「类型列写固定值…」）不是那段，照样留着。
        _nm_note = chk.get("_nomatch_note") or ""
        old = (chk.get("basis") or "").strip()
        if _nm_note:
            old = old.replace(_nm_note, "").strip("。；; ")
        # 两种来源要分开说，不能混成一句「按你钉的位置」：
        #   · anchor  = 人自己点的那一格（精确）
        #   · 学来的  = 人点了一格，工具据此推出「这张表行看规格、列看方向+部位」，
        #               再由这条规则算出来的格子。复核的人得知道这一格**不是**人点的，
        #               是推出来的 —— 万一推错了，他才知道该去怀疑哪里。
        # 校正：这一行本来已有结论，是被人重新指了位置推翻的。
        # 必须说清楚「原来是什么、现在是什么」，否则复核的人看到一个「一致」
        # 没了、一个「不一致」冒出来，会以为是工具自己变了口径。
        #
        # 工具没有上一轮的记忆，所以「原来」只能从**这一轮的数据**里推，有两种：
        #   ① `st0` 已经被别的通道判出来了（按名字匹配、分层回退…）→ 直接引它的结论。
        #   ② `st0` 是 no_match，但这一行本来**会被表级规则**算出坐标（`_alt`）——
        #      这就是用户说的那种情况：「他认为是对上了的，但他又抽取错误了」。
        #      规则把列推错一位，旁边那一列的值刚好也落在 Excel 的范围内，于是判成
        #      「一致」。这种错光看结论是看不出来的，只有把人重新指的那一格与规则
        #      原来指的那一格并排摆出来，才看得出差别。
        #      而 app 那边把人钉的坐标 update 上去时会盖掉规则的坐标，所以得靠
        #      `_alt` 把它带下来 —— 见 app 注入锚点处的注释。
        _alt = b.get("_alt") if _manual else None
        _alt_raw = ""
        if _alt:
            _alt_raw = _anchor_cell(pdf_path, _alt)
            # 那一格读出来跟人指的这格一模一样 → 根本没改动什么，别谎报「校正」
            if _alt_raw and cell and _alt_raw == cell:
                _alt = None
        _fix = False
        if _manual:
            if st0 not in ("no_match", "no_scope"):
                _fix = True
            elif _alt:
                _fix = True     # 规则本来会算这一行，现在被人改指了
        _wtxt = ""
        if _fix:
            stat["fix"] += 1
            chk["anchor"]["corrected"] = True
            if st0 not in ("no_match", "no_scope"):
                chk["anchor"]["was"] = st0
                _wtxt = "（这一行原来判的是「%s」，按你重新指定的这一格重判）" % (
                    STATUS_CN.get(st0, st0),)
            else:
                chk["anchor"]["was_rule"] = {"row": int(_alt.get("row") or 0),
                                             "col": int(_alt.get("col") or 0),
                                             "raw": _alt_raw,
                                             "label": _alt.get("rule_label") or ""}
                _wtxt = ("（这一行原来按学到的表级规则看第%d行第%d列%s，"
                         "按你重新指定的这一格重判）"
                         % (int(_alt.get("row") or 0), int(_alt.get("col") or 0),
                            ("「%s」" % _show_cell(_alt_raw)) if _alt_raw else ""))
                # —— 校正的**危险方向**，必须标出来 ——
                # 规则算出来那一格跟 Excel 对不上（本来会报「不一致」），而人钉的
                # 这一格刚好对上 → 结论从「不一致」翻成「一致」。
                # 这有两种可能，工具分不出，但**后果完全不同**：
                #   · 人改对了：规则把列推错了一位，那一列的值是别的规格的；
                #   · 真差异被盖住了：Excel 那一行填的就是别的规格的数。
                # WI-0041 实测（2026-11-02）：φ5.0 那张卡的球囊中间上/下限填了
                # 6.3 / 6.0 —— 那是 φ6.0 那一行的值。把它钉到 φ6.0 的格子上，
                # 两条真差异就变成了「一致」。
                # 不拦（人指的位置永远算数），只打标：`anchor.shade` 供界面显示，
                # 依据里把「本来会报不一致」这句话摆出来。
                # 口径：假一致比假差异更伤 —— 报错会被复核拦下，静默变一致不会。
                if not _os and chk["status"] == "ok":
                    _a_lo, _a_hi, _a_u, _a_t = _parse_value(_alt_raw)
                    _a_vals = {(_a_lo, _a_hi if _a_hi is not None else _a_lo)}
                    if _a_lo is not None and isinstance(act, tuple) and act not in _a_vals:
                        chk["anchor"]["shade"] = True
                        _wtxt += ("；注意：规则算的那一格是「%s」，与你钉的这格不同，"
                                  "按规则本来会判「不一致」—— 你这一格把它改成了「一致」。"
                                  "若那条规则是对的，这里就盖住了一条真差异，"
                                  "请再核一眼 Excel 这一行填的是不是本规格的数"
                                  % (_show_cell(_alt_raw),))
        if _from_rule:
            lead = ("按学到的表级规则（这张表行看适用规格、列看方向+部位）算出第%d行第%d列，"
                    "读到「%s」= %s%s%s"
                    % (int(b.get("row") or 0), int(b.get("col") or 0),
                       _show_cell(cell), _fmt_vals(vals, True), _tail, _wtxt))
        elif _fix:
            lead = ("校正：你重新指定了这一行的位置（%s%s），按你指的格子读到「%s」= %s%s%s"
                    % (where,
                       ("，%s" % b["pdf_name"]) if b.get("pdf_name") else "",
                       _show_cell(cell), _fmt_vals(vals, True), _tail, _wtxt))
        else:
            lead = ("锚点通道：按你钉的位置（%s%s）读到「%s」= %s%s%s"
                    % (where,
                       ("，%s" % b["pdf_name"]) if b.get("pdf_name") else "",
                       _show_cell(cell), _fmt_vals(vals, True), _tail, _wtxt))
        if _mnote:
            lead += _mnote
        if _mnote2:
            lead += _mnote2
        # 指纹认不出、靠「表名 + 原文」重新认回来的，得明说 —— 否则人会以为
        # 指纹还灵，等哪天真失效了也不知道该怪什么。
        if b.get("relinked") or any(c.get("relinked") for c in (b.get("cells") or ())):
            lead += ("（这份文件的表格指纹对不上了，是按表名「%s」+ 那一格的原文重新认"
                     "回来的，位置可能和上次不同，扫一眼再信）" % (b.get("tname") or "这张表"))
        chk["basis"] = (lead + "。" + old) if old and old != lead else lead
    return checks, stat


def _lex_recs(er, vindex):
    """按词库把 Excel 名换成 PDF 原名，再取记录。没登记就返回空列表。"""
    if not _LEX_ALIAS:
        return []
    dst = _LEX_ALIAS.get(_literal_key(er.get("item")))
    if not dst:
        return []
    return vindex.get(_LIT_PREFIX + _literal_key(dst)) or []


# --------------------------------------------------------------------------
#  「附表页」结构维度收窄（2026-10-30，独立通道，见 page_tier.py）
#
#  背景：WI-0042 的 12 张附表是 4 组（RX-90 / RX-150·200 / OTW-90 / OTW-150）×
#  3 个长度档（15-80 / 100-150 / 200-300mm）。四组的位置值域完全不重叠，
#  而**速度四组全同** —— 于是速度一直判得好好的，位置却永远收成十几套。
#  这四个维度在抽取时没进记录（`sub=''`/`prog=''`/`dia=None`），所以收窄链
#  拿不到。现在由 `page_tier.attach()` 从每页标题补上 `_len_tier` / `_pg_rx`。
#
#  ⚠️ 三条保险（任何一条不满足就**原样放过**，宁可多套候选也不误杀）：
#     ① 本行同名候选**全都**带维度才动手（有一条没解析出来 → 放过）；
#     ② 收窄后为空 → 放过；
#     ③ 产品形态只在**卡片大类唯一**时才用（卡片同时含 RX 与 OTW → 弃权）。
#  保险 ① 的关键在于：它是按「候选集」判的，不是按「整份 PDF」判 ——
#  所以同一份 PDF 里有一张没标题维度的表（如 p7 长度对照表）不会让别的表失效。
# --------------------------------------------------------------------------
_CARD_TOKENS = set()      # 卡片的产品形态，如 {"RX"}；空 = 弃权


def set_card_tokens(tokens=None):
    """注入产品形态（由 app/run_batch 从 page_tier.card_tokens 取来）。"""
    global _CARD_TOKENS
    _CARD_TOKENS = set(tokens or ())


def _page_dim_filter(er, recs):
    """按「长度档 + 产品形态 + 导管长度」收窄候选。不满足条件时**原样返回**。

    三个维度合起来才够用：WI-0042 的 12 张附表是
    4 组（RX-90 / RX-150·200 / OTW-90 / OTW-150）× 3 个长度档，
    少收一个维度就还剩三四套值（2026-11-02 用户实测：36 条多套候选）。
    **这三个维度全部写在表外**（附表标题那一行 `For NEO RX-90cm，Balloon Length
    15-80mm`），不在参数值里 —— 所以只能从这里认，不能靠名字。
    """
    if not recs:
        return recs
    if not all(r.get("_len_tier") is not None for r in recs):
        return recs                                  # 保险 ①
    try:
        import page_tier
    except Exception:
        return recs
    want = page_tier.spec_dims(er.get("spec"))
    out = recs
    # —— 长度档：与参数卡那一段**有重叠**就留 ——
    # 老写法给的是单点（`1.25 X 15 90cm` 的 15），新写法给的是区间
    # （`QL(15-80)`）—— 两种都按「重叠」判，单点自然退化成原先的「落在区间内」。
    ln, ln_hi = want.get("len"), want.get("len_hi")
    if ln is not None:
        hi = ln_hi if ln_hi is not None else ln
        sub = [r for r in out
               if not (r["_len_tier"][1] < ln or r["_len_tier"][0] > hi)]
        if sub and len(sub) < len(out):
            out = sub
    # —— 产品形态 + 导管长度：必须**同时**对上 ——
    # 只对形态不变量（卡片同时含 RX 与 OTW 就弃权，保险 ③）。导管长度按集合相交：
    # Excel 写 `CL(150,200)`，PDF 那一页写 `RX-150/200cm`，交集非空才算对得上。
    cms = want.get("rx_cms") or set()
    if cms and len(_CARD_TOKENS) == 1:
        form = next(iter(_CARD_TOKENS))
        sub = [r for r in out
               if any(f == form and (cms & set(c))
                      for (f, c) in (r.get("_pg_rx") or ()))]
        if sub and len(sub) < len(out):
            out = sub
    return out                                            # 保险 ②：只剩空就返回原样


def _excel_prog_ctx(er):
    """把 Excel 采集项拆成「程序列 / 溶液 / 步骤 / 序号 / 方向 / 单位 / 参数名」。

    '程序A(底层溶液)  产品上提01(mm/s)'      → A / 底层溶液 / – / 1 / up / mm/s
    '底层溶液步骤9移动溶液槽到设定位置：mm/s' → A / 底层溶液 / 9 / – / –  / mm/s
    '程序号：'                              → 参数名=程序号（单独处理）
    """
    s = str(er.get("item") or "")
    program, solution = None, ""
    m = _PROG_IN_ITEM.search(s)
    if m:
        program = (m.group(1) or "").upper() or None
        solution = m.group(2) or ""
        s = s[:m.start()] + " " + s[m.end():]
    if not solution:
        m = _SOL_IN_ITEM.search(s)
        if m:
            solution = m.group(1)
            s = s[:m.start()] + " " + s[m.end():]
    if program is None and solution:
        program = "A" if solution.startswith("底") else "B"
    step = None
    m = _STEP_IN_ITEM.search(s)
    if m:
        step = int(m.group(1))
        s = s[:m.start()] + " " + s[m.end():]
    unit = _excel_unit(er)
    base = re.sub(r"[（(][^）)]*[）)]", " ", s)
    base = re.sub(r"[:：]\s*[^\u4e00-\u9fff]*$", " ", base)
    base = _clean(base).strip("：: ")
    labels = set()
    ck = _canonical_param(base)
    if ck:
        labels.add(ck)
    if unit in _UNIT_LABEL:
        labels.add(_UNIT_LABEL[unit])
    m = _TAIL_ORD.search(base)
    return {"program": program, "solution": solution, "step": step, "unit": unit,
            "labels": labels, "ordinal": int(m.group(1)) if m else None,
            "direction": _direction(base), "base": base,
            "is_prog_no": bool(_PROG_NO_ITEM.search(base))}


def _prog_row(r):
    """把一条程序附表记录预编译成可比对的形式（值/单位/标签订成集合）。"""
    vals = _extract_all_values(r.get("text"))
    labels = set()
    lab = _canonical_param(r.get("label")) if r.get("label") else None
    if lab:
        labels.add(lab)
    if not r.get("label"):                      # 没有子标签时，用参数名当标签
        ck = _canonical_param(r.get("param"))
        if ck:
            labels.add(ck)
    m = _TAIL_ORD.search(str(r.get("param") or ""))
    return {
        "program": r.get("program"), "solution": r.get("solution") or "",
        "steps": set(r.get("steps") or ()), "param": r.get("param") or "",
        "label": r.get("label") or "", "labels": labels,
        "ordinal": int(m.group(1)) if m else None,
        "direction": _direction(r.get("param")),
        "units": {_norm_unit(u) for (_a, _b, u) in vals if u},
        "vals": {(a, b) for (a, b, _u) in vals},
        "raw": r.get("raw") or "", "page": r.get("page"),
    }


def _program_subtables(recs):
    """把程序附表记录按「附表」聚合成一张张可独立定位的子表。"""
    subs = {}
    for r in recs:
        key = (r.get("sub") or "", r.get("equip_raw") or "", r.get("prog_no_raw") or "",
               tuple(r.get("dwg") or ()))
        d = subs.get(key)
        if d is None:
            d = subs[key] = {
                "sub": r.get("sub") or "", "equip_raw": r.get("equip_raw") or "",
                "equip": set(r.get("equip") or ()), "prog_no_raw": r.get("prog_no_raw") or "",
                "dwg": tuple(r.get("dwg") or ()), "rows": [], "pages": set(),
            }
        d["equip"] |= set(r.get("equip") or ())
        d["pages"].add(r.get("page"))
        d["rows"].append(_prog_row(r))
    for d in subs.values():
        d["label"] = _prog_sub_label(d)
    return list(subs.values())


def _prog_sub_label(d):
    """附表的可读出处：'附表1-a（T-1654·程序号8）'。"""
    bits = []
    if d["equip_raw"]:
        bits.append(d["equip_raw"])
    if d["prog_no_raw"]:
        bits.append("程序号%s" % d["prog_no_raw"])
    if d["dwg"]:
        bits.append("/".join(d["dwg"]))
    return "%s%s" % (d["sub"] or "未编号附表", "（%s）" % "·".join(bits) if bits else "")


def _prog_match(ctx, rows):
    """在一张附表里找该 Excel 行对应的记录；逐层收窄，能收多细收多细。

    返回匹配到的记录列表（不是值集合）—— 候选值要连「出现在第几步」一起给出来，
    否则收不窄时用户仍然看不出该采信哪一个。
    """
    cand = rows
    if ctx["program"]:
        c2 = [r for r in cand if r["program"] == ctx["program"]]
        if not c2:
            return []
        cand = c2
    if ctx["unit"]:
        c2 = [r for r in cand if not r["units"] or ctx["unit"] in r["units"]]
        if not c2:
            return []                 # 单位对不上就别硬凑（mm 不是 mm/s）
        cand = c2
    if ctx["step"] is not None:
        c2 = [r for r in cand if ctx["step"] in r["steps"]]
        if not c2:
            return []
        cand = c2
    if ctx["labels"]:
        c2 = [r for r in cand if r["labels"] & ctx["labels"]]
        if not c2:
            return []
        cand = c2
    if ctx["ordinal"] is not None:             # 「上提01」↔「提拉上升01」
        c2 = [r for r in cand if r["ordinal"] == ctx["ordinal"]]
        if c2:
            cand = c2
    if ctx["direction"]:
        c2 = [r for r in cand if r["direction"] == ctx["direction"]]
        if c2:
            cand = c2
    return cand


def _row_step(r):
    st = sorted(r.get("steps") or ())
    return "步%d" % st[0] if st else ""


def _vals_of(rows):
    vals = set()
    for r in rows:
        vals |= r["vals"]
    return vals


def _card_context(records):
    """把同一张参数卡的行归到一起，提取设备号 / 程序号 / 每行的上下文与现值。"""
    card = {"equip": set(), "prog_no": None, "ctxs": []}
    for er in records:
        card["equip"] |= set(_equip_set(er.get("equip")))
        ctx = _excel_prog_ctx(er)
        if ctx["is_prog_no"] and card["prog_no"] is None:
            v = er.get("fixed")
            if v is None:
                v = er.get("lo")
            if v is None and er.get("btrue"):
                v = None
            if v is not None:
                card["prog_no"] = _as_float(v)
        lo, hi = _as_float(er.get("lo")), _as_float(er.get("hi"))
        if lo is None and er.get("fixed") is not None:
            lo = hi = _as_float(er.get("fixed"))
        ctx["actual"] = (lo, hi) if lo is not None else None
        ctx["item"] = er.get("item")
        card["ctxs"].append(ctx)
    return card


def _sub_has_no(d, no):
    if no is None:
        return True
    for a, _b, _u in _extract_all_values(d["prog_no_raw"]):
        if abs(a - no) < 1e-9:
            return True
    return False


def _pick_subtables(card, subs):
    """选附表：设备号 → 程序号 → 卡内各行投票；选不出唯一就返回并列的几张。"""
    pool = [s for s in subs
            if not card["equip"] or not s["equip"] or (card["equip"] & s["equip"])]
    if not pool:
        return []
    if card["prog_no"] is not None:
        p2 = [s for s in pool if _sub_has_no(s, card["prog_no"])]
        if p2:
            pool = p2
    if len(pool) <= 1:
        return pool
    scored = []
    for s in pool:
        hit = miss = 0
        for c in card["ctxs"]:
            if c["actual"] is None or not (c["program"] or c["step"] is not None):
                continue
            vals = _vals_of(_prog_match(c, s["rows"]))
            if not vals:
                continue
            if c["actual"] in vals:
                hit += 1
            else:
                miss += 1
        scored.append((hit - miss, hit, s))
    best = max(scored, key=lambda x: (x[0], x[1])) if scored else None
    if not best or best[0] <= 0:
        return pool          # 认不出来就退回设备/程序号筛过的全部，由上层标记「多套候选」
    return [s for sc, _h, s in scored if sc == best[0]]


def _subseq_matches(seq_vals, want, cap=40):
    """在 seq_vals 里找所有「按下标递增且取值依次等于 want」的取法（最多 cap 种）。"""
    out = []

    def dfs(start, j, cur):
        if len(out) >= cap:
            return
        if j == len(want):
            out.append(list(cur))
            return
        for i in range(start, len(seq_vals)):
            if seq_vals[i] == want[j]:
                cur.append(i)
                dfs(i + 1, j + 1, cur)
                cur.pop()

    dfs(0, 0, [])
    return out


def _ordinal_hints(card, subs):
    """把卡里成组的「序号型」行（上提01/上提02…）按步骤先后钉到具体行。

    依据是工艺顺序常识：「上提01」必然写在「上提02」之前。把同组候选行按步号
    排成序列后，要求 Excel 里的值序列是它的**子序列**：
      · 恰好一种取法（或多种取法但取值全都一样）→ 收窄到那一行 = 唯一正确值；
      · 一种取法都没有 → 值与 PDF 的步骤先后矛盾（多半把相邻两行对调了）→ 报不一致；
      · 多种取法且取值不同      → 不硬猜，保留候选并标「多套候选」。

    注意：只有在每行的值本来就出现在该列的候选里时才会做顺序推断，
    所以这条规则只用于「把宽集合收窄」，不会凭空造出不存在的期望值。
    """
    chosen = _pick_subtables(card, subs)
    if not chosen:
        return {}
    groups = {}
    for i, c in enumerate(card["ctxs"]):
        if c["ordinal"] is None or not c["labels"] or c["actual"] is None:
            continue
        key = (c["program"], c["unit"], tuple(sorted(c["labels"])), c["direction"])
        groups.setdefault(key, []).append(i)

    out = {}
    for key, idxs in groups.items():
        if len(idxs) < 2:
            continue
        idxs = sorted(idxs, key=lambda i: card["ctxs"][i]["ordinal"])
        protos = [card["ctxs"][i] for i in idxs]
        base = dict(protos[0])
        base["ordinal"] = None
        rows = []
        for s in chosen:
            rows.extend(_prog_match(base, s["rows"]))
        seq, seen = [], set()
        for r in rows:
            if len(r["vals"]) != 1:
                continue
            st = sorted(r["steps"])
            v = next(iter(r["vals"]))
            if (tuple(st), v) in seen:
                continue
            seen.add((tuple(st), v))
            seq.append((st, v))
        seq.sort(key=lambda x: x[0][0] if x[0] else 10 ** 6)
        if not seq:
            continue
        # 只有当每个现值本来就落在该列的候选里时，才做顺序推断 ——
        # 否则「对不上」的原因是值本身有问题，不是顺序问题。
        if any(p["actual"] not in _vals_of(_prog_match(p, rows)) for p in protos):
            continue
        want = [p["actual"] for p in protos]
        found = _subseq_matches([v for _st, v in seq], want)
        if not found:
            for i in idxs:
                out[i] = {"conflict": True}
            continue
        for j, i in enumerate(idxs):
            vs = {seq[k][1] for k in {m[j] for m in found}}
            st = seq[found[0][j]][0]
            out[i] = {"vals": vs, "unique": True,
                      "steps": "步%s" % "、".join(str(x) for x in st)}
    return out


def _program_expected(er, ctx, card, subs, hint=None):
    """程序附表路径：返回 {vals, scope, basis, subs}；不适用时返回 None。

    scope:
      prog_strict  已定位到唯一（或数值完全一致的几张）附表 → 期望值是那一个值
      prog_multi   附表定位到了，但同一附表同一列里仍有多行同值域（如「上提01」
                   行名对不上 PDF 的「提拉上升01」）→ 如实给候选 + 每一步的步号，
                   并在界面上标「多套候选」，不假装这一条是唯一确定的
    """
    if not (ctx["program"] or ctx["step"] is not None or ctx["is_prog_no"]):
        return None                                   # 这行没有程序表上下文，交给别的路径
    chosen = _pick_subtables(card, subs)
    if not chosen:
        return None
    basis = "、".join(s["label"] for s in chosen[:2])
    if len(chosen) > 2:
        basis += "…另%d张" % (len(chosen) - 2)
    if ctx["is_prog_no"]:
        vals = set()
        for s in chosen:
            vals |= {(a, b) for (a, b, _u) in _extract_all_values(s["prog_no_raw"])}
        if not vals:
            return None
        return {"vals": vals, "scope": "prog_strict", "basis": basis + " 的程序号"}
    if ctx["program"]:
        basis += " 程序%s" % ctx["program"]
    if ctx["step"] is not None:
        basis += " 步骤%d" % ctx["step"]
    if hint and hint.get("vals") is not None:
        # 顺序推断成功：这就是唯一的正确值（附上落到了第几步，方便回原文档复核）
        note = "，按步骤先后定位%s" % hint["steps"] if hint.get("steps") else ""
        return {"vals": set(hint["vals"]), "scope": "prog_strict",
                "basis": basis + note}
    rows = []
    for s in chosen:
        rows.extend(_prog_match(ctx, s["rows"]))
    if not rows:
        return None
    vals = _vals_of(rows)
    if len(vals) == 1:
        return {"vals": vals, "scope": "prog_strict", "basis": basis}
    # 收不窄：把「哪个值出现在第几步」一起给出来，用户能拿着去 PDF 里对
    where = {}
    for r in rows:
        for v in r["vals"]:
            where.setdefault(v, set()).add(_row_step(r))
    detail = []
    for v in sorted(where):
        steps = [x for x in sorted(where[v]) if x]
        label = _fmt_num(v[0]) if v[0] == v[1] else "%s~%s" % (_fmt_num(v[0]), _fmt_num(v[1]))
        detail.append("%s%s" % (label, "(" + "、".join(steps[:6]) + ")" if steps else ""))
    return {"vals": vals, "scope": "prog_multi",
            "basis": "%s；候选 %s" % (basis, "，".join(detail[:8]))}


# 分段写法 → 总名的**兜底**别称表。只在「本名一条记录都没有」时才退回
# （见 `_vertical_expected`），所以不会把真的分段参数混成一个。
_PARAM_FALLBACK = {
    "开始温度": "温度",
    "结束温度": "温度",
    "开始时间": "时间",
    "结束时间": "时间",
    # 2026-11-02（WI-0040 实测）：词典里「时间」这一族被拆成了两个键 ——
    #   Excel 写「焊接时间」/ PDF 写「持续时间」「保持时间」→ 键 `焊接时间`
    #   PDF 写「Time 时间」/「Time」            → 键 `时间`
    # 拆开是为了照顾 WI-0006（PDF 写「持续时间」），可代价是两半永远碰不上面：
    # 两边只剩「互相包含」给 1 分，而 1 分按规矩不能单独立案 → 判「无对应」。
    # WI-0040 就栽在这里：温度（`温度`↔`温度`）、风量（`风量`↔`风量`）都判得对，
    # 唯独「焊接时间」↔`Time 时间` 判不出来，明明表里白纸黑字写着 `7-9s`、
    # Excel 也填 7~9、设备号 T-1208 都对得上。
    # 用退路而不是改 `_canonical_param` 的返回值：改返回值会动 vindex 的分桶，
    # 所有 WI 的桶同时变化，风险面是整个引擎；退路只在**精确桶空着**时才走，
    # 桶里有东西就一步不挪 —— 与上面 4 条同一个性质。
    "焊接时间": "时间",
}


# ================================================ 统一收窄链（三种版式共用）
#
# 一条参数值总是「在某个适用范围下」才成立：某台设备 / 某个规格 / 某个程序号。
# 收窄顺序按证据强度从强到弱：单位 → 规格(直径) → 设备 → 产品线。
# 每一层只在「能收窄且不为空」时才生效 —— 宁可少收一层，也不要把候选清空后
# 假装「找不到」；收完仍多值就如实标 multi 并在界面上提示人工确认。
#
# 为什么要有这一层：横表/程序附表本来就带适用范围（所以 WI-0021/2105/2013/3852
# 能收到唯一值），只有竖表与千分尺表把范围丢了。补上范围之后三种版式共用这条链，
# 就不必「你给我一个 WI，我改一个 WI」。

def _narrow(cand, test, label, used):
    sub = [r for r in cand if test(r)]
    if sub and len(sub) < len(cand):
        used.append(label)
        return sub
    return cand


# ------------------------------------------------ 一格里的「型号分档」（11-02 新增）
#
# WI-0030 第 4 页的焊接参数表，温度那**一格**里写了两档：
#
#     1  Temperature ℃   175-181 for 014PTA ℃  181-185 for 018PTA ℃
#
# 一份参数卡只对应其中一档（这张卡是 DWG-0000012-(0.018)，就该看 181-185）。
# 这里原先有两个毛病，都是用户 2026-11-02 点出来的：
#
#   ① `014PTA` / `018PTA` 里的 014 / 018 被数字扫描当成 **14 / 18** 两个候选值 ——
#      于是界面「应为」里出现 14、18、175-181、181-185 四个数，其中两个根本不是值。
#   ② 不按型号挑就永远「多套候选」，Excel 填哪个都算对 —— 等于这一行没核。
#
# 只在**同一个格子**里出现 ≥2 个不同型号标记时才动手：那是制表人自己把两档写在
# 一格的明确信号。只出现一个标记的格子一律不碰（多半是设备号 `for T-1203` 或
# 一句废话），拿它去筛会错杀真值。本卡的型号在格子里找不到就两档都留着、如实报
# 多套 —— 不猜。

_MODEL_TAG_RE = re.compile(
    r"(?<![A-Za-z])(?:for|适用(?:于)?|对应)\s*(0?\.\d+|\d{2,4})\s*(?:PTA|PTCA|MM)?",
    re.I)


def _model_key(s):
    """'0.018' / '018' / '18' → 18（统一按「千分之一英寸」口径）。

    认不出、或者大得不像型号（>9999）→ None。
    """
    t = str(s or "").strip()
    if not t:
        return None
    try:
        v = float(t)
    except (TypeError, ValueError):
        return None
    if v <= 0:
        return None
    if v < 1:
        v *= 1000                      # 0.018 英寸 → 18
    k = int(round(v))
    return k if 0 < k < 10000 else None


def _model_text(k):
    """18 → '0.018'、150 → '150'（只是给人看的写法，判定用 `_model_key` 那个整数）。"""
    if k is None:
        return "?"
    return ("0.%03d" % k) if k < 100 else str(k)


def _card_model(er):
    """本卡的型号：卡名 `DWG-0000012(0.018)-标记环固定-0001` → 18。认不出返回 None。

    **只认括号里那一段**（图纸编号的直径/规格位），不去全文找数字 ——
    卡名里还带着 2000329 这种大号，全文扫必然认错。
    """
    s = str((er or {}).get("card") or "")
    for m in re.finditer(r"[（(]\s*(0?\.\d+|\d{2,4})\s*[)）]", s):
        k = _model_key(m.group(1))
        if k:
            return k
    return None


def _tagged_values(text):
    """把一格按「值 …for 型号」切成 [(型号 或 None, **不含型号标记的**值原文)]。

    '175-181 for 014PTA ℃ 181-185 for 018PTA ℃'
      → [(14, '175-181'), (18, '℃ 181-185'), (None, '℃')]

    ⚠️ 切出来的那一段必须**不含 `for 014PTA` 本身**：留着它，数字扫描会把
       `018` 读成候选值 18（第一版就漏了这一层，界面「应为」里还是多一个 18）。
       同理，型号标记前面没东西的那一段直接丢掉，绝不去解析标签里的数字。
    """
    s = str(text or "")
    ms = list(_MODEL_TAG_RE.finditer(s))
    if not ms:
        return [(None, s)]
    out, prev = [], 0
    for m in ms:
        piece = s[prev:m.start()]
        if piece.strip():
            out.append((_model_key(m.group(1)), piece))
        prev = m.end()
    tail = s[prev:]
    if tail.strip():
        out.append((None, tail))
    return out or [(None, s)]


def _model_narrow(cand, model, used):
    """一格里有 ≥2 档型号时，按本卡型号收窄；认不出就只剔掉型号号本身。

    返回**新的**记录列表（改的是记录里的 `vals`，所以必须复制，不能就地改 ——
    同一批记录还要给别的行用，就地改会串行）。
    """
    if not cand:
        return cand

    def _vals_of(texts):
        vals, units = set(), set()
        for v in texts:
            for (lo, hi, u) in _extract_all_values(v):
                vals.add((lo, hi))
                if u:
                    units.add(_norm_unit(u))
        return vals, units

    out, by_model = [], False
    before = set()
    for r in cand:
        before |= set(r.get("vals") or ())
    for r in cand:
        segs = _tagged_values(r.get("text") or "")
        tags = [t for t, _ in segs if t is not None]
        if len(set(tags)) < 2:
            out.append(r)
            continue
        keep = [v for t, v in segs if (model and t == model)]
        vals, units = _vals_of(keep) if keep else (set(), set())
        if vals:
            by_model = True
        else:
            # 认不出本卡是这一档（或者挑出来的那段没数字）→ 两档都留，
            # 但**不含型号号**（14 / 18 不是值）
            keep = [v for _t, v in segs]
            vals, units = _vals_of(keep)
        if not vals:
            out.append(r)               # 一段都解不出数 → 原样放过，不硬来
            continue
        r2 = dict(r)
        r2["vals"] = vals
        if units:
            r2["units"] = units
        out.append(r2)
    after = set()
    for r in out:
        after |= set(r.get("vals") or ())
    if after and after < before:
        used.append("型号" if by_model else "型号分档")
    return out


def _pick_by_model(cell, model):
    """锚点通道用：一格里有 ≥2 档型号时，挑出本卡那一档的原文。

    挑不出（没有型号、本卡型号不在这格里、或者挑出来的那段解不出数）→
    返回各档拼起来的一段，只保证「型号号本身不被当成值」，语义仍是多档并列。
    """
    segs = _tagged_values(cell)
    tags = [t for t, _ in segs if t is not None]
    if len(set(tags)) < 2:
        return cell
    keep = [v for t, v in segs if (model and t == model)]
    if keep and any(_extract_all_values(v) for v in keep):
        return " ".join(keep)
    return " ".join(v for _t, v in segs)


_CJK_RE = None


def _cjk(s):
    """只留中文（含全角括号里的字），把英文、数字、标点全部丢掉。

    为什么要丢：卡号 `DWG2001529-标记环固定-0001` 与表名
    `标记环焊接参数表（DWG-0000007-02）` 共有的拉丁片段是 `DWG200`，
    拿它去算「共有词」会把两张**毫不相干**的表算成一家人。
    中文词才是工序的身份。
    """
    global _CJK_RE
    if _CJK_RE is None:
        _CJK_RE = re.compile(r"[^\u4e00-\u9fff]+")
    return _CJK_RE.sub("", str(s or ""))


def _lcs_len(a, b):
    """最长公共**连续**子串的长度（滚动数组，几十个字的小串，够快）。

    用连续子串而不是最长公共子序列：`标记环固定` 与 `铂环焊接参数表` 的
    子序列可以有 `环`+`焊`… 拼出 2 个字的假命中；连续子串只认真连在一起的，
    `标记环` 就是 3、`环` 就是 1 —— 这才是「同一族」与「只是都有个环字」的区别。
    """
    a, b = str(a or ""), str(b or "")
    if not a or not b:
        return 0
    if len(a) < len(b):
        a, b = b, a
    prev = [0] * (len(b) + 1)
    best = 0
    for ca in a:
        cur = [0] * (len(b) + 1)
        for j, cb in enumerate(b, 1):
            if ca == cb:
                cur[j] = prev[j - 1] + 1
                if cur[j] > best:
                    best = cur[j]
        prev = cur
    return best


def _name_score(key, tname):
    """卡上的工序词 与 表名 的「像不像」，0 = 不像。

    取最长公共中文子串，但把 `参数表`/`过程`/`工艺流程` 这类**通用词**剔掉再算：
    所有表名都以「参数表」结尾，不剔的话随便两张表都白送 3 个字。
    """
    a = _cjk(key)
    b = _cjk(tname)
    for w in ("参数表", "工艺参数", "过程参数", "参数", "过程"):
        a = a.replace(w, "")
        b = b.replace(w, "")
    return _lcs_len(a, b)


def _tname_narrow(cand, er, used):
    """表名层：卡的工序词 与 表名 共有的中文，是比设备号更硬的**身份**证据。

    为什么排在设备前面（2026-11-13）：设备号只说明「这张表管哪几台」，
    表名说明「这张表是什么工序」。WI-0028 卡 0001 的工序是「标记环固定」，
    PDF 第 4 页里 表2/表3/表4 叫「标记环焊接参数表」、表6 叫「铂环焊接参数表」。
    设备层那道「PDF 设备组严格包含 Excel」的规则会锁到 表6（3 台 ⊇ 2 台），
    一旦锁上，表2/表3（各一台）连候选都进不去。名字层先过一遍就没有这个坑。

    三条保守规矩，缺一条都可能错杀真值：
      ① 只拿**有名字**的表比；没印表名的表永远留在候选里（规则 ④ 同理：缺
         信息不等于不属于这一档）；
      ② 全批最高分不到 2 个字就当「名字没提供信息」，原样放行；
      ③ 只有真的排掉了东西才记 `used` —— 依据里多一句「表名」，人才能顺着查。
    """
    key = "%s %s" % (er.get("proc") or "", er.get("card") or "")
    if len(_cjk(key)) < 2:
        return cand
    # ⚠️ 设备组**恰好相同**时不许名字否决（2026-11-13 实测踩到）：
    #    WI-0028 的卡 0005 设备写 T-1857/T-1922/T-2101，与表6（铂环）**一模一样**，
    #    Excel 那边的值 180-187 / 8-10 也确实出自表6。可卡名还叫「标记环固定-0005」，
    #    名字层一过就把唯一正确的那张表摘掉了 —— 一改就从「一致」变成「不一致」。
    #    设备组精确相等是「这张表就是给这几台机写的」，没有歧义；包含关系才是弱证据
    #    （3 台包含 2 台太容易成立，WI-0028 的卡 0001 就是这样被锁错表6的）。
    #    所以：精确相等 → 名字不参与否决；只有包含/交集时，名字才说话。
    eq = _equip_set(er.get("equip"))
    if eq and any(r.get("equip") and set(r["equip"]) == set(eq) for r in cand):
        return cand
    named = [r for r in cand if str(r.get("tname") or "").strip()]
    if len(named) < 2:
        return cand
    scored = [(_name_score(key, r.get("tname")), r) for r in named]
    best = max((s for s, _ in scored), default=0)
    if best < 2:
        return cand
    keep = [r for s, r in scored if s == best]
    keep += [r for r in cand if not str(r.get("tname") or "").strip()]
    if len(keep) >= len(cand):
        return cand
    used.append("表名")
    return keep


# ---------------------------------------------------------------- 行级语义打分
#
# 上面那串 `_xxx_narrow` 走的是**排他**：型号不符就整批摘掉、设备组不符就整批
# 摘掉。排他擅长「排除明显不对的」，但它有个致命性质 —— 任何一层判错，整批
# 候选一起归零，最后只能空手回一句「PDF 里没有该设备对应的参数」。可 PDF 里
# 明明写着那个数，人一眼就能看见（用户 2026-11-14：「为什么人类的眼睛就可以
# 识别，为什么你们就不能识别呢」）。
#
# 打分是另一种口径：某一层对不上只**扣分**，不判死。它不该取代排他（排他更准，
# 而且已经被四万条语料校准过），只在排他收到 0 条时兜底 —— 从同名池里挑最像的
# 那一组，把「查无此项」降级成「这是我的最佳猜测，请复核」。
def _semantic_score(er, r):
    """一条 PDF 记录有多像这个 Excel 行要问的那条 —— 纯语义证据，不看坐标。

    刻意**不含**页码、第几张表、行列号：那三项换个模板就全废，只会让模型
    把老模板的先验带进来（用户担心的「再学习也只是对这个模板学习了」）。
    留下来的都是跨模板仍成立的东西：参数名、量纲、设备组、规格、图号、表名。
    """
    sc = 0
    # ① 参数名 —— 最硬。Excel 与 PDF 本来就是照同一份 WI 抄的。
    a, b = _literal_key(er.get("item")), _literal_key(r.get("param"))
    na, nb = _norm_text(er.get("item")), _norm_text(r.get("param"))
    if a and b and a == b:
        sc += 10
    elif na and nb and na == nb:
        sc += 8
    elif na and nb:
        # 名字不完全一样时按「最长公共中文词」给分（`浸没时间` vs `下移浸没01`
        # 这类改写的唯一交集）。上限 4，免得一个字碰巧相同就拿高分。
        sc += min(4, _lcs_len(_cjk(na), _cjk(nb)))
    # ② 量纲：单位都不对，基本不是同一个参数（WI-0039 的「温度30min」）。
    u = _excel_unit(er)
    if u:
        ru = [_norm_unit(x) for x in (r.get("units") or ())]
        if not [x for x in ru if x]:
            sc += 1                    # 这条记录没标单位，不奖不罚
        elif u in ru:
            sc += 4
        else:
            sc -= 5
    # ③ 设备组：恰好相同 = 「这张表就是给这几台机写的」，最硬；
    #    PDF 只写了 Excel 其中一台（子集）= 这张表就为那一台写的，也硬；
    #    PDF 比 Excel 多（包含）= 弱 —— 3 台包含 2 台太容易成立，WI-0028 的
    #    表6（铂环，T-1857/T-1922/T-2101）就是这么把表2/表3 压住的；
    #    无交集 = 基本不是这台机的表。
    eq = _equip_set(er.get("equip"))
    if eq:
        re_ = set(r.get("equip") or ())
        if not re_:
            pass
        elif re_ == eq:
            sc += 8
        elif re_ < eq:
            sc += 2
        elif eq < re_:
            sc += 1
        elif re_ & eq:
            pass
        else:
            sc -= 8
    # ④ 规格：复合区间优先，再看直径档。
    spec_s = er.get("spec")
    if spec_s:
        rr = r.get("spec_ranges")
        if rr:
            sc += {2: 5, 1: 2}.get(_range_match_score(_spec_ranges(spec_s), rr), -3)
        spec_t = _spec_range(spec_s)
        if spec_t and r.get("dia"):
            sc += {2: 4, 1: 1}.get(_dia_score(spec_t, r["dia"]), -2)
    # ⑤ 图号：同一批设备在不同图号下参数不同（WI-0016 的 85-115 / 110-130℃）。
    pns = _pn_full(er.get("spec") or "") or _pn_cores(er.get("card") or "")
    if pns:
        rp = set(r.get("pn_full") or r.get("pn") or ())
        if rp:
            sc += 4 if (rp & pns) else -4
    # ⑥ 表名：换版式之后内容会变，表名不会 —— 所以它是能跨模板累计的那一条。
    #    权重给到 ×2（上限 6）：这是**身份**证据，不是范围证据。WI-0028 实测，
    #    只给 3 分时表6（铂环）靠「设备包含」就能跟表2 打平；乘 2 之后
    #    表2/表3（标记环）= 6、表6（铂环）= 2，设备那点弱优势再也压不住。
    key = "%s %s" % (er.get("proc") or "", er.get("card") or "")
    tn = str(r.get("tname") or "")
    if tn and len(_cjk(key)) >= 2:
        sc += min(6, _name_score(key, tn) * 2)
    return sc


def _rank_pick(er, recs, floor=7, trusted=False):
    """从同名池里挑「最像」的那一组。返回 (记录列表, 分数)；挑不出来回 (None, 0)。

    **只在排他筛到 0 条时才用**。同分全取而不是随便挑一个 —— 同一份 WI 里
    一个参数常写成两张表（用户 2026-11-11：「它对应了两个表卷，它没有分开写」），
    两张得分相等正是「两张都该算」的信号，取并集才对；硬挑一张反而会把
    两边合起来才完整的范围（157~181）截成半截。

    `trusted` 说的是**这个池子可不可信**：
      · True  —— 池子来自 `vindex` 的同键桶，词典已经认定「这些是同一个参数」。
                 此时名字字面对不上是**正常的**：Excel 写「焊接时间」、PDF 写
                 `Set Time`，两边中文压根没交集，但它们是同一个参数。所以
                 名字不许当否决项。
      · False —— 池子是临时拼的（全量记录 / 跨键合并）。这时没有词典背书，
                 名字必须自己站得住，否则 `退出条件加热中间模` 会跟
                 `中间模编号` 配上 —— 后者写的是 2004 这种模具编号，接住就是
                 假差异（实测过）。用户那句「表格里有的就该认出来」的前提是
                 **认对了**；认错一个比漏一个更伤，那是伪造结论。

    兜底的定性（比分数更重要）：宁可拒，不许错。拒了界面上还是「PDF 里没有
    该设备对应的参数」，人自己翻一眼就补上了；错了则是一条盖了章的错误结论。
    """
    if not recs:
        return None, 0
    import os as _os
    if _os.environ.get("WI_NO_RANK"):
        return None, 0          # AB 用：关掉兜底，看它到底改了什么
    a = _literal_key(er.get("item"))
    na = _norm_text(er.get("item"))
    if not a and not na:
        return None, 0
    eq = _equip_set(er.get("equip"))
    scored = []
    for r in recs:
        # 设备完全不是同一批 → **直接出局**，不参与打分。
        # 这一条不能靠扣分了事：名字相同给 10、单位给 4、表名给 6，凑一凑
        # 总能把 -8 抵回来，最后拿另一台机器的参数值去给这台机定罪。
        # 「这是不是这台机的表」是个是非问题，不是程度问题。
        if eq:
            re_ = set(r.get("equip") or ())
            if re_ and not (re_ & eq):
                continue
        b = _literal_key(r.get("param"))
        nb = _norm_text(r.get("param"))
        ok_name = bool((a and b and a == b) or (na and nb and na == nb))
        if not ok_name and not trusted:
            continue
        sc = _semantic_score(er, r)
        if not ok_name:
            # 名字一个字都对不上时把「名字」那档分数补回来：池子是可信的，
            # 没名字分不是它的错，但也不能因此白捡分 —— 就按同键算满。
            sc += 10
        scored.append((sc, r))
    if not scored:
        return None, 0
    best = max(s for s, _ in scored)
    if best < floor:
        return None, 0
    return [r for s, r in scored if s == best], best


def _scope_candidates(er, recs, family=None):
    """按适用范围逐层收窄候选。返回 (候选记录 或 None, 生效的层次名)。

    顺序：单位 → 设备 → 规格 → 产品线。

    设备排在规格前面的理由：设备号是「精确的归类键」，规格是「区间贴合度」，
    后者会把「没标直径」的块整个排除掉。WI-0008 的 表4（T-1729/T-10050，
    表头只写「PTCA产品使用…海波管」、没有直径）就吃过这个亏：
    先按规格挑，它连候选都进不去，于是明明 PDF 里有、却报「无对应」。
    """
    used = []
    cand = list(recs)
    # 型号层放最前：卡名括号里那个直径（`-(0.018)`）是**卡自己的身份**，
    # 比单位、设备、规格都硬 —— 它说的是「这份卡管哪一档」，不是推测出来的贴合度。
    cand = _model_narrow(cand, _card_model(er), used)
    # 表名层紧跟在型号层后面（两者都是「这张卡是谁」的身份信息，不是适用范围）：
    # 先把「不是这个工序的表」整批摘掉，后面设备/规格那一串才好比。
    cand = _tname_narrow(cand, er, used)
    unit = _excel_unit(er)
    if unit:
        cand = _narrow(cand, lambda r: (not r.get("units")) or unit in r["units"], "单位", used)
    eq = _equip_set(er.get("equip"))
    pre_eq = list(cand)                 # 设备层之前的候选，规格层放宽时要回到这里
    eq_exact = []
    if eq:
        # 没有设备信息的记录不能因为「无法比对」被丢掉 —— 它们只是范围未标注，
        # 一律保留（丢了就会把「本来能核对的」变成「无对应」）。
        # 「设备号组合」是表格的身份，不只是过滤条件（用户 2026-09-27 点破）：
        # WI-0005 里 Excel 写 T-2052/T-2069/T-1533 三台的行，PDF 里既有
        # 「恰好这三台」的表、也有「这三台 + T-1176」的四台表 —— 交集匹配把
        # 四台表也捞进来，焊机温度混出 2~3 套值。表格的设备组合应是固定的：
        # 先找「整组恰好相同」的表，命中就锁定；找不到再退回交集（老口径，
        # 一行只写部分设备的情形仍能命中）。
        # ——设备号匹配的四种场景（用户 2026-09-28 定的规则）——
        #   ① 数量一一对应：Excel 与 PDF 的设备数相等、且是同一组 → 确定匹配
        #   ② 各只有一个：两边都只有 1 台且一致 → 匹配，但**必须再校验规格**
        #      （同一台设备常常有两张表、只是规格档不同，光看设备号分不开）
        #   ③ Excel 少于 PDF：Excel 的设备号是 PDF 那一组的**子集** → 包含关系成立，匹配
        #   ④ Excel 多于 PDF：PDF 的设备号比 Excel 少 → 建立不了包含关系，**不匹配**
        # ④ 是这一轮补上的：老口径「有交集就算」，于是 WI-0019 的 Excel（5 台）把
        # 只有 2 台的 表4 也捞了进来，温度混成 170-190 / 195-200 两套候选，
        # 明明只对应一张表却报「多套候选」。
        n_eq = len(eq)
        eq_exact = [r for r in cand if r.get("equip") and set(r["equip"]) == set(eq)]
        if eq_exact:
            used.append("设备组")
            cand = eq_exact
            # 精确命中后，无设备信息的记录不再保留 —— 它们既不属于这一组，
            # 也不是这行的归属；留着只会把唯一值搅成多套（本例实测）。
        else:
            # 场景③：PDF 这一组**包含** Excel 的（且比它多）
            eq_super = [r for r in cand if r.get("equip") and set(eq) < set(r["equip"])]
            if eq_super:
                if len(eq_super) < len(cand):
                    used.append("设备")
                cand = eq_super
            else:
                # 前两层都没命中 → 退回老口径（交集）。
                # ⚠️ 这里**不能**顺手把「PDF 设备比 Excel 少」的表排掉：
                # 元信息项（机器编号）的 Excel 侧常写 3~5 台、PDF 侧是一台台列在
                # 设备清单里的，它们天生只能靠交集命中（WI-0011/2012/2013 实测：
                # 一排就是整片「PDF 里查无此设备」的假错报）。
                # 场景④要防的「强行关联」已经由上面 eq_exact / eq_super 的锁定解决 ——
                # 锁定到那一组之后，设备数更少的那张表根本进不来。
                ne = [r for r in cand if (not r.get("equip")) or (set(r["equip"]) & set(eq))]
                # —— 2026-10-30：交集为空时，不许「只剩未标设备的记录」冒充收窄成功 ——
                # 上面那句「没有设备信息的记录一律保留」是对的，但它有个前提没写明：
                # 池子里得**至少有一条**真属于这台设备。两者都不成立时，留下的那批
                # 未标设备记录根本不是这台设备的值，拿它们当期望值就是凭空定罪。
                #
                # WI-0039 实测（`扩口 温度（℃）`，Excel 写 T-2025/T-1278）：PDF 竖表
                # 16 条温度全挂 T-1279/T-2084，一条都不沾边；只有 p4 的
                # `温度30min`（值 30、单位 min）因为没标设备而幸存 → 剩 1 条 →
                # `strict` → 判「应为 30」，210 行假差异。单位层修好之后这类
                # 「单位都不对」的记录进不来，但**同单位**的无关记录仍会这样混进来，
                # 所以这一层必须单独堵：标了设备却一条都对不上 = 这台设备 PDF 里没有。
                _lab = [r for r in cand if r.get("equip")]
                _hit = [r for r in ne if r.get("equip")]
                if ne and _lab and not _hit:
                    # 有设备标注、但一条都不沾边 → 老老实实说「PDF 里没有这台设备」，
                    # 让上层走 no_equip_match，别硬塞一个值。
                    return None, used
                if ne:
                    if len(ne) < len(cand):
                        used.append("设备")
                    cand = ne
                else:
                    return None, used
    spec = _spec_range(er.get("spec"))
    if spec:
        # 复合规格（WI-3264：`φ(4.5-5.0)xQL(5-60)` 含直径+长度两组区间）：
        # 单区间比对会把 5-60 和 80-200 混成同一档，必须**整组区间**比。
        er_ranges = _spec_ranges(er.get("spec"))
        with_ranges = [r for r in cand if r.get("spec_ranges")]
        if with_ranges and er_ranges:
            scored = [(_range_match_score(er_ranges, r["spec_ranges"]), r)
                      for r in with_ranges]
            best = max(s for s, _ in scored)
            if best >= 1:
                top = [r for s, r in scored if s == best]
                if top and len(top) < len(cand):
                    used.append("规格" if best == 2 else "规格≈")
                    cand = top
            scored = []
        else:
            scored = [(_dia_score(spec, r["dia"]), r) for r in cand if r.get("dia")]
        if scored:
            best = max(s for s, _ in scored)
            if best >= 1:
                top = [r for s, r in scored if s == best]
                if top and len(top) < len(cand):
                    used.append("规格" if best == 2 else "规格≈")
                    cand = top
            elif eq_exact:
                # 「整组恰好相同」锁到了一组**对不上规格**的表上 —— WI-0002 实测：
                # Excel 卡0001 写 T-1112/T-1114、Ø2.0-3.5，PDF 里整组恰好是这两台的
                # 只有 表1-g（Ø3.75-5.0），而 Ø2.0-3.5 的 表1-a/表1-c 各只写了一台。
                # 制表人把两台参数相同的行合并成一行，于是「设备组精确」反而锁错了档。
                # 规格号是三层匹配里更靠后的**确认层**：这一组里压根没有该规格，
                # 说明这次精确是巧合，放宽到交集让规格来定夺。
                relax = [r for r in pre_eq
                         if (not r.get("equip")) or (set(r["equip"]) & set(eq))]
                rescored = [(_dia_score(spec, r["dia"]), r) for r in relax if r.get("dia")]
                if rescored:
                    best2 = max(s for s, _ in rescored)
                    if best2 >= 1:
                        cand = [r for s, r in rescored if s == best2]
                        used = [("设备" if u == "设备组" else u) for u in used]
                        used.append("规格" if best2 == 2 else "规格≈")
    if family:
        cand = _narrow(cand, lambda r: (not r.get("families")) or (set(r["families"]) & family),
                       "产品线", used)
    # 图号：同一批设备在不同图号下参数不同（WI-0016：T-1279 在 DWG-0000006 是
    # 85-115℃、在 DWG-0000004 是 110-130℃），设备号收窄完还是两套，只能靠图号定。
    # 图号写在子表标题 / 正文句子里（`DWG-0000001`），Excel 侧从卡号前缀取。
    # 一样用 `_narrow`：认不出图号的记录（pn 为空）一律保留，绝不因为多了这层
    # 就把真有值说成「无对应」。
    #
    # 2026-10-29 两处补充：
    # ① 优先用 Excel 的「适用规格」（如 `DWG-0000007-01`）而不是卡号 ——
    #    卡号只表明这一行属于哪张卡，规格才说明「这一行管哪一档」；
    #    同一张卡里 4 行分属 4 个规格，只拿卡号永远分不开。
    # ② 用完整图号（带 `-01` 后缀），不再只比数字主干。
    pns = _pn_full(er.get("spec") or "") or _pn_cores(er.get("card") or "")
    if pns:
        cand = _narrow(cand, lambda r: (not (r.get("pn_full") or r.get("pn")))
                       or (set(r.get("pn_full") or r.get("pn")) & pns),
                       "图号", used)
        # 二级：图号**能把候选分开**时，没有图号的记录就不再并列。
        # 它们无法证明自己属于这一档，留着就是「多套对应值」的来源 ——
        # 而多套值一旦存在，这一行的一致性判定就等于没有（只要 Excel 的值
        # 命中任意一套就判「一致」，那是抽奖，不是结论）。
        # 只在「带图号且命中的那批恰好落成**一套值**」时才排空；
        # 若那批本身还是多套，说明图号没起到分档作用，原样保留、绝不乱杀。
        #
        # ⚠️ 只认**行级**图号 `pn_full`，不认表级 `pn`（2026-10-29 实测修正）。
        # 表级 pn 是「这张表整体属于哪个图号」，它挂在部分记录上并不等于
        # 其余记录不属于这个图号 —— 拿它去排空会**错杀真值**。
        # WI-0015 实测：T-1114 / 程序21 / 550-750mW 才是 Excel 550~750 对应的那一行，
        # 只因为它没有表级 pn 而被排掉，剩下的 600~650 反而报了「不一致」——
        # 一条假差异，比漏报更伤（用户 2026-09-24 定的口径）。
        # 行级 pn_full 不同：那是「这一行自己写着这个 P/N」，有资格排空。
        _withpn = [r for r in cand if r.get("pn_full")]
        if _withpn and len(_withpn) < len(cand):
            _pv = set()
            for _r in _withpn:
                _pv |= set(_r.get("vals") or [])
            if len(_pv) == 1:
                cand = _withpn
                used.append("图号唯一")
    # 程序名（第 42 轮）：亲水涂层那类 WI 里，同一台机挂着 6~9 个涂层程序
    # （QSDS-45/80/120/150…），每个程序一套参数。设备号收窄完还剩这么多套，
    # 只能靠程序名定 —— 它是**卡级**信息（写在卡内「涂层程序名称」那一行上，
    # 由 `compare()` 下发到 `er["card_prog"]`）。
    cprogs = er.get("card_prog") or set()
    if cprogs:
        cand = _narrow(cand,
                       lambda r: (not r.get("prog"))
                       or (set(str(r.get("prog")).split("|")) & cprogs),
                       "程序", used)
    return cand, used


def _group_by_equip(recs):
    """按设备分组给出候选值 —— 用于「收不窄时如实展示」而不是硬塞一个值。"""
    out = {}
    for r in recs:
        key = "/".join(r.get("equip") or ()) or (r.get("equip_raw") or "（未标设备）")
        if r.get("vals"):
            out.setdefault(key, set()).update(r["vals"])
        elif r.get("val") is not None:
            out.setdefault(key, set()).add((r["val"], r["val"]))
    return out


_ITEM_STEP_RE = re.compile(r"步骤\s*(\d+)")


def _item_step(name):
    """从 Excel 采集项名里剥出步骤号：「步骤2的激光能量」→ 2；不是步骤式则 None。"""
    m = _ITEM_STEP_RE.search(str(name or ""))
    return int(m.group(1)) if m else None


# Excel 侧竖表型采集项：`(底)下移浸没01  位置( mm)` —— 三个信息挤在一个名字里：
# 并列侧（底/表）、参数名、属性（位置/速度）。PDF 那一格却把两个值写在一起
# （「位置: 1400 mm；速度: 50 mm/s」），所以两边都必须拆开才对得上。
_VX_SIDE_RE = re.compile(r"^\s*[\(（]\s*(底|表|底层|表层)\s*[\)）]")
# 侧标不写在括号里、而是**写在名字里**：「使用底层溶液次数」「使表层溶液次数」——
# 这类名字里的「底层 / 表层」本身就是并列侧（PDF 那边底层一行、表层一行）。
# 只在括号侧标没命中时才用它，且**不从名字里删掉**（它是名字的一部分，删了
# 主干就对不上 PDF 的参数名了）。
_VX_WORD_SIDE_RE = re.compile(r"(底层|表层)")
# 「底涂 / 表涂」是同一件事的另一种写法（WI-0042：Excel 写「底涂 下沉浸没1」，
# 对应 PDF 附表里的「程序 A (底层溶液)」那一列）。
# ⚠️ 只认**名字开头**：这两个词也常出现在别处（「涂层厚度（底涂）」），
# 出现在中间时不足以判定并列列 —— 认错侧就等于拿程序B的值去核程序A。
_VX_PREFIX_SIDE_RE = re.compile(r"^\s*(底涂|表涂)")
_VX_ATTR_RE = re.compile(r"(位置|速度|距离|转速|次数)")
_VX_PAREN_RE = re.compile(r"[\(（][^)）]*[)）]")


def parse_vertical_item(name):
    """Excel 竖表型采集项 → (并列侧, 参数名, 属性)。

      '(底)下移浸没01  位置( mm)'      → ('底', '下移浸没01', '位置')
      '(表) 校正产品位置 速度 (mm/s)'  → ('表', '校正产品位置', '速度')
      '使用底层溶液次数 (次)'           → ('底', '使用底层溶液', '次数')

    「校正产品位置」里本身就带「位置」二字，所以属性取**最后**那个匹配 ——
    取第一个会把参数名吃掉。认不出侧或属性时返回 None（调用方按老口径走，绝不猜）。
    """
    s = str(name or "").strip()
    side = None
    m = _VX_SIDE_RE.match(s)
    if m:
        side = "底" if m.group(1).startswith("底") else "表"
        s = s[m.end():].strip()
    else:
        m2 = _VX_WORD_SIDE_RE.search(s)
        if m2:
            side = "底" if m2.group(1).startswith("底") else "表"
        else:
            m3 = _VX_PREFIX_SIDE_RE.match(s)
            if m3:
                side = "底" if m3.group(1).startswith("底") else "表"
    hits = list(_VX_ATTR_RE.finditer(s))
    attr = None
    if hits:
        h = hits[-1]
        attr = h.group(1)
        s = (s[:h.start()] + " " + s[h.end():])
    s = _VX_PAREN_RE.sub(" ", s)                 # 去掉单位括号 (mm) / (mm/s)
    s = re.sub(r"\s+", " ", s).strip(" :：·-")
    return side, s, attr


def _meta_field_of(name):
    """采集项名（或 PDF 参数名）是不是「表头元信息」那几行；是则返回该下发的字段。"""
    n = re.sub(r"[:：\s]+$", "", str(name or "").strip())
    return _V_META_ROWS.get(n)


def _vertical_expected(er, vindex, fallback, family=None):
    """竖表路径：带适用范围地取期望值。

    scope:
      strict          收窄到只剩一套值
      multi           收窄完仍多套（如不同产品线值不同）→ 如实给候选，标记需人工确认
      name_only       定位不到适用范围，退回全 PDF 同名并集（口径最宽，界面会提示）
      no_equip_match  PDF 里查不到这台设备对应的参数
      name_unknown    参数名不在识别词典内，**根本没查** —— 与「PDF 里没有」不是一回事
    """
    # —— 竖表复合采集项：先剥掉「并列侧 / 属性 / 单位」再定键 ——
    # Excel 写「(表) 校正产品位置 速度 (mm/s)」，PDF 那一行的参数名却叫「校正产品位置」
    # （「速度」只是**属性**，写在值里：`位置: 1210 mm；速度: 50 mm/s`）。
    # 拿整串去定键会被词典的「速度」抢走（动词归类排在「位置」前面），
    # 于是 Excel 查「速度」桶、PDF 把记录放在「位置」桶 —— 两边根本见不到面，
    # 实测后果：属性筛选后 0 条，收窄链一层没走，期望值直接是 23 个跨页并集。
    # 只有在**认出了属性、且主干自己也被词典认得出**时才改用主干；否则照老样子。
    # 第二条是硬约束：「拉伸速度右端（近端）（mm/s）」剥掉「速度」后主干变成
    # 「拉伸右端（近端）」，词典认不出 —— 拿它去查就整行掉进 name_unknown
    # （WI-0004 的「转速」实测就是这么被刮掉的：主干剥完是空串）。
    _vside, _vbase, _vattr = parse_vertical_item(er["item"])
    ck = _canonical_param(er["item"])
    if _vattr and _vbase:
        # ⚠️ 归一到标准名后若**以该属性词结尾**，说明这词是参数名自带的后缀，
        #    而不是「格内属性」，绝不能剥。
        #    实测（WI-0013/2025 近端·远端焊接「移动距离」）：Excel 写
        #    「步骤1的移动距离」，剥掉「距离」得「步骤1的移动」，再归一化却是
        #    **「速度」** —— 于是拿**移动速度**的记录（`速度 0.5mm/s`）去核
        #    **移动距离**（mm）的值。两者单位都不同，永远对不上，
        #    65 条数值差异里 49 条是它。
        #    「移动距离 / 移动速度 / 转动速度」这一类，属性词本身就是名字的结尾，
        #    按词表就能判别；而真正的格内属性（「校正产品位置 速度」里的速度）
        #    归一化后是「位置」，并不以「速度」结尾 —— 两者分得开。
        if not (ck and ck.endswith(_vattr)):
            _ck2 = _canonical_param(_vbase)
            if _ck2:
                ck = _ck2
        else:
            _vattr = None        # 名字自带的词，不是格内属性 → 也不能拿它筛记录
    # —— 属性也可能写在「单位」那一列（2026-10-29，WI-0042）——
    # 这张卡把属性揉进了「单位」格：`位置（mm）` / `速度（mm/s）`，参数名那一列
    # 只写「底涂 下沉浸没1」。只看名字就认不出属性 —— 位置行与速度行于是拿到
    # **同一个并集**（实测 {950, 945, 20} 三合一），位置填成速度的值也判一致。
    # 只在名字里没认出属性时才从单位补：名字自带的（「移动距离」）优先，
    # 且必须排在上面那段「以属性词结尾就不算格内属性」之后，免得把那道判断绕过去。
    if _vattr is None:
        _um = _VX_ATTR_RE.search(str(er.get("unit") or ""))
        if _um:
            _vattr = _um.group(1)
    # —— 字面名通道（第 54 轮）——
    # 词典不认这个名字，不等于 PDF 里没有：Excel 与 PDF 的参数名本来都是照同一份
    # WI 抄的，**字面相同就是最硬的证据**。以前 ck is None 就直接判「未收录」，
    # 于是 WI-0003 rev22 的两条「锻打频率」（PDF 里白纸黑字 `Forging frequency/
    # 锻打频率 15-35Hz`）整项漏掉，界面上写「无对应」。
    _lit = _literal_key(er["item"])
    _by_name = False
    _hit_key = None
    _alias_from = None          # 走的是「同义叫法」时记下这一对，写进依据
    if ck is None:
        # ⚠️ 字面通道有个必要条件：**这一行得带单位**。
        # 名字相同只是「同一个词」，还得有量纲做旁证 —— 用户的原话：
        # 「尤其是那个没录入名字的，不是有单位吗，应该能匹配上才对」。
        # 反例（A/B 实测跑出来的）：WI-3140 的采集项就叫「分段」，固定值 1，
        # 不带任何单位；PDF 里也有一个字面叫「分段」的东西。字面相同就硬配，
        # 12 行从「无对应」翻成「不一致」—— 那是**假差异**。
        # 没有单位的名字多半是编号 / 序号 / 开关，字面相同不足以证明是同一个参数，
        # 宁可退回「无对应」让人自己看一眼。
        recs = (vindex.get(_LIT_PREFIX + _lit) or []) if (_lit and _excel_unit(er)) else []
        if not recs:
            # 词库兜一层（2026-10-30）：人确认过「这个名字在 PDF 里叫什么」时，
            # 走词库登记的名字。仍查不到才算真的认不出。
            recs = _lex_recs(er, vindex)
            if recs:
                # 走的是词库，不是字面相同 —— 依据里不能写「参数名直配」，
                # 复核的人得看得出这条结论来自人工确认过的映射。
                _alias_from = "词库 %s→%s" % (str(er.get("item") or "").strip(),
                                              _LEX_ALIAS.get(_literal_key(er.get("item"))))
        if not recs:
            # 名字没被词典收录、PDF 里也没有同名字的参数 —— 这才真是「认不出」。
            # 单列一档，让它自己出来认领：这是**代码侧**待补（词表缺条），
            # 不是数据侧的问题，别跟「PDF 里没有」混成一样。
            return {"vals": set(), "scope": "name_unknown",
                    "basis": "参数名未收录（不在识别词典内），PDF 里也没有同名的参数",
                    "by_equip": {}}
        _by_name = not _alias_from
    else:
        # `or []`：词典认得、但这份 PDF 里压根没这条时 vindex 返回 None，
        # 而下面「属性 / 并列侧 / 名称」三道筛都在 `if recs:` **之前**跑 ——
        # 拿 None 去迭代会直接把整份 WI 打成「核对异常」（WI-0020 / WI-3852 实测）。
        recs = []
        for _k in _ck_alts(ck):
            recs = vindex.get(_k) or []
            if recs:
                _hit_key = _k
                break
        if not recs and ck in _PARAM_FALLBACK:
            recs = vindex.get(_PARAM_FALLBACK[ck]) or []
        # —— 同义叫法（2026-10-29）——
        # 只在「词典认得这个名字、但这份 PDF 里一条同键记录都没有」时救一次。
        # PDF 自己写的就是这个名字时不会走到这里（上面已经查到了），
        # 所以它**不会**把本来并存的两类并成一类 —— 那是假一致的来源。
        if not recs and ck in _PARAM_ALIAS:
            _dst = _PARAM_ALIAS[ck]
            recs = vindex.get(_LIT_PREFIX + _dst) or []
            if recs:
                _alias_from = "%s→%s" % (ck, _dst)
        if not recs and _lit and _excel_unit(er):
            # 大类桶也没有 —— 再退到字面名（例如 PDF 写的名字词典同样不认，
            # 但两边字面一致）。找不到就还是走下面的老路（空池 → 无对应）。
            recs = vindex.get(_LIT_PREFIX + _lit) or []
            _by_name = bool(recs)
        if not recs:
            # 词库兜一层（2026-10-30）：字典键 / 大类桶 / 同义词 / 字面名都不中时，
            # 才问人确认过的映射。放在最后 = 一旦有更硬的证据就不抢。
            recs = _lex_recs(er, vindex)
            if recs:
                _by_name = False
                _alias_from = "词库 %s→%s" % (str(er.get("item") or "").strip(),
                                              _LEX_ALIAS.get(_literal_key(er.get("item"))))
        # —— 同名优先，**放在收窄之前** ——
        # 同一个大类桶里可能并排躺着**不同的参数**：`压力` 桶里既有
        # 「锻打压力 20-60psi」（锻打机气压），也有 `Pressure/压力 20psi`
        # （压模机模具压力，设备 T-1708）。名字完全一样的那几条才是这一行要问的，
        # 先挑出来再谈设备/图号 —— 否则「图号对不上」会把真记录筛掉，只剩
        # 那条没有图号的无关记录来当期望值（WI-0003 card 0098 就是这么被
        # 判成「20~60 不一致」的，假不一致）。
        #
        # ⚠️ 只在**名字比词典键更具体**时才切这一刀（`_lit != ck`）。
        #    名字本身就是大类词（Excel 只写「时间」）时不能切：WI-0008 实测，
        #    切到 PDF 里字面叫「时间」的那几条，真正对应的「焊接时间」被切掉，
        #    ok/strict 直接变 mismatch —— 那正是第 42 轮踩过的坑。
        _wide = None
        if _lit and ck and _lit != ck:
            _same = [r for r in recs if _literal_key(r.get("param")) == _lit]
            if _same and len(_same) < len(recs):
                # 切之前把整池留一份：这一刀要是把后面的设备/图号收窄带进死胡同，
                # 必须能原样退回来。WI-0015 的「（Step1）激光能量（mW）」实测踩过 ——
                # 同名记录只覆盖部分子表，切完一台设备都对不上，
                # 14 行从 一致/不一致 掉成「无对应」（漏报），当场被 A/B 拦下。
                _wide = recs
                recs, _by_name = _same, True
    # —— 竖表「一格两量」的第二维（第 42 轮）——
    # PDF 那一格写「位置: 1400 mm；速度: 50 mm/s」，Excel 却拆成「位置」「速度」两行。
    # 不按属性 / 并列侧筛，两行拿到的是**同一个并集**（PDF 侧已拆好，见
    # `_vertical_expand`），于是位置填成速度的值照样判一致 —— 假一致里最伤的一种。
    # 两条硬筛都是「筛完为空就原样放过」：老版式没有这些字段，与改动前逐字一致。
    if _vattr:
        _sub = [r for r in recs if r.get("attr") == _vattr]
        if _sub and len(_sub) < len(recs):
            recs = _sub
    if _vside and len({r.get("ord") for r in recs if r.get("ord") is not None}) > 1:
        _want = 0 if _vside == "底" else 1
        _sub = [r for r in recs if r.get("ord") == _want]
        if _sub:
            recs = _sub
    # —— 附表页维度：长度档 / 产品形态（2026-10-30，独立通道）——
    # 放在属性、并列侧之后：这两维是「这条参数在哪一页」，比「哪一列」更靠外，
    # 但必须先于设备/规格收窄 —— 同一台设备在这 12 页里各有一套值。
    if recs:
        recs = _page_dim_filter(er, recs)
    off = 0                      # 步骤号偏移：0=原样对齐，1=PDF 从 Step0 起算、后移一位
    if recs:
        # 步骤号对齐：Excel 写「步骤2的激光能量」，PDF 那张表按 Step0/1/2…分列。
        # 不同步骤压根不是同一个值（Step2 是 480-530、Step3 是 450-520），不收窄的话
        # 三步的值并成一个集合，Excel 填哪一步都判「一致」——那是**假一致**，
        # 比漏判更伤（工具的全部可信度建立在「报出来的不一致都是真的」上）。
        # 只在 PDF 记录真带 step 时才收窄：老版式没有这个字段，筛完为空就原样放过，
        # 与改动前的行为逐字一致。
        want = _item_step(er["item"])
        stepped = [r for r in recs if r.get("step") is not None]
        if want is not None and stepped:
            # 只有 PDF 这批记录里**真的带步骤号**才做步骤判定。老版式的竖表没有
            # step 字段，此时一律不介入 —— 与改动前的行为逐字一致。
            # —— 先按「设备 + 规格」收窄，**不看步骤** —— 拿到这台设备这个规格下
            # PDF 到底摆了哪几步，再决定要不要挪位。
            #
            # ⚠️ 上一版（第十三轮）在这里写错了两次，都栽在「先试原样、不行再挪」：
            #   · 第一次看「步骤号存不存在」→ 别张表有 Step3，offset=0 就算命中；
            #   · 第二次改成看「收窄是否成功」→ 但 0 起算的表里 Step2 本来就存在、
            #     设备规格也对得上，offset=0 **照样收窄成功**，于是永远轮不到偏移。
            # 两次都是同一类错：把「能不能对上」当成了「该不该挪」的判据。
            #
            # 正确判据是**这张表自己从几开始编号**：表头若从 Step0 起（min==0），
            # 它就是 0 起算表，Excel 步骤 k ↔ PDF Step k-1，**必须挪**；否则原样。
            pool, _u = _scope_candidates(er, stepped, family)
            if pool is None:
                pool, _u = [], []
            # 按子表分组：同一页同一张表里可能还夹着两张（p5 的 表1-a / 表1-b）
            groups = {}
            for r in pool:
                groups.setdefault((r.get("page"), r.get("table"),
                                   tuple(r.get("equip") or ())), []).append(r)
            picked = []
            for g in groups.values():
                gset = {x["step"] for x in g}
                # 「这张表从几开始编号」要问**表头**，不能问幸存下来的记录：
                # 值是 NA 的步骤在建点时就丢了，幸存集里没有 Step0 不代表表头没有。
                decl = set()
                for x in g:
                    decl.update(x.get("steps_all") or ())
                base = decl or gset
                shift = 1 if min(base) == 0 else 0
                s = want - shift
                if s in gset:
                    picked.extend(x for x in g if x["step"] == s)
                    off = shift
            if not picked:
                # 整组精确匹配（第十一轮）偶尔锁过头：Excel 写 8 台一组，PDF 里
                # step1 那格标的是 8 台、**step3 却拆成两个 4 台组**分别标注
                # （WI-0006 p5 表1-b，两组都是 1.8~2.3）。精确档只能捞到 step1，
                # step3 的记录被整组排除 → 报「该表里没有步骤3」，那是**假不一致**：
                # Excel 填的 1.8~2.3 其实是对的。
                # 退一步：先按步骤号筛，再交回收窄链（交集档认得下这两个 4 台组）。
                decl_all = set()
                for r in stepped:
                    decl_all.update(r.get("steps_all") or ())
                want0 = (want - 1) if min(decl_all or {r["step"] for r in stepped}) == 0 else want
                for s in dict.fromkeys([want, want0]):
                    sub = [r for r in stepped if r["step"] == s]
                    if not sub:
                        continue
                    c3, u3 = _scope_candidates(er, sub, family)
                    if c3 is not None:
                        recs, cand, used = sub, c3, u3
                        off = 1 if s != want else 0
                        picked = sub
                        break
            if not picked:
                # 表头声明了 Step0/1/2，Excel 却要步骤3 → 这与「PDF 里没有这台
                # 设备」是两码事。混着说会让人去翻设备号，翻半天翻不出东西。
                have_txt = "、".join("步骤%d" % x for x in sorted({r["step"] for r in stepped}))
                return {"vals": set(), "scope": "no_step",
                        "basis": "PDF 该表里没有步骤%d（表头只有 %s）" % (want, have_txt),
                        "by_equip": {}}
            if recs is not picked:
                recs = picked
                cand, used = _scope_candidates(er, recs, family)
            if cand is None:
                cand, used = pool, _u
            if off:
                # 挪过位就要说出来：不然界面写「步骤2」、PDF 里其实是 Step1，
                # 人工复核时会对不上，又变成一笔糊涂账。
                used = list(used or []) + [
                    "步骤对齐(PDF 该表从 Step0 起算，已按 Excel 步骤号后移 1 位)"]
        else:
            cand, used = _scope_candidates(er, recs, family)
        if cand is None:
            # 收不窄时按「更宽的那一档」再试一次，口径回到改动前：
            #   · `_wide`      —— 同名优先切之前的整池（名字一刀切过头）
            #   · 大类桶        —— 精确（大类|限定）桶套不上这批设备
            _alt = _wide
            if _alt is None and _hit_key and "|" in _hit_key:
                _alt = vindex.get(_hit_key.split("|", 1)[0]) or []
            if _alt and len(_alt) > len(recs):
                _c2, _u2 = _scope_candidates(er, _alt, family)
                if _c2:
                    recs, cand, used = _alt, _c2, _u2
        if cand is None:
            # —— 行级语义打分兜底（2026-11-14）——
            # 排他筛到 0 条，**不代表 PDF 里没有**：可能只是某一层筛子太硬，
            # 也可能几张表各写一半（WI-0028 的 157~181 就分在两张表里，单看
            # 哪一张都只有半截）。从同名池里按语义挑最像的一组，把「查无此项」
            # 降级成「最佳猜测，请复核」—— 依据里必须写明是打分推的，
            # 不能让人以为这是硬筛命中。
            cand, _rs = _rank_pick(er, _wide or recs, trusted=True)
            if cand:
                used = list(used or []) + ["按语义打分挑的(%d分·硬筛没命中)" % _rs]
        if cand is None:
            # 步骤式参数要带上步骤号再说一遍：只说「没有该设备对应的参数」会让人
            # 去翻设备号，而问题其实是**这台设备那张表里没有这一步**。
            basis = ("PDF 里没有该设备在步骤%d 下的参数" % want
                     if want is not None else "PDF 里没有该设备对应的参数")
            # —— 2026-10-30：把「没标设备的同名记录」写进依据 ——
            # 这类记录不能当本行设备的期望值（就是 `_scope_candidates` 那道新闸挡掉的
            # 东西），但也不能一声不吭：复核的人拿到「PDF 里没有该设备对应的参数」，
            # 唯一的动作就是回去翻设备号，可 PDF 里明明白纸黑字写着数。
            # WI-0039 实测：Excel 写 T-2025/T-1278，PDF 同页有一条 `温度30min`（30）
            # 没标设备 —— 不说出来，人就会以为 PDF 里没有温度这一项。
            _unl = set()
            for _r in recs:
                if not _r.get("equip"):
                    _unl |= set(_r.get("vals") or ())
            if _unl:
                _txt = "、".join(
                    ("%g~%g" % t if t[0] != t[1] else "%g" % t[0])
                    for t in sorted(_unl)[:6])
                basis += ("；未标设备的同名记录写着 %s（无设备标注，不能算作这台设备的值）"
                          % _txt)
            return {"vals": set(), "scope": "no_equip_match",
                    "basis": basis, "by_equip": {}}
        vals = set()
        for r in cand:
            vals |= r["vals"]
        # —— 参数名确认层：收窄完**仍多套值**时，按名字本身再切一刀 ——
        # 词典把「浸没时间」「静止时间」「UV固化计时」「提拉上升01…05」全压进
        # 「时间」「速度」这类泛键，于是设备+程序+并列侧都锁到只剩一张表了，
        # 期望值还有好几套（WI-0031 实测：浸没时间 {0,10}、提拉上升位置 9 套）。
        # 而两边**名字本来就是一样的**，只是词典先把它压扁了。
        #
        # ⚠️ 必须放在收窄**之后**且「只在仍多值时」才动刀。放在收窄前（第 42 轮
        # 第一版实测）是有害的：WI-0008 的「时间」被切到 PDF 里字面叫「时间」的
        # 那几条（16/18），真正对应的「焊接时间」被切掉，ok/strict(9~11) 直接
        # 变成 mismatch/multi。确认层的语义就是「已经收到了，再验一次」——
        # 唯一值时不介入，切完没变小也不介入。
        _nb = _norm_text(_vbase) if _vbase else ""
        # —— 运动方向 + 序号（2026-10-29，WI-0042）——
        # 名字里的动词并不可靠：Excel 写「下沉浸没1」，PDF 写「下移浸没01」。
        # 字面对不上，但「往哪个方向动、第几次」是这张表的结构本身，两边一致。
        # 放在字面名那层**之前**：它的判据更宽（不要求字面相等），
        # 但只在「位置/速度/并列侧都收窄完仍多套」时才动手，
        # 且切完没变小就不采纳（与下面那层同一条规矩）。
        _md, _mo = _motion_coord(_vbase)
        if _md and len(vals) > 1:
            _sub = [r for r in cand if _motion_coord(r.get("param")) == (_md, _mo)]
            if _sub and len(_sub) < len(cand):
                _v2 = set()
                for r in _sub:
                    _v2 |= r["vals"]
                if _v2 and len(_v2) < len(vals):
                    cand, vals = _sub, _v2
                    used = list(used or []) + ["运动方向+序号"]
        if _nb and len(vals) > 1:
            _sub = [r for r in cand if _norm_text(r.get("param")) == _nb]
            if _sub:
                _v2 = set()
                for r in _sub:
                    _v2 |= r["vals"]
                if _v2 and len(_v2) < len(vals):
                    cand, vals = _sub, _v2
                    used = list(used or []) + ["参数名"]
        # —— 收窄完仍多套：用语义打分再排一次（2026-11-14）——
        # 与上面两层同一个性质：只在**仍然多套**时才动刀，且只在最高分明显
        # 高出时才采纳。打平（差距 < 3）说明它分不开 —— 硬挑一个就是瞎猜，
        # 宁可如实报「多套候选」让人自己看一眼。
        #
        # 为什么这一刀是安全的：候选池已经过了型号/表名/设备/规格/图号五层
        # 硬筛，池子里**每一张都是这台机器的表**，这里只是问「这几张里哪张更
        # 像这一行」，不会把无关的表拉进来。
        if len(vals) > 1 and len(cand) > 1:
            _sc = [(_semantic_score(er, r), r) for r in cand]
            _top = max(s for s, _ in _sc)
            _nxt = max([s for s, _ in _sc if s < _top], default=-999)
            if _top - _nxt >= 3:
                _sub = [r for s, r in _sc if s == _top]
                _v3 = set()
                for r in _sub:
                    _v3 |= r["vals"]
                if _v3 and len(_v3) < len(vals):
                    cand, vals = _sub, _v3
                    used = list(used or []) + ["语义打分%d:%d" % (_top, _nxt)]
        if vals:
            labels = []
            for r in cand:
                if r["basis"] not in labels:
                    labels.append(r["basis"])
            if _alias_from:
                # 「按同义词配上的」要写在依据里：叫法不同就说明这不是靠词典
                # 推出来的，复核的人得知道该不该信这一条。
                # 词库来的映射本身已带「词库」前缀，不再套「同义词」。
                used = ([_alias_from] if _alias_from.startswith("词库")
                        else ["同义词 %s" % _alias_from]) + list(used or [])
            if _by_name:
                # 说清楚这一档是「按名字配上的」，不是靠词典推出来的 ——
                # 复核的人看到依据不同，才知道该不该信。
                used = ["参数名直配"] + list(used or [])
            prefix = "＋".join(used) + "：" if used else ""
            # 精度签名：与 `_step_expected` 同口径，供「3.0 与 3 不能算对」的守卫用。
            # 以前这里不传，竖表路径整条没有精度校验（见 `_rec_dps` 的说明）。
            _vdp = {}
            for r in cand:
                for tup, dpp in _rec_dps(r).items():
                    _vdp.setdefault(tup, set()).add(dpp)
            res = {"vals": vals,
                   "scope": "strict" if len(vals) == 1 else "multi",
                   "basis": prefix + "；".join(labels[:4]),
                   "by_equip": _group_by_equip(cand),
                   "val_dp": _vdp}
            # 不等号方向：候选里所有带方向的记录必须**同向**才认定；
            # 有 lo 有 hi（或方向不明）就不带方向，判定层按老口径走。
            sides = {r.get("side") for r in cand} - {None}
            if len(sides) == 1:
                res["side"] = sides.pop()
            return res
    cand = fallback.get(ck, set()) if ck else set()
    unit = _excel_unit(er)
    vals = {(lo, hi) for (lo, hi, u) in cand if not unit or not u or _norm_unit(u) == unit}
    if not vals:
        return None
    return {"vals": vals, "scope": "name_only",
            "basis": "PDF 全部同名值（定位不到具体设备/规格，口径偏宽）", "by_equip": {}}


def _meta_tokens(v):
    """把 'T-1362/T-10161/T-10058'、'SPC2000019-01' 这类值拆成一个个编号。"""
    return [t for t in re.split(r"[/／、,，;；\s]+", str(v or "")) if t]


def _contains_code(hay, needle):
    """在已 _flat 的文本里找编号，要求两侧不是字母数字。

    没有这个边界，'T-1114' 会命中 'T-11141'，把错了的编号判成一致。
    """
    n = _flat(needle)
    if not n:
        return False
    return bool(re.search(r"(?:^|[^0-9a-z])" + re.escape(n) + r"(?![0-9a-z])", hay))


def _meta_actual(er, kind):
    """元信息项的「实际值」：(类型, 值)；没有可核对的值返回 None。

    取值同样不死认类型列（与参数行一致）：下限/上限有值就按数值读，固定值列有值就按它读。
    特殊的只有「机器编号」—— 它的答案本来就不是一个参数值，而是「这台机是哪台机」，
    所以没填值时退回该行的设备名称列（Excel 里这一项的值就写在那里）。
    """
    lo = _as_float(er.get("lo"))
    if lo is not None:
        return ("num", lo)
    fx = er.get("fixed")
    if fx not in (None, ""):
        f = _as_float(fx)
        if f is not None:
            return ("num", f)
        s = str(fx).strip()
        if s.upper() not in ("N/A", "NA", "NIL", "NONE", "无"):
            return ("text", s)
    if kind == "机器编号" and er.get("equip"):
        return ("text", er.get("equip"))
    return None


# 「设备号 + 括号机器序号」：T-1114（65） / T-1697(52)
# 同一台机器在**不同直径档下的序号不同**（T-1114 在 2.0-4.0mm 档是 65、4.5-7.0mm 是 66、
# 8.0mm 是 89）。只比设备号会把「填了隔壁档的机器序号」判成一致 —— 那恰恰是最该抓的错。
_META_SUFFIXED = re.compile(r"(T-\d+)\s*[（(]\s*([0-9]+)\s*[）)]", re.I)


def _meta_suffixed(text):
    """取出文本里所有「设备号 + 序号」二元组（已归一）。"""
    return {(_flat(a), b) for a, b in _META_SUFFIXED.findall(str(text or ""))}


def _meta_match(act, vals):
    """实际值是否命中期望值集合。

    期望值是 PDF 表头原文，可能是合并值（'T-1114（65）/T-1281（50）'）、
    带机器序号（'T-1697（52）'）、或斜杠分隔的程序号（'1/1'），要拆开逐个比。

    两条要点：
    ① 实际值**带了机器序号**时，期望值必须出现同样的「设备号+同样序号」，
       不许退化成只比设备号 —— 否则「2.0mm 档填成 8.0mm 档的 T-1114（89）」会漏判。
    ② 实际值没带序号（Excel 只写 'T-1254'）时按设备号宽松匹配；且要从实际值里
       **再抠一遍纯设备号**，因为 Excel 常写成 'T-2057-DWG-0000007-03,04'（设备号+工装号），
       只按分隔符拆会得到整串，在 PDF 侧（只有 'T-2057'）永远匹配不上 —— 那是误报。
    """
    kind, v = act
    want = toks = None
    if kind == "text":
        want = _meta_suffixed(v)
        toks = set(_meta_tokens(v)) | set(re.findall(r"T-\d+", str(v), re.I))
    for raw in vals:
        s = _flat(raw)
        if not s:
            continue
        if toks is not None:
            if want:
                if want & _meta_suffixed(s):
                    return True
                continue            # 带了序号就不许退化 —— 宁可报不一致让人复核
            for t in toks:
                if _contains_code(s, t):
                    return True
        else:
            try:
                num = float(v)
            except (TypeError, ValueError):
                continue
            got = {float(x) for x in re.findall(r"\d+(?:\.\d+)?", s)}
            if num in got:
                return True
    return False


def _meta_expected(er, mindex, family=None):
    """元信息路径（机器编号 / 程序号 / 保护套 / 芯轴）：带适用范围地取期望值。

    为什么要有这一路：这几项的答案就写在参数表表头，版式层早就抽到了，但比对层
    从没把它们当成「参数」核对 —— 字符串型的「机器编号」被整类当采集项跳过，
    数值型的「程序号」则因为 PDF 里查不到同名参数而报「无对应」。
    结果是每个 WI 都漏掉一整类采集项，而它们恰恰最容易填错。

    scope:
      meta_strict  收窄到唯一（或唯一档位）→ 硬判定
      meta_multi   收窄完仍多套，且**确有候选没被现值命中** → 如实给候选，
                   需人工确认（哪一套才是对的，会影响一致与否）
      meta_all     收窄完多套，但**每一套候选都被现值命中** → 用哪一套答案都一样，
                   判定的可信度与唯一命中同级 → 硬判定。
                   （PDF 表头常把整条产线的机器编号写在一个合并格子里，
                    候选就是那份清单，Excel 填的整组设备本来就都在里面 ——
                    这种「候选」提供不了任何新信息，列进「宽松判定明细」纯噪音。）
      no_meta      PDF 表头里没有这一项
      no_equip_match  PDF 里没有该设备对应的元信息
    不是元信息项时返回 None（交回原先的路径）。
    """
    kind = _meta_kind_of(er.get("item"))
    if not kind:
        return None
    # 第二道闸（2026-10-30）：元信息的值永远是**标识符**（T-2084 / DWG…-01 /
    # 程序A），不可能是「30~40」这种测量区间。名字再像，只要这一行是
    # 浮点且给了上下限区间，就交回普通参数路径 —— 否则「读起来像保护套」的
    # 参数会被元信息专线整条抢走，最后报一句和事实无关的「表头里没有这一项」。
    if (er.get("type") == "浮点" and er.get("lo") is not None
            and er.get("hi") is not None and er.get("lo") != er.get("hi")):
        return None
    recs = list(mindex.get(kind) or [])
    if not recs:
        return {"kind": kind, "vals": set(), "scope": "no_meta", "by_equip": {},
                "basis": "PDF 表头里没有「%s」这一项" % kind}
    cand, used = _scope_candidates(er, recs, family)
    if cand is None:
        return {"kind": kind, "vals": set(), "scope": "no_equip_match", "by_equip": {},
                "basis": "PDF 表头里没有该设备的「%s」" % kind}
    vals, labels = set(), []
    for r in cand:
        vals.add(r["value"])
        if r["basis"] not in labels:
            labels.append(r["basis"])
    prefix = "＋".join(used) + "：" if used else ""
    by_equip = {}
    for r in cand:
        by_equip.setdefault("/".join(r.get("equip") or ()) or "（未标设备）", set()).add(r["value"])
    # 候选全部来自「设备清单表」兜底时只能做存在性核对 —— 如实降档，不冒充唯一命中
    scope = "meta_strict" if len(vals) == 1 else "meta_multi"
    if all(r.get("weak") for r in cand):
        scope = "meta_list"
    return {"kind": kind, "vals": vals, "scope": scope,
            "basis": prefix + "；".join(labels[:4]), "by_equip": by_equip}


def _dia_eq(a, b):
    """两个直径档是不是同一档 —— 按**数值**比，不按字符串。

    Excel 规格写 '5.00 X 20 150cm'（两位小数），针规表印的是 '5.0x20-150mm'，
    字符串比 '5.00' != '5.0' → 明明有 φ5.0 档却判成「表里没有」，退到 step_down
    拿 φ4 推算（0.03″），界面上还会写一句「查找表没有 φ5.00 档」—— 假差异 + 假说法。
    """
    fa, fb = _as_float(a), _as_float(b)
    if fa is not None and fb is not None:
        return abs(fa - fb) < 1e-9
    return str(a or "").strip() == str(b or "").strip()


def _micrometer_expected(er, kind, mic, family=None):
    """千分尺/针规查找表：先按产品线，再按直径，最后按设备收窄。

    老口径只按直径取，于是 WI-0011「Fold 头闭口针规 过 / φ2.0」会同时拿到
    表2-a 的 0.026″ 与 表2-b 的 0.029″，填哪边都算对。带上范围后才分得开。

    顺序必须先产品线后直径：表2-c 只属于 POT 线、表2-a 属于 产品线A 线，
    先按直径挑会挑到 POT 的 φ2.25 档（0.029″），而 产品线A 表里根本没有
    φ2.25 这一档 —— 正确做法是先限定 产品线A 线，再取「不大于该直径的最大档」。

    直径档位在 PDF 里是断档的（表2-a 只有 1.25/1.75/2.0/2.5/3.0/4.0）。
    Excel 对中间的 φ1.5、φ2.25 等填的是「不大于该直径的最大档」的值
    （0.022″ 就是 φ1.25 档）—— 这里照此推算，但**必须注明是推算**，
    界面上标成「按较小档 φ1.25 推算」，绝不冒充成精确命中。

    长度也是一把刀（2026-10-30）：同一根径按长度分两档、值不同，
    φ5.0 就同时有 0.035″（20-150mm）与 0.036″（200-300mm）。
    先直径后长度再设备 —— 只按直径取会让 Excel 填 0.035 被判成
    「应为 0.036」或者干脆退到 φ4 的 0.03″，两种都是假差异。
    """
    dia = _diameter_from_spec(er.get("spec"))
    recs = list(mic.get(kind) or [])
    if not recs:
        return {"vals": set(), "scope": "no_match", "basis": "该 PDF 里没有这类查找表", "by_equip": {}}
    used = []
    pool = recs
    if family:
        pool = _narrow(recs, lambda r: bool(set(r.get("families") or ()) & family), "产品线", used)
    derived = ""
    cand = [r for r in pool if _dia_eq(r.get("dia"), dia)] if dia else list(pool)
    if not cand and dia:
        target = _as_float(dia)
        lower = [r for r in pool if _as_float(r.get("dia")) is not None
                 and _as_float(r["dia"]) <= target + 1e-9]
        if lower:
            best = max(_as_float(r["dia"]) for r in lower)
            cand = [r for r in lower if _as_float(r["dia"]) == best]
            derived = "查找表没有 φ%s 档，按不大于该直径的最大档 φ%s 推算" % (dia, _fmt_num(best))
    if not cand:
        return {"vals": set(), "scope": "no_match",
                "basis": "查找表里没有 φ%s 档，也没有更小的档位可推算" % (dia or "?"),
                "by_equip": {}}
    # 长度档收窄（2026-10-30 加）。实案 WI-0043 p6 的针规表：
    #   φ3.5x20-150mm → 过 0.026″ / 不过 0.027″
    #   φ3.5x200-300mm→ 过 0.028″ / 不过 0.029″
    #   φ5.0x20-150mm → 过 0.035″ / 不过 0.036″
    #   φ5.0x200-300mm→ 过 0.036″ / 不过 0.037″
    #   φ6.0 同形（0.039/0.040 与 0.040/0.041）
    # 只按直径取会让 φ5.0 同时命中 0.035 与 0.036，填哪个都算对（scope=multi）；
    # 更糟的是 φ5.0 收不到档时走 step_down 退到 φ4（0.03/0.031）——
    # 用户截图里「Excel 填 0.035、工具说应为 0.03」就是这个原因。
    # Excel 规格 '5.00 X 20 150cm' 的第二个数就是长度，`φ5.0xQL(200-300)`
    # 的 QL 区间也是长度 —— 两种写法都由 `_length_from_spec` 归一成区间，
    # 再按**重叠**判（2026-11-02：原来只认老写法，QL 写法直接 None →
    # 这道收窄整段跳过，φ5.0 两档都留，界面「应为」成了 0.035、0.036 两个值）。
    seg_len = _length_from_spec(er.get("spec"))
    len_note = ""
    if seg_len is not None and cand:
        keep = [r for r in cand if _len_span_overlap(r.get("len"), seg_len) is not False]
        if keep and len(keep) < len(cand):
            cand = keep
            used.append("长度档 %smm" % _span_text(seg_len))
        elif not keep:
            # 这个长度不在任何一档里（例如 Excel 写 15、表里只有 20-150/200-300）
            # —— 宁可不动，也不能把候选清空变成「无法匹配」。
            len_note = "长度 %s 不在任何长度档内，未按长度收窄；" % _span_text(seg_len)
    eq = _equip_set(er.get("equip"))
    if eq:
        ne = [r for r in cand if (not r.get("equip")) or (set(r["equip"]) & set(eq))]
        if ne and len(ne) < len(cand):
            cand = ne
            used.append("设备")
    vals = {(r["val"], r["val"]) for r in cand if r.get("val") is not None}
    if not vals:
        return {"vals": set(), "scope": "no_match", "basis": "查找表里没有可用值", "by_equip": {}}
    labels = []
    for r in cand:
        lb = "%s%s" % (r.get("sub") or "（未标附表）",
                       "/" + "/".join(r["families"]) if r.get("families") else "")
        if lb not in labels:
            labels.append(lb)
    prefix = "＋".join(used) + "：" if used else ""
    if derived:
        prefix = derived + "；" + prefix
    if len_note:
        prefix = len_note + prefix
    return {"vals": vals,
            "scope": "step_down" if derived else ("strict" if len(vals) == 1 else "multi"),
            "basis": prefix + "；".join(labels[:4]),
            "by_equip": _group_by_equip(cand)}


# ----------------------------------------------------------------- 比对

# ⚠️ 2026-09-30 清掉两个「幽灵常量」：本文件曾在 10537 行再定义一次 `_fmt_num`，
# 而**函数体相同的那一份**在文件末尾（搜索 `def _fmt_num` 只应命中一处）。
# Python 从上往下执行，后定义的静默覆盖先定义的 —— 中间这段的调用点看着像在用
# 这一份，实际用的是末尾那一份。改错了地方、跑起来毫无变化，是最难查的一类 bug。
# 同批删掉的还有 `NUMERIC_TYPES` / `COLLECTION_TYPES`（全项目零引用，
# 「类型」列的归一早已由下面的 `_type_kind` 自带关键词表负责）。
# 判据：先跨 wi_core / app / _regress 三处确认零引用，再删。


def _type_kind(t):
    """把「类型」列的任意写法归到 num / fixed / bool / str / unknown。

    换一份表时枚举名可能不同（数值/数字/小数/常量/文本…），这里做容错归一。
    """
    s = str(t or "").strip().lower()
    if not s:
        return "unknown"
    for k in ("浮点", "整数", "十进制", "小数", "数值", "数字", "float", "double",
              "decimal", "number", "numeric", "int"):
        if k in s:
            return "num"
    for k in ("固定", "常量", "定值", "constant", "fixed"):
        if k in s:
            return "fixed"
    for k in ("布尔", "bool", "boolean"):
        if k in s:
            return "bool"
    for k in ("字符", "文本", "string", "text", "char"):
        if k in s:
            return "str"
    return "unknown"


def _as_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _precision_mismatch(er, actual, expected, val_dp):
    """「3.0 与 3 不能算对」的精度守卫。

    在 `actual in expected` 数值命中之后调用：数值虽相等，但 Excel 填的值与
    PDF 期望值的小数位数不一致时，返回 True（应判 mismatch）；其余返回 False
    （不干预，保持原判定）。

    判据（务实、不误伤）：
      · Excel 侧精度：取 er 的 lo_dp / hi_dp / fixed_dp（纯数字的小数位数）；
        只要有一端是 None（非纯数字，如带单位/区间），就不做精度校验。
      · PDF 侧精度：取 val_dp 里「数值 == actual」那一项的精度签名集合；
        集合为空或全 None 时同样不校验。
      · 命中后：Excel 精度签名 ∈ PDF 精度签名集合 → 一致；否则 → 不一致。

    为什么用「∈ 集合」而不是「逐位相等」：同一数值可能在多张子表里写法不一
    （表1 写 3、表2 写 3.0），此时 Excel 填 3 或 3.0 都该算对 —— 只要 PDF 里
    真出现过这种写法。
    """
    if not actual or not expected or not val_dp:
        return False
    lo, hi = actual
    # 只对「数值精确命中」的值做精度校验；命中不了说明数值本就不等，交给原判定
    matched = None
    for tup in expected:
        if tup == actual:
            matched = tup
            break
    if matched is None:
        return False
    dps = val_dp.get(matched)
    if not dps:
        return False
    # Excel 侧精度：按 actual 的来源取对应字段
    if er.get("lo_dp") is not None or er.get("hi_dp") is not None:
        lo_dp = er.get("lo_dp")
        hi_dp = er.get("hi_dp")
    elif er.get("fixed_dp") is not None:
        lo_dp = hi_dp = er.get("fixed_dp")
    else:
        return False
    # 单侧填法（只填下限/上限）：只看填了的那一端精度
    if lo == hi:
        # 单值或单侧：两端同一精度即可（取任一端非 None 的）
        want = lo_dp if lo_dp is not None else hi_dp
        if want is None:
            return False
        return not any(p[0] == want for p in dps if p[0] is not None)
    # 区间：两端精度都要对得上（3.0~5.0 vs 3~5 → 两端都不同 → 不一致）
    if lo_dp is None or hi_dp is None:
        return False
    return not any(p[0] == lo_dp and p[1] == hi_dp for p in dps
                   if p[0] is not None and p[1] is not None)


def _fmt_precision_basis(er, actual, expected, val_dp):
    """把「精度不一致」的差异说成人话：PDF 写 3（0 位小数）、Excel 填 3.0（1 位小数）。

    不能用 _fmt_num / _fmt_vals —— 它们把 3.0 抹平成 3，精度差异会被显示层吃掉，
    看的人对着「PDF 写 3、Excel 填 3」根本看不出为什么判错。这里直接按精度签名
    还原两边的原始写法，让差异一眼可见。
    """
    lo, hi = actual
    # PDF 侧：取该值在 val_dp 里的精度签名（可能有多种，取第一种即可）
    pdf_dps = val_dp.get(actual) or set()
    pdf_txt = _dp_txt(lo, hi, next(iter(pdf_dps)) if pdf_dps else None)
    # Excel 侧精度
    if er.get("lo_dp") is not None or er.get("hi_dp") is not None:
        xl_dp = (er.get("lo_dp"), er.get("hi_dp"))
    else:
        xl_dp = (er.get("fixed_dp"), er.get("fixed_dp"))
    xl_txt = _dp_txt(lo, hi, xl_dp)
    return ("数值相等，但小数位数与 PDF 不一致 —— PDF 写 %s、Excel 填 %s"
            % (pdf_txt, xl_txt))


def _dp_txt(lo, hi, dp):
    """按精度签名把 (lo, hi) 还原成带正确小数位数的文字。

    dp 为 (lo_dp, hi_dp)；某端为 None 时该端退回「不标小数位数」的 _fmt_num。
    3.0（1 位）→ '3.0'；3（0 位）→ '3'；区间 3.0~5.0 → '3.0~5.0'。
    """
    def _one(v, d):
        if d is None:
            return _fmt_num(v)
        return ("%." + str(d) + "f") % float(v)
    if hi is None or abs(float(lo) - float(hi)) < 1e-9:
        return _one(lo, dp[0] if dp else None)
    return "%s~%s" % (_one(lo, dp[0] if dp else None), _one(hi, dp[1] if dp else None))


def _one_sided_verdict(actual, expected, side):
    """单侧限值的核对口径（Excel 只填了下限或只填了上限时走这里）。

    老口径是「元组完全相等」：Excel 的 (0.52, 0.52) 必须逐个等于 PDF 的某一组
    (lo, hi)。可 Excel 在「只填一侧」时，那个数**未必**是精确值 ——
    例：WI-0039 的「压力（Mpa）」，下限列是空的、上限列写 0.52，
    PDF 那边写的是 `> 0.52Mpa`（下限约束），`是否允许超上下限`=是。
    拿它当「必须等于 0.52」去比，PDF 写 0.40 就会被判成不一致 —— 那是冤枉。

    所以分两档：

      'same'   PDF 的取值与 Excel 填的那个数**完全相同** → 硬判定（可当唯一命中用）
      'within' PDF 的取值落在 Excel 声明的这一侧之内（上限侧：≤ 该值；
               下限侧：≥ 该值）→ 一致，但口径偏宽，单列一档让人看得见
      'out'    落在界外 → 不一致

    side == 'hi' 表示 Excel 只填了上限，side == 'lo' 表示只填了下限。
    expected 是 PDF 声明的取值集合，每项形如 (lo, hi) 的二元组。
    """
    if not expected or not actual:
        return "out"
    lim = _as_float(actual[1] if side == "hi" else actual[0])
    if lim is None:
        return "out"
    same, within = True, True
    for pair in expected:
        if not isinstance(pair, (tuple, list)) or len(pair) < 2:
            return "out"
        p_lo, p_hi = _as_float(pair[0]), _as_float(pair[1])
        if p_lo is None or p_hi is None:
            return "out"
        if not (p_lo == lim and p_hi == lim):
            same = False
        if side == "hi":
            if p_hi > lim:
                within = False
        elif p_lo < lim:
            within = False
    if same:
        return "same"
    return "within" if within else "out"


def compare(pdf_by_name: dict, excel_records: list, micrometer=None, pdf_precise=None,
            pdf_steps=None, pdf_programs=None, pdf_vertical=None, family=None,
            pdf_meta=None, pdf_step_blanks=None, balloon=None, pdf_path=None) -> dict:
    """类型分派核对：浮点/固定值 → 规格核对；字符串/布尔值 → 采集项跳过。

    核对优先级：
      1. 表头元信息（机器编号 / 程序号 / 保护套 / 芯轴）—— 答案在参数表表头里
      2. 千分尺查找表（直径 + 设备 + 产品线）
      3. WI-0008 精确匹配（P/N + 尺寸 + 设备）
      4. 步骤式横表匹配（Excel「步骤N的XX」 ↔ PDF「参数 | Step N」，按设备对齐）
      5. 程序附表匹配（Excel「程序A(底层溶液) 产品上提01」 ↔ PDF「附表N / 程序A列 / 第N行」）
      6. 竖表匹配（Excel 采集项 ↔ PDF「参数名称 / 参数值」块，带设备/规格/产品线收窄）

    第 1 条是本轮新增：元信息以前根本没被核对过（字符串型整类跳过、数值型报无对应）。
    第 6 条以前是「通用名称并集」——只按参数名把全 PDF 的同名值并成一个大集合，
    于是「光强」同时给出 T-1018 组的 1.5-2.0 与 T-1540 组的 3-5，填错也判一致。
    现在改成与前面几条同一条收窄链：收不窄就如实标 multi（多套候选）。
    """
    # 附表页维度（长度档 / 产品形态）：只往记录上补字段，不改任何既有判定。
    # 解析不到就什么都不加，下游 `_page_dim_filter` 见不到 `_len_tier` 会原样放过。
    if pdf_path and pdf_vertical:
        try:
            import page_tier
            page_tier.attach(pdf_vertical, pdf_path)
        except Exception:
            pass

    # 表头元信息：{类别: [带适用范围的记录…]}
    meta_index = {}
    for r in (pdf_meta or []):
        meta_index.setdefault(r["kind"], []).append(r)

    # 「未认领格」索引：{规范名: [PDF 里写了字、但我没解出值的条目…]}
    # 用途是**不让工具把「我读不懂」说成「文件里没有」**（见 `LAST_UNPARSED`）。
    # 有了它，no_match 的依据才能分清：
    #   · 索引里没有 → PDF 里真没这个参数，依据照旧说「可能写在正文里」
    #   · 索引里有   → PDF 白纸黑字写了，是我没读出来，必须如实写出来给人看
    unparsed_index = {}
    for u in (LAST_UNPARSED or []):
        for k in {_canonical_param(u.get("param") or ""), _canonical_param(u.get("raw") or "")}:
            if k:
                unparsed_index.setdefault(k, []).append(u)

    # 把 PDF 参数按规范名聚合：{规范名: 值集合}（末路兜底用）
    canonical_lookup = {}
    for name, vals in (pdf_by_name or {}).items():
        key = _canonical_param(name)
        if key:
            canonical_lookup.setdefault(key, set()).update(vals)
    # 竖表参数按规范名建索引：{规范名: [带适用范围的记录…]}
    # 同一批记录再按**字面名**建一份（前缀 `名:` 隔开），词典不认的名字靠它兜住，
    # 详见 `_literal_key`。两条索引指向同一批记录对象，不复制。
    vindex = {}
    for r in (pdf_vertical or []):
        key = _canonical_param(r["param"])
        if key:
            vindex.setdefault(key, []).append(r)
            # 同时挂一份到**大类桶**（`压力|储料第四段` → `压力`）。
            # 精确键查不到时要能退回大类，而这份 PDF 里每条压力参数都带着限定词、
            # 大类桶本来是空的 —— 不退这一档，「储料压力（Bar）第四段」就会从
            # 一致/差异掉成「无对应」（实测 2 条）。
            # 退回后的口径与改动前完全相同：整类候选 + 设备/规格/单位收窄 + 名称确认层。
            bare = key.split("|", 1)[0]
            if bare != key:
                vindex.setdefault(bare, []).append(r)
        lk = _literal_key(r["param"])
        if lk:
            vindex.setdefault(_LIT_PREFIX + lk, []).append(r)
    if pdf_precise:
        for p in pdf_precise:
            p["sizes"] = _size_set(p["spec"])

    # 横表参数建索引：{(规范名, 步骤号): [记录…]}
    # blank_index：同键，但只放「格子存在、明确写空」的记录（`------`）。
    # step_tables：{规范名: {(子表编号, 直径): {equip, steps}}} —— 该参数一共有哪几张
    # 子表、每张有哪些步骤。用于识别「该规格的表里根本没有这一步」（横表按直径
    # 分档，不同档步骤数不一样：T-1175 的 2.0-5.5mm 档只有 2 步，8.0mm 档有 4 步）。
    #
    # steps 用**完整集合**（含有值格 + 空格）统计：这里问的是「表头声明了哪几步」，
    # 而不是「哪几步碰巧有值」。否则表头有 Step3、值写着 `------` 的档会被误判成
    # 「没有这一步」（no_step），而正确答案其实是「有这一步，但值为空」。
    step_index = {}
    blank_index = {}
    step_tables = {}
    # sub_params：{(子表编号, 直径): {"equip": {设备号…}, "params": {规范名: {步骤…}}}}
    # ——「这张表里到底有哪些参数」。与 `step_tables` 是同一批记录的**两个切面**：
    #   · step_tables 按参数查「哪几张表有它、有几步」
    #   · sub_params  按表查「这张表印了哪些参数」
    # 第二个切面是第 53 轮补的：光有前者只能问「有没有这一步」，问不出
    # 「这张表压根没有这个采集项」—— 而那是比步骤更靠前的结构性问题
    # （见 `_param_gap`）。
    sub_params = {}
    for r in list(pdf_steps or []) + list(pdf_step_blanks or []):
        # ⚠️ 一条记录可能要挂到**多个**规范名上：PDF 列头写「加热端模」（左右
        # 同值）时，Excel 的「加热左端模」「加热右端模」都得能查到它
        # （见 `_param_keys`）。只挂主名，另一半采集项在自己该看的那张表里
        # 一格都找不到，只能去借隔壁设备的同名参数 —— 那是假一致。
        cks = _param_keys(r["param"])
        if not cks:
            continue
        dia = tuple(r["dia"]) if r.get("dia") else ()
        skey = (r.get("sub") or "", dia)
        sp = sub_params.setdefault(skey, {"equip": set(), "params": {}})
        sp["equip"].update(r.get("equip") or ())
        for ck in cks:
            (blank_index if r.get("empty") else step_index).setdefault(
                (ck, r["step"]), []).append(r)
            # 没有直径档的子表也要进索引。像 WI-0024 的参数表标题是
            # 「表1-3 设备/工具编号：T-1281尖端成型参数表」——**不带 (φ2.0-6.0mm)**，
            # 靠设备号就能把它定位死。以前这里 `if not dia: continue` 把这类表整个
            # 排除在外，于是「该设备这张表有没有这一步」永远问不出来，
            # 步骤数判不出来（表1-3 只有 Step1，Excel 步骤2 填 N/A 会被误报成「未填值」）。
            t = step_tables.setdefault(ck, {}).setdefault(
                skey, {"equip": set(), "steps": set()})
            t["equip"].update(r.get("equip") or ())
            t["steps"].add(r["step"])
            sp["params"].setdefault(ck, set()).add(r["step"])

    # 序位近似兜底：Excel 的「适用规格」与 PDF 分档**大面积**对不上时，
    # 按「从小到大第 N 个规格 ↔ 第 N 档」建立映射（见 _ordinal_maps）。
    # 只对「精确套不上任何一档」的行生效，且判定强度单列一档（ordinal）。
    ordinal_maps = ({} if os.environ.get("POC_NO_ORDINAL") == "1"
                    else _ordinal_maps(excel_records, pdf_steps, pdf_step_blanks))

    # 程序附表：按卡聚合记录 + 预解析每行上下文（位置/速度、上提/下移、步骤号…）
    prog_subs = _program_subtables(pdf_programs) if pdf_programs else []
    card_ctxs = {}
    card_hints = {}
    if prog_subs:
        by_card = {}
        for er in excel_records:
            by_card.setdefault(er.get("card"), []).append(er)
        for k, v in by_card.items():
            card, rows_of = _card_context(v), v
            card_ctxs[k] = (card, rows_of)
            card_hints[k] = _ordinal_hints(card, prog_subs)

    # 球囊专属规矩的总开关（第 2.6 节 detect_balloon）。传 None 时自己认 ——
    # 这样老调用点（批量 / 回归脚本）不改代码也吃得到这条规矩。
    if balloon is None:
        balloon = detect_balloon(pdf_steps, excel_records, pdf_path=pdf_path)

    # 卡级元信息下发（第 42 轮，与 PDF 侧 `_vertical_expand` 的第 ③ 步对应）：
    # 「涂层程序名称 / 设备编号」这类只写在卡内**某一行**上的信息，同卡其它行要拿它
    # 收窄 —— 不下发的话，整张卡只剩设备号一个维度，同一台机 6~9 个程序的值会全并起来。
    # 卡里没有这类行的：**一律不动**（er 上不会出现 card_prog，收窄链跳过这一层）。
    card_progs = {}
    for _e in excel_records:
        if _meta_field_of(_e.get("item")) != "prog":
            continue
        _v = _e.get("fixed")
        if _v is None:
            _v = _e.get("lo") if _e.get("hi") is None else _e.get("lo")
        _v = str(_v or "").strip()
        if _v:
            card_progs.setdefault(_e.get("card"), set()).add(_v)

    checks = []
    for er in excel_records:
        kind = _type_kind(er["type"])
        chk = dict(er)
        if er.get("card") in card_progs:
            chk["card_prog"] = card_progs[er["card"]]
        # 「显示上下限 / 允许超上下限」按类型规矩**应该**填什么：由第 2.6 节的
        # 规则表算出来（自检与界面共用，不在这里另写一套）。界面把「Excel 实际
        # 填的」和「规矩要求的」上下并排，填反了才看得出来。
        #
        # 球囊规矩会在**类型层面**表态（PDF 带 >/< → 应为「字符串」）：那种情况
        # 下两个开关要按**应有的类型**去算，不能按表里填错的那个类型算，
        # 否则「应为」一列会跟着一起错。
        _shape = dominant_shape(pdf_shapes_for(er, step_index))
        # 球囊/非球囊对 ref 的口径相反（2026-10-14 用户定）：
        #   · 球囊 WI：PDF 标 (ref) 的格子不填（空着对、填了错）；
        #   · 非球囊 WI：ref 不享受特殊待遇，照常填写、照常数值比对。
        # 所以非球囊时要把「这一格是 ref」降级回数值形态（区间/单边/单值/空），
        # 否则 ref 会一路带进 flag_contradiction / decide_missing_cell，
        # 把「该填」的行判成「不必填、不该填」。
        if not balloon and _shape == "ref":
            _shape = _demote_ref_shape(er, step_index)
        _op_raw = pdf_op_raw(er, step_index) if balloon else ""
        _want_typ = expected_type_for(er, _op_raw, balloon)
        _rule = type_fill_rule(_want_typ or er.get("type"), _shape)
        chk["want_show"] = _rule["show_lim"]
        chk["want_over"] = _rule["allow_over"]
        chk["want_type"] = _want_typ or ""
        # PDF 那一格带 (ref) 标记 → 「应为」显示时要跟着写 (ref)。
        # 不写的话界面上「Excel 空着、应为 10」看着就是漏填却判了对 ——
        # 实际上这是参考值、规矩上本就不填，必须标出来（2026-09-29 用户点名）。
        chk["ref_marked"] = (_shape == "ref")
        # 类型 ↔ 两个「是/否」开关对不对得上（第 2.6 节 flag_contradiction）。
        # 放在最前面判：这是配置层的错，比数值更靠前。判到就直接结束这一条 ——
        # 类型/开关都不对时，往下比数值没有意义。
        _fc = flag_contradiction(er, _shape, op_raw=_op_raw, balloon=balloon)
        if _fc:
            chk["status"] = "mismatch"
            chk["scope"] = "type_flag"
            chk["expected"] = []
            chk["basis"] = _fc
            checks.append(chk)
            continue

        # 表头元信息（机器编号 / 程序号 / 保护套 / 芯轴）必须最先判：
        # 它的答案在参数表表头，而不是在参数行里。放在「采集项跳过」之后，
        # 字符串型的「机器编号」就会被整类跳过（这正是它一直没被核对的原因）。
        if pdf_meta is not None:
            mres = _meta_expected(er, meta_index, family)
            if mres is not None:
                chk["scope"] = mres["scope"]
                chk["meta_kind"] = mres["kind"]
                chk["basis"] = mres.get("basis")
                chk["expected"] = sorted(mres["vals"])
                act = _meta_actual(er, mres["kind"])
                if act is None:
                    # 这一项压根没填值（如「程序号」类型写布尔值、空着）——
                    # 先判「未填值」，不要去谈期望值存不存在，否则会把
                    # 「没填」报成「不一致」，冤枉人。
                    chk["status"] = "novalue"
                elif not mres["vals"]:
                    # 空期望值分两种，语义完全不同：
                    #   no_equip_match —— Excel 指名的设备在 PDF 里查无此机。
                    #     对「机器编号/程序号」这类项，这就是填错了
                    #     （「填了个 PDF 里不存在的设备号」正是最常见的错），必须判不一致。
                    #   no_meta —— PDF 表头压根没有这一项（版式不同或提取不到），
                    #     这不是 Excel 的错，报「无对应」即可。
                    chk["status"] = ("mismatch" if mres["scope"] == "no_equip_match"
                                     else "no_match")
                elif _meta_match(act, mres["vals"]):
                    chk["status"] = "ok"
                    # 「每一套候选都被现值命中」时，一致与否跟「哪一套候选才对」无关 ——
                    # 这种行不该混进「宽松判定明细」让人逐条去确认。
                    # 反过来，只要还有候选没被命中，那一套万一是对的就会翻成不一致，
                    # 那才是真正需要人看的一类。
                    if (chk["scope"] == "meta_multi" and mres["vals"]
                            and all(_meta_match(act, [v]) for v in mres["vals"])):
                        chk["scope"] = "meta_all"
                else:
                    chk["status"] = "mismatch"
                    chk["expected_kind"] = act[0]
                    chk["actual"] = act[1]
                checks.append(chk)
                continue

        # 采集项（字符串 / 布尔值）：按设计不参与规格核对，
        # 但要显式记下来 —— 以前是直接 continue，界面上看不出有这些行。
        if kind in ("str", "bool"):
            chk["status"] = "collection"
            chk["expected"] = []
            checks.append(chk)
            continue

        # 取值看「哪一列真的有值」，不是死认类型列：
        # 重填/录入时常见「类型列写固定值，值却填进了下限/上限」——
        # 老代码只读固定值列，于是判「未填值」整行跳过，错值就这样漏网。
        actual = None
        val_note = None
        one_sided = None
        lo, hi = _as_float(er["lo"]), _as_float(er["hi"])
        f = _as_float(er["fixed"])
        if lo is not None:
            actual = (lo, hi if hi is not None else lo)
            if hi is None:
                # 只填了下限、没填上限（与下面「只填上限」同源）。
                one_sided = ("lo", lo)
                val_note = "Excel 只填了下限 %s（上限未填）—— 只能按单侧核对" % _fmt_num(lo)
            elif kind == "fixed":
                val_note = "类型列写「固定值」，但值填在下限/上限列 —— 已按下限/上限读取"
        elif hi is not None:
            # **只填了上限、没填下限**。老代码这里没有分支，actual 保持 None，
            # 整行被当成「未填值」跳过 —— 可是 Excel 里明明写着 0.52。
            # 实测 WI-0039 的「压力（Mpa）」：Excel 上限 0.52、PDF「Pressure压力 0.52」、
            # 参数名也认得，三方齐全却报「未填值」，看的人只会以为工具没读出来。
            # 单侧限值照样可以核，只是可核的只有一侧，措辞里必须说清是哪一侧。
            actual = (hi, hi)
            one_sided = ("hi", hi)
            val_note = "Excel 只填了上限 %s（下限未填）—— 只能按单侧核对" % _fmt_num(hi)
        elif f is not None:
            actual = (f, f)
            if kind in ("num", "unknown"):
                val_note = "类型列写「%s」，但值填在固定值列 —— 已按固定值读取" % (er["type"] or "?")
        chk["_actual"] = actual
        chk["_one_sided"] = one_sided
        if val_note:
            chk["basis"] = val_note
        # Excel 侧状态（第 2.7 节）：value / na / blank / unreadable。
        xstate = excel_cell_state(er, kind, actual)

        if xstate == "unreadable":
            # 值列有字、但读不出可比较的数（「中间模编号」填 `T-2054-49` 这类）。
            # 这是**填了**，不是没填 —— 绝不能进缺失场景：那边一看到 PDF 侧有值
            # 就会判「漏填」，等于把「填了、我读不懂」说成「没填」（凭空造的假错）。
            chk["status"] = "novalue"
            # 打一个标记：这一行**是填了的**，只是填的不是数字（代码/文本），
            # 工具读不出可比较的值。界面/报告据此改口径 ——
            # 不能跟「Excel 空着没填」共用「未填值」四个字（2026-10-29）。
            chk["unreadable"] = True
            chk["expected"] = []
            chk["basis"] = ("Excel 填的是「%s」，本工具只能核对数值 —— 未参与核对"
                            % _show_cell(str(er.get("fixed") or er.get("lo")
                                             or er.get("hi") or "").strip()))
            checks.append(chk)
            continue

        if xstate in ("blank", "na"):
            # 单元格没值，但**不能一律当「未填值」** —— 是对是错完全取决于 PDF
            # 那一格。判定的唯一入口是第 2.7 节那张表（`decide_missing_cell`）：
            # 以前「Excel 空」和「Excel 填 N/A」各写一条 if/elif 链，同一句规矩
            # 改一遍漏一遍（实测「按类型不必填」在空着时判对、填 N/A 时判漏填）。
            res = (_step_expected(er, step_index, step_tables, blank_index,
                                  ordinal=ordinal_maps)
                   if step_index else None)
            st, sc, basis, exp = decide_missing_cell(
                er, res, pdf_shapes_for(er, step_index) if step_index else (),
                excel_state=xstate)
            chk["status"] = st
            chk["scope"] = sc
            chk["expected"] = exp
            if basis:
                chk["basis"] = basis
            _mark_ordinal(chk, res)
            checks.append(chk)
            continue

        # ref 参考值是**双向**的规矩：空着对，**填了错**。
        # 以前只实现了前一半（「所有空位不用填写」→ 空着判 type_exempt），
        # 后一半漏了 —— Excel 一旦填了值，就掉进下面的普通数值比对，
        # 数字对上照样判一致。WI-0020 表4!164「加热中间模/6(℃)」填了下限 115、
        # PDF 那格是 `115(ref)`，结果判成一致（2026-09-29 用户点名）。
        # 这条与 PDF 空格「不该填」（blank_filled）是同一类错：填了不该填的。
        # ⚠️ 2026-10-14 加球囊前提：这条「ref 不填」只对**球囊 WI** 生效；
        # 非球囊 WI 里 ref 照常填写，_shape 已被降级，ref_marked 也已是 False，
        # 这里再兜一层 balloon 条件，双保险防回归。
        if balloon and chk.get("ref_marked") and actual is not None and kind == "num":
            _res = (_step_expected(er, step_index, step_tables, blank_index,
                                   ordinal=ordinal_maps)
                    if step_index else None)
            _vals = (_res or {}).get("vals") or set()
            chk["status"] = "mismatch"
            chk["scope"] = "ref_filled"
            chk["expected"] = []
            chk["basis"] = ("PDF %s 这一格是 ref 参考值（%s），按类型规矩所有空位都"
                            "不填；Excel 却填了 %s —— 不该填"
                            % ((_res or {}).get("sub") or "该档参数表",
                               _fmt_vals(_vals, True) or "—",
                               _fmt_vals({actual}) or "—"))
            checks.append(chk)
            continue


        # 千分尺查找表（按直径 + 设备 + 产品线）
        kind = _micrometer_kind(er["item"])
        if kind and micrometer:
            res = _micrometer_expected(er, kind, micrometer, family)
            chk["scope"] = res["scope"]
            chk["basis"] = res.get("basis", "")
            chk["by_equip"] = res.get("by_equip") or {}
            expected = res["vals"]
            chk["expected"] = sorted(expected)
            if not expected:
                chk["status"] = "no_match"
            elif actual in expected:
                chk["status"] = "ok"
            else:
                chk["status"] = "mismatch"
        elif (pdf_precise is not None and _pn_core(er["equip"]) is not None
              and (precise_vals := _expected_values(er, pdf_precise)) is not None):
            # 参数表精确匹配（设备号 + 保护套 P/N + 尺寸）—— 三维度同时命中，
            # 收窄最狠，必须显式声明判定强度，否则界面上会落进「未标注」被当成宽松判定。
            #
            # 注意：**拿不到值时不在这一支判「无对应」**，而是落到下面的通用链
            # （横表 / 程序附表 / 竖表）继续判。老写法在这里直接 no_match，
            # 于是换一份没见过的版式，带 P/N 的行会整批变成「无对应」——
            # 表现上就是「对新文件识别不出来」，而通用链本来是判得了的。
            chk["scope"] = "precise"
            chk["basis"] = "按「设备号 + 保护套P/N + 尺寸」三维度精确命中该行的唯一取值区间"
            chk["expected"] = sorted(precise_vals)
            # 多套候选不拍板 —— 全流程统一口径，见下面竖表分支的长注释。
            if len(precise_vals) > 1:
                chk["multi"] = len(precise_vals)
                chk["status"] = "multi"
            else:
                chk["status"] = "ok" if actual in precise_vals else "mismatch"
        else:
            # 3) 步骤式横表匹配：Excel「步骤2的激光能量」 ↔ PDF「激光强度 / Step 2」
            res = (_step_expected(er, step_index, step_tables, blank_index,
                                  ordinal=ordinal_maps)
                   if step_index else None)
            if res:
                scope = res["scope"]
                chk["scope"] = scope
                chk["by_equip"] = res.get("by_equip") or {}
                if scope in ("strict", "approx", "union", "union_same", "device"):
                    step_expected = res["vals"]
                    chk["expected"] = sorted(step_expected)
                    # 多套候选不拍板（2026-10-29）。原来只有 scope=="device"
                    # 时记一个 `multi` 数，判定仍照「命中任一即 ok」—— 那正是
                    # 用户说的「两三套值一起判，一致性完全失效」。这里统一堵死：
                    # 只要还剩两套以上，就标成「多套候选，需人工确认」。
                    if len(step_expected) > 1:
                        chk["multi"] = len(step_expected)
                        chk["status"] = "multi"
                    elif actual in step_expected and _precision_mismatch(
                            er, actual, step_expected, res.get("val_dp")):
                        # 数值相等但小数位数不一致（3.0 vs 3）→ 判不一致。
                        # scope 保持原收窄层级（strict/union/device）不动 —— 它是
                        # 「期望值收窄到哪一层」的语义，不是「为什么错」；「为什么错」
                        # 用独立字段 precision_mismatch 标记，界面据此单独措辞。
                        # 依据必须说清「不是数值不对，是精度写法不一致」，否则
                        # 看的人会以为是数字填错了。
                        chk["status"] = "mismatch"
                        chk["precision_mismatch"] = True
                        chk["exp_dp"] = _dp_hint(res.get("val_dp") or {}, actual)
                        chk["basis"] = _fmt_precision_basis(
                            er, actual, step_expected, res.get("val_dp") or {})
                    else:
                        chk["status"] = "ok" if actual in step_expected else "mismatch"
                    if res.get("sub"):
                        chk["sub"] = res.get("sub")
                        chk["dia"] = res.get("dia")
                elif scope == "no_step":
                    # 该规格对应的参数表里根本没有这一步：填了值就是错，
                    # 正确做法是像 PDF 一样留空/N/A。
                    chk["status"] = "mismatch"
                    chk["expected"] = []
                    chk["basis"] = _nostep_basis(res, " —— 此栏应填 N/A")
                # ⚠️ 这里**没有** `no_param` 分支 —— 那一档不在这条链上产生。
                # 「PDF 该档没有这个采集项」是**兜底之后**的最后一问，挂在
                # `_apply_fallback` 之后（见 `_apply_no_param`）。塞在这里会让
                # 本该被兜底救回的行（如靠表下方备注取值的「充气速度」）被截胡：
                # R53 第一版就是这么写的，mismatch 当场从 4 条涨到 25 条。
                elif scope in ("pdf_blank", "pdf_na"):
                    # 与 no_step 不同：这一步在表里**有**，只是那一格没有值
                    # （写着 `------`、或明确写着 N/A）。表在声明「此处无值」，
                    # 填上值反而是错。
                    # 档位换成 blank_filled：不然会跟「PDF 为空且一致」混进同一档统计，
                    # 一个说对、一个说错，看着自相矛盾。
                    chk["scope"] = "blank_filled"
                    chk["status"] = "mismatch"
                    chk["expected"] = []
                    chk["basis"] = (("PDF %s 里这一格写的是「%s」（表格在声明此步不适用）"
                                     % (res.get("sub") or "该档参数表",
                                        _show_cell(res.get("raw") or "N/A"))
                                     if scope == "pdf_na" else
                                     "PDF %s 里这一格本身是空值（表格里写的是「%s」）"
                                     % (res.get("sub") or "该档参数表",
                                        _show_cell(res.get("raw") or "------")))
                                    + " —— 此栏应留空或填 N/A")
                elif scope == "no_spec_match":
                    chk["status"] = "no_match"
                    chk["expected"] = []
                    chk["basis"] = _nospec_basis(res, er)
                elif scope == "none":
                    # Excel 没给设备号 → 不给期望值，也不硬判对错，如实说明
                    chk["status"] = "no_scope"
                    chk["expected"] = []
                else:   # no_equip_match
                    chk["status"] = "no_match"
                    chk["expected"] = []
                _mark_ordinal(chk, res)
            else:
                # 4) 程序附表匹配（WI-0009 / WI-3852 型转置双列附表）
                res = None
                pair = card_ctxs.get(er.get("card"))
                if pair:
                    card, _rows = pair
                    idx = next((i for i, r in enumerate(_rows) if r is er), None)
                    ctx = card["ctxs"][idx] if idx is not None else None
                    if ctx:
                        hint = (card_hints.get(er.get("card")) or {}).get(idx)
                        if hint and hint.get("conflict"):
                            # 值本身都在候选里，但拼不出与 PDF 步骤顺序一致的取法 →
                            # 多半是把相邻两行（上提01/上提02）的值对调了
                            alt = _program_expected(er, ctx, card, prog_subs, None)
                            chk["scope"] = "prog_order"
                            chk["basis"] = ("该卡这几行的值与 PDF 的步骤先后不符："
                                            "按 PDF 步骤顺序推不出这组取值（疑似相邻行对调）")
                            chk["status"] = "mismatch"
                            chk["expected"] = sorted(alt["vals"]) if alt else []
                            checks.append(chk)
                            continue
                        res = _program_expected(er, ctx, card, prog_subs, hint)
                if res:
                    chk["scope"] = res["scope"]
                    chk["basis"] = res["basis"]
                    chk["expected"] = sorted(res["vals"])
                    # 多套候选不拍板，口径同上。
                    if len(res["vals"]) > 1:
                        chk["multi"] = len(res["vals"])
                        chk["status"] = "multi"
                    else:
                        chk["status"] = "ok" if actual in res["vals"] else "mismatch"
                    checks.append(chk)
                    continue
                # 5) 竖表路径（带适用范围收窄）
                # ⚠️ 传 `chk` 而不是 `er`：卡级元信息（程序名）只写在 chk 上
                # （见上面 card_progs 那一段）。传 er 的话程序号那一层根本看不见，
                # 静默跳过 —— 界面上「设备组」锁到了、期望值却还有 9 个。
                # 而**不能**反过来把 er 换成副本：上面第 4 步程序附表用
                # `r is er` 这个**身份**去定位本行在卡里的位置，换了副本就整条
                # 路径失效（WI-0009 实测：24 行 prog_strict 退回 multi）。
                res = _vertical_expected(chk, vindex, canonical_lookup, family)
                if res is None:
                    # 只有「参数名认得、但这份 PDF 里确实找不到它」才会走到这里
                    # （名字不认得会返回 name_unknown 那一档，设备对不上会返回
                    #   no_equip_match）。以前这里不写依据，界面上只剩一句「无对应」，
                    # 看的人没法判断是工具没读出来、还是 PDF 里真没有 ——
                    # WI-0039 的「开始温度/结束温度」就在这一档上被怀疑成识别故障。
                    #
                    # 措辞要把「三种都算正常」说全，这是用户 2026-09-24 定的口径：
                    # 「这种东西真没遇到的话也没有关系，你就标明出来就行了，
                    #   让人为知道就行。」—— 工具不去猜，只负责标明，人来判。
                    # 三种：① 写在 PDF 正文里（WI-0039 焗炉 75℃：正文有、参数表没有）
                    #      ② Excel 额外引入/默认的条件，PDF 本就不写
                    #        （DWG2002300 的「测试压力 14~14.5atm」挂 WI-3852，
                    #          值实际写在 WI-0010 的表1 里，WI-3852 里根本没有）
                    #      ③ 单位/写法对不上（如 Excel 写 atm、PDF 写 Mpa）
                    chk["status"] = "no_match"
                    chk["expected"] = []
                    _nm = ("参数名已收录，但这份 PDF 的参数表里没有它 —— "
                           "可能写在正文里，或本就是 Excel 额外引入的默认条件"
                           "（PDF 里不写），也可能是单位/写法对不上；请人工核一眼")
                    # ⚠️ 但**先查对账袋**：如果 PDF 里其实写了这个参数、只是我没能
                    # 解出数值（`SL1+3 ref` / `>Step5` / `中间模：135` 这类），那就
                    # 不能说「文件里没有」—— 那是把工具的读不懂说成文件里没有，
                    # 人会照着这句去改 Excel，越改越错。必须把原文摆出来。
                    _hits = _unparsed_hits(er, unparsed_index)
                    if _hits:
                        _where = "、".join(
                            sorted({"第%d页%s" % (h.get("page") or 0,
                                                 ("「%s」" % h.get("sub")) if h.get("sub") else "")
                                    for h in _hits[:3]}))
                        _what = "、".join(sorted({_show_cell(h.get("raw") or "")
                                                 for h in _hits[:3]}))
                        _nm = ("PDF 里**有**这一项，但我没读出数值 —— %s 那一格写的是"
                               "「%s」，我没解出可用的值（不是「文件里没有」）。"
                               "这类写法需要人工核一眼，别按「PDF 没有」去改 Excel"
                               % (_where, _what))
                    else:
                        # 图片块：文字层里没有 ≠ PDF 上没有。
                        # WI-0039 第 4 页的表7 整块是图片，印着 35/60/100/140/170/
                        # 220/270/330mm，和 Excel 那八行全对得上 —— 只是工具读不到
                        # （2026-11-02 用户自己看出来的：「看着像是图片类型的PDF」）。
                        # 这句必须补上，否则上面那句「PDF 的参数表里没有它」会被
                        # 读成「文件里没有」，人就会去改本来没错的 Excel。
                        _imgp = sorted(pdf_image_blocks(pdf_path)) if pdf_path else []
                        if _imgp:
                            _nm += ("。另外：这份 PDF 的第%s页印着图片块，图里的字"
                                    "没有文字层、工具读不到 —— 若这一项确实印在 PDF 上，"
                                    "多半就在那几页的图里，请对着原页面看一眼"
                                    % "、".join(str(p) for p in _imgp[:4]))
                    # 别把「取值时已经发现的问题」冲掉：比如
                    # 「类型列写浮点、值却填在固定值列 —— 已按固定值读取」。
                    # 原先这里是直接赋值，那句提示就没了（实测 ⑤ 号情形）。
                    # 但**单侧**行的取值提示放在下面统一收尾处拼（那里会按
                    # 判定结果改写措辞），这里必须排除，否则同一句会拼两遍。
                    _old = "" if chk.get("_one_sided") else (chk.get("basis") or "").strip()
                    chk["basis"] = (_old + "。" + _nm) if _old and _old != _nm else _nm
                    # 记下「这段是没找到时的说法」：锚点通道救回来之后要把这段摘掉，
                    # 不然会出现「已按你钉的位置读到值」＋「PDF 里没有它」两句并存。
                    chk["_nomatch_note"] = _nm
                else:
                    scope = res["scope"]
                    chk["scope"] = scope
                    chk["basis"] = res.get("basis", "")
                    chk["by_equip"] = res.get("by_equip") or {}
                    if scope in ("no_equip_match", "name_unknown"):
                        # name_unknown 与 no_equip_match 同样是「没找到」，但成因不同：
                        # 前者是词表缺条目（改代码），后者是这台机在 PDF 里没参数（改文档）。
                        # 都写成 no_match，但 scope 分开带走，界面据此各说各的话。
                        chk["status"] = "no_match"
                        chk["expected"] = []
                        # 同样是「没找到时的说法」，锚点救回来之后要摘掉（见 _apply_anchors）
                        chk["_nomatch_note"] = chk.get("basis") or ""
                    else:
                        chk["expected"] = sorted(res["vals"])
                        # ★★ 多套候选不再拍板（2026-10-29，用户定）
                        #
                        # 只要候选还剩两套以上，这一行判出来的「一致」就是抽奖：
                        # Excel 的值命中其中**任意一套**就算过，而真值完全可能是
                        # 另一套 —— 那不是结论，是概率。用户原话：
                        # 「多套对应值它的一致性就完全失效了……只能保留一套
                        #   对应值。你不能说你选那么两三个表格来。」
                        #
                        # 所以这里不猜、也不含糊地判 ok，如实标成「多套候选，
                        # 需人工确认」：候选套数、每套的出处都摆在依据里，
                        # 人一眼能看出该去核哪几张表。
                        # 收窄链已经收到**一套**的行完全不受影响 —— 那才是多数。
                        if len(res["vals"]) > 1:
                            chk["multi"] = len(res["vals"])
                            chk["status"] = "multi"
                        # 精度守卫（「3.0 与 3 不能算对」）—— 与横表分支同一份判据。
                        # 之前只有横表路径做了这一步，竖表（WI-0013/2025 焊接、
                        # 亲水涂层这类）填 3 代替 3.0 一律判一致，用户 2026-10-13
                        # 明确要求「书写一致才通过」。
                        elif actual in res["vals"] and _precision_mismatch(
                                er, actual, res["vals"], res.get("val_dp")):
                            chk["status"] = "mismatch"
                            chk["precision_mismatch"] = True
                            chk["exp_dp"] = _dp_hint(res.get("val_dp") or {}, actual)
                            chk["basis"] = _fmt_precision_basis(
                                er, actual, res["vals"], res.get("val_dp") or {})
                        else:
                            chk["status"] = "ok" if actual in res["vals"] else "mismatch"
                        # 竖表值的不等号方向（`> 0.52Mpa` → 下限约束）。
                        # 只在有方向时才写键 —— 绝大多数 PDF 值没有不等号，
                        # 走老口径。注意必须用非下划线键名：app 的显示层
                        # 读的是同一批 dict，但下划线开头的是引擎内部标记，
                        # 约定不对外。这里 one_side_wrong 才是对外的。
                        if res.get("side"):
                            chk["_pdf_side"] = res["side"]

        # 单侧限值的最终裁决（Excel 只填了下限或只填了上限）。
        # 上面每一条分支用的都是「元组完全相等」，对单侧填法是偏严的：
        # 详见 _one_sided_verdict 的说明。这里统一按单侧口径重判一次，
        # 并把「这个数是单侧读来的」写进依据 —— 不然看的人以为 Excel 给的是精确值。
        # 只在真的取到期望值时才改写；「无对应 / 无法判定 / 未填值」保持原样，
        # 因为那几种情况本来就没在比数值。
        if chk.get("_one_sided"):
            side, lim = chk["_one_sided"]
            side_cn = "上限" if side == "hi" else "下限"
            pdf_side = chk.get("_pdf_side")
            note = ("Excel 只填了%s %s（另一侧留空）—— 只有这一侧能核"
                    % (side_cn, _fmt_num(lim)))
            if pdf_side and pdf_side != side and chk.get("expected") and \
                    chk.get("status") in ("ok", "mismatch"):
                # PDF 明确写了不等号（如 `> 0.52Mpa` 是**下限**约束），
                # Excel 却把值填在了**另一侧**的列里 → 填错了侧。
                # 这正是用户 2026-09-27 点破的缺陷：值碰巧相同时（0.52==0.52），
                # 老口径会判「一致」，把填错列完全盖住 ——
                # 「上限 0.52 + 允许超上下限=是」等于任何值都能过，这格白填了。
                pdf_cn = "上限" if pdf_side == "hi" else "下限"
                sign_cn = "大于" if pdf_side == "lo" else "小于"
                chk["scope"] = "one_sided"
                chk["status"] = "mismatch"
                chk["one_side_wrong"] = True
                note = ("PDF 写的是「%s」约束（%s），%s 应填在「%s」列；"
                        "Excel 却把它填在「%s」列 —— 填错了侧"
                        % (sign_cn, _fmt_num(lim), _fmt_num(lim), pdf_cn, side_cn))
            elif chk.get("expected") and chk.get("status") in ("ok", "mismatch"):
                vd = _one_sided_verdict(chk.get("_actual"), chk["expected"], side)
                if vd == "same":
                    chk["scope"] = "one_sided"
                    chk["status"] = "ok"
                    note += "；PDF 的取值与它相同"
                    if pdf_side == side:
                        note += "，且 PDF 写的正是这侧的单边约束，方向一致"
                elif vd == "within":
                    chk["scope"] = "one_sided_loose"
                    chk["status"] = "ok"
                    note += "；PDF 的取值未超出该%s" % side_cn
                    if pdf_side == side:
                        note += "（PDF 写的正是这侧的单边约束，方向一致）"
                else:
                    chk["scope"] = "one_sided"
                    chk["status"] = "mismatch"
                    note += "；PDF 的取值超出了该%s" % side_cn
            old = (chk.get("basis") or "").strip()
            chk["basis"] = (note + "。" + old) if old and old != note else note
        checks.append(chk)

    # ——「适用规格」判定与门禁（用户 2026-09-28 定的规则）——
    # PDF 的参数表在**设备号下面**还写着适用规格（「T-1175激光焊接参数表(1.5-4.0mm)」
    # 「…(4.5-6.0)」）。它必须像「机器编号」一样单独核对，而且是步骤参数的前置门。
    checks = _apply_spec_gate(checks, step_tables)

    # —— 锚点通道（人钉过位置的行，按坐标直接读值）——
    # 必须在下面「分层回退」之前：人指的位置是最硬的依据，不该被启发式抢走。
    # 也要在算「对不上率」之前 —— 救回来的行不该继续算在错误率里。
    checks, anchor_stat = _apply_anchors(checks, pdf_path)

    # —— 分层回退匹配（第 7.5 节）——
    # 前提：PDF 是标准矢量版式时，抽出来的是「读到但匹配不上」，不是「没读到」。
    # 严格链判不出来的行，按 关键词 → 设备号+单位 → 纯数字/规格 → 值集近似 回退。
    # 第 4 层还要看整份的「对不上率」：> 30% 才启用（用户定的门槛）。
    tables = _pdf_tables(pdf_steps, pdf_vertical, pdf_programs)
    judged0 = sum(1 for c in checks
                  if c.get("status") in ("ok", "mismatch", "no_match", "no_scope"))
    failed0 = sum(1 for c in checks if c.get("status") in ("no_match", "no_scope"))
    rate = (failed0 / judged0) if judged0 else 0.0
    fuzzy_on = rate > _FUZZY_RATE
    checks, fb_stat = _apply_fallback(checks, excel_records, tables, fuzzy_on)
    # 兜底之后的最后一档救助：连分层匹配都救不回来的行，再问一次「PDF 该档
    # 有没有这个采集项」（见 `_apply_no_param`）。放在这里而不是主循环里，
    # 是因为它不能抢掉任何已有路径 —— 兜底能给出期望值的行优先级更高。
    np_stat = _apply_no_param(checks, sub_params)
    points = sum(len(t["recs"]) for t in tables.values())

    # ——「读到了但匹配不上」与「压根没读到」必须分开说 ——
    # 扫描件那种真读不出的情况由 pdf_scan_warning / pdf_text_warning 在界面上
    # 单独拦一道；这里说的是另一回事：PDF 抽出了 N 张表、M 个参数点（文字层
    # 正常），这一行仍然对不上 —— 那是**匹配**问题，不是读取问题。
    for c in checks:
        if c.get("status") != "no_match":
            continue
        if c.get("scope") == "spec_block":
            continue        # 适用规格门禁：另有说法，不混进来
        head = ("PDF 已读出 %d 张表 / %d 个参数点（文字层正常）—— "
                "本行是**匹配不上**，不是没读到" % (len(tables), points))
        if c.get("fb_best") is not None:
            head += ("；值集近似匹配度最高仅 %.0f%%（门槛 %d%%）"
                     % (c["fb_best"] * 100, int(_FUZZY_HIT * 100)))
        if c.get("fb_tables"):
            head += ("；匹配度达标的候选有 %s，彼此咬得太近，交给人定"
                     % "、".join(c["fb_tables"]))
        old = (c.get("basis") or "").strip()
        c["basis"] = (head + "。" + old) if old and head not in old else head

    # —— M8：匹配阶段与一致性校验阶段必须分开（规范第五条）——
    # 上面分层匹配只回答「对上了哪一张表」；**数值对不对是另一件事**，
    # 由这里按 status 单独给出状态。已经判出结果的行也补一份匹配元信息，
    # 界面才说得清「匹配成功」与「数据一致」不是一回事。
    for c in checks:
        if c.get("match"):
            if c.get("status") == "mismatch":
                c["match"]["state"] = M_DIFF
            continue
        st = c.get("status")
        if st in ("collection", "novalue") or c.get("spec_gate_row"):
            state = M_SKIP
        elif st == "mismatch":
            state = M_DIFF            # 对应上了，但数值不一致
        elif st == "ok":
            # 收窄到唯一的算精确匹配；多套候选/仅按名称命中的算辅助匹配
            state = (M_EXACT if c.get("scope") in
                     ("precise", "meta_strict", "meta_all", "spec_ok", "strict",
                      "union_same", "prog_strict", "one_sided", "step_down",
                      "no_step", "pdf_blank", "meta_list",
                      "anchor", "anchor_loose")
                     else M_ASSIST)
        elif st == "no_match":
            state = M_CONFIRM if c.get("fb_tables") else M_NONE
        else:
            state = M_NONE
        c["match"] = {"layer": 1 if state in (M_EXACT, M_DIFF) else 0,
                      "state": state, "score": 1.0 if state == M_EXACT else 0.0,
                      "why": "（严格链已判定，未经分层匹配）",
                      "candidates": [], "table": ""}

    counts = {"ok": 0, "mismatch": 0, "multi": 0, "no_match": 0,
              "collection": 0, "novalue": 0, "no_scope": 0}
    for c in checks:
        counts[c["status"]] = counts.get(c["status"], 0) + 1

    return {"checks": checks, "counts": counts,
            "read": {"tables": len(tables), "points": points,
                     "rate": round(rate, 3), "fuzzy": fuzzy_on, "fb": fb_stat,
                     "anchor": anchor_stat}}


def _pdf_spec_dias(eq, step_tables):
    """这台设备在 PDF 里出现过的**全部规格档** [(lo, hi)]（按步骤横表索引统计）。

    复用 `step_tables`（compare 已经建好的 (子表, 直径档) → 设备集合 索引），
    不另写一套收窄 —— 两处实现漂移比漏判更难查。
    """
    out = set()
    for _ck, tbl in (step_tables or {}).items():
        for (_sub, dia), info in tbl.items():
            if not dia:
                continue
            e = set(info.get("equip") or ())
            if not e:
                continue
            # ① 整组相同 ③ Excel 是 PDF 的子集，两场景都算这台设备的档
            if e == set(eq) or set(eq) < e or (e & set(eq)):
                out.add(tuple(dia))
    return sorted(out)


def _dia_cn(d):
    """(1.5, 4.0) → '1.5~4.0mm'；(6.0, 6.0) → '6.0mm'。"""
    if not d:
        return ""
    a, b = float(d[0]), float(d[1])
    if abs(a - b) < 1e-9:
        return "%smm" % _fmt_num(a)
    return "%s~%smm" % (_fmt_num(a), _fmt_num(b))


def _spec_gate(er, step_tables):
    """「一张参数卡 + 一台设备」的适用规格门禁结论。

    None = 无从判定（Excel 没给适用规格 / 没给设备号），不生成判定行、也不禁门。
    """
    eq = _equip_set(er.get("equip"))
    spec = _spec_range(er.get("spec"))
    if not eq or not spec:
        return None
    dias = _pdf_spec_dias(eq, step_tables)
    g = {"excel": str(er.get("spec") or ""), "excel_dia": spec, "pdf_dias": dias}
    if not dias:
        # PDF 压根没给这台设备分档 —— 无从判定，不禁门。
        # 新规则不能把「本来能核对的」变成「不能核对」，那是另一种事故。
        g["status"] = "unknown"
        g["hit"] = []
        return g
    full = [d for d in dias if _dia_score(spec, d) == 2]     # 规格整段落在这一档内
    if len(full) == 1:
        g["status"], g["hit"] = "ok", full
        return g
    if len(full) > 1:
        g["status"], g["hit"] = "vague", full
        return g
    part = [d for d in dias if _dia_score(spec, d) == 1]     # 沾边但没被完整包含
    if part:
        # 含糊：Excel 的规格跨在两档之间。不门禁，但把几档都列出来由人定夺 ——
        # 强行挑一档判下去，判出来的「一致」是没有根据的。
        g["status"], g["hit"] = "vague", part
        return g
    g["status"], g["hit"] = "mismatch", []
    return g


def _spec_gate_row(host, g):
    """把门禁结论做成一条**独立的判定行**（明细表 / 导出 / 交接包都会带上）。

    参数卡那一列留空：它是一道门，不属于某个采集项。
    """
    pdf = "、".join(_dia_cn(d) for d in (g.get("pdf_dias") or []))
    hit = g.get("hit") or []
    if g["status"] == "ok":
        scope = "spec_ok"
        basis = "Excel 适用规格落在 PDF 该设备的 %s 档内" % _dia_cn(hit[0])
    elif g["status"] == "vague":
        scope = "spec_vague"
        basis = ("Excel 适用规格同时挨着 PDF 的多档（%s）—— 离线定不了该用哪一档，"
                 "下面数值按含糊口径判，需人工确认" % pdf)
    else:
        scope = "spec_mismatch"
        basis = ("Excel 适用规格不在 PDF 该设备的任何一档（%s）内" % pdf)
    return {
        # —— 身份（2026-11-11 补）——
        # 原写法这四项全空，于是这条判定在「逐条对拍」里只能显示成
        # 「适用规格#3340」，看不出是哪份 WI 的、哪一行。它一变，人没法判断
        # 是「这份文件的规格档抽错了」还是「别的文件把它带歪了」——
        # 之前一度被当成「跨 WI 状态泄漏」查了半天，其实只是身份没写。
        "wi": host.get("wi"),
        "sheet": host.get("sheet"),
        "row": host.get("row"),
        "card": "",                       # 左侧留空：它是一道门，不挂在某个采集项下
        "item": "适用规格",
        "proc": host.get("proc") or "",
        "type": "字符串",
        "unit": "",
        "spec": g.get("excel", ""),
        "equip": host.get("equip"),
        "lo": None, "hi": None,
        "fixed": g.get("excel", ""),      # 「现值」= Excel 适用规格列的原文
        "btrue": None, "bfalse": None,
        "meta_kind": "适用规格",
        "status": "mismatch" if g["status"] == "mismatch" else "ok",
        "scope": scope,
        "expected": [_dia_cn(d) for d in (hit or g.get("pdf_dias") or [])],
        "basis": basis,
        "spec_gate_row": True,
    }


def _apply_spec_gate(checks, step_tables):
    """「适用规格」判定行 + 步骤参数的前置门。返回**新的** checks（行已插好）。

    用户 2026-09-28 定的规则：PDF 的参数表标题在设备号下面还写着适用规格
    （「T-1175激光焊接参数表(1.5-4.0mm)」「…(4.5-6.0)」），这一项必须像机器编号
    一样单独核对；**只有当它与 Excel「适用规格」列对上，才往下判步骤二/三**。

    三种结论：
      · **对上某一档** → 放行，步骤二/三照常判定（记下命中的那一档）；
      · **一档都对不上** → 门禁：本组步骤参数**不再判定**，如实报「未判定」并说清
        卡在哪一格 —— 规格都不对，下面的值对上了也只是巧合；
      · **PDF 没标规格** → 无从判定，**不禁门**也不出行，照常判。

    含糊（Excel 的直径跨在两档之间）时不门禁，判「一致」但档位标成含糊，
    让人一眼看出这一条要自己看。
    """
    cache = {}
    for chk in checks:
        eq = _equip_set(chk.get("equip"))
        if not eq or not chk.get("spec"):
            chk["spec_gate"] = None
            continue
        # 同一张卡 + 同一台设备 + 同一个适用规格 → 同一个结论。
        # 注意不能按「卡」分组：一张卡里并排放着 5~6 台设备（各 10 行），
        # 按卡分组只会算出第一台的档，然后套到所有设备上。
        key = (chk.get("card"), tuple(sorted(eq)), str(chk.get("spec")))
        if key not in cache:
            cache[key] = _spec_gate(chk, step_tables)
        chk["spec_gate"] = cache[key]

    # 门禁：只对**真正要核对数值**的那些行生效。机器编号这类元信息项不依赖
    # 规格档（它核对的是「这台设备在不在 PDF 里」），不能跟着一起停。
    for chk in checks:
        g = chk.get("spec_gate")
        if not g or g["status"] != "mismatch":
            continue
        if chk.get("meta_kind") or chk.get("status") in ("collection", "novalue"):
            continue
        why = (
            "适用规格「%s」不在 PDF 该设备的任何一档（%s）里 —— "
            "按规则本组参数不予判定，请先确认适用规格是否抄错"
            % (g["excel"] or "（空）",
               "、".join(_dia_cn(d) for d in g["pdf_dias"]) or "PDF 未标规格档"))
        if chk["status"] == "no_match":
            # 底层本来就判不出来（多半也是同一个规格原因）。状态不变，
            # 但**把根因写成适用规格** —— 否则这些行散在「无对应」里，
            # 看的人会以为要一处处去 PDF 里翻参数，其实要改的是适用规格那一格。
            chk["scope"] = "spec_block"
            chk["basis"] = why
            continue
        chk["status"] = "no_match"
        chk["scope"] = "spec_block"
        chk["expected"] = []
        chk["sub"] = None
        chk["basis"] = why

    # 独立判定行：紧跟在「机器编号」那一条下面。
    out = []
    for chk in checks:
        out.append(chk)
        g = chk.get("spec_gate")
        if chk.get("meta_kind") == "机器编号" and g and g["status"] != "unknown":
            out.append(_spec_gate_row(chk, g))
    return out



# ==========================================================================
#  第 7.5 节  分层匹配（Layered Matching）
#
#  规范（2026-09-29 版）要求的八个模块，各自独立、可单独加规则：
#    M1 文本标准化        _norm_text
#    M2 设备编号匹配      _equip_match
#    M3 参数名称匹配      _param_match
#    M4 单位与物理量类型   UNIT_TYPES / _unit_type / _unit_conv / _unit_match
#    M5 数值特征          _num_seq / _num_feature
#    M6 候选结果评分      _score_candidates
#    M7 匹配结果确认      _confirm_match（输出 7 种状态）
#    M8 最终一致性校验    在 compare() 里做 —— **不混进匹配阶段**
#
#  三条本项目的老规矩，在这一节同样生效：
#    ① 只对「严格链判不出来」的行介入，已判出结果的行一字不动（零回归靠这条）。
#    ② 第三层绑定表格必须 leave-one-out（排除本行），否则就是用本行的值自证。
#    ③ 高分 ≠ 正确：模糊匹配只负责「找出最可能的那张表」，
#       数值对不对由 M8 单独判，两者在界面上分开显示。
# ==========================================================================

def _cfg(name, default, cast=float):
    """读环境变量，读不到或读坏了就用默认值 —— 阈值不写死（规范第九条）。"""
    v = os.environ.get(name)
    if v is None or str(v).strip() == "":
        return default
    try:
        return cast(str(v).strip())
    except Exception:
        return default


# 第三层数值匹配的门槛。规范写「例如 80%」，上一轮用户口头说的是 50%~60%，
# 两者不一致 → 做成可配置，默认取 0.60（沿用上一轮），要改不用动代码。
_FUZZY_HIT = _cfg("WI_MATCH_MIN", 0.60)
# 第一名要甩开第二名这么多的绝对值，否则算「几张表都像」
_FUZZY_MARGIN = _cfg("WI_MATCH_MARGIN", 0.10)
# 整份「对不上率」（= no_match+no_scope / 可核项）超过多少才启用第 4 层
# 注意：它量的是**匹配**没成功的比例，不是「识别/OCR 出错」。
# WI-0039 表4 九条全对不上 → 100%，而那份 PDF 的 14 张表读得好好的。
# 2026-10-30 界面上曾把这一句写成「识别错误率」，已改掉。
_FUZZY_RATE = _cfg("WI_FUZZY_RATE", 0.30)
# 数值「近似相同」的容差：相对误差 ≤1% 或绝对差 ≤0.05
_REL_TOL = _cfg("WI_REL_TOL", 0.01)
_ABS_TOL = _cfg("WI_ABS_TOL", 0.05)

_XT_STEP = re.compile(r"[Ss]tep\s*#?\s*(\d+)")
_XT_TAIL = re.compile(r"/\s*(\d{1,2})\s*[）)]?\s*$")
_XT_NUM = re.compile(r"\d+(?:\.\d+)?")

# 7 种匹配状态（规范第五条）
M_EXACT = "精确匹配"        # 第一层：设备号/参数名等明确字段对上
M_ASSIST = "辅助匹配"       # 第二层：靠单位与物理量类型缩小候选后对上
M_FUZZY = "模糊匹配"        # 第三层：靠数值特征找到高相似候选
M_CONFIRM = "待人工确认"    # 多个候选分数接近 / 依据不足
M_NONE = "未找到匹配"       # 三层都没找到
M_DIFF = "匹配后存在差异"    # 对应上了，但数值不一致（匹配成功 ≠ 数据一致）
M_SKIP = "未参与匹配"       # 采集项 / 未填值等，本就不参与

# ---------------------------------------------------------------- M1 文本标准化
def _norm_text(s):
    """去首尾空格 → 全角转半角 → 统一大小写 → 去掉不影响含义的标点与空格。

    只做「不改变含义」的归一：不删数字、不删字母、不做同义替换
    （同义替换是 M3 的事，混在一起就没法追溯了）。
    """
    s = str(s or "")
    out = []
    for ch in s:
        o = ord(ch)
        if 0xFF01 <= o <= 0xFF5E:          # 全角 ASCII
            out.append(chr(o - 0xFEE0))
        elif o == 0x3000:                   # 全角空格
            out.append(" ")
        else:
            out.append(ch)
    s = "".join(out).strip().lower()
    s = re.sub(r"[\s\-_/\\·．.,，;；:：()（）\[\]【】「」\"'°]+", "", s)
    return s


def _norm_unit_key(u):
    k = _norm_text(u)
    return {"secs": "s", "sec": "s", "second": "s", "seconds": "s",
            "minute": "min", "minutes": "min", "hour": "h", "hours": "h",
            "millimetre": "mm", "millimeter": "mm", "centimeter": "cm",
            "metre": "m", "meter": "m", "degree": "℃"}.get(k, k)


# ------------------------------------------------------- M4 单位与物理量类型
UNIT_TYPES = {
    "mm": "长度", "cm": "长度", "m": "长度", "um": "长度", "mil": "长度",
    "mm/s": "速度", "cms": "速度", "ms_": "速度", "m/min": "速度",
    "rpm": "转速", "r/min": "转速",
    "℃": "温度", "c": "温度", "f": "温度", "℉": "温度", "°c": "温度", "celsius": "温度",
    "mpa": "压力", "kpa": "压力", "pa": "压力", "bar": "压力",
    "atm": "压力", "psi": "压力",
    "mj": "能量", "j": "能量",
    "w": "功率", "mw": "功率", "kw": "功率",
    "s": "时间", "min": "时间", "h": "时间", "ms": "时间",
    "v": "电压", "a": "电流", "ma": "电流",
    "g": "质量", "kg": "质量", "n": "力",
}

# 换算到该类型的基准单位（温度不是比例量，不做换算）
UNIT_FACTOR = {
    "长度": {"mm": 1.0, "cm": 10.0, "m": 1000.0, "um": 0.001, "mil": 0.0254},
    "时间": {"s": 1.0, "min": 60.0, "h": 3600.0, "ms": 0.001},
    "压力": {"pa": 1.0, "kpa": 1000.0, "mpa": 1000000.0,
             "bar": 100000.0, "atm": 101325.0, "psi": 6894.76},
    "能量": {"j": 1.0, "mj": 0.001},
    "功率": {"w": 1.0, "mw": 0.001, "kw": 1000.0},
    "速度": {"mm/s": 1.0, "cms": 10.0, "m/min": 16.6667},
    "质量": {"g": 1.0, "kg": 1000.0},
    "电流": {"a": 1.0, "ma": 0.001},
}


def _unit_type(u):
    """单位 → 物理量类型（"mm"→长度）。认不出返回空串 —— 空串不等于失败。"""
    return UNIT_TYPES.get(_norm_unit_key(u), "")


def _unit_conv(v, u_from, u_to):
    """单位换算。同类型且两边都有换算系数才换；换不了返回 None（绝不硬猜）。"""
    ta, tb = _unit_type(u_from), _unit_type(u_to)
    if not ta or ta != tb:
        return None
    if ta == "温度":
        return None
    fa = UNIT_FACTOR.get(ta, {}).get(_norm_unit_key(u_from))
    fb = UNIT_FACTOR.get(tb, {}).get(_norm_unit_key(u_to))
    if not fa or not fb:
        return None
    return v * fa / fb


def _unit_match(u_a, u_b):
    """(等级, 说明)：3=完全相同 2=同类型可换算 1=同类型但换不了 0=不同类型/-1=任一方缺。

    规范第三条第 4 点：相同单位**不能**单独认定是同一个参数 —— 所以这里只给
    「缩小候选范围」用的等级，最终判定不由它单独决定。
    """
    ka, kb = _norm_unit_key(u_a), _norm_unit_key(u_b)
    if not ka or not kb:
        return -1, "有一侧单位缺失（不据此处判定）"
    if ka == kb:
        return 3, "单位完全相同（%s）" % ka
    ta, tb = _unit_type(u_a), _unit_type(u_b)
    if ta and ta == tb:
        if _unit_conv(1.0, u_a, u_b) is not None:
            return 2, "同属%s，可换算（%s ↔ %s）" % (ta, ka, kb)
        return 1, "同属%s，但无明确换算关系（%s ↔ %s）" % (ta, ka, kb)
    return 0, "物理量类型不同（%s ↔ %s）" % (ta or ka, tb or kb)


# ----------------------------------------------------------- M2 设备编号匹配
def _equip_match(a, b):
    """(等级, 说明)：4=整组完全相同 3=包含/子集 2=有交集 1=一方无设备号 0=不相干。

    规范第六条：**设备数量不同不判错**。Excel 只写其中两台（子集）是常态，
    那是包含关系，不是错误；没给变更范围时，多出来的设备列为待确认项。
    """
    ea, eb = set(_equip_set(a)), set(_equip_set(b))
    if not ea or not eb:
        return 1, "有一侧没有设备号"
    if ea == eb:
        return 4, "设备组完全相同（%d 台）" % len(ea)
    if ea < eb or eb < ea:
        return 3, "包含关系（%d 台 ⊂ %d 台，数量不同不判错）" % (
            min(len(ea), len(eb)), max(len(ea), len(eb)))
    if ea & eb:
        return 2, "有交集（%d 台重合）" % len(ea & eb)
    return 0, "设备号不相干"


# ----------------------------------------------------------- M3 参数名称匹配
def _param_match(name_a, name_b):
    """(等级, 说明)：3=规范化后完全一致 2=同义词表同一条 1=互相包含 0=对不上。

    规范第二条：优先精确匹配，避免无必要的模糊 —— 所以「互相包含」只给 1 分，
    且只能当辅助特征，不能单独定案。
    """
    na, nb = _norm_text(name_a), _norm_text(name_b)
    if not na or not nb:
        return 0, ""
    if na == nb:
        return 3, "参数名一致（%s）" % str(name_a).strip()
    ca, cb = _canonical_param(name_a), _canonical_param(name_b)
    if ca and ca == cb:
        return 2, "参数名同义（都归到「%s」）" % ca
    if na in nb or nb in na:
        return 1, "参数名互相包含（%s ↔ %s）" % (str(name_a).strip(),
                                             str(name_b).strip())
    return 0, ""


# ------------------------------------------------------------- M5 数值特征
def _num_seq(vals):
    """把一堆 (lo,hi) / 数字摊平成数值序列，保留顺序。"""
    out = []
    for v in vals:
        if isinstance(v, (tuple, list)):
            for x in v:
                if isinstance(x, (int, float)):
                    out.append(float(x))
        elif isinstance(v, (int, float)):
            out.append(float(v))
    return out


def _num_eq(a, b):
    """完全相等"""
    return abs(a - b) < 1e-9


def _num_near(a, b):
    """近似相等：完全相等，或相对误差 ≤ _REL_TOL，或绝对差 ≤ _ABS_TOL。"""
    if _num_eq(a, b):
        return True
    if abs(a - b) <= _ABS_TOL:
        return True
    d = max(abs(a), abs(b))
    return (abs(a - b) / d) <= _REL_TOL if d else False


def _num_feature(seq_a, seq_b, ordered=True):
    """比较两组数值序列，返回 (分数, 明细)。

    规范 3.1：不能只比字符串，要综合
      · 相同的数量      → 逐位比，完全相同记 1
      · 相对误差        → 近似相同记 0.8
      · 序列长度        → min/max 作为长度相似度
      · 排列顺序        → ordered=True 按位对齐（步骤/位置有业务意义）
                          ordered=False 排序后对齐（集合匹配）
      · 分布特征        → 用「逐位比值是否稳定」近似体现，已含在逐位分里
    明细里逐个列明：哪些完全相同、哪些只是近似、哪些真的不一样 ——
    第三层必须让人看见差异具体在哪（规范第八条）。
    """
    a = list(seq_a)
    b = list(seq_b) if ordered else sorted(seq_b)
    if not ordered:
        a = sorted(a)
    if not a and not b:
        return 0.0, {}
    n = max(len(a), len(b))
    same, near, diff = [], [], []
    s = 0.0
    for i in range(n):
        x = a[i] if i < len(a) else None
        y = b[i] if i < len(b) else None
        if x is None or y is None:
            diff.append((x, y))
            continue
        if _num_eq(x, y):
            same.append((i, x))
            s += 1.0
        elif _num_near(x, y):
            near.append((i, x, y))
            s += 0.8
        else:
            diff.append((x, y))
    len_sim = (min(len(a), len(b)) / n) if n else 0.0
    base = s / n if n else 0.0
    score = 0.75 * base + 0.25 * len_sim
    return round(score, 4), {
        "same": same, "near": near, "diff": diff,
        "len": (len(a), len(b)), "ordered": ordered,
    }


# ------------------------------------------------- 表聚合（M6 之前的准备工作）
def _tail_step(name):
    """Excel「压力/7」「加热中间模/6(℃)」→ 7 / 6。末尾是「/数字」才认。

    英文版式（PDF 表头写 Step 4/5/6…）唯一能对齐的锚点。
    """
    m = _XT_TAIL.search(str(name or "").strip())
    return int(m.group(1)) if m else None


def _step_no(r):
    """PDF 记录的步骤号：字段优先，认不出再从参数名里读「Step N」。"""
    s = r.get("step")
    if s is None:
        m = _XT_STEP.search(str(r.get("param") or ""))
        s = int(m.group(1)) if m else None
    return s


def _nums_of_text(s):
    return {float(x) for x in _XT_NUM.findall(str(s or ""))}


def _spec_numbers(spec):
    """规格串里的数字：'φ2.0-3.5,φ4.0xQL(80-150)mm' → {2.0,3.5,4.0,80.0,150.0}"""
    return _nums_of_text(spec)


def _flatten_nums(x):
    """把任意嵌套的 (a,b) / [a,b] / 数字 摊平成 {float}。"""
    out = set()
    if x is None or isinstance(x, (str, bytes)):
        return out
    if isinstance(x, (int, float)):
        out.add(float(x))
        return out
    try:
        for i in x:
            out |= _flatten_nums(i)
    except TypeError:
        return out
    return out


def _units_of(r):
    u = r.get("units")
    if u:
        return set(u)
    one = _norm_unit(r.get("unit"))
    return {one} if one else set()


def _vals_of_recs(recs):
    """一组 PDF 记录的取值集合 → {(lo, hi)}。三种来源字段不同，这里统一。"""
    out = set()
    for r in recs:
        vs = r.get("vals")
        if vs:
            out |= {tuple(v) for v in vs}
            continue
        lo, hi = _as_float(r.get("lo")), _as_float(r.get("hi"))
        if lo is not None:
            out.add((lo, hi if hi is not None else lo))
        elif hi is not None:
            out.add((hi, hi))
        else:
            for v in _nums_of_text(r.get("text") or r.get("raw")):
                out.add((v, v))
    return {v for v in out if v[0] is not None}


def _pdf_tables(pdf_steps, pdf_vertical, pdf_programs):
    """把抽出来的记录按「一张表」聚合：{key: tbl}。

    tbl = {src, key, equip, units, names, steps, nums, specs, vals, recs}
    三种版式（步骤横表 / 竖表 / 程序附表）归到同一个结构里，后面的模块就
    不用关心自己问的是哪一种 —— 版式判断是抽取层的事。
    """
    out = {}

    def add(src, key, r, equip, units, names, steps, nums, vals, specs):
        t = out.get(key)
        if t is None:
            t = out[key] = {"src": src, "key": key, "equip": set(), "units": set(),
                            "names": set(), "steps": set(), "nums": set(),
                            "vals": set(), "specs": set(), "recs": []}
        t["equip"] |= equip
        t["units"] |= units
        t["names"] |= names
        t["steps"] |= steps
        t["nums"] |= nums          # 表里出现过的**参数值**
        t["specs"] |= specs        # 这张表**标称的规格/直径档**
        t["vals"] |= vals
        t["recs"].append(r)

    for src, lst in (("step", pdf_steps or ()), ("vert", pdf_vertical or ()),
                     ("prog", pdf_programs or ())):
        for r in lst:
            ck = _canonical_param(r.get("param"))
            names = {ck} if ck else set()
            st = _step_no(r)
            steps = {st} if st is not None else set()
            steps |= set(r.get("steps_all") or ())
            specs = _flatten_nums(r.get("dia")) | _flatten_nums(r.get("spec_ranges"))
            if src == "step":
                lo, hi = _as_float(r.get("lo")), _as_float(r.get("hi"))
                nums = {x for x in (lo, hi) if x is not None}
                vals = set()
                if lo is not None:
                    vals.add((lo, hi if hi is not None else lo))
                key = ("step", r.get("page"), r.get("table"), r.get("sub") or "",
                       tuple(r.get("dia") or ()))
                units = {_norm_unit(r.get("unit"))} - {None}
            elif src == "vert":
                vs = {tuple(v) for v in (r.get("vals") or ())}
                nums = {x for a, b in vs for x in (a, b) if x is not None}
                vals = vs
                key = ("vert", r.get("page"), r.get("table"), r.get("sub") or "",
                       tuple(r.get("dia") or ()))
                units = set(_units_of(r))
            else:
                nums = _nums_of_text(r.get("text") or r.get("raw"))
                vals = {(v, v) for v in nums}
                key = ("prog", r.get("page"), r.get("sub") or "",
                       r.get("program") or "", r.get("solution") or "")
                units = set()
            add(src, key, r, set(r.get("equip") or ()), units, names, steps,
                nums, vals, specs)
    return out


def _tbl_label(t):
    """表的只读名字，写进依据里让人能找到是 PDF 哪一页哪张表。"""
    src, a, b, c, d = t["key"]
    if src == "vert":
        return "第%s页第%s张表" % (a, b)
    if src == "step":
        return "%s%s" % (c or "第%s页第%s张表" % (a, b),
                         ("(%s)" % _dia_cn(d)) if d else "")
    return "附表%s%s" % (c, d)


def _fb_target(er, t, strict=False):
    """在已绑定的表里给这一行找**一一对应**的值（匹配阶段的最后一步）。

    顺序：参数名 → 步骤号(+单位) → 认不出就返回 None（绝不明猜）。
    strict=True 时只认「按参数名」这一种 —— 弱命中必须走到这一步才敢用。
    """
    ck = _canonical_param(er.get("item"))
    if ck:
        hit = [r for r in t["recs"] if _canonical_param(r.get("param")) == ck]
        vals = _vals_of_recs(hit)
        if vals:
            return vals, "按参数名「%s」对应" % ck
    if strict:
        return None, ""
    st = _tail_step(er.get("item"))
    if st is None:
        st = _item_step(er.get("item") or "")
    u = _excel_unit(er)
    if st is not None and st in t["steps"]:
        hit = [r for r in t["recs"] if _step_no(r) == st]
        if u:
            hu = [r for r in hit if u in _units_of(r)]
            if hu:
                hit = hu
        vals = _vals_of_recs(hit)
        if vals:
            return vals, "按步骤%d%s对应" % (st, ("＋单位「%s」" % u) if u else "")
    return None, ""


# ------------------------------------------------------------- M6 候选评分
#
# ⚠️ 这一节原本是**一串 if / continue**：六级判据连同各自的分值硬编码在函数
# 体里，加一档要读懂整段才知道该插在哪；层号 1/2/3 是魔法数，而 `_FB_NAME`
# 那张「层号 → 层名」表明明已经定义好了却**从来没被接上**（三处都查不到引用）。
# 2026-09-30 重构：判据抽成规则表，层号统一引用 `_FB_NAME` 的键。

class MatchRule(object):
    """分层匹配的一条**判据**（规范第五条：逐层降级）。

    tier  层号：1=文本字段精确 2=单位/物理量辅助 3=数值特征模糊（键见 `_FB_NAME`）
    hit   fn(ctx) -> bool    这条判据成不成立
    score fn(ctx) -> float   命中时的分数
    why   fn(ctx) -> str     依据文案（要能让人一眼看出是靠哪条命中的）

    ⚠️ 规则**按表里的顺序逐条试，命中第一条即定案**（短路）：
    已经能用上层可靠匹配的数据，绝不让下层覆盖。所以顺序不能随便调。
    """

    __slots__ = ("tier", "name", "hit", "score", "why")

    def __init__(self, tier, name, hit, score, why):
        self.tier = tier
        self.name = name
        self.hit = hit
        self.score = score
        self.why = why


class _TierCtx(object):
    """一张候选表 + Excel 这一行，能凑出的**全部信号**。

    规则表里的每条判据都只读这个对象 —— 判据因此与「信号怎么来的」解耦：
    将来要多用一路信号（比如产品线图号），只改这里，不用动规则表。

    `fuzzy` 做了惰性缓存：第三层要算 `_num_feature`，而分数与依据都要用它，
    不缓存就会对同一张表算两遍。
    """

    __slots__ = ("t", "e_lv", "e_why", "p_lv", "p_why", "u_lv", "u_why",
                 "spec_nums", "row_seq", "_fuzzy")

    def __init__(self, t, e_lv, e_why, p_lv, p_why, u_lv, u_why,
                 spec_nums, row_seq):
        self.t = t
        self.e_lv, self.e_why = e_lv, e_why
        self.p_lv, self.p_why = p_lv, p_why
        self.u_lv, self.u_why = u_lv, u_why
        self.spec_nums = spec_nums
        self.row_seq = row_seq
        self._fuzzy = None

    @property
    def fuzzy(self):
        if self._fuzzy is None:
            self._fuzzy = _num_feature(
                self.row_seq, _num_seq(sorted(self.t["nums"])), ordered=False)
        return self._fuzzy


def _fuzzy_why(ctx):
    sc, det = ctx.fuzzy
    return "数值序列相似度 %.0f%%（同%d/近%d/异%d）" % (
        sc * 100, len(det.get("same") or []),
        len(det.get("near") or []), len(det.get("diff") or []))


MATCH_TIER_RULES = (
    # —— 第一层：明确字段（设备号 / 参数名 / 规格）——
    MatchRule(1, "设备号＋参数名",
              lambda c: c.e_lv >= 3 and c.p_lv >= 2,
              lambda c: 0.90 + 0.02 * c.e_lv,
              lambda c: "设备号＋参数名：%s；%s" % (c.e_why, c.p_why)),
    MatchRule(1, "参数名命中",
              lambda c: c.p_lv >= 2,
              lambda c: 0.80,
              lambda c: "参数名命中：%s" % c.p_why),
    # —— 第二层：单位与物理量类型（辅助，**不能单独定案**）——
    MatchRule(2, "设备号＋规格档",
              lambda c: c.e_lv >= 3 and bool(c.spec_nums & c.t["specs"]),
              lambda c: 0.75,
              lambda c: "设备号相同＋规格档对上：%s；%s" % (
                  c.e_why,
                  "/".join(_fmt_num(x)
                           for x in sorted(c.spec_nums & c.t["specs"])[:4]))),
    MatchRule(2, "设备号",
              lambda c: c.e_lv >= 3,
              lambda c: (0.60 + 0.05 * c.u_lv) if c.u_lv > 0 else 0.60,
              lambda c: "设备号：%s%s" % (
                  c.e_why, ("；" + c.u_why) if c.u_lv > 0 else "")),
    MatchRule(2, "单位与物理量",
              lambda c: c.u_lv >= 2,
              lambda c: 0.55,
              lambda c: "单位与物理量：%s" % c.u_why),
    # —— 第三层：只看数值特征（忽略中英文名称与单位）——
    MatchRule(3, "数值特征",
              lambda c: bool(c.row_seq),
              lambda c: c.fuzzy[0],
              _fuzzy_why),
)


def _tier_ctx_for(er, t, row_seq):
    """凑出一张候选表的信号袋。"""
    e_lv, e_why = _equip_match(er.get("equip"), t["equip"] or set())
    p_lv, p_why = 0, ""
    ck = _canonical_param(er.get("item"))
    if ck and ck in t["names"]:
        p_lv, p_why = 3, "参数名「%s」在这张表里" % ck
    u_lv, u_why = -1, ""
    for u in (t["units"] or ()):
        lv, why = _unit_match(er.get("unit"), u)
        if lv > u_lv:
            u_lv, u_why = lv, why
    return _TierCtx(t, e_lv, e_why, p_lv, p_why, u_lv, u_why,
                    _spec_numbers(er.get("spec")), row_seq)


def _score_candidates(er, tables, other_rows):
    """给每张 PDF 表打分，返回 [(tbl, 层号, 分数, 说明)]，按 (层, -分) 升序。

    层号见 `_FB_NAME`：1=文本字段精确 2=单位/物理量辅助 3=数值特征模糊。
    判据全部在 `MATCH_TIER_RULES` 表里 —— 加一档改那张表即可，本函数不动。
    """
    out = []
    row_seq = _num_seq([r["val"] for r in other_rows]) if other_rows else []
    for _k, t in tables.items():
        ctx = _tier_ctx_for(er, t, row_seq)
        for rule in MATCH_TIER_RULES:
            if rule.hit(ctx):
                out.append((t, rule.tier, rule.score(ctx), rule.why(ctx)))
                break
    out.sort(key=lambda x: (x[1], -x[2]))
    return out


def match_tier_report():
    """把三级匹配的判据打出来 —— 给「版式自检」看，别再另写一套硬编码。"""
    out = []
    for tier in sorted(_FB_NAME):
        out.append("  第 %d 层（%s）：" % (tier, _FB_NAME[tier]))
        for r in MATCH_TIER_RULES:
            if r.tier == tier:
                out.append("      · %s" % r.name)
    return out


# ------------------------------------------------------------- M7 结果确认
def _confirm_match(ranked, strict_target=None):
    """把评分结果翻译成 7 种状态之一 + 决定用哪张表。

    返回 {state, layer, score, tbl, why, candidates, second}
    ranked 为空 → 未找到匹配。
    """
    if not ranked:
        return {"state": M_NONE, "layer": 0, "score": 0.0, "tbl": None,
                "why": "三层都没有找到候选表", "candidates": [], "second": None}
    top_layer = min(r[1] for r in ranked)
    top = [r for r in ranked if r[1] == top_layer]
    top.sort(key=lambda x: -x[2])
    best = top[0]
    second = top[1] if len(top) > 1 else None
    cands = [(_tbl_label(r[0]), r[1], round(r[2], 3)) for r in ranked[:4]]
    if top_layer == 1:
        state = M_EXACT
    elif top_layer == 2:
        state = M_ASSIST
    else:
        state = M_FUZZY
    # 第三层才看阈值与「是否甩开第二名」；前两层是明确字段命中，不设阈值。
    if top_layer == 3:
        if best[2] < _FUZZY_HIT:
            return {"state": M_NONE, "layer": 3, "score": best[2], "tbl": None,
                    "why": "数值相似度最高仅 %.0f%%，未达门槛 %d%%"
                           % (best[2] * 100, int(_FUZZY_HIT * 100)),
                    "candidates": cands, "second": second}
        if second and (best[2] - second[2]) < _FUZZY_MARGIN:
            return {"state": M_CONFIRM, "layer": 3, "score": best[2], "tbl": None,
                    "why": "多个候选分数接近（%s）—— 不替人选" % "、".join(
                        "%s %.0f%%" % (c[0], c[2] * 100) for c in cands[:3]),
                    "candidates": cands, "second": second}
    elif len(top) > 1 and abs(best[2] - (second[2] if second else 0)) < 1e-9:
        return {"state": M_CONFIRM, "layer": top_layer, "score": best[2],
                "tbl": None,
                "why": "同层有多个一样好的候选（%s）" % "、".join(
                    c[0] for c in cands[:3]),
                "candidates": cands, "second": second}
    return {"state": state, "layer": top_layer, "score": best[2], "tbl": best[0],
            "why": best[3], "candidates": cands, "second": second}


def _apply_fallback(checks, excel_records, tables, fuzzy_on):
    """分层匹配主入口。**只改 status ∈ {no_match, no_scope} 的行**。

    返回 (新的 checks, 统计)。
    """
    if not tables:
        return checks, {}
    stat = {"l1": 0, "l2": 0, "l3": 0, "amb": 0, "kept": 0}

    by_card = {}
    for i, c in enumerate(checks):
        by_card.setdefault(c.get("card"), []).append(i)

    for card, idxs in by_card.items():
        for i in idxs:
            chk = checks[i]
            if chk.get("status") not in ("no_match", "no_scope"):
                continue
            if chk.get("meta_kind") or chk.get("status") == "collection":
                continue      # 元信息项 / 采集项各有各的路，不在这里猜
            if chk.get("spec_gate_row"):
                continue
            # 同卡其它行（**排除本行**）—— 第三层绑定表格用
            others = []
            for j in idxs:
                if j == i:
                    continue
                r = checks[j]
                v = None
                for x in (_as_float(r.get("lo")), _as_float(r.get("hi")),
                          _as_float(r.get("fixed"))):
                    if x is not None:
                        v = x
                        break
                if v is None:
                    continue
                others.append({"st": (_tail_step(r.get("item"))
                                      or _item_step(r.get("item") or "")),
                               "val": v})
            ranked = _score_candidates(chk, tables, others)
            if not fuzzy_on:
                # 错误率没过门槛 → 不启用第三层（数值模糊）
                ranked = [r for r in ranked if r[1] <= 2]
            res = _confirm_match(ranked)
            chk["match"] = {"layer": res["layer"], "state": res["state"],
                            "score": round(res["score"], 3),
                            "why": res["why"],
                            "candidates": res["candidates"],
                            "table": _tbl_label(res["tbl"]) if res["tbl"] else ""}
            if not res["tbl"]:
                if res["state"] == M_CONFIRM:
                    stat["amb"] += 1
                    chk["fb_tables"] = [c[0] for c in res["candidates"][:3]]
                else:
                    stat["kept"] += 1
                    chk["fb_best"] = round(res["score"], 3)
                continue
            t = res["tbl"]
            # 弱命中（第二层、且分数低于 0.6）必须能按参数名一一对应
            vals, why = _fb_target(chk, t, strict=(res["layer"] == 2
                                                   and res["score"] < 0.6))
            if not vals:
                stat["kept"] += 1
                chk["match"]["state"] = M_NONE
                chk["match"]["why"] = (res["why"] + "；但表内定位不到"
                                       "一一对应的取值（参数名对不上）")
                continue
            act = chk.get("_actual")
            if act is None:
                stat["kept"] += 1
                continue
            chk["expected"] = sorted(vals)
            chk["status"] = "ok" if act in vals else "mismatch"
            chk["scope"] = {1: "fb_kw", 2: "fb_equip", 3: "fb_fuzzy"}[res["layer"]]
            # 匹配成功 ≠ 数据一致（规范第五条）：匹配层只说「找到了」，
            # 差多少由 M8（compare 的数值比较）单独说。
            if chk["status"] == "mismatch":
                chk["match"]["state"] = M_DIFF
            old = (chk.get("basis") or "").strip()
            lead = ("分层匹配·第%d层（%s）分数 %.0f%%：命中「%s」—— %s；表内%s"
                    % (res["layer"], res["state"], res["score"] * 100,
                       _tbl_label(t), res["why"], why))
            if res["layer"] == 3:
                lead += "；本行**没有参与**这次绑定（留一法）"
            chk["basis"] = (lead + "。" + old) if old and lead not in old else lead
            if len(vals) > 1:
                chk["multi"] = len(vals)
            stat["l%d" % res["layer"]] = stat.get("l%d" % res["layer"], 0) + 1
    return checks, stat


_FB_NAME = {1: "关键词/设备号", 2: "单位与物理量类型", 3: "数值特征"}


def _apply_no_param(checks, sub_params):
    """兜底 pass（第 53 轮）：连分层匹配都救不回来的 `no_match` 行，再问一次
    「这台设备这个规格档的参数表里，到底有没有这个采集项」。

    ⚠️ 必须放在**所有匹配路径之后**（含 `_apply_fallback`），绝不能提前 ——
    提前就会抢掉兜底的活。实测 WI-0036 card 0008 的「充气速度/10」：
    「充气速度」不在表1 的**列**里，但表下方备注写着「充气速率：一般为
    0.2-15」，`_enstep_notes` 专门为它记了一条。放在前面判，这条会被
    「表里没这一项」截胡、从 ok 翻成 mismatch —— 凭空造出 21 条假错。

    只对**仍是无对应**的行生效，且只认 scope=「规格套不上档」那一类：
    那一档原本的语义是「我不知道该看哪张表」，本 pass 把它细化成
    「看得出来是哪张表、而那张表里确实没有这一项」—— 于是可以判对。
    其余无对应（参数名未收录 / 设备查无此机 …）保持原样，绝不顺手扩大。
    """
    if not sub_params:
        return 0
    hit = 0
    for chk in checks:
        if chk.get("status") != "no_match":
            continue
        if chk.get("scope") != "no_spec_match":
            continue
        if chk.get("meta_kind") or chk.get("status") == "collection":
            continue
        # ⚠️ 只救「Excel 这一格没填值」的行。Excel 填了值而 PDF 该档没有这个
        # 采集项 → 那是**多填**，方向相反，绝不能在这里判对。
        # 这一档本轮先不接（维持 R52 的「无对应」），等用户确认口径再说 ——
        # 「不表态」永远好过「凭空造一个结论」。
        if chk.get("_actual") is not None:
            continue
        step, base = parse_step_item(chk.get("item"))
        if step is None:
            continue
        ck = _canonical_param(base)
        if not ck:
            continue
        gap = _param_gap(sub_params, ck,
                         _equip_set(chk.get("equip")), _spec_range(chk.get("spec")))
        if gap is None:
            continue
        # 判定的唯一入口仍是第 2.7 节那张表（`decide_missing_cell`）——
        # 依据文案、Excel 空/N/A 两种状态的一致性都由它统一管，这里不另写一套。
        st, sc, basis, exp = decide_missing_cell(
            chk, _no_param_res(gap), (),
            excel_state="na" if excel_says_na(chk) else "blank")
        chk["status"] = st
        chk["scope"] = sc
        chk["expected"] = exp
        if basis:
            chk["basis"] = basis
        hit += 1
    return hit


# ==========================================================================
#  第 8 节  批量核对
#  一次跑完目录下所有 WI
# ==========================================================================

# -*- coding: utf-8 -*-
"""批量核对：一次把目录下所有 WI 都跑一遍。

解决两个具体麻烦：

1. **要一个 WI 一个 WI 手点**。10 个 WI 就是 10 轮「选 PDF → 选 Excel → 点比对」。
2. **两份参数表的 WI 不重叠**：大表只有 WI-0021/2105/2106，小表只有
   WI-0007~2016/2038/2264/3852。切错了表就会「看着全对、其实只核了一半」，
   而界面上完全看不出来 —— 这和之前「工作表锁在表4、少看 300 倍」是同一类坑。
   批量模式对每份 PDF **自动在所有参数表里找它的数据**，找不到才如实报「无数据」，
   不会因为选错文件而静默漏掉。

内存：批量不保留全部明细（10 个 WI × 上万行会到几百 MB），
只留每个 WI 的计数 + 非「一致」的那些行（通常几百条），够用且轻。
"""


# 这些状态不算问题，批量汇总里不列明细
_QUIET = ("ok", "collection")


@dataclass
class BatchItem:
    wi: str = ""
    pdf: str = ""
    xlsx: str = ""
    sheet: str = ""
    counts: dict = dc_field(default_factory=dict)
    bad_rows: list = dc_field(default_factory=list)   # 已格式化的文字行（界面/报告用）
    bad: list = dc_field(default_factory=list)        # 结构化明细（导出 Excel 用）
    error: str = ""
    page_warning: str = ""          # PDF 文字质量提示（原先声明了却从没被写入）
    suspect: int = 0                # 抽取自检疑点数
    selfcheck_verdict: str = ""     # 抽取自检判词（顶部那一行）

    def judged(self):
        c = self.counts or {}
        return (c.get("ok", 0) + c.get("mismatch", 0), c)

    def verdict(self):
        c = self.counts or {}
        if self.error:
            return "失败：%s" % self.error[:40]
        if not c:
            return "未执行"
        if c.get("mismatch") or c.get("multi"):
            # 2026-11-09：多套候选并入「不一致」计数（用户：「多套就直接判错」）。
            # 判词里仍分开写「其中多套 N 条」，因为这两种的下一步动作不同：
            # 填错要改 Excel，多套要钉格把位置记下来。
            _mm = c.get("mismatch", 0) + (c.get("multi") or 0)
            _m = c.get("multi") or 0
            return "不一致 %d 条%s" % (_mm,
                                     ("（其中多套 %d 条）" % _m) if _m else "")
        # no_match / no_scope / novalue 的行**都不进** ok/mismatch（ok 只数判了
        # 对错的）。不把这几个数写出来，它们就会被「全部一致」吃掉 —— 而下面偏偏
        # 还逐条列着它们，等于自己拆自己的台（WI-0009 实测就是这样自相矛盾的：
        # 判定写「全部一致」，底下却列了 3 条）。
        # 用短词、放括号里：这一列只有 34 字符宽，写成整句会把右边的「← 表」挤走；
        # 每条具体是什么问题，下面那几行明细已经用中文写全了。
        rest = []
        if c.get("no_match"):
            rest.append("无对应 %d 条" % c["no_match"])
        if c.get("no_scope"):
            rest.append("缺设备号 %d 条" % c["no_scope"])
        if c.get("novalue"):
            rest.append("没填值 %d 条" % c["novalue"])
        tail = "（%s）" % "、".join(rest) if rest else ""
        # 多套候选已经并进上面的「不一致」计数了（2026-11-09），这里不再单列 ——
        # 留着会和新口径打架：上面说「不一致 2 条（其中多套 2 条）」，下面又补一句
        # 「多套候选待确认 2 条」，同一件事说两遍，还让人以为总数是 4。
        if c.get("ok"):
            return ("一致" + tail) if tail else "全部一致"
        # 「一条都没核」不能显示成「全部一致」—— 那正是这个工具要防的假安心。
        # 采集项全是字符串/布尔这类不参与规格核对的类型时，ok 会是 0。
        skipped = c.get("collection", 0)
        head = ("无规格类参数（%d 条均为字符串/布尔）" % skipped) if skipped else "无可核对项"
        return (head + tail) if tail else head


@dataclass
class BatchReport:
    items: list = dc_field(default_factory=list)
    xlsxs: list = dc_field(default_factory=list)
    notes: list = dc_field(default_factory=list)

    def total(self, key):
        return sum((it.counts or {}).get(key, 0) for it in self.items)

    def summary(self):
        out = [("批量核对 · %d 个 WI" % len(self.items), "t")]
        out.append(("  参数表 %d 份：%s"
                    % (len(self.xlsxs),
                       "、".join(os.path.basename(x)[:26] for x in self.xlsxs)), "dim"))
        out.append(("  合计：一致 %d ｜ 不一致 %d ｜ 多套待确认 %d ｜ 无对应 %d ｜ 未填值 %d"
                    % (self.total("ok"), self.total("mismatch"), self.total("multi"),
                       self.total("no_match"), self.total("novalue")), "h"))
        # 批量最危险的正是「十个都绿、其实有一个没读全」—— 判词必须提到合计后面
        n_sc = sum(1 for it in self.items if it.suspect)
        if n_sc:
            out.append(("  %d/%d 个 WI 的抽取自检有疑点 —— 这几个 WI 的「无对应」"
                        "可能混有「没读出来」的项，见下方逐条"
                        % (n_sc, len(self.items)), "warn"))
        else:
            out.append(("  抽取自检：%d 个 WI 全部未发现结构性疑点"
                        % len(self.items), "ok"))
        n_noverify = sum(1 for it in self.items if not (it.counts or {}).get("ok")
                         and not it.error)
        if n_noverify:
            out.append(("  其中 %d 个 WI 没有规格类参数可核（采集项全是字符串/布尔），"
                        "不代表「核对通过」" % n_noverify, "note"))
        out.append(("", None))
        for it in self.items:
            c = it.counts or {}
            if c.get("mismatch"):
                tag = "bad"
            elif it.error or c.get("no_match"):
                tag = "warn"
            elif not c.get("ok"):
                # 一条都没核：不能显示成绿色 OK，那正是这个工具要防的假安心
                tag = "note"
            else:
                tag = "ok"
            out.append(("  %-9s %-34s %s"
                        % (it.wi or "?", it.verdict(), ("← %s" % it.sheet) if it.sheet else ""),
                        tag))
            if it.error:
                out.append(("      " + it.error[:78], "warn"))
            if it.suspect:
                out.append(("      %s" % (it.selfcheck_verdict or "抽取自检有疑点"), "warn"))
            if it.page_warning:
                out.append(("      " + it.page_warning[:78], "warn"))
            if it.bad_rows:
                for line in it.bad_rows[:2]:
                    out.append(("      " + line[:78], "note"))
                if len(it.bad_rows) > 2:
                    out.append(("      …… 另有 %d 条" % (len(it.bad_rows) - 2), "dim"))
        out.append(("", None))
        for n in self.notes:
            out.append(("  · " + n, "dim"))
        return out

    def text(self):
        L = ["【批量核对报告】"]
        L.append("=" * 62)
        L.append("参数表：%s" % "、".join(os.path.basename(x) for x in self.xlsxs))
        L.append("合计：一致 %d ｜ 不一致 %d ｜ 多套待确认 %d ｜ 无对应 %d ｜ 未填值 %d"
                 % (self.total("ok"), self.total("mismatch"), self.total("multi"),
                    self.total("no_match"), self.total("novalue")))
        n_sc = sum(1 for it in self.items if it.suspect)
        if n_sc:
            L.append("%d/%d 个 WI 的抽取自检有疑点 —— 这几个 WI 的「无对应」可能"
                     "混有「没读出来」的项" % (n_sc, len(self.items)))
        else:
            L.append("抽取自检：%d 个 WI 全部未发现结构性疑点" % len(self.items))
        L.append("")
        for it in self.items:
            c = it.counts or {}
            L.append("%-9s %s" % (it.wi or "?", it.verdict()))
            L.append("    PDF   %s" % os.path.basename(it.pdf))
            if it.sheet:
                L.append("    数据  %s ｜ %s ｜ 记录 %d 行"
                         % (os.path.basename(it.xlsx)[:40], it.sheet, c.get("records", 0)))
            if c:
                L.append("    结果  一致 %d ｜ 不一致 %d ｜ 多套待确认 %d ｜ 无对应 %d ｜ 未填值 %d ｜ 非规格类跳过 %d"
                         % (c.get("ok", 0), c.get("mismatch", 0), c.get("multi", 0),
                            c.get("no_match", 0), c.get("novalue", 0),
                            c.get("collection", 0)))
            if it.error:
                L.append("    错误  %s" % it.error)
            if it.suspect:
                L.append("    抽取  %s" % (it.selfcheck_verdict or "有疑点"))
            if it.page_warning:
                L.append("    文字  %s" % it.page_warning)
            for line in it.bad_rows[:20]:
                L.append("      " + line)
            if len(it.bad_rows) > 20:
                L.append("      …… 另有 %d 条" % (len(it.bad_rows) - 20))
            L.append("")
        for n in self.notes:
            L.append("· " + n)
        return "\n".join(L)


def _fmt_num(x):
    """数值打印：整数不带小数点尾巴（18.0 → 18），其余原样。

    **全项目唯一的一份**（`grep -c "^def _fmt_num"` 必须 = 1）。
    2026-09-30 之前这里是一对同名函数，另一份在 10537 行 —— 那份才是死代码，
    已被删掉，详见删除处的注释。

    只认 `float`：非浮点一律 `str(x)` 原样返回。这是**已实测**的口径 ——
    全语料 21 份 PDF、29574 次调用，入参 100% 是 Python `float`，一次字符串
    都没有（探针 `/tmp/fmt_probe2.py`）。所以「字符串数字要不要转成 18」这种
    分歧在当前调用面上不存在；若将来新增调用点可能传字符串，先补断言再改。
    """
    if isinstance(x, float) and x == int(x):
        return str(int(x))
    return str(x)


# 判定词说人话：汇总里只有一列放下它，必须最短且自解释。
# 键是 scope，另有三个只看 status 的（no_match / no_scope / 兜底的 mismatch）。
_FLAG_CN = {
    "no_spec_match": "无对应（规格对不上任何一档）",
    "no_param": "不填正确（PDF 该档没有这个采集项）",
    "no_equip_match": "设备号在 PDF 里查无此机",
    "name_unknown": "参数名未收录（不在识别词典内）",
    "no_meta": "PDF 表头没有这一项",
    "missing": "漏填（PDF 有值、Excel 空）",
    "na_wrong": "漏填（PDF 有值、Excel 却填了 N/A）",
    # 按类型规矩不必填（如 PDF 给单一值、Excel 是浮点）→ 空着是对的，不算漏填
    "type_exempt": "不填正确（按类型规矩不必填）",
    "blank_filled": "多填（PDF 空、Excel 有值）",
    # ref 参考值：规矩是双向的 —— 空着对，填了错（表4!164 就是填了却判一致的那条）
    "ref_filled": "多填（PDF 是 ref 参考值、Excel 不该填）",
    "type_flag": "类型与「显示上下限 / 允许超上下限」开关不符",
    "prog_order": "与步骤顺序不符",
}


def _flag_cn(v):
    """把一条非 ok 判定的原因说成一句中文。原来这里直接印英文 scope 名
    （precise / no_meta / no_equip_match），看的人只能猜。"""
    st = v.get("status") or ""
    if st == "multi":
        # 候选还剩两套以上 —— 这不是「对」也不是「错」，是**工具判不了**。
        # 措辞必须把这一点说死：写成「不一致」会让人去改一个可能本来就对的
        # 单元格；写成「一致」就是拿多套值抽奖（2026-10-29 用户点名）。
        n = v.get("multi") or 0
        return ("多套候选（%d 套），工具无法唯一确定 —— 请人工核对"
                % n) if n else "多套候选，工具无法唯一确定 —— 请人工核对"
    if st == "novalue":
        # 「这一格没填值」时不去看 scope —— 一行本身就没值可核，再标成
        # 「设备号在 PDF 里查无此机」是把次要事实当成主因（WI-0009 的实测行
        # 就这么被标歪过）。单份界面里这类行显示的是「—（未填值）」，与此对齐。
        # 措辞要短：汇总里每条明细按 78 字符截断，标长一点整句就会被切掉半截。
        #
        # 但「填了文字/代码」不是「没填值」——那是我们没核，不是文件空着。
        # 两句话必须分开说，否则一张卡九十几行全印「未填值」，
        # 读的人只会以为这表是空的（2026-10-29 用户点名）。
        if v.get("unreadable"):
            return "未参与核对（填的是文字/代码）"
        return "未填值（Excel 空）"
    if st == "no_match":
        return _FLAG_CN.get(v.get("scope") or "", "无对应")
    if st == "no_scope":
        return "缺设备线索（Excel 没填设备号）"
    sc = v.get("scope") or ""
    if sc in _FLAG_CN:
        return _FLAG_CN[sc]
    if st == "mismatch":
        return "不一致"
    return STATUS_CN.get(st, st)


def _bad_line(v):
    """把一条非 ok 判定压成一行，供批量汇总显示 / 文本报告用。

    三个值必须**印对**。原来直接 %s 打原始字段，实测出来是这样：
        现值  ｜ 应为 (18.0, 22.0) ｜ precise
    三条全废：现值取错了键（判定字典里是 fixed / lo+hi / meta_kind，没有
    value）、期望值把 (下限, 上限) 元组原样印出来、判定列印的是英文 scope 名 ——
    看的人只会以为工具坏了。

    注意：这里是**给汇总一行用的最简格式化**，不是把 app 的 _value_str /
    _fmt_expected 又抄一份（那两个要处理「按设备分组列候选、元信息候选上限、
    候选真的不止一个」等十几种情形，且只该在 GUI 侧）。若将来真要合一，
    正确方向是把它们**下移到 wi_core** 供两边共用，而不是两边各留一份。
    """
    if v.get("lo") is not None:
        val = "%s ~ %s" % (_fmt_num(v["lo"]), _fmt_num(v["hi"]))
    elif v.get("fixed") is not None:
        val = _fmt_num(v["fixed"])
    elif v.get("meta_kind") and v.get("equip"):
        # 元信息项（机器编号等）没填在任何数值列里，它的值就是该行设备名
        val = str(v["equip"])
    else:
        val = "—"
    if v.get("unit") and val != "—":
        val += str(v["unit"])
    # 期望值：数值型是 [(下限, 上限), …]，元信息型是 ["T-1175", …]
    exp = v.get("expected")
    if isinstance(exp, (list, tuple, set)) and exp:
        parts = []
        for one in exp:
            if isinstance(one, (list, tuple)) and len(one) == 2:
                a, b = one
                parts.append(_fmt_num(a) if a == b
                             else "%s~%s" % (_fmt_num(a), _fmt_num(b)))
            else:
                parts.append(str(one))
        exp_s = "、".join(parts)
    else:
        exp_s = "—"
    return "%s · %s ｜ 现值 %s ｜ 应为 %s ｜ %s" % (
        v.get("card", ""), v.get("item", ""), val, exp_s, _flag_cn(v))


def run_batch(pdf_paths: list, xlsx_paths: list, progress=None,
              keep_bad: int = 400) -> "BatchReport":
    """逐个 PDF 核对：自动在所有参数表里找数据。

    progress: 可选回调 (i, total, text)，用于界面显示进度。
    """
    rep = BatchReport(xlsxs=list(xlsx_paths))
    total = len(pdf_paths)
    for i, pdf in enumerate(pdf_paths, 1):
        if progress:
            progress(i, total, os.path.basename(pdf))
        it = BatchItem(pdf=pdf)
        try:
            it.wi = detect_wi(pdf) or ""
        except Exception as e:
            it.error = "识别 WI 失败：%s" % e
            rep.items.append(it)
            continue
        if not it.wi:
            # ⚠️ 认不出编号时**必须停手**，不能继续往下走。
            # `extract_excel_params(x, "")` 里 `wi_name` 是空串 → 关掉 WI 过滤
            # （见 records() 的 `if wi_name and wi != wi_name`）→ 于是拿**整张卡**
            # 的采集项去比一份不知道是谁的 PDF，报出一大片 no_match。
            # 那看上去像「PDF 里没写这些参数」，实际是「没认出这是哪一份 WI」——
            # 两种结论对使用者的下一步动作完全相反，绝不能混。
            it.error = ("认不出这份 PDF 的 WI 编号（文字层里没有「文件编号」，"
                        "文件名里也没有 WI 编号）—— 不知道该取哪一批采集项来核，"
                        "已跳过。这不是「数据对不上」")
            rep.items.append(it)
            continue

        hit = None
        for x in xlsx_paths:
            try:
                recs, ts = extract_excel_params(x, it.wi, with_report=True)
            except Exception:
                continue
            if recs:
                hit = (x, ts, recs)
                break
        if hit is None:
            it.error = ("所有参数表里都没有这个 WI 的采集项 —— "
                        "可能是本次没提供它的参数表")
            rep.items.append(it)
            continue

        x, ts, recs = hit
        it.xlsx = x
        it.sheet = ts.chosen.label if ts.chosen else ""
        try:
            reset_pdf_failures()
            # 缺口 B：批量恰恰是「一次跑十个、没人在旁边盯着」的场景 —— 抽取自检
            # 与文字质量提示必须跟着跑，否则单份比对看得到的那层信号一跑批量就
            # 整个丢掉。必须放在 release_caches 之前：三者共用同一批缓存页面，
            # 放到 finally 之后等于把每份 PDF 重新解析一遍（10 份慢一倍）。
            sc = self_check(pdf)
            it.suspect = int(sc.get("suspect") or 0)
            it.selfcheck_verdict = (sc.get("lines") or [""])[0]
            it.page_warning = pdf_text_warning(pdf)
            # 词库（可选）：批量也要跟单份同一口径，否则同一份文件两条路的
            # 名字对齐结果不一样。词库为空时这步等于没做。
            try:
                import name_lex as _nl
                _nl.push_to_core(it.wi, doc_family(x))
            except Exception:
                pass
            try:                       # 产品形态（RX/OTW），供附表页维度收窄
                import page_tier as _pt
                set_card_tokens(_pt.card_tokens(x))
            except Exception:
                set_card_tokens(())
            # 关键字传参：compare 有 10 个位置参数，顺序被打乱会静默错位
            r = compare(
                pdf_by_name=extract_pdf_params(pdf),
                excel_records=recs,
                micrometer=extract_pdf_micrometer(pdf),
                pdf_precise=extract_pdf_params_precise(pdf),
                pdf_steps=extract_pdf_steps(pdf),
                pdf_programs=extract_pdf_programs(pdf),
                pdf_vertical=extract_pdf_vertical(pdf),
                family=doc_family(x),
                pdf_meta=extract_pdf_meta(pdf),
                pdf_step_blanks=extract_pdf_step_blanks(pdf),
                pdf_path=pdf)
        except Exception as e:
            it.error = "核对失败：%s" % e
            rep.items.append(it)
            continue
        finally:
            release_caches()

        c = dict(r["counts"])
        c["records"] = len(recs)
        it.counts = c
        for v in r["checks"]:
            if v.get("status") in _QUIET:
                continue
            if len(it.bad_rows) < keep_bad:
                it.bad_rows.append(_bad_line(v))
                exp = v.get("expected")
                if isinstance(exp, (list, tuple, set)):
                    exp = "/".join(str(x) for x in sorted(map(str, exp)))
                it.bad.append({
                    "wi": it.wi,
                    "card": v.get("card", ""),
                    "item": v.get("item", ""),
                    "type": v.get("type", ""),
                    "value": v.get("value", v.get("got", "")),
                    "expected": "" if exp is None else exp,
                    "unit": v.get("unit", ""),
                    "status": status_cn_of(v),
                    "scope": v.get("scope", ""),
                })
        rep.items.append(it)

    rep.notes.append("每个 WI 的数据是从全部参数表里自动找的 —— 两份表的 WI 通常不重叠，"
                     "手动选表很容易只核到一半。")
    # —— 卡里引用了、但本次没给 PDF 的 WI，逐个点名 ——
    # 不点名的话，这类只会表现为「所有参数表里都没有这个 WI 的采集项」，
    # 而那句话反过来也能读成「PDF 没提供」—— 使用者分不清是补 PDF、
    # 还是改参数表，只能一张张翻（用户 2026-10-13 反馈的「无法校验」里，
    # 有一部分就是这个：现行目录里根本没有那份 PDF，共 77 种 WI / 87 张卡）。
    _got = {it.wi for it in rep.items if it.wi}
    _miss = {}
    for x in xlsx_paths:
        try:
            names = TableSchema.load(x).wi_names()
        except Exception:
            continue
        for w in names:
            w = _norm_wi_no(w)
            if w and re.match(r"^WI-", w) and w not in _got:
                _miss.setdefault(w, []).append(os.path.basename(x))
    if _miss:
        _keys = sorted(_miss)
        rep.notes.append(
            "这些 WI 参数卡里有、但本次没提供对应 PDF（共 %d 个）：%s"
            % (len(_keys),
               "、".join(_keys[:20]) + ("…等" if len(_keys) > 20 else "")))
    rep.notes.append("内存考虑：明细只保留非「一致」的行；完整逐行结果请用单份比对导出。")
    return rep
