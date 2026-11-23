# -*- coding: utf-8 -*-
"""page_tier.py —— 「附表页」的结构维度通道（长度档 / 产品形态 / 附表号）

为什么需要它：
    WI-0042 的 12 张设备参数附表，标题是
        （For NEO RX-90cm，Balloon Length 15-80mm）
        （For NEO RX-150/200cm，Balloon Length 15-80mm）
        （For NEO OTW-90cm，Balloon Length 15-80mm）
        （For NEO OTW-150cm，Balloon Length 15-80mm）
    也就是 **4 组（RX-90 / RX-150·200 / OTW-90 / OTW-150）× 3 个长度档**。
    这四个组的「位置」值域完全不重叠（360~1050 / −15~450 / 425~1070 / −15~460），
    而**速度四组全同**（8/20/50）—— 所以速度从来没暴露问题，位置一收就是十几套。

    可这四个维度**在抽取时根本没进记录**：竖表记录里 `sub=''`、`prog=''`、`dia=None`。
    收窄链拿不到「这条属于哪张附表、哪个长度档」，于是只能把 12 页的候选并成一个集合。

本模块只做三件事（都在抽取之后、判定之前，**不碰任何版式函数**）：
    1. `page_dims(pdf)`  逐页解析标题 → {页码: {sub, len, forms}}
    2. `attach(recs, pdf)` 把上面三个维度就地补到该页每条记录上
       （`_pg_sub` / `_len_tier` / `_pg_rx`；解析不到的页不加，下游据此跳过收窄）
    3. `spec_dims(spec)` 把参数卡的 `1.25 X 15 90cm` 拆成 直径 / 球囊长度 / 有效长度(cm)
       以及 `card_tokens(xlsx)` 从「表2 大类」取产品形态（RX / OTW）

收窄规则由 `wi_core._page_dim_filter` 使用，三条保险：
    · 只在**本行同名候选全都带维度**时才动手（有一条没解析出来 → 原样放过）；
    · 收窄后为空 → 原样放过（宁可多套候选，也不误杀）；
    · 产品形态只在**卡片大类唯一**时才用（卡片同时含 RX 与 OTW 就弃权）。
"""
import os
import re

# （For NEO RX-90cm，Balloon Length 15-80mm）
_TITLE_PAREN = re.compile(r"[（(]\s*([^（）()]*(?:Balloon|球囊长度)[^（）()]*)[）)]", re.I)
_LEN = re.compile(r"(?:Balloon\s*Length|球囊长度)\s*[:：]?\s*"
                  r"(\d+(?:\.\d+)?)\s*[-~－—至]\s*(\d+(?:\.\d+)?)\s*(mm|cm)?", re.I)
_FORM = re.compile(r"\b(RX|OTW)\s*[-－—]\s*(\d+(?:\s*/\s*\d+)*)\s*(cm|mm)", re.I)
_SUBNO = re.compile(r"附表\s*(\d+)")

_DIMS_CACHE = {}
_TOKENS_CACHE = {}
_SPEC_CACHE = {}


# ------------------------------------------------------------------ 1. 页面维度
def _parse_title(line):
    """从一行标题里抽 (长度档, [(形态, {cm…})])。抽不到的部分给 None / ()。"""
    ln = None
    m = _LEN.search(line)
    if m:
        lo, hi = float(m.group(1)), float(m.group(2))
        if (m.group(3) or "").lower() == "cm":      # 长度用 cm 写的不认，避免与 mm 混
            lo, hi = None, None
        if lo is not None and lo > hi:
            lo, hi = hi, lo
        if lo is not None:
            ln = (lo, hi)
    forms = []
    for fm, cms, u in _FORM.findall(line):
        vals = set()
        for x in re.split(r"\s*/\s*", cms):
            try:
                v = float(x)
            except ValueError:
                continue
            # 统一按 cm 存（与参数卡第三段同口径）：cm 原样，mm 折成 cm
            vals.add(v if u.lower() == "cm" else round(v / 10.0, 3))
        if vals:
            forms.append((fm.upper(), frozenset(vals)))
    return ln, forms


def page_dims(pdf_path):
    """{页码: {"sub": int|None, "len": (lo,hi)|None, "forms": [(RX,{90.0})], "raw": str}}

    逐页解析，解析不到任何维度的页不进结果 —— 下游据此判断「这页没有维度」。
    """
    key = os.path.abspath(str(pdf_path))
    if key in _DIMS_CACHE:
        return _DIMS_CACHE[key]
    out = {}
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            for pno, page in enumerate(pdf.pages, 1):
                try:
                    txt = page.extract_text() or ""
                except Exception:
                    continue
                heads = [ln for ln in txt.splitlines() if _TITLE_PAREN.search(ln)]
                if not heads:
                    continue
                ln, forms = None, []
                for h in heads:
                    a, b = _parse_title(h)
                    ln = ln or a
                    forms = forms or b
                if ln is None and not forms:
                    continue
                m = _SUBNO.search(txt)
                out[pno] = {"sub": int(m.group(1)) if m else None,
                            "len": ln, "forms": forms,
                            "raw": heads[0][:80]}
    except Exception:
        out = {}
    _DIMS_CACHE[key] = out
    return out


def attach(records, pdf_path):
    """把维度就地补到记录上。返回补了多少条（0 = 这份 PDF 没有这类附表）。"""
    if not records or not pdf_path:
        return 0
    dims = page_dims(pdf_path)
    if not dims:
        return 0
    n = 0
    for r in records:
        d = dims.get(r.get("page"))
        if not d:
            continue
        if d["len"]:
            r["_len_tier"] = d["len"]
        if d["forms"]:
            r["_pg_rx"] = tuple(d["forms"])
        if d["sub"] is not None:
            r["_pg_sub"] = d["sub"]
        n += 1
    return n


# ⚠️ 只给界面看、**不参与判定**的一句人话：「这一页是哪一档」。
#
# 用户原话（2026-11-02）：「我在选这个表的时候……它根本没有表以外的数据信息。
# 那我根本不知道这个规格参数是哪一页，哪张表，我还得去翻原 PDF 里面的。」
# 那 12 张附表的页眉只写「设备参数附表1…12」，光看表名根本分不出 RX-90 和 OTW-150；
# 唯一的线索就是标题那一行括号里的 `For NEO RX-90cm，Balloon Length 15-80mm`。
_SPEC_PAREN = re.compile(r"[（(]([^（）()]{4,120}(?:Balloon|球囊长度|RX|OTW)[^（）()]{0,80})[）)]", re.I)


def spec_text(pdf_path):
    """{页码: 表外那句规格原文}，例如
    8 → 「For NEO RX-90cm，Balloon Length 15-80mm」。

    抓的是**表格框外面**那一行标题（`page_dims` 里已经存下来了），
    所以不会把表里的「参数值或范围（N，RX90＜80）」误当成页规格。
    """
    key = os.path.abspath(str(pdf_path))
    if key in _SPEC_CACHE:
        return _SPEC_CACHE[key]
    out = {}
    for pno, d in page_dims(pdf_path).items():
        raw = d.get("raw") or ""
        m = _SPEC_PAREN.search(raw)
        if m:
            out[pno] = m.group(1).strip()
        elif raw:
            out[pno] = raw.strip()[:60]
    _SPEC_CACHE[key] = out
    return out


# ------------------------------------------------------------------ 2. 参数卡侧
#
# 参数卡上的「适用规格」有两种写法，必须都认，否则附表页那 12 张表收不窄
# （2026-11-02 用户点名：WI-0042 又出 36 条多套候选）：
#
#   ① 老写法（球囊规格）  `1.25 X 15 90cm` / `φ1.25xQL15` / `φ1.5xQL(15-150)`
#        → 直径、球囊长度、有效长度(cm)
#   ② 新写法（导管规格）  `QL(15-80)xCL90` / `QL(100-150)xCL(150,200)`
#        → **没有直径**；QL(...) 是球囊长度区间，CL... 是导管有效长度（单位 cm 省略不写）
#
# ② 这一种以前落进了 ① 的通用数字扫描里：`QL(15-80)xCL90` → nums=[15,80,90]、
#    没有 "cm" 字样于是 rx=None，长度档只拿到 80 这个**点**，形态/导管长度那一层
#    干脆整段跳过 —— 结果 12 张附表一张都没被排除，收窄形同虚设。
_QL_RANGE = re.compile(r"QL\s*[（(]?\s*(\d+(?:\.\d+)?)\s*[-~－—至]\s*(\d+(?:\.\d+)?)", re.I)
_QL_ONE = re.compile(r"QL\s*[（(]?\s*(\d+(?:\.\d+)?)\s*[）)]?", re.I)
_CL_LIST = re.compile(r"CL\s*[（(]\s*(\d+(?:\.\d+)?(?:\s*[,，/、]\s*\d+(?:\.\d+)?)*)\s*[）)]", re.I)
_CL_ONE = re.compile(r"CL\s*(\d+(?:\.\d+)?)", re.I)
_NUMU = re.compile(r"(\d+(?:\.\d+)?)\s*(cm|mm)?")


def spec_dims(spec):
    """参数卡「适用规格」→ 三个维度（解析不到的部分给 None / 空集）。

    `1.25 X 15 90cm`          → dia=1.25, len=(15,15), rx_cms={90}
    `φ1.5xQL(15-150)`         → dia=None, len=(15,150), rx_cms=∅
    `QL(15-80)xCL90`          → dia=None, len=(15,80),  rx_cms={90}
    `QL(100-150)xCL(150,200)` → dia=None, len=(100,150), rx_cms={150,200}

    ⚠️ QL 写法一律 dia=None —— 那种规格里**根本没有直径**（只有球囊长度区间
    和导管有效长度）。φ1.5 那个前缀在真实参数卡里也常被省略，硬抽反而会
    抽到区间里的数字。`dia` 这条线目前无人消费（`_page_dim_filter` 只用
    len / len_hi / rx_cms），所以留空比猜错安全。

    返回 dict：
      dia     直径（老写法才有）
      len     球囊长度**下界**（老口径就是那个单值，兼容原来的调用）
      len_hi  球囊长度**上界**（单值时与 len 相同）
      rx_cms  导管有效长度的集合（cm）
      rx_cm   兼容旧字段：集合里最小的那个
    """
    s = str(spec or "").strip()
    out = {"dia": None, "len": None, "len_hi": None, "rx_cms": set(), "rx_cm": None}
    if not s:
        return out

    # ① QL / CL 写法（新的导管规格）—— 命中就整段按它算，不再做通用数字扫描，
    #    否则 15/80/150/200 会互相抢位置。
    mq = _QL_RANGE.search(s) or _QL_ONE.search(s)
    if mq:
        if _QL_RANGE.search(s):
            m = _QL_RANGE.search(s)
            out["len"], out["len_hi"] = float(m.group(1)), float(m.group(2))
        else:
            out["len"] = out["len_hi"] = float(mq.group(1))
        mcl = _CL_LIST.search(s)
        if mcl:
            out["rx_cms"] = {float(x) for x in
                             re.split(r"\s*[,，/、]\s*", mcl.group(1)) if x}
        else:
            m1 = _CL_ONE.search(s)
            if m1:
                out["rx_cms"] = {float(m1.group(1))}
        if out["rx_cms"]:
            out["rx_cm"] = min(out["rx_cms"])
        return out

    # ② 老写法：整串按出现顺序取「数字 + 可选单位」，带 cm 的是有效长度。
    #    `1.25 X 15 90cm` 的第三段是 `15 90cm`，按 X 切段会吞掉 90。
    nums, rx = [], None
    for m in _NUMU.finditer(s):
        v = float(m.group(1))
        if (m.group(2) or "").lower() == "cm":
            rx = v                      # 有效长度
        else:
            nums.append(v)
    out["dia"] = nums[0] if nums else None
    ln = nums[1] if len(nums) > 1 else None
    out["len"] = out["len_hi"] = ln
    if rx is not None:
        out["rx_cms"] = {rx}
        out["rx_cm"] = rx
    return out


def card_tokens(xlsx_path):
    """从「表2 · 大类」抽产品形态（RX / OTW）。返回 set，例如 {"RX"}。

    为什么用它：参数卡本身没有「形态」这一列，但表2 的大类写着
    「装入盘管组件NEO RX」—— 与附表标题里的 RX-90cm / OTW-90cm 同口径。
    读不到表2、或大类里同时出现两种形态 → 返回空集（调用方据此弃权）。
    """
    key = os.path.abspath(str(xlsx_path))
    if key in _TOKENS_CACHE:
        return _TOKENS_CACHE[key]
    toks = set()
    try:
        import openpyxl
        wb = openpyxl.load_workbook(xlsx_path, data_only=True, read_only=True)
        try:
            if "表2" in wb.sheetnames:
                ws = wb["表2"]
                rows = ws.iter_rows(max_row=60, values_only=True)
                head = None
                for r in rows:
                    vals = [str(c).strip() if c is not None else "" for c in r]
                    if head is None:
                        if "大类" in vals:
                            head = vals.index("大类")
                        continue
                    if head is not None and head < len(vals):
                        t = vals[head].upper()
                        for fm in ("OTW", "RX"):
                            if re.search(r"\b%s\b" % fm, t):
                                toks.add(fm)
                    if len(toks) > 1:
                        break
        finally:
            try:
                wb.close()
            except Exception:
                pass
    except Exception:
        toks = set()
    _TOKENS_CACHE[key] = toks
    return toks


# ------------------------------------------------------------------ 3. 自检
def report(pdf_path):
    """给人看的自检：这份 PDF 的维度解析结果。"""
    dims = page_dims(pdf_path)
    L = ["【附表页维度】共 %d 页带维度" % len(dims)]
    for pno in sorted(dims):
        d = dims[pno]
        L.append("  第%-3d页 附表%-3s 长度档=%-12s 形态=%s"
                 % (pno, d["sub"], d["len"] or "—",
                    "、".join("%s-%s" % (f, "/".join("%g" % x for x in sorted(c)))
                             for f, c in d["forms"]) or "—"))
    return L
