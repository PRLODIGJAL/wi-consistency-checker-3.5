# -*- coding: utf-8 -*-
"""xl_locate.py —— 把「工具说的行号」翻译成「用户手里那张表的行号」

为什么要这个模块：
  同一个工作簿里往往有好几张「像参数表」的表，粒度完全不同。以
  `DWG-0000014,rev05,…xlsx` 为例：

    表4      280 行  参数卡模板：一张卡几行采集项定义，没有设备/产品展开
    rev06  14492 行  展开后的实际数据：同一采集项按 设备 × 产品 各占一行

  引擎必须读 rev06 —— 只有它带设备号/适用规格，收窄链才有依据；换成表4
  会把 14490 条核对缩成 279 条，覆盖率直接崩掉。
  但使用者核对时打开的是表4，工具报的却是 rev06 的行号，于是「第 58 行」
  在表4 里翻出来是另一条东西 —— 用户原话：「行号与实际 excel 中表4 的
  内容好像对应不上」。

  对不上的根源是**没说是哪张表的第几行**，不是行号算错了。所以这里不去改
  数据源，只做一件事：按「数据采集项名称」在表4 里反查行号，让界面能同时
  给出两个行号（rev06!58 · 表4 第 2 行）。

三条边界：
  · 只读；不参与任何判定；
  · 表4 不存在 / 读不动 / 名字对不上 → 一律返回空，界面显示「—」；
  · 结果按 (路径, mtime, 表名) 缓存，一次运行只读一遍。
"""
import os

_CACHE = {}          # {(path, mtime, sheet): {item: [row, ...]}}


def _norm(s):
    """名字比对用的归一化：去掉所有空白与全角空格。

    「焊接温度（℃）」在表4 与 rev06 里字形一致，但中间可能夹 `\\xa0`；
    只去空白、不去括号和标点 —— 名字就是名字，不做同义归并。
    """
    return "".join(str(s or "").split()).replace("\u3000", "")


def find_sheet(path, want="表4"):
    """在工作簿里找 `表4` 这张表（允许有前后缀空格）。找不到返回 None。"""
    try:
        mt = os.path.getmtime(path)
    except OSError:
        return None
    ck = ("find", path, mt, want)
    if ck in _CACHE:
        return _CACHE[ck]
    got = None
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        _CACHE[ck] = None
        return None
    try:
        for n in wb.sheetnames:
            if str(n).strip() == want:
                got = n
                break
    finally:
        try:
            wb.close()
        except Exception:
            pass
    _CACHE[ck] = got
    return got


def _header_cols(ws, max_scan=6):
    """找表头行与「数据采集项名称」所在的列（1-based）。找不到返回 (None, None)。

    表4 的表头在第 1 行；但同族文件的模板表不一定 —— 表头行不固定时靠
    「哪一行同时出现『数据采集项名称』」来认，不写死行号。
    """
    for ri, row in enumerate(ws.iter_rows(max_row=max_scan, values_only=True), 1):
        for ci, v in enumerate(row or (), 1):
            if v is not None and _norm(v) == "数据采集项名称":
                return ri, ci
    return None, None


def index_of(path, sheet="表4"):
    """{归一化采集项名: [该名在 sheet 里的物理行号, ...]}；失败返回 {}。"""
    try:
        mt = os.path.getmtime(path)
    except OSError:
        return {}
    ck = (path, mt, sheet)
    if ck in _CACHE:
        return _CACHE[ck]
    out = {}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception:
        _CACHE[ck] = out
        return out
    try:
        if sheet not in wb.sheetnames:
            _CACHE[ck] = out
            return out
        ws = wb[sheet]
        hr, ic = _header_cols(ws)
        if hr is None:
            _CACHE[ck] = out
            return out
        for ri, row in enumerate(ws.iter_rows(min_row=hr + 1, values_only=True), hr + 1):
            v = row[ic - 1] if ic - 1 < len(row) else None
            if v is None or not str(v).strip():
                continue
            out.setdefault(_norm(v), []).append(ri)
    except Exception:
        out = {}
    finally:
        try:
            wb.close()
        except Exception:
            pass
    # 只留前 20 个行号：同一条采集项在多张卡里重复时，二十个足够定位，
    # 存全量会让缓存无限长（表4 才 280 行，但同族文件的表可能上万行）。
    for k in list(out):
        if len(out[k]) > 20:
            out[k] = out[k][:20]
    _CACHE[ck] = out
    return out


def rows_in(path, item, sheet="表4"):
    """这条采集项名在 sheet 里占哪几行（没有则空列表）。"""
    return index_of(path, sheet).get(_norm(item), [])


def describe(path, item, sheet="表4"):
    """给界面用的一句话：「表4 第 2 行」/「表4 第 2、5 行」/「表4 里没有」。"""
    if not path or not item:
        return "—"
    idx = index_of(path, sheet)
    rs = idx.get(_norm(item), [])
    if not rs:
        # 分两种说法：表在、只是没这个名字（说明这条采集项只存在于展开表里，
        # 表4 里确实没有对应行）；表不在（这份文件没有「表4」）。
        return ("%s 里没有" % sheet) if idx else "—"
    if len(rs) <= 3:
        return "%s 第 %s 行" % (sheet, "、".join(str(x) for x in rs))
    return "%s 第 %s 行…共 %d 行" % (sheet, "、".join(str(x) for x in rs[:3]), len(rs))
