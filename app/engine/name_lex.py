# -*- coding: utf-8 -*-
"""name_lex.py —— 参数名词库（学习闭环的存储与扫描）

设计口径（与 wi_core 的分工）：
  · 词库只回答「Excel 里这个名字，在 PDF 里写成什么」—— 即**名字对齐**；
  · 数值比较、收窄、判定归 wi_core，词库一行都不碰；
  · 所以词库为空时，工具行为与不存在这个模块时**逐条一致**。

为什么是「Excel 名 → PDF 原名」而不是「→ 标准键」：
  现有 `_PARAM_ALIAS` 已经验证过这条路（`停留时间 → 浸没时间` 之后去查
  `vindex["名:浸没时间"]`）。PDF 原名在字面桶里一定有记录（竖表通道对
  词典不认的名字也会挂字面桶），所以目标名只要在 PDF 里出现过就能命中；
  若改成「→ 标准键」，反而要承担「两个名字并进同一个键」的假一致风险。

存储：与程序同目录的 `词库.json`（便携，不依赖 SQLite）。
      每次保存自动备份成 `词库.json.bak-时间戳`，保留最近 10 份 —— 这就是
      「任何词库变更之后必须能一键回滚到上一版本」的实现方式。

三类记录：
  aliases      {excel_lit: {"pdf": PDF原名, "unit":, "scope":, "src":, ...}}
  non_params   {name_lit: {"kind": 侧标/工序名/说明/确认项, "scope":, ...}}
  tiers        {wi: [{"sub": 附表号, "lo":, "hi":, "side":}]}   # 预留给长度档映射
"""
import os
import json
import re
import difflib
import datetime
import sys as _sys

# ── 目录四层自举（2026-11-03 目录改造）：让下面的 `from paths import …` 有得找 ──
_APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _APP_DIR not in _sys.path:
    _sys.path.insert(0, _APP_DIR)
try:
    import paths as _paths
    _paths.install()
except Exception:       # pragma: no cover —— 单独拷文件调试时允许没有 paths
    pass

SCHEMA = 1
FILE_NAME = "词库.json"
_HERE = os.path.dirname(os.path.abspath(__file__))
# 2026-10-30：词库挪到 数据\（程序目录只放代码，用户改出来的东西不进交付包）。
# 路径由 paths.py 统一回答；万一 paths 不可用（单独拷这个文件调试）就退回同目录。
try:
    from paths import data_path as _data_path
    _PATH = _data_path(FILE_NAME)
except Exception:       # pragma: no cover
    _PATH = os.path.join(_HERE, FILE_NAME)

# 适用范围：{"type": "global"} / {"type": "wi", "value": "WI-0042"} /
#           {"type": "family", "value": "NEO"} / {"type": "oem"}
_CACHE = {"lex": None, "path": _PATH, "mtime": 0}


# ---------------------------------------------------------------- 基础读写
def _blank():
    return {"version": SCHEMA, "updated": "", "aliases": {},
            "non_params": {}, "tiers": {}}


def path_of():
    return _CACHE["path"]


def set_path(p):
    """给测试/多份词库用。切换后缓存失效。"""
    _CACHE["path"] = p
    _CACHE["lex"] = None
    _CACHE["mtime"] = 0


def load(force=False):
    """读词库；文件不存在当空库（**不是错误**，首次运行就是这样）。"""
    p = _CACHE["path"]
    try:
        mt = os.path.getmtime(p)
    except OSError:
        _CACHE["lex"], _CACHE["mtime"] = _blank(), 0
        return _CACHE["lex"]
    if force or _CACHE["lex"] is None or mt != _CACHE["mtime"]:
        try:
            with open(p, encoding="utf-8") as f:
                lex = json.load(f)
        except Exception:
            # 词库坏了不能让工具起不来：退回空库，并把坏文件留个证据
            lex = _blank()
            try:
                os.replace(p, p + ".corrupt")
            except OSError:
                pass
        for k in ("aliases", "non_params", "tiers"):
            lex.setdefault(k, {})
        _CACHE["lex"], _CACHE["mtime"] = lex, mt
    return _CACHE["lex"]


def _backup(p):
    if not os.path.exists(p):
        return
    ts = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    try:
        with open(p, "rb") as src:
            blob = src.read()
        with open("%s.bak-%s" % (p, ts), "wb") as dst:
            dst.write(blob)
    except OSError:
        return
    # 只留最近 10 份
    d = os.path.dirname(p)
    base = os.path.basename(p) + ".bak-"
    olds = sorted(x for x in os.listdir(d) if x.startswith(base))
    for x in olds[:-10]:
        try:
            os.remove(os.path.join(d, x))
        except OSError:
            pass


def save(lex=None):
    """保存（先备份）。返回备份文件路径（无旧文件则 None）。"""
    lex = lex or load()
    lex["version"] = SCHEMA
    lex["updated"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    p = _CACHE["path"]
    _backup(p)
    tmp = p + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(lex, f, ensure_ascii=False, indent=1, sort_keys=True)
    os.replace(tmp, p)
    _CACHE["lex"], _CACHE["mtime"] = lex, os.path.getmtime(p)
    push_to_core()
    return p


# ---------------------------------------------------------------- 适用范围
def _scope_ok(scope, wi=None, family=None):
    """一条记录在当前这份文件上是否适用。默认（缺 scope）= 全局。"""
    if not scope:
        return True
    t = scope.get("type") or "global"
    if t == "global":
        return True
    v = str(scope.get("value") or "").strip().upper()
    if not v:
        return True
    if t == "wi":
        return str(wi or "").strip().upper() == v
    if t == "family":
        return v in str(family or "").upper()
    if t == "oem":
        return "OEM" in str(family or "").upper()
    return True


def alias_map(wi=None, family=None):
    """当前适用范围内的 {Excel 字面键: PDF 原名}。"""
    lex = load()
    out = {}
    for k, rec in (lex.get("aliases") or {}).items():
        if not rec.get("pdf"):
            continue
        if _scope_ok(rec.get("scope"), wi, family):
            out[k] = rec["pdf"]
    return out


def non_param_set(wi=None, family=None):
    lex = load()
    return {k for k, rec in (lex.get("non_params") or {}).items()
            if _scope_ok(rec.get("scope"), wi, family)}


def push_to_core(wi=None, family=None):
    """把词库推给 wi_core（它自己不会去读文件，避免循环依赖）。"""
    try:
        import wi_core
        wi_core.set_lexicon(alias_map(wi, family), non_param_set(wi, family))
    except Exception:
        pass


# ---------------------------------------------------------------- 写入
def add_alias(excel_name, pdf_name, unit="", scope=None, src=None, note=""):
    import wi_core
    k = wi_core._literal_key(excel_name)
    if not k or not pdf_name:
        return False
    lex = load()
    lex["aliases"][k] = {
        "excel": str(excel_name).strip(), "pdf": str(pdf_name).strip(),
        "unit": unit or "", "scope": scope or {"type": "global"},
        "src": src or {}, "note": note, "enabled": True,
        "ts": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    save(lex)
    return True


def add_non_param(name, kind="侧标", scope=None, src=None, note=""):
    import wi_core
    k = wi_core._literal_key(name)
    if not k:
        return False
    lex = load()
    lex["non_params"][k] = {
        "name": str(name).strip(), "kind": kind or "侧标",
        "scope": scope or {"type": "global"}, "src": src or {},
        "note": note, "enabled": True,
        "ts": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    save(lex)
    return True


def disable(kind, key):
    lex = load()
    rec = (lex.get(kind) or {}).get(key)
    if rec is None:
        return False
    rec["enabled"] = False
    save(lex)
    return True


def enable(kind, key):
    lex = load()
    rec = (lex.get(kind) or {}).get(key)
    if rec is None:
        return False
    rec["enabled"] = True
    save(lex)
    return True


def remove(kind, key):
    """删掉一条（2026-11-02 学习记录窗用）。

    与 `disable` 的区别：停用只是不生效、记录还在（随时能翻出来看当时为什么这么判）；
    删除是连记录一起没了。界面把两件事分开给入口，别让「停用」变成「删了没法查」。
    """
    lex = load()
    if (lex.get(kind) or {}).pop(key, None) is None:
        return False
    save(lex)
    return True


def clear_kind(kind):
    """清空某一类（aliases / non_params）。返回清掉了多少条。"""
    lex = load()
    n = len(lex.get(kind) or {})
    lex[kind] = {}
    save(lex)
    return n


def records(kind):
    """某一类的全部记录，最近改动在前。每条多带一个 `key`。"""
    lex = load()
    out = []
    for k, rec in (lex.get(kind) or {}).items():
        d = dict(rec)
        d["key"] = k
        d["_kind"] = kind
        out.append(d)
    out.sort(key=lambda r: r.get("ts") or "", reverse=True)
    return out


# ---------------------------------------------------------------- 候选排序
_WORD_RE = re.compile(r"[a-z][a-z0-9]*")


def _sim(a, b):
    a, b = (a or "").lower(), (b or "").lower()
    if not a or not b:
        return 0.0
    r = difflib.SequenceMatcher(None, a, b).ratio()
    # 一个串完全包含另一个：这类「加了限定词」的情形相似度会被长度惩罚，
    # 单独提一档（例：「底涂 停留时间」vs「浸没时间」不会命中，但
    # 「校正产品位置 速度」vs「校正产品位置」会）。
    if a in b or b in a:
        r = max(r, 0.75)
    return r


def _word_sim(a, b):
    """中英混写的名字：只比两边的**英文词**，返回 0 表示没有共同词。

    为什么必须在**原始字号**上算、不能用 `_literal_key` 的结果：
    `_literal_key` 会把空格吃掉，「Fold Head Close Stop Micrometer Setting」
    粘成 `foldheadclosestopmicrometersetting` 一整串，一个词都切不出来。

    Excel 常把 PDF 里的英文词原样抄过来（「Fold 头关闭时针规设定值 不过」），
    两边整串比只有 0.16 —— 中文那半截和英文完全不重合，SequenceMatcher
    只会把它埋进 107 个候选的第 5 位。而「fold」这个词是重合的，它就是最硬的线索：
    命中一个英文词提到 0.55，命中越多越高（上限 0.9，仍低于「字面包含」的 0.75 之上
    但不越俎代庖）。仍然只是排序，不产生任何判定。
    """
    wa = set(_WORD_RE.findall(str(a or "").lower()))
    wb = set(_WORD_RE.findall(str(b or "").lower()))
    if not wa or not wb:
        return 0.0
    common = wa & wb
    if not common:
        return 0.0
    return min(0.9, 0.45 + 0.15 * len(common))


def candidates(name, pool, unit="", pool_units=None, top=5, min_score=0.34):
    """给一个 Excel 名找 PDF 侧候选。**只用于排序给人看，不产生结论。**

    pool_units: {PDF 名: 单位集合}；单位相同者加一个 ★ 提示（不加分排序之外的
    任何判断）—— 因为「单位相同」是旁证，不是证明。

    min_score（2026-10-30）：默认 0.34 是「列出来也是噪声」的下限，用于
    Combobox 那种只放得下几项的场合。界面改成「全部列出来 + 搜索框」之后
    要传 0.0 —— 像「真空跑漏」和 PDF 里任何名字都算不出相似度，
    卡阈值的结果就是候选框**空的**，等于没得选（用户反馈原话：
    「PDF 内的预设名字至少得提取出来给我选择吧」）。
    """
    import wi_core
    k = wi_core._literal_key(name)
    # 单位只用于「★ 单位相同」这个提示，**不参与任何判定**，所以可以在这里
    # 单独归一：Excel 侧常写「（Inch）」「（mW）」这种带括号的，拿着括号去和
    # PDF 的 'inch' / 'mw' 比永远不等，★ 就永远不亮 —— 而中英文名字相似度恒为 0
    # 的时候，★ 是唯一还能起作用的旁证。
    unit = str(wi_core._unwrap_unit(unit) or "").strip().lower()
    out = []
    for p in pool:
        pk = wi_core._literal_key(p)
        s = _sim(k, pk)
        # 英文词重合度（中英混写的名字常有「Fold / Pleat / Hub」这类原样抄来的词）
        sw = _word_sim(name, p)
        if sw > s:
            s = sw
        # min_score=0 表示「一个都不滤」。用 <= 会把相似度正好 0 的（毫不相像的
        # 那些）一起删掉 —— 右栏就只剩一两个，等于没给选。
        if min_score > 0 and s <= min_score:
            continue
        same_unit = False
        if unit and pool_units:
            us = {str(u).strip().lower() for u in (pool_units.get(p) or ())}
            same_unit = bool(us) and (unit in us)
        out.append({"pdf": p, "score": round(s, 3), "same_unit": same_unit})
    # 排序：相似度 → 「单位相同」→ 名字。
    # 相似度打平时单位优先，是因为中英文对照的名字**相似度恒为 0**：
    # 「Fold 头关闭时针规设定值 不过」对上 `Fold Head Close Stop Micrometer
    # Setting`，SequenceMatcher 给出 0.0，于是它和 100 个无关名字并列 ——
    # 只按字母序排的话它沉在第 40 多位，等于没给人看。单位相同是这里唯一
    # 还能用的旁证，所以让它把并列的一堆顶上来（仍然只是排序，不是判定）。
    out.sort(key=lambda x: (-x["score"], not x["same_unit"], x["pdf"]))
    return out[:top]


# ---------------------------------------------------------------- 扫描
def scan_unknowns(checks, wi=None, family=None, min_rows=1):
    """从 compare() 的结果里，按**名字**清点「词典不认识、词库也没登记」的项。

    ⚠️ 为什么不按 `scope == "name_unknown"` 扫（阶段 0 实测该档 = 0 条）：
       名字不认的行绝大多数在进入判定出口之前就被分成了 `collection`
       （类型列是布尔值/字符串、无值）或 `novalue`（值不是数字）。
       只扫那一档，弹窗永远不会出来。

    返回 [{name, rows, unit, statuses, kinds}]，按行数降序。
    """
    import wi_core
    lex_alias = alias_map(wi, family)
    nonp = non_param_set(wi, family)
    agg = {}
    for c in checks or []:
        item = (c.get("item") or "").strip()
        if not item:
            continue
        if wi_core._meta_kind_of(item):        # 机器编号/程序号等，走元信息路径
            continue
        k = wi_core._literal_key(item)
        if not k or k in nonp or k in lex_alias:
            continue
        if wi_core._canonical_param(item) is not None:
            continue
        rec = agg.setdefault(k, {"name": item, "rows": 0, "unit": "",
                                 "statuses": set(), "kinds": set(),
                                 "rowlist": [], "sheet": ""})
        rec["rows"] += 1
        # 具体行号（2026-10-30）：只给「9 行」这种数字没法回 Excel 定位 ——
        # 用户原话「行号与实际 excel 中表4 的内容好像对应不上」，根子就在这里：
        # 他要的是「去第几行改」，不是「一共几行」。前 8 个够点了。
        _r = c.get("row")
        if _r is not None and len(rec["rowlist"]) < 8:
            rec["rowlist"].append(_r)
        if not rec["sheet"]:
            rec["sheet"] = (c.get("sheet") or "").strip()
        if not rec["unit"] and (c.get("unit") or "").strip():
            rec["unit"] = (c.get("unit") or "").strip()
        rec["statuses"].add(c.get("status") or "")
        t = (c.get("type") or "").strip()
        if t:
            rec["kinds"].add(t)
    out = [v for v in agg.values() if v["rows"] >= min_rows]
    out.sort(key=lambda x: (-x["rows"], x["name"]))
    return out


def pdf_name_pool(pdf_vertical=None, pdf_steps=None, with_units=True):
    """PDF 侧出现过的全部参数名（去重）。候选就从这个池子里挑。

    返回 (names, units_by_name)
    """
    names, units = [], {}
    for r in list(pdf_vertical or []) + list(pdf_steps or []):
        p = (r.get("param") or "").strip()
        if not p:
            continue
        if p not in units:
            names.append(p)
            units[p] = set()
        for u in (r.get("units") or ()):
            if u:
                units[p].add(str(u).strip().lower())
    return names, units


# ---------------------------------------------------------------- 统计
def stats():
    lex = load()
    a = lex.get("aliases") or {}
    n = lex.get("non_params") or {}
    return {"aliases": sum(1 for r in a.values() if r.get("enabled", True)),
            "aliases_off": sum(1 for r in a.values() if not r.get("enabled", True)),
            "non_params": sum(1 for r in n.values() if r.get("enabled", True)),
            "path": _CACHE["path"]}
