# -*- coding: utf-8 -*-
"""anchors.py —— 锚点库：把「Excel 这一行」钉在「PDF 这张表这一格」上

为什么要有这个东西（2026-11-02 用户定方向的原话）：

  「不合理的（错）问题会源源不断这样子的，不应该（一处处）去改了啊……
    我们最初的目标就是为了各种各样的模板给它统计起来，然后去从底层改。」

  「你比如说 30~50 这个，然后前面有一堆什么什么长度什么的，那我还要去 PDF
    里面找这个关键字眼，还要匹配上，而且你不一定截取得到这个关键字眼。
    但是反过来，我们最终要的是数据。那我直接选取 PDF 中的数据，让他自己去
    反推是哪个关键字眼，这样子行不行？」

要解决的问题
------------
现在整条链子的命门是**名字字符串**：Excel 采集项名 → 去 PDF 的文字层里找
一个「长得像」的名字 → 再用那个名字取值。字符串是这份文件里最不可靠的东西：
换个模板就换个写法、被换行截断、被塞进合并格、中英混排时相似度算出来是 0。
所以每遇到一种新模板就掉一批「找不到对应」，修一条来一条。

锚点换掉的就是这一环：**不看名字，看位置。**

  一条锚 = 「Excel 里这一行」→「第几页、第几张表、第几行第几列」

这个位置由**人在界面上点出来**（他看得见那张表，一眼就能点中 30~40 那一格），
工具只负责照着这个坐标把格子里的字读出来、解析成数值、跟 Excel 比。
名字长得像不像，从此与这一行无关。

为什么「位置」比「名字」耐用
--------------------------
一条锚记三样东西，缺一不可：
  1. **模板指纹**（`fp`）：表号 + 表名 + 前几个字段名 —— 见 `pdf_tables.fingerprint()`
  2. **行锚**：那一行的第一列（行标签，如 P/N `SABERX.014`）
  3. **列锚**：那一列的字段名（如「拉伸前切定长」）

下次换一份文件，只要还是**同一张表**（指纹对得上），就先用行标签/列名重新定位
一次，定位不到才退回原来的行列序号。这样即使新版本多插了一行、调了列序，
锚点照样有效 —— 这才是「把各种模板统计起来」的意思：统计的是**表的结构**，
不是某一次的名字写法。

三条边界（与 name_lex.py / xl_locate.py / pdf_tables.py 完全一致）
------------------------------------------------------------------
  · 只读源文件：锚点只记坐标，从不改 PDF / Excel；
  · 只影响**配了锚点的那一行**：没配的行一格都不碰，锚点库为空时行为与没有
    这个模块时逐条一致（`_apply_anchors` 第一行就 return）；
  · 不猜：锚点指向的表在这份 PDF 里认不出来 → 这一条不生效，如实说明
    「你标的那张表这份文件里没有」，绝不拿一张形状相近的表顶上。
"""
import datetime
import json
import os
import re
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
FILE_NAME = "锚点.json"

try:
    from paths import data_path as _data_path
    _PATH = _data_path(FILE_NAME)
except Exception:       # pragma: no cover
    _PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), FILE_NAME)

_CACHE = {"lex": None, "path": _PATH, "mtime": 0}


# ---------------------------------------------------------------- 基础读写
def _blank():
    return {"version": SCHEMA, "updated": "", "anchors": {}, "rules": {}}


def path_of():
    return _CACHE["path"]


def set_path(p):
    _CACHE["path"] = p
    _CACHE["lex"] = None
    _CACHE["mtime"] = 0


def load(force=False):
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
            lex = _blank()
            try:
                os.replace(p, p + ".corrupt")
            except OSError:
                pass
        lex.setdefault("anchors", {})
        lex.setdefault("rules", {})
        _CACHE["lex"], _CACHE["mtime"] = lex, mt
    return _CACHE["lex"]


def _backup(p):
    if not os.path.exists(p):
        return None
    ts = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
    dst = "%s.bak-%s" % (p, ts)
    try:
        with open(p, "rb") as src, open(dst, "wb") as out:
            out.write(src.read())
    except OSError:
        return None
    d = os.path.dirname(p)
    base = os.path.basename(p) + ".bak-"
    olds = sorted(x for x in os.listdir(d) if x.startswith(base))
    for x in olds[:-10]:
        try:
            os.remove(os.path.join(d, x))
        except OSError:
            pass
    return dst


def save(lex=None):
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
    return p


# ---------------------------------------------------------------- 适用范围
def _scope_ok(scope, wi=None, family=None):
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
    return True


def _key(name, excel_row=None):
    """锚点的存储键：`字面名#行号`。

    ⚠️ 为什么必须带行号（2026-11-02 端到端实测踩出来的）：
       WI-0039 的参数卡表4 里有 **8 行同名**采集项（都是「使用模板将球囊保护套
       切定长，长度」），只是各行填的值不同、各自对应 PDF 表7 里的不同一行。
       只按名字做键的话，钉一次会把这 8 行**一起**指到同一格 —— 实测直接产生
       8 条假差异。一个人指的是「这一行」，不是一个名字。

    行号用 Excel 数据源里的**数据行号**（与 check["row"] 同源）：它在同一份参数卡
    里是唯一的，`xl_locate` 那套「表4 第几行」也是从它翻译出来的。
    """
    try:
        import wi_core
        lit = wi_core._literal_key(name)
    except Exception:
        lit = re.sub(r"[\s\u3000\u00a0]+", "", str(name or "")).lower()
    if not lit:
        return ""
    if excel_row in (None, ""):
        return lit + "#*"      # 不带行号 = 这份表里同名行都适用（老格式，兼容用）
    try:
        return "%s#%d" % (lit, int(excel_row))
    except (TypeError, ValueError):
        return lit + "#*"


# ---------------------------------------------------------------- 写入
def _cell_rec(c):
    """把「一格」规范成存盘格式。字段名与主格那一套同名 —— 换文件重定位时
    走的是同一个 `pick_tab` + `_relocate`，不另开一条路。

    `fp` 按**这一格自己所在的那张表**算，不是按主格那张：多选的全部意义就是
    「同一个采集项分写在两张不同的表里」（WI-0028 的焊接时间在表2/表3、
    焊机温度在表4/表5），两张表指纹不同，用主格的指纹去认第二格必然认错。
    """
    tab = c.get("item") or {}
    fp = str(c.get("fp") or "")
    if tab:
        try:
            import pdf_tables
            fp = pdf_tables.fingerprint(tab) or fp
        except Exception:
            pass
    try:
        tno = c.get("tno")
        tno = tab.get("tno") if tno is None else tno
        return {"fp": fp,
                "name": str(tab.get("name") or c.get("name") or "").strip(),
                "label": str(tab.get("label") or c.get("label") or "").strip(),
                "page": int(tab.get("page") or c.get("page") or 0),
                "tno": int(tno or 0),
                "row": int(c.get("row") or 0),
                "col": int(c.get("col") or 0),
                "row_label": str(c.get("row_label") or "").strip(),
                "col_header": str(c.get("col_header") or "").strip(),
                "value": str(c.get("value") or "").strip()}
    except (TypeError, ValueError):
        return {}


def add(excel_name, item, row, col, pdf_name="", row_label="", col_header="",
        value="", scope=None, src=None, note="", excel_row=None, cells=None):
    """钉一条锚点。`item` 是 pdf_tables.catalog() 的一项。

    `excel_row` 是 Excel 数据源里的行号（就是 check 里那个 `row`）。
    **强烈建议带上** —— 不带就是「同名行全适用」，一张表里有两行同名采集项时
    会把它们一起指到同一格（见 `_key` 的注释）。

    `cells` 是**多选**时人点中的全部格子（含主格，顺序即点选顺序）。
    单项 `row/col/item` 仍然写，值等于主格那一格 —— 老版本的读取方只认这几个
    字段，多出来的 `cells` 它们看不见就用不上，正好退回「只钉一格」的老行为。
    为什么要有它：2026-11-11 用户原话「它这里对应了两个表，它没有分开写。也就是
    说这两个表的话，取一起的话都是 157 到 181」—— 同一个采集项在一份 PDF 里
    常常分写两张表，只钉一格等于逼人在两者中二选一。

    返回 (ok, message)。message 是给人看的一句话（成功也说，失败也说清楚为什么）。
    """
    k = _key(excel_name, excel_row)
    if not k:
        return False, "Excel 名字是空的"
    if not item:
        return False, "没选表"
    try:
        import pdf_tables
        fp = pdf_tables.fingerprint(item)
    except Exception:
        fp = ""
    if not fp:
        return False, ("这张表认不出模板指纹（既没有表号/表名，字段名也不够）—— "
                       "记下来的锚点没法在别的文件里复用，所以不存")
    lex = load()
    _rec = {
        "excel": str(excel_name).strip(),
        "excel_row": (int(excel_row) if str(excel_row or "").strip() not in ("", "None")
                      else None),
        "fp": fp,
        "name": str(item.get("name") or "").strip(),
        "label": str(item.get("label") or "").strip(),
        "page": int(item.get("page") or 0),
        "tno": int(item.get("tno") or 0),
        "row": int(row),
        "col": int(col),
        "row_label": str(row_label or "").strip(),
        "col_header": str(col_header or "").strip(),
        "pdf_name": str(pdf_name or "").strip(),
        "value": str(value or "").strip(),
        "scope": scope or {"type": "global"},
        "src": src or {},
        "note": note, "enabled": True,
        "ts": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
    _cs = [_cell_rec(c) for c in (cells or ())]
    _cs = [c for c in _cs if c and c.get("row") is not None]
    if len(_cs) > 1:
        _rec["cells"] = _cs
    else:
        _rec.pop("cells", None)      # 单格就是老格式，别往库里塞多余的字段
    lex["anchors"][k] = _rec
    save(lex)
    _n = len(_rec.get("cells") or ())
    if _n > 1:
        return True, "已记住：%s → %d 格（%s 等）" % (
            _rec["excel"], _n, _rec["cells"][-1].get("label") or "这张表")
    return True, "已记住：%s → %s 第%d行第%d列" % (
        _rec["excel"], item.get("label") or "这张表", row, col)


def disable(k, off=True):
    lex = load()
    rec = (lex.get("anchors") or {}).get(k)
    if rec is None:
        return False
    rec["enabled"] = not off
    save(lex)
    return True


def remove(k):
    lex = load()
    if (lex.get("anchors") or {}).pop(k, None) is None:
        return False
    save(lex)
    return True


def clear():
    """清空**全部**锚点（调用方负责确认；这里只顺手备份一份再清）。"""
    lex = load()
    n = len(lex.get("anchors") or {})
    lex["anchors"] = {}
    save(lex)
    return n


# ---------------------------------------------------------------- 查询
def records(wi=None, family=None, only_on=True):
    """当前适用范围内的锚点，按最近改动在前。"""
    lex = load()
    out = []
    for k, rec in (lex.get("anchors") or {}).items():
        if only_on and not rec.get("enabled", True):
            continue
        if not _scope_ok(rec.get("scope"), wi, family):
            continue
        d = dict(rec)
        d["key"] = k
        out.append(d)
    out.sort(key=lambda r: r.get("ts") or "", reverse=True)
    return out


def resolve(cat, wi=None, family=None):
    """把这批锚点对到**这份 PDF 的表目录**上。

    返回 {(字面名, Excel行号 或 None): 绑定}。查的时候先按 (名, 行) 精确查，
    查不到再按 (名, None) 查 —— 后者是「没带行号的老格式 / 同名行都适用」。

    绑定 = {page, tno, row, col, row_label, col_header, pdf_name, fp, stale}

    `stale=True` 表示：锚点记的表在这份 PDF 里认不出来（指纹对不上）。
    这种条目**不参与判定**，只会在界面上说一句「这份文件里没有你标的那张表」——
    拿一张形状相近的表顶上去会造出假一致，那比不套用更糟。
    """
    lex = load()
    out = {}
    _pend = 0        # 就地补全了几条（≥1 才写盘一次，不为了显示去反复改库）
    for k, rec in (lex.get("anchors") or {}).items():
        if not rec.get("enabled", True):
            continue
        if not _scope_ok(rec.get("scope"), wi, family):
            continue
        lit = k.split("#", 1)[0]
        _row = rec.get("excel_row")
        try:
            _row = int(_row) if _row not in (None, "") else None
        except (TypeError, ValueError):
            _row = None
        out[(lit, _row)] = {"page": rec.get("page"), "tno": rec.get("tno"),
                            "row": rec.get("row"), "col": rec.get("col"),
                            "row_label": rec.get("row_label") or "",
                            "col_header": rec.get("col_header") or "",
                            # 表名（钉的时候这张表印的名字）。指纹失效后，靠它续命。
                            "tname": rec.get("name") or "",
                            "pdf_name": rec.get("pdf_name") or "",
                            "fp": rec.get("fp") or "",
                            "value": rec.get("value") or "",
                            "excel": rec.get("excel") or "",
                            "excel_row": _row,
                            # 多选时人点中的全部格子（含主格）。老记录没有这个键，
                            # 下游一律当「只有主格一格」处理。
                            "cells": list(rec.get("cells") or ()),
                            "stale": True}
    if not out:
        return out
    try:
        import pdf_tables
    except Exception:
        return out
    for k, b in out.items():
        try:
            _cands = pdf_tables.match_tables(cat, b["fp"])
        except Exception:
            _cands = []
        # 指纹认不出（换版式最常见）→ 退回「表名 + 原文」再找一次。
        # 认回来就标 relinked，让依据里明说「这次是靠表名找到的」——
        # 不标的话人会以为指纹还灵，下次真改版了也不知道为什么突然失效。
        if not _cands:
            _cands = _semantic_cand(cat, b)
            if _cands:
                b["relinked"] = True
        try:
            tab = pick_tab(_cands, b)
        except Exception:
            tab = None
        # 多选的其余格子：**每格按自己的指纹各认各的表**。主格认不出（stale）
        # 不代表其它格也认不出 —— 所以这段放在 `if not tab: continue` 之前。
        _cs = []
        for cell in (b.get("cells") or ()):
            c2 = dict(cell)
            c2["stale"] = True
            try:
                _cc = pdf_tables.match_tables(cat, c2.get("fp"))
            except Exception:
                _cc = []
            if not _cc:
                _cc = _semantic_cand(cat, c2)
                if _cc:
                    c2["relinked"] = True
            try:
                tab2 = pick_tab(_cc, c2)
            except Exception:
                tab2 = None
            if tab2:
                c2["stale"] = False
                c2["page"], c2["tno"] = tab2.get("page"), tab2.get("tno")
                c2["label"] = str(tab2.get("label") or c2.get("label") or "").strip()
                c2["row"], c2["col"] = _relocate(tab2, c2)
            _cs.append(c2)
        b["cells"] = _cs
        if not tab:
            continue        # 保持 stale=True
        b["stale"] = False
        # 老记录里没有表名（11-02 之前钉的只存了表号），就地补一次并写回。
        # 只有**指纹逐字相同**才补 —— 那说明「这就是同一张表的同一个模板」，
        # 名字一定是它的；靠 page/tno 猜出来的不算，宁可不补也不写错。
        if not (b.get("tname") or "").strip():
            _nm = str(tab.get("name") or "").strip()
            if _nm:
                try:
                    if pdf_tables.fingerprint(tab) == (b.get("fp") or ""):
                        b["tname"] = _nm
                        _rec = (lex.get("anchors") or {}).get(k) or {}
                        if not (str(_rec.get("name") or "").strip()):
                            _rec["name"] = _nm
                            _pend += 1
                except Exception:
                    pass
        # 先用行标签 / 列名重新定位（新版本可能插行、调列序），定不到才用原序号
        b["page"], b["tno"] = tab.get("page"), tab.get("tno")
        # 表标签取**这份文件里**那张表的叫法（「第4页 表2」），不拿钉的时候那句 ——
        # 依据里写「第4页 第5张表」和 PDF 上印的「表2」对不上号，人会以为找错了表。
        b["label"] = str(tab.get("label") or "").strip()
        r, c = _relocate(tab, b)
        b["row"], b["col"] = r, c
        # 老格式就地补全（**只影响显示与下次换文件重定位，坐标和判定一格不动**）：
        # 第一版「指定位置」窗有两个毛病，已经写进 .json 的记录带着这两个毛病：
        #   ① col_header 存的是空串 —— 那时只取第 0/1 行那一格，两级表头取不到；
        #   ② pdf_name 取成了行标签（WI-0041 表3 全变成「1.25 mm」），
        #      因为它把「数据格本身」当成列名排除了，剩下的只有行标签。
        # 不写回文件，只在这一次读取时算好，免得为了修一句显示去改用户的库。
        try:
            if not (b.get("col_header") or "").strip():
                b["col_header"] = _col_short(_grid_of(tab), c, r)
        except Exception:
            pass
        _stored = (b.get("pdf_name") or "").strip()
        _rl = (b.get("row_label") or "").strip()
        # 老格式的 pdf_name 是「行号 / 那一列的列头」拼出来的
        # （`1 / 设备/工具说明：热焊机 参数值或范围`）—— 那是行号和表头，不是参数名。
        # 认这个形状就地重算，不写回用户的库（理由同上：改一句显示不动数据）。
        _oldfmt = bool(_rl) and _stored.startswith(_rl + " / ")
        if not _stored or _stored == _rl or _oldfmt:
            b["pdf_name"] = _infer_name(tab, r, c) or _stored
    if _pend:
        # 只补了「表名」这一个字段，坐标、判定口径一格没动；写盘失败也不影响判定。
        try:
            save(lex)
        except Exception:      # noqa: BLE001
            pass
    return out


def _grid_of(tab):
    """取一张表的原始格子（按 `set_pdf()` 记下的那份文件）。"""
    try:
        import pdf_tables
        return pdf_tables.grid(_LAST_PDF.get("path") or "",
                               tab.get("page"), tab.get("tno"))
    except Exception:
        return []


def _table_has_cell(tab, want):
    """这张表里有没有哪一格就写着 `want`（已归一化比较）。"""
    if not want:
        return False
    for row in _grid_of(tab) or ():
        for cell in row or ():
            if _norm(cell) == want:
                return True
    return False


_CJK_ONLY = re.compile(r"[^\u4e00-\u9fff]+")


def _cjk_str(s):
    """只留中文。卡号/图号里那些 `DWG-0000007` 不算「像不像」的证据。"""
    return _CJK_ONLY.sub("", str(s or ""))


def _name_score(a, b):
    """两段中文「像不像」：最长公共**连续**子串，先剔掉所有表名都有的通用词。

    参数表 / 过程 / 参数 这些词不做剔除的话，随便两张表都白送 3 个字，
    名字这一层就变得永远「有话说」，反而把设备层搅乱。
    """
    try:
        import wi_core
        return wi_core._name_score(a, b)
    except Exception:
        pass
    x, y = _cjk_str(a), _cjk_str(b)
    for w in ("参数表", "工艺参数", "过程参数", "参数", "过程"):
        x = x.replace(w, "")
        y = y.replace(w, "")
    if not x or not y:
        return 0
    if len(x) < len(y):
        x, y = y, x
    prev = [0] * (len(y) + 1)
    best = 0
    for ca in x:
        cur = [0] * (len(y) + 1)
        for j, cb in enumerate(y, 1):
            if ca == cb:
                cur[j] = prev[j - 1] + 1
                if cur[j] > best:
                    best = cur[j]
        prev = cur
    return best


def _semantic_cand(cat, b):
    """指纹认不出时，按「表名 + 那一格的原文」在整份文件里再找一次。

    为什么要有（2026-11-13）：指纹 = 表号|表名|前 4 个字段名。模板一改版，
    字段名加个后缀、表号从「表2」变「表2-a」，指纹整条对不上，锚点全部判
    stale —— 用户钉过的东西等于白钉，每换一版都要重新钉一遍。
    可人记的从来不是指纹，是「标记环焊接参数表里的 Set Time」。所以这里退回
    两条**语义**线索：表名（长得像）+ 那一格原文 / 行标签（还在不在）。

    只在**恰好剩一张**时才认；多张一律返回 []，宁可判 stale 也不猜 ——
    套错表会把别人的值安到这一行上，制造假一致。
    """
    tn = str(b.get("tname") or "").strip()
    if not tn:
        return []
    tabs = [c for c in (cat or ()) if c.get("kind") != "header"]
    hit = [c for c in tabs if _name_score(tn, c.get("name")) >= 2]
    if not hit:
        return []
    if len(hit) > 1:
        # 名字一样时按三样打分（2026-11-13 实测：WI-0028 第 4 页有 3 张
        # **同名**的「标记环焊接参数表」，光靠名字分不开）：
        #   那一格原文还在 → +2（内容对上是最实的证据）
        #   行标签还在     → +1
        #   还是原来的页码+第几张表 → +2（改版常常只改字段名，位置没动）
        # 只认唯一最高分；打平就返回空 —— 宁可 stale 也不猜。
        v = _norm(b.get("value") or "")
        lab = _norm(b.get("row_label") or "")
        pg, tn0 = b.get("page"), b.get("tno")

        def _score(c):
            s = 0
            if v and _table_has_cell(c, v):
                s += 2
            if lab:
                for row in _grid_of(c) or ():
                    if _norm(row[0] if row else "") == lab:
                        s += 1
                        break
            if pg is not None and str(c.get("tno")) == str(tn0) and c.get("page") == pg:
                s += 2
            return s

        _sc = [(_score(c), c) for c in hit]
        _best = max(s for s, _ in _sc)
        if _best <= 0:
            return []
        hit = [c for s, c in _sc if s == _best]
    return hit if len(hit) == 1 else []


def _by_tname(cands, b):
    """同指纹、同页多张表时，用「表名」再收一道。

    换版式之后那一格写的内容会变，表名不会 —— 名字是身份，内容是状态。
    """
    tn = str(b.get("tname") or "").strip()
    if not tn or len(cands) < 2:
        return cands
    hit = [c for c in cands if _name_score(tn, c.get("name")) >= 2]
    if not hit or len(hit) >= len(cands):
        return cands
    return hit


def pick_tab(cands, b):
    """同指纹命中多张表时，选**到底钉的是哪一张**。

    为什么要这一步（2026-11-02 实案 WI-0030，用户看到的 0.021）：
      第 4 页「第 2 张表」（锻打，`1 Die 0.0210″`）与「第 4 张表」（焊接，
      `1 Temperature 175-181/181-185`）字段名完全一样，指纹撞车。老代码取
      目录里第一张 → 用户钉温度，读到隔壁的 Die 尺寸 0.021。

    选法（强度从强到弱，任何一档不唯一就往下走）：
      ① 只有一张 → 就是它；
      ② 用「钉的时候那一格的原文」去认：整批候选里只有**一张**含有这段文字才算数
         （同一页两张同构表不可能写同一个值；换版本、表挪了位置也照样认得出）；
      ③ 平移到别的文件时内容可能改了，退回锚点里记的「第几页第几张表」——
         用户钉的时候就在那张表里，同一份文件里这个坐标是确定的；
      ④ 都不唯一 → 返回 None，由调用方标 stale（**不猜**，宁可不套用）。
    """
    cands = list(cands or ())
    if not cands:
        return None
    if len(cands) == 1:
        return cands[0]
    # ①′ 表名：换版式后内容会改，名字不会。只在「恰好剩一张」时采信。
    cands = _by_tname(cands, b)
    if len(cands) == 1:
        return cands[0]
    want = _norm(b.get("value") or "")
    if want:
        hit = [c for c in cands if _table_has_cell(c, want)]
        if len(hit) == 1:
            return hit[0]
    pg, tn = b.get("page"), b.get("tno")
    same = [c for c in cands
            if c.get("page") == pg and str(c.get("tno")) == str(tn)]
    if len(same) == 1:
        return same[0]
    return None


def find(binds, name, excel_row=None):
    """按「名 + 行号」查一条绑定；查不到退回「名 + 全适用」。"""
    try:
        import wi_core
        lit = wi_core._literal_key(name)
    except Exception:
        lit = re.sub(r"[\s\u3000\u00a0]+", "", str(name or "")).lower()
    if not lit:
        return None
    if excel_row not in (None, ""):
        try:
            b = binds.get((lit, int(excel_row)))
            if b:
                return b
        except (TypeError, ValueError):
            pass
    return binds.get((lit, None))


def _norm(s):
    return re.sub(r"[\s\u3000\u00a0]+", "", str(s or "")).lower()


def _relocate(tab, b):
    """在**新的**表里找「同一行同一列」。

    顺序：① 行标签对得上 → 用它的行号 + 原来的列号
          ② 列名对得上 → 用它的列号（两级/三级表头按整串比）
          ③ 都对不上 → 原样返回锚点里记的行列序号（表对了，位置大概率没变）
    返回值是 0 起的 (row, col)，与 pdf_tables.grid() 的行列一致。

    ⚠️ 列名要按**整串列头**比（`_col_stacks`），不能只比某一格：
       两级表头里「上限」只印在合并区的第一列（WI-0041 表3 第 4 列），
       第 5 列那格是空串。只比单格的话，列一旦挪位就永远对不上。
    """
    r0, c0 = int(b.get("row") or 0), int(b.get("col") or 0)
    try:
        import pdf_tables
        g = pdf_tables.grid(_LAST_PDF.get("path") or "", b.get("page"), b.get("tno"))
    except Exception:
        g = []
    if not g:
        return r0, c0
    lab = _norm(b.get("row_label") or "")
    hdr = _norm(b.get("col_header") or "")
    if lab:
        for i, row in enumerate(g):
            v = _norm(row[0] if row else "")
            if v and v == lab:
                return i, c0
    if hdr:
        stacks = _col_stacks(g, list(range(0, min(max(r0, 1), len(g)))))
        for j, st in enumerate(stacks):
            if _norm(st) == hdr:
                return r0, j
        for j, st in enumerate(stacks):
            if hdr and (hdr in _norm(st) or _norm(st) in hdr):
                return r0, j
        # 老格式（只存了一格表头）也认一下
        for hr in range(min(2, len(g))):
            for j, v in enumerate(g[hr]):
                if _norm(v) and _norm(v) == hdr:
                    return r0, j
    return r0, c0


_LAST_PDF = {"path": ""}


def set_pdf(p):
    """记住当前这份 PDF 的路径 —— `_relocate` 要用它去读格子。

    单独一个 setter 而不是塞进 resolve 的参数里，是为了让 resolve 的签名
    与 name_lex.alias_map(wi, family) 保持一致（app 那边调用形状相同）。
    """
    _LAST_PDF["path"] = p or ""


def _infer_name(tab, r, c, path=None):
    """从坐标反推出「这一格叫什么」。

    这是用户那句「让他自己去反推是哪个关键字眼」的实现：
    行首（第 0 列）是这个参数的行标签，表头是它的列名，两个拼起来就是名字。
    拼不出东西就返回空串 —— 名字只用于显示和写依据，不参与匹配。

    `path` 不给就用 `set_pdf()` 记下的那一份。给路径是必要的：界面上的
    「指定位置」窗可能在**还没跑过比对**的时候打开，那时 set_pdf 一个值都没有，
    反推就会拿到空网格、名字变成空 —— 界面上看着像「工具没反推出来」。
    """
    path = path if path is not None else _LAST_PDF.get("path") or ""
    try:
        import pdf_tables
        g = pdf_tables.grid(path, tab.get("page"), tab.get("tno"))
    except Exception:
        g = []
    lab = ""
    if g and 0 <= r < len(g) and g[r]:
        lab = (g[r][0] or "").strip()
    col = ""
    if g and r > 0:
        # 用整串列头（含合并表头的向右继承），而不是只看第 0/1 行那一格 ——
        # WI-0041 表3 的「上限」印在第 4 列，第 5/6 列是空串，直接取格子会漏掉。
        col = _col_short(g, c, r)
    # 有「参数名称」这一列的（三列参数表：序号 | 参数名称 | 参数值或范围），
    # 真名字就在那一列。不先用它，拼出来是
    # 「1 / 设备/工具说明：热焊机 参数值或范围」—— 那是行号和表头，不是参数名
    # （WI-0030 实测，用户看到的「0.021」那条依据里就是这个）。
    if r < len(g):
        for j in range(len(g[r])):
            for h in range(0, min(3, len(g))):
                head = _norm((g[h] or [])[j] if j < len(g[h]) else "")
                if head in ("参数名称", "参数名", "参数", "项目名称",
                            "项目", "检测项目", "名称"):
                    nm = ((g[r] or [])[j] if j < len(g[r]) else "") or ""
                    if nm.strip() and re.search(r"[A-Za-z\u4e00-\u9fff]", nm):
                        return nm.strip()
    # 行标签是个纯编号（序号 1/2/3）就别当名字用：「1 / …」看着像工具认错了
    if lab and re.fullmatch(r"[0-9.]+", lab):
        lab = ""
    if lab and col and lab != col:
        return "%s / %s" % (lab, col)
    return col or lab or ""


def _col_short(grid, c, upto=None):
    """某一列的列名，**取最后两截**，别把三级表头全糊上去。

    `_col_stacks` 拼出来的是整串，WI-0041 表3 第 5 列会拼成
    「Maximum Balloon Diameter （螺旋测微仪上限） Middle End/ 中间」——
    写进判定依据里长达 50 字，人一眼扫不到重点（实测截图里换行换得很难看）。
    最具体的信息在最后两截：方向（上限/下限）+ 部位（近端/中间/远端）。
    英文那截重复且长，去掉。
    """
    n = max(1, int(upto or 0)) if upto else max(1, len(grid))
    parts = []
    for h in range(0, min(n, len(grid))):
        row = grid[h] or []
        v = (row[c] if c < len(row) else "") or ""
        if not v and c > 0:
            for k in range(c - 1, -1, -1):
                if k < len(row) and row[k]:
                    v = row[k]
                    break
        v = (v or "").strip()
        if v and v not in parts:
            parts.append(v)
    if not parts:
        return ""
    return " ".join(parts[-2:])


# ---------------------------------------------------------------- 表级规则
#  「我标了一条，凭什么剩下 41 条不跟着学会？」—— 用户 2026-11-02 的原话：
#     「如果说你引入了学习机制呢，那比方说有10条，它其实本质上都是这一个表的，
#       这个由于这个参数名为收录导致的，那么有可能我把第一条给让它识别出来了，
#       它是不是剩下的也能够识别出来」
#
#  一条锚点只钉「这一行」。但人指的这一刻，其实同时暴露了**这张表怎么读**：
#  WI-0041 第4页表3「螺旋测微仪参数设定」的长相是
#
#      行  = 规格（1.25 mm / 1.5 mm / …）        ← 由 Excel「适用规格」定
#      列  = 两级：上限/下限（最大/最小直径）× 部位（近端/中间/远端）
#
#  而 Excel 里的 6 个采集项正好是「球囊 + 部位 + 上限/下限」×16 个规格 = 96 格。
#  所以钉中「球囊中间上限 + φ1.25xQL15」这一格，等于告诉工具：
#      「这个表的行看规格、列看『部位 + 上下限』」—— 剩下的行都能按这条推。
#
#  ⚠️ 三条保险（任何一条不满足就不立规则，宁可让人一条条点）：
#     ① 用户钉的那个名字必须**同时**含「方向词」（上限/下限）和「部位词」
#        （近端/中间/远端）—— 说明这一列确实是这样组织的，不是瞎猜；
#     ② 规则必须能**复现出人钉的那一格**（用规则重新算行和列，结果必须一样）。
#        复现不出来 = 我对这张表的理解是错的，那就别用；
#     ③ 只对**同一张表**（指纹相同）生效，且只碰仍未判定的行。
_J = re.compile(r"[\u4e00-\u9fff]")
_UP = ("上限", "最大", "max")
_DN = ("下限", "最小", "min")
_SITE = ("近端", "中间", "中端", "远端", "近端部", "proximal", "middle", "distal")


def _cjk(s):
    return set(_J.findall(str(s or "")))


def _dir_of(name):
    """名字里的方向词 → ('up'/'down', 那个词)；没有就 (None, '')。"""
    s = str(name or "")
    low = s.lower()
    for t in _UP:
        if t in low:
            return "up", t
    for t in _DN:
        if t in low:
            return "down", t
    return None, ""


def _site_of(name):
    """名字里的部位词（近端/中间/远端…）；没有就空串。"""
    low = str(name or "").lower()
    for t in _SITE:
        if t in low:
            return t
    return ""


def _nums(s):
    """一串文字里的数字（含区间两端）。

    「φ1.25xQL(200-300)」→ [1.25, 200.0, 300.0]；
    「2.5x（200-300）mm」→ [2.5, 200.0, 300.0]。
    """
    out = []
    for m in re.finditer(r"\d+(?:\.\d+)?", str(s or "")):
        try:
            v = float(m.group(0))
        except ValueError:
            continue
        if v not in out:
            out.append(v)
    return out


def _col_stacks(grid, hdr_rows):
    """每一列的「表头竖着拼起来」的整串文字（两级/三级表头都拼上）。

    ⚠️ 必须补「合并表头向右继承」（2026-11-02 实测踩出来的）：
       WI-0041 第4页表3 的表头是两级合并的 ——

           [0] Balloon Diameter | Minimum … | · | · | Maximum … | · | ·
           [1] 球囊直径        |（下限）   | · | · |（上限）   | · | ·
           [2] ·  | Proximal | Middle | Distal | Proximal | Middle | Distal

       「Maximum Balloon Diameter /（螺旋测微仪上限）」只写在**合并区的第一列**
       （第 4 列），第 5、6 列（Middle / Distal）在 `grid()` 里就是空串。
       不补的话第 5 列的表头只剩「Middle End/中间」，拼不出「上限」两个字，
       `pick_col` 就会拿第 4 列（近端）和第 5 列（中间）打平 → 弃权 → 学不出规则。
       Excel 里 6 个采集项正是靠「方向 × 部位」定位的，少了方向就全废。

    补法：表头行里遇到空格子，就继承**左边最近的非空格子**（合并单元格就是这么显示的）。
    只对表头行做，数据行不碰；列数不变，所以列号仍然对得上 `grid()`。
    """
    n = max((len(r) for r in grid), default=0)
    out = []
    for c in range(n):
        parts = []
        for h in hdr_rows:
            if h >= len(grid) or c >= len(grid[h]):
                continue
            v = grid[h][c]
            if v:
                parts.append(v)
                continue
            # 空格子：往左找同行的第一个非空表头（合并区的延续）
            for k in range(c - 1, -1, -1):
                if k < len(grid[h]) and grid[h][k]:
                    parts.append(grid[h][k])
                    break
        out.append(" ".join(parts))
    return out


def pick_col(stacks, name):
    """按名字给这一列打分 → (列号 或 None, 分数说明)。

    打分（越具体越靠前）：
      · 名字里的汉字与这一列表头文字的**重合字数**
      · 方向词（上限/下限）两边都有 → +3
      · 部位词两边都有 → +3

    必须**唯一最高**才认；并列就当没找到（宁可让人点，不许猜错列）。
    """
    nc = _cjk(name)
    d, _dw = _dir_of(name)
    site = _site_of(name)
    if not nc:
        return None, ""
    scored = []
    for c, st in enumerate(stacks):
        sc = len(nc & _cjk(st))
        low = st.lower()
        if d == "up" and any(t in low for t in _UP):
            sc += 3
        if d == "down" and any(t in low for t in _DN):
            sc += 3
        if site and site in low:
            sc += 3
        scored.append((sc, c))
    if not scored:
        return None, ""
    scored.sort(reverse=True)
    best, bc = scored[0]
    second = scored[1][0] if len(scored) > 1 else -1
    if best <= 0 or best == second:
        return None, ""
    if d is None and site is None:
        return None, ""      # 连方向/部位都没有，说明不了什么，不立规则
    return bc, "%s 与「%s」对得上" % (name, stacks[bc][:40])


def pick_row(grid, spec, hdr_rows):
    """按 Excel 的「适用规格」找这一行 → 行号 或 None。

    规则：
      · 直径（第一个数）必须相等；
      · 再比「规格里的其余数字有多少落在这个行标签里」（QL 的档位 200-300 就靠它区分）；
      · 打平就取**数字最少**的那一行（「2.5 mm」比「2.5x（200-300）mm」更基本，
        QL(20-150) 这种档位就该落到它）。
    """
    sn = _nums(spec)
    if not sn:
        return None
    best = None
    for r in range(hdr_rows[-1] + 1 if hdr_rows else 0, len(grid)):
        row = grid[r] or []
        lab = row[0] if row else ""
        rn = _nums(lab)
        if not rn:
            continue
        if abs(rn[0] - sn[0]) > 1e-6:
            continue
        cover = len(set(sn) & set(rn))
        key = (cover, -len(rn))
        if best is None or key > best[0]:
            best = (key, r)
    return best[1] if best else None


def derive(item, grid, row, col, excel_name, spec):
    """从「人钉的那一格」反推「这张表怎么读」。推不出/复现不了 → 返回 None。

    这是整个学习机制的核心：一条锚点 → 一条表级规则。
    """
    if not grid or row is None or col is None:
        return None
    if row <= 0:
        return None                     # 钉在表头上，没有数据行可比，立不了规则
    hdr_rows = list(range(0, row))
    d, _ = _dir_of(excel_name)
    site = _site_of(excel_name)
    if not d or not site:
        # 名字里没有「上下限 + 部位」这层结构 —— 说明这张表的列不是这么组织的，
        # 那就老老实实只钉这一格，别硬套。
        return None
    stacks = _col_stacks(grid, hdr_rows)
    pc, why = pick_col(stacks, excel_name)
    pr = pick_row(grid, spec, hdr_rows)
    # 保险②：规则必须复现人钉的那一格，否则说明理解错了
    if pc != col or pr != row:
        return None
    try:
        import pdf_tables
        fp = pdf_tables.fingerprint(item)
    except Exception:
        return None
    if not fp:
        return None
    return {"fp": fp,
            "label": item.get("label") or "",
            "name": item.get("name") or "",
            "hdr_rows": hdr_rows,
            "row_dim": "spec",
            "col_dim": "dir+site",
            "witness": {"excel": str(excel_name).strip(), "spec": str(spec or ""),
                        "row": int(row), "col": int(col),
                        "col_text": (stacks[col] if col < len(stacks) else "")[:80]},
            "why": why,
            "enabled": True,
            "ts": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}


def save_rule(rule, src=None):
    if not rule or not rule.get("fp"):
        return False
    lex = load()
    r = dict(rule)
    r["src"] = src or {}
    lex.setdefault("rules", {})[r["fp"]] = r
    save(lex)
    return True


def rule_records(only_on=True):
    lex = load()
    out = []
    for fp, r in (lex.get("rules") or {}).items():
        if only_on and not r.get("enabled", True):
            continue
        d = dict(r)
        d["key"] = fp
        out.append(d)
    out.sort(key=lambda x: x.get("ts") or "", reverse=True)
    return out


def drop_rule(fp):
    lex = load()
    if (lex.get("rules") or {}).pop(fp, None) is None:
        return False
    save(lex)
    return True


def toggle_rule(fp, off=None):
    lex = load()
    r = (lex.get("rules") or {}).get(fp)
    if r is None:
        return False
    r["enabled"] = (not r.get("enabled", True)) if off is None else (not off)
    save(lex)
    return True


def _scope_rule_ok(r, wi, family):
    return _scope_ok(r.get("scope"), wi, family)


def resolve_rules(cat, checks, pdf_path, wi=None, family=None):
    """用学到的表级规则，把还没判定的行算出「该看哪一格」。

    返回 {(字面名, Excel行号): 绑定} —— 与 `resolve()` 同构，直接喂同一个
    `_apply_anchors`，不另开一条判定路径（这是本设计最要紧的一点：
    规则推出来的行与人工钉的行，走的是完全一样的取数/比对/判定代码）。

    `checks` 可以是**比对结果**（带 status），也可以直接是 **Excel 采集项**
    （extract_excel_params 的原样输出）。后者是给 `app._run` 用的 ——
    那时比对还没跑，拿不到 status，所以这里对「没有 status 的记录」一律
    当候选。这不影响正确性：`_apply_anchors` 自己只改 no_match / no_scope，
    已经判出来的行一格都不会被这条规则碰到。
    """
    lex = load()
    rules = [r for r in (lex.get("rules") or {}).values()
             if r.get("enabled", True) and _scope_rule_ok(r, wi, family)]
    if not rules or not checks:
        return {}
    try:
        import pdf_tables
        import wi_core
    except Exception:
        return {}
    out = {}
    cache = {}
    for rule in rules:
        # 同指纹可能命中好几张表（WI-0030 第4页第2张表和第4张表就是）。
        # 规则是「这张表怎么读」，所以逐张试：哪一张能按规则推出「列头带方向+部位、
        # 行标签带规格」的那些格子，就用哪一张。一张都推不出来 → 这条规则在这份
        # 文件里不适用，原样跳过（绝不拿形状相近的表硬套）。
        cands = pdf_tables.match_tables(cat, rule.get("fp"))
        if not cands:
            continue                    # 这份 PDF 里没有那张表，规则不适用
        for tab in cands:
            key = (tab["page"], tab["tno"])
            if key not in cache:
                g = pdf_tables.grid(pdf_path, tab["page"], tab["tno"])
                hr = [h for h in (rule.get("hdr_rows") or []) if h < len(g)]
                cache[key] = (g, hr, _col_stacks(g, hr) if g and hr else [])
            g, hr, stacks = cache[key]
            if not g or not hr:
                continue
            got = {}
            for chk in checks:
                st = chk.get("status")
                if st is not None and st not in ("no_match", "no_scope"):
                    continue
                name = (chk.get("item") or "").strip()
                if not name:
                    continue
                if not _dir_of(name) or not _site_of(name):
                    continue            # 名字不带「方向+部位」，这条规则管不了
                k = (wi_core._literal_key(name), chk.get("row"))
                if k in out:
                    continue
                c, _why = pick_col(stacks, name)
                if c is None:
                    continue
                r = pick_row(g, chk.get("spec"), hr)
                if r is None:
                    continue
                got[k] = {"page": tab["page"], "tno": tab["tno"],
                          "row": r, "col": c,
                          "row_label": (g[r][0] if r < len(g) and g[r] else ""),
                          "col_header": (stacks[c] if c < len(stacks) else "")[:60],
                          "pdf_name": "", "fp": rule.get("fp") or "",
                          "excel": name, "excel_row": chk.get("row"),
                          "from_rule": True, "stale": False,
                          "rule_label": rule.get("label") or ""}
            if got:
                out.update(got)
                break                   # 这张表能读通，不再试同指纹的其它表
    return out


# ---------------------------------------------------------------- 统计
def stats():
    lex = load()
    a = lex.get("anchors") or {}
    return {"anchors": sum(1 for r in a.values() if r.get("enabled", True)),
            "anchors_off": sum(1 for r in a.values() if not r.get("enabled", True)),
            "path": _CACHE["path"]}
