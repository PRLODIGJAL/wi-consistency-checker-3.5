# -*- coding: utf-8 -*-
"""结构自检 —— 把「一个正确答案必然满足的形状」写成不变量，跑完就查。

为什么需要它
------------
2026-11-02 一天里连着出了三个同类 bug（2534 多段多套、2535 应为两个值、
2535 把物料表当成针规表）。症状各不相同，机制是同一个：

    解析器只认识「我见过的写法」。遇到第二种写法时它返回 None，
    上层把 None 当成「文件里没有这个信息」继续往下走 —— 于是**静默变宽**：
    候选从 1 个变 2 个、从 2 个变 4 个，值仍然「对得上」，界面上一点错都不报。

这类错误靠「写代码更小心」避免不了（每次都是少认一种写法，还会有第四种）。
能避免的是「它不吭声」：把必然成立的形状写成断言，跑完立刻查。

两级结论（这个分级很重要，别合并）
--------------------------------
  errors —— 几乎一定是错的，值得停下来看；
  notes  —— 可能正常，只是「值得瞄一眼」，不该当错误报，否则天天狼来了。

纯函数：只吃 checks 列表，不碰文件、不碰界面、不改任何判定。
"""
from collections import Counter, defaultdict


def _pairs(v):
    """把 expected 拆成「套」：每一套是一个区间 (a, b) 或一个单值。

    ⚠️ 不能用「摊平后数不同值」来数候选：`(215, 218)` 是一个**区间**，
    摊平会变成 215/218 两个数，被误当成「两套候选」。
    （这套逻辑的第一版就这么错过一次，14490 行里误报 2100 条。）
    """
    if not v:
        return []
    if isinstance(v, (list, tuple, set)):
        items = list(v)
        if items and all(isinstance(x, (list, tuple, set)) for x in items):
            return [tuple(x) for x in items]
        return [tuple(items)]
    return [(v,)]


def _flat(v):
    out = []

    def walk(x):
        if isinstance(x, (list, tuple, set)):
            for y in x:
                walk(y)
            return
        if x is not None:
            out.append(x)
    walk(v)
    return out


def _row(chk, kind, note):
    return {"kind": kind, "note": note,
            "wi": chk.get("wi") or "", "card": chk.get("card") or "",
            "item": str(chk.get("item") or ""), "spec": chk.get("spec") or "",
            "status": chk.get("status") or "", "row": chk.get("row"),
            "expected": chk.get("expected")}


def audit(checks):
    """扫一遍判定结果。返回 {"errors": [...], "notes": [...]}。"""
    errors, notes = [], []
    checks = [c for c in (checks or []) if isinstance(c, dict)]

    # ── 错误级 ────────────────────────────────────────────────────────
    # E1 判「一致」却没有任何期望值 —— 那这个一致是凭什么判的
    for c in checks:
        if c.get("status") == "ok" and not _flat(c.get("expected")):
            errors.append(_row(c, "一致但期望值为空",
                               "status=ok，但 expected 是空的"))

    # E2 期望值里出现与该行直径相同的数 —— 典型的「隔壁表串台」
    #    （WI-0043 第3页物料表的 NO 列被当成针规值，dia=4 / val=4.0）
    for c in checks:
        dia = c.get("dia")
        try:
            d = float(dia) if dia is not None else None
        except (TypeError, ValueError):
            d = None
        if d is None:
            continue
        for v in _flat(c.get("expected")):
            try:
                fv = float(v)
            except (TypeError, ValueError):
                continue
            if abs(fv - d) < 1e-9:
                errors.append(_row(c, "期望值等于直径（疑似串到隔壁表）",
                                   "直径 %s，期望值里却有 %s" % (dia, v)))
                break

    # E3 「过」与「不过」取到同一组值 —— 找列时把同一列读了两次
    pair = defaultdict(dict)
    for c in checks:
        it = str(c.get("item") or "")
        if "不过" in it:
            pair[(c.get("wi"), it.replace("不过", "过"))]["nopass"] = c
        elif "过" in it:
            pair[(c.get("wi"), it)]["pass"] = c
    for _k, d in pair.items():
        a, b = d.get("pass"), d.get("nopass")
        if not a or not b:
            continue
        ea, eb = _pairs(a.get("expected")), _pairs(b.get("expected"))
        if ea and ea == eb:
            errors.append(_row(a, "过/不过取到同一组值",
                               "「过」与「不过」的期望值完全相同：%s" % (ea,)))

    # E4 同一采集项里候选套数不一致 —— 收窄只对一部分行生效
    #    （WI-0043 修之前就是：有的行 1 套、有的 2 套、有的 3 套）
    grp = defaultdict(list)
    for c in checks:
        grp[(c.get("wi"), str(c.get("item") or ""))].append(c)
    for _k, rows in grp.items():
        if len(rows) < 3:
            continue
        ns = Counter(len(_pairs(r.get("expected"))) for r in rows)
        if len([n for n in ns if n > 0]) > 1:
            worst = max(rows, key=lambda r: len(_pairs(r.get("expected"))))
            errors.append(_row(worst, "同一采集项候选套数不一致",
                               "该采集项 %d 行，候选套数分布 %s（收窄可能只对部分行生效）"
                               % (len(rows), dict(ns))))

    # ── 提示级 ────────────────────────────────────────────────────────
    # N1 判「一致」但给了多套候选 —— 允许，是「撞上其中一个」，必须让人看见
    for c in checks:
        if c.get("status") != "ok":
            continue
        ps = _pairs(c.get("expected"))
        if len(ps) > 1:
            notes.append(_row(c, "多套候选中撞上一个",
                              "候选 %d 套：%s" % (len(ps), ps)))

    # N2 规格不同、期望值却完全一样 —— 多数是「PDF 只有一张通用表」的正常情形
    #    （实测 WI-0040 焊接温度：那张表按设备组列，根本没有规格列，各规格同值是对的）
    for _k, rows in grp.items():
        spec_of = {}
        for r in rows:
            sp = str(r.get("spec") or "")
            if not sp:
                continue
            spec_of.setdefault(sp, []).append(tuple(_pairs(r.get("expected"))))
        if len(spec_of) < 2:
            continue
        exps = {v[0] for v in spec_of.values() if v}
        if len(exps) == 1 and next(iter(exps)):
            notes.append(_row(rows[0], "规格不同但期望值相同",
                              "%d 种规格都给出同一组值 %s（PDF 若是通用表则正常）"
                              % (len(spec_of), next(iter(exps)))))

    return {"errors": errors, "notes": notes}


def summary(a):
    """一行话，给状态栏用。"""
    a = a or {}
    ne, nn = len(a.get("errors") or []), len(a.get("notes") or [])
    if ne:
        return "结构自检：%d 条可疑（点开看）" % ne
    if nn:
        return "结构自检通过（%d 条提示）" % nn
    return "结构自检通过"
