# -*- coding: utf-8 -*-
"""tab_health.py —— 判断 PDF 里每张表「读懂了没有」。

为什么要有（2026-11-10）：
  界面上现在只有「引擎抽出 N 条」这个绝对数，可 N=0 有两种完全不同的含义：
    · 物料明细表 / 工具设备明细表 —— 本来就没有参数，抽 0 条是对的；
    · 开口尺寸设定表 —— 明明有参数，一条都没抽到，这是真漏抽。
  两种混在一个 0 里，人就没法判断「是没抽出来」还是「抽出来但没对应上」。
  这个模块只回答一件事：**这张表读懂了没有、没读懂是因为什么。**

判定口径（只用表内能看到的证据，不看机器学出来的东西）：
  not_param  不像参数表：表名/首列/表头里没有参数类字样 —— 抽 0 条正常，不报
  bad        像参数表，却一条都没抽到                    —— 需检查（真漏抽）
  weak       抽到了但很零散（抽出/行数 < 20%）或表头没认出来 —— 需检查
  ok         其余

为什么用文字特征而不是「行数/条数比」单指标：
  实测 157 张表，比值 <0.01 的有 73 张（46%），其中绝大多数是物料明细表。
  只按比值判，会把半份文档都标成需检查，等于没标。先分「该不该有参数」，
  再判「抽得够不够」，才分得开。
"""
from __future__ import annotations

# 出现在表名 / 首列 / 表头里，说明「这张表该有参数」
PARAM_WORDS = (
    "参数", "设定", "尺寸", "规格", "工艺", "作业", "标准值",
    "直径", "长度", "压力", "速度", "时间", "温度", "功率", "强度",
)

# 明确不是参数表的东西（优先级高于 PARAM_WORDS）
NOT_PARAM_WORDS = (
    "物料明细", "物料明显", "物料清单", "bom", "工具设备明细", "设备明细表",
    "变更记录", "修订记录", "版本历史", "签核", "会签",
)

# 表头里出现这些，说明列结构被认出来了
HEADER_HINTS = ("参数名称", "参数值", "参数", "名称", "数值", "设定值", "单位")

# 少于这么多行的，装不下一张参数表 —— 实测里都是封面/页眉信息条
# （第一版没这条，`rows=2` 的「工序名称｜文件编号」条被当成参数表报了 16 次假阳性）
MIN_ROWS_FOR_PARAM = 5

MIN_ROWS_FOR_RATIO = 8      # 行数太少时，比值没意义（3 行抽出 1 条也是 33%）
RATIO_WEAK = 0.2            # 抽出 / 行数 低于这个，算零散
BLANK_ROWS_RATIO = 0.15     # 有内容的行 / 总行 低于这个，判「空白/纯图页」


def _text_of(cat_item: dict) -> str:
    """把这张表所有能读到的文字拼起来（表名 + 副标题 + 表头 + 首列）。"""
    parts = [
        cat_item.get("name") or "",
        cat_item.get("en") or "",
        cat_item.get("title") or "",
    ]
    parts.extend(str(x) for x in (cat_item.get("headers") or []))
    parts.extend(str(x) for x in (cat_item.get("first_col") or []))
    return " ".join(parts).lower()


def kind_of_table(cat_item: dict) -> str:
    """这张表像不像参数表。返回 param / not_param / unknown。

    「明确不是参数表」的词优先 —— 有的物料明细表里也会出现「规格」列名，
    那是物料规格，不是工艺参数，不能因此当参数表。
    """
    txt = _text_of(cat_item)
    if any(w in txt for w in NOT_PARAM_WORDS):
        return "not_param"
    if any(w in txt for w in PARAM_WORDS):
        return "param"
    return "unknown"


def assess(cat_item: dict, page_points: int | None = None) -> dict:
    """给 `pdf_tables.catalog()` 的一项补一份体检结论。

    Args:
        cat_item: catalog() 返回的一条（要有 page / rows / nrec / headers / first_col）。
        page_points: 这一页被**四条抽取路径**（横表/竖表/程序附表/千分尺）一共产出
            多少条记录。不传就只按本表的 nrec 判 —— 那是错的，见下面「为什么必须有它」。

    Returns:
        {"level": "ok"|"weak"|"bad"|"not_param",
         "ratio": float,          # 抽出 / 行数，行数为 0 时给 0.0
         "why": str,              # 一句话原因，直接给界面用
         "kind": "param"|"not_param"|"unknown"}

    为什么必须有 page_points（2026-11-10 实测踩坑）：
        `catalog` 的 `nrec` 只统计**竖表**那一路的记录。WI-0011 第 8~11 页的
        「开口/闭口尺寸设定表」竖表确实是 0 条，但千分尺路径在这四页分别抽到了
        42/36/20/23 条 —— 数据是读到了的，只是走的是另一条路。
        只看 nrec 就会把 4 张正常表标成「漏抽」，正是用户最反感的假警报。
        判「这页到底有没有被读懂」必须以四路合计为准。
    """
    rows = int(cat_item.get("rows") or 0)
    nrec = int(cat_item.get("nrec") or 0)
    nhead = len(cat_item.get("headers") or [])
    # 分母用「有内容的行数」而不是物理行数：参数表的物理行里常常一半是空行，
    # 用物理行算比值会把「4 个参数抽全了」的窄表（WI-0024 第 6 页，32 行里只有
    # 4 个参数行）判成漏抽。实测这一条是最后一处误报的来源。
    denom = int(cat_item.get("filled_rows") or 0) or rows
    ratio = (nrec / float(denom)) if denom else 0.0
    kind = kind_of_table(cat_item)
    page_pts = None if page_points is None else int(page_points)

    if cat_item.get("kind") != "table":
        return {"level": "not_param", "ratio": ratio,
                "why": "页眉条", "kind": kind}
    if kind == "not_param":
        # 物料表里偶尔也能抽到几条（「规格」列被当成参数名了）。这种情况不能写
        # 「无参数」—— 表格自己就矛盾了。照实说抽到几条，但仍按物料表处理，
        # 不计入「参数表」的体检统计（它本来不该参与规格核对）。
        if nrec > 0:
            return {"level": "not_param", "ratio": ratio,
                    "why": "物料表，已抽 %d 条" % nrec, "kind": kind}
        return {"level": "not_param", "ratio": ratio,
                "why": "物料明细表，无参数", "kind": kind}
    if rows < MIN_ROWS_FOR_PARAM:
        return {"level": "not_param", "ratio": ratio,
                "why": "仅 %d 行，信息条" % rows, "kind": kind}
    # 整页空白 / 纯图页：表壳在、格子里没字。WI-0023、WI-0024 的第 2 页就是这样
    # （35 行里只有页眉那 1 行有字）。判据用「有内容的行数占比」而不是「非空格占比」：
    # 窄竖表（WI-0024 第 6 页，参数只在前两列）用格子占比会被误伤成空白页。
    filled = cat_item.get("filled_rows")
    if filled is not None and rows and filled / float(rows) < BLANK_ROWS_RATIO:
        return {"level": "not_param", "ratio": ratio,
                "why": "空白/图页（%d 行有字）" % int(filled), "kind": kind}

    # 到这里说明「像参数表」或「判断不了」。
    if nrec == 0:
        if page_pts:                       # 同页别的表读到了 —— 数据在，只是不在这张表
            return {"level": "ok", "ratio": 0.0,
                    "why": "本表 0 条，同页另读到 %d 点" % page_pts, "kind": kind}
        if kind == "param":
            return {"level": "bad", "ratio": 0.0,
                    "why": "像参数表，整页 0 条", "kind": kind}
        return {"level": "weak", "ratio": 0.0,
                "why": "整页 0 条，不像参数表", "kind": kind}

    if nhead < 2:
        return {"level": "weak", "ratio": ratio,
                "why": "表头仅 %d 字段" % nhead, "kind": kind}
    if denom >= MIN_ROWS_FOR_RATIO and ratio < RATIO_WEAK:
        return {"level": "weak", "ratio": ratio,
                "why": "%d 个参数行只抽到 %d 条" % (denom, nrec), "kind": kind}

    return {"level": "ok", "ratio": ratio,
            "why": "抽 %d/%d 参数行" % (nrec, denom), "kind": kind}


def page_points(*sources) -> dict:
    """把多路抽取结果折成 {页码: 条数}，供 `assess` 判「这页到底读到没有」。

    能吃的形态（都在引擎里现成存在）：
      · 列表：横表 / 竖表 / 程序附表的记录列表
      · 字典：千分尺查找表 `{"open": [...], "pass": [...], "nopass": [...]}`

    Args:
        *sources: 任意多路记录，`None` 会被跳过。

    Returns:
        {页码(int): 条数}
    """
    out: dict[int, int] = {}

    def eat(rec):
        if not isinstance(rec, dict):
            return
        p = rec.get("page")
        if p:
            out[int(p)] = out.get(int(p), 0) + 1

    for src in sources:
        if not src:
            continue
        if isinstance(src, dict):
            for v in src.values():
                if isinstance(v, (list, tuple, set)):
                    for r in v:
                        eat(r)
        elif isinstance(src, (list, tuple, set)):
            for r in src:
                eat(r)
    return out


def assess_all(cat: list[dict], bypage: dict | None = None) -> list[dict]:
    """给整个 catalog 逐张补体检结论，返回**新列表**（不改原 dict）。

    Args:
        cat: `pdf_tables.catalog()` 的结果。
        bypage: `page_points(...)` 的结果。给了才能避免「竖表 0 条就喊漏抽」
            的假警报（见 `assess` 的说明）。不给仍然能跑，但结论偏严。

    Returns:
        每项 = 原 catalog 项 + `health` 字段（assess 的返回值）。
    """
    bypage = bypage or {}
    out = []
    for it in cat or ():
        d = dict(it)
        d["health"] = assess(it, bypage.get(it.get("page")))
        out.append(d)
    return out


def summary(cat_with_health: list[dict]) -> dict:
    """数一下体检结果：「需检查」有几张、能不能核。

    Returns:
        {"total": 参数表张数（不含页眉条与物料表）,
         "ok": n, "weak": n, "bad": n,
         "need_check": weak+bad,
         "all_tables": 目录里全部张数}
    """
    c = {"total": 0, "ok": 0, "weak": 0, "bad": 0, "need_check": 0,
         "all_tables": len(cat_with_health or [])}
    for it in cat_with_health or ():
        h = it.get("health") or {}
        lv = h.get("level")
        if lv == "not_param":
            continue
        c["total"] += 1
        if lv in ("ok", "weak", "bad"):
            c[lv] += 1
    c["need_check"] = c["weak"] + c["bad"]
    return c
