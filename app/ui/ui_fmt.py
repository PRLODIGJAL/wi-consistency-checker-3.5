"""纯函数：把判定结果、数值、状态翻成人话。

从 app.py 拆出来的（2026-11-03 第二轮）。这里**没有任何 self、不碰磁盘、
不建控件**——给它一组 check 字典，它还你一段能贴进表格或报告的文字。
想验证措辞对不对，单独 import 这个文件就能跑，不必拉起整个窗口。
"""


# ── 目录四层自举（与其他模块同一套写法，幂等）────────────────────────────
import os as _os
import sys as _sys
_APP_DIR = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))
if _APP_DIR not in _sys.path:
    _sys.path.insert(0, _APP_DIR)
import paths as _paths                                   # noqa: E402
_paths.install()

import os
import re
import unicodedata
import tkinter as tk

from ui_const import (BAD_SCOPES, DEFAULT_SHEET_ALIASES, HARD_SCOPES,
                      SCOPE_ORDER, _BAD_RATE_MIN_N, _BAD_RATE_TRIG,
                      _SCOPE_ORDER_MAP, _VALUE_RE)



def _looks_like_value(s):
    return bool(_VALUE_RE.search(str(s or "")))




def preferred_sheet(names):
    """在打开的工作表名里挑出默认该选的那张（表4）。挑不到返回 None。"""
    for n in names or []:
        if str(n).strip().replace(" ", "").replace("　", "") in DEFAULT_SHEET_ALIASES:
            return n
    return None




# ── 拖文件进来的两块零件（2026-10-13 用户要求「一口气拖多个文件」）────────
# 拆成模块级函数：它们不依赖窗口已经建好，回归脚本不建界面也能直接测。
def drop_paths(data: str, splitlist=None) -> list:
    """`<<Drop>>` 事件的数据 → 一串路径。

    ⚠️ 不能拿 `.split()` 硬切：tkdnd 丢过来的是**一个 Tcl list**，路径里有
       空格 / 花括号时那一项会被花括号整个包起来，硬切会把一条路径劈成两条
       不存在的路径。`splitlist` 传窗口的 `tk.splitlist`（Tcl 官方解析）。
    """
    text = data or ""
    if not text.strip():
        return []
    try:
        items = list(splitlist(text)) if callable(splitlist) else text.split()
    except Exception:  # noqa: BLE001 —— 解析器不认就退回按空白切
        items = text.split()
    return [s for s in (str(p).strip() for p in items) if s]




def _walk_widgets(root, limit=400) -> list:
    """`root` 之下所有后代控件（不含 root 自己），深度优先。

    拖放要逐个控件注册（见 `App._setup_dnd`），所以需要一个「把整棵树列出来」。
    `limit` 只是防呆：万一哪天界面里塞了动态生成的控件，别让注册动作失控。
    """
    out, stack = [], [root]
    while stack and len(out) < limit:
        w = stack.pop()
        try:
            kids = list(w.winfo_children())
        except Exception:  # noqa: BLE001 —— 控件已销毁
            kids = []
        for k in kids:
            if len(out) >= limit:
                break
            out.append(k)
            stack.append(k)
    return out




def _short_names(paths, maxw=None) -> list:
    """把一串完整路径变成「只显示文件名」，给清单用（2026-10-28 用户要求）。

    为什么不是简单 `basename` 了事：
      · 同一批里**可能有两个同名文件**（不同文件夹、同一个名字，比如两个
        项目下各有一份「表4.xlsx」）。这时只显示名字就分不出谁是谁，
        所以**仅在重名时**把上级文件夹名补在前面（`[项目A] 表4.xlsx`）；
        不重名的一律只留文件名，界面干净。
      · `maxw=None`（默认）**不截断** —— 清单本身带横向滚动条，名字再长也能
        拉着看全（2026-10-29 用户要求：「反正有拖进度条可以显示更全面一点」）。
        只有确实需要定宽的地方才传数字。
    ⚠️ 这里只改**显示**，`self.pdf_list` / `self.xlsx_list` 里存的仍是完整
       路径 —— 比对、导出、写盘记忆全都靠真路径，显示层不能污染数据层。
    """
    raw = [str(p or "") for p in paths]
    bases, count = [], {}
    for p in raw:
        b = os.path.basename(p.rstrip("\\/")) or p
        bases.append(b)
        count[b] = count.get(b, 0) + 1
    out = []
    for p, b in zip(raw, bases):
        s = b
        if count[b] > 1:
            par = os.path.basename(os.path.dirname(p.rstrip("\\/")))
            if par:
                # 重名时把**能区分的那个词**放在最前面：两个长文件名的前 40
                # 个字往往一模一样，差异藏在后面，放尾巴等于看不见。
                s = "[%s] %s" % (_fit(par, 14), b)
        out.append(s if maxw is None else _fit(s, maxw))
    return out




def _disp_width(s):
    """估算字符串的显示宽度：全角算 2、半角算 1。

    Tk 的 tab stop 是按屏幕距离算的，而中文一个字顶两个半角 ——
    不估算宽度就没法判断某一段会不会把后面的列挤走。
    """
    return sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
               for ch in str(s))




def _fit(s, maxw):
    """超宽就截断加省略号。

    摘要表是按固定列位排的：第一列一旦超过它的列宽，Tk 会把后面的内容
    顶到下一个 tab stop，整块表格就歪了 —— 宁可截断，也不能错位。
    """
    s = str(s)
    if _disp_width(s) <= maxw:
        return s
    out, w = [], 0
    for ch in s:
        cw = 2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1
        if w + cw > maxw - 1:
            break
        out.append(ch)
        w += cw
    return "".join(out) + "…"




def _join_card_item(card, item, maxw):
    """把「参数卡 · 采集项」拼成一列，超宽时先让卡片名让位。

    直接整体截断的话，像「DWG-0000102-亲水和亲水后正压测漏-0006 · 亲水和…」
    这种长卡片名会把采集项名整个吃掉 —— 而「错的是哪个参数」恰恰是最要紧的
    信息，宁可把卡片名截短。
    """
    full = "%s · %s" % (card, item)
    if _disp_width(full) <= maxw:
        return full
    room = maxw - _disp_width(item) - 3          # 3 ≈ 「 · 」
    if room >= 8:
        return "%s · %s" % (_fit(card, room), item)
    return _fit(full, maxw)                      # 采集项名本身就超宽，只能整体截




def _fmt(v):
    """把 225.0 显示为 225，225.5 保持 225.5；兼容数字字符串(如 '14.5')。"""
    try:
        f = float(v)
    except (ValueError, TypeError):
        return str(v)
    return str(int(f)) if f == int(f) else str(f)




def _one_sided_side_cn(er):
    """Excel 只填了一侧时，判断是哪一侧 —— 「上限」/「下限」，两列都有值或都空则「一侧」。

    直接从 lo/hi 两列判断，**不去读引擎的内部标记**：界面/导出拿到的行是
    compare 的返回副本，多余字段未必一路带过来，用数据本身判断最稳。
    """
    lo, hi = er.get("lo"), er.get("hi")
    if hi is not None and lo is None:
        return "上限"
    if lo is not None and hi is None:
        return "下限"
    return "一侧"




def _row_num(er):
    """只要行号（摘要区用）。

    「DWG-0000009-球囊成型-0001 · 加热中间模/6(℃)」这一串会把摘要的位置列占掉
    半行，而回 Excel 里改的时候真正要找的是**第几行** —— 卡片名在下方明细表里
    本来就有（2026-09-29 用户要求：「前面一堆只写第几行就好」）。
    """
    r = er.get("row")
    return "—" if r is None else str(r)




def _row_label(er):
    """这条数据在 Excel 里的行号（明细表最左列）。

    判定结果最终要回 Excel 里改，看到「哪张卡的哪一项」还不够 —— 同名采集项
    在同一张表里可能出现几十次，没有行号就得靠肉眼翻。
    没有行号信息（如「适用规格」门禁行，它不属于任何一行）就显示「—」。
    """
    r = er.get("row")
    if r is None:
        return "—"
    sh = er.get("sheet")
    return ("%s!%d" % (sh, r)) if sh else str(r)




def _yn_cn(v):
    """Excel 里「是/否」那一格 → 界面显示。

    空格显示「—」而不是冒充「否」：没填和填了「否」是两回事，
    一律显示「否」会把「这一列整片没填」这类系统性问题藏起来。
    """
    if v is None:
        return "—"
    s = str(v).strip()
    if not s:
        return "—"
    if s in ("是", "Y", "y", "True", "true", "1"):
        return "是"
    if s in ("否", "N", "n", "False", "false", "0"):
        return "否"
    return s




def _flag_cn(b):
    """按类型规矩算出来的 True / False / None → 界面显示。

    None = 「规矩管不到这一项」（认不得的类型），显示「—」，不硬凑一个「否」。
    """
    if b is None:
        return "—"
    return "是" if b else "否"




def _value_str(er):
    """根据类型生成 Excel 侧「现值」：浮点→范围，固定值→数值，布尔值→真/假。

    表头元信息项（机器编号）没填在任何数值列里 —— 它的值就是该行的设备名，
    所以要单独回退到 equip，否则「现值」列会是一片「—」。
    """
    if er.get("lo") is not None:
        if er.get("hi") is None:
            # 只填了下限、上限留空：这不是「没填值」，是 Excel 只给了一侧。
            # 以前 lo 有值就无条件拼 f"{lo} ~ {hi}"，hi 是 None 时会显示成
            # 「0.52 ~ None」；更糟的是「只填上限」连这个分支都进不来，
            # 现值直接显示「—」—— 看着就像 Excel 里什么都没有。
            return f"≥ {_fmt(er['lo'])}"
        return f"{_fmt(er['lo'])} ~ {_fmt(er['hi'])}"
    if er.get("hi") is not None:
        # 只填了上限（WI-0039 的「压力（Mpa）」就是这种：上限列 0.52、下限留空）。
        return f"≤ {_fmt(er['hi'])}"
    if er.get("fixed") is not None:
        return str(er["fixed"])
    if er.get("meta_kind") and er.get("equip"):
        return str(er["equip"])
    if er.get("btrue") is not None:
        return str(er["btrue"])
    if er.get("bfalse") is not None:
        return str(er["bfalse"])
    return "—"




def _fmt_by_equip(chk, limit=4):
    """期望值说不清时，按设备分组如实列出候选，而不是丢一个并集给人猜。"""
    be = chk.get("by_equip") or {}
    if not be:
        return "？"
    out = []
    for i, k in enumerate(sorted(be)):
        if i >= limit:
            out.append("…另 %d 台" % (len(be) - limit))
            break
        vs = "、".join(_fmt(a) if a == b else f"{_fmt(a)}~{_fmt(b)}" for a, b in sorted(be[k]))
        out.append(f"{k}: {vs}")
    return " ； ".join(out)




def _fmt_meta_expected(chk, limit=None):
    """元信息（机器编号/程序号/保护套/芯轴）的「应为」：本身是文字，不是数值区间。

    其余路径的 expected 是 (下限, 上限) 元组，这里必须区别对待 ——
    直接按元组解包字符串会抛异常。
    """
    exp = chk.get("expected") or []
    if not exp:
        if chk.get("scope") == "no_equip_match":
            return "（PDF 里查无该设备）"
        if chk.get("scope") == "no_meta":
            return "（PDF 表头没有这一项）"
        return "—"
    out = [str(x) for x in (exp if limit is None else exp[:limit])]
    if limit is not None and len(exp) > limit:
        out.append(f"…共{len(exp)}套之一")
    return "、".join(out)




def _ref_suffix(chk):
    """PDF 那一格带 (ref) 时，「应为」后面要跟着写 (ref)。

    不写的话界面上「Excel 空着 / 应为 10」看着就是漏填却判了一致 ——
    实际上这是参考值、按类型规矩本就不填（2026-09-29 用户点名：
    「不然这样的一致容易让人误判为我忘记填了，但是又给判对了」）。
    """
    return "(ref)" if chk.get("ref_marked") else ""




def _fmt_dp(v, dp):
    """按 PDF 的原始小数位数显示一个数：dp=1 → 2.0、dp=0 → 2、dp=None → 老样子。

    2026-10-30 用户截图（光强 1/2）：「PDF 内是 1.5--2.0，但是这个显示方面
    出错了」—— 判定没错（Excel 那格是整数 2、PDF 写 2.0），错的是显示：
    应为值被 _fmt 抹成 1.5~2，跟 Excel 填的一模一样，人看不出差在哪。
    """
    if dp is None or v is None:
        return _fmt(v)
    try:
        return ("%." + str(int(dp)) + "f") % float(v)
    except (TypeError, ValueError):
        return _fmt(v)




def _fmt_exp_pair(chk, a, b):
    """一条期望值 → 文字（区间 a~b，单值 v），有 exp_dp 时按 PDF 位数还原。"""
    dp = chk.get("exp_dp") or None
    lo_dp = hi_dp = None
    if dp:
        try:
            lo_dp, hi_dp = dp[0], dp[1]
        except (TypeError, IndexError):
            lo_dp = hi_dp = None
    if a == b:
        return _fmt_dp(a, lo_dp)
    return "%s~%s" % (_fmt_dp(a, lo_dp), _fmt_dp(b, hi_dp))




def _fmt_expected(chk):
    """格式化「应为」：范围→a~b，单值→v；没有确定期望值时给按设备分组的候选。"""
    if chk.get("meta_kind"):
        return _fmt_meta_expected(chk)
    if not chk.get("expected"):
        return _fmt_by_equip(chk) if chk.get("by_equip") else "？"
    out = [_fmt_exp_pair(chk, a, b) for a, b in chk["expected"]]
    return "、".join(out) + _ref_suffix(chk)




def _fmt_expected_brief(chk, limit=3):
    """一致时也要把「应为」亮出来，只是太长时折叠 —— 让人能确认不是碰巧对上。"""
    if chk.get("meta_kind"):
        return _fmt_meta_expected(chk, limit=limit)
    exp = chk.get("expected") or []
    if not exp:
        return "—"
    out = [_fmt_exp_pair(chk, a, b) for a, b in exp[:limit]]
    if len(exp) > limit:
        out.append(f"…共{len(exp)}套之一")
    return "、".join(out) + _ref_suffix(chk)




def scope_note(chk):
    """一致时在「核对」列里的说明：写清这一致是唯一命中还是多套候选里撞上一个。"""
    scope = chk.get("scope")
    # 有些通道只置了 scope=multi，没填 `multi` 数量（千分尺查找表那条就是）。
    # 直接取 0 会拼出「0套候选之一」这种自相矛盾的话（2026-11-02 用户截图），
    # 所以退一步用期望值本身有几套来数。
    multi = chk.get("multi") or len(chk.get("expected") or []) or 0
    if scope == "precise":
        return "一致(设备+规格+保护套精确命中)"
    if scope == "meta_strict":
        return "一致(表头元信息唯一命中)"
    if scope == "meta_multi":
        return "一致(表头元信息多套候选之一)"
    if scope == "meta_all":
        # 候选虽多，但每一套都被现值命中 —— 一致与否跟「用哪一套」无关。
        return "一致(表头元信息候选全部命中)"
    if scope == "meta_list":
        return "一致(该设备在 PDF 设备清单里)"
    if scope == "spec_ok":
        return "一致(适用规格落在该档内)"
    if scope == "spec_vague":
        return "一致(适用规格含糊：跨在 PDF 多档之间)"
    if scope == "no_step":
        return "一致(N/A 正确：该档参数表没有此步骤)"
    if scope == "pdf_blank":
        # 与 no_step 不同：这一步在表里有，只是那一格写着 `------`。
        # 填 N/A 和留空都对，所以分开措辞，让人看得出是哪种。
        if str(chk.get("fixed") or "").strip().upper() in ("N/A", "NA"):
            return "一致(N/A 正确：PDF 该格本身是空值)"
        return "一致(PDF 该格本身是空值，不填正确)"
    if scope == "prog_strict":
        return "一致(已定位到附表)"
    if scope == "prog_multi":
        return f"一致(附表多套候选{multi}套)"
    if scope == "one_sided":
        # Excel 只填了一侧（下限或上限）时，那个数**未必**是精确值 ——
        # 措辞里必须点出来，否则会被当成「Excel 给的就是精确值、PDF 也一模一样」。
        side = _one_sided_side_cn(chk)
        return f"一致(Excel 只填了{side}，PDF 取值与它相同)"
    if scope == "one_sided_loose":
        side = _one_sided_side_cn(chk)
        return f"一致(Excel 只填了{side}，PDF 取值未超出这一侧)"
    if scope == "strict":
        return "一致(唯一值)"
    if scope == "anchor":
        # 11-02：这一行的 PDF 值不是「找名字找到的」，是**人指定了坐标**读出来的。
        # 措辞必须点明，否则看的人会以为工具自己认出了这个名字。
        where = (chk.get("anchor") or {}).get("label") or ""
        return f"一致(按你钉的位置读数{('：' + where) if where else ''})"
    if scope == "anchor_loose":
        side = _one_sided_side_cn(chk)
        where = (chk.get("anchor") or {}).get("label") or ""
        return (f"一致(按你钉的位置读数{('：' + where) if where else ''}；"
                f"Excel 只填了{side}，PDF 落在该侧内)")
    if scope == "step_down":
        return "一致(按直径档推算)"
    if scope == "name_only":
        return "一致(仅按名称命中)"
    if scope == "multi":
        return f"一致({multi}套候选之一)"
    if scope == "device":
        return f"一致(设备级{multi}套之一)"
    if scope == "union":
        sub = chk.get("sub") or ""
        return f"一致(规格跨多张子表取并集{('：' + sub) if sub else ''})"
    if scope == "union_same":
        # 与 union 的区别要说出来：并集里没有分歧，所以这一致是「唯一」级别的。
        sub = chk.get("sub") or ""
        return f"一致(规格跨多张子表、取值一致{('：' + sub) if sub else ''})"
    if scope == "approx":
        sub = chk.get("sub") or ""
        return f"一致(按直径近似{(' ' + sub) if sub else ''})"
    if scope == "ordinal":
        # 这一行的档位是**按位次**对出来的，不是按规格算的 —— 措辞里必须说清，
        # 否则会被当成和「规格级唯一命中」一样硬。
        sub = chk.get("sub") or ""
        return f"一致(按规格位次近似{('对到 ' + sub) if sub else ''})"
    if scope == "no_spec_match":
        return "无对应(PDF 里没有与该规格对应的档位)"
    # —— 分层回退：一致也要说清是第几层退出来的，别让人当成「唯一命中」——
    if scope == "fb_kw":
        return "一致(回退第1层：关键词)"
    if scope == "fb_equip":
        return "一致(回退第2层：设备号/单位)"
    if scope == "fb_num":
        return "一致(回退第3层：规格数字)"
    if scope == "fb_fuzzy":
        return "一致(回退第4层：值集近似，需人工确认)"
    if multi > 1:
        return f"一致({multi}套候选之一)"
    return "一致"




def scope_buckets(judged):
    """按判定强度分档计数（未标注的归入 'other'）。"""
    known = set(_SCOPE_ORDER_MAP)
    out = {}
    for v in judged:
        s = v.get("scope") or "other"
        out[s if s in known else "other"] = out.get(s if s in known else "other", 0) + 1
    return out




def scope_summary_lines(judged):
    """生成「判定强度」两行：先给硬/宽松总数，再给逐档明细。"""
    b = scope_buckets(judged)
    hard = sum(b.get(k, 0) for k in HARD_SCOPES)
    bad = sum(b.get(k, 0) for k in BAD_SCOPES)
    soft = len(judged) - hard - bad
    head = (f"判定强度：硬判定（期望值唯一，或结论本身确定）{hard} 条"
            f" / 宽松判定（候选不止一个，或 Excel 只填了一侧限值）{soft} 条")
    if bad:
        head += f" / 判为不一致 {bad} 条"
    detail = "；".join(f"{name} {b[k]} 条" for k, name in SCOPE_ORDER if b.get(k))
    if b.get("other"):
        detail += ("；" if detail else "") + f"未标注 {b['other']} 条"
    return [head, "　· " + detail] if detail else [head]




def _plain_verdict(c, n_gap=0, soft=0, bad=0):
    """把几个计数翻成一句人话 —— 窗口最上面那一行。

    人打开窗口问的是「这份表到底有没有问题、我该干什么」，不是「一致多少条」。
    所以这里先说结论、再点出去哪儿看；数字排在句子里，不再单独占一张计数表。
    （计数表原来单独成段，一眼扫过去全是「0 条」，真正要看的「3 条填错」被
    三个 0 稀释掉了。）
    """
    n_ok, n_bad, n_none = c["ok"], c["mismatch"], c["no_match"]
    n_ns = c.get("no_scope", 0)
    n_multi = c.get("multi", 0)
    total = n_ok + n_bad + n_none + n_ns + n_multi
    # 分句先各自拼好再统一连 —— 直接往上一句屁股后面接「、%d 条…」，会拼出
    # 「…没有填错的、1 条 PDF 里找不到对应」这种半截话（第一版就错在这儿）。
    rest = []
    if n_bad:
        rest.append("%d 条填错要改" % n_bad)
    if n_multi:
        # 2026-11-09 改口径：多套候选**按不一致计**（用户：「多套就直接判错呗……
        # 它一个是给判对，又不标红，这样这个机器又学习不到」）。
        # 措辞不能说「填错」—— PDF 里那个参数确实按设备/规格分了好几档，
        # 是工具没定出该用哪一档；但也不能再让人以为「可以不管」。
        rest.append("%d 条取数不唯一（按不一致计）" % n_multi)
    if n_none:
        # 「（PDF 已读到，只是没匹配上）」→「在 PDF 里找不到对应」。
        # 后半句原来是在解释「读到了 vs 没读到」的区别，但对看表的人来说，
        # 他要知道的是「这一项在 PDF 里到底有没有答案」。而实测有一类情况
        # 恰恰是**PDF 里根本没有那一项**（WI-0039：正文说「如下表7」，全篇却
        # 没有表7），写「只是没匹配上」等于替文件开脱（2026-10-30 改）。
        rest.append("%d 条在 PDF 里找不到对应" % n_none)
    if n_ns:
        rest.append("%d 条 Excel 没填设备号判不了" % n_ns)
    if not rest:
        # 全干净时别再报一遍同一个数 ——「核了 14652 条，14652 条对得上」很蠢。
        out = ["可核对 %d 条，全部对得上。" % total]
    else:
        out = ["可核对 %d 条，%d 条对得上，%s。" % (total, n_ok, "、".join(rest))]
    # 「要处理的都在下面，先看它们。」删于 2026-09-23 —— 结论句已经说了有几条要改，
    # 紧接着就是【要处理的】段头，再说一句「下面看」纯属占位。
    if n_gap:
        out.append("「找不到对应」里有 %d 条是「规格对不上任何一档」。" % n_gap)
    if n_multi:
        out.append("取数不唯一的行，是 PDF 里同一个参数按设备/规格/程序分了好几档，"
                   "工具没能收窄到唯一的一档 —— 已按不一致计入。"
                   "请双击那一行、在右侧 PDF 上点中正确的那一格再「钉住这格」，"
                   "同一张表下次就直接按位置取数。")
    if soft:
        out.append("另有 %d 条判了「一致」但口径没锁到最严（候选不止一个，"
                   "或 Excel 只填了一侧限值），采用前建议对着 PDF 过一眼。" % soft)
    return "".join(out)




def _unpassed_rate(c):
    """未通过率 = 没判「一致」的占已判定的比例。达标的返回 (比率, 未通过数, 总数)。

    为什么要有这一条（2026-09-28 用户的实际场景）：工具是要交给**产线员工**用的，
    他们只看一个结论。
      · 全对 → 就是全对，不用解释；
      · 错一两条 → 照常逐条标出来，员工去改表，这是工具的正常用法；
      · **错一大片 → 通常不是表填错了，是这份 PDF 没被读透**（换了版式、
        掺了扫描页、PDF 语法畸形……），这时候照着"不一致"去改表，会把本来
        没错的数据改坏 —— 必须先把结论的可信度说清楚。

    它**不是判定口径**，只是给看结果的人提个醒；不参与计数、不改任何 status。
    """
    total = (c["ok"] + c["mismatch"] + c["no_match"] + c.get("no_scope", 0)
             + c.get("multi", 0))
    if total < _BAD_RATE_MIN_N:
        return None
    n_bad = total - c["ok"]
    rate = n_bad / total
    return (rate, n_bad, total) if rate >= _BAD_RATE_TRIG else None




def _unpassed_hint(c):
    """未通过率达标时要说的那句话；没达标返回 None。

    「反正都建议再核一遍，一句话就够了」——不够。同样是未通过率 60%，
    **主体是「找不到对应」还是「取值不一致」，该给的建议完全相反**：
      · 找不到对应为主 → 这份 PDF 没被读透，照着改表会把没错的数据改坏；
      · 取值不一致为主 → 更可能是参数卡真的跟 PDF 对不上（换过版本 / 拿错文件），
        这时候第一件事是确认「这两份文件是不是同一批」。
    所以一句「核对一下」不够，得把方向指对；但也就到此为止，不多解释。
    """
    got = _unpassed_rate(c)
    if not got:
        return None
    rate, n_bad, total = got
    head = "未通过率 %d%%（%d/%d）" % (round(rate * 100), n_bad, total)
    if c["no_match"] + c.get("no_scope", 0) > c["mismatch"]:
        tail = "：大部分是「PDF 里找不到对应」，多半是没读透，先别照着改表"
    else:
        tail = "：这么大一片取值对不上，先确认 PDF 与参数卡是不是同一批文件"
    return head + tail + " —— 建议人工再核一遍，或把这份 PDF 丢给 WorkBuddy 再核对。"




def _norm_item(s):
    return re.sub(r"\s+", "", str(s or ""))




def sibling_value_hints(schema, checks):
    """「无对应」时的核对线索：同名采集项是否写在其它 WI 的参数卡里。

    用途：有的采集项在 Excel 里挂在 A WI 名下，取值却出自 B WI 的作业指导书
    （例：DWG-0000102 表4 把「测试压力 14~14.5atm」挂在 WI-0102 下，而
    14+0.5/-0 atm 实际写在 WI-0010 的表1 里）。这时只换 A WI 的 PDF 永远
    「无对应」—— 这里如实地把线索指出来，由人决定用哪份 PDF 核对。

    例外：scope=no_spec_match 的「无对应」不是「这份 PDF 里没这个参数」，
    而是「有这个参数、但规格对不上任何一档」—— 这时提示「换一份 WI 的 PDF」
    是往错误方向指，所以整类跳过。
    """
    # 没有一条「无对应」就没什么可给的 —— 而下面要整表读一遍（大表要好几秒、
    # 几十 MB），这个提前返回别省。
    if not any(c.get("status") == "no_match" and c.get("scope") != "no_spec_match"
               for c in checks):
        return {}
    try:
        all_recs = schema.records(None)
    except Exception:
        return {}

    def _val(r):
        if r.get("lo") is not None:
            return _fmt(r["lo"]) if r["lo"] == r["hi"] else f"{_fmt(r['lo'])}~{_fmt(r['hi'])}"
        if r.get("fixed") is not None:
            return _fmt(r["fixed"])
        return None

    # 先按「归一化采集项名」建一次索引：否则每条「无对应」都要把全表重扫一遍，
    # 几千条记录 × 上百条无对应 = 反复做同一份字符串归一。
    by_item = {}
    for r in all_recs:
        by_item.setdefault(_norm_item(r.get("item")), []).append(r)

    hints = {}
    for c in checks:
        if c.get("status") != "no_match":
            continue
        if c.get("scope") == "no_spec_match":
            continue
        key = _norm_item(c.get("item"))
        if not key:
            continue
        hits = {}
        for r in by_item.get(key, ()):
            if (r.get("wi") or "") == (c.get("wi") or ""):
                continue
            v = _val(r)
            if v is None:
                continue
            hits.setdefault(r.get("wi") or "未知WI", set()).add(v)
        if hits:
            hints[(c.get("card"), c.get("item"))] = {
                wi: sorted(vs) for wi, vs in hits.items()}
    return hints
