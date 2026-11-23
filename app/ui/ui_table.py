"""明细表与概览的绘制：填行、上色、切页、列显隐。

从 app.py 的 App 类拆出来的（2026-11-03 第二轮）。
方法体一个字没改，只是搬了家；App 用多重继承把它们拼回去，
所以 `self.xxx` 的调用关系、执行顺序、窗口行为都跟拆之前完全一样。
"""

# ── 目录四层自举（与其他模块同一套写法，幂等）────────────────────────────
import os as _os
import sys as _sys
_APP_DIR = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))
if _APP_DIR not in _sys.path:
    _sys.path.insert(0, _APP_DIR)
import paths as _paths                                   # noqa: E402
_paths.install()

import re
from tkinter import messagebox
from datetime import datetime

import ui_theme as theme
from ui_const import BAD_SCOPES, HARD_SCOPES, _SEP, _SUB_TAG
from ui_fmt import _fit, _flag_cn, _fmt_by_equip, _fmt_expected, _fmt_expected_brief, _plain_verdict, _row_label, _row_num, _unpassed_hint, _value_str, _yn_cn, scope_buckets, scope_note

try:
    import xl_locate
except Exception:      # 缺了就降级，不影响核对
    xl_locate = None


_EQ_NO = re.compile(r"^(T-\d+)-")


def _slim_equip(p):
    """设备串里每一项只留设备号：`T-2068-DWG-0000007-03,04` → `T-2068`。

    2026-11-11 通用性巡检发现（成功1 语料）：采集项「过渡点焊接 机器编号」的值
    就是整串设备名，但每台后面还挂着自己的 DWG 件号（`T-2068-DWG-0000007-03,04`），
    每项 28 个字符。`_short_list` 按 `/` 留 3 段之后仍量出 592px，塞进 74px 的
    限值列里只剩 `T-2068-DWG20…` —— 人看不出那是几台设备。
    核对时真正要认的是设备号，件号在「依据」列和导出的全串里都还在。
    """
    m = _EQ_NO.match(str(p).strip())
    return m.group(1) if m else p


def _short_list(s, keep=3, unit_word="台"):
    """把「A/B/C/D/…」这样的长列表压成「A/B/C… 共 N 台」。

    只在**元信息行**上用（那几行的值是整串设备号）。单值原样返回，免得把
    「T-1030」这种也加上「共 1 台」；多项时顺便把每项的设备号后面的 DWG
    件号摘掉（见 `_slim_equip`），否则三项仍是六百像素宽。
    """
    t = str(s or "")
    parts = [p for p in t.split("/") if p.strip()]
    if len(parts) <= 1:
        return t
    slim = [_slim_equip(p) for p in parts]
    if len(parts) <= keep:
        return "/".join(slim)
    return "%s… 共 %d %s" % ("/".join(slim[:keep]), len(parts), unit_word)


def _fit_px(s, max_px, fnt):
    """按**像素**截断加省略号（分组条的「规格 · 设备」用）。

    为什么不复用 ui_fmt._fit：那是按「显示宽度」（CJK=2）近似的，给的是文本框
    这类等宽场景；分组条是 13 号加粗的比例字体，同样的字串两种算法能差出 40%，
    截断点会落在字中间。这里直接拿 font.measure 量，量到哪截到哪。
    fnt 给不了（极端环境）就原样返回，宁可被 Tk 裁一点也不要错截。
    """
    s = str(s or "")
    if not max_px or fnt is None:
        return s
    if fnt.measure(s) <= max_px:
        return s
    out = []
    for ch in s:
        if fnt.measure("".join(out) + ch + "…") > max_px:
            break
        out.append(ch)
    return "".join(out) + "…"


def _bulk_px(vals, q=0.9, spread=1.8):
    """一组「内容宽度」里**真正该拿来定列宽**的那个值：核心簇的分位，不是最大值。

    为什么不能用 max：一条散文式的异常内容就能把整列撑开，把后面所有列推到
    窗口右缘。实测 WI-0029 —— 182 行里 180 行的采集项只有 144px，另外 2 行
    把「采集项」写成了一整句话（`球囊远端焊接参数确认首件制作，参数与下面
    焊接参数一致/1PC`，358px）。按 max 定宽 → 采集项 382px → 类型 / 上下限 /
    超上下限 / 限值 四列整体右移 170px，中间白出一大块（用户 2026-11-11：
    「与左边那个判定结果中间隔了很大的位置，这是通用性的问题」）。

    做法：先按**中位数**圈出「核心簇」（不超过中位数 `spread` 倍的那些值），
    离群的那几条不参与定宽，然后在核心簇里取 q 分位。

    为什么不用「p90 + 最大值比值」那种写法（2026-11-11 上一版）：行数少的表
    会失效。WI-0028 一共 12 行，其中 10 行 88px、2 行 295px —— p90 落在
    int(12*0.9)=10 那一格上，正好是 295，于是「离群值」反而成了定宽依据，
    采集项列被撑到 319px，把右边的核对结果挤到只剩 215px，设备号被切掉。
    按中位数圈簇就与行数无关：88 的核心簇里没有 295，取出来就是 88。

    内容本来就都长（比如全表都是长名字）时，核心簇 == 全部，q 分位 ≈ max，
    不会无故把正常内容也截掉。
    """
    v = sorted(int(x) for x in vals if x)
    if not v:
        return 0
    med = v[len(v) // 2]
    core = [x for x in v if x <= med * spread] or v
    return core[min(len(core) - 1, int(len(core) * q))]


def _pending(chk):
    """这一行是不是「有事要管」—— 全表只有这一处口径。

    分组条上那句「N 条待处理」、自动收起哪些组、「只看异常」时留住哪些行，
    **三个地方必须数同一批行**。原来这三处各写一遍条件（`status not in (...)`
    抄了三遍），只要有一处改了口径，就会出现「表头说 3 条、屏幕上 4 行」
    这种谁也说不清的对不上。所以归到一个函数里。

    「一致 / 采集项 / 未填值」这三种不是错，也不需要人动手，其余都算待处理
    （不一致、找不到对应、缺设备线索、多套候选）。
    """
    return bool(chk) and chk.get("status") not in ("ok", "collection", "novalue")


def _card_split(card):
    """把参数卡名拆成「前缀 + 编号」：`DWG-0000009-球囊成型-0001` → (`DWG-0000009-球囊成型`, `0001`)。

    为什么要拆：这类名字前面那一大串在同一份表里几乎每行都一样（实测 486 行里
    481 行是同一个前缀），真正区分一张卡的是尾巴上的 `0001`。前缀当分组条上的
    「大表头」只出现一次，编号留在行里，眼睛就不用在一列长字符串里找不同
    （用户原话：「把那个 000几给抽取出来，然后把那个前面一大给它当做一个大的
    表头展示出来」）。

    认不出编号（整串没数字）就原样返回，前缀留空 —— 调用方会退回显示全名。
    """
    s = str(card or "").strip()
    m = re.match(r"^(.*?)(\d{2,})$", s)
    if not m:
        return "", ""
    pref = m.group(1).rstrip("-—_ /")
    return (pref, m.group(2)) if pref else ("", "")



class TableMixin:
    """明细表与分组的绘制：填行、上色、折叠、列宽自适应。"""


    def _show(self, content, alert=False):
        """原来往「概览」页写的那份长报告 —— 概览页删掉后（2026-11-03 第四轮），
        这里只剩「留存」和「该不该让人看见」两件事。

        用户原话：「那这样子的话是不是就可以把那个概览的功能给平替掉，然后概览
        那个按钮就可以删掉，以及那一系列功能可以删掉。」

        但 `_show` 的调用点里有一半是**跑不下去时的失败提示**（表格自检失败 /
        PDF 是扫描件 / 没取到采集项 / 批量核对失败），这些不能跟着一起消失 ——
        删了就变成「点了没反应」，正是用户最反感的那种。所以分两路：

        Args:
            content: str，或 [(行文本, 标签名), …] 的分段列表。
            alert: True 表示「这条必须让人看见」→ 弹窗；False 只留档。
        """
        if isinstance(content, str):
            blocks = [(content, None)]
        else:
            blocks = list(content)
        self._blocks = blocks
        if not alert:
            return
        lines = []
        for line, _tag in blocks:
            if isinstance(line, list):
                lines.append("".join(str(seg) for seg, _t in line))
            else:
                lines.append(str(line))
        text = "\n".join(lines).strip()
        if text:
            messagebox.showwarning("提示", text)



    # ------------------------------------------------------------ 分组（按参数卡）
    def _group_view(self, seq):
        """把 seq 铺成「分组条 + 数据行」，收起的组只留一条分组条。

        为什么按参数卡分（2026-11-03 用户定）：
          · 参数卡名是 `DWG-0000009-球囊成型-0001` 这种「一大串 + 000几」，
            同一份表里前缀往往只有两三种 —— 486 行里 481 行的前缀完全一样，
            按行重复印 486 遍，等于白占一列 152px 还把真正的信息挤没了；
          · 一张卡就是「一次设定」，复核时人本来就是一张卡一张卡地看。

        分组条本身也带数字（几行、几条待处理），所以**收起时**照样能看出哪组要管，
        不必先展开再数。
        """
        by, order, seed = {}, [], {}
        for er, chk in seq:
            card = str((er or {}).get("card") or "").strip() or "（无参数卡）"
            if card not in by:
                by[card] = []
                order.append(card)
            # 「规格 + 设备」在**过滤之前**先收一遍：带着整串设备号的正是元信息行，
            # 等下面把它过滤掉再回头找，分组条上那串设备号就整条消失了。
            # 一张卡 = 一次设定，卡内每行的 spec/equip 一样，取第一个非空即可。
            _e0 = er or {}
            _sp, _eq = seed.get(card, ("", ""))
            seed[card] = (_sp or str(_e0.get("spec") or ""),
                          _eq or str(_e0.get("equip") or ""))
            # 元信息行（「机器编号」这类）判成一致的**整行不进数据区**。
            # 2026-11-11 用户原话：「我圈起来的这一行，你看它其实没有任何的信息，
            # 而且它左侧也写的是机器编号，那这个就应该属于是那种设备号的一种，
            # 那这个地方就完全不需要核对，直接放到最上面那一栏。而且是，这就是
            # 最上面那一栏，它与它重复了」。
            #
            # 重复是事实：分组条上那串「设备 T-2101/T-1857」取的正是这一行的
            # `equip`，所以数据区再占两行（现值 + ↳ 应为）把同一串字印一遍，
            # 等于让人对着两个一模一样的东西「核对」。
            #
            # 只删**判成没问题**的。判出问题的那种（「这行填的设备在 PDF 里查无
            # 此机」「表头没有这一项」）必须留着 —— 那正是要人去改的一行，删了
            # 就变成静默放过。同时「N 行」「M 条待处理」都在这里数，删完之后
            # 表头说的行数和屏幕上数出来的行数才对得上。
            if (chk or {}).get("meta_kind") and not _pending(chk):
                continue
            card = str((er or {}).get("card") or "").strip() or "（无参数卡）"
            if card not in by:
                by[card] = []
                order.append(card)
            by[card].append((er, chk))
        bad_of = {}
        for card in order:
            bad_of[card] = sum(1 for _e, _c in by[card] if _pending(_c))
        # 自动收起（2026-11-03）：刚跑完一次比对时，**没有待处理行的组先收起来**。
        # 为什么要有这条：实测这份文件 130 张卡里有 112 张只有 1 行、且全部一致，
        # 全展开是 1458 行（55 屏）；按这个口径默认只留「有问题的」展开，
        # 一屏就能看见全部要管的东西，剩下 112 张卡的结论各占一行，扫一眼就知道
        # 「这些没问题」。人手动点过的展开/收起（_collapsed）在这一轮里一直有效，
        # 下次重新跑比对才会按新结果重算。
        #
        # 2026-11-04：有异常的组也不再整组摊开，而是**只留那几行待处理的**
        # （_trimmed，见 _apply_group）。用户原话：「对于有异常的那个表头的那一项，
        # 我觉得应该把那其中没有影响的也给收纳起来，就是只保留红色的露出来，
        # 这样子就会十分明显，显得效果就跟之前的那里的简洁是一样的。」
        # 也就是说跑完一次比对，屏幕上第一眼看到的就该是「要改的那几行」，
        # 而不是「打开了几张有问题的卡、里面混着几十行一致的」。
        if self._auto_fold:
            self._collapsed = {c for c in order if not bad_of[c]}
            self._trimmed = {c for c in order if bad_of[c]}
            self._auto_fold = False
        # 「几条待处理」这份账要留着给 _fold_all 用：收起动作**永远不碰有异常的组**
        # （见 _fold_all 的注释），得知道哪些组是有异常的。
        self._grp_bad = dict(bad_of)
        # 每张卡的「规格 + 设备」（2026-11-04 第六轮，用户原话：「每个号后面它对应
        # 的那个规格和设备号是不是应该也要展示出来？因为它们每一个都是不一样的
        # 嘛」）。实测同一张卡里的 58 行规格/设备完全一致（一张卡 = 一次设定），
        # 所以取第一份就够，不用合并。
        # 值来自上面那个 `seed` —— 它在**过滤元信息行之前**就收好了，否则这一行
        # 一被过滤，分组条上那串设备号也跟着没了（2026-11-11 踩到）。
        self._grp_meta = dict(seed)
        # 有异常的组**排到最前面**（2026-11-03 第四轮）。
        # 为什么非要排序：这份语料 130 张卡里只有 4 张有问题，其余 126 条灰色
        # 分组条按原顺序排在前面 —— 「只看异常」点下去，屏幕上一条红的都看不见，
        # 得往下翻几十屏才找得到。那等于这个按钮没起作用（第一版截图就是这样）。
        # 排序用「稳定排序」：有问题的一组互相之间、没问题的一组互相之间，
        # 都还是原来的卡号顺序，所以找某一号还是按顺序往下翻。
        # 注意这个顺序**与折叠状态无关** —— 否则点一下「全部展开」整张表就重排，
        # 人的视线会被打断。
        _idx = {c: i for i, c in enumerate(order)}
        order = sorted(order, key=lambda c: (0 if bad_of[c] else 1, _idx[c]))
        view = []
        for card in order:
            rows = by[card]
            view.append(("grp", (card, len(rows), bad_of[card],
                                 card in self._collapsed)))
            # 收起的组**照样把行铺出来**，铺完在 _render 末尾统一 detach。
            # 为什么不在这里跳过：折叠/展开走的是 detach / move，靠的是
            # self._grp_children 里那份「这一组有哪些行」的名单 —— 这里跳过了，
            # 名单就是空的，点开收起的那一组会变成「展开了一条也没有」（实测踩到）。
            for er, chk in rows:
                view.append(("row", er, chk))
        return view

    def _emit_group(self, g):
        """画一条分组条，并把「后面插进来的行归这一组」记下来（给折叠用）。

        版式（2026-11-03 用户原话：「把那个 000几给抽取出来，把前面一大当做一个
        大的表头」）：最左那格放「箭头 + 0001」，采集项那格放前缀大表头。
        前缀每张卡都印一次，但**一行只出现一次**（不是 486 行印 486 遍），
        眼睛找的是那个变化的 0001。
        """
        card, n, bad, folded = g
        iid = "grp::" + card
        pref, tail = _card_split(card)
        head = ("%s · %s" % (pref or card, tail)) if (pref and tail) else card
        self._grp_rows[iid] = card
        self._grp_children[card] = []
        # 这一组每一条子行「该不该留」—— 收起/展开时用不着，但「只看异常」
        # 要在**组内部**再筛一道，靠的就是这份表（见 _apply_group）。
        self._grp_keep[card] = {}
        self._grp_n[card] = n
        self._cur_grp = card
        # 有问题的组亮蓝色、没有的走灰色：130 条分组条全是同一个蓝色时，
        # 那几条真要管的会被自己人淹没（第一版截图就是这样）。
        #
        # 2026-11-04 第十轮：规格 · 设备不再单占一列，改成接在「58 行 · 显示 2 条」
        # 后面，两段拼成**一句话**写在 match 里 —— 用户原话：「把设备规格给列在
        # 那个核对结果的右侧……没有必要展示设备规格这个表头这个名字」。
        # 2026-11-11 又拆开：左半句（设备 · 规格）留在 match，右半句（N 行 · M 条）
        # 挪到 pad 那格右对齐 —— 拼一句时两样总要牺牲一样（见 _grp_line）。
        self.tree.insert("", "end", iid=iid,
                         values=(("▸ " if folded else "▾ ") + (tail or ""),
                                 "", pref or head,
                                 "", "", "", "", "", "", "",
                                 self._grp_line(card, self._match_budget_px()),
                                 self._grp_tail(card, self._pad_budget_px())),
                         tags=("grp" if bad else "grp2",))

    def _grp_tail(self, card, budget_px=0):
        """分组条**右端**那半句：「12 行 · 3 条待处理」/「12 行 · 显示 3 条」。

        三种状态说三句不同的话，因为这个数字必须和**屏幕上能数出来的行**对得上：
          · 全部一致        → 12 行 · 全部一致
          · 全展开          → 12 行 · 3 条待处理（这时 12 行都在下面）
          · 只显示待处理的  → 12 行 · 显示 3 条（下面真的只有 3 行）
        「12 行」永远是这张卡的大小（它回答了「这张卡一共几行」），后面那半句
        才是「现在让你看见几条」—— 两件事分开说，就不会觉得表头在骗人。

        它住在 `pad` 那一格（右缘那块留白，见 ui_layout.basic_cols），跟左边的
        「设备 · 规格」隔着 类型/上下限/限值/单位 几列 —— 左边是**这张卡是什么**
        （要拿去和 Excel 核的），右边是**这张卡有几行**（导航用的），本来就该
        一头一尾。放在同一格里挤不下的原因见 `_grp_line`。
        """
        n = (getattr(self, "_grp_n", None) or {}).get(card, 0)
        bad = (getattr(self, "_grp_bad", None) or {}).get(card, 0)
        if not n:
            # 这张卡的行全是「机器编号」那种元信息行，已经并到标题上了
            # （见 _group_view 的过滤）。硬写「0 行」像是在说这张卡是空的。
            return "全部一致"
        if not bad:
            full, short = "%d 行 · 全部一致" % n, "%d 行" % n
        elif card in (getattr(self, "_trimmed", None) or ()):
            full, short = "%d 行 · 显示 %d 条" % (n, bad), "%d 行" % n
        else:
            full, short = "%d 行 · %d 条待处理" % (n, bad), "%d 行" % n
        fnt = self._grp_font()
        if not budget_px or fnt is None:
            return full
        if fnt.measure(full) <= budget_px:
            return full
        # 右缘那格有多宽完全看窗口：窄了就只留「12 行」——「显示 3 条」虽然
        # 更要紧，但半句被切成「12 行 · 显示…」比只写「12 行」更容易看错。
        if fnt.measure(short) <= budget_px:
            return short
        return ""

    def _grp_line(self, card, budget_px=0):
        """分组条**左半句**：`设备 T-2101/T-1857 · ALL`（行数那半句见 `_grp_tail`）。

        2026-11-11 第一轮用户原话：「设备号应该也放到那个蓝色的表头里面去，
        因为设备号比较长，你放底下又塞不下，就直接单独把它拎出来放在上面，
        进行一种核验」。设备排在最前、带「设备」二字。

        第二轮把**降级梯子整个拆了**（用户原话）：
          「你看你有些是写设备 T-10194 等两台，是不是你这个你就不要等了，
            你还是全部展开，你这个位置挺宽的。写不下了再说。尽量多展示一些内容，
            在最上面这个标题栏这里。」
          「我就把右边的框给拉大了，拉大了之后我发现左边这个设备它居然不显示了
            ……它既然能左右拖动，你干嘛要把它压缩起来呢？就应该在标题栏那边
            全部显示出来。」

        所以这里**不再按列宽预压缩**：有几台写几台，一行写完。真放不下时由列
        自己裁（表格有横向拖动条），而不是提前把「等 2 台」「共 3 台」这种半截
        话印上去 —— 那种话既不能拿去和 Excel 核，窗口拉宽了也不会自己长回来，
        只能算「看起来有、实际没有」。列宽由 `_match_need_px` 按内容算，
        所以正常情况下这句话是完整显示的。

        `budget_px` 形参留着只为兼容调用点（`_emit_group` / `_refresh_grp_texts`
        / `_match_need_px` 都在传），值一概不用。
        """
        return self._grp_meta_text(card)

    def _grp_meta_text(self, card, budget_px=0):
        """分组条左半句：「设备 T-2101/T-1857 · ALL」。

        2026-11-04 第六轮，用户原话：「每个号后面它对应的那个规格和设备号是不是
        应该也要展示出来？因为它们每一个都是不一样的嘛……右边留的这么多空位就
        可以用上」。一张卡 = 一次设定，规格和设备串在卡内几十行里完全一致，
        所以在分组条上印**一次**就够，数据行那一格留空。

        两段各管一件事：
          · **设备** —— 唯一要人拿 Excel 对着核一眼的东西（用户说的「进行一种
            核验」）。**全量列出**，不缩写、不「等 N 台」。
          · **规格** —— 区分两张卡（φ2.0xQL15mm vs QL20mm），跟在设备后面。

        设备项只留设备号（`_slim_equip`）：Excel 里常写成
        `T-2068-DWG-0000007-03,04`，后面那串件号对「这台机器是不是这一台」
        没有任何帮助，却让一句话长三倍。
        """
        spec, equip = (getattr(self, "_grp_meta", None) or {}).get(card) or ("", "")
        # 设备串按过滤后的列表算：语料里有整串就是个「/」的噪音值，
        # 按原串判真假会拼出「设备 / · ALL」这种废话。
        eqs = [_slim_equip(p) for p in str(equip).split("/") if p.strip()]
        head = ("设备 %s" % "/".join(eqs)) if eqs else ""
        if head and spec:
            return "%s · %s" % (head, spec)
        return head or str(spec or "")

    def _grp_font(self):
        """分组条的字体（F_HEAD 加粗）—— 截断量宽必须按它算，按正文字号量会失真。"""
        f = getattr(self, "_grp_font_obj", None)
        if f is None:
            try:
                import tkinter.font as tkf
                f = tkf.Font(family=theme.FONT, size=theme.F_HEAD, weight="bold")
            except Exception:      # noqa: BLE001 —— 量不了就不截断
                f = False
            self._grp_font_obj = f
        return f or None

    def _match_budget_px(self):
        """分组条那句话画得下多少像素（match 列宽 − 行内左右呼吸）。"""
        try:
            return max(140, int(self.tree.column("match", "width")) - 12)
        except Exception:      # noqa: BLE001
            return 240

    def _pad_budget_px(self):
        """分组条右端「N 行 · M 条」画得下多少像素（pad 列宽 − 呼吸）。

        呼吸只留 8px：pad 是块**留白**列，不是内容列，右边多留白没有意义；
        多留 6px 就会把「58 行 · 显示 2 条」挤成只剩「58 行」——
        「显示几条」恰恰是唯一能说明「还有几行藏着」的信息。
        """
        try:
            return max(0, int(self.tree.column("pad", "width")) - 8)
        except Exception:      # noqa: BLE001
            return 0

    def _match_need_px(self):
        """「核对结果」这列**按内容**算需要多宽（结果缓存在 `_match_need` 上）。

        2026-11-11 用户原话：「把那个右边的什么类型啊、上下限什么的往左移一点，
        你看与左边那个判定结果中间隔了很大的位置」。
        起因：这一列原来是**窗口宽度的 34%（最少 300px）**，而判定词普遍只有
        四五个字，于是列里白出一两百像素，把 类型 / 上下限 一路推到窗口右边，
        看着像两截不相干的表。

        现在量两样，**都按最宽的那条给**：
          · 数据行的判定词（`取数不唯一(2 套)…` 这类长句量到 281px）。
            截断它等于把结论切掉半句。
          · 分组条上的「设备 · 规格」。2026-11-11 第二轮改成取 **max**：
            原来取中位数是为了「别让 9% 的卡给 91% 的卡白占 150px」，但代价是
            那 9% 的卡上设备号被硬切一半 —— 而设备号恰恰是唯一要人对着 Excel
            核一眼的东西，切掉它等于这条分组条白做。用户原话：「你还是全部展开，
            你这个位置挺宽的。写不下了再说。尽量多展示一些内容」。真要省地方，
            窗口拉窄时下面那道 `avail - tail_need` 会自然把它压回去，不需要提前砍。
        两者取大再加 18px 呼吸，夹在 200~460（460 是「一行分组条再长也不该
        超过半个屏幕」的经验上限；再长的卡靠横向拖动看）。
        """
        cache = getattr(self, "_match_need", 0)
        if cache:
            return cache
        try:
            import tkinter.font as tkf
            from tkinter import ttk as _ttk
            fname = _ttk.Style().lookup("Treeview", "font") or "TkDefaultFont"
            body = tkf.Font(font=fname)
        except Exception:      # noqa: BLE001 —— 量不了就退回默认宽度
            return 300
        head = self._grp_font() or body
        px_v, px_g = 0, []
        for card in list(getattr(self, "_grp_children", None) or ()):
            try:
                s = self._grp_meta_text(card, 0)   # 0 = 不降级，量「设备 · 规格」全量
            except Exception:      # noqa: BLE001
                continue
            if s:
                px_g.append(head.measure(s))
        try:
            for iid in self.tree.get_children(""):
                if iid.startswith("grp::"):
                    continue
                s = str(self.tree.set(iid, "match") or "")
                # 分隔行是「─」铺满一格，量它会得出几千像素，必须跳过
                if not s or not s.strip("─ ·"):
                    continue
                px_v = max(px_v, body.measure(s))
        except Exception:      # noqa: BLE001
            pass
        px = max(200, min(460, max(px_v, max(px_g) if px_g else 0) + 18))
        self._match_need = px
        return px

    def _fit_match_col(self):
        """定三列宽：match 按内容、unit 按内容、pad 吃掉右缘剩下的全部。

        · **match**（核对结果）：同时装判定词和分组条那句「设备 · 规格」，
          宽度按内容算（见 `_match_need_px`，200~320），**不按窗口比例**。
        · **unit**（单位）：固定一个「mW / mm」够用的宽度，不伸张 —— 让它去
          吃右缘会变成一个 200px 宽、中间只有一个「—」的列。
        · **pad**：空白占位列，唯一 stretch。它负责让总宽正好等于树宽，
          分组条那条色带才铺到最右边（用户第十轮：「颜色至少要延展到最右侧」），
          右边不留一条白。
        """
        try:
            cols = self.tree["displaycolumns"] or ()
            if isinstance(cols, str):
                cols = cols.split()
            if not cols or cols in (("#all",), ["#all"]):
                cols = list(self.tree["columns"])
            w = self.tree.winfo_width()
            if w <= 100:
                return
            others = sum(int(self.tree.column(c, "width"))
                         for c in cols if c not in ("match", "unit", "pad"))
            avail = w - others - 44 - 2
            # 先给右端那句「14 行 · 1 条待处理」留一块**刚好够**的地方，剩下的
            # 才归 match。留着不预留的后果实测过：设备号一展开（不再「等 2 台」）
            # match 就从 320 涨到 410，右端那格被压到 134px，正好差 6px 装不下
            # 「2 行 · 2 条待处理」，于是静默退化成光秃秃的「2 行」—— 看起来像
            # 「这组只显示 2 行」，其实是「还有 2 条待处理没说出口」。
            # 尾句是**定量**（每张卡就那几个字），设备号是**变量**，该让的是变量。
            tail_need = self._tail_need_px() + 6
            mw = min(self._match_need_px(), max(160, avail - tail_need))
            # 窗口太窄时别硬撑：先让出位置，宁可截判定词也别把后面几列挤没。
            mw = max(180, min(mw, w - others - 70))
            uw = 44
            self.tree.column("match", width=mw)
            self.tree.column("unit", width=uw)
            self.tree.column("pad", width=max(tail_need, avail - mw))
        except Exception:      # noqa: BLE001 —— 控件还没建好就保持默认宽度
            pass

    def _tail_need_px(self):
        """右端那句「N 行 · M 条待处理」最宽要多宽（`_fit_match_col` 拿它预留）。

        量的是**当前这一屏**（`_grp_bad` / `_trimmed` 已经按本轮结果算好），
        窗口变宽变窄、折叠状态变了都会重新量 —— 这是唯一一处「定宽依赖内容」
        的地方能接受的开销，因为要量的只有几张卡里的几个固定格式字符串。
        """
        cache = getattr(self, "_tail_need", 0)
        if cache:
            return cache
        fnt = self._grp_font()
        if fnt is None:
            return 120
        px = 0
        for card in list(getattr(self, "_grp_children", None) or ()):
            try:
                px = max(px, fnt.measure(self._grp_tail(card, 0)))
            except Exception:      # noqa: BLE001
                continue
        px = max(90, min(260, px + 8))
        self._tail_need = px
        return px

    def _refresh_grp_texts(self):
        """窗口变宽/变窄后，按新预算重铺各分组条那两半句（match + pad）。"""
        budget = self._match_budget_px()
        padb = self._pad_budget_px()
        for card in list(getattr(self, "_grp_children", None) or ()):
            try:
                iid = "grp::" + card
                self.tree.set(iid, "match", self._grp_line(card, budget))
                self.tree.set(iid, "pad", self._grp_tail(card, padb))
            except Exception:      # noqa: BLE001 —— 这一组已经不在表上了
                pass

    def _on_tree_resize(self):
        """树的尺寸变了：match 定宽跟上、灰框跟着走、分组条文案按新宽度重铺。"""
        self._fit_match_col()
        self._refresh_grp_texts()
        ring = getattr(self, "_sel_ring", None)
        if ring is not None:
            ring.refresh()

    def _add_kid(self, iid, keep=True):
        """把刚插进去的一行登记到当前组。

        `keep` 是「只看异常时这一条要不要露出来」：
          · 待处理的数据行 → True（连同它下面那行「↳ 应为」）
          · 一致/采集项/未填值的数据行 → False（连同它的「↳ 应为」）
          · 分隔线 → 一律 False（留着会在被藏掉的空档里画出一条断线）
        """
        card = getattr(self, "_cur_grp", None)
        if card is None:
            return
        self._grp_children.setdefault(card, []).append(iid)
        self._grp_keep.setdefault(card, {})[iid] = bool(keep)

    def _apply_group(self, card):
        """按这一组当前的状态，把它的子行摆成该有的样子。

        三个状态，其实是两把开关：
          · card in _collapsed → 一行都不显示（只剩分组条）
          · card in _trimmed   → 只显示「待处理」那几行
          · 两个都不在         → 全部显示

        为什么统一走 detach / move 而不是整表重画：486 行的表重画一次要几百
        毫秒，观感是「卡一下」；detach/move 是立即操作，是「唰」一下。
        也正因为是 move，**顺序**必须一条条摆回去（不能一批 move 完算数）——
        否则展开一组的行会跑到别组后面去。
        """
        iid = "grp::" + card
        kids = self._grp_children.get(card) or []
        keep = self._grp_keep.get(card) or {}
        folded = card in (self._collapsed or ())
        trimmed = (not folded) and (card in (self._trimmed or ()))
        try:
            idx = self.tree.index(iid) + 1
        except Exception:      # 表已经重画过、这条分组条不在了
            return
        for k in kids:
            show = (not folded) and (not trimmed or keep.get(k, True))
            try:
                if show:
                    self.tree.move(k, "", idx)
                    idx += 1
                else:
                    self.tree.detach(k)
            except Exception:
                pass
        # 箭头和那半句数字跟着状态走（收起时是 ▸，其余是 ▾）
        try:
            _p, _t = _card_split(card)
            self.tree.set(iid, "row", ("▸ " if folded else "▾ ") + (_t or ""))
            self.tree.set(iid, "match",
                          self._grp_line(card, self._match_budget_px()))
            self.tree.set(iid, "pad",
                          self._grp_tail(card, self._pad_budget_px()))
        except Exception:
            pass

    def _apply_all_groups(self):
        """把每一组都摆一遍（_render 铺完行之后调一次）。"""
        # 折叠 / 展开会换掉右端那句话（「N 条待处理」↔「显示 N 条」↔「全部一致」），
        # 三个字符串宽度不一样，所以这块缓存必须跟着作废，不能跨状态复用。
        self._tail_need = 0
        for card in list(self._grp_children):
            self._apply_group(card)
        self._ring_refresh()

    def _ring_refresh(self):
        """选中行的灰框跟上最新版面（行被藏/挪/重画之后必须调）。"""
        ring = getattr(self, "_sel_ring", None)
        if ring is not None:
            ring.refresh()

    def _toggle_group(self, card):
        """点分组条 → 收起 / 展开这一组。

        从「只看异常」留下的那几行点进来时，**展开就是全展开**：人点开一组，
        想看的就是这张卡的全部上下文（哪几行是一致的、一致的填的什么）。
        想看回那几行，点右上角那个「只看异常」即可。
        """
        self._set_group(card, card not in self._collapsed)

    def _set_group(self, card, folded):
        """把某一组设成收起（folded=True）或展开。已经是这个状态就不动。"""
        if folded:
            if card in self._collapsed:
                return
            self._collapsed.add(card)
            # 收起时把「裁过」这个状态也放下：再点开就是全展开（见 _toggle_group）。
            # 留着的话，分组条上那半句会一直写着「显示 1 条」而下面一行都没有。
            self._trimmed.discard(card)
        else:
            if card not in self._collapsed and card not in self._trimmed:
                return
            self._collapsed.discard(card)
            self._trimmed.discard(card)
        self._apply_group(card)
        self._ring_refresh()

    def _fold_all(self, fold=True):
        """一键：`fold=True` 是「只看异常」，`fold=False` 是「全部展开」。

        2026-11-03 第四轮，用户原话：「我想要就是收起的时候把那些异常项就不要给
        它收集起来。那这样子的话是不是就可以把那个概览的功能给平替掉，然后概览
        那个按钮就可以删掉。」
        2026-11-04 再进一步：有异常的组**也不整组摊开**，只留待处理那几行
        （用户原话见 _group_view 里那段）。于是点一下「只看异常」，屏幕上
        就真的是「一条一条要改的行」，没有一行是绿的。

        展开则没有例外：要全展开就全展开，人才看得到完整上下文。
        """
        if fold:
            self._collapsed, self._trimmed = set(), set()
            for card in list(self._grp_children):
                if (self._grp_bad or {}).get(card):
                    self._trimmed.add(card)
                else:
                    self._collapsed.add(card)
        else:
            self._collapsed, self._trimmed = set(), set()
        self._apply_all_groups()
        self._sync_fold_btn()

    def _toggle_fold_all(self):
        """工具栏上那一个按钮：现在有收着/裁着的就全展开，否则「只看异常」。"""
        self._fold_all(not (self._collapsed or self._trimmed))

    def _sync_fold_btn(self):
        """按钮上的字跟着状态走 —— 它永远说「点下去会发生什么」。

        只要还有组是收着或裁着的，右半屏就不是全貌，此时该给的是「全部展开」；
        全都摊开时该给的是「只看异常」。
        """
        btn = getattr(self, "btn_fold", None)
        if btn is None:
            return
        try:
            btn.configure(text="全部展开" if (self._collapsed or self._trimmed)
                          else "只看异常")
        except Exception:
            pass

    def _on_tree_motion(self, event):
        """鼠标划过列边界：把「左右双箭头」光标按回去（2026-11-03 第四轮）。

        用户原话：「既然不能调整了，那这个应该把那些代码给删掉，我现在鼠标放到
        这个地方，它居然中间还有左右移动键，虽然不能调整。」

        光标是 Tk 的类绑定（`ttk::treeview::Motion`）在发现 region=separator 时
        改的，不是我们的代码设的 —— widget 级绑定先于类绑定跑，在这里 return
        "break" 才能拦住它。只在 separator 上拦，别的地方照旧走 Tk 自己的逻辑
        （行高亮、hover 都靠它）。
        """
        try:
            if self.tree.identify_region(event.x, event.y) == "separator":
                if str(self.tree.cget("cursor") or "") != "arrow":
                    self.tree.configure(cursor="arrow")
                return "break"
            if str(self.tree.cget("cursor") or ""):
                self.tree.configure(cursor="")
        except Exception:      # noqa: BLE001 —— 控件正在销毁
            return

    def _on_tree_click(self, event):
        """明细表上的左键：① 拦住拖列宽；② 点分组条折叠/展开。

        列宽为什么一点都不能动（2026-11-03 第四轮把最后一处「能拉伸」也关了）：
        每一列的字段长度是**固定**的（「表4!21」「浮点」「是/否」「℃」），被拖窄
        只会剩「…」，拖宽是白占地方。采集项的宽度按内容算（`_fit_item_col`），
        也不是靠人拖。真正需要看长内容的时候，横向条还在，窗口拉大也在。
        """
        try:
            region = self.tree.identify_region(event.x, event.y)
        except Exception:
            return
        self._grp_pending = None        # 新的一次按下：上一次那条作废
        if region == "separator":
            return "break"              # 列边界不再是可拖的把手
        if region != "cell":
            return
        card = (self._grp_rows or {}).get(self.tree.identify_row(event.y))
        if card:
            # 2026-11-04 第九轮：折叠**改到松手那一刻**再做。
            # 原来是按下就折 —— 而这一轮表上要能按住左键拖（滑条撤掉了），
            # 「想在分组条上起手拖」会先被折一次，画面当场跳一格。
            # 现在只记下是哪张卡，交给 _on_tree_click_done 在「按了没动」时补。
            self._grp_pending = card
            return "break"              # 拦住 ttk：分组条不该被选中

    def _on_tree_click_done(self, event=None):
        """左键「按下 → 原地松开」之后收尾（见 `_PanDrag(on_click=...)`）。

        只做一件拖不动的事：折叠/展开被点中的那条分组条。真的拖动过就什么都不做
        —— 那一下是要挪画面，不是要折表。
        """
        card, self._grp_pending = getattr(self, "_grp_pending", None), None
        if not card:
            return
        self._toggle_group(card)
        self._sync_fold_btn()

    def _on_row_dbl(self, event=None):
        """双击明细某一行 → 打开「指定这一行在 PDF 的哪一格」。

        不做「只有找不到对应才能双击」的限制：人想给一条已经判对的行也钉个位置
        是合理的（换模板时更稳），多拦一次反而添堵。真正钉了没意义的是
        「采集项 / 未填值」这两类 —— 它们本来就不比数值，这里直接说清楚，
        不要让人点开一个窗、选完格子、最后什么也没发生。
        """
        if event is None or not getattr(self, "result", None):
            return
        # 双击落在列边界上：Tk 自带「双击自动适配列宽」，那正是这一版要杜绝的
        # （列宽全部固定死），所以直接吃掉。
        try:
            if self.tree.identify_region(event.x, event.y) == "separator":
                return "break"
        except Exception:
            pass
        iid = self.tree.identify_row(event.y)
        if str(iid).startswith("grp::"):
            # 双击分组条：它自己就是折叠开关，别往上找「前一行」——
            # 那会落到上一组最后一条数据上，稀里糊涂打开一个不相干的窗。
            return
        m = self._row2chk or {}
        # 双击落在「↳ 应为」那一行或分隔行上 → 往上找它的数据行
        _guard = 0
        while iid and iid not in m and _guard < 3:
            iid = self.tree.prev(iid)
            _guard += 1
        chk = m.get(iid)
        if not chk:
            return
        st = chk.get("status")
        if st in ("collection", "novalue"):
            messagebox.showinfo(
                "这一行不用指定位置",
                "「%s」这一行本来就不比数值（%s），钉了位置也不会产生判定。"
                % (chk.get("item") or "", "采集项" if st == "collection" else "未填值"))
            return
        self._pin_dialog(chk)



    def _render(self):
        r = self.result
        self._refresh_audit()
        self.tree.delete(*self.tree.get_children())
        # 分组的三份登记表跟着这表一起清空：上一轮的行 iid 已经随 delete 一起没了，
        # 留着只会让「点开某一组」去 move 一批不存在的行（异常被 try 吞掉，
        # 表现就是点了没反应）。
        self._grp_rows, self._grp_children = {}, {}
        self._grp_keep, self._grp_n = {}, {}
        self._grp_meta = {}
        self._cur_grp = None
        # 「核对结果」列该多宽要重新量（换了一份结果，内容全变了）
        self._match_need = 0
        # 右端尾句（「N 行 · M 条待处理」）的预留宽度同理
        self._tail_need = 0
        # 新一轮结果 → 重新按「有没有待处理行」决定哪些组默认收起（见 _group_view）
        self._auto_fold = True
        # 双击某一行 → 打开「指定这一行在 PDF 的哪一格」。
        # 这张映射表就是给它用的：Treeview 的行是「数据行 + 应为行」成对插入的，
        # 光看 iid 认不出背后是哪个 check，所以插入时顺手记一份（2026-11-02）。
        self._row2chk = {}

        excel_recs = self.excel_recs
        checks = r["checks"]
        # 判定结果必须**逐行**对上，绝不能用 (卡片, 采集项名) 做 key ——
        # 同一张卡里的同名采集项有几十行（同一项 × 5 台设备 × N 个规格），
        # 用 key 会互相覆盖，界面上所有同名行都显示**最后一条**的判定。
        # 实测 WI-0021 的 16650 行里有 13320 行被顶掉 ——
        # 不一致的行会被显示成「一致」，看着就像假一致。
        # compare 是逐条 append 的，所以下标天然一一对应，直接按顺序配对。
        # 「适用规格」门禁行是引擎自己插进来的（不属于任何一行 Excel），
        # 配对时要先把它摘出去，否则行数对不上、整片判定会掉进 cmap 兜底。
        plain = [c for c in checks if not c.get("spec_gate_row")]
        if len(plain) == len(excel_recs):
            pairs = list(zip(excel_recs, plain))
        else:
            cmap = {}
            for c in plain:
                cmap.setdefault((c["card"], c["item"], c.get("seq"),
                                 str(c.get("equip")), str(c.get("spec"))), c)
            pairs = [(er, cmap.get((er["card"], er["item"], er.get("seq"),
                                    str(er.get("equip")), str(er.get("spec")))))
                     for er in excel_recs]
        # 再按 checks 的**原始顺序**铺开：门禁行插在「机器编号」那一条之后
        # （引擎就是这么放的），左边「参数卡」留空 —— 它是一道门，不挂在某个采集项下。
        seq, pi = [], 0
        for c in checks:
            if c.get("spec_gate_row"):
                seq.append((c, c))          # 门禁行自带展示所需的全部字段
            elif pi < len(pairs):
                seq.append(pairs[pi])
                pi += 1
        # 分组（2026-11-03 第三轮，用户定「按参数卡分，分组条点击折叠」）：
        # 把 seq 铺成「分组条 + 数据行」的序列，收起的组只留分组条。
        # self._seq 留着 —— 折叠/展开时不必重算一遍配对，直接照它重画。
        self._seq = seq
        view = self._group_view(seq)
        for _i, _ent in enumerate(view):
            if _ent[0] == "grp":
                self._emit_group(_ent[1])
                continue
            er, chk = _ent[1], _ent[2]
            if chk is None:
                tag, match, should = "na", "—（采集项）", "—"
            elif chk["status"] == "ok":
                # 一致时同样把「应为」亮出来（绿色行）：只给一个「一致」没法确认
                # Excel 的值真的能在 PDF 里找到对应，假一致就藏在看不见里。
                # 说明文字由 scope_note 统一生成 —— 底层收窄到哪一档，这里就说哪一档。
                tag, match = "ok", scope_note(chk)
                if (chk.get("anchor") or {}).get("shade"):
                    # 这一行是**人钉过**的：钉的那一格把它从「不一致」改成了「一致」，
                    # 而按学到的规则本来会报不一致。橙色 + 一句话，不能让它混在
                    # 一片绿色里被划过去（假一致比假差异伤，见 wi_core 里那段注释）。
                    tag = "warn"
                    # 明细最后一列大约只放得下 20 个字就会截断（实测截图），
                    # 所以「盖住了一条差异」必须排在最前面 —— 那是这一行唯一
                    # 不能漏掉的信息，剩下的细节在「依据」里。
                    match = "一致（盖住了一条差异，请复核）"
                should = _fmt_expected_brief(chk)
            elif chk["status"] == "mismatch":
                if chk.get("meta_kind") and chk.get("scope") == "no_equip_match":
                    # 元信息项填了个 PDF 里查不到的机器号 —— 这不是「无对应」，是填错了
                    tag, match = "bad", "不一致(这行填的设备在 PDF 里查无此机)"
                elif chk.get("scope") == "prog_order":
                    tag, match = "bad", "不一致(与步骤顺序不符)"
                elif chk.get("scope") == "missing":
                    tag, match = "bad", "不一致(漏填：PDF 该步骤有值)"
                elif chk.get("scope") == "blank_filled":
                    tag, match = "bad", "不一致(PDF 该格为空值，不该填)"
                elif chk.get("scope") == "ref_filled":
                    tag, match = "bad", "不一致(PDF 是 ref 参考值，这一格不该填)"
                elif chk.get("scope") == "type_flag":
                    tag, match = "bad", "不一致(类型与是/否开关不符)"
                elif chk.get("precision_mismatch"):
                    # 数值相等、但小数位数不同（PDF 写 3、Excel 填 3.0）——
                    # 这是「写法精度不一致」，不是数值填错。必须单独措辞，
                    # 否则看的人会对着「应为 3」和「填了 3.0」以为数字没错、
                    # 不理解为什么判错。
                    tag, match = "bad", "不一致(小数位数与 PDF 不符)"
                elif chk.get("one_side_wrong"):
                    # PDF 明确写了不等号方向（如 `> 0.52` 是下限约束），
                    # Excel 却把值填在了另一侧的列 —— 比超限更根本的错，单独说。
                    tag, match = "bad", "不一致(填错了侧：PDF 是单边约束，值应填另一列)"
                elif chk.get("scope") == "one_sided":
                    # 单侧限值判成不一致，只可能是「PDF 的取值超出了 Excel 填的那一侧」——
                    # 光写「不一致」看不出是哪种不符，这里说清。
                    tag, match = "bad", "不一致(超出 Excel 所填的一侧限值)"
                elif chk.get("scope") == "spec_mismatch":
                    # 门禁行本身：适用规格对不上，是比任一参数值更靠前的错 ——
                    # 规格档都错了，下面填什么值都无从谈对错。
                    tag, match = "bad", "不一致(适用规格不在 PDF 任何一档内)"
                else:
                    tag, match = "bad", "不一致"
                should = _fmt_expected(chk)
            elif chk["status"] == "no_scope":
                tag, match = "warn", "无法判定(Excel 无设备号)"
                should = _fmt_by_equip(chk)
            elif chk["status"] == "multi":
                # 候选还剩两套以上 = 工具取数没能收窄到唯一。
                #
                # 2026-11-09 用户定的新口径：「多套就直接判错呗……它一个是给判对，
                # 又不标红，这样这个机器又学习不到。」
                # 所以这里从「黄色待确认」改成**红色按不一致计**：
                #   · 标红 → 一定会被看见，不会被「一致 251」那块绿色糊过去；
                #   · 计入不一致 → 汇总数字与实际待处理量对得上；
                #   · 措辞仍说清是「取数不唯一」而不是「你填错了」—— 标红是手段，
                #     不是给数据定罪。用户看到红字会去点这一格，点了就落锚点，
                #     下次同一张表直接按位置取，这条链才学得会。
                _n = chk.get("multi") or 0
                tag = "bad"
                match = ("取数不唯一(%d 套)，按不一致计 —— 请钉住这一格" % _n) if _n \
                    else "取数不唯一，按不一致计 —— 请钉住这一格"
                # 期望值那一列要**全部摆出来** —— 只显示第一套会让人以为
                # 那就是答案，反而比不显示更误导。
                should = _fmt_expected(chk)
            elif chk["status"] == "no_match":
                if chk.get("scope") == "name_unknown":
                    # 默认 5 列里就能看见的区别：不是「PDF 没有」，是「工具不认得这个叫法」。
                    tag, match = "warn", "无对应(参数名未收录)"
                elif chk.get("scope") == "no_step":
                    # 步骤式附表：这台设备那张表的表头只声明了 Step0/1/2，Excel 却要步骤3。
                    # 与「设备号查无此机」是两码事，混着说会让人去翻设备号。
                    tag, match = "warn", "无对应(PDF 该表里没有这一步)"
                elif chk.get("scope") == "spec_block":
                    # 门禁拦下来的：不是「PDF 里没有」，是**按规则不判**。
                    # 说成「无对应」会误导人去 PDF 里翻参数，其实要改的是适用规格那一格。
                    tag, match = "warn", "未判定(适用规格对不上，本组参数已跳过)"
                else:
                    # 「无对应」这个名字会让人以为 PDF 没读到。绝大多数情况是
                    # **读到、但匹配不上**（PDF 是标准矢量版式时尤其如此）；
                    # 真读不出的（扫描件）在整份结论那里单独拦一道，不混在这里。
                    # 2026-10-30：措辞与概览抬头、状态栏统一成「找不到对应」——
                    # 同屏两个词指同一批行，看的人得自己换算。顺带把括号里那句
                    # 「已读到 PDF，只是没对上」也改掉：实测有 PDF 里确实没有那一项
                    # 的情况（WI-0039 正文说「如下表7」却全篇没有表7），
                    # 「只是没对上」等于替文件开脱。
                    # 「几张表都像，需人工定」只在**确实存在多张表候选**的时候说。
                    # 走到「参数名认得、但这份 PDF 的参数表里没有它」那一档
                    # （引擎里 `_vertical_expected` 返回 None，scope 因此为空）时，
                    # `fb_tables` 只是兜底层顺手挂上的近似候选 —— 挑哪张都没用，
                    # 那是让人去挑一张本来就不含这个参数的表。
                    # 2026-10-30：WI-0039 表4 那 8 条「切定长长度」就在这一档，
                    # 原来一律写「几张表都像，需人工定」，等于把「表里没有」
                    # 说成「你选一张」。依据列本来就说清了真正的三种可能。
                    if chk.get("fb_tables") and chk.get("scope"):
                        tag, match = "warn", "找不到对应(几张表都像，需人工定)"
                    elif (getattr(self, "pdf_selfcheck", None) or {}).get("image_pages"):
                        # 这份 PDF 有图片块 → 文字层里找不到，不代表 PDF 上没有。
                        # WI-0039 表7 就是整块图片，印着 35/60/100/140/170/220/270/330mm，
                        # 和 Excel 八行全对得上，只是工具读不到（2026-11-02 用户自己发现的）。
                        # 说「PDF 里没有这一项」是错的，必须把「去看图」这条路指出来。
                        tag, match = "warn", "找不到对应(可能印在图片里)"
                    else:
                        tag, match = "warn", "找不到对应(PDF 里没有这一项)"
                should = _fmt_expected(chk) if chk.get("meta_kind") else "？"
            elif chk["status"] == "collection":
                tag, match, should = "na", "—（采集项，不核对）", "—"
            elif chk["status"] == "novalue":
                if chk.get("unreadable"):
                    # 值是**填了的**，只是填的不是数值（「定位标记模板：T-2046-56」
                    # 这类代码/文本），工具读不出可比较的数，所以没参与核对。
                    # 2026-10-29 用户点名：「表格明明填着，为什么说未填值」——
                    # 一栏九十几行全写「未填值」，等于告诉人「这表是空的」，
                    # 而实际上是我们没核。措辞必须把话说准。
                    tag, match, should = "na", "—（填的是文字/代码，未参与核对）", "—"
                else:
                    tag, match, should = "na", "—（未填值）", "—"
            else:
                tag, match, should = "na", "—（采集项）", "—"
            # 现值取 chk（它是 er 的副本，另外还带 meta_kind 等判定标记）——
            # 「机器编号」这类元信息项的值就是该行设备名，要能回退到 equip 才能显示出来。
            val = _value_str(chk if chk else er)
            # 元信息行（机器编号那几条）的「现值」其实是**整串设备号**——
            # 实测用户那份表是 25 台用「/」串起来、184 个字符，塞进 118px 的限值列
            # 里连前三个都看不全（用户原话：「它如果是设备号，它占的位置就太多了」）。
            # 这里压成「前 3 台… 共 25 台」：一眼能认出是哪批设备，要看全的
            # 到「依据」列或导出里看，那里有全串。
            # 判据不只认 meta_kind：任何「≥4 段用 / 串起来」的值都是这种列表形状，
            # 一份表里 meta_kind 可能没标上，但形状骗不了人。
            _meta = bool((chk or er or {}).get("meta_kind"))
            if _meta or str(val).count("/") >= 3:
                val = _short_list(val)
            # 元信息行（机器编号）的「↳ 应为」**不画**（2026-11-11 用户原话：
            # 「设备号应该也放到那个蓝色的表头里面去……那底下那一行是不是就可以
            #   不用核对了？根本不需要有核对这一项，省去一个步骤」）。
            # 判据只有两条：① 这是元信息行（值就是设备号）；② 判成一致。
            #
            # 上一轮还多要了一个「现值字面等于应为」，那是过度保守 —— 巡检在成功1
            # 上抓到了后果：Excel 写 `T-2068-件号/T-1254-件号/…`，PDF 写
            # `T-1254、T-1534、…共5套之一`，**同一批设备的两种写法**，引擎判 ok，
            # 但因为字面不等，照样在下面画了一行 `↳ 应为 T-1254、T-15…`。
            # 那不是「最该看的一眼」，那是把已经核对完的事再摊一遍 —— 正是用户
            # 说要省掉的那一步。真要不一致时 status 就不是 ok 了，两行照旧画。
            _skip_sub = (bool(_meta) and (chk or {}).get("status") == "ok")
            basis = (chk.get("basis") if chk else None) or "—"
            # 匹配依据挂在「依据」列最前面（规范第八条）：匹配层级 / 状态 / 分数 /
            # 对应表。这样既能在默认 5 列里翻到（该列可展开），也不新增列破坏版式。
            m = (chk.get("match") if chk else None) or {}
            if m and m.get("state"):
                head = "[匹配 L%d %s" % (m.get("layer") or 0, m["state"])
                if m.get("score"):
                    head += " %.0f%%" % (m["score"] * 100)
                if m.get("table"):
                    head += "｜" + m["table"]
                head += "] "
                basis = head + basis
            # ---- 上行：Excel 现在填的（含表里自带的两个「是/否」开关）----
            # 这一条在「只看异常」里容不容易留下来：只看它有没有事要管，
            # 与颜色无关（橙色那几种「找不到对应」也得留着，那是要去核的）。
            _keep = _pending(chk)
            _iid = self.tree.insert("", "end", values=(
                _row_label(er), er["card"], er["item"],
                er["type"] or "—", _yn_cn(er.get("show_lim")),
                _yn_cn(er.get("allow_over")), val,
                er["unit"] or "—", er["proc"] or "—", basis,
                match, ""), tags=(tag,))
            if chk is not None:
                self._row2chk[_iid] = chk
            self._add_kid(_iid, _keep)
            # ---- 下行：应为值（行号 / 参数卡留空，值落在「限值」正下方）----
            # 两个「是/否」列这里放的是**按类型规矩应该填的**（引擎算好带过来），
            # 与上行的实际值上下对齐：值错、是/否错、类型错，三种都能直接看出来。
            # 「类型」列只在规矩**对类型本身有要求**时才写（球囊那条：PDF 带 > / <
            # → 应为「字符串」）；规矩没表态就留空，别拿「—」冒充「不用改」。
            _wt = (chk.get("want_type") if chk else "") or ""
            _sub = None
            if not _skip_sub:
                # 「↳ 应为」那一格同样要压：它是**另一条路**进来的（`should` 来自
                # 判定结果，没经过上行那个 `_short_list`）。2026-11-11 巡检在成功1
                # 上发现：采集项「机器编号」的应为值也是整串设备名，592px 塞进
                # 74px 的限值列，只露出 `T-2068-DWG20…`。判据与上行一致 ——
                # 只有「≥4 段用 / 串起来」的列表形状才压，普通值一个字不动。
                _should_show = should or "—"
                if str(_should_show).count("/") >= 3:
                    _should_show = _short_list(_should_show)
                _sub = self.tree.insert("", "end", values=(
                    "", "", "↳ 应为", ("↳ " + _wt) if _wt else "",
                    _flag_cn(chk.get("want_show")) if chk else "—",
                    _flag_cn(chk.get("want_over")) if chk else "—",
                    _should_show,
                    "", "", "", "", ""), tags=(_SUB_TAG.get(tag, "na2"),))
                # 「↳ 应为」跟着它上面那条走：上面留在屏上，它才留 ——
                # 只留「应为」而没有「现值」，等于给了一句话却不知道在说哪一行。
                self._add_kid(_sub, _keep)
            # ---- 分隔行：两条数据之间拉一条横线，把每条数据「框」起来 ----
            # ttk Treeview 没有行边框（macOS 的 Aqua 主题还会忽略 tag 的背景色），
            # 所以用一条由「─」组成的空行来当分隔线 —— 纯文本，两个平台都画得出来。
            # 只在「下一条还是同一组里的数据行」时才画：跨到下一组的分组条上，
            # 分组条本身就是一条视觉分断，再加一条「─」等于两道线叠在一起。
            if (self.show_sep.get() and _i + 1 < len(view)
                    and view[_i + 1][0] == "row"):
                _s = self.tree.insert("", "end",
                                      values=(_SEP,) * len(self.all_cols),
                                      tags=("sep",))
                # 分隔线在「只看异常」里一律不露：它夹在两条数据之间，而这两条
                # 里往往正好有一条被藏掉了 —— 留着就是一条凭空出现的断线。
                self._add_kid(_s, False)

        # 行铺完了，再把每一组摆成它该有的样子（此刻名字才齐）
        self._apply_all_groups()

        # ---------------- 摘要区排版 ----------------
        # 列位置由 Text 的 tab stop 固定（见 _build_ui 里的 tabs=…），不靠空格凑 ——
        # 中英文宽度不同，用空格拼出来的「表」在真实界面里永远是歪的。
        # 每段文字带一个标签名（t/h/ok/bad/warn/dim/note），只影响样式，不改内容。
        c = r["counts"]
        src = self.schema.chosen.label if self.schema and self.schema.chosen else "未识别"
        # 只留「五种来源各多少」这一组数：它还用在下面对「全都没读到」的兜底判断。
        # 原来还有 n_blank_pdf / n_manual / n_name 三个，只服务于【PDF 参数来源】
        # 那一段，随那一段一起删了（2026-10-30）。
        n_step = len(getattr(self, "pdf_steps", []) or [])
        n_prog = len(getattr(self, "pdf_programs", []) or [])
        n_vert = len(getattr(self, "pdf_vertical", []) or [])
        mic = getattr(self, "pdf_micrometer", None) or {}
        n_mic = sum(len(v) for v in mic.values() if isinstance(v, (list, set)))
        n_meta_items = len(getattr(self, "pdf_meta", []) or [])

        B = []                                  # [(行文本, 标签或 None)]

        def add(text="", tag=None):
            B.append((text, tag))

        def sec(title):
            add()
            add("【%s】" % title, "h")

        def row(label, value="", note="", indent=2, tag=None):
            """一行三列：条目 / 数值 / 备注 —— 列位靠 tab stop，不靠空格。"""
            s = " " * indent + str(label)
            if note:
                s += "\t" + str(value) + "\t" + str(note)
            elif str(value) != "":
                s += "\t" + str(value)
            add(s, tag)

        def note(text, tag="note"):
            add("      · " + text, tag)

        judged = [v for v in r["checks"] if v["status"] in ("ok", "mismatch")]
        b = scope_buckets(judged)
        hard = sum(b.get(k, 0) for k in HARD_SCOPES)
        bad = sum(b.get(k, 0) for k in BAD_SCOPES)
        soft = len(judged) - hard - bad

        # 结论先行：抬头一行交代「比对的是哪两边」，第二行是人话结论。
        # 抬头说清**这是什么**，而不是「跑完了」。
        # 原来写的是「WI-0039 · 比对完成」—— 「比对完成」是状态、不是身份，而且
        # 和底下那句人话结论打架：抬头说「完成」，下面说「9 条对不上」，看着像
        # 自相矛盾（2026-10-30 用户原话：「那个对比完成那个地方，它也比较含糊」）。
        # 「完成」这件事状态栏已经在说了，抬头这一行只用来交代比对的两边是谁。
        # （更早那版还有「来源/采集项/时间」+「产品线」两行独立抬头，信息量很低
        # 却占掉摘要区三分之一，早已合并进下面那行灰色的 meta。）
        _n_gap = sum(1 for v in r["checks"] if v.get("scope") == "no_spec_match")
        add("%s ｜ Excel 参数卡 对照 PDF 作业指导书" % self.wi_name, "t")
        add("  " + _plain_verdict(c, _n_gap, soft, bad),
            "bad" if (c["mismatch"] or c.get("multi")) else
            ("warn" if (c["no_match"] or c.get("no_scope")) else "ok"))
        # 未通过率告警紧跟结论：它是「这份结果能不能照着改」的前提，
        # 比下面的来源、meta 都重要。错一两条时不出现（正常用法，逐条改就行）。
        _hint = _unpassed_hint(c)
        if _hint:
            add("  " + _hint, "warn")
        _meta = ["Excel 来源 %s" % src, "采集项 %d 条" % len(excel_recs)]
        if self.family:
            _meta.append("产品线 " + "、".join(self.family))
        _meta.append(datetime.now().strftime("%H:%M:%S"))
        add("  " + " ｜ ".join(_meta), "dim")
        # —— 行号口径（2026-10-30）——
        # 这份文件里有 17 张表，参数明细在两张粒度完全不同的表里都「像回事」：
        #   「表4」280 行 = 参数卡模板，一张卡的采集项定义就这几行；
        #   「rev06」14492 行 = 同一批采集项按 设备 × 产品 展开后的实际数据。
        # 引擎必须读后者（只有它带设备号/适用规格），可行号一报出去，人拿它去
        # 表4 里翻，翻出来是另一条 —— 用户原话「行号与实际 excel 中表4 的内容
        # 好像对应不上」。所以行号一律写成「表名!行」，并在摘要里点明这一点。
        if xl_locate is not None and src and src.strip() != "表4":
            try:
                _t4n = xl_locate.find_sheet(self.xlsx_var.get())
            except Exception:
                _t4n = None
            if _t4n:
                add("  行号是「%s」的（%s 是参数卡模板，行号不对应）" % (src, _t4n), "dim")
        # ——「读到了多少」与「匹配上了多少」分开报 ——
        # 界面上那些「无法匹配」常被误读成「PDF 没读出来」。PDF 是标准矢量版式
        # 时，绝大多数其实是**读到、但没对上**。这里把底账摆出来：PDF 抽出了
        # 几张表、几个参数点，以及分层回退各自救回多少条。
        _rd = getattr(self, "read_info", None) or {}
        if _rd:
            _bits = ["PDF 读出 %d 张表 / %d 个参数点" % (_rd.get("tables", 0),
                                                    _rd.get("points", 0))]
            _fb = _rd.get("fb") or {}
            saved = sum(_fb.get("l%d" % i, 0) for i in (1, 2, 3, 4))
            if saved:
                _bits.append("分层回退救回 %d 条（关键词%d/设备号%d/规格%d/值集近似%d）"
                             % (saved, _fb.get("l1", 0), _fb.get("l2", 0),
                                _fb.get("l3", 0), _fb.get("l4", 0)))
            if _fb.get("amb"):
                _bits.append("%d 条候选表咬得太近、已交人工定" % _fb["amb"])
            # 锚点通道（11-02）：这一栏不写出来，人就看不出「这几条是我自己钉的位置，
            # 不是工具认出来的」—— 那种一致性最硬，但也最该知道它从哪来。
            _an = _rd.get("anchor") or {}
            if _an.get("hit"):
                _bits.append("按你钉的位置读值 %d 条（一致%d/不一致%d）"
                             % (_an["hit"], _an.get("ok", 0), _an.get("diff", 0)))
            if _an.get("stale"):
                _bits.append("%d 条锚点这份文件里认不出那张表、未生效" % _an["stale"])
            if _rd.get("fuzzy"):
                # 这里原来是「识别错误率 100% > 30%」—— 词用错了。
                # `rate` 在引擎里是 failed0 / judged0，即**没有匹配上的比例**，
                # 跟「识别」半点关系都没有（WI-0039 表4 九条全对不上 → 100%，
                # 而那份 PDF 的 14 张表 / 99 个参数点读得好好的）。
                # 一句话同时说错两件事：既冤枉了识别层，又让人以为文件坏了。
                # 2026-10-30 改成「对不上率」，并点明这个数只是第4层的触发门槛。
                _bits.append("对不上率 %.0f%% 超过 %d%%，已额外启用第4层近似匹配"
                             % ((_rd.get("rate") or 0) * 100, 30))
            add("  " + " ｜ ".join(_bits), "dim")
            if c["no_match"]:
                _imgp = (getattr(self, "pdf_selfcheck", None) or {}).get("image_pages") or []
                if _imgp:
                    # 有图片块的时候，那句话就不能只说「是这一项没对上」了 ——
                    # 用户 2026-11-02 的原话：「这个部分看着像是图片类型的PDF啊，
                    # 所以才会读取不到吧」。他说得对，工具得跟上。
                    add('      「找不到对应」%d 条：PDF 的文字层里确实没有，'
                        '但第%s页印着图片块（图里的字没有文字层，工具读不到）—— '
                        '先对照那几页看图，别当成 PDF 里没有'
                        % (c["no_match"], "、".join(str(p) for p in _imgp[:4])), "warn")
                else:
                    add('      「找不到对应」%d 条 = PDF 的文字层读得出，是这一项没对上'
                        '（真读不出的只有扫描件，那种整份就被拦下了）' % c["no_match"],
                        "dim")
        if getattr(self, "code_warning", ""):
            add("  " + self.code_warning, "warn")

        # 缺口 A：抽取判词顶在最上面。它回答的是「下面这张表值不值得信」，
        # 比一致/不一致那几个数字更该先看到。原来只压在报告中段的【抽取自检】
        # 段里 —— 而人只会看最上面一行，等于没有。
        sc_top = getattr(self, "pdf_selfcheck", None) or {}
        if sc_top.get("lines"):
            _sus = bool(sc_top.get("suspect"))
            add("  " + sc_top["lines"][0], "warn" if _sus else "ok")
            # 缺口 C（诚实版）：把「漏抽」和「无对应」显式连起来。
            # 做不成「逐条 no_match 打标」—— 实测引擎里 6 个 no_match 分点全都
            # 不保留「这条查的是哪一页/哪张子表」的血缘（1/1 条既无 sub 也无
            # page），要逐条标就得改 5 个抽取器的返回契约 + compare 全链路，
            # 那是引擎改造、不是补一个交叉引用；而且当前语料 0 疑点，改完
            # 在真实数据上根本验证不了。所以这里只做**能说清事实**的那一层：
            # 点名可疑的具体位置，说清它对「无对应」意味着什么。
            if _sus and c["no_match"]:
                _where = []
                if sc_top.get("hollow_subs"):
                    _where.append("子表 " + "、".join(sc_top["hollow_subs"][:4]))
                if sc_top.get("gap_subs"):
                    _where.append("断号子表 " + "、".join(sc_top["gap_subs"][:4]))
                if sc_top.get("missed_pages"):
                    _where.append("整页漏读 p" + "、p".join(
                        str(p) for p in sc_top["missed_pages"][:6]))
                # 判词那一句已经交代过「疑点会混进『无对应』」，这里只补它没说的
                # 两件事：疑点具体在哪、以及它跟本次「无对应」条数的关系 ——
                # 换了段精简之后，再把同一句话说第二遍就是纯占地方了。
                _tail = ("本次「无对应」%d 条里若夹着这几处的项，那是没读出来，"
                         "不是 PDF 里真没有。" % c["no_match"])
                if _where:
                    add("      可疑位置：%s —— %s" % (" ｜ ".join(_where), _tail), "warn")
                else:
                    add("      " + _tail, "warn")

        d = getattr(self, "snapshot_diff", None)
        if d is not None:
            # 基准时间并进同一行（原来单独占一行「对比基准：上一次比对快照」，
            # 是句废话；真正想知道的是「跟哪个时间点的比」，那就把时间写出来）。
            # 时间戳的真格式是 save_snapshot 里的 %Y%m%d_%H%M%S（15 位，如
            # 20260922_145652）—— 按 %Y-%m-%d %H:%M:%S 去左/右切片会切出
            # 「3_145652」这种乱码。所以先验位数与下划线位置，对不上就不显示，
            # 绝不让一个错的时间戳出现在结论区。
            _pt = getattr(self, "snapshot_prev_time", "") or ""
            if len(_pt) == 15 and _pt[8] == "_" and _pt.replace("_", "").isdigit():
                _pt = "（基准 %s-%s %s:%s）" % (_pt[4:6], _pt[6:8], _pt[9:11], _pt[11:13])
            else:
                _pt = ""
            if d.get("total"):
                # 只留一行汇总。原来再挂 3 行样例（2 条新引入 + 1 条已修正），
                # 那是给改引擎的人看的；交给员工的那版只要「跟上次比有没有变坏」
                # 这一句就够了 —— 具体哪几条在导出报告里查。
                add("  与上次比对%s：修好 %d ｜ 新引入 %d ｜ 其它 %d ｜ 新增 %d ｜ 消失 %d"
                    % (_pt, d["fixed"], d["broke"], d["changed"], d["added"], d["gone"]),
                    "bad" if d["broke"] else "note")
            else:
                add("  与上次比对%s：判定结果没有变化。" % _pt, "note")

        # 【核对结果】整段已删：一致/不一致/无对应 三个数已在上面那句人话里，
        # 而且「占已判定 %」这种比例没人会拿它做决定 —— 重复一遍只是把结论摊薄。
        # 【判定强度】整段已删，只留下「宽松判定」这一句（只在真有时才出现）：
        #   ① 硬判定/宽松判定两个数本身就重复（一个 = 总数减另一个）；
        #   ② 「逐档（只列非零）」列的是底层 scope 名，是引擎内部口径 ——
        #      明细表每一行的「依据」列已经写明了这一行为什么这么判，还能排序筛选，
        #      在窗口里再铺一张十几行的档位表，等于让人读两遍同一件事。
        #      逐档明细在**导出报告**里仍然完整保留（那里是留档，越全越好）。
        # 放在分隔线之前：它是「结论」的一部分，不是下面某一段的开场白。
        if soft:
            add("  其中 %d 条是「候选不止一个、命中了其中之一」的宽松判定 ——"
                " 已在下面【要人工确认的】里逐条列出。" % soft, "note")
        add("─" * 56, "rule")

        # ---------------- 要处理的那几条 ----------------
        # 「错在哪一条」是人打开窗口后第一件想确认的事，所以紧跟结论；
        # 来源 / 抽取自检 / 要人工确认的 都是补充信息，排在它后面。
        bads = [v for v in r["checks"]
                if v["status"] in ("mismatch", "no_match", "no_scope")]

        def flag_of(v):
            """一句话说清这条错在哪 —— 判定列只有一列，必须最短且自解释。"""
            if v["status"] == "no_match":
                if v.get("scope") == "name_unknown":
                    # 「工具不认识这个叫法」和「PDF 里真没有」是两件完全不同的事：
                    # 前者要在词表里补一条，后者要去改文档。混成一句「无对应」，
                    # 看的人会拿着它去改本来没错的文档（WI-0039 的「速率」实例）。
                    return "名字未收录"
                # 与明细表「核对结果」列、概览抬头、状态栏用同一个词（2026-10-30）。
                return "找不到对应"
            if v["status"] == "no_scope":
                return "缺设备线索"
            sc = v.get("scope")
            if sc == "missing":
                return "漏填"
            if sc == "blank_filled":
                return "多填"
            if sc == "ref_filled":
                return "多填(ref)"
            if sc == "type_flag":
                return "类型/开关不符"
            if sc == "one_sided" and v.get("one_side_wrong"):
                # PDF 写明了是哪一侧的约束（`> 0.52` 是下限），Excel 填到另一侧去了。
                return "填错侧"
            if sc == "one_sided":
                # 会走到这里的单侧行只有一个原因：PDF 的取值超出了 Excel 填的那一侧。
                return "超出所填一侧"
            if sc == "prog_order":
                return "步骤顺序不符"
            if v.get("meta_kind") and sc == "no_equip_match":
                return "设备号查无此机"
            return "不一致"

        def _with_unit(text, unit):
            """「—」和「？」不是值，别在它们后面缀单位。"""
            return text if text in ("—", "？", "") else text + (unit or "")

        def _same_tilde(text):
            """现值说「720 ~ 760」、应为说「740~780」—— 并排看很别扭，统一成一种。"""
            return re.sub(r"\s*~\s*", " ~ ", text)

        if bads:
            sec("要处理的（共 %d 条）" % len(bads))
            add("      行号\t类型\t显示上下限\t允许超\t现值\t判定", ("cols", "dim"))
            marks = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳"
            # **全部列出，不截断**（2026-09-29 用户要求）：摘要区是明细表的压缩显示，
            # 截断等于「上面说有 63 条、只给看 40 条，剩下 23 条自己下去翻几千行」——
            # 那正是这块区域要替人省掉的事。长就长，概览页整页都是它，自己的滚动条管够
            # （2026-10-29 起概览独占一页，不再和明细表抢高度）。
            _basis_done = {}          # {(卡片, 依据): 已打过} —— 同一张卡同一条依据只说一遍
            for i, v in enumerate(bads):
                mk = marks[i] if i < len(marks) else "%d)" % (i + 1)
                u = v.get("unit") or ""
                flag = flag_of(v)
                # 每条**两行**（与下方明细表统一）：
                #   上行 = 行号 / 类型 / 两个是-否 / 现值 / 判定
                #   下行 = 应为值
                # 位置列改成只写行号：原来写「卡片 · 采集项」一串，占掉半行还看不出
                # 该去改哪一行（2026-09-29 用户要求）。
                add([("  %s %s\t%s\t%s\t%s\t%s\t" % (
                        mk, _row_num(v),
                        _fit(v.get("type") or "—", 6),
                        _yn_cn(v.get("show_lim")), _yn_cn(v.get("allow_over")),
                        _fit(_with_unit(_value_str(v), u), 14)),
                      "cols"),
                     (flag, ("cols", "flag"))])
                # 下行**每一格都写**（不是只写一个应为值）：类型错、开关填反、值漏填
                # 是三种不同的改法，只给一个值的话人不知道该动哪几列。
                # 与上行**不一样的那几格用红色**，一样的保持灰色 —— 一眼看出
                # 「这条要改的只有这几格」（2026-09-29 用户：上下显示内容要统一）。
                #
                # 两条不标红的边界（宁可让人自己看，也不凭空造错）：
                #   · 表里那格空着 → 不标（没填 ≠ 填错）
                #   · 规矩没表态（应为 None）→ 不标
                _yn_show = _yn_cn(v.get("show_lim"))
                _yn_over = _yn_cn(v.get("allow_over"))
                _wt = v.get("want_type") or ""
                _want = [
                    (_wt or "—", _wt != "" and _wt != (v.get("type") or "")),
                    (_flag_cn(v.get("want_show")),
                     v.get("want_show") is not None and _yn_show != "—"
                     and _flag_cn(v.get("want_show")) != _yn_show),
                    (_flag_cn(v.get("want_over")),
                     v.get("want_over") is not None and _yn_over != "—"
                     and _flag_cn(v.get("want_over")) != _yn_over),
                ]
                # 应为值那一格同理：判定已经是「不一致」、且现值跟应为对不上时才标红。
                # 只在**判定层已经认定这是错**时才比 —— 「无法匹配」那些应为值是空的
                # （工具不知道答案），标红等于说「应该删掉」，那是凭空造错。
                _exp_val = _same_tilde(_with_unit(_fmt_expected_brief(v), u))
                _val_diff = (v.get("status") == "mismatch"
                             and _with_unit(_value_str(v), u) != _exp_val)
                _segs = [("      ↳ 应为\t", ("cols", "dim"))]
                for _txt, _diff in _want:
                    _segs.append((_txt + "\t", ("cols", "flag" if _diff else "dim")))
                _segs.append((_exp_val + "\t", ("cols", "flag" if _val_diff else "dim")))
                add(_segs)
                # 依据只在「光看判定列看不出原因」时补一行：
                # 漏填/多填/不一致 本身已经说明了问题，逐条再打一遍依据
                # 会把摘要撑成几十行，反而把真正要看的埋掉。
                # 「类型/开关不符」「多填(ref)」也补一句：判定列只说「不对」，
                # 不说**为什么**（球囊那条要看 PDF 原文才知道是 `>90℃` 触发的）。
                if flag in ("无对应", "缺设备线索", "步骤顺序不符", "设备号查无此机",
                            "类型/开关不符", "多填(ref)") and v.get("basis"):
                    # 同一条依据在一张卡里往往管十几行（如「规格对不上任何一档」）：
                    # 行还是要一行行列（不能归并，否则等于把别的行藏起来），
                    # 但那句解释没必要原样重复十几遍 —— 重复的只说「同上」。
                    bk = (v["card"], v["basis"])
                    if _basis_done.get(bk):
                        add("      ↳ 依据同上", "dim")
                    else:
                        _basis_done[bk] = True
                        add("      ↳ %s" % v["basis"], "note")
                hint = (self.sibling_hints or {}).get((v["card"], v["item"]))
                if hint:
                    txt = "；".join("%s: %s" % (wi, "、".join(vs2))
                                    for wi, vs2 in sorted(hint.items()))
                    add("      ↳ 本 PDF 里没有该参数；同名采集项还写在其它 WI 的参数卡下：%s"
                        % txt, "warn")
                    add("        （该值很可能出自那份 WI 的作业指导书 —— "
                        "请换对应 WI 的 PDF 再核对这一条）", "warn")
                # 每条之间一条浅灰横线 —— 与下方明细表同一个手法（纯文本画线，
                # 两个平台都画得出来）。一条占两行（现值 / 应为），不隔开的话
                # 上下两条会糊成一片 —— 值接近时（1850 与 1900）根本分不清
                # 哪一行属于哪一条。最后一条后面不画。
                if i + 1 < len(bads):
                    add("─" * 56, "rule")
            # 「按判定类型给一句怎么改」曾经做过，2026-09-23 用户看过实际界面后
            # 判定为没必要，已整段删除（连同 app 里的 _FIX_HINT 表）。
            # 想恢复的话：判定词 → 动作句的映射要放在**表尾**、同类只说一遍，
            # 否则会被逐条重复撑成两倍行数。别恢复成逐条重复。

        # ---- 说明段瘦身 ----------------------------------------------------
        # 原来 11 条长段落，一半在讲「引擎是怎么做到的」（附表级怎么定位、直径档
        # 怎么推算、表头元信息取哪几行）—— 那是实现说明，不是这份表的结果，
        # 却占掉摘要区好几屏。现在只留三种「判了一致、但你别全信」的告警，
        # 各自压成一句、且只在真出现时才说；口径说明只留在导出报告里。
        # （曾经还有一份随包分发的《使用说明.txt》，2026-09-28 用户已删：
        #   用法由他当面讲给使用的人听，随包文档没人看、还会过期。别再写回来。）
        n_pm = b.get("prog_multi", 0)
        n_name_only = b.get("name_only", 0)
        n_ordinal = b.get("ordinal", 0)
        n_osl = b.get("one_sided_loose", 0)
        n_blank = b.get("pdf_blank", 0)
        n_missing = b.get("missing", 0)
        n_filled = b.get("blank_filled", 0)

        # 账外说明：**一行**说完「哪些没进统计」。原来拆成「空值：…」
        # 「另有 … —— 这些都不进统计。」两行，占两行却只回答一个问题。
        # _foot / _ex_parts 两个中间变量已于 2026-09-28 合并掉，别再加回来。
        _tail = []
        if n_blank:
            _tail.append("%d 格 PDF 本就写空" % n_blank)
        if n_missing:
            _tail.append("%d 条漏填" % n_missing)
        if n_filled:
            _tail.append("%d 条多填" % n_filled)
        if _tail:
            _tail[0] = "空值：" + _tail[0]
        _tail2 = []
        if c["collection"]:
            _tail2.append("%d 条字符串 / 布尔值采集项按设计不核对" % c["collection"])
        if c["novalue"]:
            _tail2.append("%d 条类型可核对、但这一格本就没填值" % c["novalue"])
        if _tail2:
            _tail.append("；".join(_tail2) + "（都不进统计）")
        if _tail:
            add()
            add("  " + " ｜ ".join(_tail) + "。", "dim")

        # 【PDF 参数来源】整段删除（2026-10-30 用户点名）。
        # 用户原话：「概览底下那个 PDF 参数来源没有人会去看的啊」。
        # 它原本是判断「无对应」可不可信的底账，但五种来源各占一行、绝大多数
        # 时候全是 0 点，每天开工先划过去的就是这一段。
        # 底账并没有丢 —— 只换了地方：
        #   · 导出报告里照旧是完整的一段（见 `_export` 里那行「PDF 参数来源：…」）；
        #   · 真读不出来的时候，「抽取自检」那一段会自己冒出来说（见下）；
        #   · 每条「无对应」的**依据**列本来就写着「PDF 已读出 N 张表 / M 个参数点」。
        # 唯一保留的一句：**五种来源全没读到** = 这份 PDF 根本没读进来，这种时候
        # 必须当场说，不能等用户去导报告才发现。
        _src_all = (n_step, n_vert, n_meta_items, n_prog, n_mic)
        if not any(_src_all):
            add("  这份 PDF 里一种参数来源都没读到 —— 先确认它是电子版 PDF"
                "（有文字层），不是打印后扫描的图片。", "warn")

        # 【抽取自检】的判词已经在最上面出现过一次，这里再打一遍是重复；
        # 原来紧跟着的「参数点：横表 48 / 竖表 18 …」又和上面【PDF 参数来源】
        # 是同一批数。于是：干净时整段不出现，有疑点时只列判词之外的具体疑点。
        sc_result = getattr(self, "pdf_selfcheck", None)
        if sc_result and sc_result.get("suspect"):
            _sc_detail = [ln for i, ln in enumerate(sc_result.get("lines") or [])
                          if i > 0 and not ln.startswith("参数点：")]
            if _sc_detail:
                sec("抽取自检 · 疑点")
                for _ln in _sc_detail:
                    add("  " + _ln, "warn")      # 这一段就是「疑点」，整段标橙
        wi_warn = getattr(self, "wi_warning", "")
        if wi_warn:
            add("  " + wi_warn, "warn")
        warn = getattr(self, "pdf_text_warning", "")
        if warn:
            add("  " + warn, "warn")

        # 上面那句「其中 N 条是宽松判定」数的就是这个列表 —— 两边必须同一口径，
        # 否则会出现「说 26 条、只列 24 条」的账对不上。meta_multi（表头元信息
        # 两套候选，如 T-1114（232）/T-1114（255））正属于「候选真的不止一个」。
        # one_sided_loose 也必须在这里：它的期望值虽然唯一，但 Excel 那一侧只填了
        # 一半、PDF 只是「没超出」，跟「两边完全相同」不是一回事，不能当硬判定混过去。
        loose = [v for v in r["checks"]
                 if v.get("scope") in ("prog_multi", "meta_multi", "multi", "device",
                                       "union", "approx", "one_sided_loose",
                                       "ordinal", "name_only")]
        if loose:
            # 标题不能只说「期望值没锁死到唯一值」：one_sided_loose 的期望值恰恰是
            # 唯一的，松的是 Excel 那一侧（只填了一侧、PDF 只是没超出）。
            # 说成「口径没锁到最严」才两种都装得下。
            sec("要人工确认的（%d 条：判了「一致」，但口径没锁到最严）" % len(loose))
            # 原来这里是「三句解释 + 12 条样例」，最长能到 15 行 —— 员工根本不会读，
            # 反而把上面【要处理的】顶出视野。改成：**一句总括 + 3 条样例**，
            # 剩下的一律指去明细表/导出报告（那里是全的）。
            _warn_bits = []
            if n_ordinal:
                _warn_bits.append("%d 条按规格位次硬对" % n_ordinal)
            if n_name_only:
                _warn_bits.append("%d 条只按参数名命中" % n_name_only)
            if n_pm:
                _warn_bits.append("%d 条附表级多套候选" % n_pm)
            if n_osl:
                _warn_bits.append("%d 条 Excel 只填了一侧限值" % n_osl)
            if _warn_bits:
                add("  " + "；".join(_warn_bits) + " —— 都没判成不一致，采用前请对着 PDF 看一眼。",
                    "warn")
            for v in loose[:3]:
                add("   · %s · %s ｜ 现值 %s ｜ 候选 %s"
                    % (v["card"], v["item"],
                       _same_tilde(_value_str(v) + (v["unit"] or "")),
                       _same_tilde(_fmt_expected(v))), "warn")
            if len(loose) > 3:
                add("   …另有 %d 条：明细表里筛「依据」列可全列出（导出报告里也是全的）。"
                    % (len(loose) - 3), "dim")

        self._show(B)
        self._update_tabs(c, len(r["checks"]))
        # 状态栏和概览抬头用同一个词。原来这里写「无法匹配」、抬头写「找不到对应」，
        # 同屏两个说法指同一批行，看的人得自己换算（2026-10-30 统一）。
        # 2026-11-04 第九轮再压一轮（用户：「把所有的字都给精简一点」）：
        # 「多套待确认」→「多套」、「找不到对应」→「无对应」—— 四个数字一排在
        # 880px 的最小窗口里本来就快顶到边，省下这 7 个字正好留出余量。
        # 2026-11-09：多套候选改按不一致计（用户：「多套就直接判错」），
        # 所以这里写进括号里，不再单列一个「多套」让它在视觉上另立门户。
        _mm_txt = str(c["mismatch"])
        if c.get("multi"):
            _mm_txt += "(含多套 %d)" % c["multi"]
        self.status.set(f"完成：{self.wi_name} 一致 {c['ok']} 不一致 {_mm_txt} "
                        f"无对应 {c['no_match']}")
        # 表格有行了 → 收掉空态那行引导（2.1 改版加的）
        self._sync_empty_hint()
        # 采集项列宽按这一份数据的实际内容定（不靠窗口拉伸，见 _fit_item_col）
        self._fit_item_col()
        # match（核对结果 / 分组条那句话）跟着定宽、文案按最终预算重铺 —— 必须在
        # _fit_item_col **之后**：item 宽度一变，右缘剩多少也跟着变。
        self._fit_match_col()
        self._refresh_grp_texts()
        # 分组按钮上的字跟着「现在有多少组是收着的」走
        self._sync_fold_btn()
        # 整张表重画完，选中行的灰框跟上（多半此时是空的，但收着也无妨）
        self._ring_refresh()



    def _fit_item_col(self):
        """按**内容**定「采集项」列宽，而不是按窗口宽度拉伸。

        2026-11-03 第四轮，用户原话：「你都给它固定死了，包括采集项这边都没占满，
        你是不是应该把这个间距给它调整一下呢？把它往左移一点，使右边空出来更多
        位置，让右边到时候以后或许要加一些新的东西。」

        为什么不能留 stretch=True：窗口一宽，那一列就吃掉几百像素空白，看着像
        「这一列没填满」，而且列位会随窗口漂。这里改成量最长的那条采集项名，
        加上一点左右留白，夹在 200~420 之间 —— 内容短的文件不会白占地方，
        右边的空档就留给以后要加的列。
        """
        rows = getattr(self, "_seq", None) or []
        if not rows:
            return
        try:
            import tkinter.font as tkf
            from tkinter import ttk as _ttk
            fname = _ttk.Style().lookup("Treeview", "font") or "TkDefaultFont"
            fnt = tkf.Font(font=fname)
        except Exception:      # noqa: BLE001 —— 量不出来就保持默认宽度
            return
        px_items = []
        for er, _chk in rows:
            # 判成一致的元信息行已经不画在表上了（见 _group_view），
            # 它的采集项却往往是全表最长的（「将定位标记固定在内管上并热焊接
            # 机器编号」这种一长串），再拿它去定列宽，等于给一行不存在的字
            # 留出 100px 空白 —— 实测 WI-0028 采集项列因此是 319px。
            if (_chk or {}).get("meta_kind") and not _pending(_chk):
                continue
            s = str((er or {}).get("item") or "")
            if s:
                px_items.append(fnt.measure(s))
        # 取「大部分行需要多宽」，不是「最长那条多宽」（见 _bulk_px）。
        # WI-0029 实测：182 行里 180 行的采集项 144px，另外 2 行把采集项写成了
        # 一整句话（358px）—— 按 max 定宽 382px，整张表被这两行推到右边，
        # 类型/上下限 离判定结果隔出一大块空白（用户 2026-11-11 点名的就是这个）。
        px = _bulk_px(px_items)
        # 分组条的前缀（DWG-0000009-球囊成型）也占这一列，而且它是 F_HEAD 加粗、
        # 比正文宽三成 —— 不量它就会截成「DWG-0000009-球囊成」（第六轮实测踩到）。
        # 前缀是**每张卡的标识**，属于硬需求：不参与分位淘汰，直接抬下限。
        pref_px = 0
        try:
            gf = self._grp_font() or fnt
            for card in list(getattr(self, "_grp_children", None) or ()):
                pref, _tail = _card_split(card)
                s = pref or str(card)
                if s:
                    pref_px = max(pref_px, gf.measure(s))
        except Exception:      # noqa: BLE001 —— 量不了就只按数据行走
            pass
        px = max(px, pref_px)
        # 下限 200 → 170（2026-11-04 第六轮）：内容量出来普遍 150~175px，
        # 200 的下限是白白占 30px。上限 420 → 360（2026-11-11）：再宽就不是
        # 「采集项」而是「采集项 + 一整段说明」，多出来的宽度全给右边的核对列。
        w = max(170, min(360, px + 24))
        try:
            if abs(int(self.tree.column("item", "width")) - w) > 2:
                self.tree.column("item", width=w)
        except Exception:      # noqa: BLE001
            pass



    def _update_tabs(self, c, n_rows=0):
        """结论数字写在按钮旁边（2026-10-29）。

        原来写进 Notebook 的标签标题（「概览（3 条要改）」）；标签条为了省
        29px 高度已经去掉了，数字改挂到按钮右边那行小字上 —— 作用不变：
        扫一眼就知道这份表要不要管，点不点进「概览」都行。
        """
        bits = []
        if c.get("mismatch"):
            bits.append("%d 条要改" % c["mismatch"])
        if c.get("multi"):
            bits.append("%d 条待确认" % c["multi"])
        if c.get("no_match"):
            bits.append("%d 条对不上" % c["no_match"])
        note = "、".join(bits) if bits else "没有要改的"
        self.page_note.configure(
            text=note + ("　·　明细 %d 条" % n_rows if n_rows else ""),
            foreground=theme.BAD_FG if bits else theme.OK_FG)
