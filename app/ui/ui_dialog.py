"""四个弹窗：名字确认、学习记录、钉格、结构自检。

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

import os
import re
import bisect
import tkinter as tk
from tkinter import messagebox, ttk

import ui_theme as theme
from ui_fmt import _looks_like_value

try:
    import name_lex
except Exception:      # 缺了就降级，不影响核对
    name_lex = None

try:
    import pdf_tables
except Exception:      # 缺了就降级，不影响核对
    pdf_tables = None

try:
    import tab_health
except Exception:      # 缺了就降级：表目录不显示「体检」列，别的照旧
    tab_health = None

# 表目录「体检」列的字。四种结论各给一个能被看懂的短词 ——
# 都说成「需检查」反而看不出是漏抽还是抽得零散（这两个的后续动作不一样）。
_HEAL_TXT = {"ok": "正常", "weak": "零散", "bad": "漏抽", "not_param": "—"}

try:
    import xl_locate
except Exception:      # 缺了就降级，不影响核对
    xl_locate = None

try:
    import anchors
except Exception:      # 缺了就降级，不影响核对
    anchors = None

try:
    import pdf_render
except Exception:      # 缺了就降级 —— ④ 那栏变成一句说明，别的照旧
    pdf_render = None

try:
    from PIL import ImageTk as _ImageTk
except Exception:      # 缺了就降级（Pillow 是 pdfplumber 的依赖，正常都会有）
    _ImageTk = None



# 「按住左键拖」那一套搬去了 ui_pan（主界面、主页清单也要用，不能只长在这个
# 窗里）。这里仍按原名导入，下面 _pin_dialog 里的调用一个字都不用改。
from ui_pan import _PanGrab, _PanDrag                     # noqa: E402
from ui_pan import tv_click as _tv_click                  # noqa: E402


class DialogMixin:
    """四个弹窗：名字确认、学习记录、钉格、结构自检。"""


    # ------------------------------------------------------------------ 词库
    #  说明见 name_lex.py 文件头。三条边界在这里再写一遍：
    #   1) 只影响「名字对齐」，数值比较与判定归引擎，一行不碰；
    #   2) 跳过 / 关闭不影响任何已有结论（未确认的名字照旧落原来的档）；
    #   3) 确认后写进同目录的 `词库.json`，保存前自动备份，可整份删掉回滚。
    #
    #  2026-11-04 第七轮：原来这里还有一个 `_maybe_ask_names()` —— 比对完
    #  自动把「新参数名待确认」窗弹出来。它已经不是调用路径（更早一轮就改成
    #  手动进了），用户这一轮又把「弹窗」这件事整个否掉：「现在已经很明确的
    #  标记出来哪里有问题了，根本不需要有弹窗这个东西存在」。所以直接删掉，
    #  `_ask_names()` 本体留着 —— 它是个纯手动窗，日后要挂入口随时能挂。
    def _ask_names(self, items):
        """模态确认窗：每个名字只问一次，确认后永久记住。

        默认关闭在「关掉就走」—— 不点也不会改变本次结论，避免打断出报告。

        —— 2026-10-30 重做（用户反馈「上手有点难度」）——
        两处根本问题：
          ① 候选只给「相似度 > 0.34 的前 8 个」。「真空跑漏」和 PDF 里任何一个
             名字都算不出相似度，下拉框就是空的 —— 等于没法选。改成：PDF 抽到的
             **全部**参数名列出来（按相似度排序、单位相同打 ·），上面带搜索框。
          ② 「行数」只给个数字，回 Excel 定位不了。改成「行号」：给出数据源工作表
             的**具体行号**，外加「表4 第几行」—— 工具读的是展开表（rev06），
             使用者核对时打开的是表4，两个行号天生不相等，只报一个必然翻错行。
        """
        try:
            pool, units = name_lex.pdf_name_pool(getattr(self, "pdf_vertical", None),
                                                 getattr(self, "pdf_steps", None))
        except Exception:
            pool, units = [], {}
        # —— 表目录（2026-10-30）——
        # 用户原话：「先选中是 P、D、F 当中的哪一个表，然后再提取这个表中的哪一个字段」
        #           「可以在比对的时候加一个它是第几页的表几，既能方便人工核对，
        #             也方便机器提高准确率」。
        # 这里只算出「每个 PDF 名字出自哪张表」，用于显示和「按表筛选候选」；
        # 不筛掉任何候选、不改任何结论（候选池、相似度、词库都一字未动）。
        _cat, _tkey = [], {}
        try:
            if pdf_tables is not None:
                _cat = pdf_tables.catalog(self.pdf_var.get().strip(),
                                          getattr(self, "pdf_vertical", None))
                _tkey = pdf_tables.name_where(getattr(self, "pdf_vertical", None))
        except Exception:
            _cat, _tkey = [], {}
        # 只在「步骤通道」里出现的名字没有页/表可言，单独标一下，不混进「—」
        _stepnames = set()
        for _r in (getattr(self, "pdf_steps", None) or ()):
            if (_r.get("param") or "").strip():
                _stepnames.add(_r["param"].strip())
        # 「第6页 表2」+ 表名 —— 右栏那列放得下的短写法（完整表名见「表目录」）
        _lbl_by_key, _nm_by_key = {}, {}
        for _c in _cat:
            _lbl_by_key[(_c["page"], _c["tno"])] = _c["label"]
            _nm_by_key[(_c["page"], _c["tno"])] = _c.get("name") or ""
        # 数据源工作表名（引擎实际读的那张）与 Excel 路径：两者一起才能说清行号。
        _src = ""
        try:
            _src = (self.schema.chosen.name if self.schema and self.schema.chosen else "")
        except Exception:
            _src = ""
        _xlsx = self.xlsx_var.get()
        # 「表4」在这次的文件里到底叫什么（找不到就 None，界面显示「—」）
        _t4 = None
        if xl_locate is not None:
            try:
                _t4 = xl_locate.find_sheet(_xlsx)
            except Exception:
                _t4 = None
        top = tk.Toplevel(self.root)
        top.configure(bg=theme.BG)      # 弹窗也走三级底色，不留系统灰
        top.title("新参数名确认 —— %d 个（跳过不影响结论）" % len(items))
        top.transient(self.root)
        top.geometry("1230x640")
        top.minsize(1000, 520)
        changed = {"n": 0}

        head = ttk.Frame(top, padding=(10, 8))
        head.pack(fill="x")
        ttk.Label(head, justify="left", text=(
            "这些名字本工具的词典里没有。确认一次会记进「词库.json」，以后同类 WI 自动生效。\n"
            "· 左栏是 Excel 里的写法，右栏是 PDF 里抽到的名字 —— 点一下就行，不用打字；\n"
            "· 行号：%s 的行号；「%s」那一列是这个采集项在你平时看的表4 里第几行"
            "（两张表粒度不同，行号不会相等）。\n"
            "· 右栏「表」那一列：这个名字出自「第几页、文档自己印的表几」。"
            "摸不清该挑哪个名字时，先按「表目录」看那张表叫什么、有哪些字段。\n"
            "· 不确定就跳过 —— 跳过只是保持现状，不会让结论出错。"
            % (("「%s」" % _src) if _src else "数据源工作表", _t4 or "表4"))).pack(anchor="w")

        mid = ttk.Frame(top, padding=(10, 0))
        mid.pack(fill="both", expand=True)

        # ---- 左：未知名字清单（行号是「数据源表!行」，表4 列是反查出来的）----
        lf = ttk.Frame(mid)
        lf.pack(side="left", fill="both", expand=True)
        cols = ("name", "rows", "t4", "unit", "st")
        tv = ttk.Treeview(lf, columns=cols, show="headings", height=10)
        for c, t, w, anc in (
                ("name", "参数名（Excel 里的写法）", 230, "w"),
                ("rows", "行号（%s）" % (_src or "Excel"), 130, "w"),
                ("t4", _t4 or "表4", 95, "w"),
                ("unit", "单位", 52, "center"),
                ("st", "现落在", 83, "w")):
            tv.heading(c, text=t)
            tv.column(c, width=w, anchor=anc)
        tv.pack(side="left", fill="both", expand=True)
        _PanDrag(tv)                       # 没滚动条：按住左键拖（第九轮统一）
        for i, it in enumerate(items):
            _rl = it.get("rowlist") or []
            if _rl:
                _rt = ",".join(str(x) for x in _rl[:4])
                if it["rows"] > 4:
                    _rt += "…共 %d 行" % it["rows"]
            else:
                _rt = "—"
            _t4txt = "—"
            if xl_locate is not None and _t4:
                try:
                    _t4txt = xl_locate.describe(_xlsx, it["name"], _t4)
                except Exception:
                    _t4txt = "—"
            tv.insert("", "end", iid=str(i), values=(
                it["name"], _rt, _t4txt, it["unit"] or "—",
                "/".join(sorted(x for x in it["statuses"] if x)) or "—"))

        # ---- 右：PDF 里的全部参数名（可选 + 可搜 + 出自哪张表）----
        rf = ttk.LabelFrame(mid, text="PDF 里抽到的参数名（共 %d 个）" % len(pool),
                            padding=(8, 6))
        rf.pack(side="left", fill="both", padx=(10, 0))
        q_v = tk.StringVar()
        qrow = ttk.Frame(rf)
        qrow.pack(fill="x")
        ttk.Label(qrow, text="搜索").pack(side="left")
        ttk.Entry(qrow, textvariable=q_v, width=16).pack(side="left", padx=(4, 6))
        # 「表目录」紧贴搜索框 —— 这两个都是「找名字」的入口，放一起顺着用。
        # 按表筛选的取消按钮走右边（默认不显示，不占位）。
        ttk.Button(qrow, text="表目录", width=8,
                   command=lambda: _show_tables()).pack(side="left")
        tab_off = ttk.Button(qrow, text="× 取消表筛选", width=13,
                             command=lambda: _tab_filter(None))
        # 两个列表都可能在窄屏上放不下：给各自留一条横向滚动条。
        # 另外选中右栏某一条时，把它的完整名字和出处写到下面提示行 ——
        # 宽 1230 的窗里「表」那列必然要截字，截了不能让人靠猜。
        tf = ttk.Frame(rf)
        tf.pack(fill="both", expand=True)
        ctree = ttk.Treeview(tf, columns=("pdf", "t", "u"), show="headings", height=11)
        for c, t, w, anc in (("pdf", "PDF 参数名", 180, "w"),
                             ("t", "表（第几页 表几 表名）", 240, "w"),
                             ("u", "单位（同）", 96, "center")):
            ctree.heading(c, text=t)
            ctree.column(c, width=w, anchor=anc)
        # 两条滑条都撤了（第九轮）：按住左键拖，上下左右都行；放不下就往右拖。
        ctree.pack(side="left", fill="both", expand=True)
        _PanDrag(ctree)

        # —— 下半：处置区（一行放完，省下来的高度给上面的候选列表）——
        box = ttk.LabelFrame(top, text="处理选中的这条", padding=(10, 6))
        box.pack(fill="x", padx=10, pady=(8, 0))
        r1 = ttk.Frame(box)
        r1.pack(fill="x")
        ttk.Label(r1, text="Excel 名").grid(row=0, column=0, sticky="w")
        name_v = tk.StringVar()
        ttk.Entry(r1, textvariable=name_v, width=32, state="readonly").grid(
            row=0, column=1, sticky="w", padx=(6, 16))
        ttk.Button(r1, text="确认：是同一参数", width=17,
                   command=lambda: _as_alias()).grid(row=0, column=2)
        ttk.Label(r1, text="   或者").grid(row=0, column=3, sticky="w")
        kind_v = tk.StringVar(value="侧标")
        cb = ttk.Combobox(r1, textvariable=kind_v, width=10, state="readonly",
                          values=("侧标", "工序名", "说明", "确认项", "表头"))
        cb.grid(row=0, column=4, sticky="w", padx=(4, 4))
        ttk.Button(r1, text="标记：不是参数", width=15,
                   command=lambda: _as_non()).grid(row=0, column=5)
        ttk.Button(r1, text="跳过这条", width=10,
                   command=lambda: _skip()).grid(row=0, column=6, padx=(14, 0))
        hint = ttk.Label(top, foreground=theme.TEXT2, text="")
        hint.pack(fill="x", padx=12, pady=(4, 0))

        def _cur():
            sel = tv.selection()
            if not sel:
                hint.configure(text="先在列表里点一条。")
                return None
            return items[int(sel[0])]

        _cand = {"list": []}
        _tabf = {"key": None}      # 正在按哪张表筛（(页, tno)；None = 不筛）

        def _unit_of(p):
            us = sorted({str(u).strip() for u in (units.get(p) or ()) if str(u).strip()})
            return "/".join(us) if us else "—"

        def _tab_of(p):
            """这个名字出自哪张表（只显示）。步骤通道里的名字没有页/表，单说。"""
            k = _tkey.get(p)
            if k is not None and k in _lbl_by_key:
                lab, nm = _lbl_by_key[k], _nm_by_key.get(k) or ""
                if nm:
                    return "%s %s" % (lab, nm if len(nm) <= 14 else nm[:14] + "…")
                return lab
            return "步骤表" if p in _stepnames else "—"

        def _tab_filter(key):
            """按表筛候选。key=None 取消。

            ⚠️ 这里筛的只是**右栏给人看的候选列表**，`_cand["list"]` 一个都没删 ——
            取消筛选、保存、判定全都还按完整候选池走。所以它不会让任何一条结论
            从「能选」变成「选不到」。
            """
            _tabf["key"] = key
            if key:
                lab = ""
                for c in _cat:
                    if (c["page"], c["tno"]) == key:
                        lab = c["label"]
                        break
                tab_off.configure(text="× %s" % (lab or "表"))
                tab_off.pack(side="right", padx=(0, 4))
            else:
                tab_off.pack_forget()
            _refresh_cands()

        def _refresh_cands(_e=None):
            """按搜索词 / 选中的表过滤右栏。都不设 = **全部列出**，只是相似度高的排前面。

            故意**不自动选中**任何一条：默认选中 + 一键确认 = 把「相似度最高」
            当成「就是它」，而相似度只是排序，不是证据（「真空跑漏」和
            「真空测漏」很像，但它也可能压根不是参数）。宁可多一次点击。
            """
            q = (q_v.get() or "").strip().lower()
            tk_ = _tabf["key"]
            ctree.delete(*ctree.get_children())
            n = 0
            for i, c in enumerate(_cand["list"]):
                if q and q not in c["pdf"].lower():
                    continue
                if tk_ is not None and _tkey.get(c["pdf"]) != tk_:
                    continue
                ctree.insert("", "end", iid=str(i), values=(
                    c["pdf"], _tab_of(c["pdf"]),
                    (_unit_of(c["pdf"]) + (" ·" if c["same_unit"] else ""))))
                n += 1
            if tk_ is not None:
                hint.configure(text="只列这张表的字段：%d 个 —— 点一个再确认"
                                    "（按错了点搜索框右边那个「×」取消）。" % n)
            elif q:
                hint.configure(text="搜「%s」筛出 %d 个 —— 点一个再确认。" % (q, n))
            else:
                hint.configure(text="右栏 %d 个候选，越靠上越像 —— 点一个再确认。" % n)

        def _show_tables():
            """表目录（只读）：这份 PDF 有哪些表、文档印的表号是什么、有哪些字段。

            双击一行 = 「只看这张表的字段」，给右栏加一道筛选 —— 这就是用户说的
            「先选中是哪一个表，再提取这张表里的字段」那一步的人工入口。
            筛选只动「显示」，候选池与判定一个都不动（见 _tab_filter）。
            """
            if not _cat:
                hint.configure(text="这份 PDF 没有识别出表格 —— 表目录是空的。")
                return
            rows = [c for c in _cat if c.get("kind") == "table"]
            if not rows:
                hint.configure(text="这份 PDF 只认出页眉条，没有数据表。")
                return
            # 每张表「读懂了没有」——纯函数，缺 tab_health 时整列显示「—」，不影响别的。
            # 必须把四路记录折成「每页读到几条」一起传进去：catalog 的 nrec 只算竖表，
            # 只拿它判会把「千分尺路径读到的表」误标成漏抽（WI-0011 第 8~11 页踩过）。
            _hmap = {}
            if tab_health is not None:
                try:
                    _bypage = tab_health.page_points(
                        getattr(self, "pdf_steps", None),
                        getattr(self, "pdf_vertical", None),
                        getattr(self, "pdf_programs", None),
                        getattr(self, "pdf_micrometer", None))
                    _hmap = {(it["page"], it["tno"]): it["health"]
                             for it in tab_health.assess_all(rows, _bypage)}
                except Exception:
                    _hmap = {}

            def _health_of(c):
                return _hmap.get((c["page"], c["tno"]))

            w = tk.Toplevel(top)
            w.configure(bg=theme.BG)
            w.title("表目录 —— %s" % os.path.basename(self.pdf_var.get().strip()))
            w.transient(top)
            w.geometry("1180x560")
            ttk.Label(w, padding=(12, 8), justify="left", text=(
                "表号是文档自己印的「表N」，不是工具内部序号 —— 照这个号去 PDF 里翻才翻得到。\n"
                "「—」= 这张表上方没印表号（或离得太远），此时只按「第几页第几张表」称呼，不猜号。\n"
                "「抽出/参数行」= 工具从这张表里取到几条参数 / 这张表有几个真正有字的行（空行不算）。\n"
                "「体检」红 = 像参数表却整页一条都没读到（真漏抽）；橙 = 抽得零散或表头没认出来；"
                "灰 = 物料表、图页，本来就没参数。\n"
                "黑字「本表 0 条，同页另读到 N 点」= 这张表本身没读出，但同一页别的表读到了 —— 数据没丢，只是走的是另一条路。\n"
                "双击一行 = 只看这张表的字段：回到确认窗，右栏就只剩这张表的参数名。"))\
                .pack(anchor="w")
            fr = ttk.Frame(w, padding=(12, 0))
            fr.pack(fill="both", expand=True)
            cols = ("lbl", "name", "en", "rows", "nrec", "heal", "hd", "why")
            tvw = ttk.Treeview(fr, columns=cols, show="headings")
            for c, t, wd, anc in (
                    ("lbl", "位置", 96, "w"),
                    ("name", "表名（文档印的）", 200, "w"),
                    ("en", "英文/副标题", 128, "w"),
                    ("rows", "行数", 40, "center"),
                    ("nrec", "抽出/参数行", 82, "center"),
                    ("heal", "体检", 58, "center"),
                    ("hd", "表头字段", 176, "w"),
                    ("why", "为什么", 268, "w")):
                tvw.heading(c, text=t)
                tvw.column(c, width=wd, anchor=anc)
            # 体检结论直接染在整行字上：红=真漏抽，橙=抽得零散，灰=本来没参数。
            # 只染字不铺底 —— 铺底会把「表头行」的层级压掉（第十轮定的规矩）。
            tvw.tag_configure("h_bad", foreground=theme.BAD_FG)
            tvw.tag_configure("h_weak", foreground=theme.WARN_FG)
            tvw.tag_configure("h_na", foreground=theme.MUTED)
            tvw.pack(side="left", fill="both", expand=True)
            _PanDrag(tvw)                  # 没滚动条：按住左键拖（第九轮统一）
            keys = []
            for c in rows:
                keys.append((c["page"], c["tno"]))
                hl = _health_of(c)
                lv = hl.get("level") if hl else None
                tag = {"bad": "h_bad", "weak": "h_weak",
                       "not_param": "h_na"}.get(lv)
                nrec = c.get("nrec")
                _den = c.get("filled_rows") or c["rows"]
                tvw.insert("", "end", iid=str(len(keys) - 1), tags=(tag,) if tag else (),
                           values=(
                               c["label"], c.get("name") or "—", c.get("en") or "—",
                               c["rows"],
                               "—" if nrec is None else "%d/%d" % (nrec, _den),
                               _HEAL_TXT.get(lv, "—"),
                               " | ".join(c.get("headers") or []) or "—",
                               (hl or {}).get("why") or ""))

            def _pick(_e=None):
                sel = tvw.selection()
                if not sel:
                    return
                key = keys[int(sel[0])]
                w.destroy()
                _tab_filter(key)
                try:
                    ctree.focus_set()
                except Exception:
                    pass

            tvw.bind("<Double-1>", _pick)
            ft = ttk.Frame(w, padding=(12, 8))
            ft.pack(fill="x")
            ttk.Button(ft, text="只看这张表的字段", width=18,
                       command=_pick).pack(side="left")
            ttk.Button(ft, text="关闭", width=10,
                       command=w.destroy).pack(side="right")
            # 高度跟着表数走：写死 560 的时候，只有 10 张表就空一大片，
            # 几十张表又滚不动。必须**等按钮也建完**再量，否则脚底那两个按钮
            # 会被算出窗外（第一版就是漏了这一步）。上限压在屏幕内。
            try:
                w.update_idletasks()
                _sw, _sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
                _ww = min(1180, max(900, _sw - 80))
                _hh = min(max(w.winfo_reqheight() + 8, 360), _sh - 120)
                w.geometry("%dx%d" % (_ww, _hh))
            except Exception:
                pass

        def _peek(_e=None):
            """选中右栏某一条 → 提示行把完整名字和出处写全。

            列宽是有限的（1230 的窗里「表」那列必然截字），截了不能让人靠猜：
            「Fold Head Close Stop M…」这种，到这里就变成
            「选中：Fold Head Close Stop Micrometer Setting ｜出自 第6页 表3 球囊折叠工艺参数表」。
            """
            sel = ctree.selection()
            if not sel:
                return
            p = str(ctree.item(sel[0], "values")[0])
            k = _tkey.get(p)
            if k is not None and k in _lbl_by_key:
                lab, nm = _lbl_by_key[k], _nm_by_key.get(k) or ""
                where = ("｜出自 %s %s" % (lab, nm)) if nm else ("｜出自 %s" % lab)
            elif p in _stepnames:
                where = "｜出自步骤表（横表，没有页/表号）"
            else:
                where = ""
            hint.configure(text="选中：%s%s" % (p, where))

        def _fill(_e=None):
            it = _cur()
            if not it:
                return
            name_v.set(it["name"])
            # min_score=0：一个都不滤掉。像「真空跑漏」这种和 PDF 里任何名字都算不出
            # 相似度的，卡阈值的结果就是右栏空的 —— 那正是「没法选」的根源。
            _cand["list"] = name_lex.candidates(
                it["name"], pool, unit=it["unit"], pool_units=units,
                top=max(len(pool), 1), min_score=0.0)
            q_v.set("")          # 换一条名字就清掉上一条的搜索词（trace 会重建右栏）

        def _drop(iid):
            if tv.exists(iid):
                tv.delete(iid)

        def _as_alias():
            it = _cur()
            if not it:
                return
            sel = ctree.selection()
            pdf = str(ctree.item(sel[0], "values")[0]).strip() if sel else ""
            if not pdf:
                hint.configure(text="先在右栏点一个 PDF 里抽到的参数名。")
                return
            ok = name_lex.add_alias(it["name"], pdf, unit=it["unit"],
                                    src=self._src_tag())
            if ok:
                changed["n"] += 1
                _drop(tv.selection()[0])
                hint.configure(text="已记住：%s ↔ %s" % (it["name"], pdf))

        def _as_non():
            it = _cur()
            if not it:
                return
            ok = name_lex.add_non_param(it["name"], kind_v.get(),
                                        src=self._src_tag())
            if ok:
                changed["n"] += 1
                _drop(tv.selection()[0])
                hint.configure(text="已记为「%s」，以后不再问：%s" % (kind_v.get(), it["name"]))

        def _skip():
            sel = tv.selection()
            if sel:
                _drop(sel[0])

        def _close():
            if changed["n"] and messagebox.askyesno(
                    "词库已更新", "记住 %d 条。现在按新词库重跑一遍吗？\n"
                    "（重跑只影响名字对齐，数值结论仍由原判定给出）" % changed["n"]):
                top.destroy()
                self._no_names_dialog = True
                try:
                    self._run()
                finally:
                    self._no_names_dialog = False
                return
            top.destroy()

        tv.bind("<<TreeviewSelect>>", _fill)
        q_v.trace_add("write", lambda *_: _refresh_cands())
        ctree.bind("<<TreeviewSelect>>", _peek)
        # 右栏双击 = 直接确认（少按一次按钮）。单击只是选中，不改变任何东西。
        ctree.bind("<Double-1>", lambda _e: _as_alias())
        foot = ttk.Frame(top, padding=(10, 8))
        foot.pack(fill="x")
        ttk.Label(foot, foreground=theme.TEXT2, text="词库文件：%s" % name_lex.path_of()).pack(side="left")
        ttk.Button(foot, text="关闭", width=10, command=_close).pack(side="right")
        # —— 窗口高度跟着内容走（10-30）——
        # 固定 640 会把最底下「确认：是同一参数」那一行挤出屏幕：实测内容要 778 高。
        # 这不是本轮引入的（改版前 600 更矮，一样挤出），是这轮量出来顺手修掉的。
        # 上限压在屏幕内，1366x768 的笔记本上也不会冒出屏幕外。
        try:
            top.update_idletasks()
            _sw, _sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
            _w = min(1290, max(1000, _sw - 60))
            _h = min(max(top.winfo_reqheight() + 8, 520), _sh - 90)
            top.geometry("%dx%d+%d+%d" % (_w, _h, max(0, (_sw - _w) // 2),
                                          max(0, (_sh - _h) // 3)))
        except Exception:
            pass
        if tv.get_children():
            tv.selection_set(tv.get_children()[0])
            _fill()




    # ------------------------------------------------------------- 学习记录
    #  用户 2026-11-02 的原话，这个窗存在的全部理由：
    #    「它有个就是你认不出来的话，它不是有判定吗？在那个界面，我觉得你给他
    #      加个入口，我判定过的什么，我是不是可以把那个记录给清空，我没判定过的，
    #      我是不是可以自己手动添加，不然每次只有在对比的时候才弹出来一次。」
    #
    #  以前只有「新参数名确认」那一个模态窗，比对时弹一次，关掉就再也找不回来；
    #  教给工具的东西存在 词库.json / 锚点.json 里，界面上没有任何地方看得见。
    #  这个窗就是那两个文件的界面：看得见、能停用、能删、能清空。
    #
    #  两个文件都自带 .bak-时间戳 备份（见 name_lex._backup / anchors._backup），
    #  所以「清空」是可回滚的 —— 清之前不用手动拷，备份已经在旁边。

    def _src_tag(self):
        """一条学习记录的「出处」：它是在哪一对文件上记下来的。

        2026-11-04 用户原话：「那个你分不清它是对比的什么与什么之间的……
        或许你可以在哪个地方给它加上去，是哪个表与哪个 PDF 对比的时候产生的
        （用简写就好）」。

        为什么必须**在生成那一刻**记下来：学习记录是跨文件累积的 —— 一条别名
        在别的 WI 上也照样生效，等回过头翻这一页时，「当初是哪份表哪份 PDF
        教出来的」只有当时知道。以前只存了 wi（PDF 那边的 WI 号），Excel 那半边
        一个字段都没有，于是记录一多，就分不清哪几条是同一次比对里产生的。

        路径只存**文件名**（不带目录）：这两个 json 是要在几台机器之间拷来拷去的，
        把某人桌面上的目录结构写进去既没用也碍事；显示层只需要认得出是哪份文件。
        """
        try:
            _ch = getattr(self, "schema", None)
            sheet = _ch.chosen.label if (_ch is not None and _ch.chosen) else ""
        except Exception:
            sheet = ""

        def _bn(p):
            try:
                return os.path.basename(str(p or "").strip())
            except Exception:
                return str(p or "")

        return {"wi": getattr(self, "wi_name", "") or "",
                "pdf": _bn(self.pdf_var.get()),
                "xlsx": _bn(self.xlsx_var.get()),
                "sheet": sheet}

    def _open_learn(self):
        """学习记录 —— 把「工具从你这儿学到的东西」讲成人话。

        2026-11-03 改版（用户原话：「学习记录这边的规则也比较乱糟糟的，让人不能
        直接一眼看懂，甚至我自己我都理解不懂有一些。想想办法解决。」）：

        上一版的毛病是**按内部实现分页**：位置锚点 / 别名 / 不核对 / 表级规则，
        四个名字都是代码里的概念，人得先学会这套词才看得懂；表格里的列名也
        照抄字段（「反推出来的字段」「适用范围」），第四页那一列还永远写着
        同一句废话。于是这一版只做三件事：

        ① **页签换成「你做了什么」的说法**：我指过的格子 / 学会的表 /
           换个写法也认得 / 这些不用核 —— 顺序就是事情发生的顺序。
        ② **每页顶上一句话**说清「这一页记的是什么、删了会怎样」，底下不再
           挂大段说明（用户上一轮刚说过「文字太多了」）。
        ③ **列名换成人话**，并把「学会的表」那页真正的内容摆出来
           （怎么认行 / 怎么认列 / 当初钉的那格），不再是同一句话重复 N 遍。

        动作（停用 / 删除 / 清空）全部沿用原来的后端调用与确认文案，
        只换了门面和文案 —— 底下的备份、回滚能力一点没动。
        """
        if name_lex is None and anchors is None:
            messagebox.showinfo("学习记录", "词库/锚点模块没加载起来，这个窗打不开。")
            return
        if getattr(self, "_learn_win", None) is not None:
            try:
                self._learn_win.deiconify()
                self._learn_win.lift()
                self._learn_win.focus_force()
                return
            except Exception:
                pass
        top = tk.Toplevel(self.root)
        top.configure(bg=theme.BG)
        self._learn_win = top
        top.title("学习记录 —— 我教给工具的东西")
        top.transient(self.root)

        def _howto():
            messagebox.showinfo(
                "这四类分别管什么",
                "它们是四个不同的层次，从「一条」到「一批」：\n\n"
                "· 我指过的格子：你指一次，它记住这一行该看哪一格。管一行。\n"
                "· 学会的表：指的那一格让它看出这张表怎么读（行看适用规格、\n"
                "  列看方向+部位），同结构的行一起解决。管一批。\n"
                "· 换个写法也认得：名字对不上、其实是同一个东西，认一次就记住。\n"
                "· 这些不用核：表头、工序名、说明行，本来就不是参数。\n\n"
                "「来自」这一列说的是：这条是在哪一对文件上记下来的 ——\n"
                "前面是 PDF 的 WI 号，后面是 Excel 的工作表名（如 WI-0041 ↔ rev06）。\n"
                "它只说明出身：同一条别名/规则在别的 WI 上也照样生效。\n\n"
                "这四类都不会改判定规则本身 —— 它们只回答「去哪一格取数」和\n"
                "「这个名字等于那个名字」。判定算式在代码里，这里改不动。")

        # ── 顶部：一句话总览（数字是活的，每次刷新都要重算）──
        head = ttk.Frame(top, padding=(10, 8, 10, 0))
        head.pack(fill="x")
        sum_lbl = ttk.Label(head, foreground=theme.TEXT, text="")
        sum_lbl.pack(side="left")
        ttk.Button(head, text="这四类分别管什么", style="Small.TButton",
                   command=_howto).pack(side="right")

        nb = ttk.Notebook(top)
        nb.pack(fill="both", expand=True, padx=10, pady=(6, 0))

        # 页签按「你做了什么」排，顺序即阅读顺序
        SUB = {
            "anchor": "你亲手指过的「这一行该看哪一格」。只管那一行 —— 删掉，"
                      "那一行就退回「对不上」。",
            "rule": "钉过一格之后自动学会的：这张表怎么认行、怎么认列。"
                    "一条能顶十几条单条记录。",
            "alias": "同一样东西的两种叫法：Excel 里这么写 ＝ PDF 里那么写。",
            "non": "表头、工序名、说明行这类，本来就不是要核的参数。"
                   "列在这里，免得天天报「对不上」。",
        }
        PAGES = (("anchor", "我指过的格子"), ("rule", "学会的表"),
                 ("alias", "换个写法也认得"), ("non", "这些不用核"))
        frame, bar, hint_of = {}, {}, {}
        for key, title in PAGES:
            f = ttk.Frame(nb, padding=(8, 6))
            nb.add(f, text=title)
            frame[key] = f
            # 顺序要紧：先钉底部的动作条，再放顶上那句话，最后才是表格 ——
            # 表格 expand 会把中间剩的全吃掉，顺序反了会把说明挤没。
            bar[key] = ttk.Frame(f)
            bar[key].pack(side="bottom", fill="x", pady=(6, 0))
            ttk.Label(f, foreground=theme.TEXT2, text=SUB[key],
                      wraplength=980, justify="left").pack(side="top", anchor="w",
                                                           pady=(0, 4))

        CW = {}          # 每页的列宽，供 _fit() 按像素裁字用

        def _mk_tree(key, cols, empty_text):
            """一页的表格：内容 + 两根滚动条 + 一句空态提示（没记录时才浮出来）。"""
            wrap = ttk.Frame(frame[key])
            wrap.pack(side="top", fill="both", expand=True)
            tv = ttk.Treeview(wrap, columns=[c[0] for c in cols],
                              show="headings", height=11)
            for cid, cap, w, anc in cols:
                tv.heading(cid, text=cap)
                tv.column(cid, width=w, minwidth=52, anchor=anc)
            CW[key] = {c[0]: c[2] for c in cols}
            _PanDrag(tv)                   # 没滚动条：按住左键拖（第九轮统一）
            tv.grid(row=0, column=0, sticky="nsew")
            wrap.rowconfigure(0, weight=1)
            wrap.columnconfigure(0, weight=1)
            tv.tag_configure("off", foreground=theme.NA_FG)
            # 空态：一整片白的表最容易被当成「坏了」。这里不写「暂无数据」这类
            # 死话（项目 UI 指南第 5.3 节），直接说下一步该做什么。
            hl = tk.Label(wrap, text=empty_text, bg=theme.CARD, fg=theme.TEXT2,
                          justify="center", font=(theme.FONT, theme.F_SMALL))
            hint_of[key] = (hl, wrap)
            return tv

        # 按**像素**裁字（2026-11-03）：原来靠猜「一个汉字约 13px」，Mac 的
        # PingFang 下一格汉字要 14px，猜窄了就被下一列吃掉半个词 —— 截图里
        # 「第 4 行第 5」后面直接没了。Treeview 不会自己省略，只能自己量。
        import tkinter.font as _tkfont
        _F = _tkfont.nametofont("TkDefaultFont")

        def _fit(s, px):
            s = str(s or "")
            if _F.measure(s) <= px:
                return s
            while s and _F.measure(s + "…") > px:
                s = s[:-1]
            return (s + "…") if s else ""

        def _keep_sel(tv):
            """重画表格前记下选中行，画完再选回去。

            没有这一下，「停用/启用」点一次之后选中就丢了 —— 想再点回来还得重新
            找那一行（2026-11-02 冒烟里实测出来的：第二下点了没反应）。
            """
            return set(tv.selection())

        def _restore(tv, keys):
            for k in keys:
                try:
                    if tv.exists(k):
                        tv.selection_set(k)
                except Exception:
                    pass

        def _scope_cn(sc):
            """适用范围说成人话（原来是「仅这份 WI: xxx」「产品线: xxx」）。"""
            sc = sc or {}
            t = sc.get("type") or "global"
            if t == "global":
                return "通用"
            return "%s：%s" % ({"wi": "只这份 WI", "family": "同产品线"}.get(t, t),
                               sc.get("value") or "")

        def _base(p):
            """文件名去掉目录和后缀：`/x/y/WI-0041.pdf` → `WI-0041`。"""
            try:
                return os.path.splitext(os.path.basename(str(p or "")))[0]
            except Exception:
                return ""

        def _from_cn(src):
            """「PDF ↔ Excel 表」压成一行短标签，例如 `WI-0041 ↔ rev06`。

            2026-11-04 用户原话：「分不清它是对比的什么与什么之间的……加上去，
            是哪个表与哪个 PDF 对比的时候产生的（用简写就好）」。

            为什么只写这两位：PDF 那边人认的是 WI 号，Excel 那边人认的是工作表名
            （rev06 / 参数卡模板 / 表4）；完整路径和文件名一格放不下，也没人读。
            PDF 号优先取跑比对时记下的 WI 名，没有就退回 PDF 文件名；
            Excel 表名取记录里的 sheet —— 2026-11-04 之前的记录没这一项，
            退回文件名，至少能看出是哪个 Excel。
            """
            src = src or {}
            wi = str(src.get("wi") or "").strip() or _base(src.get("pdf"))
            m = re.search(r"WI[\s\-_]*(\d{3,5})", wi, re.I)
            if m:
                wi = "WI-" + m.group(1)
            sheet = str(src.get("sheet") or "").strip() or _base(src.get("xlsx"))
            if not wi and not sheet:
                return "早先记的（没存出处）"
            return "%s ↔ %s" % (wi, sheet) if (wi and sheet) else (wi or sheet)

        def _ts_cn(ts):
            """「2026-11-04 08:42:31」→「26-11-04 08:42」。

            年份只留两位（记录都是近期的，两位足够区分）、秒没人看 ——
            省下来的宽度给「来自」那一列。认不出的格式原样返回，不猜。
            """
            t = str(ts or "").strip()
            if len(t) >= 16 and t[4] == "-" and t[7] == "-":
                return t[2:16]
            return t

        def _draw(tv, key):
            """画完一页之后：空的把提示浮上来，非空的收起来。"""
            hl, wrap = hint_of[key]
            if tv.get_children():
                hl.place_forget()
            else:
                hl.place(in_=tv, relx=0.5, rely=0.5, anchor="center")

        # ================================================== ① 我指过的格子
        t1 = _mk_tree("anchor", (
            ("xrow", "Excel 行", 66, "center"),
            ("excel", "参数名（Excel 里写的）", 210, "w"),
            ("from", "来自（PDF ↔ 表）", 158, "w"),
            ("where", "指向哪一格", 148, "w"),
            ("val", "那格写的是", 110, "w"),
            ("pdfn", "PDF 里叫它", 190, "w"),
            ("st", "状态", 58, "center"),
            ("ts", "记于", 116, "w")),
            "还没有指过格子。\n\n双击明细表里「对不上」的那一行，\n"
            "在弹窗里点中它真正该看的那一格，这里就会多一条。")

        def _fill1():
            _sel = _keep_sel(t1)
            t1.delete(*t1.get_children())

            def _c(k, v):
                return _fit(v, CW["anchor"][k] - 10)
            if anchors is not None:
                for r in anchors.records(only_on=False):
                    where = "%s 第%d行第%d列" % (r.get("label") or "?",
                                                int(r.get("row") or 0),
                                                int(r.get("col") or 0))
                    t1.insert("", "end", iid=r["key"], values=(
                        _c("xrow", ("%s" % r["excel_row"]) if r.get("excel_row") else "整份"),
                        _c("excel", r.get("excel")),
                        _c("from", _from_cn(r.get("src"))),
                        _c("where", where), _c("val", r.get("value")),
                        _c("pdfn", r.get("pdf_name")),
                        _c("st", "启用" if r.get("enabled", True) else "已停用"),
                        _c("ts", _ts_cn(r.get("ts")))),
                        tags=() if r.get("enabled", True) else ("off",))
            _restore(t1, _sel)
            _draw(t1, "anchor")

        def _a_sel():
            s = t1.selection()
            return s[0] if s else None

        def _a_toggle():
            k = _a_sel()
            if not k:
                return
            r = anchors.load()["anchors"].get(k) or {}
            anchors.disable(k, off=bool(r.get("enabled", True)))
            _fill1()

        def _a_del():
            k = _a_sel()
            if not k:
                return
            if not messagebox.askyesno(
                    "删掉这一条", "这条记录一删，这一行就退回「对不上」\n"
                    "（除非别的办法还能认出它）。\n\n确定删掉？\n"
                    "（删除前会自动备份 锚点.json，能回滚）"):
                return
            anchors.remove(k)
            _fill1()

        def _a_clear():
            st = anchors.stats()
            n = st["anchors"] + st["anchors_off"]
            if not n:
                return
            if not messagebox.askyesno(
                    "清空全部「指过的格子」",
                    "现在有 %d 条，清空之后一条都不剩。\n\n"
                    "· 只影响「照记录直接取值」这条通道，判定算式一格不变；\n"
                    "· 清空前会自动备份 锚点.json（带时间戳，可整份换回来）。\n\n"
                    "确定清空？" % n):
                return
            n = anchors.clear()
            _fill1()
            messagebox.showinfo("已清空", "清掉了 %d 条。备份就在 data/ 目录里。" % n)

        ttk.Button(bar["anchor"], text="停用 / 启用", command=_a_toggle).pack(side="left")
        ttk.Button(bar["anchor"], text="删除选中", command=_a_del).pack(side="left", padx=6)
        ttk.Button(bar["anchor"], text="清空这一页", command=_a_clear).pack(side="left")
        ttk.Label(bar["anchor"], foreground=theme.TEXT2,
                  text="只管一行，删了不影响别的记录").pack(side="right")

        # ================================================== ② 学会的表
        t4 = _mk_tree("rule", (
            ("tbl", "哪张表", 200, "w"),
            ("from", "来自（PDF ↔ 表）", 158, "w"),
            ("how", "它怎么读这张表", 300, "w"),
            ("wit", "当初钉的那一格（见证）", 236, "w"),
            ("st", "状态", 56, "center"),
            ("ts", "记于", 116, "w")),
            "还没有学会任何表。\n\n指过一格之后自动产生 —— 它看到的是"
            "「行看适用规格、\n列看方向+部位」，于是同结构的行不用再一条条指。")

        def _fill4():
            _sel = _keep_sel(t4)
            t4.delete(*t4.get_children())

            def _c(k, v):
                return _fit(v, CW["rule"][k] - 10)
            if anchors is not None:
                for r in anchors.rule_records(only_on=False):
                    w = r.get("witness") or {}
                    hr = r.get("hdr_rows") or []
                    # 「它怎么读这张表」合成一句人话（2026-11-03）。
                    # 上一版把「行怎么认」「列怎么认」拆两列，每列都只剩半句，
                    # 合起来还得自己在脑子里拼 —— 用户说「一眼看不懂」就是这个。
                    # 更早那版更糟：这一列永远是「行看适用规格、列看方向+部位」，
                    # 每行一模一样，三列宽度全白占。
                    how = "行：适用规格（表头 %d 行）　列：方向+部位" % len(hr)
                    # 坐标写在最前面：这列窄了会从尾巴开始截，坐标是最不能丢的
                    wit = "第 %d 行第 %d 列 ← 「%s」%s" % (
                        int(w.get("row") or 0), int(w.get("col") or 0),
                        w.get("excel") or "?", (w.get("spec") or "")[:12])
                    nm = (r.get("name") or "").strip()
                    tbl = (r.get("label") or "?") + ((" " + nm) if nm else "")
                    t4.insert("", "end", iid=r["key"], values=(
                        _c("tbl", tbl), _c("from", _from_cn(r.get("src"))),
                        _c("how", how), _c("wit", wit),
                        _c("st", "启用" if r.get("enabled", True) else "已停用"),
                        _c("ts", _ts_cn(r.get("ts")))),
                        tags=() if r.get("enabled", True) else ("off",))
            _restore(t4, _sel)
            _draw(t4, "rule")

        def _d_toggle():
            s = t4.selection()
            if not s:
                return
            anchors.toggle_rule(s[0])
            _fill4()

        def _d_del():
            s = t4.selection()
            if not s:
                return
            if not messagebox.askyesno(
                    "删掉这条规则",
                    "删掉之后，靠它推出来的那些行会退回「对不上」\n"
                    "（你单独指过的格子在上一页，不受影响）。\n\n确定删掉？\n"
                    "（删除前会自动备份 锚点.json，能回滚）"):
                return
            anchors.drop_rule(s[0])
            _fill4()

        def _d_clear():
            rs = anchors.rule_records(only_on=False)
            if not rs:
                return
            if not messagebox.askyesno(
                    "清空全部「学会的表」",
                    "现在有 %d 条，清空之后一条都不剩。\n\n"
                    "· 只影响「按学到的表结构推坐标」这一层，你手动指过的格子\n"
                    "  （上一页）一条不动；\n"
                    "· 清空前会自动备份 锚点.json。\n\n确定清空？" % len(rs)):
                return
            for r in rs:
                anchors.drop_rule(r["key"])
            _fill4()

        ttk.Button(bar["rule"], text="停用 / 启用", command=_d_toggle).pack(side="left")
        ttk.Button(bar["rule"], text="删除选中", command=_d_del).pack(side="left", padx=6)
        ttk.Button(bar["rule"], text="清空这一页", command=_d_clear).pack(side="left")
        ttk.Label(bar["rule"], foreground=theme.TEXT2,
                  text="一条管一批；推错了就停用，它不再参与推断").pack(side="right")

        # ================================================== ③ 换个写法也认得
        t2 = _mk_tree("alias", (
            ("excel", "Excel 里这么写", 265, "w"),
            ("pdf", "PDF 里那么写", 265, "w"),
            ("from", "来自（PDF ↔ 表）", 158, "w"),
            ("unit", "单位", 62, "center"),
            ("scope", "管到哪儿", 140, "w"),
            ("st", "状态", 58, "center"),
            ("ts", "记于", 116, "w")),
            "还没有别名。\n\n同一件东西两边叫法不同、按名字对不上时，\n"
            "工具会问一次「这俩是同一个吗」——答了「是」就记在这里。")

        def _fill2():
            _sel = _keep_sel(t2)
            t2.delete(*t2.get_children())

            def _c(k, v):
                return _fit(v, CW["alias"][k] - 10)
            if name_lex is not None:
                for r in name_lex.records("aliases"):
                    t2.insert("", "end", iid=r["key"], values=(
                        _c("excel", r.get("excel")), _c("pdf", r.get("pdf")),
                        _c("from", _from_cn(r.get("src"))),
                        _c("unit", r.get("unit")),
                        _c("scope", _scope_cn(r.get("scope"))),
                        _c("st", "启用" if r.get("enabled", True) else "已停用"),
                        _c("ts", _ts_cn(r.get("ts")))),
                        tags=() if r.get("enabled", True) else ("off",))
            _restore(t2, _sel)
            _draw(t2, "alias")

        def _b_toggle():
            s = t2.selection()
            if not s:
                return
            r = name_lex.load()["aliases"].get(s[0]) or {}
            (name_lex.disable if r.get("enabled", True) else name_lex.enable)(
                "aliases", s[0])
            _fill2()

        def _b_del():
            s = t2.selection()
            if not s:
                return
            if messagebox.askyesno("删掉这一条",
                                   "删掉之后这个名字又会回到「不认得」的状态。确定删掉？"):
                name_lex.remove("aliases", s[0])
                _fill2()

        def _b_clear():
            st = name_lex.stats()
            n = st["aliases"] + st["aliases_off"]
            if not n:
                return
            if messagebox.askyesno("清空全部别名",
                                   "现在有 %d 条，清空之后一条都不剩。\n\n"
                                   "· 只影响「名字怎么对上」这一层，判定算式一格不变；\n"
                                   "· 清空前会自动备份 词库.json。\n\n确定清空？" % n):
                n = name_lex.clear_kind("aliases")
                _fill2()
                messagebox.showinfo("已清空", "清掉了 %d 条。" % n)

        ttk.Button(bar["alias"], text="停用 / 启用", command=_b_toggle).pack(side="left")
        ttk.Button(bar["alias"], text="删除选中", command=_b_del).pack(side="left", padx=6)
        ttk.Button(bar["alias"], text="清空这一页", command=_b_clear).pack(side="left")
        ttk.Label(bar["alias"], foreground=theme.TEXT2,
                  text="「Excel 里这么写」＝「PDF 里那么写」").pack(side="right")

        # ================================================== ④ 这些不用核
        t3 = _mk_tree("non", (
            ("name", "名字", 300, "w"),
            ("from", "来自（PDF ↔ 表）", 158, "w"),
            ("kind", "类型", 90, "center"),
            ("scope", "管到哪儿", 150, "w"),
            ("st", "状态", 58, "center"),
            ("ts", "记于", 116, "w")),
            "还没有登记「不用核」的项。\n\n工具把表头、工序名当成参数来核的时候，\n"
            "会问一次「这个不是参数吧」——答了「是」就记在这里。")

        def _fill3():
            _sel = _keep_sel(t3)
            t3.delete(*t3.get_children())

            def _c(k, v):
                return _fit(v, CW["non"][k] - 10)
            if name_lex is not None:
                for r in name_lex.records("non_params"):
                    t3.insert("", "end", iid=r["key"], values=(
                        _c("name", r.get("name")),
                        _c("from", _from_cn(r.get("src"))),
                        _c("kind", r.get("kind")),
                        _c("scope", _scope_cn(r.get("scope"))),
                        _c("st", "启用" if r.get("enabled", True) else "已停用"),
                        _c("ts", _ts_cn(r.get("ts")))),
                        tags=() if r.get("enabled", True) else ("off",))
            _restore(t3, _sel)
            _draw(t3, "non")

        def _c_toggle():
            s = t3.selection()
            if not s:
                return
            r = name_lex.load()["non_params"].get(s[0]) or {}
            (name_lex.disable if r.get("enabled", True) else name_lex.enable)(
                "non_params", s[0])
            _fill3()

        def _c_del():
            s = t3.selection()
            if not s:
                return
            if messagebox.askyesno("删掉这一条",
                                   "删掉之后它又会被当成参数来核。确定删掉？"):
                name_lex.remove("non_params", s[0])
                _fill3()

        def _c_clear():
            n = name_lex.stats()["non_params"]
            if not n:
                return
            if messagebox.askyesno("清空全部",
                                   "现在有 %d 条，清空之后一条都不剩。\n"
                                   "清空前会自动备份 词库.json。\n\n确定清空？" % n):
                n = name_lex.clear_kind("non_params")
                _fill3()

        ttk.Button(bar["non"], text="停用 / 启用", command=_c_toggle).pack(side="left")
        ttk.Button(bar["non"], text="删除选中", command=_c_del).pack(side="left", padx=6)
        ttk.Button(bar["non"], text="清空这一页", command=_c_clear).pack(side="left")
        ttk.Label(bar["non"], foreground=theme.TEXT2,
                  text="这些不再参与核对，也不会再报「对不上」").pack(side="right")

        # ================================================== 刷新：一次算全
        def _refresh_all():
            n_anchor = n_rule = n_alias = n_non = 0
            if anchors is not None:
                st = anchors.stats()
                n_anchor = st["anchors"] + st["anchors_off"]
                try:
                    n_rule = len(anchors.rule_records(only_on=False))
                except Exception:
                    n_rule = 0
            if name_lex is not None:
                st = name_lex.stats()
                n_alias = st["aliases"] + st["aliases_off"]
                n_non = st["non_params"]
            for key, title in PAGES:
                nb.tab(frame[key], text="%s（%d）" % (title, {
                    "anchor": n_anchor, "rule": n_rule,
                    "alias": n_alias, "non": n_non}[key]))
            sum_lbl.configure(
                text="工具现在记住：我指过的格子 %d 条 · 学会的表 %d 张 · "
                     "换个写法 %d 条 · 不用核 %d 条"
                     % (n_anchor, n_rule, n_alias, n_non))
            return n_anchor, n_rule, n_alias, n_non

        foot = ttk.Frame(top, padding=(10, 8))
        foot.pack(fill="x")

        def _rel(p):
            """文件路径只留「从交付根开始」那一截。

            原来写全路径，两个 path 加起来比窗口还宽，直接把右边「关闭」按钮
            顶出了可视区（截图里只剩半个「关」字）。这两个路径本来就在包的
            data/ 下，写「data/锚点.json」既够定位又不占地方。
            """
            try:
                return os.path.relpath(p, os.path.dirname(os.path.dirname(p)))
            except Exception:
                return p

        _files = []
        if anchors is not None:
            _files.append("锚点：%s" % _rel(anchors.path_of()))
        if name_lex is not None:
            _files.append("词库：%s" % _rel(name_lex.path_of()))
        ttk.Button(foot, text="关闭", width=10,
                   command=lambda: (top.destroy(), setattr(self, "_learn_win", None))
                   ).pack(side="right")
        ttk.Label(foot, foreground=theme.TEXT2, text="　｜　".join(_files)
                  ).pack(side="left")

        top.protocol("WM_DELETE_WINDOW",
                     lambda: (top.destroy(), setattr(self, "_learn_win", None)))
        _fill1()
        _fill2()
        _fill3()
        _fill4()
        _refresh_all()
        try:
            top.update_idletasks()
            sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
            w = min(1180, max(940, sw - 120))
            h = min(max(top.winfo_reqheight() + 8, 480), sh - 120)
            top.geometry("%dx%d+%d+%d" % (w, h, max(0, (sw - w) // 2),
                                          max(0, (sh - h) // 3)))
        except Exception:
            pass



    # ------------------------------------------------------- 指定 PDF 位置
    #  「先让我选是 PDF 里哪一页的哪一张表，选出来之后根据表里面的数值」
    #  —— 用户 2026-11-02 定的交互顺序：先选表，再选值，名字由工具自己反推。
    #
    #  2026-11-02 第二次改版（用户原话）：
    #    「如果说比方说有42条对应不上，那我就把42条全部列出来……原逻辑就是一条一条
    #      让我手动去标记。但是如果说你引入了学习机制呢，那比方说有10条，它其实
    #      本质上都是这一个表的……我把第一条给让它识别出来了，它是不是剩下的也能
    #      够识别出来，就直接也辅助清空掉」
    #    「我只能选中一行，而不是选中每一个单元格，这个就不对……我选中这个单元格，
    #      那它是不是根据行列和纵列的表头，我就让他进行机器学习」
    #
    #  所以这一版是三栏，和旧版的差别全在「一」和「三」：
    #    ① 左：**一次列出全部对不上的行**（不再是双击哪条只处理哪条），钉完自动消失
    #    ② 中：PDF 表目录（第几页 表几 + 表名 + 行数）
    #    ③ 右：那张表的原始格子 —— 点哪**格**选中哪格，不是整行
    #  钉住之后：记一条锚点 + 反推一条表级规则 + 自动重跑，左栏条数当场变少。
    #
    #  为什么右栏要用一堆 Label 摆成真表格、而不是继续用 Treeview：
    #    ttk.Treeview 只能选**行**，单元格高亮它做不到。用户要的正是「像一张表格
    #    那样点中一格」，所以这里用 Canvas + 一格一个 Label，选中格直接换底色。
    #    表最多二十来行十几列（WI-0041 表3 = 18×7），几百个 Label 毫无压力。
    def _pin_dialog(self, chk=None):
        """在主界面**右侧**展开 PDF 原图，把人指的那一格记下来。

        2026-11-04 第十三轮：从独立弹窗改成常驻面板。用户原话：
        「像 2 的话就只是个页面而已，1 的话它完全和那个主界面是重叠的，完全没必要
          那我是不是可以现在改为原始界面，就是主界面，我双击之后不要弹出这个新的了，
          而是在主界面的右侧直接扩宽，产生一个显示 PDF 的这个位置表格，让我可以进行
          选取。而且以默认的这个 PDF 比例显示……是高 4 宽 3 的这种比例」

        于是三块内容各归其位：
          · ① 待钉的行 —— 主界面那张明细表本来就是它，删；
          · ② 页码列表 —— 顶栏一行「‹ 第 N/M 页 ›」就够，删；
          · 剩下的整页 PDF + 热区 + 提取结果预览，正好填满右侧那一栏。

        面板只搭一次（挂在 self._pin_panel 上），之后每双击一次只换行、换页。
        这样第二次以后没有任何构建开销，也不会闪一下白框。
        """
        if anchors is None or pdf_tables is None:
            messagebox.showinfo("指定位置", "锚点/表目录模块没加载起来，这个功能用不了。")
            return
        pdf = (self.pdf_var.get() or "").strip()
        if not pdf or not os.path.exists(pdf):
            messagebox.showinfo("指定位置", "先在上面选一份 PDF，再指定位置。")
            return
        if not getattr(self, "result", None):
            messagebox.showinfo(
                "指定位置",
                "先点一次「开始比对」——要先把「哪些行对不上」列出来，才谈得上逐条指位置。")
            return
        cat = [c for c in pdf_tables.catalog(pdf, getattr(self, "pdf_vertical", None))
               if c.get("kind") != "header"]
        if not cat:
            messagebox.showinfo("指定位置",
                                "这份 PDF 里没认出任何表格，没法按位置指定。")
            return

        # ── 面板已经搭好、而且看的是同一份 PDF → 只换行，不重建 ──
        # 换过 PDF（主页上点了别的文件）就必须重建：cat / 页数 / 缓存全是旧的。
        P = getattr(self, "_pin_panel", None)
        if P:
            try:
                if P["root"].winfo_exists() and P.get("pdf") == pdf:
                    P["open"]()
                    P["show"](chk)
                    return
            except Exception:      # noqa: BLE001
                pass
            self._pin_panel = None

        dock = getattr(self, "pdf_dock", None)
        pg_split = getattr(self, "main_split", None)
        if dock is None or pg_split is None:
            messagebox.showinfo("指定位置", "界面还在加载，稍等一下再双击。")
            return
        for _w in list(dock.winfo_children()):
            try:
                _w.destroy()
            except Exception:      # noqa: BLE001
                pass

        _NP = pdf_tables.npages(pdf)
        # 缩放档位。为什么是「重渲染」而不是把图拉大：放大之后要看清的是 PDF 上
        # 那些小字，拿同一张位图放大只会更糊（Retina 下尤其明显）。
        # 1.0 是「适宽」—— 按栏宽反推 scale，正好铺满，不用拖。
        _ZSTEPS = (1.0, 1.5, 2.0, 3.0, 4.0)
        # 展开时窗口往右长出去多少 —— 这个数就是 PDF 栏的宽度。
        # 430px 是「整页 A4 铺满 + 底下一行提取结果还读得下去」的实际下限；
        # 再宽明细表就白占屏幕，再窄 PDF 上的小字看不清。
        _DOCK_W = 430
        _ST_CN = {"ok": "一致", "mismatch": "不一致", "multi": "多套候选",
                  "no_match": "对不上", "no_scope": "无适用"}

        S = {"chk": None, "tab": None, "row": None, "col": None,
             "rule": None,
             # grid / nhdr = 当前那张表的提取结果与表头行数（只取数，不画）
             "grid": [], "nhdr": 1,
             # page = 现在看的是第几页；sel = 点中的那一格 (page, tno, r, c)。
             # 为什么带 tno：整页视图里同一页可能有两三张表，光有行列号分不清是哪张。
             # sels = 多选时**全部**选中的格（顺序即点选顺序，最后一格是主格）。
             #   2026-11-11 用户原话：「用 Ctrl 键或者在 Mac 这边是 Command 键嘛，
             #   就可以进行多个选择，这样子就可以多个表来选」—— 同一个采集项在
             #   一份 PDF 里常常分写两张表（WI-0028 的焊机温度在表4、表5 各一份），
             #   只钉一格等于逼人在两格里二选一，而它们本来就该一起看。
             "page": None, "sel": None, "sels": [],
             "zoom": 1.0,
             # pmeta = render_ex 的换算信息；pboxes = 图像像素热区；
             # phot / sel_rect = 悬停框 / 选中框的 Canvas item id。
             "pmeta": None, "pboxes": [], "phot": None, "sel_rect": None,
             "pimg": None, "pkey": None,
             # widen_px = 这次为了放 PDF 把窗口**加宽**了多少像素。
             #   0 = 没加宽（屏幕放不下，退回「挤压原窗口」的老做法）。
             #   收起时要照着它把窗口瘦回去，否则右边会留一块空地。
             # geo0/geo1 = 展开前 / 我们改成的几何。"收起"会把窗口完整还原成
             #   geo0 —— 除非这中间用户自己拖过窗口，那就只减宽度、别拉他回去。
             "widen_px": 0, "geo0": None, "geo1": None,
             "opened": False}

        root_w = ttk.Frame(dock)
        root_w.pack(fill="both", expand=True)

        # ── 顶栏：翻页 + 放大 + 收起 ──────────────────────────────
        # 原来这里是 ① 一行说明、② 一张页码表。页码表删掉之后，翻页就靠这一对
        # 箭头 —— 钉格子时绝大多数时候是「跟着这一行走」，手翻只是兜底。
        head = ttk.Frame(root_w, padding=(6, 4, 6, 2))
        head.pack(fill="x")
        ttk.Button(head, text="收起", width=5, style="Small.TButton",
                   command=lambda: _collapse()).pack(side="right")
        S["zoomtab"] = tk.BooleanVar(value=False)
        # 勾上 = 只钉着这一张表看（裁切 + 放大）。名字描述「勾上会发生什么」，
        # 不描述「现在是什么状态」—— 默认是整页，所以不能叫「整页」。
        ttk.Checkbutton(head, text="放大", variable=S["zoomtab"],
                        command=lambda: _pdf_show(force=True)
                        ).pack(side="right", padx=(0, 6))
        ttk.Button(head, text="‹", width=2, style="Small.TButton",
                   command=lambda: _step(-1)).pack(side="left")
        _pgmark = ttk.Label(head, text="", foreground=theme.TEXT2)
        _pgmark.pack(side="left", padx=3)
        ttk.Button(head, text="›", width=2, style="Small.TButton",
                   command=lambda: _step(+1)).pack(side="left")

        # ── 图：独占剩下的全部高度 ────────────────────────────────
        cbox = ttk.Frame(root_w)
        cbox.pack(fill="both", expand=True, padx=6)
        pcv = tk.Canvas(cbox, bg=theme.CARD, highlightthickness=1,
                        highlightbackground=theme.BORDER)
        pcv.pack(fill="both", expand=True)

        # ── 缩放条：贴在图**正下方**（左键拖=平移，中键=放大）────────
        zbar = ttk.Frame(root_w, padding=(6, 3, 6, 0))
        zbar.pack(fill="x")
        ttk.Button(zbar, text="−", width=2, style="Small.TButton",
                   command=lambda: _pdf_zoom(-1)).pack(side="left")
        zlab = ttk.Label(zbar, text="100%", foreground=theme.TEXT2,
                         width=5, anchor="center")
        zlab.pack(side="left")
        ttk.Button(zbar, text="+", width=2, style="Small.TButton",
                   command=lambda: _pdf_zoom(+1)).pack(side="left")
        # 「适宽」用链接式扁平字：放大到 400% 之后要回原样，连点减号太笨；
        # 但它不是常用操作，不配抢第四个按钮位。
        zfit = ttk.Label(zbar, text="适宽", foreground=theme.TEXT2, cursor="hand2")
        zfit.pack(side="left", padx=(8, 0))
        zfit.bind("<Button-1>", lambda _e: _pdf_zoom(0, reset=True))

        # ── 「这一格提取出来是什么」（用户点名要的）────────────────
        # 三行：位置 / 原文 / 上下文。值的来源必须是 S["grid"][r][c]，
        # 也就是**判定引擎真正读到的那串字**，不是重新裁一次 —— 按格裁切取字
        # 实测只有 89.2% 与 extract() 一致，差异全是「文字画了两遍」
        # （DWG-0000009-02 → DDWWGG2200101001--0022）。拿它给人看会让人以为
        # 工具读错了，其实判定用的那串是好的。
        pk = tk.Frame(root_w, bg=theme.FIELD, highlightthickness=1,
                      highlightbackground=theme.BORDER)
        pk.pack(fill="x", padx=6, pady=(4, 0))
        pk_in = tk.Frame(pk, bg=theme.FIELD)
        pk_in.pack(fill="x", padx=7, pady=4)
        _WRAP = 330
        pick_t1 = tk.Label(pk_in, text="还没点格子", bg=theme.FIELD,
                           fg=theme.TEXT2, anchor="w", justify="left",
                           wraplength=_WRAP)
        pick_t1.pack(fill="x")
        # 提取值：12 号加粗，整栏只有这一处放大字号 —— 它是这一屏里唯一
        # 「要你核对对不对」的东西，值得抢一次注意力。
        pick_t2 = tk.Label(pk_in, text="在图上点一下 → 这里显示提取结果",
                           bg=theme.FIELD, fg=theme.TEXT2, anchor="w",
                           justify="left", wraplength=_WRAP,
                           font=(theme.FONT, 12, "bold"))
        pick_t2.pack(fill="x", pady=(1, 0))
        pick_t3 = tk.Label(pk_in, text="", bg=theme.FIELD, fg=theme.TEXT2,
                           anchor="w", justify="left", wraplength=_WRAP)
        pick_t3.pack(fill="x", pady=(1, 0))

        # ── 底栏：一句状态 + 一颗钉住 ──────────────────────────────
        # 2026-11-03 用户原话：「在那个对不上的界面当中，底下还是太多文字了」。
        # 所以只有一行，钉住之后的结果（成功/失败/重判中）也走同一行 ——
        # 同一时刻只有一个状态，不用上下扫两趟。
        bar = ttk.Frame(root_w, padding=(6, 4, 6, 5))
        bar.pack(fill="x")
        # 按钮上的字跟着选中格数走：一格「钉住这格」、多格「钉住 N 格」——
        # 否则多选了却还写着「这格」，人不知道点下去会钉几个。
        btn_pin = ttk.Button(bar, text="钉住这格 →", width=11,
                             command=lambda: _pin())
        btn_pin.pack(side="right")
        hint = ttk.Label(bar, justify="left", foreground=theme.TEXT2,
                         wraplength=250,
                         text="在图上点那一格（按住 ⌘/Ctrl 可多选）")
        hint.pack(side="left", fill="x", expand=True)

        def _wrap_to(host, labels, pad=14, lo=120):
            """让文字的折行宽度跟着容器走。

            固定 `wraplength` 在面板被拖窄时**不是折行而是截断** —— Tk 只在
            wraplength 处折行，超出容器宽度的部分直接看不见。面板从「固定
            430px」改成可拖宽之后，这两个数就不该再是常数（多选提示多出来的
            那截字，在窄面板上会被切掉尾巴）。
            """
            def _on(e):
                try:
                    if int(e.width) < 60:
                        return          # 还没排版完的宽度（1、2px）别当真
                    wd = max(lo, int(e.width) - pad)
                    for _lb in labels:
                        if abs(int(_lb.cget("wraplength")) - wd) > 6:
                            _lb.configure(wraplength=wd)
                except Exception:      # noqa: BLE001
                    pass
            host.bind("<Configure>", _on, add="+")

        # 绑在 hint 自己身上而不是 bar 上：bar 还含着右边那颗按钮，按 bar 的宽度
        # 算出来的折行宽会比 hint 实际拿到的宽，右边照样被切。
        _wrap_to(hint, [hint], pad=6)
        _wrap_to(pk_in, [pick_t1, pick_t2, pick_t3])

        # ============================================================
        #  以下全是内部实现（顺序可读性优先，不是调用顺序）
        # ============================================================

        def _place_sash():
            """把分隔条摆到合适位置。两套口径，取决于窗口有没有被我们撑宽：

              · 撑宽了（S["widen_px"] 有值）→ 右边**固定**就那一条，
                左边回到展开前的原样，一像素都不挪；
              · 没撑宽（窗口本来就贴满屏幕）→ 退回挤压：PDF 占 28% 窗宽，
                不低于 340px。

            第十三轮后面板是全窗高的，不再按「高 4 宽 3」从高度反推 ——
            那个比例是给 A4 纸的，面板还包了标题/缩放/预览/底栏约 150px。
            """
            try:
                self.root.update_idletasks()
                total = int(pg_split.winfo_width())
                if total < 400:
                    return
                add = int(S.get("widen_px") or 0)
                if add >= 240:
                    w = min(add, max(240, total - 300))
                else:
                    w = max(340, min(520, int(total * 0.28)))
                pg_split.sashpos(0, max(180, total - w))
            except Exception:      # noqa: BLE001
                pass

        def _grow(on):
            """窗口让位：展开时往右长一条，收起时把那一条收回去。

            用户原话（2026-11-09）：「能不能直接扩大原窗口，强制扩大原窗口。
            除非这个原窗口是全屏状态，就它只要不是全屏状态，就直接强制给它
            扩大一部分，在那扩大的一部分上面再去放 PDF，这样就不会挤压原空间。」

            三档处理：
              1. 右边有地方 → 往右长出 _DOCK_W（右边界顶到屏幕就整体左移）；
              2. 右边不够、左边还有 → 整体挪到靠左，再长；
              3. 屏幕真放不下（最大化 / 贴满 / 已接近全屏）→ 一个像素都不改，
                 widen_px 归零，_place_sash 自动退回「挤压」老做法。
            """
            r = self.root
            try:
                w, h = int(r.winfo_width()), int(r.winfo_height())
                x, y = int(r.winfo_x()), int(r.winfo_y())
                sw, sh = int(r.winfo_screenwidth()), int(r.winfo_screenheight())
            except Exception:      # noqa: BLE001
                return
            m = 16                       # 离屏幕边的安全距离
            if w < 240 or h < 200:       # 窗口还没映射出来，量不到真尺寸
                return
            if not on:
                add = int(S.get("widen_px") or 0)
                if add:
                    g0, g1 = S.get("geo0"), S.get("geo1")
                    cur = (w, h, x, y)
                    if g0 and g1 and all(abs(cur[i] - g1[i]) <= 4
                                         for i in range(4)):
                        # 这中间用户没动过窗口 → 原样还回去（位置也还回去，
                        # 别让窗口收起来之后凭空往左跳了 200px）
                        try:
                            r.geometry("%dx%d+%d+%d" % g0)
                            r.update_idletasks()
                        except Exception:      # noqa: BLE001
                            pass
                    else:
                        # 用户挪过/拖过 → 只把我们加的那一条减掉，别的照旧
                        nw = max(620, w - add)
                        nx = max(m, min(x, sw - m - nw))
                        try:
                            r.geometry("%dx%d+%d+%d" % (nw, h, nx, y))
                            r.update_idletasks()
                        except Exception:      # noqa: BLE001
                            pass
                S["widen_px"] = 0
                S["geo0"] = S["geo1"] = None
                return
            avail = sw - m * 2
            if y + h > sh - m:           # 加宽后别让底边出屏
                y = max(m, sh - m - h)
            nw = min(w + _DOCK_W, avail)
            add = nw - w
            if add < 240:                # 挤不出像样的一条 → 不硬撑，走挤压
                S["widen_px"] = 0
                return
            nx = max(m, min(x, sw - m - nw))
            S["widen_px"] = add
            S["geo0"] = (w, h, x, y)
            S["geo1"] = (nw, h, nx, y)
            try:
                r.geometry("%dx%d+%d+%d" % (nw, h, nx, y))
                r.update_idletasks()
            except Exception:      # noqa: BLE001
                S["widen_px"] = 0

        def _open():
            if S.get("opened"):
                return
            try:
                pg_split.add(dock, weight=1)
            except Exception:      # noqa: BLE001
                return
            S["opened"] = True
            # 先把窗口撑宽，再摆线。反过来的话 pane 会按窄窗口先分一次宽度，
            # 图先按小尺寸渲染一遍再重渲染，白闪一下。
            _grow(True)
            # 摆三次：after_idle 时 pane 可能还没分到真实宽度；120ms 后那次
            # 才基本准了；260ms 兜底 —— 改窗口大小是交给窗口管理器的异步操作，
            # macOS 上偶尔比 120ms 还慢。
            self.root.after_idle(_place_sash)
            self.root.after(120, _place_sash)
            self.root.after(260, _place_sash)

        def _collapse():
            """收起面板 —— 等价于原来的「关闭」，只是不销毁面板本体。

            收起后再双击任意一行，面板原样回来（还停在上次那一页），
            不用重搭、也不用重新渲染缓存过的图。窗口一起瘦回去。
            """
            try:
                pg_split.forget(dock)
            except Exception:      # noqa: BLE001
                pass
            S["opened"] = False
            _grow(False)

        # ---------------- 读表（只取数，不画）----------------
        def _hdr_count(tab, g):
            """这张表前面几行是表头（只用于「哪几格是数据」，不参与判定）。

            两条判据合用，取更保守的那个：

            ① **像测量值的行一定是数据行**。命中 `±`、`~数字`、或以数字/φ 开头
               带着单位的，就是数据。这一条必须优先 —— 反过来判错（把数据行当
               表头）会让那一格点不动，用户要的「点中单元格」直接失效。
            ② 否则，非空格子大半能在这张表的字段名（`catalog` 的 `headers`）里
               找到 → 表头。

            ⚠️ 两次实测踩出来的教训：
               · 只数「开头几行非空」→ WI-0039 表4-1 的第 2 行是
                 「DWG-0000013-01 | 拉伸前切定长 | 820±10mm」，会被算成表头，
                 820±10mm 就点不动了；
               · 只用判据② → WI-0041 表3 的 `headers` 只有英文那行，第 1、2 行
                 （「球囊直径」「Proximal End/近端」）不算表头，顶栏就只剩英文，
                 「上限/下限/近端」这些关键信息全丢（截图上一眼看出来的）。
            最多认 3 行表头，至少 1 行。
            """
            names = [str(h or "").strip() for h in (tab.get("headers") or [])]
            names = [h for h in names if h]
            n = 0
            for row in g[:3]:
                cells = [str(x or "").strip() for x in row if str(x or "").strip()]
                if not cells:
                    break
                if any(_looks_like_value(c) for c in cells):
                    break                       # 像测量值 → 这是数据行，到此为止
                hit = 0
                for c in cells:
                    for h in names:
                        if c in h or h in c:
                            hit += 1
                            break
                if names and hit == 0 and n >= 1 and len(cells) < 2:
                    # 既不像值、跟字段名无关、又只占一格：是合并的标题/说明行，不算表头。
                    # （占 2 格以上的照样算表头 —— 「近端 | 中间 | 远端」这种二级表头
                    #   跟字段名一个都对不上，但它确实是表头。）
                    break
                n += 1
            return max(1, n)

        def _tab_of(page, tno):
            """(页, 表号) → catalog 里那条记录；找不到给 None。

            为什么靠 (page, tno) 找而不是靠下标：图上每一格自带的坐标就是这两个
            数，拿它直接反查是最短的路径 —— 不需要「先选表、再点格」两步。
            """
            for _t in cat:
                if _t.get("page") == page and _t.get("tno") == tno:
                    return _t
            return None

        def _load(page, tno):
            """把某张表的格子读进内存。不画 —— 画的是原图，格子只是热区。"""
            tab2 = _tab_of(page, tno)
            if tab2 is None:
                S["grid"], S["nhdr"] = [], 1
                return None
            try:
                g = pdf_tables.grid(pdf, page, tno) or []
            except Exception:      # noqa: BLE001
                g = []
            S["grid"] = g
            S["nhdr"] = _hdr_count(tab2, g) if g else 1
            return tab2

        def _cell_ctx(page, tno, r, c):
            """(页, 表, 行, 列) → (表记录, 行标签, 格子原文, 列头整串)。

            多选拼 `cells` 用。主格那张表直接用已经读好的 `S["grid"]`（和界面上
            预览的是同一份，不会两处不一致）；别的表现读一次 —— 只取数不重画，
            点一下钉住不该为此把整个面板重渲染一遍。
            """
            t2 = _tab_of(page, tno)
            _cur = S.get("sel")
            if _cur and tuple(_cur[:2]) == (page, tno):
                g2, nh = (S.get("grid") or []), (S.get("nhdr") or 1)
            else:
                try:
                    g2 = pdf_tables.grid(pdf, page, tno) or []
                except Exception:      # noqa: BLE001
                    g2 = []
                nh = (_hdr_count(t2, g2) if (t2 and g2) else 1)
            rl2 = (g2[r][0] if r < len(g2) and g2[r] else "")
            v2 = (g2[r][c] if r < len(g2) and c < len(g2[r]) else "")
            try:
                st = anchors._col_stacks(g2, list(range(0, max(1, nh))))
                ct2 = (st[c] if c < len(st) else "") or ""
            except Exception:          # noqa: BLE001
                ct2 = ""
            return t2, rl2, (v2 or ""), ct2

        def _box_of(page, tno, r, c):
            """这一格在当前这张图上的像素框（没画出来/没这一格 → None）。"""
            for b in (S.get("pboxes") or ()):
                if b[0] == page and b[1] == tno and b[2] == r and b[3] == c:
                    return b
            return None

        def _col_text(c, upto=None):
            """这一列的表头整串（含合并表头的向右继承，与 anchors._col_stacks 同口径）。

            `upto` 不给就用表头行数。反推规则时传「人钉的那一行」——
            `anchors.derive` 里 hdr_rows 正是 `range(0, row)`，这里必须跟它一致，
            否则面板里的预览和真存下来的规则会不是同一件事。
            """
            g = S.get("grid") or []
            try:
                n = upto if upto else max(1, S.get("nhdr") or 1)
                st = anchors._col_stacks(g, list(range(0, n)))
                return (st[c] if c < len(st) else "") or ""
            except Exception:
                return ""

        def _pick(page, tno, r, c, add=False):
            """点中一格 —— (页, 表号, 行, 列)。

            整页视图下必须先说清「你点的是哪张表」，否则同一页两张表都叫
            「第 3 行第 2 列」，钉下去必错。所以坐标里把 page/tno 也带上。

            `add=True`（按住 Ctrl / Command 点）是**加选**：已经在集合里的再点
            一次就取消。为什么要多选 —— 2026-11-11 用户原话：「用 Ctrl 键或者在
            Mac 这边是 Command 键嘛，就可以进行多个选择，这样子就可以多个表来选。
            ……它这里对应了两个表，它没有分开写，取一起的话都是 157 到 181。
            那这个我自己盯的话，也不知道该盯哪个才能辅助机器学习呢」。
            多选的格子在 `_pin` 里合成**一条**锚点（`cells`），判定时按并集取范围。
            """
            key = (page, tno, r, c)
            if add:
                sels = list(S.get("sels") or [])
                if key in sels:
                    sels.remove(key)           # 再点一下 = 取消它
                else:
                    sels.append(key)           # 顺序即点选顺序，最后一格是主格
            else:
                sels = [key]
            S["sels"] = sels
            if sels:
                pg, tn = sels[-1][0], sels[-1][1]
                cur = S.get("sel")
                if not cur or tuple(cur[:2]) != (pg, tn):
                    _load(pg, tn)              # 主格换了表才重新读格子
                S["sel"] = sels[-1]
                S["tab"] = _tab_of(pg, tn)
                S["row"], S["col"] = sels[-1][2], sels[-1][3]
            else:
                S["sel"] = None
                S["row"] = S["col"] = None
                S["tab"] = _tab_of(page, tno)
            _hot_sel()                         # 选中框（常亮，和悬停框分开）
            _preview()                         # 「提取出来是什么」
            _describe()

        # ---------------- 原图上的热区 ----------------
        def _build_hot(page2, img, meta):
            """把这一页上每张表每格的 bbox 换算成图上的像素框，供鼠标命中。

            条目是 `(page, tno, r, c, x0, y0, x1, y1)`。换算不出来（meta 缺、
            格子没坐标、框在图外）就一格都不建 —— 宁可点不动，也不画偏的框：
            偏的框比没有更误导人。
            """
            S["pboxes"] = []
            if not meta or page2 is None:
                return
            W, H = img.size
            out = []
            for t in cat:
                if t.get("page") != page2:
                    continue
                try:
                    boxes = pdf_tables.cell_boxes(pdf, page2, t.get("tno"))
                except Exception:      # noqa: BLE001
                    continue
                if not boxes:
                    continue
                # 表头那几行**不建热区**：钉锚点是为了「这一行该读哪个数」，
                # 表头是列名，钉上去取到的是一串标题。不然鼠标在表头上也变手型、
                # 点下去还是个合法动作，白钉一次。
                try:
                    _g = pdf_tables.grid(pdf, page2, t.get("tno")) or []
                    nh = _hdr_count(t, _g) if _g else 1
                except Exception:      # noqa: BLE001
                    nh = 1
                for r, row in enumerate(boxes or ()):
                    if r < nh:
                        continue
                    for c, cb in enumerate(row or ()):
                        if not cb:
                            continue        # 合并格 / 老版取不到坐标 —— 不猜
                        px = pdf_render.to_px(cb, meta)
                        if not px:
                            continue
                        x0, y0, x1, y1 = px
                        if x1 <= x0 or y1 <= y0:
                            continue
                        # 完全在图外的丢掉：表被裁掉一半、坐标算飘时都会留一堆
                        # 越界框，让命中判定白跑一遍
                        if x1 <= 0 or y1 <= 0 or x0 >= W or y0 >= H:
                            continue
                        out.append((page2, t.get("tno"), r, c, x0, y0, x1, y1))
            S["pboxes"] = out

        def _hot_at(x, y):
            """画布坐标 → 命中的那一格；没命中给 None。"""
            for b in (S.get("pboxes") or ()):
                if b[4] <= x < b[6] and b[5] <= y < b[7]:
                    return b
            return None

        def _hot_paint(b):
            """悬停提示框：1px 灰虚线，只在鼠标下面那格闪一下。

            用虚线 + 中性灰（不用主色）：它只说明「这一格可以点」，不表示
            「已经选中」—— 后者是 `_hot_sel` 那个 3px 实线主色框。
            """
            old = S.get("phot")
            if old is not None:
                try:
                    pcv.delete(old)
                except Exception:      # noqa: BLE001
                    pass
                S["phot"] = None
            if not b:
                return
            try:
                S["phot"] = pcv.create_rectangle(
                    b[4], b[5], b[6], b[7],
                    outline=theme.TEXT2, width=1, dash=(4, 2), tags="hot")
            except Exception:          # noqa: BLE001
                S["phot"] = None

        def _hot_sel():
            """已经点中的格子：主色实线，一直亮着。

            为什么不复用悬停框：悬停框每次鼠标一动就删了重建，点完就看不见自己
            选了哪格了。而钉之前人往往要拖一下图、放大看看，选中框必须扛得住。

            多选（Ctrl / Command）时按 `S["sels"]` 全部画出来：**主格**（最后点的
            那格）3px 实线，其余 2px 虚线 —— 主格决定底下「提取结果」显示哪一格，
            副格只是「这一条也要一起看」，两者必须看得出差别。
            用 tag 删而不是记 id：格数会变，一个个记账容易漏删。
            """
            try:
                pcv.delete("sel")
            except Exception:      # noqa: BLE001
                pass
            S["sel_rect"] = None
            sels = list(S.get("sels") or ())
            if not sels and S.get("sel"):
                sels = [S["sel"]]
            for i, s in enumerate(sels):
                b = _box_of(s[0], s[1], s[2], s[3])
                if not b:
                    continue
                main = (i == len(sels) - 1)
                try:
                    _rid = pcv.create_rectangle(
                        b[4], b[5], b[6], b[7], outline=theme.PRIMARY,
                        width=3 if main else 2,
                        dash=() if main else (4, 3), tags="sel")
                    if main:
                        S["sel_rect"] = _rid
                except Exception:      # noqa: BLE001
                    pass

        def _pdf_motion(e):
            # 正在拖动平移时别抢光标、也别刷框 —— 一刷就跟 _PanGrab 的 hand2
            # 打架，指针抖得看不清
            try:
                if pcv.grab_current() is not None:
                    return
            except Exception:      # noqa: BLE001
                pass
            b = _hot_at(pcv.canvasx(e.x), pcv.canvasy(e.y))
            _hot_paint(b)
            # 能不能点，鼠标形状先告诉人 —— 只在形状真变了才 configure，
            # 每次移动都设一遍会让指针发颤
            cur = "hand2" if b else "fleur"
            try:
                if str(pcv.cget("cursor")) != cur:
                    pcv.configure(cursor=cur)
            except Exception:      # noqa: BLE001
                pass

        def _pdf_click(e, add=False):
            """点中一格 → 记下来 + 底下显示「提取出来是什么」。

            `add=True` 是按住 Ctrl / Command 的多选（见 `_pick`）。
            """
            b = _hot_at(pcv.canvasx(e.x), pcv.canvasy(e.y))
            if not b:
                return
            page2, tno2, r2, c2 = b[0], b[1], b[2], b[3]
            try:
                _pick(page2, tno2, r2, c2, add=add)
            except Exception:          # noqa: BLE001
                pass

        # ---------------- 把某一页画到画布上 ----------------
        def _draw_table_boxes(page2, meta, tab2):
            """整页视图：把这一页**所有认出来的表**都圈出来。

            2026-11-11 用户原话：「现在是选中某一个框，它才会标出来这一个大的表格
            周围的那个红线。我觉得你提前先把这个红线给标记出来，这样就知道哪里有
            表格，哪里的表格我们可以选择，哪里的表格你没有识别到。」

            于是分两档画（同一套 `to_px` 换算，和格子蓝框、选中框永远对齐）：
              · 认出来的每张表 → 1px 虚线红框（这些区域都能点）；
              · 当前正在看的那张 → 2px 实线红框（叠在上面，语义和以前一样）。
            没有框的空白 = 这份 PDF 这一块没有参数表，或者工具没认出来 ——
            这正是用户想看见的「哪里你没识别到」。
            """
            if not meta or page2 is None:
                return 0
            n = 0
            cur = (tab2 or {}).get("tno")
            for t in cat:
                if t.get("page") != page2 or not t.get("bbox"):
                    continue
                try:
                    tb = pdf_render.to_px(t["bbox"], meta)
                except Exception:      # noqa: BLE001 —— 这一张圈不出来就跳过
                    continue
                if not tb:
                    continue
                x0, y0, x1, y1 = tb
                if x1 <= x0 or y1 <= y0:
                    continue
                _on = (cur is not None and t.get("tno") == cur)
                try:
                    pcv.create_rectangle(
                        x0, y0, x1, y1, outline=theme.BAD_FG,
                        width=2 if _on else 1,
                        dash=() if _on else (3, 3), tags="tbox")
                    n += 1
                except Exception:      # noqa: BLE001
                    pass
            return n

        def _pdf_show(page2=None, tab2=None, force=False, anchor=None,
                      keep_pos=False):
            """按当前页把 PDF 画上去。

            `force=False` 时同一页、同一种看法、宽度没变就不重画 —— Configure
            是滚动、拖分隔线都会触发的，每次都重渲染会把拖拽拖成幻灯片。

            `anchor=(x, y)`：缩放时把鼠标底下那个内容点**钉在原地**。
            `keep_pos=True`：栏宽变了（拖分隔线）时保住当前看的位置。
            """
            try:
                if not root_w.winfo_exists():
                    return
            except Exception:      # noqa: BLE001
                return
            if page2 is None:
                page2 = S.get("page")
            if tab2 is None:
                tab2 = S.get("tab")
            if page2 is None:
                _pdf_msg("这一页没加载出来")
                return
            if pdf_render is None or _ImageTk is None or not pdf_render.available():
                _pdf_msg("本机没装渲染 PDF 的组件（pypdfium2），这一栏看不了原图。")
                return
            S["page"] = page2
            # 「放大」勾了、而且确实有一张表 → 只画那一块；否则整页（默认）。
            zoomtab = False
            try:
                zoomtab = bool(S["zoomtab"].get())
            except Exception:      # noqa: BLE001
                zoomtab = False
            whole = not (zoomtab and tab2)
            bbox = None if whole else tab2.get("bbox")
            # 缩放前先记住「鼠标下 / 左上角」那个点在内容里的相对位置（0~1）。
            # 必须在换图之前量 —— 换完 bbox 就是新图的了。
            keep = None
            try:
                _bb0 = pcv.bbox("all")
                if _bb0 and _bb0[2] > 2 and _bb0[3] > 2:
                    if anchor is not None:
                        keep = (pcv.canvasx(anchor[0]) / float(_bb0[2]),
                                pcv.canvasy(anchor[1]) / float(_bb0[3]))
                    elif keep_pos:
                        keep = (max(0.0, pcv.canvasx(0)) / float(_bb0[2]),
                                max(0.0, pcv.canvasy(0)) / float(_bb0[3]))
            except Exception:      # noqa: BLE001
                keep = None
            try:
                want = int(pcv.winfo_width()) - 2
            except Exception:      # noqa: BLE001
                want = 0
            if want < 260:                 # 还没布局完，先别急着渲染
                want = 520
            # 缩放：按 pane 宽度反推 scale，所以「放大」不是把同一张图拉大（那只会
            # 更糊），而是按更大的像素重新渲染一次，字是真的变清楚。
            want = max(200, min(4000, int(want * float(S.get("zoom") or 1.0))))
            key = (page2, (tab2 or {}).get("tno") if not whole else None,
                   whole, want // 40)
            if not force and S.get("pkey") == key:
                return
            S["pkey"] = key
            img = None
            meta = None
            try:
                # 用 render_ex 而不是 render：要把「裁切原点 / 缩放倍率 / 旋转角」
                # 一起拿到手，否则换算不出每个格子在图上落在哪几个像素。
                img, meta = pdf_render.render_ex(pdf, page2, bbox,
                                                 target_w=want, whole=whole)
            except Exception:      # noqa: BLE001
                img, meta = None, None
            if img is None:
                _pdf_msg("这一页画不出来（PDF 可能加密或损坏）。",
                         theme.BAD_FG)
                return
            pcv.delete("all")
            S["phot"] = None
            S["sel_rect"] = None
            try:
                S["pimg"] = _ImageTk.PhotoImage(img)
            except Exception:      # noqa: BLE001
                _pdf_msg("图片显示失败（Pillow 装了但 Tk 不支持这个位图）。")
                return
            pcv.create_image(0, 0, anchor="nw", image=S["pimg"])
            W, H = img.size
            pcv.configure(scrollregion=(0, 0, W, H))
            # 图换了 → 热区必须跟着重算（缩放、拖分隔线、切页都会走到这）
            S["pmeta"] = meta
            _build_hot(page2, img, meta)
            if keep:
                # 让那个内容点回到原来的屏幕位置：左上角分数 = (rx·W − 鼠标x) / W
                _ox = anchor[0] if anchor is not None else 0
                _oy = anchor[1] if anchor is not None else 0
                pcv.xview_moveto(max(0.0, min(1.0, (keep[0] * W - _ox) / float(W))))
                pcv.yview_moveto(max(0.0, min(1.0, (keep[1] * H - _oy) / float(H))))
            else:
                pcv.xview_moveto(0)
                pcv.yview_moveto(0)
            # 整页视图要把**这一页上所有识别到的表**都圈出来 —— 2026-11-11 用户原话：
            # 「现在是选中某一个框，它才会标出来这一个大的表格周围的那个红线。
            #   我觉得你提前先把这个红线给标记出来，这样就知道哪里有表格，
            #   哪里的表格我们可以选择，哪里的表格你没有识别到。」
            #
            # 以前只圈「当前那张」（`if whole and tab2`），于是打开一页什么提示都没有，
            # 只能靠鼠标乱扫看哪里变手型；而「没圈的空白」到底是没有表、还是没认出来，
            # 也分不清。现在：
            #   · 认出来的每一张表都画 1px 虚线红框（可选区域，一眼看全）；
            #   · 当前正在看的那张再叠一层 2px 实线红框（和以前一样的语义）；
            #   · 没框的地方就是「这份 PDF 这里没有参数表，或者工具没认出来」——
            #     用户要的「哪里你没识别到」就是这块空白本身。
            # 放大模式不画：整张图就是那张表，再套框只是噪声。
            if whole:
                _draw_table_boxes(page2, meta, tab2)
            # 重画之后选中框要重新画回去，不然刚点的那格一下子找不到了
            _hot_sel()
            try:
                # label 自带「第12页 」，再拼页号会印成「第 12 页 · 第12页 表1」
                _ntab = len([t for t in cat if t.get("page") == page2])
                if not whole:
                    _tl = re.sub(r"^第\s*\d+\s*页\s*", "",
                                 str((tab2 or {}).get("label") or "表"))
                    _mark = "%s（放大）" % _tl
                else:
                    _mark = "%d 表" % _ntab
                if _NP:
                    _pgmark.configure(text="第 %d/%d 页 · %s" % (page2, _NP, _mark))
                else:
                    _pgmark.configure(text="第 %d 页 · %s" % (page2, _mark))
            except Exception:      # noqa: BLE001
                pass

        def _pdf_msg(txt, color=None):
            """画布上放一句说明（读不到 / 缺库 / 还没加载）。"""
            pcv.delete("all")
            S["pimg"] = None
            S["pboxes"] = []
            S["pmeta"] = None
            pcv.create_text(10, 10, anchor="nw", text=txt, width=300,
                            fill=color or theme.TEXT2, font=(theme.FONT, 9),
                            tags="msg")
            pcv.configure(scrollregion=(0, 0, 340, 90))

        # 分隔线一拖 / 窗口一缩，图的宽度就变了 —— 抖一下再按新宽度重渲染
        # （不是每像素一次，那会把拖动拖成幻灯片）。
        _pdf_job = {"id": None}

        def _pdf_later(_e=None):
            try:
                if _pdf_job["id"]:
                    root_w.after_cancel(_pdf_job["id"])
            except Exception:      # noqa: BLE001
                pass
            _pdf_job["id"] = root_w.after(
                260, lambda: _pdf_show(force=True, keep_pos=True))
        pcv.bind("<Configure>", _pdf_later)

        # ---------------- 缩放：中键 / 底下按钮 ----------------
        # 参照「PDF 对比工具」：缩放锚在鼠标位置（滚到哪就放大到哪）。
        # 它那边用滚轮；这里滚轮留给「上下滚」，中键用来放大。
        def _z_text():
            z = float(S.get("zoom") or 1.0)
            try:
                # 一律写百分比（适宽 = 100%）：旁边那颗叫「适宽」，标签再写
                # 「适宽」会变成「适宽 适宽」，看不出到底是不是最大。
                zlab.configure(text="%d%%" % round(z * 100))
            except Exception:      # noqa: BLE001
                pass

        def _z_set(z, anchor=None):
            """切到某个倍率。倍率没变就什么都不做（别白渲染一张图）。"""
            z = float(z)
            if abs(z - float(S.get("zoom") or 1.0)) < 1e-6:
                return
            S["zoom"] = z
            _z_text()
            _pdf_show(force=True, anchor=anchor)

        def _pdf_zoom(step, anchor=None, reset=False):
            """`step=+1` 放大一档、`-1` 缩小一档；`reset=True` 回适宽。

            放大是**按更大像素重渲染**，不是把同一张图拉大 —— 后者只会更糊，
            放大等于白放。
            """
            cur = float(S.get("zoom") or 1.0)
            if reset:
                _z_set(1.0, anchor)
                return
            i = min(range(len(_ZSTEPS)), key=lambda k: abs(_ZSTEPS[k] - cur))
            i = max(0, min(len(_ZSTEPS) - 1, i + step))
            _z_set(_ZSTEPS[i], anchor)

        def _pdf_wheel(e):
            if e.state & 0x0001:                  # Shift
                pcv.xview_scroll(-1 if e.delta > 0 else 1, "units")
            else:
                pcv.yview_scroll(-1 if e.delta > 0 else 1, "units")
            return "break"

        pcv.bind("<MouseWheel>", _pdf_wheel)
        pcv.bind("<Button-2>", lambda e: _pdf_zoom(+1, anchor=(e.x, e.y)))
        pcv.bind("<Shift-Button-2>", lambda e: _pdf_zoom(-1, anchor=(e.x, e.y)))
        # 左键拖 = 平移（放大之后没有滚动条，全靠这个）；中键在这里是缩放，
        # 所以不挂「中键上下拖」。滚轮照样滚 —— 放大到 400% 之后竖向翻更省事。
        _PanGrab(pcv, middle=False, on_click=_pdf_click)
        # Ctrl（Windows / X11）/ Command（macOS）点 = **加选**一格（再点一次取消）。
        # 2026-11-11 用户原话：「用 Ctrl 键或者在 Mac 这边是 Command 键嘛，就可以
        # 进行多个选择，这样子就可以多个表来选」。
        # 单独绑这两个序列是有意的：Tk 对同一个控件只触发**最具体**的那个绑定，
        # 所以按住 Ctrl 点时 `_PanGrab` 的「按下就拦」不会跑，这一下既不会被当成
        # 拖动，也不会在松手时被补成一次「只选一格」的普通点击。
        for _seq in ("<Control-Button-1>", "<Command-Button-1>"):
            try:
                pcv.bind(_seq, lambda e: _pdf_click(e, add=True), add="+")
            except Exception:      # noqa: BLE001 —— 老 Tk 不认 Command
                pass
        pcv.bind("<Motion>", _pdf_motion)

        # ---------------- 「提取结果」预览 ----------------
        def _preview():
            """刷新底下那块。

            值的来源是 `S["grid"][r][c]` —— 判定引擎真正读到的那串字符（原因
            见上面三行控件的注释）。
            """
            sel = S.get("sel")
            _n_sel = len(S.get("sels") or ())
            try:
                # 按钮上的字跟着选中格数走（多选时点下去钉的是 N 格里的并集）
                btn_pin.configure(text=("钉住 %d 格 →" % _n_sel) if _n_sel > 1
                                  else "钉住这格 →")
            except Exception:      # noqa: BLE001
                pass
            if not sel:
                pick_t1.configure(text="还没点格子")
                pick_t2.configure(text="在图上点一下 → 这里显示提取结果",
                                  fg=theme.TEXT2)
                pick_t3.configure(text="")
                return
            page2, tno2, r2, c2 = sel
            tab2 = _tab_of(page2, tno2)
            g = S.get("grid") or []
            val = ""
            if r2 < len(g) and c2 < len(g[r2] or ()):
                val = str(g[r2][c2] if g[r2][c2] is not None else "")
            try:
                # label 本身已经带「第12页 」，这里再拼一遍页号会印成
                # 「第 12 页 · 第12页 表1 · 第 1 行第 1 列」。剥掉前缀，只留表名。
                _lbl = (tab2 or {}).get("label") or "表"
                _lbl = re.sub(r"^第\s*\d+\s*页\s*", "", str(_lbl)) or "表"
                pick_t1.configure(
                    text="第 %d 页 · %s · 第 %d 行第 %d 列" % (
                        page2, _lbl, r2, c2))
            except Exception:      # noqa: BLE001
                pass
            if val.strip():
                # 原文照印，只压掉首尾空白（PDF 里的换行/多空格会让人以为值不对）
                pick_t2.configure(text="提取结果：%s" % val.strip(),
                                  fg=theme.TEXT)
            else:
                pick_t2.configure(text="提取结果：（这一格是空的）",
                                  fg=theme.BAD_FG)
            # 上下文：行首那一格 + 列头。写「行首」不写「行」——
            # 这批 WI 的「Step #」列是垂直居中的合并格，提取出来会落到上一行，
            # 于是「Step 4」那一行的第一个格写着「Step 4」、参数格却写着「Step5+5」。
            # 写「行：Step 4」会让人以为工具把行认错了；写「行首：Step 4」就是
            # 在陈述事实 —— 这一行第一格里印的是它。
            try:
                rl = str(g[r2][0]).strip() if r2 < len(g) and g[r2] else ""
            except Exception:      # noqa: BLE001
                rl = ""
            cl = ""
            try:
                cl = (_col_text(c2, r2) or "").strip()
            except Exception:      # noqa: BLE001
                cl = ""
            # 列头里混着合并格留下的「N/A」占位（三行表头会拼成
            # 「Setting Parameters Temp. (℃) N/A N/A N/A」）—— 显示时剔掉，
            # 但不改 `_col_text()` 本身：那个值要原样存进锚点记录里当凭据。
            if cl:
                _toks = [x for x in cl.split() if x.upper() != "N/A"]
                cl = " ".join(_toks)
            bits = []
            if rl:
                bits.append("行首：%s" % rl[:32])
            if cl:
                bits.append("列：%s" % cl[:48])
            pick_t3.configure(text="　".join(bits))

        # ---------------- 状态行 ----------------
        def _pending_rows():
            """还没结论的行（no_match / no_scope）—— 只用来算「还差几条」。"""
            out = []
            for c in (self.result or {}).get("checks") or []:
                if c.get("status") in ("no_match", "no_scope"):
                    out.append(c)
            return out

        def _rule_reach(rule, g, pend):
            """这条规则能覆盖待钉的哪几条（只算坐标，不看值）。"""
            hr = rule.get("hdr_rows") or []
            stacks = anchors._col_stacks(g, hr)
            out = []
            for ch in pend:
                nm = (ch.get("item") or "").strip()
                if not nm or not anchors._dir_of(nm) or not anchors._site_of(nm):
                    continue
                cc, _ = anchors.pick_col(stacks, nm)
                if cc is None:
                    continue
                rr = anchors.pick_row(g, ch.get("spec"), hr)
                if rr is None:
                    continue
                out.append(ch)
            return out

        def _now_text(chk2):
            """「这一行现在判成什么、读的是哪一格」—— 校正时必须先看得见。"""
            st = _ST_CN.get(chk2.get("status"), chk2.get("status") or "")
            t = "%s · 第 %s 行 · 现在判成「%s」" % (
                (chk2.get("item") or "").strip(), chk2.get("row") or "?", st)
            a = chk2.get("anchor") or {}
            if a:
                # `label` 在锚点通道里就是「第N页 第X行第Y列」这一整串，不能再拼
                # 一遍行列 —— 拼了会印成「第4页 第3行第2列 第3行第2列」。
                t += " · 读 %s「%s」%s" % (
                    a.get("label") or "第%s页 第%s行第%s列" % (
                        a.get("page") or "?", a.get("row") or "?", a.get("col") or "?"),
                    a.get("raw") or "（空）",
                    "（学的）" if a.get("from_rule") else "（你钉的）")
            elif chk2.get("match"):
                m = chk2["match"]
                if m.get("table"):
                    # 找到的是**表**，不是参数。这一行明明「对不上」，却写「按名字对到了
                    # 第4页第1张表」，看的人会以为工具自相矛盾。
                    if chk2.get("status") in ("no_match", "no_scope"):
                        t += " · 名字对到 %s，但表里没这个参数" % m["table"]
                    else:
                        t += " · 名字对到 %s" % m["table"]
            return t

        def _describe():
            chk2 = S.get("chk")
            c, r = S.get("col"), S.get("row")
            g = S.get("grid") or []
            _nsel = len(S.get("sels") or ())
            if chk2 is None:
                hint.configure(text="在图上点那一格（按住 ⌘/Ctrl 可多选）",
                               foreground=theme.TEXT2)
                return
            if r is None or c is None or not g:
                # 选中了行、还没点格子：先说清这一行现在是什么状况，
                # 校正时人才知道自己要改的是哪一条结论。
                hint.configure(text=_now_text(chk2), foreground=theme.TEXT)
                return
            cell = (g[r][c] if r < len(g) and c < len(g[r]) else "")
            name = (chk2.get("item") or "").strip()
            rule = None
            try:
                # S["tab"] 可能是 None（还没在任何格子上点过、按行也没猜出表）——
                # 那时 `anchors.derive` 会抛，直接当成「学不会整张表」处理。
                rule = anchors.derive(S["tab"], g, r, c, name, chk2.get("spec"))
            except Exception:
                rule = None
            S["rule"] = rule
            reach = 0
            if rule:
                try:
                    reach = len(_rule_reach(rule, g, _pending_rows()))
                except Exception:
                    reach = 0
            # 一行说清：点的是哪一格、这格写着什么、钉完能带会多少条。
            txt = "第 %d 行第 %d 列 · 这格写「%s」 · " % (
                r, c, cell or "（空）")
            if _nsel > 1:
                # 多选时先说清「这一条会看好几处」—— 否则底下只写主格，
                # 人会以为副格白点了（钉子下去其实是取并集）。
                txt = "%d 格取并集 · " % _nsel + txt
            _decided = chk2.get("status") not in ("no_match", "no_scope")
            if rule:
                txt += "连带 %d 条同结构的行" % reach
                if _decided:
                    txt += "（本行也重判）"
            elif not anchors._dir_of(name) or not anchors._site_of(name):
                txt += "只钉这一条（名字里没有方向/部位）"
            else:
                txt += "只钉这一条（按名字算的行列对不上，不猜）"
            hint.configure(text=txt,
                           foreground=theme.OK_FG if cell else theme.BAD_FG)

        # ---------------- 换页 / 换行 ----------------
        def _goto(page2, tab2="keep"):
            page2 = max(1, int(page2 or 1))
            S["page"] = page2
            if tab2 != "keep":
                S["tab"] = tab2
            elif S.get("tab") and S["tab"].get("page") != page2:
                # 同一页上有多张表时「当前表」留着（红框还圈着它）；换到别的页就
                # 清掉，否则红框会画到这一页上根本不存在的那张表的位置。
                S["tab"] = None
            _pdf_show(page2, S.get("tab"))
            _hot_sel()
            _preview()
            _describe()

        def _step(d):
            """翻一页。换页之后选中框必须清掉 —— 同一对 (行,列) 在另一页是
            **完全不同的格子**，留着上一页的框会让人以为「工具还记得我选的那格」，
            钉下去就错了。"""
            p = int(S.get("page") or 1) + d
            if p < 1 or (_NP and p > _NP):
                return
            # 多选集合同样要清：留着它，界面上一个框都看不见，按钮却还写着
            # 「钉住 2 格」—— 点下去会钉到屏幕上根本不存在的格子里。
            S["sels"] = []
            S["sel"] = None
            S["row"] = S["col"] = None
            S["tab"] = None
            _goto(p)

        def _default_tab():
            """没有线索时先看哪张表。

            ① 参数行数最多的那张（表目录算出了 `nrec` 时）；
            ② `nrec` 缺失（这份 PDF 没算行数）→ **第一张参数表**；
            ③ 连参数表都认不出 → 第一张。

            为什么不能退回「第一张」：第 1 张往往是「工艺流程图」（一个框一行的
            示意表），对着它看原图对复核毫无帮助。语料实测：WI-0020 表1 是流程图、
            表2 起才是真正的参数表；WI-0028 更典型 —— `nrec` 全是 None，
            于是打开面板停在第 2 页那张流程图，而真正要钉的参数表在第 4 页，
            用户原话：「十几页的 PDF 我要自己逐页去翻」（2026-11-11）。
            """
            best, best_n = None, -1
            for i, c in enumerate(cat):
                n = c.get("nrec")
                if isinstance(n, int) and n > best_n:
                    best, best_n = i, n
            if best_n > 0:
                return best
            for i, c in enumerate(cat):
                if _like_param(c):
                    return i
            return best

        def _like_param(t):
            """这张表像不像「参数表」（而不是流程图 / 物料清单 / 页眉条）。

            2026-11-11 加。起因是「双击一条判错的行，面板停在第 2 页」——
            第 2 页那张表是**工艺流程图**（`工序名称 › 确定MKB…`），它的表头里
            恰好有一个跟「焊接时间」撞上的两字窗口，于是 `_guess_tab` 的粗匹配
            把它排到第一；真正写着 Set Time 的第 4 页参数表反而一个字都没撞上
            （Excel 是中文、PDF 是英文）。人看到的就是「它带我去了错的地方，
            还得自己一页页翻」。
            """
            hs = " ".join(str(h or "") for h in (t.get("headers") or []))
            if not hs:
                return False
            return any(k in hs for k in ("参数名称", "参数值", "参数",
                                         "Parameter", "parameter", "规格"))

        def _guess_tab(chk2):
            """这一行的参数名**最可能**落在哪张表里（返回 cat 下标或 None）。

            只用来把视野预先跳到「这一行所在的那张表」—— 不参与任何判定，
            猜错也只是人自己去翻页。

            线索强弱：① 已经钉过的锚点（人自己指过，最准）→ ② 参数名在 PDF
            正文/表里的位置（`pdf_vertical`）→ ③ 表头做两字窗口粗匹配，
            **参数表优先**（见 `_like_param`）。
            """
            # ① 这一行钉过位置 → 直接去那一页。复核一条老记录时这是唯一不会错的路。
            try:
                _anc = (chk2 or {}).get("anchor") or {}
                _ap, _at = _anc.get("page"), _anc.get("tno")
                if _ap is not None:
                    # 先按 (页, 第几张表) 精确找：同一页常有好几张表（WI-0028 第 4 页
                    # 有 6 张），只按页返回本页第一张会把人带到隔壁那张去 ——
                    # 用户看到的就是「红框圈着锻打那张表，可这一行讲的是焊接」。
                    if _at is not None:
                        for i, c in enumerate(cat):
                            if c.get("page") == _ap and c.get("tno") == _at:
                                return i
                    for i, c in enumerate(cat):
                        if c.get("page") == _ap:
                            return i
            except Exception:      # noqa: BLE001
                pass
            nm = ((chk2 or {}).get("item") or "").strip()
            if not nm:
                return None
            try:
                where = pdf_tables.name_where(getattr(self, "pdf_vertical", None)) or {}
            except Exception:      # noqa: BLE001
                where = {}
            hit = where.get(nm)
            if hit:
                for i, c in enumerate(cat):
                    if (c.get("page"), c.get("tno")) == tuple(hit):
                        return i
            # 退一步：按**表头字段**做一次粗匹配。
            # 为什么需要退这一步：这一行之所以在这里，多半就是因为它的名字
            # （「焗炉温度」）在 PDF 里根本没出现过 —— name_where 里当然查不到。
            # 用两字窗口去撞表头（「温度」），撞得多的是哪张就先看哪张；
            # 撞不上就返回 None，让视野停在原地（宁可不动，也别乱跳）。
            grams = set()
            for k in range(len(nm) - 1):
                g2 = nm[k:k + 2]
                if g2.strip():
                    grams.add(g2)
            if not grams:
                return None
            best_p, hit_p = None, 0        # 参数表里撞得最多的
            best_o, hit_o = None, 0        # 其它表里撞得最多的
            for i, c in enumerate(cat):
                hs = " ".join(str(h or "") for h in (c.get("headers") or []))
                if not hs:
                    continue
                n_hit = sum(1 for g2 in grams if g2 in hs)
                if _like_param(c):
                    # 一条都没撞上时（中文采集项 vs 英文表头，最常见）也要记住
                    # 第一张参数表 —— 否则 best_p 永远是 None，下面那句
                    # 「回到参数表」根本执行不到，面板就停在第 1 张流程图上了。
                    if best_p is None or n_hit > hit_p:
                        best_p, hit_p = i, n_hit
                elif best_o is None or n_hit > hit_o:
                    best_o, hit_o = i, n_hit
            # 非参数表（流程图那种）必须**明显**撞得更多才采信（≥2 个窗口），
            # 否则一律回到参数表 —— 用户要的是「跳到判错那一页」，不是「跳到
            # 一页刚好有同名两个字的地方」。
            if best_p is not None and (hit_p > 0 or hit_o < 2):
                return best_p
            if best_o is not None and hit_o > 0:
                return best_o
            return None

        def _show(chk2):
            """外部（双击某一行）→ 面板：换行、跳到那一行大概在的那一页。"""
            S["chk"] = chk2
            # 换了行，上一行选的那几格就没意义了（它们指的是**另一个采集项**的
            # 位置），和翻页同理：不清就会「看不见框、按钮还写着钉 2 格」。
            S["sels"] = []
            S["sel"] = None
            S["row"] = S["col"] = None
            _hot_sel()
            _preview()
            gi = _guess_tab(chk2) if chk2 else None
            if gi is not None:
                _t = cat[gi]
                S["tab"] = _t
                # 整页视图下**不自动切到放大**：换行时视图一会儿整页一会儿放大，
                # 眼睛会断片。只在整页图上用红框圈出「这一行大概在这张表」。
                if S.get("page") == _t.get("page"):
                    _pdf_show(force=False)
                    _hot_sel()
                    _preview()
                    _describe()
                else:
                    _goto(_t.get("page"), _t)
            else:
                _pdf_show(force=False)
                _hot_sel()
                _preview()
                _describe()

        # ---------------- 钉住 ----------------
        def _pin():
            chk2 = S.get("chk")
            tab = S.get("tab")
            r, c = S.get("row"), S.get("col")
            if chk2 is None:
                hint.configure(text="先在左边双击一条对不上的行", foreground=theme.BAD_FG)
                return
            if tab is None or r is None or c is None:
                hint.configure(text="先在图上点中写着这个数的那一格",
                               foreground=theme.BAD_FG)
                return
            g = S.get("grid") or []
            rl = (g[r][0] if r < len(g) and g[r] else "")
            cell = (g[r][c] if r < len(g) and c < len(g[r]) else "")
            ct = _col_text(c, r)
            name = (chk2.get("item") or "").strip()
            before = len(_pending_rows())
            try:
                nm = anchors._infer_name(tab, r, c, path=pdf) or ""
            except Exception:
                nm = cell or rl
            # 多选：把点中的每一格都记下来（主格仍是上面这一格）。判定那边按
            # 「这几处取并集」用 —— 同一个采集项分写两张表时，只记一格就等于
            # 让人在两者中二选一（用户 2026-11-11 说的正是这件事）。
            _sels = list(S.get("sels") or ())
            if not _sels and S.get("sel"):
                _sels = [S["sel"]]
            _cells = []
            if len(_sels) > 1:
                for _pg, _tn, _rr, _cc in _sels:
                    _t3, _rl3, _v3, _ct3 = _cell_ctx(_pg, _tn, _rr, _cc)
                    if not _t3:
                        continue
                    _cells.append({"item": _t3, "page": _pg, "tno": _tn,
                                   "row": _rr, "col": _cc, "row_label": _rl3,
                                   "col_header": _ct3, "value": _v3})
            ok, msg = anchors.add(
                name, tab, r, c, pdf_name=nm, row_label=rl, col_header=ct,
                value=cell,
                # 带上 Excel 行号 —— 一张卡里同名采集项常有七八行，不带行号会一次
                # 指中它们全部（WI-0039 表4 那 8 行「切定长」实测就踩过）。
                excel_row=chk2.get("row"),
                src=self._src_tag(), cells=_cells)
            if not ok:
                messagebox.showwarning("没法记", msg)
                return
            rule = S.get("rule")
            learned = False
            if rule:
                try:
                    _sr = self._src_tag()
                    _sr.update({"excel": name, "row": chk2.get("row")})
                    learned = anchors.save_rule(rule, src=_sr)
                except Exception:
                    learned = False
            # 钉完立刻重跑：左表自己变短，人亲眼看得见「学会」这件事发生了。
            # 重跑这一下是唯一的卡顿点（整份 PDF 重判），先写一行进度再说 ——
            # 干等没有任何反馈，用起来就像「卡死了」。
            hint.configure(text="正在按你指的格子重判…", foreground=theme.TEXT2)
            root_w.configure(cursor="watch")
            try:
                root_w.update()
                self._no_names_dialog = True
                try:
                    self._run()
                finally:
                    self._no_names_dialog = False
            finally:
                root_w.configure(cursor="")
            # 重跑之后 result 换了一份，S["chk"] 指着旧对象 —— 按 (行号, 采集项)
            # 在新结果里找回同一条，否则底下那句话永远停在「现在判成 对不上」。
            _old_row = chk2.get("row")
            _old_item = (chk2.get("item") or "").strip()
            for c3 in (self.result or {}).get("checks") or []:
                if c3.get("row") == _old_row \
                        and (c3.get("item") or "").strip() == _old_item:
                    S["chk"] = c3
                    break
            # 重判之后这张表可能被重新解析过 → 重新读一遍格子，再按新图重画：
            # 选中框和预览都要跟得上，否则人会以为刚才点的白点了。
            _sel0 = S.get("sel")
            if _sel0:
                try:
                    _load(_sel0[0], _sel0[1])
                except Exception:      # noqa: BLE001
                    pass
            try:
                _pdf_show(force=True)
                _hot_sel()
                _preview()
            except Exception:          # noqa: BLE001
                pass
            after = len(_pending_rows())
            cnt = {}
            for c2 in (self.result or {}).get("checks") or []:
                _st = c2.get("status")
                cnt[_st] = cnt.get(_st, 0) + 1
            _ok = cnt.get("ok", 0) + cnt.get("multi", 0)
            _mm = cnt.get("mismatch", 0)
            parts = ["记下了：%s" % msg]
            if learned:
                parts.append("这张表也学会了，同结构的行一起解决")
            parts.append("待钉 %d → %d 条　全文件：一致 %d / 不一致 %d"
                         % (before, after, _ok, _mm))
            hint.configure(text="　".join(parts), foreground=theme.OK_FG)

        # ---------------- 起手 ----------------
        _z_text()
        P = {"root": root_w, "show": _show, "open": _open,
             "close": _collapse, "pdf": pdf, "state": S}
        self._pin_panel = P
        _open()
        # 先给一页看 —— 打开就看到一片空白（或者一句「先选表」）是最糟的开场，
        # 人连「该点哪」都没法判断。
        _show(chk)
        if S.get("page") is None:
            _di = _default_tab()
            _pg0 = (cat[_di].get("page") if _di is not None
                    else (cat[0].get("page") if cat else 1)) or 1
            S["tab"] = cat[_di] if _di is not None else None
            _goto(_pg0, S.get("tab"))


    def _show_audit(self):
        """点状态栏那行小字 → 打开「结构自检」清单。"""
        a = getattr(self, "audit_result", None)
        if not a:
            return
        errs, notes = a.get("errors") or [], a.get("notes") or []
        if not errs and not notes:
            messagebox.showinfo("结构自检",
                                "全部通过。\n\n"
                                "说明：这不是在验「PDF 写得对不对」，而是在验"
                                "「这次判定出来的结果形不形状正常」——\n"
                                "比如有没有一行判了一致却根本没拿到期望值、"
                                "有没有一行把隔壁表的数当成了本行的答案、"
                                "有没有同一个采集项有的行收窄了有的行没收窄。")
            return

        win = tk.Toplevel(self.root)
        win.configure(bg=theme.BG)
        win.title("结构自检 — %s" % (self.wi_name or "本次比对"))
        win.geometry("1080x520")
        try:
            win.transient(self.root)
        except Exception:
            pass

        head = ("发现 %d 条可疑（红色，建议逐条看）、%d 条提示（灰色，多数是正常情形）"
                % (len(errs), len(notes)))
        ttk.Label(win, text=head, anchor="w",
                  foreground=theme.DANGER if errs else theme.TEXT2).pack(
            fill="x", padx=10, pady=(8, 4))

        cols = ("lv", "kind", "item", "spec", "note")
        tv = ttk.Treeview(win, columns=cols, show="headings", height=16)
        # kind 要放得下最长的那句「期望值等于直径（疑似串到隔壁表）」= 16 个汉字，
        # 小四号下约 248px；留 4px 余量给 Treeview 自身的左右内边距。
        for cid, cap, w in (("lv", "级别", 52), ("kind", "这一类", 252),
                            ("item", "采集项", 196), ("spec", "适用规格", 114),
                            ("note", "说明", 430)):
            tv.heading(cid, text=cap)
            tv.column(cid, width=w, anchor="w")
        tv.tag_configure("err", foreground=theme.DANGER)
        tv.tag_configure("note", foreground=theme.TEXT2)
        for r in errs:
            tv.insert("", "end", values=("可疑", r["kind"], r["item"],
                                         r["spec"], r["note"]), tags=("err",))
        for r in notes:
            tv.insert("", "end", values=("提示", r["kind"], r["item"],
                                         r["spec"], r["note"]), tags=("note",))
        _PanDrag(tv)                       # 没滚动条：按住左键拖（第九轮统一）

        foot = ("这些不是说「你的文件错了」，而是「这次判定出来的形状不寻常」，"
                "需要你瞄一眼。自检不改任何判定结论。")
        ttk.Label(win, text=foot, foreground=theme.TEXT2, wraplength=1060,
                  justify="left").pack(side="bottom", fill="x", padx=10, pady=(0, 6))
        tv.pack(side="left", fill="both", expand=True, padx=10, pady=(0, 10))
