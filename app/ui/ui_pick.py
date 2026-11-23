"""两个文件池：加/删/清、选中、拖放、计数。

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
import tkinter as tk
from tkinter import filedialog, ttk

import ui_theme as theme
from ui_state import _pick_dir
from ui_fmt import _fit, _short_names, _walk_widgets, drop_paths
from ui_pan import _PanDrag


# ── 「把文件拖进来」的可选能力（与 app.py 原来那段同一套降级写法）────────
try:
    import tkinterdnd2 as _dnd
    from tkinterdnd2.TkinterDnD import require as _dnd_require
    DND_FILES = _dnd.DND_FILES
except Exception:       # 没装就退化为「无拖放」，「＋ 添加文件…」照常
    _dnd = None
    _dnd_require = None
    DND_FILES = None


class PickMixin:
    """两个文件池：加/删/清、选中、拖放、计数。"""


    # ------------------------------------------------------- 清单（一次挑一份）
    def _build_pick_head(self, parent, kind, label, where="left"):
        """一栏的表头：标题 + 份数，左右**镜面对称**（2026-10-29 用户要求）。

        用户原话：「那个 0 份放在那个蓝色框线的右边一点，也就是让它形成一个
        镜面对称的效果」「那个 WI 作业指导书 PDF 也是可以放在最右侧」。

        所以两边都是「份数贴中间那根线、标题贴大框外沿」：
          左：[工程参数表 Excel] …………………… [4 份]
          右：[0 份] …………………… [WI 作业指导书 PDF]
        """
        head = ttk.Frame(parent, style="Card.TFrame")
        # 表头坐在白色卡片上，所以用 Card.* 样式（普通 TLabel 的底是窗口灰，
        # 贴在卡片上会出现一条灰带，看着像没对齐）。
        title = ttk.Label(head, text=label, style="Card.TLabel",
                          font=(theme.FONT, theme.F_BODY, "bold"))
        cnt = ttk.Label(head, text="0 份", style="CardMuted.TLabel")
        if where == "left":
            title.pack(side="left", padx=(2, 0))
            cnt.pack(side="right", padx=(0, 7))     # 离细线留一点，别贴着
        else:
            cnt.pack(side="left", padx=(7, 0))      # 刚从细线右边出来
            title.pack(side="right", padx=(0, 2))
        if kind == "xlsx":
            self.cnt_xlsx = cnt
        else:
            self.cnt_pdf = cnt
        return head



    def _build_pick_list(self, parent, kind, where="left"):
        """一栏的清单（含竖条），按格摆进那个大框。

        `where` 决定竖条贴哪边：'left' 栏的竖条在清单**左边**（也就是整个
        大框的最左沿），'right' 栏的在清单**右边**（大框最右沿）——
        2026-10-29 用户原话：「左边上下滑动的条放在最左侧，右边那个框上下
        滑动的条放在最右侧，左右滑动条分别放在他们的底下。」

        清单是**单选**的 —— 一栏放几十份备着，每次只挑一份，两边各一份对着比。
        单选不是能力退化，是把「比什么」从一堆组合里收敛成一条明确的对应关系：
        这份参数卡 对 这份 WI，出了错能指名道姓，报告也不用再按 WI 汇总。

        2026-10-29 第三次改：三个按钮从这里**搬出去**了（见 `_build_pick_btns`）。
        它们原来在栏内、跟着细线一起被挤，用户原话：「怎么底下的那个添加文件、
        移除清单、还有清空，这个按钮你也添加了这个效果吧？我觉得把这个三个
        按钮在 PDF 那边给它固定死在最右侧就行了。」现在它们钉死在大框外沿。
        """
        # ⚠️ 清单**不能**保留自己的凹陷边框（Listbox 默认 relief=sunken +
        #    borderwidth=1）：那会画成「两个框」，用户要的是「一个大框中间
        #    一根细线」。压平之后，唯一的分隔就是那根线；文字也被裁在线的
        #    边缘，正好是「左边的字被细线覆盖、右边从线的端点开始」。
        #
        # ⚠️ width=1 是关键：Listbox 默认按 20 个字符要宽度，请求宽度会顶住
        #    grid，中间那根线就拖不到比它更窄的地方。把请求宽度压到最小，
        #    实际宽度交给 grid 的 minsize / weight 决定。
        #
        # height 6→5：用户连着三轮都在要「底下的面积再大一点」，这一轮把
        # 「概览 / 明细」、按钮行整块往上提，清单再让一行是最直接的一刀。
        lb = tk.Listbox(parent, height=theme.LIST_ROWS, width=1,
                        selectmode="browse",
                        activestyle="none", exportselection=False,
                        font=(theme.FONT, theme.F_SMALL), relief="flat", borderwidth=0,
                        highlightthickness=0, bg=theme.CARD, fg=theme.TEXT,
                        # 选中的那一条就是「本次要比的」。
                        # 2026-11-04 第九轮把底色调淡一档（TINT_BAND → TINT_HEAD）：
                        # 主页上一左一右两块都铺满饱和蓝太抢眼（用户：「颜色采用
                        # 地点有点多了，有点晃眼」），现在是「一眼看得出选中，又不
                        # 跟旁边的红线蓝带打架」。
                        selectbackground=theme.TINT_HEAD,
                        selectforeground=theme.TEXT)
        # 滑条 2026-11-04 第九轮全部撤掉（用户：「对于所有的有滑条的地方，都可以
        # 把那个滑条进行删除」）—— 按住左键拖就能看，上下左右都行。
        # 清单自己的按下行为（选中那一份）必须留着，所以用 _PanDrag 而不是 _PanGrab：
        # 它只在「拖」的那一段拦一手，免得拖动时一路刷选中。
        _PanDrag(lb)
        if where == "left":
            lb.grid(row=1, column=self._COL_LIST_L, sticky="nsew")
        else:
            lb.grid(row=1, column=self._COL_LIST_R, sticky="nsew")

        # 用 <<ListboxSelect>> 而不是鼠标事件：点击、键盘上下、Home/End
        # —— 凡是能改变选择的动作都会走到这里，一套逻辑全覆盖。
        lb.bind("<<ListboxSelect>>", lambda _e, k=kind: self._on_pick_select(k))
        # 清单**自己**变宽变窄时刷新标题那句份数（见 `_sync_counts`：
        # 窄了只报「N 份」，「· 已选 1」会被边缘切掉半个字）。
        # 不能只挂在大框的 Configure 上：首帧那次大框的 Configure 早于清单
        # 摆好宽度，量到的还是 1px，之后就没有事件了。
        lb.bind("<Configure>", lambda _e, k=kind: self._sync_counts())

        if kind == "xlsx":
            self.lst_xlsx = lb
        else:
            self.lst_pdf = lb
        self._refresh_pick_list(kind)



    def _build_pick_btns(self, parent):
        """两排按钮：左组钉死大框最左、右组钉死大框最右（2026-10-29）。

        用户原话：「我觉得把这个三个按钮在 PDF 那边给它固定死在最右侧就行了。」
        做法是把两组塞进同一个跨全宽的格子里，一个 side="left"、一个
        side="right" —— 它们跟中间那根细线**完全没有关系**了：
        细线拖到哪儿，这两组按钮都还在原来的位置。
        （之前它们在栏内，栏一窄就被挤成「⋯」菜单，那正是用户不满的地方。）
        """
        row = ttk.Frame(parent, style="Card.TFrame")
        self.btn_groups = {}
        for kind, side in (("xlsx", "left"), ("pdf", "right")):
            grp = ttk.Frame(row, style="Card.TFrame")
            grp.pack(side=side)
            ttk.Button(grp, text="＋ 添加文件…", style="Small.TButton",
                       command=lambda k=kind: self._pick_many(k)).pack(side="left")
            ttk.Button(grp, text="移出清单", style="Small.TButton",
                       command=lambda k=kind: self._remove_selected(k)
                       ).pack(side="left", padx=4)
            ttk.Button(grp, text="清空", style="Small.TButton",
                       command=lambda k=kind: self._clear_list(k)).pack(side="left")
            self.btn_groups[kind] = grp
        return row



    def _sel_of(self, kind):
        """本次选中的那一份（最多一条）；没选返回空列表。"""
        lst = self._list_of(kind)
        sel = self.sel_xlsx if kind == "xlsx" else self.sel_pdf
        return [sel[0]] if sel and sel[0] in lst else []



    def _set_sel(self, kind, paths):
        if kind == "xlsx":
            self.sel_xlsx = list(paths)[:1]
        else:
            self.sel_pdf = list(paths)[:1]



    def _label_of(self, kind):
        return "Excel" if kind == "xlsx" else "WI"



    def _paint_sel(self, kind):
        """把 `sel_*` 画成清单里的蓝色高亮（程序化改选中时用）。"""
        lb, lst = self._listbox_of(kind), self._list_of(kind)
        sel = self._sel_of(kind)
        self._syncing_sel = True
        try:
            lb.selection_clear(0, "end")
            if sel:
                i = lst.index(sel[0])
                lb.selection_set(i)
                lb.see(i)
        finally:
            self._syncing_sel = False



    def _auto_pick_first(self, kind):
        """清单里还没选中时，自动把第一份选上（2026-10-29 用户要求）。

        用户原话：「拖进来那一刻，默认选择第一个文件当做那个默认对比项……
        这样子就不用每次拖进来一个文件就比选一次……如果 PDF 多个的话，
        那它就默认当选第一个。」

        **只补空、不抢**：已经有选中的话原样不动 —— 连拖几份时把刚点好的
        那份顶掉，比多重一次点击更烦人。所以「拖进来就自动选上」这个体感
        只在「还没有任何选中」时发生，正好是用户抱怨的那一步。
        """
        lst = self._list_of(kind)
        if not lst or self._sel_of(kind):
            return False
        self._set_sel(kind, [lst[0]])
        self._paint_sel(kind)
        self._sync_counts()
        # ⚠️ 已经比过一次就不动 `*_var`：这两个变量记的是「手上这份结论是
        #    拿哪两份文件算出来的」，导出报告 / 交接包都按它写路径。
        #    比完再拖一份新的进来自动选中，只是换个下次要比的对象 ——
        #    这时把 var 改掉，导出的旧结论就会挂上新文件名（张冠李戴）。
        #    新选中那份会在点「开始比对」时（_run 里）才写进 var。
        if self.result:
            return True
        var = self.xlsx_var if kind == "xlsx" else self.pdf_var
        if var.get() != lst[0]:
            var.set(lst[0])
        return True



    def _on_pick_select(self, kind):
        """清单里选中的那一条 → 写回 `sel_*`。"""
        if getattr(self, "_syncing_sel", False):
            return
        lb, lst = self._listbox_of(kind), self._list_of(kind)
        cur = lb.curselection()
        picked = [lst[cur[0]]] if cur and 0 <= cur[0] < len(lst) else []
        self._set_sel(kind, picked)
        self._sync_counts()
        if not picked:
            return
        # 清单里只有文件名，这里把完整路径报出来 —— 状态栏在最底下占了
        # 一整行，放得下完整路径，不必再截短。
        self.status.set("已选 %s：%s" % (self._label_of(kind), _fit(picked[0], 150)))
        if kind == "xlsx" and self.xlsx_var.get() != picked[0]:
            self.xlsx_var.set(picked[0])
            self._load_sheets(restore=None)



    def _sync_counts(self):
        """刷新两栏标题右侧的份数（含「已选 1」）。"""
        for kind, cnt in (("xlsx", getattr(self, "cnt_xlsx", None)),
                          ("pdf", getattr(self, "cnt_pdf", None))):
            if cnt is None:
                continue
            try:
                n_pool = len(self._list_of(kind))
                has = bool(self._sel_of(kind))
                # 栏被那根细线拖窄之后，「N 份 · 已选 1」会顶到右边缘被切掉
                # 半个字。窄了就只说份数 —— 选没选中看蓝色那一条更直接。
                narrow = self._listbox_of(kind).winfo_width() < 290
                if has and not narrow:
                    text = "%d 份 · 已选 1" % n_pool
                else:
                    text = "%d 份" % n_pool
                cnt.configure(text=text,
                              foreground=theme.TINT_BAND_FG if has else theme.TEXT2)
            except Exception:  # noqa: BLE001 —— 控件正在销毁
                pass



    def _list_of(self, kind):
        return self.xlsx_list if kind == "xlsx" else self.pdf_list



    def _listbox_of(self, kind):
        return self.lst_xlsx if kind == "xlsx" else self.lst_pdf



    def _refresh_pick_list(self, kind):
        """把清单灌进对应的 Listbox，并同步单份变量（pdf_var / xlsx_var 指向第一项）。

        ⚠️ 清单里**只显示文件名**（2026-10-28 用户要求），完整路径仍留在
           `self.pdf_list` / `self.xlsx_list` 里。所以要严格保证
           「Listbox 第 i 行 ⟷ 列表第 i 项」这个一一对应关系：
           取选中项一律用 `curselection()` 的下标回去索引列表，
           **绝不能**反过来拿 Listbox 里显示的文字去当路径用。
        """
        lst = self._list_of(kind)
        lb = self._listbox_of(kind)
        lb.delete(0, "end")
        for s in _short_names(lst):
            lb.insert("end", s)
        lb.xview_moveto(0)          # 重排后把横向滚动条归零，否则新名字可能是半截
        # 清单变了 → 选中项跟着收拾：已经不在清单里的那份要剔掉
        # （否则删了文件还挂着一条幽灵选择）。
        self._set_sel(kind, [p for p in self._sel_of(kind) if p in lst])
        self._paint_sel(kind)       # 重排后把蓝色高亮照着 sel_* 重新画一遍
        self._sync_counts()
        # 同步单份变量：交接包路径、单份比对都读它。
        # 三种情形：
        #   · 选中了某一份 → 指向它；
        #   · **已经比过一次**（self.result 有东西）却什么都没选 → 一个字都不动。
        #     这时 pdf_var / xlsx_var 记的是「这份结论是拿哪两个文件算出来的」，
        #     清单里随手删几行就把它改掉，导出的报告和 JSON 会张冠李戴；
        #   · 还没比过 → 退回清单第一项（启动时要靠它去读工作表，不能清空）。
        sel_now = self._sel_of(kind)
        if sel_now:
            head = sel_now[0]
        elif self.result:
            head = None                      # None = 本次不改
        else:
            head = lst[0] if lst else ""
        if head is not None:
            if kind == "xlsx":
                self.xlsx_var.set(head)
            else:
                self.pdf_var.set(head)



    def _add_paths(self, kind, paths):
        """往清单追加路径（同一路径加两次直接跳过），并写盘记忆。

        返回真正加进去的条数（`0` = 全是重复项）。拖入处拿它来告诉用户
        「加了几个、跳过几个」，免得拖完看不出到底有没有反应。
        """
        lst = self._list_of(kind)
        added = 0
        for p in paths:
            p = str(p).strip()
            if p and p not in lst:
                lst.append(p)
                added += 1
        if not added:
            return 0
        self._refresh_pick_list(kind)
        self._save_last()
        # 拖进来 / 加进来就把第一份选上，省掉那一次点击（2026-10-29）。
        self._auto_pick_first(kind)
        return added



    def _pick_many(self, kind):
        """「＋ 添加文件…」—— 一次可多选（按住 Cmd / Ctrl，或框选）。"""
        if kind == "xlsx":
            title = "添加工程参数表（可按住 Cmd / Ctrl 多选）"
            types = [("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        else:
            title = "添加 WI 作业指导书（可按住 Cmd / Ctrl 多选）"
            types = [("PDF 文件", "*.pdf"), ("所有文件", "*.*")]
        cur = self._list_of(kind)
        start = _pick_dir(*(os.path.dirname(p) for p in cur),
                          os.path.expanduser("~/Desktop"))
        got = filedialog.askopenfilenames(title=title, initialdir=start, filetypes=types)
        if got:
            self._add_paths(kind, list(got))
            if kind == "xlsx" and self.xlsx_list:
                # 换了 Excel 文件，重新列工作表（清单第一项是新的那批里第一个）
                self._load_sheets(restore=None)



    def _absorb_paths(self, paths):
        """把一批路径按后缀分进左右两栏 —— 拖放的唯一入口。

        原来还有第二个调用方（中间那个「＋ 添加文件…」按钮）。2026-10-29
        用户把中间框拆掉了，只剩拖放这一条路，但分流逻辑照旧只留一份：
        分错一次就是「Excel 跑到右栏、怎么比都比不出东西」，不能有两份实现。

        返回 dict：新增数、原始数（用来算「重复几个」）、文件夹、不认的。
        """
        xl, pd, skip, dirs = [], [], [], []
        for p in paths:
            ext = os.path.splitext(str(p))[1].lower()
            if ext in (".xlsx", ".xls", ".xlsm"):
                xl.append(p)
            elif ext == ".pdf":
                pd.append(p)
            elif os.path.isdir(p):
                # 拖文件夹是常见误操作（对着文件夹整个往里拖）。不收是对的
                # —— 收下去只会把一堆无关 PDF 灌进清单；但必须**说清楚**，
                # 否则用户看到的就是「拖了没反应」。
                dirs.append(os.path.basename(str(p).rstrip("\\/")) or p)
            else:
                skip.append(os.path.basename(str(p)) or p)
        n_xl = self._add_paths("xlsx", xl) or 0
        if n_xl:
            self._load_sheets(restore=None)
        n_pd = self._add_paths("pdf", pd) or 0
        return {"xl": n_xl, "xl_all": len(xl), "pd": n_pd, "pd_all": len(pd),
                "dirs": dirs, "skip": skip}



    def _remove_selected(self, kind):
        """把清单里选中的项从清单里去掉。

        2026-10-29 用户的话：「你说不会清除本地文件，这个也是没必要的。
        因为你要必须保证你这个工具就是检测嘛，不会动任何原本的文件……
        反而会给人增加一种信任危机感。」
        —— 所以这里既不再弹确认框，也不再说「不会删你的文件」：
        这一栏本来就是个待办列表，删的只是列表里的一行，没什么好保证的。
        """
        lb = self._listbox_of(kind)
        lst = self._list_of(kind)
        sel = sorted(lb.curselection(), reverse=True)
        if not sel:
            self.status.set("先在清单里点中要移出的那份（一次一份，点了就蓝）")
            return
        for i in sel:
            if 0 <= i < len(lst):
                del lst[i]
        self._refresh_pick_list(kind)
        self._save_last()



    def _clear_list(self, kind):
        """清空整张清单（不删文件本身）。"""
        if not self._list_of(kind):
            return
        if kind == "xlsx":
            self.xlsx_list = []
        else:
            self.pdf_list = []
        self._refresh_pick_list(kind)
        self._save_last()



    # ── 拖文件进来（2026-10-13 用户要求「一口气拖多个文件」）──────────────
    def _setup_dnd(self) -> bool:
        """把「接拖入」挂到窗口上，挂不上返回 False（界面照常，只是没这功能）。

        ⚠️ 2026-10-28 改成**逐个控件注册**，不再只挂 `self.host` 一层。

        为什么：Windows 上的拖放走 OLE 的 RegisterDragDrop，它认的是**鼠标底下
        那个 HWND**。tkdnd 自己会沿父链往上找已登记的祖先（`FindWindowWithCommonTypes`），
        理论上只挂 host 就够 —— 但这条路在「注册的是 Frame、鼠标底下是
        Listbox / Treeview」时靠的是 OLE 自己的祖先回溯，实测并不可靠
        （用户反映「拖不进去」，而且完全没有提示，比报错更难查）。
        逐个控件登记是笨办法，但每个可能被拖到的 HWND 都确实登记了，
        不依赖任何一层回溯行为，坏不了。
        """
        if _dnd_require is None:
            return False
        try:
            # 先给窗口补上拖放能力（require 只该调一次；失败只是没这功能）
            _dnd_require(self.root)
        except Exception:  # noqa: BLE001
            return False
        targets = [self.host] + _walk_widgets(self.host)
        ok = 0
        for w in targets:
            if not hasattr(w, "drop_target_register"):
                continue
            try:
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop)
                ok += 1
            except Exception:  # noqa: BLE001 —— 单个控件挂不上不影响其它
                continue
        return ok > 0



    def _on_drop(self, e):
        """松手了 —— 按扩展名把落下的文件分进 Excel / PDF 两个清单。

        拖进来的一堆里既有 .xlsx 又有 .pdf，自动按后缀归类：Excel 进第一行、
        右栏；PDF 进右栏、Excel 进左栏。都不认的（如 .docx）静默忽略，不塞脏数据。
        """
        # ⚠️ 必须用 Tcl 的 splitlist，而且要从 **root** 上取：
        #    `self` 是 App（普通 Python 对象）不是控件，`self.tk` 根本不存在 ——
        #    原来这里 getattr 拿不到就退回 `str.split()`，而 tkdnd 送来的
        #    是一个 Tcl list，含空格 / 逗号的路径会被整条套上花括号，
        #    硬切直接切成两条不存在的路径（本项目的文件名几乎全带空格和逗号）。
        splitlist = getattr(self.root, "tk", None)
        splitlist = getattr(splitlist, "splitlist", None)
        paths = drop_paths(getattr(e, "data", ""), splitlist)
        if not paths:
            self.status.set("拖进来的是空的，没有加入任何文件")
            return "copy"
        # 分类与入清单统一走 `_absorb_paths`（中间那个「＋ 添加文件…」
        # 用的是同一个入口，两边不会分得不一致）。
        st = self._absorb_paths(paths)
        n_xl, n_pd = st["xl"], st["pd"]
        dirs, skip = st["dirs"], st["skip"]
        # 拖完必须有回声：加了几个、重复几个、哪些不认 —— 都不说，
        # 用户就只能靠猜，猜不出来就会以为「这功能没做」。
        msg = []
        if n_xl:
            msg.append("Excel %d 个" % n_xl)
        if n_pd:
            msg.append("PDF %d 个" % n_pd)
        dup = (st["xl_all"] - n_xl) + (st["pd_all"] - n_pd)
        if dup:
            msg.append("已在清单里的 %d 个（跳过）" % dup)
        if dirs:
            msg.append("文件夹 %d 个（不支持拖文件夹，请打开它再拖里面的文件）：%s"
                       % (len(dirs), "、".join(dirs[:2])))
        if skip:
            msg.append("不认识的 %d 个（只收 .pdf / .xlsx）：%s"
                       % (len(skip), "、".join(skip[:3]) + ("…" if len(skip) > 3 else "")))
        if msg:
            self.status.set("已拖入：" + "；".join(msg))
            # 2026-10-29 起拖进来就自动把第一份选上，所以这里只剩「缺一边」的
            # 情况需要点一句（比如只拖了 PDF，左栏还是空的）。两边都有了就不再
            # 啰嗦 —— 选中是自动的，用不着教他手工点。
            lack = []
            if not self._sel_of("xlsx"):
                lack.append("Excel（左栏）")
            if not self._sel_of("pdf"):
                lack.append("WI（右栏）")
            if lack:
                self.status.set(self.status.get() + "　→ 还差%s" % "、".join(lack))
        else:
            self.status.set("拖进来的文件都没加入清单（可能都在清单里了）")
        return "copy"



    # ── 拖放挂不上时的一句话说明（不再有常驻提示条）──────────────────────
    # 2026-10-29 用户原话：「中间有一行字的蓝色提示弹窗，这个是没有必要的。」
    # 说得对：拖入是个动作，拖一次就知道能用；为它常驻一条横幅，等于天天在
    # 教一个已经会用的操作，还把界面往「说明文档」的方向带。
    # 所以只保留**坏消息**：真挂不上时在状态栏说一句、并指路按钮，不占版面。
    def _dnd_note(self):
        """拖放不可用时在状态栏说明原因；可用就什么也不说。"""
        if getattr(self, "dnd_ok", False):
            return
        if _dnd is None:
            self.status.set("拖入未启用（缺拖放组件）—— 用「＋ 添加文件…」同样可以多选")
        else:
            self.status.set("拖入未启用（拖放组件与当前环境不匹配）"
                            "—— 用「＋ 添加文件…」同样可以多选")
