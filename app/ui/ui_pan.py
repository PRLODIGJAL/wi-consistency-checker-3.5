# -*- coding: utf-8 -*-
"""ui_pan.py —— 把「滚动条」换成「按住左键拖」的一套小手艺。

2026-11-04 第八、九轮，用户拿了同机的「PDF 对比工具 6.1」当参照：
「它对于文件以及任何可以有位移的格式内，它具有拖拽机制……所有地方的滑条
都采用左键，然后按住然后左右拖拽」。所以这个程序里**一条滚动条都不留**，
凡是能滚的地方都按住左键拖。

两个类，区别只在「按下那一下要不要拦住」：

  · `_PanGrab`  —— 按下就 `return "break"`，把控件自己的按下行为挡掉，
    松手时如果几乎没动，再调 `on_click(event)` 把「点一下」补回来。
    用在「按下会引发换内容」的地方：指位置窗的 ① 行表、② 表目录、③ 格子。

  · `_PanDrag`  —— 按下**不拦**，让控件自己先处理（Treeview 选中行、
    Listbox 选中项、分组条折叠都照旧），只拦「拖」那一段（`<B1-Motion>`），
    免得拖动过程中把选中一路刷下去。用在主界面明细表和主页文件清单。

⚠️ 两个通用坑：

  ① **事件要合并**。鼠标/触控板的移动事件比「重绘一帧」密得多（一台
     1000Hz 的鼠标，1ms 一个事件；重绘一张 500 格的表要 11ms）。不合并的话
     事件队列越堆越长，拖起来是「越拖越飘」的橡皮筋手感 —— 用户说的
     「卡卡的」正是这个。所以这里只**累加**增量，每 8ms 落一次账
     （`_FRAME_MS`），一帧只重绘一次。
  ② **平移量按分数算**，不按 units：`yview_scroll(3, "units")` 在 Canvas 是
     3 像素、在 Treeview 是 3 行，同一个数字在两处速度差十倍。

2026-11-04 第九轮收盘，用户把这套动作定成了**通用规矩**（原话）：

  「对于那种可拖动的界面来说，以后也是尽量直接用鼠标中键上下滚动。然后鼠标
   直接左右拖动，不需要加滑动条；而对于那种中间是用来放大的地方，那就直接
   鼠标左键拽住，上下左右拖动也不需要加滑动条。这个时候同时为了方便调试，
   不只（是）用鼠标中键可以放大，还要加一个单独的按钮用来放大缩小。」

所以现在每个可拖区域同时具备三件事，见 `bind_wheel`：

  · 滚轮 / 中键上下滚 —— 竖向；Shift + 滚轮横着走；
  · 按住左键拖 —— 上下左右都行（`_Pan*` 提供）；
  · 按住中键上下拖 —— 给没有滚轮的鼠标和触控板兜底（放大区除外，那里
    中键是缩放）。
"""

FRAME_MS = 8            # 一帧的合并窗口；再小就白重绘，再大能感觉到延迟
THRESH = 4              # 超过这么多像素才算「拖」，否则算「点」
WHEEL_PX = 44           # 滚轮一格走多少像素（跟一行文字的高度差不多）


def pan_px(w, dx, dy):
    """按像素增量平移任意「有 view 的」控件：Canvas / Treeview / Listbox 通吃。

    为什么不直接 `yview_scroll(dy, "units")`：单位对 Canvas 是像素、对 Treeview
    是行、对 Listbox 是项，同一个数字在三处走的速度差十倍。分数换算跟控件类型
    无关，滚轮和拖动就能共用一套手感。
    """
    for axis, d in (("x", dx), ("y", dy)):
        if not d:
            continue
        try:
            f, l = (w.xview() if axis == "x" else w.yview())
        except Exception:                  # noqa: BLE001
            continue
        span = float(l) - float(f)
        if span <= 1e-6 or span >= 0.999999:
            continue                       # 这个方向没得拖（内容比视口小）
        size = w.winfo_width() if axis == "x" else w.winfo_height()
        total = max(1.0, float(size) / span)
        nf = max(0.0, min(1.0 - span, float(f) - d / total))
        try:
            (w.xview_moveto if axis == "x" else w.yview_moveto)(nf)
        except Exception:                  # noqa: BLE001
            pass


class _PanBase:
    """拖动的公共部分：累计增量、按帧落账、光标、分数换算。"""

    THRESH = THRESH

    def __init__(self, w):
        self.w = w
        self._p0 = None
        self._moved = False
        self._pend = [0.0, 0.0]
        self._job = None

    # ---- 事件 ----
    def _start(self, e):
        self._p0 = (e.x, e.y)
        self._moved = False
        self._pend = [0.0, 0.0]
        self._cursor("fleur")

    def _motion(self, e):
        if self._p0 is None:
            return None
        x0, y0 = self._p0
        dx, dy = e.x - x0, e.y - y0
        if not self._moved and abs(dx) < self.THRESH and abs(dy) < self.THRESH:
            return "break"                 # 还在「点」的抖动范围内，先不动
        self._moved = True
        self._p0 = (e.x, e.y)
        self._pend[0] += dx
        self._pend[1] += dy
        self._schedule()
        return "break"

    def _end(self):
        p0, self._p0 = self._p0, None
        self._flush()                      # 松手前把最后那点增量补上
        self._cursor("")
        moved, self._moved = self._moved, False
        return p0, moved

    # ---- 合并 ----
    def _schedule(self):
        if self._job is not None:
            return
        try:
            self._job = self.w.after(FRAME_MS, self._tick)
        except Exception:                  # noqa: BLE001 —— 控件没了就当场做完
            self._tick()

    def _tick(self):
        self._job = None
        self._flush()

    def _flush(self):
        dx, dy = self._pend
        self._pend = [0.0, 0.0]
        if not dx and not dy:
            return
        try:
            if self._job is not None:
                self.w.after_cancel(self._job)
        except Exception:                  # noqa: BLE001
            pass
        self._job = None
        self._pan(int(round(dx)), int(round(dy)))

    # ---- 动作 ----
    def _cursor(self, cur):
        try:
            self.w.configure(cursor=cur)
        except Exception:                  # noqa: BLE001
            pass

    def _pan(self, dx, dy):
        """按像素增量平移（换算逻辑见模块级 `pan_px`）。"""
        pan_px(self.w, dx, dy)


class _PanGrab(_PanBase):
    """按下就拦（控件自己的按下行为不要了），松手补一次 `on_click`。

    为什么要挡：ttk 的行选中是在 `<ButtonPress-1>` 的类绑定里做的 —— 不挡的话
    「想拖」这一下会先顺手把光标下那行选了，指位置窗的 ① 里一拖就跳表、③ 整块
    重画，看着像抽风。挡掉之后由 `on_click(event)` 在「按下到松开几乎没动」时
    补一次选中，效果和原来单击完全一样。

    ⚠️ 列宽分隔线要放行（`identify_region == "separator"`）—— 拖列宽是
    Treeview 自带的手艺，抢过来就变成平移了。
    """

    def __init__(self, w, on_click=None, wheel=True, middle=True):
        _PanBase.__init__(self, w)
        self.on_click = on_click
        w.bind("<ButtonPress-1>", self._press, add="+")
        w.bind("<B1-Motion>", self._motion, add="+")
        w.bind("<ButtonRelease-1>", self._release, add="+")
        if wheel:
            bind_wheel(w, middle=middle)

    def _press(self, e):
        try:
            if (self.w.winfo_class() == "Treeview"
                    and self.w.identify_region(e.x, e.y) == "separator"):
                # 2026-11-04 第十轮：列宽改成「按比例跟着容器算」，于是分隔线
                # 不给拖了 —— 用户原话：「给它固定成某一个宽度之后就给它固定死
                # 在那里，让它不要拖动嘛」。这里 `return "break"` 把 ttk 自己的
                # 「按下分隔线 → 进入拖列宽状态」也挡掉（只 return None 的话，
                # 列宽还是拖得动，而且顺带把这一下当成平移）。
                self._p0 = None
                return "break"
        except Exception:                  # noqa: BLE001
            pass
        self._start(e)
        return "break"

    def _motion(self, e):
        if self._p0 is None:
            return "break"                 # 分隔线上按下的那一下：什么都不做
        return _PanBase._motion(self, e)

    def _release(self, e):
        p0, moved = self._end()
        if p0 is None:
            return None
        if not moved and self.on_click is not None:
            try:
                self.on_click(e)
            except Exception:              # noqa: BLE001
                pass
        return "break"


class _PanDrag(_PanBase):
    """按下**不拦**（控件的选中/折叠照旧），只拦「拖」那一段。

    Listbox 的 `<B1-Motion>` 类绑定会一路刷选中区、Treeview 会拖选/自动滚 ——
    这两样在「想拖画面」的时候都是帮倒忙，所以拖动时直接 `return "break"`。
    """

    def __init__(self, w, on_click=None, wheel=True, middle=True):
        _PanBase.__init__(self, w)
        self.on_click = on_click
        w.bind("<ButtonPress-1>", self._press, add="+")
        w.bind("<B1-Motion>", self._motion, add="+")
        w.bind("<ButtonRelease-1>", self._release, add="+")
        if wheel:
            bind_wheel(w, middle=middle)

    def _press(self, e):
        self._start(e)
        return None                        # 不拦：选中/折叠交给控件自己

    def _release(self, e):
        p0, moved = self._end()
        if p0 is not None and not moved and self.on_click is not None:
            try:
                self.on_click(e)
            except Exception:              # noqa: BLE001
                pass
        return None


class _MiddlePan(_PanBase):
    """按住中键上下拖 = 竖向平移（只走竖的，横的留给左键）。"""

    def __init__(self, w):
        _PanBase.__init__(self, w)
        w.bind("<ButtonPress-2>", self._press, add="+")
        w.bind("<B2-Motion>", self._motion, add="+")
        w.bind("<ButtonRelease-2>", self._release, add="+")

    def _pan(self, dx, dy):
        pan_px(self.w, 0, dy)              # 中键只管上下

    def _press(self, e):
        self._start(e)
        return "break"

    def _release(self, e):
        self._end()
        return "break"


def bind_wheel(w, middle=True, px=WHEEL_PX):
    """给一个可拖控件补上「滚轮」和「中键上下拖」（2026-11-04 用户定的通用规矩）。

    · 滚轮 / 中键滚 —— 上下走一格 `WHEEL_PX` 像素；
    · Shift + 滚轮 —— 横着走（宽表看右边几列）；
    · 中键按住上下拖 —— `middle=True` 时挂上（放大区不挂，那里中键是缩放）。

    滚轮的手感要跟左键拖一致，所以也走 `pan_px` 那套分数换算，不调
    `yview_scroll`：后者在 Listbox 上一格是「一项」、在 Canvas 上是一像素。
    """
    def _notches(e):
        d = getattr(e, "delta", 0) or 0
        if abs(d) >= 60:                   # Windows：一格是 120 的倍数
            d = d / 120.0
        return int(-round(d))              # 往上滚（delta 正）= 看前面 = dy 正

    def _wheel(e):
        n = _notches(e)
        if not n:
            return "break"
        try:
            shift = bool(int(e.state) & 0x0001)
        except Exception:                  # noqa: BLE001
            shift = False
        if shift:
            pan_px(w, n * px, 0)
        else:
            pan_px(w, 0, n * px)
        return "break"

    def _updown(n):
        """X11 的滚轮（Button-4/5）：按住 Shift 就横着走。

        X11 没有单独的 `<Shift-MouseWheel>` 事件序列 —— Shift 只体现在
        `state` 的 bit0 上，所以这里要自己看。
        """
        def _h(e):
            try:
                shift = bool(int(e.state) & 0x0001)
            except Exception:              # noqa: BLE001
                shift = False
            if shift:
                pan_px(w, n * px, 0)
            else:
                pan_px(w, 0, n * px)
            return "break"
        return _h

    w.bind("<MouseWheel>", _wheel, add="+")     # macOS / Windows
    w.bind("<Shift-MouseWheel>", _wheel, add="+")
    w.bind("<Button-4>", _updown(+1), add="+")  # X11 上滚
    w.bind("<Button-5>", _updown(-1), add="+")  # X11 下滚
    if middle:
        _MiddlePan(w)


def tv_click(w):
    """Treeview 的「点一下选中」—— 默认行为被 `_PanGrab` 挡掉后由这里补回来。"""
    def _do(e):
        try:
            iid = w.identify_row(e.y)
        except Exception:                  # noqa: BLE001
            return
        if not iid:
            return
        try:
            w.focus(iid)
            if iid not in w.selection():
                w.selection_set(iid)       # 触发 <<TreeviewSelect>>，与单击同一条路
        except Exception:                  # noqa: BLE001
            pass
    return _do
