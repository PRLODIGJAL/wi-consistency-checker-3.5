"""骨架与分栏：整窗搭建、左右分栏拖拽、两页切换、空态。

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

import tkinter as tk
from tkinter import ttk

import ui_theme as theme
from ui_const import APP_VERSION, AUTO_SHEET
from ui_pan import _PanDrag

# 两栏的宽度权重（2026-11-03 用户定：左 DW 宽、右 WI 窄 = 8 : 2）。
# 只影响「两栏各占多少」，跟里面的内容无关；要换比例改这两个数就行。
POOL_W_L, POOL_W_R = 8, 2


class _SelRing:
    """明细表选中行的灰框（2026-11-04 第六轮）。

    用户原话：「我在选中某一行的时候，怎么点是蓝色的，表头也是蓝色，这样子两个
    色彩冲突了……要不我选中的时候就给它不标颜色，就只标一个框最好。而且这个框
    用灰色或者什么颜色代替。」

    ttk 的 Treeview 只会给选中行**铺底色**，画不了「只描边不铺底」的行框；
    而铺任何底色都会和分组条的蓝色带抢视线。所以：
      · 底色：选中不铺色（Wide.Treeview 的 selected background = 卡片白）；
      · 框：  用 4 条 2px 的灰细线绕着行四周摆一圈 —— Tk 里唯一能「框住一行
              又不挡住内容」的画法（tk.Frame 有实心底，会把行里的字盖住）。

    什么时候要 refresh（漏一处，框就停在旧位置）：
      · 选中变了   → <<TreeviewSelect>>
      · 表滚动了   → x/yscrollcommand 回调（任何方式的滚动都走这两条）
      · 行被藏/挪  → _apply_group / _render 收尾
      · 窗口变了   → <Configure>
    """

    THICK = 2               # 框线粗细（px）

    def __init__(self, tree, parent):
        self.tree = tree
        self.parent = parent
        self.lines = [tk.Frame(parent, bg=theme.SEL_RING) for _ in range(4)]

    def _hide(self):
        for ln in self.lines:
            ln.place_forget()

    def refresh(self):
        """跟着当前选中行走；没有选中/选中行不在视野里就收起来。"""
        try:
            sel = self.tree.selection()
            bb = self.tree.bbox(sel[0]) if sel else ""
        except Exception:      # noqa: BLE001 —— 控件正在销毁
            self._hide()
            return
        if not bb:
            self._hide()
            return
        x, y, w, h = bb
        ox, oy = self.tree.winfo_x(), self.tree.winfo_y()
        t = self.THICK
        # 上/下两条横线压着行的上下边（-1 让框骑在边线上，不挤占行高）；
        # 左/右两条竖线把整行箍住。lift 保证压在表格内容之上。
        geo = ((x + ox - 1, y + oy - 1, w + 2, t),
               (x + ox - 1, y + oy + h - 1, w + 2, t),
               (x + ox - 1, y + oy - 1, t, h + 2),
               (x + ox + w - 1, y + oy - 1, t, h + 2))
        for ln, (px, py, pw, ph) in zip(self.lines, geo):
            ln.place(x=px, y=py, width=pw, height=ph)
            ln.lift()


class LayoutMixin:
    """骨架与分栏：整窗搭建、左右分栏拖拽、两页切换、空态。"""


    def _build_ui(self):
        # 铺满窗口的主容器 —— 拖放能力注册在它上面（tk.Tk 本身没有
        #    drop_target_register，只有 BaseWidget 才有），落在任意子控件上的
        #    拖放都会冒泡到这里（见 _setup_dnd）。
        self.host = ttk.Frame(self.root)
        self.host.pack(fill="both", expand=True)
        # 2026-11-04 第十三轮改：PDF 面板占**整个软件窗口**的右侧全高
        # （不是只在明细表右边）。用户原话：「我所谓的右侧应该是在整个软件的
        #   窗口的右侧，直接把这个软件给拉宽变宽那样子的效果……不管是上面还是
        #   下面都要被硬拓宽出来一部分，用来放这个 PDF。」
        # 所以 host 里放一个横向 PanedWindow：左边 = 全部现有内容，右边 = PDF。
        self.main_split = ttk.PanedWindow(self.host, orient="horizontal")
        self.main_split.pack(fill="both", expand=True)
        self.left_col = ttk.Frame(self.main_split)
        self.main_split.add(self.left_col, weight=5)
        # 右侧 PDF 面板的容器（面板本体由 ui_dialog._pin_dialog 搭）。
        # 没双击过任何一行时它不在 paned 里，界面上就完全看不到。
        self.pdf_dock = ttk.Frame(self.main_split)
        # 上部：左右两栏的「文件池」（2026-10-29 用户要求从左到右并排）——
        #   左 = 工程参数表 Excel，右 = WI 作业指导书 PDF。
        #   一栏里配「＋添加文件… / 移除选中 / 清空」三个按钮，也支持把文件从
        #   资源管理器直接拖进来（见 _setup_dnd）。清单只动内存，绝不删磁盘文件。
        #
        #   用词的变化要记一笔：这里是「池」，不是「要比的」。用户的原话是
        #   「我放进来不代表我要对比」—— 先把一大批都放进来，再在清单里点选
        #   **这次**要比哪几份（蓝色高亮即选中）。没选中的默认**不参加**比对。
        pools = tk.Frame(self.left_col, bg=theme.CARD,
                         highlightthickness=1, highlightbackground=theme.BORDER)
        # ⚠️ 这里**不能** expand —— 池子一拉伸就会去抢摘要区和明细表的高度，
        #    结果底下那张明细表被压成一条表头。池子的高度交给清单自己的
        #    行数（height=5）决定，多余的空间全留给下面。
        pools.pack(fill="x", padx=10, pady=(4, 2))
        # 2026-10-29 第三次改（用户原话）：「整个大方框……那个蓝色的框线应该是
        # 从那个白色的区域开始可以移动，而不是从那个灰色的区域就开始。然后，
        # 那个 WI 作业指导书 PDF 也是可以放在最右侧。然后，那个零分放在那个
        # 蓝色框线的右边一点，也就是让它形成一个镜面对对称的效果。
        # 我觉得把这个三个按钮在 PDF 那边给它固定死在最右侧就行了。」
        #
        # 所以这个大框改成一个 5 列 × 4 行的网格，每一行各管一件事：
        #   col: 0=左竖条 1=左清单 2=细线 3=右清单 4=右竖条
        #   row0 表头（左：标题靠左、份数贴线 / 右：份数贴线、标题靠右）
        #   row1 清单 + 细线 ← 细线**只跨这一行**，也就是只在白色区域里
        #   row2 两条横条（各在自己清单底下）
        #   row3 按钮行（左组钉死最左、右组钉死最右，跟细线毫无关系）
        # ⚠️ col1 必须 weight=0 + minsize，col3 才是 weight=1：
        #    两个都给 weight 的话拖动一边、另一边会跟着均分，线就拖不动了。
        self.pool_box = pools
        # ── 2026-11-03 改：分栏从「可拖细线」改成**固定 8 : 2**（用户定的）──
        # 用户原话：「左边是 DW 右边是 WI，就保持一个大概的这个位置比例就好了。
        #   8:2，就是让底图背景那边多一点，因为反正底下有滑条，所以中间那个
        #   左右拖动的线倒是没有必要存在了。」
        # 所以这里不再有 split_w 这个状态：宽度全交给 grid 的 weight 分配，
        # 两栏都带横向滚动条，窄了就往右拖，不用先拉线再拖条。
        # 想换成别的比例只改下面两个常量，别在别处写魔数。
        pools.rowconfigure(1, weight=1)                        # 清单那行吃高度
        pools.columnconfigure(self._COL_VBAR_L, weight=0)      # 左竖条
        pools.columnconfigure(self._COL_LIST_L, weight=POOL_W_L,
                              minsize=220)                     # 左清单（DW，宽）
        pools.columnconfigure(self._COL_LINE, weight=0)        # 分隔
        pools.columnconfigure(self._COL_LIST_R, weight=POOL_W_R,
                              minsize=150)                     # 右清单（WI，窄）
        pools.columnconfigure(self._COL_VBAR_R, weight=0)      # 右竖条
        # 表头：左右互为镜像 —— 份数都贴着中间那根线，标题都贴着大框外沿
        self._build_pick_head(pools, "xlsx", "工程参数表 Excel", "left"
                              ).grid(row=0, column=self._COL_VBAR_L,
                                     columnspan=2, sticky="ew")
        self._build_pick_head(pools, "pdf", "WI 作业指导书 PDF", "right"
                              ).grid(row=0, column=self._COL_LIST_R,
                                     columnspan=2, sticky="ew")
        self._build_pick_list(pools, "xlsx", "left")
        self._build_splitter(pools)
        self._build_pick_list(pools, "pdf", "right")
        # 横条原来钉在 row2（贴着两个清单的底边）；2026-11-04 第九轮跟着滑条一起
        # 撤掉了 —— 清单里按住左键横着拖就能看完整文件名。
        # 按钮两组装在同一个跨全宽的格子里：左组 side=left、右组 side=right，
        # 天然贴死大框的两个外沿，中间的细线怎么拖都碰不到它们。
        self._build_pick_btns(pools).grid(row=3, column=self._COL_VBAR_L,
                                          columnspan=5, sticky="ew", pady=(5, 0))

        # ── 按钮行（2026-10-29 用户要求：按钮往上提，底下多留地方给结果）──
        # 原来「使用工作表」单独占一整行，跟按钮各占一行；现在把它挪到按钮行
        # 右端，等于白赚一行高度。顺便把多余的解释性小字全删了 —— 界面不是
        # 说明书，用得着的信息（份数、已选）留在标题上就够了。
        #
        # 2026-09-28 用户拍板拿掉的四个入口，理由记在这里，别再有人加回来：
        #   ·「结构自检」「表格自检」两个按钮
        #   ·「显示全部列」「展开参数来源」两个勾选框
        # 底层方法 `_audit` / `_inspect` 与两个开关变量都保留（导出报告、
        # 回归脚本还要用），只是界面上不给入口。
        btnbar = ttk.Frame(self.left_col)
        btnbar.pack(fill="x", padx=10, pady=(2, 1))

        # ── 按钮行分三段（2026-10-30 用户定）：左＝动作，中＝工作表，右＝翻页 ──
        # 用户原话：「把那个开始对比、导出报告放在最左侧那三个按钮，然后把那个
        # 概览和明细放在最右侧，使用工作表选项放在中间」。原来「概览/明细」在最
        # 左、「三个动作按钮」和「使用工作表」挤在右侧，等于**先用右键点翻页、
        # 再用左键点动作**，两次都要横穿整个窗口。现在从右往左读是「看哪一页 →
        # 用什么表 → 干什么」，动作按钮落在拇指最容易到的最左边。
        #
        # 三段各用一个容器：逐个 pack(side="right") 是**后 pack 的贴最右**，
        # 顺序容易反，还会跟中间那段互相挤。容器各管一段，顺序写死。
        leftbox = ttk.Frame(btnbar)
        leftbox.pack(side="left")
        # 主按钮全屏只有一个，就是「开始比对」：实心主色 + 白字 + bold。
        # 其余一律白底细框（Small.TButton），靠位置和留白区分主次，
        # 不靠把每颗按钮都染成彩色（指南第 5.1 节）。
        self.btn_run = ttk.Button(leftbox, text="开始比对", style="Primary.TButton",
                                  command=self._run)
        self.btn_run.pack(side="left")
        # 2026-11-02 用户要求撤掉「导出报告」「导出 JSON」两个键：
        #   「帮我把那个导出报告还有导出 JSON 这两个功能键给删掉吧，没有用了」
        # 底层 `_export` / `_export_json` 两个方法**保留未删**（逻辑完好，只是没有
        # 界面入口了）—— 哪天要接回来，把它们挂到菜单上即可，不必重写。
        # 学习记录（2026-11-02 用户要求）：人工判定过的东西得有个常驻入口 ——
        # 原话「不然每次只有在对比的时候才弹出来一次」「我判定过的……是不是可以
        # 把那个记录给清空，我没判定过的，我是不是可以自己手动添加」。
        # 这个窗在**任何时候**都能打开，不必先跑一次比对。
        self.btn_learn = ttk.Button(leftbox, text="学习记录", style="Small.TButton",
                                    command=self._open_learn)
        self.btn_learn.pack(side="left", padx=(8, 0))

        # 右段：只看异常 / 全部展开 + 结论小字（2026-11-03 第四轮）
        #
        # 「概览 / 明细」两个分段键在这一轮被撤掉了。用户原话：
        #   「现在不是可以全部展开全部收起吗？我想要收起的时候把那些异常项就
        #     不要给它收起来……那这样子的话是不是就可以把那个概览的功能给平替掉，
        #     然后概览那个按钮就可以删掉，以及那一系列功能可以删掉。」
        # 也就是说：点一下「只看异常」，屏幕上只剩分组条 + 红色行，本身就是一张
        # 待改清单 —— 概览页那份长报告是「抬头看结论」，而这张清单直接把结论
        # 标在要改的那几行上，比读一页文字短。所以那一页连同它的按钮一起删。
        rightbox = ttk.Frame(btnbar)
        rightbox.pack(side="right", padx=(0, 2))   # 跟下面表格的右边沿对齐
        # 结论小字留在最右：概览页没了，这几个数字就是「这份表要不要管」的
        # 唯一一句话总结（原来是靠切进概览页才知道）。
        self.page_note = ttk.Label(rightbox, text="", style="Muted.TLabel")
        self.page_note.pack(side="right", padx=(14, 0))
        # 只看异常 / 全部展开（2026-11-03 分组上线时加，第四轮改语义，11-04 再收一道）：
        # 组是按参数卡分的，一份表上百张卡，一张张点开不现实；默认把「没有问题」
        # 的组收起来了，所以必须留一个一键的口子。有异常的组**不收，但也不整组
        # 摊开** —— 只留那几条待处理的行（用户原话：「有异常的那个表头的那一项，
        # 把那其中没有影响的也给收纳起来，只保留红色的露出来」）。这样按下这个键，
        # 屏幕上就真的是「一条一条要改的行」，没有一行是绿的，也就真正平替了概览。
        # 字随状态变，永远说「点下去会发生什么」。
        self.btn_fold = ttk.Button(rightbox, text="只看异常",
                                   style="Small.TButton", width=9,
                                   command=self._toggle_fold_all)
        self.btn_fold.pack(side="right")

        # 中段：使用工作表（跟着左栏选中的那份 Excel 走）
        # 外层 fill+expand 吃满左右两段之间的空档，内层 expand 在空档里居中。
        # 三步缺一不可：只写 expand 不写 fill，空档分给了外层但和外层同宽，
        # 下拉框会停在空档的最左边，看着像「贴着动作按钮」而不是在中间。
        midbox = ttk.Frame(btnbar)
        midbox.pack(side="left", expand=True, fill="x")
        midinner = ttk.Frame(midbox)
        midinner.pack(expand=True)
        ttk.Label(midinner, text="使用工作表").pack(side="left", padx=(0, 6))
        self.sheet_box = ttk.Combobox(midinner, textvariable=self.sheet_var,
                                      state="readonly", values=[AUTO_SHEET],
                                      width=22, font=(theme.FONT, theme.F_SMALL))
        self.sheet_box.pack(side="left")

        # ── 底部状态栏 ──
        # 从按钮行右侧挪到最底下：一是按钮行不被长路径顶走，二是条带占满宽度，
        # 完整路径能整条显示（原来得截到 96 个字符）。
        # 底色比窗口底再深一档（#eceef1），整块界面就有了「收口」；
        # 只有这一处报统计，顶部不再重复（指南第 7 节）。
        self.status = tk.StringVar(value="就绪（自动识别 WI 文档）")
        statusbar = ttk.Frame(self.left_col, style="Status.TFrame")
        statusbar.pack(side="bottom", fill="x")
        # ⚠️ 摆放顺序有讲究：先 pack 右侧那几个**固定宽度**的，最后再 pack
        #    左侧那个 expand=True 的。反过来的话，左侧长文本会先吃掉整条宽度，
        #    右边两个被压成几个像素（实测自检那行只剩 4px，字直接被裁没）。
        # 版本号不塞进窗口标题（标题栏一长就被系统截断），放这条最右边，
        # 9 号灰字，不抢正文的视线。
        ttk.Label(statusbar, text="v%s" % APP_VERSION,
                  style="StatusTiny.TLabel").pack(side="right", padx=(0, 10))
        # ── 结构自检（2026-11-02）──
        # 右边一行小字，比对完自动跑。为什么要有：11-02 一天里连着三个同类 bug
        # （2534 多段多套、2535 应为两个值、2535 表头误判），症状不同但机制相同 ——
        # 解析器少认一种写法就**静默变宽**，界面上一点错都不报，只能靠人肉发现。
        # 这几条不变量就是这个「不吭声」的兜底：跑完立刻查，可疑就亮出来。
        # 放状态栏而不是弹窗：绝大多数时候它是「通过」，不该打断人。
        self.audit_var = tk.StringVar(value="")
        self.audit_lbl = ttk.Label(statusbar, textvariable=self.audit_var,
                                   style="Status.TLabel", anchor="e", cursor="")
        self.audit_lbl.pack(side="right", padx=10, pady=3)
        self.audit_lbl.bind("<Button-1>", lambda _e: self._show_audit())
        ttk.Label(statusbar, textvariable=self.status, style="Status.TLabel",
                  anchor="w").pack(side="left", fill="x", expand=True,
                                   padx=10, pady=3)

        # 明细表独占整块高度（2026-11-03 第四轮：「概览」页删掉之后）
        #
        # 历史：2026-10-29 把「摘要 + 明细」拆成两页，为的是两块别互相挤；
        # 现在摘要那一页整个撤了（用户：「概览那个按钮就可以删掉，以及那一系列
        # 功能可以删掉」），结论由「只看异常」+ 分组条 + 右上角那行小字承担，
        # 所以这里只剩一页，直接吃满。
        # 2026-11-04 第十三轮：PDF 面板移到全窗口右侧（main_split），
        # 这里只剩明细表本身，不再包 PanedWindow。
        self.page_holder = ttk.Frame(self.left_col)
        self.page_holder.pack(fill="both", expand=True, padx=10, pady=(2, 8))
        self.page_table = self.page_holder
        self.page_holder.rowconfigure(0, weight=1)
        self.page_holder.columnconfigure(0, weight=1)

        # 明细表（外层 frame 承放两个滚动条）
        table = self.page_table
        # 行号放在**最左**：判定结果要回 Excel 里改，第一眼就得看到是哪一行。
        #
        # 一条数据占**上下两行**（2026-09-29 用户定的版式）：
        #   上行 = Excel 现在填的；下行 = 应为（行号/参数卡留空，「限值」列
        #   正下方放应为值，两个「是/否」列放**按类型规矩应该填的**）。
        #   上下两行一比，值填错、是/否填反、类型写错，三种错都能当场看出来。
        #
        # 列序**跟 Excel 表头一致**（2026-09-29 用户要求）：采集项 → 类型 →
        # 是否显示上下限 → 是否允许超上下限 → 才到限值。原来限值在类型左边，
        # 跟源表的摆放不一样，来回对照时要重新找位置。
        cols = ("row", "card", "item", "type", "show", "over", "value",
                "unit", "proc", "basis", "match", "pad")
        # 默认显示 9 列（2026-11-04 第六轮在 8 列基础上加了最右的 pad）：
        # 行号 → 采集项 → 类型 → 两个是/否 → 限值 → 单位 → 核对结果 → 规格·设备。
        # 参数卡为什么撤：它的名字是「一大串前缀 + 0001」，486 行里 481 行的前缀
        # 一模一样，按行重复印 152px 的列宽，等于用整列去重复同一句话；而人真正
        # 靠它分辨的只有尾巴那个编号。现在前缀进**分组条**（只印一次），
        # 编号写在分组条上，这一列就不必存在了。列仍然保留在 cols 里（导出、
        # 双击定位的路由都还按 12 个字段取），只是不进 displaycolumns。
        #
        # pad 是**唯一 stretch 的列**（2026-11-04 用户原话：「右边留的这么多空位
        # 就可以用上，可以把这个表头给延展到最右侧」）：它把窗口右边那截空档吃掉，
        # 分组条的色带跟着铺到最右；分组条的「规格 · 设备」也写在这一格里
        # （见 ui_table._grp_meta_text）。数据行这一格留空 —— 同一张卡的规格/设备
        # 在分组条上只印一次，再按行重复就是当年撤掉参数卡列的反面。
        self.all_cols = cols
        # 2026-11-04 第九轮收盘，`match`（核对结果）从第 8 列挪到第 3 列 —— 紧跟在
        # 采集项后面。用户原话：「核对结果和设备规格不一定要按照最上面的那个格式
        # 去放置啊，表头可以单独放一种格式，也就是说中间左边空了这么多位置，可以
        # 把它全都往左移一点，然后显示更多内容。」
        #   起因：分组条上那句「58 行 · 显示 2 条」写的是 match 这一列，而它前面
        #   隔着 类型/上下限/超上下限/限值/单位 五列 —— 这五列在分组条上永远是空的，
        #   于是表头中间白出 250px 一截。
        #   挪过来之后：采集项 → 核对结果 紧挨着，分组条那一行也就连起来了。
        #   Treeview 的单元格不能跨列，所以「表头单独排版」的落地方式就是**换列序**。
        # 2026-11-04 第十轮，用户原话：「没有必要专门搞一列设备规格……你就把设备
        # 规格给列在那个核对结果的右侧……没有必要展示设备规格这个表头这个名字」。
        # 于是 pad 整列撤出视图（数据仍留在 cols 里给导出用），规格和设备改写在
        # **分组条那句话的末尾**（见 ui_table._grp_line）—— 一张卡只印一次，
        # 不占一整列的宽度，也不用再给它一个表头。
        #
        # 表头也按用户的话减了两个：「采集项」「核对结果」这两种东西「是人类一眼
        # 就能看出来的」，就不印名字了（留空，列还在）。
        #
        # 2026-11-11：`pad` 重新回到视图里，但**只当右缘的空白占位**，表头留空、
        # 数据行留空。为什么要它：窗口宽 1020，八列按内容量下来只占 830 ——
        # 剩下那截如果没列去吃，分组条那条蓝带就铺不到最右边（第十轮的老问题）；
        # 如果让「单位」去吃，就会看见一个 200px 宽、中间只有一个「—」的列，
        # 看着像表格排错了（本轮截图实测）。空白列既不冒充内容，也正好是用户
        # 说过两次的「右边空出来给以后加新东西的位置」。
        self.basic_cols = ("row", "item", "match", "type", "show", "over",
                           "value", "unit", "pad")
        self.tree = ttk.Treeview(table, columns=cols, show="headings")
        # 「行号」这个列名删掉了（2026-11-11 用户原话：「那个不应该叫行号……在那个
        # 表图那里它对应的那个不叫行号，那个叫序列号。你把行号的那两个字给删掉」）。
        # 为什么删得掉：这一列里放的是两种东西 —— 数据行是 `表4!226`（Excel 的行），
        # 分组条上是 `0001`（参数卡的序号）。两个都不是「行号」，硬起一个名字必然
        # 骗一半人；而数字本身一眼就知道是「去源表里第几行/第几号卡」，不必标。
        headings = {"row": "", "card": "参数卡", "item": "",
                    "value": "限值 / 应为", "type": "类型",
                    "show": "上下限", "over": "超上下限",
                    "unit": "单位", "proc": "工艺",
                    "basis": "依据（PDF 里的出处）", "match": "",
                    # pad 是右缘那块留白，不印名字 —— 它没有任何内容，
                    # 印个标题反而会让人去那一列里找东西。
                    "pad": ""}
        # 列宽（2026-11-03 第四轮，用户原话：「你都给它固定死了，包括采集项这边
        # 都没占满，你是不是应该把这个间距给它调整一下呢？把它往左移一点，使右边
        # 空出来更多位置，让右边到时候以后或许要加一些新的东西」）：
        #   · **所有列都 stretch=False**（上一轮只有采集项会拉伸）：拉伸会让
        #     「采集项」在宽屏上吃掉几百像素的空白，看着像这一列没填满；而且
        #     窗口一变宽，整张表的列位就全漂了，复核时眼睛要找位置。
        #   · 采集项宽度改为**按内容算**（见 ui_table._fit_item_col）：这一列
        #     多长取决于数据里最长的那个采集项名，不是取决于窗口。
        #   · 这样一来右边必然留一条空白 —— 那正是用户要的「以后加新东西的
        #     位置」。真要看长内容，横向条还在。
        #   · 限值 232→118 且右对齐：一格数就 3~5 个字符，原来那 232px 有 3/4 是空的。
        widths = {"row": 72, "card": 152, "item": 240, "value": 74, "type": 46,
                  "show": 46, "over": 62, "unit": 40, "proc": 76, "basis": 300,
                  "match": 282, "pad": 240}
        # 第九轮收盘这几列的宽度是按**表头字宽实测**定的（TkHeadingFont 量出来：
        # 超上下限 44、单位 22、限值 / 应为 54，再加行内左右各 9px 呼吸）——
        # 上一版凭感觉给 56/30，结果表头被切成「超上下」「单」，正是用户说的
        # 「框一缩小，字全被挡住」。
        # 各列 2026-11-04 第六轮按「内容 + 8px 呼吸」重新量到骨（一个字没截，
        # 量的是 PingFang 实测像素）：行号 96→72、类型 52→49、上下限 58→49、
        # 超上下限 74→60、限值 118→80、单位 44→34、核对结果 196→178 ——
        # 挤出来的全给了 pad，分组条上的「规格 · 设备」就是靠它们才塞得下。
        # 第九轮收盘再收一遍（用户原话：「中间左边空了这么多位置，可以把它全都
        # 往左移一点，然后显示更多内容」）：核对结果 178→300 并改成**左对齐**，
        # 前面那五列各让几像素（类型 49→46、上下限 49→46、超上下限 60→56、
        # 限值 80→72、单位 34→30）补给给它。分区条上的「58 行 · 显示 2 条」写的
        # 正是这一列，列一起左移，分组条中间那截空当也跟着收掉。
        # 下限与宽度取同一个值：拖不动（见 _on_tree_click / _on_tree_motion），
        # 万一被程序改小也不会低到只剩「…」。pad 是唯一的例外（stretch 补满右缘）。
        minw = dict(widths)
        # 伸展列从 pad 换成 match（2026-11-04 第十轮）：pad 撤出视图之后，铺到
        # 窗口右缘、把分组条那条色带拉满的责任交给「核对结果」这列 —— 它现在
        # 既是判定词的老家，也是分组条那句话（行数 + 规格 + 设备）的老家。
        # 下限收到 240：窄窗口下 ttk 会拿 minwidth 把 width 夹回去，写 282 就
        # 夹不回来了（第九轮在 pad 上踩过一次，见下面的注释）。
        # 2026-11-11 再改回来：**match 不再伸张**，宽度由内容算（ui_table.
        # _fit_match_col / _match_need_px），不跟窗口走 —— 用户点名的「判定结果
        # 和类型之间隔了很大的位置」就是因为这列按窗口比例长胖。铺满右缘的
        # 责任还给 pad（空白占位列）。
        minw["match"] = 180
        minw["unit"] = 30
        minw["pad"] = 40
        _center = ("row", "type", "show", "over", "value")
        for c in cols:
            self.tree.heading(c, text=headings[c])
            self.tree.column(c, width=widths[c], minwidth=minw[c],
                             stretch=(c == "pad"),
                             # pad 里只有分组条右端那句「N 行 · M 条待处理」，
                             # 右对齐才像一条「数量」，跟左边的设备号一头一尾。
                             anchor="e" if c == "pad"
                             else ("center" if c in _center else "w"))
        self.tree["displaycolumns"] = self.basic_cols
        self.tree.tag_configure("ok", foreground=theme.OK_FG)
        self.tree.tag_configure("bad", foreground=theme.BAD_FG)
        self.tree.tag_configure("warn", foreground=theme.WARN_FG)
        self.tree.tag_configure("na", foreground=theme.NA_FG)
        # 分组条：加粗 + 主色 + 字号放大一档 + 一条底色带（2026-11-03 第四轮）。
        # 用户原话：「关于这个表头，就是这个蓝色的字体给它变大一点，让它与底下的
        # 那个展开之后的正文部分形成一些明显的差距」。
        # 正文是 TkDefaultFont(10)，分组条用 F_HEAD(12) 加粗；
        # 行高跟着给到 24（Aqua 默认 18 会把 12 号加粗的上下各切掉一点 —— 实测）。
        #
        # ⚠️ 顺带纠正一条旧注释：这里原来写着「macOS 的 Aqua 主题会忽略 tag 的
        # background」。2026-11-03 实测**不成立** —— tag 的底色是会画的
        # （单独建一个 Treeview 设 20 号字 + 橙底，橙底正常显示，只是字被行高裁掉）。
        # 所以除了字号，再给分组条一条浅色横带：一眼就把「标题」和「数据」分开了。
        try:
            _st = ttk.Style()
            _st.configure("Wide.Treeview", rowheight=theme.ROW_H_WIDE)
            # 选中不铺色（第六轮）：主明细表的选中底 = 卡片白，和没选一样，
            # 「你在这一行」全靠 _SelRing 的灰框说。弹窗里的小表不走这个样式，
            # 选中照旧是浅蓝底 —— 它们没有分组条，不存在撞色问题。
            _st.map("Wide.Treeview", background=[("selected", theme.CARD)])
            self.tree.configure(style="Wide.Treeview")
        except Exception:      # noqa: BLE001 —— 主题不支持就退回默认行高
            pass
        # grp = 这一组有待处理的行（蓝字 + 蓝带，要抢眼）；
        # grp2 = 全部一致（灰字 + 浅灰带，130 条全亮的话，那几条真要管的会淹没）。
        # 2026-11-04 第九轮：蓝带整体配浅一档（GRP_BG / GRP_FG）—— 原来用
        # TINT_HEAD + PRIMARY，130 条铺下来饱和蓝的面积把「主色 ≤5%」破了。
        self.tree.tag_configure("grp", foreground=theme.GRP_FG,
                                background=theme.GRP_BG,
                                font=(theme.FONT, theme.F_HEAD, "bold"))
        self.tree.tag_configure("grp2", foreground=theme.TEXT2,
                                background=theme.FIELD,
                                font=(theme.FONT, theme.F_HEAD, "bold"))
        # 下行（应为行）用浅一档的同色：既看得出属于上面那条，又不会和现值行混淆
        self.tree.tag_configure("ok2", foreground=theme.OK_FG2)
        self.tree.tag_configure("bad2", foreground=theme.BAD_FG2)
        self.tree.tag_configure("warn2", foreground=theme.WARN_FG2)
        self.tree.tag_configure("na2", foreground=theme.NA_FG2)
        # 分隔行：浅灰，看着就是一条线而不是一条数据
        self.tree.tag_configure("sep", foreground=theme.SEP)

        # 两条滑条 2026-11-04 第九轮撤掉（用户：「对于所有的有滑条的地方，都可以
        # 把那个滑条进行删除」）—— 按住左键拖就是平移，上下左右都行；Shift+滚轮
        # 照旧横向移。撤掉之后 x / y 的 scrollcommand 仍然要接：它是**所有滚动
        # 方式**的共同出口（左键拖、滚轮、键盘翻页都从这儿过），灰框靠它跟着走。
        self._sel_ring = _SelRing(self.tree, table)
        self.tree.configure(yscrollcommand=lambda *a: self._sel_ring.refresh(),
                            xscrollcommand=lambda *a: self._sel_ring.refresh())
        # 「拖」用 _PanDrag：按下不拦（点行选中、点分组条折叠都照旧），只把
        # <B1-Motion> 那一段接过来平移。挂在这里（早于下面 _on_tree_click 的
        # 绑定）是为了让「记录起点」先发生 —— 后者对列边界/分组条会 return
        # "break"，一旦先跑，拖动就起不来了。
        _PanDrag(self.tree, on_click=self._on_tree_click_done)
        self.tree.bind("<<TreeviewSelect>>",
                       lambda _e: self._sel_ring.refresh(), add="+")
        # 尺寸一变：灰框跟着走，分组条的「规格 · 设备」按新宽度重截
        # （pad 是伸展列，画多宽取决于窗口，见 ui_table._on_tree_resize）
        self.tree.bind("<Configure>",
                       lambda _e: self._on_tree_resize(), add="+")
        self.tree.grid(row=0, column=0, sticky="nsew", columnspan=2)
        # Excel 的手感：Shift + 滚轮左右移。很多人不会去拖滚动条，但都会滚滚轮，
        # 而 Treeview 默认滚轮只走纵向 —— 不给这条绑定，横向条就等于只能手拖。
        self.tree.bind("<Shift-MouseWheel>", self._wheel_x)
        # 左键：拦列宽拖动 + 点分组条折叠（见 ui_table._on_tree_click）。
        # ⚠️ 这条必须绑在 Button-1 上，靠 return "break" 抢在 Treeview 的类绑定
        #    之前把 separator 的拖动吃掉 —— 换个时机（比如 ButtonPress 之后再拦）
        #    就拦不住了，列宽照样能被拖。
        # ⚠️ 必须 add="+"：这一行原来不带它，一次 bind 就把上面 _PanDrag 挂在
        #    <ButtonPress-1> 上的「记起点」冲掉了 —— 于是拖动永远起不来（第九轮
        #    实测：按下去有事件，就是不动）。
        self.tree.bind("<Button-1>", self._on_tree_click, add="+")
        # 鼠标划过列边界：连「左右双箭头」光标都不该出现（2026-11-03 第四轮）。
        # 用户原话：「既然不能调整了，那这个应该把那些代码给删掉，我现在鼠标放到
        # 这个地方，它居然中间还有左右移动键，虽然不能调整。」
        # 那个光标是 Tk 自己的类绑定（ttk::treeview::Motion）设的，只能在这里
        # 抢在类绑定之前拦掉 —— widget 级绑定先于类绑定执行，return "break" 即止。
        self.tree.bind("<Motion>", self._on_tree_motion)
        # 双击行 = 「这一行在 PDF 里到底是哪一格，我自己指给你看」。
        # 见 _pin_dialog 的长注释（用户 2026-11-02 定的交互：先选表、再选值）。
        self.tree.bind("<Double-1>", self._on_row_dbl)
        self._row2chk = {}
        # 按参数卡分组的状态（2026-11-03 第三轮加，11-04 补上「组内再筛一道」）：
        #   _collapsed     整组收起来的卡（一行都不显示）
        #   _trimmed       只显示「待处理」那几行的卡（异常组默认就这个状态）
        #   _grp_rows      分组条 iid → 参数卡名，点它就是要折叠/展开
        #   _grp_children  参数卡名 → 这一组所有子行 iid，按插入顺序
        #   _grp_keep      参数卡名 → {子行 iid: 只看异常时留不留}
        #   _grp_n         参数卡名 → 这张卡有几条数据行（分组条上那句「N 行」）
        self._collapsed = set()
        self._trimmed = set()
        self._grp_rows = {}
        self._grp_children = {}
        self._grp_keep = {}
        self._grp_n = {}
        self._cur_grp = None
        # 每次 _render 置 True → 按「这一组有没有待处理行」重算默认收起状态
        self._auto_fold = True
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)

        # ── 空态引导（2026-11-03，2.1 改版）──
        # 没跑过比对时，这张表是**一整片白**，看着像「坏了」。
        # 指南第 5.3/8 节：空态要写「该做什么」，不能留一片白，也不写
        # 「暂无数据」这种死话。所以盖一行灰字在表格正中，有数据就自己消失。
        # 用 place 叠在 Treeview 同一个格子上：不占布局、不影响列宽、
        # 也不参与任何点击（takefocus=0 + 关掉事件）。
        self.empty_hint = tk.Label(
            table,
            text="还没有比对结果\n\n把文件拖进上面的两个框，左右各点中一份，再点「开始比对」",
            bg=theme.CARD, fg=theme.TEXT2, font=(theme.FONT, theme.F_BODY),
            justify="center")
        self.empty_hint.place(relx=0.5, rely=0.42, anchor="center")
        try:
            self.empty_hint.configure(takefocus=0)
        except tk.TclError:
            pass

        # 只剩一页了，不用再「把哪页抬上来」；空态引导照旧按有没有行来判定。
        self._sync_empty_hint()



    def _sync_empty_hint(self):
        """有数据就把空态那行字藏掉，没数据就摆回来。

        判据只有一个：明细表里有没有行 —— 不去猜「是不是刚启动」。
        这样任何一条会清空/填充表格的路径（比对、批量、扫描件中止、关闭）
        都自动对，不需要在每个调用点各记一次。
        """
        hint = getattr(self, "empty_hint", None)
        if hint is None:
            return
        try:
            has_rows = bool(self.tree.get_children())
        except Exception:      # noqa: BLE001 —— 控件还没建好
            return
        if has_rows:
            hint.place_forget()
        else:
            hint.place(relx=0.5, rely=0.42, anchor="center")



    # -------------------------------------------------- 两栏之间的分隔
    def _build_splitter(self, parent):
        """两栏之间那条固定的分隔线（2026-11-03 起**不可拖**）。

        历史：2026-10-29 这里是一根可拖的细线，用户当时的诉求是「左栏能看到
        Excel 全名、右栏只露 WI 号」。现在两栏宽度定死 8 : 2（见 `POOL_W_*`），
        而且两栏底下各自都有横向滚动条 —— 要看全名就往右拖条，不必先调比例。
        用户原话：「反正底下有滑条，所以中间那个左右拖动的线倒是没有必要存在了。」

        所以这里只剩「一条看得见的分隔」：7px 空隙 + 正中 1px 线，没有光标、
        没有事件绑定。保留 `split_bar` 这个名字，是因为别处（截图/回归）还在
        引用它，改个名只会多出无谓的改动。
        """
        self.split_bar = tk.Frame(parent, bg=theme.CARD, width=7)
        # row=1 就是清单那一行 —— 分隔线**只跨白色区域**，不伸到上面的表头、
        # 也不伸到下面的横条和按钮（2026-10-29 用户原话：「那个蓝色的框线应该
        # 是从那个白色的区域开始可以移动，而不是从那个灰色的区域就开始」）。
        self.split_bar.grid(row=1, column=self._COL_LINE, sticky="ns")
        self.split_line = tk.Frame(self.split_bar, bg=theme.BORDER, width=1)
        self.split_line.place(relx=0.5, rely=0, anchor="n", relheight=1.0)
        return self.split_bar



    # ------------------------------------------------------------ 显示开关
    def _visible_cols(self):
        """明细表当前显示的列。

        2026-10-29 起默认就是 basic_cols（9 列，已含 unit）；show_all_cols 恒
        False、界面上也没有勾选框了 —— 这里留分支只为兼容旧记忆文件。
        """
        return self.all_cols if self.show_all_cols.get() else self.basic_cols



    def _toggle_cols(self):
        self.tree["displaycolumns"] = self._visible_cols()



    def _wheel_x(self, e):
        """Shift + 滚轮 → 明细表左右移（Excel 的手感）。"""
        step = -1 if getattr(e, "delta", 0) > 0 else 1
        self.tree.xview_scroll(step * 3, "units")
        return "break"          # 吞掉事件，免得同时往上滚一行
