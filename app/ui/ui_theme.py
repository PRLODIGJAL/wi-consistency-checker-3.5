# -*- coding: utf-8 -*-
"""ui/theme.py —— 视觉规格：色板、字号阶梯、ttk 样式。

依据：《项目i5统一UI美学/浅色系UI风格指南.md》v1.0
基准作品：PDF 差异对比工具 6.1（它的 `ui_theme.py` + `_build_ui()` 前 40 行
就是这份指南的落地版，本文件按同一套值写，WI 只多了「双平台字体」和
「判定语义色」两项）。

一句话：**浅底、细线、大面积留白，饱和色只用在主按钮上。**

三条底层逻辑（改任何颜色之前先读这三条）：
  ① 三级底色造层次，不用阴影、不用渐变：
     窗口 #f5f6f8 → 卡片 #ffffff → 块中块 #f7f9fc，色差只有 2~3%。
  ② 分界靠 1px #d9dee5 细线，不靠粗边框；需要强调就把线换成主色。
  ③ 需要「有颜色但不抢眼」时用同色相的极淡 tint（12%~18%），不铺纯色。

⚠️ 这里**只放值和样式**，不放任何业务行为。想知道某个颜色用在哪，
   全文搜常量名即可 —— 界面代码里不许再出现裸的十六进制色值。
"""

import sys

import tkinter as tk
from tkinter import ttk

APP_TITLE = "WI 一致性检查工具"

# 跨平台中文字体：Windows 用微软雅黑 UI，Mac 用苹方（Tk 找不到时自动回退）。
# 指南第 4 节：字号阶梯两端一样，只换字族。
FONT = "Microsoft YaHei UI" if sys.platform.startswith("win") else "PingFang SC"

# ── 1. 基础层（三级底色 + 两个功能底）──────────────────────────────────────
BG = "#f5f6f8"          # 窗口底 / 工具栏 / 弹窗底
CARD = "#ffffff"        # 卡片、清单、表格、输入框
DZ_BG = "#f7f9fc"       # 块中块（拖放区 / 输入区常态）
DZ_BG_ON = "#eaf1ff"    # 拖入悬停 / 选中态浅蓝
CANVAS = "#eef0f3"      # 预览区、图片容器（比窗口底稍深，把内容衬出来）
STATUS = "#eceef1"      # 底部状态栏（略深于窗口底，形成收口）

# ── 2. 线条与文字 ────────────────────────────────────────────────────────
BORDER = "#d9dee5"      # 全局唯一描边色，1px
TEXT = "#2a2f36"        # 正文（不用纯黑）
TEXT2 = "#7a828c"       # 次要文字 / 标签 / 占位
MUTED = "#9aa1a9"       # 图例、脚注
DISABLED = "#c3c8ce"    # 按钮禁用态

# ── 3. 强调与语义 ────────────────────────────────────────────────────────
PRIMARY = "#2f6fed"     # 主按钮底、图标、链接；**面积 ≤ 5%**
PRIMARY_D = "#2257c4"   # 主按钮 hover / pressed、深蓝文字
WARN = "#b26a00"        # 「这不是错误，是先提醒你」类标题
DANGER = "#c0392b"      # 止损红：只放在字上，不铺底

# ── 4. 大面积 tint（表格 / 区块带 / 表头）─────────────────────────────────
# ⚠️ 铁律：tint 底上不能用主色当文字（对比不够），文字必须降一档。
TINT_BAND = "#d9e7fc"       # 区块带底 / 清单选中底
TINT_BAND_FG = "#1e4fbe"    # 压在 TINT_BAND 上的字
TINT_HEAD = "#eaf1ff"       # 表头行底 / 悬停
TINT_HEAD_FG = "#2257c4"    # 压在 TINT_HEAD 上的字
FIELD = "#f1f5fa"           # 字段标签列 / 行号列
SEP = "#e3e7ec"             # 表格分隔行、摘要里的细横线
# 分组条（明细表里那张卡的「大表头」，2026-11-04 第九轮收盘又降了一档）：
# 用户原话「折叠之后的那个表头，那个蓝色也是过于深了，我觉得也要给它降低一下深度」。
# 原来底是 TINT_HEAD、字是 PRIMARY —— 130 条带子铺下来，饱和蓝的面积远超「≤5%」
# 那条铁律。现在底再淡一点、字换成去饱和的钢蓝：还认得出「这一组要管」，
# 但不再和主按钮抢眼睛。
GRP_BG = "#eef4fe"          # 分组条底（比 TINT_HEAD 再淡半步）
GRP_FG = "#3f5f9e"          # 分组条字（钢蓝：蓝，但不饱和）
# 选中行的框（2026-11-04 第六轮）：明细表的选中**不铺底色**（铺了就和分组条的
# 蓝色带撞车，用户原话「两个色彩冲突了，显得很乱」），改用 2px 灰线框住那一行。
# 用中性灰而不是主色：框只负责「你在这一行」，不该再抢一层语义色。
SEL_RING = "#9aa1a9"

# ── 5. 判定语义色（内容层：绿=一致 / 红=不一致 / 琥珀=待确认）──────────
# 主行用深一档，下行（「应为」行）用浅一档 —— 既看得出属于上一条，
# 又不会和现值行混淆。这组是 WI 特有的，6.1 没有对应物。
OK_FG = "#1b6e45"
OK_FG2 = "#4e9464"
BAD_FG = DANGER
BAD_FG2 = "#d9736a"
WARN_FG = WARN
WARN_FG2 = "#cfa05a"
NA_FG = MUTED
NA_FG2 = "#b9bfc7"

# ── 6. 字号阶梯（同一界面 ≤4 档）──────────────────────────────────────────
F_TITLE = 13    # 弹窗标题
F_APP = 12      # 应用标题
F_HEAD = 13     # 表格分组条（正文 10 号，这里 13 号加粗才看得出是标题）
F_BODY = 11     # 默认正文 / 按钮
F_SMALL = 10    # 说明、清单文件名、工具条、状态栏
F_TINY = 9      # 图例、版本号

# 行高：Treeview 一行一个字，26 是「10 号字 + 上下各 8px」的实测值；
# 再矮文字会贴着行线，再高同屏行数就少了。
ROW_H = 26
# 明细表行高：要装得下分组条那档字号。macOS Aqua 默认 rowheight 是 18，
# 而 F_HEAD(13) 加粗的 linespace 实测 24 —— 装不进去就会被裁掉上下各半截
# （2026-11-03 用 tag 设 20 号字实测过：字是变大了，但上下直接被切平）。
# F_HEAD(13) 加粗的 linespace 是 24，给 27 就是「上下各留 1.5px」。
ROW_H_WIDE = 27
# 清单高度（行）：让出高度给下面的明细表，这是用户反复要求的
LIST_ROWS = 5


def install(root=None):
    """把整套 ttk 样式装上去。必须在建任何 ttk 控件**之前**调用。

    为什么第一步就 `theme_use("clam")`：macOS 默认的 aqua 主题**忽略**
    `background` / `bordercolor` 这些选项，按钮永远是系统灰，改不动；
    clam 是 Tk 自带的跨平台主题，颜色全听配置的，两端才能长得一样。
    clam 的按钮天然带 1~2px 抗锯齿小圆角，正好对上指南第 3 节
    「桌面端不要自己做圆角」。
    """
    st = ttk.Style(root)
    try:
        st.theme_use("clam")
    except tk.TclError:      # 极端情况下 clam 不在，退回默认主题也比崩了强
        pass

    st.configure(".", background=BG, foreground=TEXT, font=(FONT, F_BODY))
    st.configure("TFrame", background=BG)
    st.configure("Card.TFrame", background=CARD)
    st.configure("Status.TFrame", background=STATUS)

    st.configure("TLabel", background=BG, foreground=TEXT, font=(FONT, F_BODY))
    st.configure("Card.TLabel", background=CARD, foreground=TEXT,
                 font=(FONT, F_BODY))
    st.configure("Muted.TLabel", background=BG, foreground=TEXT2,
                 font=(FONT, F_SMALL))
    st.configure("CardMuted.TLabel", background=CARD, foreground=TEXT2,
                 font=(FONT, F_SMALL))
    st.configure("CardTitle.TLabel", background=CARD, foreground=TEXT,
                 font=(FONT, F_SMALL, "bold"))
    st.configure("Status.TLabel", background=STATUS, foreground=TEXT2,
                 font=(FONT, F_SMALL))
    st.configure("StatusTiny.TLabel", background=STATUS, foreground=MUTED,
                 font=(FONT, F_TINY))

    # ── 按钮三档（指南 5.1，只用这三档）─────────────────────────────────
    # 次按钮：白底 + 1px 描边 + 正文色。绝大多数按钮都是这一档。
    st.configure("TButton", font=(FONT, F_BODY), padding=(8, 4),
                 background=CARD, foreground=TEXT, bordercolor=BORDER,
                 lightcolor=CARD, darkcolor=CARD, focuscolor=CARD)
    st.map("TButton",
           background=[("active", "#f0f2f5"), ("pressed", "#e6e9ee"),
                       ("disabled", CARD)],
           foreground=[("disabled", DISABLED)])
    # 小一号的次按钮：清单那排「＋添加文件 / 移出 / 清空」用
    st.configure("Small.TButton", font=(FONT, F_SMALL), padding=(6, 2),
                 background=CARD, foreground=TEXT, bordercolor=BORDER,
                 lightcolor=CARD, darkcolor=CARD, focuscolor=CARD)
    st.map("Small.TButton",
           background=[("active", "#f0f2f5"), ("pressed", "#e6e9ee"),
                       ("disabled", CARD)],
           foreground=[("disabled", DISABLED)])
    # 主按钮：实心主色 + 白字 + bold。**一屏只能一个**（这里是「开始比对」）。
    st.configure("Primary.TButton", font=(FONT, F_BODY, "bold"), padding=(12, 4),
                 background=PRIMARY, foreground="#ffffff", bordercolor=PRIMARY,
                 lightcolor=PRIMARY, darkcolor=PRIMARY, focuscolor=PRIMARY)
    st.map("Primary.TButton",
           background=[("active", PRIMARY_D), ("pressed", PRIMARY_D),
                       ("disabled", "#a9c0f3")],
           foreground=[("disabled", "#ffffff")])
    # 止损按钮：白底 + 红字 + 细框（不铺红底，理由见指南 5.1 的注）
    st.configure("Stop.TButton", font=(FONT, F_SMALL), padding=(7, 3),
                 background=CARD, foreground=DANGER, bordercolor=BORDER,
                 lightcolor=CARD, darkcolor=CARD, focuscolor=CARD)
    st.map("Stop.TButton",
           background=[("active", "#fdf1ef"), ("pressed", "#f7e0dc"),
                       ("disabled", CARD)],
           foreground=[("disabled", DISABLED)])

    # ── 概览 / 明细这类分段切换（原来是 Radiobutton style="Toolbutton"）──
    # 选中态只换底色和字重，padding / 字号必须与未选中逐项相同，
    # 否则一切换整行会左右抖一下（6.1 踩过的坑）。
    st.configure("Toolbutton", font=(FONT, F_SMALL), padding=(8, 3),
                 background=CARD, foreground=TEXT2, bordercolor=BORDER,
                 lightcolor=CARD, darkcolor=CARD, anchor="center")
    st.map("Toolbutton",
           background=[("selected", TINT_HEAD), ("active", "#f0f2f5")],
           foreground=[("selected", TINT_HEAD_FG), ("active", TEXT)])

    # ── 表格 ────────────────────────────────────────────────────────────
    st.configure("Treeview", background=CARD, fieldbackground=CARD,
                 foreground=TEXT, font=(FONT, F_SMALL), rowheight=ROW_H,
                 bordercolor=BORDER, lightcolor=CARD, darkcolor=CARD,
                 borderwidth=1, relief="flat")
    st.map("Treeview",
           background=[("selected", TINT_BAND)],
           # foreground 故意不给 ("selected", x) 映射（2026-11-04 第六轮实测）：
           # 只要样式链里任何一层存在 foreground 的 selected 映射，它就**压过**
           # 行 tag 的判定色 —— 红字「不一致」一被选中就变蓝/变黑，颜色语义全丢。
           # 清掉之后：底色仍按下面的 selected 走，字色一律由 tag（红/绿/橙）说了算。
           foreground=[])
    # 表头 2026-11-04 第九轮：底色从浅蓝 tint 收到中性 FIELD，字从亮蓝收到次要灰。
    # 用户原话：「颜色采用地点有点多了，有点晃眼，可以略微减少一点」。
    # 表头只需要「比数据淡一档、读得出是列名」，不必再占一层蓝 —— 蓝留给
    # 主按钮、分组条（那才是「有待处理」的信号）和选中带。
    st.configure("Treeview.Heading", background=FIELD,
                 foreground=TEXT2, font=(FONT, F_SMALL, "bold"),
                 relief="flat", borderwidth=1, bordercolor=BORDER,
                 lightcolor=FIELD, darkcolor=FIELD, padding=(4, 4))
    st.map("Treeview.Heading", background=[("active", "#e6ebf2")])

    # ── 下拉框 ──────────────────────────────────────────────────────────
    st.configure("TCombobox", font=(FONT, F_BODY), padding=(6, 3),
                 fieldbackground=CARD, background=CARD, foreground=TEXT,
                 arrowcolor=TEXT2, bordercolor=BORDER,
                 lightcolor=CARD, darkcolor=CARD, selectbackground=CARD,
                 selectforeground=TEXT)
    st.map("TCombobox",
           fieldbackground=[("readonly", CARD), ("disabled", BG)],
           foreground=[("readonly", TEXT), ("disabled", MUTED)],
           selectbackground=[("readonly", CARD)],
           selectforeground=[("readonly", TEXT)])
    # 下拉列表本体（Tk 的 *TCombobox*Listbox，只能走 option database）
    try:
        root.option_add("*TCombobox*Listbox.background", CARD)
        root.option_add("*TCombobox*Listbox.foreground", TEXT)
        root.option_add("*TCombobox*Listbox.selectBackground", TINT_BAND)
        root.option_add("*TCombobox*Listbox.selectForeground", TINT_BAND_FG)
        root.option_add("*TCombobox*Listbox.font", (FONT, F_BODY))
    except Exception:        # noqa: BLE001 —— 装不上只是下拉列表用系统色
        pass

    # ── 滚动条：细、浅底、悬停才深一点（不抢内容的视线）─────────────────
    st.configure("TScrollbar", background=BG, troughcolor=BG,
                 bordercolor=BG, arrowcolor=TEXT2, gripcount=0,
                 arrowsize=12, lightcolor=BG, darkcolor=BG)
    st.map("TScrollbar", background=[("active", "#c9d2dc")])

    # ── 「学习记录」那个弹窗用的页签 ────────────────────────────────────
    # 选中页签 = 卡片白（跟内容连成一片），未选中 = 窗口底（显得在后面）。
    st.configure("TNotebook", background=BG, bordercolor=BORDER,
                 tabmargins=(0, 2, 0, 0))
    st.configure("TNotebook.Tab", font=(FONT, F_SMALL), padding=(12, 5),
                 background=BG, foreground=TEXT2, bordercolor=BORDER,
                 lightcolor=BG, darkcolor=BG)
    st.map("TNotebook.Tab",
           background=[("selected", CARD), ("active", TINT_HEAD)],
           foreground=[("selected", TEXT), ("active", TEXT)],
           lightcolor=[("selected", CARD)],
           darkcolor=[("selected", CARD)])

    return st


def card(widget):
    """把任意 tk 容器的底色刷成卡片白（tk 控件不吃 ttk 样式，只能逐个设）。"""
    try:
        widget.configure(bg=CARD)
    except tk.TclError:
        pass
    return widget


__all__ = (
    "APP_TITLE", "FONT",
    "BG", "CARD", "DZ_BG", "DZ_BG_ON", "CANVAS", "STATUS",
    "BORDER", "TEXT", "TEXT2", "MUTED", "DISABLED",
    "PRIMARY", "PRIMARY_D", "WARN", "DANGER",
    "TINT_BAND", "TINT_BAND_FG", "TINT_HEAD", "TINT_HEAD_FG", "FIELD", "SEP",
    "OK_FG", "OK_FG2", "BAD_FG", "BAD_FG2", "WARN_FG", "WARN_FG2",
    "NA_FG", "NA_FG2",
    "F_TITLE", "F_APP", "F_BODY", "F_SMALL", "F_TINY",
    "ROW_H", "LIST_ROWS", "install", "card",
)
