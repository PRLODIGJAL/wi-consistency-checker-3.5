"""WI 一致性检查工具 —— 桌面窗口(Tkinter)。

双击「启动」脚本运行，完全离线。添加 WI 作业指导书 PDF 与工程参数表 Excel
（支持一次多选 / 直接从访达拖进来），点击「开始比对」，查看覆盖差异与
温度/时间数值核对结果。

Excel 侧已做「结构自适应」：
- 不写死 sheet 名、不写死列号，按列名语义自动定位采集项表与字段；
- 识别不出或筛选后 0 条时明确报出原因，不再静默显示 0。
  界面只保留「加文件（多选/拖拽）→ 开始比对 → 导出」这条主链路
  （2026-10-13 起把「批量核对」并入「开始比对」：清单里多份时自动走批量）。

文件选择会记住上一次用的 PDF / Excel / 工作表（存在同目录的「上次选择.json」）；
若上次的文件已被删除，对应栏留空，由用户重新选择。
"""

import json
import os
import re
import sys
import unicodedata
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk
from types import MappingProxyType

# ── 「把文件拖进来」（可选能力，2026-10-13 用户要求）──────────────────────
# 拖放是 tkinter 自己没有的能力，靠第三方库 tkinterdnd2 提供。它**不是必需品**：
# 没装 / 装了但平台的 tkdnd 二进制对不上 Tcl/Tk 时，拖放静默不启用，
# 「＋ 添加文件…」多选照样能用。参考实现见 PDF转换拼接3.0（同一套写法）。
try:
    import tkinterdnd2 as _dnd
    from tkinterdnd2.TkinterDnD import require as _dnd_require
    DND_FILES = _dnd.DND_FILES
except Exception:  # noqa: BLE001 —— 没装就退化为「无拖放」，界面照常
    _dnd = None
    _dnd_require = None
    DND_FILES = None

# ── 目录四层自举（2026-11-03 目录改造）────────────────────────────────────
# 本文件在 app/ui/ 里。脚本方式启动时 sys.path[0] 就是 app/ui/，
# 扁平模块名（wi_core / paths / anchors …）在那儿都找不到。
# 口诀：先让 app/ 进搜索路径 → import paths → paths.install() 把四层全挂上（幂等）。
import sys as _sys
_APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _APP_DIR not in _sys.path:
    _sys.path.insert(0, _APP_DIR)
import paths as _paths                                   # noqa: E402
_paths.install()


# ── 界面层已按职责拆成四个模块（2026-11-03 第二轮）──────────────────
# 常量 / 纯函数 / 本地状态 / 交接包各自独立，这里一次导入，下面的
# App 类与 main() 照旧用原来的名字，调用点一行都没改。
from ui_const import (_FONT, _VALUE_RE, WI_NAME, AUTO_SHEET, DEFAULT_SHEET_ALIASES,
    APP_VERSION, _SUB_TAG, _SEP, SCOPE_ORDER, _SCOPE_ORDER_MAP, HARD_SCOPES,
    BAD_SCOPES, _BAD_RATE_TRIG, _BAD_RATE_MIN_N)
from ui_state import (_CODE_DIRS, _code_stamp, _CODE_STAMP, _CODE_STAMP_FILE,
    _stamp_text, code_is_stale, CONFIG_PATH, load_last, _is_real_session, save_last,
    _pick_dir, _resolve)
from ui_fmt import (_looks_like_value, preferred_sheet, drop_paths, _walk_widgets,
    _short_names, _disp_width, _fit, _join_card_item, _fmt, _one_sided_side_cn,
    _row_num, _row_label, _yn_cn, _flag_cn, _value_str, _fmt_by_equip,
    _fmt_meta_expected, _ref_suffix, _fmt_dp, _fmt_exp_pair, _fmt_expected,
    _fmt_expected_brief, scope_note, scope_buckets, scope_summary_lines,
    _plain_verdict, _unpassed_rate, _unpassed_hint, _norm_item, sibling_value_hints)
from ui_handoff import (_HANDOFF_ENABLE, _HANDOFF_ROOT, _HANDOFF_INBOX,
    _HANDOFF_OUTBOX, _HANDOFF_DONE, _HANDOFF_SCHEMA, _HANDOFF_FULL_CAP,
    _handoff_should, _handoff_scope_cn, _focus_reason, _handoff_item,
    _handoff_focus_item, _handoff_task_block, _safe_name)

import ui_theme as theme      # 视觉规格：色板 / 字号 / ttk 样式（见 ui_theme.py）

import wi_core      # 引擎 + 解析 + 自检 + 批量 + 导出（8 个模块已合并为这一个文件）
from paths import data_path, data_dir   # 唯一一处「该往哪写文件」的答案（见 paths.py）
try:
    # 结构自检：把「一个正确答案必然满足的形状」写成不变量，跑完就查。
    # 纯函数，只吃 checks；缺了只是不做自检，绝不影响比对本身。
    from wic import audit as _audit
except Exception:       # pragma: no cover
    _audit = None
try:
    import name_lex     # 参数名词库（学习闭环）。文件缺失/损坏时降级为「无词库」，不影响比对
except Exception:       # pragma: no cover
    name_lex = None
try:
    # 行号翻译：工具读的是「展开表」（rev06），使用者核对时打开的是「表4」，
    # 两个行号天生不相等。这个模块按采集项名在表4 里反查行号（只读、不参与判定）。
    import xl_locate
except Exception:       # pragma: no cover
    xl_locate = None
try:
    # 表目录：每条 PDF 参数名出自「第几页、文档印的表几、这张表叫什么」。
    # 用户原话「先选中是哪一个表，再提取这个表里的字段」「加一个它是第几页的表几」。
    # 只读、只用于显示和人工指定，不参与任何判定。
    import pdf_tables
except Exception:       # pragma: no cover
    pdf_tables = None
try:
    # 锚点库（2026-11-02）：把「Excel 这一行」钉在「PDF 这张表这一格」上。
    # 与词库的分工：词库记「名字怎么写」，锚点记「值在哪一格」。
    # 名字对不上时，锚点是唯一还能救回来的通道（见 anchors.py 文件头）。
    import anchors
except Exception:       # pragma: no cover
    anchors = None
# 下面三个别名让原有的调用写法（engine.xxx / batch.xxx / exporter.xxx）保持不变，
# 因此全文调用点一字未改。
engine = wi_core
batch = wi_core
exporter = wi_core


# ── 界面主体：App 由七个混入拼成（2026-11-03 第二轮拆分）──────────────
# 方法体一个字没改、只是搬了家；这里按「越基础越靠前」排，读起来顺。
from ui_lifecycle import LifecycleMixin      # 构造 / 配置落盘 / 关窗
from ui_layout import LayoutMixin            # 整窗骨架 / 分栏 / 两页切换
from ui_pick import PickMixin                # 两个文件池 / 拖放
from ui_run import RunMixin                  # 比对主链路 / 自检
from ui_table import TableMixin              # 明细表 / 概览绘制
from ui_dialog import DialogMixin            # 五个弹窗
from ui_export import ExportMixin            # 导出 / 留档


class App(LifecycleMixin, PickMixin, LayoutMixin, RunMixin,
          TableMixin, DialogMixin, ExportMixin):
    """窗口本体。方法不在这个类里，在 app/ui/ui_*.py 的七个混入中。

    2026-11-03 第二轮拆的：原来这里是 4420 行一个类，光 `_render` 一个方法
    就 660 行，改任何一处都要在几千行里上下翻。现在按职责分家，这个类只留
    两样东西 —— 类级常量，和「我是谁」的继承清单。

    行为没有任何变化：方法体一字未改，`self.xxx` 的解析顺序也没变，
    构造一个 App 得到的对象与拆分前逐属性相同。
    """

    # 大框的列号（见 `_build_ui`）：0=左竖条 1=左清单 2=分隔 3=右清单 4=右竖条
    #   2026-10-29 把「左栏 | 拖放框 | 右栏」换成「左栏 | 分隔 | 右栏」时，
    #   有人还在 `columnconfigure(0, ...)` —— 结果把「左竖条」那一列撑成 430、
    #   右栏压成 0，界面上看是「右边没了」。
    #   纯坐标错位靠读代码看不出来，所以列号一律走常量，不再写魔数。
    #   2026-11-03 起两栏宽度按 POOL_W_L : POOL_W_R = 8 : 2 由 weight 分配。
    _COL_VBAR_L, _COL_LIST_L, _COL_LINE, _COL_LIST_R, _COL_VBAR_R = 0, 1, 2, 3, 4

    # 只读视图：仍指向唯一来源 exporter.STATUS_CN，但任何误写会立刻抛 TypeError，
    # 把「静默串改批量/导出的中文文案」变成「当场报错」。类里只用 .get()，兼容。
    _FLAG = MappingProxyType(exporter.STATUS_CN)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
