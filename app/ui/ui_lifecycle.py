"""窗口的生死：构造、配置落盘、关窗。

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
from datetime import datetime

import ui_theme as theme
from ui_const import APP_VERSION, AUTO_SHEET, DEFAULT_PDF, DEFAULT_XLSX
from ui_state import _resolve, _stamp_text, load_last, save_last
from ui_fmt import _fit

import wi_core
engine = wi_core


class LifecycleMixin:
    """窗口的生死：构造、配置落盘、关窗。"""


    def __init__(self, root):
        self.root = root
        root.title("WI 一致性检查工具 %s（Mac 版）  ·  引擎 %s"
                   % (APP_VERSION, _stamp_text()))
        root.geometry("1040x760")
        root.minsize(880, 620)
        # 窗口底用三级底色里的第一档（默认的系统灰会跟卡片白打架，
        # 一进来就像「两块颜色拼起来的」）。
        root.configure(bg=theme.BG)

        # 整套 ttk 样式一次性装好：clam 主题 + 按钮三档 + 表头 tint + 滚动条。
        # ⚠️ 必须在建任何控件**之前**调用 —— 2.0 只设了字体，按钮颜色全是
        #    系统默认（Mac 的 aqua 主题根本不认 background），所以界面要么全灰、
        #    要么局部白，出不来「一张米白纸」的层次。
        theme.install(root)

        # 预填：上次选过的文件还在就用它；已被删除则留空；首次运行才用内置默认路径
        last = load_last()
        pdf0 = _resolve(last, "pdf", DEFAULT_PDF)
        xlsx0 = _resolve(last, "xlsx", DEFAULT_XLSX)

        self.pdf_var = tk.StringVar(value=pdf0)
        self.xlsx_var = tk.StringVar(value=xlsx0)
        self.sheet_var = tk.StringVar(value=AUTO_SHEET)
        # 多选清单（2026-10-13 用户要求「一口气多选」）——
        #   取代原来「一份一份选」的单值输入框。两个清单各记各的：
        #   左栏 = 工程参数表 Excel（xlsx_list），右栏 = WI 作业指导书 PDF（pdf_list）。
        #   pdf_var / xlsx_var 保留着，但语义已收窄（2026-10-29）：它们记的是
        #   「这次比对用的是哪一份」，由 `_refresh_pick_list` 跟着选中项走 ——
        #   交接包 `_handoff_payload` 读的就是这两个变量。
        #   恢复时优先用列表键（多份），没有就退回单值键（旧版记忆）。
        xl_prev = [p for p in (last.get("xlsx_list") or []) if os.path.exists(str(p))]
        pd_prev = [p for p in (last.get("pdf_list") or []) if os.path.exists(str(p))]
        self.xlsx_list = xl_prev or ([xlsx0] if xlsx0 else [])
        self.pdf_list = pd_prev or ([pdf0] if pdf0 else [])
        # 本次要比的那一份（2026-10-29 定稿：一对一）——
        #   两栏清单是「备着的」，各可能几十份；sel_* 只存**一份**，
        #   就是这次要拿去比的那一份。放进来 ≠ 要比，所以默认空。
        sel_x = [p for p in (last.get("sel_xlsx") or []) if p in self.xlsx_list][:1]
        sel_p = [p for p in (last.get("sel_pdf") or []) if p in self.pdf_list][:1]
        self.sel_xlsx = sel_x
        self.sel_pdf = sel_p
        # 两栏宽度 2026-11-03 起是**固定 8 : 2**（见 ui_layout.POOL_W_L / POOL_W_R），
        # 不再读记忆里的 split。旧的「上次选择.json」里还可能留着 split 键，
        # 读的时候自然忽略；这里保留属性只为别处（导出/截图脚本）还能取到值。
        self.split_w = 0
        # 显示开关（2026-09-28 起界面上不再给勾选框，恒定取 False）：
        #   show_all_cols —— 界面上恒定 False：默认那 9 列（含 unit）就是全部
        #     要看的东西，proc / basis 两列挪去了「导出 Excel」。留着这个变量
        #     是为了兼容旧记忆文件和回归脚本（`_visible_cols` 还会读它）。
        #   expand_src 已于 2026-10-30 连同【PDF 参数来源】整段一起删掉 ——
        #     那一段在概览里没人看，折叠开关也就跟着没了用处。
        self.show_all_cols = tk.BooleanVar(value=False)
        # show_sep —— 两条数据之间拉一条横线（2026-09-29 用户要求「框起来」）。
        # 默认**开**。代价是每个数据多插一行：大表 2.6 万条 → 8 万行，
        # 先量过再定要不要给界面加开关（见 _run 里的插入循环）。
        self.show_sep = tk.BooleanVar(value=True)
        self._blocks = []           # 摘要区待显示的内容，交给 _paint()
        self.result = None      # 最近一次比对结果
        self.schema = None      # 最近一次结构识别结果
        self.audit_results = None   # 最近一次表格自检结果（[AuditResult]）
        self.coverage = None        # 最近一次覆盖面与版本链结果（CoverageReport）
        self.outliers = None        # 最近一次横向一致性结果（OutlierReport）
        self.batch_report = None    # 最近一次批量核对结果（BatchReport）
        self.snapshot_diff = None   # 与上一次同 WI 比对结果的差异
        self.snapshot_prev_time = ""
        self._cov_error = ""
        self._outl_error = ""
        self.last_mode = ""     # "" / "compare" / "audit" / "batch"，决定「导出报告」导哪份
        self.sibling_hints = {}  # 「无对应」时的跨 WI 同名采集项线索

        self._build_ui()
        root.protocol("WM_DELETE_WINDOW", self._on_close)

        # 启动时清单里还有上次的文件（且文件仍在）→ 替用户选好第一份。
        # 和拖入时是同一套「只补空、不抢」：记忆里本来记着选中的话就不动。
        # 不这么做会出现一个别扭的状态：清单恢复了、变量也指着第一份，
        # 界面上一行都没蓝，点「开始比对」反而被拦下说还没选。
        # ⚠️ 必须在下面 `_load_sheets` **之前**：它按 xlsx_var 去读工作表名，
        #    先选中谁、就读谁的表。
        for _k in ("xlsx", "pdf"):
            self._auto_pick_first(_k)

        # 工作表一有变化就（防抖后）写盘。PDF / Excel 现在走清单，写盘在
        # `_refresh_pick_lists`（清单变化时）里直接调 `_save_last`，不走这里的 trace。
        self._save_job = None
        self._ready = False     # 启动阶段不写盘：否则会把刚恢复的记忆又覆盖掉
        self.sheet_var.trace_add("write", self._schedule_save)

        if os.path.exists(self.xlsx_var.get().strip()):
            self._load_sheets(restore=last.get("sheet"))
        # 拖放：窗口建出来之后才挂得上（见 _setup_dnd）。
        #    挂得上就不提（拖一次就会用，不用教）；挂不上才在状态栏说一句。
        self.dnd_ok = self._setup_dnd()
        self._dnd_note()
        self._hint_restored(last, pdf0, xlsx0)
        self._ready = True



    def _schedule_save(self, *_a):
        """输入变化后延迟写盘：连续输入只落一次，避免频繁碰磁盘。"""
        if not getattr(self, "_ready", False):
            return
        if self._save_job is not None:
            try:
                self.root.after_cancel(self._save_job)
            except Exception:
                pass
        try:
            self._save_job = self.root.after(400, self._save_now)
        except Exception:
            self._save_job = None



    def _save_now(self):
        self._save_job = None
        self._save_last()



    def _on_close(self):
        """关窗时立刻落盘（不等防抖），手输路径也不会丢；顺手把 PDF 缓存还给系统。"""
        if self._save_job is not None:
            try:
                self.root.after_cancel(self._save_job)
            except Exception:
                pass
        self._save_job = None
        self._save_last()
        engine.release_caches()
        self.root.destroy()



    def _save_last(self):
        """把当前选择写盘（含工作表）。清单多份就存成列表，兼容旧的单值读取。"""
        xl = [p for p in self.xlsx_list if p]
        pd = [p for p in self.pdf_list if p]
        save_last({
            # 保留旧的单值键，老版本 / 交接包路径仍能读；新增列表键存多份
            "pdf": pd[0] if pd else "",
            "xlsx": xl[0] if xl else "",
            "pdf_list": pd,
            "xlsx_list": xl,
            # 本次选中的那一份也记住：第二天开着工具，连「比到哪儿了」都还在。
            "sel_pdf": self._sel_of("pdf"),
            "sel_xlsx": self._sel_of("xlsx"),
            # 分栏比例已定死 8:2，不再落盘（旧文件里的 split 键读时忽略）
            "sheet": self._sheet_arg() or "",
            "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })



    def _hint_restored(self, last, pdf0, xlsx0):
        """状态栏如实说明「这次的预填是哪来的」。

        没有这句话时，记忆本来就没生效也看不出来 —— 用户只会觉得
        「上次选的又被忘了」，却分不清是真的没记住，还是上次的文件没了。
        """
        # 文件名截一下：状态栏在按钮那一行右侧，太长了会把按钮挤走
        lost = [_fit(os.path.basename(str(last[k])), 28) for k in ("pdf", "xlsx")
                if last.get(k) and not os.path.exists(str(last[k]))]
        back = [_fit(os.path.basename(p), 28)
                for k, p in (("pdf", pdf0), ("xlsx", xlsx0)) if p and last.get(k) == p]
        if lost:
            self.status.set("上次选择的文件已不存在，该栏已留空：%s" % "、".join(lost))
        elif back:
            self.status.set("已恢复上次选择：%s" % "、".join(back))
        elif pdf0 or xlsx0:
            self.status.set("首次使用，已预填默认文件")
        else:
            self.status.set("先把文件拖进上面两个框（放着备查），"
                            "再在左右清单里各点中一份要比的")
