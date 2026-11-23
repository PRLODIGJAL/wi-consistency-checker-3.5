"""核对这条主链路：读工作表 → 比对 → 批量 → 自检 → 结论卡。

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
from tkinter import messagebox

import ui_theme as theme
from ui_const import AUTO_SHEET, WI_NAME
from ui_state import code_is_stale
from ui_fmt import preferred_sheet, sibling_value_hints

import wi_core
engine = wi_core
batch = wi_core

try:
    from wic import audit as _audit
except Exception:      # 缺了就降级，不影响核对
    _audit = None

try:
    import name_lex
except Exception:      # 缺了就降级，不影响核对
    name_lex = None

try:
    import pdf_tables
except Exception:      # 缺了就降级，不影响核对
    pdf_tables = None

try:
    import anchors
except Exception:      # 缺了就降级，不影响核对
    anchors = None



class RunMixin:
    """核对这条主链路：读工作表 → 比对 → 批量 → 自检 → 结论卡。"""


    # ------------------------------------------------------------ 文件选择
    # （原 _pick_pdf / _pick_xlsx 单文件选择已删除 —— 2026-10-13 起改用
    #   多选清单 _pick_many / 拖拽 _on_drop，见「多选清单」一节。）

    def _load_sheets(self, restore=None):
        """打开表格后自动列出它的工作表，供手动指定。

        默认值的优先级：
          ① 上次对**同一个文件**选过的表（记忆功能，换文件时不沿用）；
          ② 换文件 / 首次打开 → **表4**（模板统一后的参数卡配置表）；
          ③ 该文件里没有表4 → 退回「自动识别」（按列名完整度打分挑一张）。
        下拉里始终列出全部工作表，用户随时可改选 —— 这里只影响默认值。
        """
        p = self.xlsx_var.get().strip()
        if not os.path.exists(p):
            return
        try:
            names = engine.list_sheets(p)
        except Exception:
            names = []
        self.sheet_box["values"] = [AUTO_SHEET] + names
        if restore and restore != AUTO_SHEET and restore in names:
            self.sheet_var.set(restore)
        else:
            self.sheet_var.set(preferred_sheet(names) or AUTO_SHEET)



    def _sheet_arg(self):
        v = self.sheet_var.get()
        return None if v == AUTO_SHEET else v



    # ------------------------------------------------------------ 结构自检
    def _inspect(self):
        """结构自检：Excel 看表结构，PDF 看版式。只描述识别结果，不含数据值。"""
        xlsx = self.xlsx_var.get().strip()
        pdf = self.pdf_var.get().strip()
        if not os.path.exists(xlsx) and not os.path.exists(pdf):
            messagebox.showerror("文件不存在", "请先选择工程参数表 Excel 或 WI 作业指导书 PDF。")
            return
        self.status.set("正在识别结构…")
        self.root.update_idletasks()
        parts = []
        if os.path.exists(xlsx):
            try:
                ts, txt = engine.inspect_excel(xlsx, sheet=self._sheet_arg())
                self.schema = ts
                parts.append(txt)
            except Exception as e:
                parts.append("Excel 结构自检失败：\n%s" % e)
        if os.path.exists(pdf):
            try:
                _recs, txt = engine.inspect_pdf(pdf)
                parts.append("")
                parts.append(txt)
            except Exception as e:
                parts.append("\nPDF 版式自检失败：\n%s" % e)
        self._show("\n".join(parts))
        self.status.set("结构自检完成")



    # ------------------------------------------------------------ 表格自检
    def _audit(self):
        """表格自检：只用 Excel，不读 PDF。

        回答的是「这张表自己前后说不说得通」—— 类型与取值是否相符、必填项有没有
        预设值、上下限有没有写反、设备号有没有绑定、同名采集项的类型是否跨类。
        与「结构自检」的分工：那个说「这张表我能不能读」，这个说「它说的对不对」。
        """
        xlsx = self.xlsx_var.get().strip()
        if not os.path.exists(xlsx):
            messagebox.showerror("文件不存在", "请先选择工程参数表 Excel。")
            return
        self.status.set("正在自检表格…")
        self.root.update_idletasks()
        try:
            rs = engine.audit_excel(xlsx, sheet=self._sheet_arg())
        except Exception as e:
            self._show([("表格自检失败：", "bad"), (str(e), None)], alert=True)
            self.status.set("表格自检失败")
            return

        # 覆盖面与版本链：要用到 PDF。PDF 没选也照跑（那时只报数据侧）。
        pdf = self.pdf_var.get().strip()
        has_pdf = os.path.exists(pdf)
        cov = None
        try:
            cov = engine.audit_coverage(
                xlsx,
                pdf_paths=[pdf] if has_pdf else [],
                scan_dir=os.path.dirname(pdf) if has_pdf else None)
        except Exception as e:
            self._cov_error = str(e)

        self.audit_results = rs
        self.coverage = cov
        self.last_mode = "audit"
        blocks = [("表格自检 · %s" % os.path.basename(xlsx), "t"),
                  ("  %s" % xlsx, "dim"), ("", None)]

        if cov is not None:
            blocks.extend(cov.summary())
        elif has_pdf is False:
            blocks.append(("覆盖面（跳过）", "t"))
            blocks.append(("  未选择 WI 作业指导书 PDF，无法判断"
                           "「录入了数据的 WI 手上有没有文件可核」。", "warn"))
            blocks.append(("", None))
        else:
            blocks.append(("覆盖面（失败）", "t"))
            blocks.append(("  %s" % getattr(self, "_cov_error", ""), "warn"))
            blocks.append(("", None))

        blocks.append(("单表问题", "t"))
        for r in rs:
            blocks.extend(r.summary(with_title=True, with_notes=False))
            s = r.stats or {}
            blanks = s.get("required_blank_sample") or []
            nblank = max(0, s.get("required_expected", 0) - s.get("required_filled", 0))
            if blanks:
                blocks.append(("  必填但未预设值：共 %d 格（列出前 %d 条）"
                               % (nblank, min(len(blanks), 8)), "h"))
                for loc, txt in blanks[:8]:
                    blocks.append(("      %-15s %s" % (loc, txt[:58]), "warn"))
                blocks.append(("", None))
            blocks.append(("", None))
        if rs and rs[0].notes:
            for n in rs[0].notes:
                blocks.append(("  · " + n, "dim"))

        # 横向一致性：找「同类产品里冒出来的格子」。耗时略长（要扫全表），
        # 但它能发现逐行核对发现不了的那类问题，所以放在自检里一起跑。
        outl = None
        try:
            outl = engine.scan_excel_outliers(xlsx, sheet=self._sheet_arg())
        except Exception as e:
            self._outl_error = str(e)
        self.outliers = outl
        if outl is not None:
            blocks.extend(outl.summary())
        else:
            blocks.append(("横向一致性（失败）", "t"))
            blocks.append(("  %s" % getattr(self, "_outl_error", ""), "warn"))
            blocks.append(("", None))
        self._show(blocks)
        n_issue = (sum(len(r.issues) for r in rs)
                   + (len(cov.issues) if cov else 0)
                   + (len(outl.outliers) if outl else 0))
        self.status.set("表格自检完成：%d 条问题" % n_issue)



    # ------------------------------------------------------------ 批量核对
    def _run_batch(self, pdfs, xlsxs):
        """批量核对：对每份 PDF 在**所有给定的参数表里自动找它的数据**。

        由 `_run` 在清单多份时转调过来 —— 参数就是界面清单里那批路径，
        不再弹窗选文件（多选 / 拖拽已经把文件收进清单了）。
        找不到才如实报「本次没提供它的参数表」，不会静默漏掉。
        """
        if not pdfs:
            messagebox.showerror("文件不存在", "请先添加 WI 作业指导书 PDF。")
            return
        if not xlsxs:
            messagebox.showerror("文件不存在", "请先添加工程参数表 Excel。")
            return

        self.status.set("正在批量核对 %d 份 WI…" % len(pdfs))
        self.root.update_idletasks()

        def prog(i, n, txt):
            self.status.set("批量核对 %d/%d：%s" % (i, n, txt[:36]))
            self.root.update_idletasks()

        try:
            rep = batch.run_batch(list(pdfs), list(xlsxs), progress=prog)
        except Exception as e:
            self._show([("批量核对失败：", "bad"), (str(e), None)], alert=True)
            self.status.set("批量核对失败")
            return

        self.batch_report = rep
        self.audit_results = None    # 切模式后「导出报告」认的是这一份
        self.last_mode = "batch"
        self._show(rep.summary())
        n_bad = rep.total("mismatch")
        n_miss = sum(1 for it in rep.items if it.error)
        self.status.set("批量核对完成：%d 个 WI ｜ 不一致 %d 条 ｜ 无数据 %d 个"
                        % (len(rep.items), n_bad, n_miss))



    # ------------------------------------------------------------ 比对
    def _run(self):
        # 2026-10-29 定稿：一次只比**一份参数表 对 一份 WI**。
        # 两栏清单只是备着的（用户原话：「我放进来不代表我要对比」），
        # 真正的对象只有各选中的那一条，所以这里不再有「多份怎么组合」的分支。
        sel_x = self._sel_of("xlsx")
        sel_p = self._sel_of("pdf")
        xlsxs = [p for p in sel_x if os.path.exists(p)]
        pdfs = [p for p in sel_p if os.path.exists(p)]
        if not xlsxs or not pdfs:
            lack, where = [], []
            if not xlsxs:
                lack.append("Excel")
                where.append("左栏")
            if not pdfs:
                lack.append("WI（PDF）")
                where.append("右栏")
            messagebox.showinfo(
                "还没选要比的文件",
                "还要在%s里各点中一份%s。\n\n"
                "清单里可以先放几十份备着，但每次只比较一对：\n"
                "左边选一份参数表、右边选一份 WI，再点「开始比对」。"
                % ("、".join(where), "、".join(lack)))
            self.status.set("请先在%s里各点中一份要比较的文件"
                            % "、".join(where))
            return
        pdf = pdfs[0].strip()
        xlsx = xlsxs[0].strip()
        # 单份比对：这两个变量供交接包/报告取路径，指向**本次**这一份
        self.pdf_var.set(pdf)
        self.xlsx_var.set(xlsx)
        self._save_last()   # 手输/粘贴的路径也一并记住
        # 源码在窗口开着的时候被更新过？先说清楚，免得看着旧代码的结果去改新代码。
        self.code_warning = code_is_stale()
        self.status.set("正在识别 WI 文档…")
        self.root.update_idletasks()
        try:
            detected = engine.detect_wi(pdf)
            scan_w = engine.pdf_scan_warning(pdf)
            if scan_w:
                # 扫描件：能识别的先识别出来（WI 号已按文件名认到），再如实说明
                # 参数为什么核不了 —— 用户的口径（2026-09-27）：「如果说它能够
                # 识别，你试一下先，如果能够识别的话就建议识别一下，不能的话就算了。」
                # 硬跑只会得到「全部无对应」的假结果，所以停下来，但把话说全。
                bits = [scan_w, ""]
                if detected:
                    bits.append("· 已按文件名识别这份 PDF 为 %s；" % detected)
                    # Excel 那边能不能读，顺手也试一下 —— 让人知道问题只出在 PDF。
                    try:
                        recs, _ts = engine.extract_excel_params(
                            xlsx, detected, sheet=self._sheet_arg(), with_report=True)
                        if recs:
                            bits.append("· Excel 侧识别正常（%s 共 %d 条采集项）——"
                                        "问题只出在 PDF 这份扫描件；" % (detected, len(recs)))
                    except Exception:
                        pass
                bits.append("· 参数核对先到这里（PDF 读不出任何参数，硬跑只会得到"
                            "「全部无对应」的假结果）；")
                bits.append("· 想核对参数：建议向文件的提供方要电子版 PDF（导出生成、"
                            "非打印扫描），或先把扫描件 OCR 成带文字层的 PDF 再来。")
                self.tree.delete(*self.tree.get_children())
                self._sync_empty_hint()
                self._show("\n".join(bits), alert=True)
                self.status.set("无法核对：PDF 是扫描件（无文字层）")
                return
            if detected is None:
                # 有文字层但没找到 WI 号：保留老兜底，但**大声说出来**。
                wi_name = WI_NAME
                self.status.set("PDF 里没识别到 WI 编号（将按内置的 %s 尝试）…" % wi_name)
            else:
                wi_name = detected
                self.status.set(f"识别到 {wi_name}，正在读取 Excel 结构…")
            self.root.update_idletasks()
            excel_recs, ts = engine.extract_excel_params(
                xlsx, wi_name, sheet=self._sheet_arg(), with_report=True)
            self.schema = ts
            if not excel_recs:
                self._explain_empty(ts, wi_name)
                return
            self.status.set(f"正在抽取 PDF 参数…（Excel {len(excel_recs)} 条）")
            self.root.update_idletasks()
            engine.reset_pdf_failures()
            pdf_by_name = engine.extract_pdf_params(pdf)          # 兼容视图（按名聚合）
            pdf_vertical = engine.extract_pdf_vertical(pdf)       # 竖表：带适用范围的参数块
            pdf_precise = engine.extract_pdf_params_precise(pdf)
            micrometer = engine.extract_pdf_micrometer(pdf)       # 千分尺/针规：带适用范围
            pdf_steps = engine.extract_pdf_steps(pdf)             # 「参数 | Step N」横表
            pdf_blanks = engine.extract_pdf_step_blanks(pdf)      # 横表里「写着 ------ 的格子」
            pdf_programs = engine.extract_pdf_programs(pdf)       # 「设备参数附表N」程序A/B 双列
            pdf_meta = engine.extract_pdf_meta(pdf)               # 表头元信息（机器编号/程序号/保护套/芯轴）
            family = engine.doc_family(xlsx)                      # 产品线（区分同名附表）
            self.pdf_text_warning = engine.pdf_text_warning(pdf)
            self.status.set(f"核对中…（竖表 {len(pdf_vertical)} 点 / 横表 {len(pdf_steps)} 点"
                            f" / 空格 {len(pdf_blanks)} 格 / 程序附表 {len(pdf_programs)} 点"
                            f" / 表头元信息 {len(pdf_meta)} 项）")
            self.root.update_idletasks()
            self.pdf_by_name, self.excel_recs, self.wi_name = pdf_by_name, excel_recs, wi_name
            self.pdf_vertical = pdf_vertical
            self.pdf_steps = pdf_steps
            self.pdf_blanks = pdf_blanks
            self.pdf_programs = pdf_programs
            self.pdf_micrometer = micrometer
            self.pdf_meta = pdf_meta
            self.family = family
            # 抽取自检：换一份没见过的 PDF 时，「读出来多少、有没有整块漏」得有个交代
            self.pdf_selfcheck = engine.self_check(pdf)
            # 球囊类 WI 判定（2026-10-14）：ref「不填」这条规矩只对球囊 WI 生效。
            # 统一走引擎的 detect_balloon（文件名 + 正文 + 子表标题 + 工艺名），
            # 与批量核对 run_batch 同一口径，避免两处各写一套、结果不一致。
            balloon = engine.detect_balloon(pdf_steps, excel_recs, pdf_path=pdf)
            # 词库注入（只影响「名字对齐」；空词库 = 与没有这个机制时逐条一致）
            if name_lex is not None:
                name_lex.push_to_core(wi_name, family)
            # 产品形态（RX / OTW）：从「表2 大类」取，供附表页维度收窄用。
            # 读不到、或卡片同时含两种形态 → 空集 → 引擎弃权（不收窄）。
            try:
                import page_tier
                engine.set_card_tokens(page_tier.card_tokens(xlsx))
            except Exception:
                engine.set_card_tokens(())
            # 锚点注入（2026-11-02）：人钉过位置的行，走「按坐标读值」那条通道，
            # 不看名字。空锚点 = 与没有这个机制时逐条一致（引擎第一行就 return）。
            # ⚠️ 必须每次都注入 —— 不然上一次比对留下的绑定会漏到这一次，
            # 那就成了「换个文件结论变了，还查不出为什么」。
            if anchors is not None:
                try:
                    anchors.set_pdf(pdf)
                    _acat = (pdf_tables.catalog(pdf, pdf_vertical)
                             if pdf_tables is not None else [])
                    _abinds = anchors.resolve(_acat, wi_name, family)
                    # 表级规则（2026-11-02 学习机制）：人钉过一格之后，工具会在这张表
                    # 上推一条「行看规格、列看方向+部位」的规则。规则推出来的绑定与
                    # 人工锚点**同构**，一起喂给引擎的同一个 _apply_anchors —— 判得对
                    # 不对全靠同一段代码，不存在「学来的」和「手钉的」两套口径。
                    #
                    # 这里传 excel_recs 而不是 checks：比对还没跑，拿不到 status。
                    # 不影响正确性 —— _apply_anchors 只改 no_match / no_scope，
                    # 已经判出来的行规则碰不到。
                    _rules = anchors.resolve_rules(_acat, excel_recs, pdf, wi_name, family)
                    # 人工钉的位置优先于规则推的 —— 但**不能简单 update 掉**：
                    # 那样这一行就只剩人钉的坐标，引擎无从知道「本来规则会把它指到哪」，
                    # 也就说不出「校正」的由来。用户要的正是那句
                    # 「他认为是对上了的，但他又抽取错误了」—— 得让人看见原来读的是哪一格。
                    # 所以把人钉的那条压上去，同时把被它盖住的规则坐标挂在 `_alt` 上，
                    # 由 `_apply_anchors` 决定要不要在依据里交代（见那边的 _alt 注释）。
                    for _k, _bb in _abinds.items():
                        _old = _rules.get(_k)
                        if _old and not _old.get("from_rule"):
                            _old = None     # 保险：只把「规则推的」当被盖住的
                        if _old and (int(_old.get("row") or -1), int(_old.get("col") or -1)) != \
                                (int(_bb.get("row") or -2), int(_bb.get("col") or -2)):
                            _bb["_alt"] = _old
                        _rules[_k] = _bb
                    self.anchor_binds = _rules
                    engine.set_anchors(_rules)
                except Exception:
                    self.anchor_binds = {}
                    engine.set_anchors({})
            else:
                self.anchor_binds = {}
                engine.set_anchors({})
            # 关键字传参：engine.compare 有 10 个位置参数，裸露位置对齐一旦顺序变就静默错位
            self.result = engine.compare(
                pdf_by_name=pdf_by_name,
                excel_records=excel_recs,
                micrometer=micrometer,
                pdf_precise=pdf_precise,
                pdf_steps=pdf_steps,
                pdf_programs=pdf_programs,
                pdf_vertical=pdf_vertical,
                family=family,
                pdf_meta=pdf_meta,
                pdf_step_blanks=pdf_blanks,
                balloon=balloon,
                pdf_path=pdf)
            # 「读到了多少」的底账：用来区分「PDF 没读到」与「读到但匹配不上」。
            # PDF 是标准矢量版式时，界面上那些「无法匹配」绝大多数属于后者。
            self.read_info = self.result.get("read") or {}
            self.sibling_hints = sibling_value_hints(ts, self.result["checks"])
            self.last_mode = "compare"
            self._snapshot()
            self._render()
            # 2026-11-04 第七轮：**不再自动弹「指定位置」窗**（用户原话：
            # 「机器学习它不需要弹窗，就是现在已经很明确的标记出来哪里有问题了，
            #   所以说根本不需要有弹窗这个东西存在」）。
            # 明细表里红/绿已经把「哪条要管」标得清清楚楚，弹窗只会在人刚看完
            # 结果的时候挡一层；需要钉格的时候双击那一行就行（_on_row_dbl）。
        except Exception as e:
            messagebox.showerror("比对失败", f"处理出错：\n{e}")
            self.status.set("比对失败")



    def _explain_empty(self, ts, wi_name):
        """筛选后 0 条 → 明确说明原因，绝不静默。"""
        L = ["没有取到任何采集项，因此没有可比对的内容。", ""]
        if ts.chosen is None:
            L.append("原因：Excel 里没有识别出「采集项明细表」。")
            for e in ts.errors:
                L.append("  · " + e)
            L.append("")
            L.append("请在「使用工作表」里换一张含「数据采集项名称」列的表后重试。")
        elif ts.wi_source == "none":
            L.append("原因：选中的表「%s」里既没有 WI 列，也没有可关联出 WI 的参数卡列。" % ts.chosen.name)
            L.append("  · 请在「使用工作表」里换一张表，或补充 WI 关联表后重试。")
        else:
            known = ts.wi_names()
            L.append("原因：表「%s」里没有本次识别的 WI「%s」。" % (ts.chosen.name, wi_name))
            if known:
                L.append("  · 该表实际包含：%s" % "、".join(known))
                L.append("  · 请确认 PDF 与 Excel 是同一批文件，或换一份 PDF。")
            else:
                L.append("  · 该表无法解析出 WI 编号。")
        self.tree.delete(*self.tree.get_children())
        self._show("\n".join(L), alert=True)
        self.status.set("未取到采集项")



    def _refresh_audit(self):
        """比对完自动跑一遍结构自检，把结论写到状态栏右端。

        为什么放在这里而不是弹窗：绝大多数时候结论是「通过」，不该打断人；
        真有问题时它就在那儿亮着，点一下就开清单。

        为什么要有这一层（说清来路）：2026-11-02 一天里连着三个同类 bug
        —— 2534 多段多套、2535 应为两个值、2535 把物料表当成针规表。
        症状各不相同，机制是同一个：解析器少认一种写法 → 返回 None →
        上层当成「文件里没这个信息」继续走 → **静默变宽**（候选从 1 个变 2 个），
        值仍然对得上，界面上一点错都不报，只能靠人肉发现。
        这几条不变量就是对「它不吭声」的兜底。
        """
        self.audit_result = None
        if _audit is None or not hasattr(self, "audit_var"):
            return
        try:
            checks = (self.result or {}).get("checks") or []
            self.audit_result = _audit.audit(checks)
        except Exception as e:
            # 自检自己坏了，不能影响比对结果 —— 如实说一句就行
            self.audit_var.set("结构自检没跑起来（%s）" % str(e)[:40])
            self.audit_lbl.configure(foreground=theme.TEXT2, cursor="")
            return
        errs = self.audit_result.get("errors") or []
        self.audit_var.set(_audit.summary(self.audit_result))
        if errs:
            self.audit_lbl.configure(foreground=theme.DANGER, cursor="hand2")
        else:
            self.audit_lbl.configure(foreground=theme.TEXT2, cursor="")
