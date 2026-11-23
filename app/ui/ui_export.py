"""导出与留档：xlsx 报告、JSON 交接包、批量报告、快照。

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

import json
import os
from tkinter import filedialog, messagebox
from datetime import datetime

from paths import data_dir
from ui_const import APP_VERSION, SCOPE_ORDER
from ui_state import _CODE_STAMP_FILE, _is_real_session, _stamp_text
from ui_fmt import _flag_cn, _fmt_expected, _row_label, _unpassed_hint, _unpassed_rate, _value_str, _yn_cn, scope_buckets, scope_note, scope_summary_lines
from ui_handoff import _HANDOFF_ENABLE, _HANDOFF_FULL_CAP, _HANDOFF_INBOX, _HANDOFF_ROOT, _HANDOFF_SCHEMA, _focus_reason, _handoff_focus_item, _handoff_item, _handoff_task_block, _safe_name

import wi_core as exporter


class ExportMixin:
    """导出与留档：xlsx 报告、JSON 交接包、批量报告、快照。"""


    # ------------------------------------------------------------ 导出
    def _ask_save_path(self, title, default_name, filetypes):
        """弹「另存为」让用户挑导出位置 —— 不再默认塞进软件自己的文件夹。

        返回用户选定的完整路径；点「取消」返回 None（调用方据此跳过导出）。
        默认目录记在 self._last_export_dir 上，下次再导还从那儿起。
        """
        start = getattr(self, "_last_export_dir", "") or os.path.expanduser("~")
        path = filedialog.asksaveasfilename(
            title=title, initialdir=start, initialfile=default_name,
            defaultextension=filetypes[0][1], filetypes=filetypes)
        if not path:
            return None
        self._last_export_dir = os.path.dirname(path)
        return path



    def _export(self):
        # 界面上显示什么就导出什么：刚做过表格自检就导自检报告，否则导比对明细。
        if self.last_mode == "batch" and self.batch_report:
            self._export_batch()
            return
        if self.last_mode == "audit" and self.audit_results:
            self._export_audit()
            return
        if not self.result:
            messagebox.showinfo("提示", "请先执行「开始比对」或「批量核对」。")
            return
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = self._ask_save_path(
            "导出比对报告", f"比对报告_{ts}.txt",
            [("文本文件", "*.txt"), ("所有文件", "*.*")])
        if not out:
            return
        r = self.result
        src = self.schema.chosen.label if self.schema and self.schema.chosen else "未识别"
        with open(out, "w", encoding="utf-8") as f:
            c = r["counts"]
            f.write("WI 一致性检查报告\n")
            f.write(f"生成时间：{datetime.now():%Y-%m-%d %H:%M:%S}\n")
            f.write(f"识别文档：{self.wi_name}\n")
            f.write(f"Excel 数据来源：{src}（工作表!表头行）\n")
            f.write(f"工具版本：{APP_VERSION}    引擎：{_stamp_text()}（{_CODE_STAMP_FILE}）\n")
            if getattr(self, "code_warning", ""):
                f.write(self.code_warning + "\n")
            # 措辞与窗口抬头、状态栏保持同一个词（2026-10-30 统一为「找不到对应」）。
            # 原来写「无法匹配（已读到 PDF，只是没对上）」—— 实测有 PDF 里确实
            # 没有那一项的情况（WI-0039 正文说「如下表7」却没有表7），
            # 「只是没对上」等于替文件开脱。
            f.write(f"核对结果：规格参数一致 {c['ok']} / 不一致 {c['mismatch']} / "
                    f"多套候选(需人工确认) {c.get('multi', 0)} / "
                    f"找不到对应（PDF 文字层正常，是这一项没对上）{c['no_match']}\n")
            # 界面上有这句提醒，留档里就必须有 —— 报告是要给别人看的，
            # 「这批结果不能照着改」这句话漏了，比漏任何一条明细都危险。
            _hint = _unpassed_hint(c)
            if _hint:
                f.write(_hint + "\n")
            _sct = getattr(self, "pdf_selfcheck", None) or {}
            if (_sct.get("lines") or []):
                f.write("抽取判词：" + _sct["lines"][0].lstrip("").rstrip() + "\n")
                if _sct.get("suspect") and c["no_match"]:
                    f.write("  （疑点与本次 %d 条「无对应」要一起看：疑点所在处的项会落到"
                            "「无对应」而不是「不一致」）\n" % c["no_match"])
            judged = [v for v in r["checks"] if v["status"] in ("ok", "mismatch")]
            b = scope_buckets(judged)
            for i, line in enumerate(scope_summary_lines(judged)):
                # 第一行本身就带「判定强度：」前缀（scope_summary_lines 的头），
                # 这里不能再拼一次 —— 之前导出的报告里那行是
                # 「判定强度：判定强度：硬判定…」，看了让人怀疑报告本身有没有问题。
                f.write(("" if i == 0 else "  ") + line.lstrip("　· ") + "\n")
            for k, name in SCOPE_ORDER:
                if b.get(k):
                    f.write(f"  · {name}：{b[k]} 条\n")
            if b.get("other"):
                f.write(f"  · 未标注：{b['other']} 条\n")
            if c.get("no_scope"):
                f.write(f"  · 无设备线索未判定：{c['no_scope']} 条 —— "
                        f"Excel 该行无设备号，期望值收不窄\n")
            f.write(f"未参与核对：采集项 {c['collection']} 条 / 未填值 {c['novalue']} 条\n")
            f.write(f"PDF 参数来源：程序附表 {len(getattr(self, 'pdf_programs', []) or [])} 点"
                    f" / 横表 {len(getattr(self, 'pdf_steps', []) or [])} 点"
                    f" / 千分尺·针规 {sum(len(v) for v in (getattr(self, 'pdf_micrometer', None) or {}).values() if isinstance(v, (list, set)))} 点"
                    f" / 竖表(带适用范围) {len(getattr(self, 'pdf_vertical', []) or [])} 块"
                    f" / 表头元信息 {len(getattr(self, 'pdf_meta', []) or [])} 项\n")
            sc_result = getattr(self, "pdf_selfcheck", None)
            if sc_result and sc_result.get("lines"):
                f.write("抽取自检：" + ("发现 %d 处疑点" % sc_result["suspect"]
                                       if sc_result.get("suspect") else "未发现结构性疑点") + "\n")
                for line in sc_result["lines"]:
                    f.write("  " + line + "\n")
            if getattr(self, "family", None):
                f.write(f"产品线识别：{'、'.join(self.family)}\n")
            warn = getattr(self, "pdf_text_warning", "")
            if warn:
                f.write(warn + "\n")
            f.write("\n—— 规格参数明细 ——\n")
            hints = getattr(self, "sibling_hints", {}) or {}
            for v in r["checks"]:
                flag = exporter.status_cn_of(v)
                if v["status"] == "ok":
                    flag = scope_note(v).lstrip("")
                elif v["status"] == "no_match" and v.get("scope") == "name_unknown":
                    # 与窗口核查列、导出 xlsx 的「判定档位」保持同一句话 ——
                    # 报告里只写「无对应」的话，看报告的人还是会以为是 PDF 里没有。
                    flag = "找不到对应(参数名未收录)"
                elif v["status"] == "no_match" and v.get("scope") == "no_step":
                    flag = "找不到对应(PDF 该表里没有这一步)"
                elif v["status"] == "mismatch" and v.get("meta_kind") and v.get("scope") == "no_equip_match":
                    flag = "不一致(PDF 里查无此设备)"
                elif v["status"] == "mismatch" and v.get("scope") == "prog_order":
                    flag = "不一致(与 PDF 步骤顺序不符)"
                elif v["status"] == "mismatch" and v.get("scope") == "missing":
                    flag = "不一致(漏填：PDF 该步骤有值，此处空白)"
                elif v["status"] == "mismatch" and v.get("scope") == "blank_filled":
                    flag = "不一致(PDF 该格本身是空值，不该填)"
                elif v["status"] == "mismatch" and v.get("scope") == "ref_filled":
                    flag = "不一致(PDF 该格是 ref 参考值，按规矩应留空)"
                elif v["status"] == "mismatch" and v.get("scope") == "type_flag":
                    flag = "不一致(类型与是/否开关不符)"
                elif v["status"] == "mismatch" and v.get("one_side_wrong"):
                    flag = "不一致(PDF 是单边约束，Excel 填错了侧)"
                elif v["status"] == "mismatch" and v.get("scope") == "one_sided":
                    flag = "不一致(超出 Excel 所填的一侧限值)"
                exp = _fmt_expected(v)
                f.write(f"{v['card']} | {v['item']}({v['type']}) | 现值 {_value_str(v)}{v['unit'] or ''}"
                        f" | 应为 {exp}{v['unit'] or ''} | {flag}")
                if v.get("basis"):
                    f.write(f" | 依据 {v['basis']}")
                f.write("\n")
                # 匹配依据单独一行（规范第八条）：匹配层级/分数/最终状态/候选/
                # 具体差异 —— 「匹配成功」和「数据一致」必须分开看得见。
                m = v.get("match") or {}
                if m:
                    bits = ["匹配L%d %s" % (m.get("layer") or 0, m.get("state") or "")]
                    if m.get("score"):
                        bits.append("分数 %.0f%%" % (m["score"] * 100))
                    if m.get("table"):
                        bits.append("对应表 " + m["table"])
                    if m.get("candidates"):
                        bits.append("候选 " + "、".join(
                            "%s(L%d %.0f%%)" % (c[0], c[1], c[2] * 100)
                            for c in m["candidates"][:3]))
                    f.write("    ↳ 匹配：%s\n" % " ｜ ".join(bits))
                    if m.get("why"):
                        f.write("      理由：%s\n" % m["why"])
                hint = hints.get((v["card"], v["item"]))
                if hint:
                    txt = "；".join("%s: %s" % (wi, "、".join(vs)) for wi, vs in sorted(hint.items()))
                    f.write(f"    ↳ 本 PDF 无该参数；同名采集项还写在其它 WI 参数卡下：{txt}\n")
        self.status.set(f"报告已导出：{out}")
        messagebox.showinfo("导出成功", f"报告已保存到：\n{out}")



    # -------------------------------------------------- 历史快照 / 导出 xlsx
    def _snapshot(self):
        """与上一次同 WI 的判定签名对比（先读后写），再存下本次。

        顺序不能反 —— 先存就会拿到自己，对比结果永远是「无变化」。
        """
        self.snapshot_diff = None
        self.snapshot_prev_time = ""
        # 自动化测试（s48 / s56 / s37 这些真开窗口跑真文件）会把「历史快照」灌满
        # 测试姿态，那目录是给用户看的「这个 WI 上次/这次有什么变化」，掺了假记录
        # 反而更乱。两道门：
        #   ① 测试脚本显式设 WI_NO_SNAPSHOT=1；
        #   ② 兜底 `_is_real_session()` —— 只有用户双击起来的界面才写（同 save_last）。
        # 真要验证快照对比逻辑，直接单测 exporter 里的纯函数（latest_snapshot /
        # diff_snapshot / save_snapshot），不要靠开界面。
        if os.environ.get("WI_NO_SNAPSHOT"):
            return
        if not _is_real_session():
            return
        try:
            base = data_dir()     # 历史快照落在 数据\历史快照\，不放在程序目录
            prev = exporter.latest_snapshot(base, self.wi_name)
            sig = exporter.build_signature(self.result["checks"])
            if prev and prev.get("sig"):
                self.snapshot_diff = exporter.diff_snapshot(prev["sig"], sig)
                self.snapshot_prev_time = prev.get("time", "")
            exporter.save_snapshot(sig, base, self.wi_name)
        except Exception:
            # 快照只是加分项，坏了不能影响主流程
            self.snapshot_diff = None



    def _compare_sheets(self):
        r = self.result
        c = r["counts"]
        ov = [["一致", c["ok"]], ["不一致", c["mismatch"]],
              ["多套候选(需人工确认)", c.get("multi", 0)],
              ["找不到对应", c["no_match"]],
              ["未参与核对-采集项", c["collection"]], ["未参与核对-未填值", c["novalue"]]]
        # 「工序」原先只在窗口明细表里有，导出里缺 —— 窗口一旦默认折叠这几列，
        # 这个字段就会在成品里彻底消失，所以先把它补进导出。
        # 两个「是/否」导出时各出两列：表里的**现值**与按类型规矩的**应为**。
        # 窗口里是上下两行对比，Excel 里没法上下并排，就并成列 —— 错在哪一列
        # 照样一眼看得到（2026-09-29 与界面版式同步补的）。
        cols = ["行号", "卡片", "采集项", "类型", "应类型", "现值", "应为", "单位", "工序",
                "显示上下限", "应显示上下限", "允许超上下限", "应允许超上下限",
                "判定", "判定档位", "匹配层级", "匹配状态", "匹配分数",
                "匹配对应表", "依据"]
        rows = []
        for v in r["checks"]:
            m = v.get("match") or {}
            rows.append([_row_label(v), v.get("card", ""), v.get("item", ""),
                         v.get("type", ""), v.get("want_type", "") or "",
                         _value_str(v), _fmt_expected(v), v.get("unit", "") or "",
                         v.get("proc", "") or "",
                         _yn_cn(v.get("show_lim")), _flag_cn(v.get("want_show")),
                         _yn_cn(v.get("allow_over")), _flag_cn(v.get("want_over")),
                         self._FLAG.get(v["status"], v["status"]),
                         v.get("scope", "") or "",
                         "L%d" % (m.get("layer") or 0), m.get("state") or "",
                         ("%.0f%%" % (m["score"] * 100)) if m.get("score") else "",
                         m.get("table") or "", v.get("basis", "") or ""])
        return [("概览", ["项目", "数值"], ov), ("比对明细", cols, rows)]



    def _audit_sheets(self):
        sheets = []
        cov = self.coverage
        if cov is not None:
            sheets.append(("覆盖面与版本链", ["级别", "类别", "WI", "问题", "详情"],
                           [["必改" if i.level == "bad" else "确认", i.cat, i.loc, i.what, i.got]
                            for i in cov.issues]))
        for a in self.audit_results or []:
            sheets.append(("单表问题-%s" % a.sheet, ["级别", "类别", "位置", "问题", "实际值"],
                           [["必改" if i.level == "bad" else "确认", i.cat, i.loc, i.what, i.got]
                            for i in a.issues]))
        o = getattr(self, "outliers", None)
        if o is not None and o.outliers:
            sheets.append(("横向一致性",
                           ["级别", "WI", "工序", "采集项", "规格组", "位置", "参数卡",
                            "产品编码", "该格值", "组内多数", "多数条数", "组内总数"],
                           [["零星" if x.level == "bad" else "成组", x.wi, x.proc, x.item,
                             x.grp, x.loc, x.card, x.pn, x.val, x.major, x.major_n, x.total]
                            for x in o.outliers]))
        return sheets



    def _batch_sheets(self):
        """批量核对导出两张表：一张总览（每 WI 一行）、一张全部非一致明细。"""
        rep = self.batch_report
        ov = []
        for it in rep.items:
            c = it.counts or {}
            if it.error:
                src = it.error[:80]
            else:
                src = "%s ｜ %s" % (os.path.basename(it.xlsx), it.sheet)
            ov.append([it.wi or "?", os.path.basename(it.pdf), src,
                       it.verdict(), c.get("records", 0), c.get("ok", 0),
                       c.get("mismatch", 0), c.get("no_match", 0), c.get("novalue", 0)])
        sheets = [("批量总览",
                   ["WI", "PDF 文件", "数据来源 / 错误", "结论", "记录行数",
                    "一致", "不一致", "找不到对应", "未填值"], ov)]
        rows = [[b["wi"], b["card"], b["item"], b["type"], b["value"], b["expected"],
                 b["unit"], b["status"], b["scope"]]
                for it in rep.items for b in it.bad]
        if rows:
            sheets.append(("非一致明细",
                           ["WI", "参数卡", "采集项", "类型", "现值", "应为",
                            "单位", "状态", "判定依据"], rows))
        return sheets



    # -------------------------------------------------- 交接包（给 WorkBuddy）
    def _handoff_payload(self):
        """组装交接包：顶层键与顺序写死（模板见 `_HANDOFF_SCHEMA`），不随数据变。

        结构固定是这份文件的全部价值 —— 对方每次拿到的形状一模一样，
        不用先猜结构再找重点。所以这里用字面量 dict 一次写全，
        不做「有数据才加这个键」那类 conditional 字段。
        """
        r = self.result
        c = r["counts"]
        checks = r["checks"]
        got = _unpassed_rate(c)
        if got:
            rate, n_bad, _total = got
            reason = ("离线判定未通过率 %d%%（%d/%d 条），已过 20%% 告警线 "
                      "—— 这批结果不能照着改表" % (round(rate * 100), n_bad, _total))
            rate_txt = "%d%%" % round(rate * 100)
            alarm = "是"
        else:
            reason = "未达 20% 告警线（或已判定不足 10 条），由人工主动发起复核"
            rate_txt = "未达告警线"
            alarm = "否"

        full_items = []
        focus_items = []
        for i, v in enumerate(checks, 1):
            full_items.append(_handoff_item(v, i))
            why = _focus_reason(v)
            if why:
                # 「序号」是重点清单内的连续号，「原序号」指回「完整条目」的位置
                focus_items.append(_handoff_focus_item(v, len(focus_items) + 1, i, why))

        cut = "否"
        if len(full_items) > _HANDOFF_FULL_CAP:
            cut = ("是（只留前 %d 条，共 %d 条；「原序号」仍指原始位置。"
                   "重点标注项不受影响，是全的）" % (_HANDOFF_FULL_CAP, len(checks)))
            full_items = full_items[:_HANDOFF_FULL_CAP]

        judged = (c["ok"] + c["mismatch"] + c["no_match"] + c.get("no_scope", 0)
                  + c.get("multi", 0))
        return {
            "模板版本": _HANDOFF_SCHEMA,
            "任务说明": _handoff_task_block(reason),
            "路径信息": {
                "WI编号": self.wi_name or "",
                "WI_PDF路径": self.pdf_var.get(),
                "参数表Excel路径": self.xlsx_var.get(),
                "使用工作表": (self.schema.chosen.label
                               if self.schema and self.schema.chosen else "未识别"),
                "导出时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "离线引擎版本": _stamp_text(),
            },
            "核对统计": {
                "已判定总数": judged,
                "一致": c["ok"],
                "不一致": c["mismatch"],
                "多套候选_需人工确认": c.get("multi", 0),
                "无对应": c["no_match"],
                "无法判定_缺设备线索": c.get("no_scope", 0),
                "未参与核对_采集项": c["collection"],
                "未参与核对_未填值": c["novalue"],
                "未通过率": rate_txt,
                "未通过率是否触发告警": alarm,
                "重点标注条目数": len(focus_items),
                "重点占比": ("%d%%" % round(len(focus_items) * 100.0 / judged)
                             if judged else "0%"),
                "完整条目是否截断": cut,
            },
            "完整条目": full_items,
            "重点标注项": focus_items,
        }



    def _export_json(self):
        """导出交接包 —— 只在这个按钮被点时才写文件。

        批量模式不出交接包：批量一次是几十份 WI，合并成一份就没有「哪份要复核」
        这个焦点了；要复核哪份，就在单文件比对里导那份。
        """
        if self.last_mode == "batch" and self.batch_report:
            messagebox.showinfo(
                "提示", "批量模式不导出交接包：一次批量涉及多份 WI，\n"
                        "合并成一份就分不清要复核哪份。\n\n"
                        "请先用「开始比对」核单份 WI，再导出 JSON。")
            return
        if not self.result:
            messagebox.showinfo("提示", "请先执行「开始比对」。")
            return
        payload = self._handoff_payload()
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        name = "交接JSON_%s_%s.json" % (_safe_name(self.wi_name, "未知WI"), ts)
        out = self._ask_save_path(
            "导出交接包 JSON", name,
            [("JSON 文件", "*.json"), ("所有文件", "*.*")])
        if not out:
            return
        try:
            # indent=2 + 不转义中文：这份文件既给程序读，也要能让人直接打开看
            with open(out, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("导出失败", "写入 JSON 时出错：\n%s" % e)
            return
        # 自动投递没开（_HANDOFF_ENABLE=False）：现在只写本地，不碰 inbox。
        # 开关打开后进这一支，同一份文件再往 ~/WI交接/inbox 放一份。
        sent = ""
        if _HANDOFF_ENABLE:
            try:
                box = os.path.join(os.path.expanduser("~"), _HANDOFF_ROOT, _HANDOFF_INBOX)
                os.makedirs(box, exist_ok=True)
                with open(os.path.join(box, "task_%s_%s.json"
                                       % (_safe_name(self.wi_name, "未知WI"), ts)),
                          "w", encoding="utf-8") as f:
                    json.dump(payload, f, ensure_ascii=False, indent=2)
                sent = "\n\n（已同时放入 %s）" % box
            except Exception as e:
                sent = "\n\n（放入 inbox 失败：%s）" % e
        n_focus = payload["核对统计"]["重点标注条目数"]
        # 0 条重点时必须说破：全对还照样导一份让人「交给 WorkBuddy」，
        # 等于让人白跑一趟 —— 导出什么结论就得陪什么结论。
        if n_focus:
            msg = ("交接包已保存到：\n%s\n\n重点标注 %d 条（占已判定 %s），\n"
                   "请把这份 JSON 交给 WorkBuddy，让它优先复核这些条目。"
                   % (out, n_focus, payload["核对统计"]["重点占比"]))
        else:
            msg = ("交接包已保存到：\n%s\n\n本次离线判定全部对得上，重点标注 0 条 ——\n"
                   "这种情况通常不需要再交给 WorkBuddy 复核。" % out)
        self.status.set("已导出 JSON：%s（重点 %d 条）" % (name, n_focus))
        messagebox.showinfo("导出成功", msg + sent)



    def _export_audit(self):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = self._ask_save_path(
            "导出表格自检报告", f"表格自检报告_{ts}.txt",
            [("文本文件", "*.txt"), ("所有文件", "*.*")])
        if not out:
            return
        with open(out, "w", encoding="utf-8") as f:
            f.write("工程参数表自检报告\n")
            f.write(f"生成时间：{datetime.now():%Y-%m-%d %H:%M:%S}\n")
            f.write("检查范围：Excel 自身一致性（类型/必填/上下限/设备号/组序号/填写规范）、"
                    "覆盖面与 WI 版本链、同规格横向一致性；不读取 PDF 内容\n")
            f.write(f"工具版本：{APP_VERSION}    引擎：{_stamp_text()}（{_CODE_STAMP_FILE}）\n")
            f.write("=" * 62 + "\n\n")
            if self.coverage is not None:
                f.write(self.coverage.text() + "\n\n")
                f.write("=" * 62 + "\n\n")
            for r in self.audit_results:
                f.write(r.text() + "\n\n")
            if getattr(self, "outliers", None) is not None:
                f.write(self.outliers.text() + "\n\n")
        self.status.set(f"报告已导出：{out}")
        messagebox.showinfo("导出成功", f"报告已保存到：\n{out}")



    def _export_batch(self):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = self._ask_save_path(
            "导出批量核对报告", f"批量核对报告_{ts}.txt",
            [("文本文件", "*.txt"), ("所有文件", "*.*")])
        if not out:
            return
        with open(out, "w", encoding="utf-8") as f:
            f.write(f"生成时间：{datetime.now():%Y-%m-%d %H:%M:%S}\n")
            f.write(f"工具版本：{APP_VERSION}    引擎：{_stamp_text()}（{_CODE_STAMP_FILE}）\n")
            f.write("=" * 62 + "\n")
            f.write(self.batch_report.text() + "\n")
        self.status.set(f"报告已导出：{out}")
        messagebox.showinfo("导出成功", f"报告已保存到：\n{out}")
