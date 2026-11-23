"""交给 WorkBuddy 的交接包（JSON）字段与组装。

从 app.py 拆出来的（2026-11-03 第二轮）。单独一个文件是因为这一整套
**字段顺序写死、下游要按位置读**——改动风险高，埋在 5000 行里容易误碰。
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

import wi_core as exporter      # 界面里的老别名，见 app.py 那三行

from ui_const import HARD_SCOPES, _SCOPE_ORDER_MAP
from ui_fmt import _fmt_by_equip, _fmt_expected, _fmt_meta_expected, _value_str



# --------------------------------------------------- 交给 WorkBuddy 的交接包（JSON）
# 闭环（2026-09-28 定）：离线软件判不动（未通过率过线）→ 人点「导出 JSON」
# → 把这份文件交给 WorkBuddy → 它打开 PDF 原件复核 → 输出 result_*.json
# → 离线软件回读。现在只落地第一步（点了按钮才写本地文件），目录结构与
# 字段模板先定死，将来接通自动投递时不用再改结构、也不用再改 WorkBuddy 的读法。
_HANDOFF_ENABLE = False           # 打开后导出时**同时**往 inbox 放一份；当前只写本地


_HANDOFF_ROOT = "WI交接"           # 交接根目录（放在用户目录下）


_HANDOFF_INBOX = "inbox"          # 软件 → WorkBuddy：task_<WI>_<时间戳>.json


_HANDOFF_OUTBOX = "outbox"        # WorkBuddy → 软件：result_<WI>_<时间戳>.json（待回读）


_HANDOFF_DONE = "done"            # 软件已回读的结果移进来留档


_HANDOFF_SCHEMA = "WI-HANDOFF-1"  # 模板版本：字段结构一改就必须升号，让读的一方立刻发现


# 「完整条目」的安全阀：超过就截断。**重点标注项不截断** —— 那才是要复核的东西。
# 几万条的 JSON 没有哪一方读得动，闷头生成 10MB 文件反而害事；截断会在
# 「核对统计」里写明，绝不静默。
_HANDOFF_FULL_CAP = 20000




def _handoff_should(chk):
    """交接包里的「应为」：与界面同一套口径，但给全量（界面上一致时是折叠的）。"""
    if chk.get("meta_kind"):
        return _fmt_meta_expected(chk)
    if chk.get("status") == "no_match":
        return "？"
    if chk.get("status") == "no_scope":
        return _fmt_by_equip(chk)
    return _fmt_expected(chk)




def _handoff_scope_cn(chk):
    """判定档位的中文名：底层 scope 是英文短码，交给对方前翻成人话。"""
    sc = chk.get("scope") or ""
    return _SCOPE_ORDER_MAP.get(sc, sc) if sc else "未标注"




def _focus_reason(chk):
    """这条要不要单列进「重点标注项」；要则返回理由，不要返回 None。

    进重点的两类，缺一类都会漏：
      · 结论本身就有问题（不一致 / 无对应 / 判不了）—— 当然要复核；
      · **判了「一致」却落在宽松档**（候选不止一个、Excel 只填了一侧）——
        这种「一致」是撞上的而不是锁死的，假一致就藏在这里，比漏判更危险。
    反过来，硬判定的一致不进重点：把 1 万条里的 9 千条都标成重点，等于没标。
    """
    st = chk.get("status")
    if st == "mismatch":
        return "离线判为「不一致」，请照 PDF 原件确认真实取值"
    if st == "no_match":
        if chk.get("scope") == "name_unknown":
            return "离线不认得这个参数名（未收录），请在 PDF 里人工找"
        if chk.get("scope") == "no_step":
            return "PDF 该档参数表里没有这一步，请确认是不是拿错了档"
        if chk.get("scope") == "spec_block":
            return ("适用规格与 PDF 的规格档对不上，离线按规则**没有判**这一组参数 —— "
                    "请先确认 Excel 适用规格那一列有没有抄错，再谈参数值")
        return "离线在 PDF 里找不到对应 —— 可能 PDF 真没有，也可能只是没读透，两者必须分清"
    if st == "no_scope":
        return "Excel 该行没填设备号，期望值收不窄，离线判不了"
    if st == "ok":
        sc = chk.get("scope")
        if sc and sc not in HARD_SCOPES:
            return ("判了「一致」但口径没锁到最严（%s）：是候选之一碰上的，不是唯一命中"
                    % _handoff_scope_cn(chk))
    return None




def _handoff_item(chk, i):
    """交接包里的一条：字段名与顺序写死，与数据无关（有数据就填，没数据填空串）。"""
    return {
        "序号": i,
        "参数卡": chk.get("card", "") or "",
        "采集项": chk.get("item", "") or "",
        "类型": chk.get("type", "") or "",
        "现值": _value_str(chk),
        "应为": _handoff_should(chk),
        "单位": chk.get("unit", "") or "",
        "工序": chk.get("proc", "") or "",
        "判定": exporter.status_cn_of(chk),
        "判定档位": _handoff_scope_cn(chk),
        "依据": chk.get("basis", "") or "",
    }




def _handoff_focus_item(chk, i, seq_full, reason):
    """「重点标注项」里的一条：比普通条目多「标记」「原序号」「关注原因」三栏。

    字段顺序同样写死，且刻意把「标记」放在**第一栏** —— 对方扫 JSON 时
    第一眼就知道这条是要优先看的，不用先读「关注原因」再回头判断。
    """
    return {
        "标记": "重点",
        "序号": i,
        "原序号": seq_full,
        "参数卡": chk.get("card", "") or "",
        "采集项": chk.get("item", "") or "",
        "类型": chk.get("type", "") or "",
        "现值": _value_str(chk),
        "应为": _handoff_should(chk),
        "单位": chk.get("unit", "") or "",
        "工序": chk.get("proc", "") or "",
        "判定": exporter.status_cn_of(chk),
        "判定档位": _handoff_scope_cn(chk),
        "依据": chk.get("basis", "") or "",
        "关注原因": reason,
    }




def _handoff_task_block(reason):
    """任务说明：固定 6 个键、固定顺序，只有「本次触发原因」随本次结果变。

    为什么连文案都要写死：对方（WorkBuddy）每次拿到的形状完全一样，
    不用先猜结构再找重点，解析最快 —— 这就是「固定模板」的全部意义。
    """
    return {
        "这是什么": "一份 WI 作业指导书 PDF 与工程参数表 Excel 的离线比对结果",
        "为什么要交给你": "离线软件只能读 PDF 的文字层，换版式、扫描页、合并单元格都会读漏；"
                          "本次离线结果不足以直接采信，需要你打开 PDF 原件复核",
        "本次触发原因": reason,
        "你要做什么": [
            "① 打开「路径信息」里的 PDF 与 Excel，用你自己的方式读 PDF（不受离线解析器限制）",
            "② 逐条复核「重点标注项」：每条给出 PDF 里真实写的应为值、你的判定"
            "（一致 / 不一致 / 无法判断）、依据（页码或表格位置）",
            "③ 「完整条目」只作概览浏览：未列入重点的条目**不代表已确认正确**，"
            "发现明显问题同样要指出",
        ],
        "不要做什么": [
            "不要直接修改 Excel 或 PDF —— 你只出复核结论",
            "不要把「离线判为一致」当成依据：宽松档位的一致是碰上的，不是锁死的",
            "不要把「无对应」当成 PDF 里一定没有 —— 它同样可能只是离线没读到",
        ],
        "输出要求": "复核结果写成 result_<WI>_<时间戳>.json 放入 output 目录：条目字段沿用本文件"
                    "「重点标注项」的字段名，另补「你的判定」「你的依据」「复核来源」三项；"
                    "离线软件会自动回读，并把它标为 AI 复核结果（不覆盖本地判定）",
    }




def _safe_name(s, fallback):
    """文件名里不能带的字符一律换掉 —— WI 编号里出现 / 会让写出直接失败。"""
    s = re.sub(r'[\\/:*?"<>|]+', "_", str(s or "").strip())
    return s or fallback
