# -*- coding: utf-8 -*-
"""历史快照 —— 把一次核对的结果压成签名，存盘，用于「这次和上次比」。

为什么单独成一个模块（2026-10-30 拆分第一步）：
  这几十行与判定毫无关系，只吃 checks 列表、吐 json，是整份引擎里最干净的
  一块。先把它搬出来，等于先验证「拆包 + 对拍」这条流水线本身是通的，
  再去动真正的判定。

对外四个函数：
  build_signature(checks)            判定结果 → 签名
  save_snapshot(sig, base, wi, ...)  存一份（同一 WI 只留最近 keep 份）
  latest_snapshot(base, wi)          取上一次
  diff_snapshot(old, new)            两次对比（修好多少 / 新坏多少）

⚠️ 这一块不参与任何判定，搬动前后 15208 行结论必须逐行一致。
"""
import json
import os
from datetime import datetime

SNAP_DIRNAME = "历史快照"


def ensure_snap_dir(base):
    d = os.path.join(base, SNAP_DIRNAME)
    os.makedirs(d, exist_ok=True)
    return d


def build_signature(checks):
    """把判定结果压成 {键: {状态, 档位, 应为值}}，用于跨版本对比。"""
    out, seen = {}, {}
    for v in checks:
        base = (str(v.get("card") or ""), str(v.get("item") or ""))
        n = seen.get(base, 0)
        seen[base] = n + 1
        key = "%s|%s|%d" % (base[0], base[1], n)
        out[key] = {
            "s": str(v.get("status") or ""),
            "sc": str(v.get("scope") or ""),
            "e": _exp_str(v.get("expected")),
        }
    return out


def _exp_str(e):
    if e is None:
        return ""
    if isinstance(e, (list, tuple, set)):
        return ";".join(sorted(str(x) for x in e))
    return str(e)


def save_snapshot(sig, base_dir, wi, keep=3, note=""):
    """存一份快照；同一 WI 只保留最近 keep 份，避免目录无限膨胀。"""
    d = ensure_snap_dir(base_dir)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    p = os.path.join(d, "快照_%s_%s.json" % (wi or "WI", ts))
    with open(p, "w", encoding="utf-8") as f:
        json.dump({"wi": wi, "time": ts, "note": note, "sig": sig},
                  f, ensure_ascii=False)
    prefix = "快照_%s_" % (wi or "WI")
    files = sorted(x for x in os.listdir(d)
                   if x.startswith(prefix) and x.endswith(".json"))
    for old in files[:-keep]:
        try:
            os.remove(os.path.join(d, old))
        except OSError:
            pass
    return p


def latest_snapshot(base_dir, wi):
    """取该 WI 最近一份快照（本次保存之前调用，拿到的就是上一次）。"""
    d = os.path.join(base_dir, SNAP_DIRNAME)
    if not os.path.isdir(d):
        return None
    prefix = "快照_%s_" % (wi or "WI")
    files = sorted(x for x in os.listdir(d)
                   if x.startswith(prefix) and x.endswith(".json"))
    if not files:
        return None
    try:
        with open(os.path.join(d, files[-1]), encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return None


def diff_snapshot(old_sig, new_sig, sample=8):
    """对比两份签名，返回 {'fixed','broke','added','gone','changed','samples'}。

    fixed = 上次不一致、这次一致（改对了）
    broke = 上次一致、这次不一致（新引入）
    """
    res = {"fixed": 0, "broke": 0, "added": 0, "gone": 0, "changed": 0,
           "samples": {"fixed": [], "broke": [], "changed": []}}
    old_keys, new_keys = set(old_sig), set(new_sig)
    for k in old_keys & new_keys:
        a, b = old_sig[k], new_sig[k]
        if a["s"] != b["s"]:
            if a["s"] == "mismatch" and b["s"] == "ok":
                res["fixed"] += 1
                if len(res["samples"]["fixed"]) < sample:
                    res["samples"]["fixed"].append((k, a.get("e", ""), b.get("e", "")))
            elif a["s"] == "ok" and b["s"] == "mismatch":
                res["broke"] += 1
                if len(res["samples"]["broke"]) < sample:
                    res["samples"]["broke"].append((k, a.get("e", ""), b.get("e", "")))
            else:
                res["changed"] += 1
                if len(res["samples"]["changed"]) < sample:
                    res["samples"]["changed"].append(
                        (k, "%s→%s" % (a["s"], b["s"]), ""))
        elif a.get("e") != b.get("e"):
            res["changed"] += 1
            if len(res["samples"]["changed"]) < sample:
                res["samples"]["changed"].append((k, a.get("e", ""), b.get("e", "")))
    res["added"] = len(new_keys - old_keys)
    res["gone"] = len(old_keys - new_keys)
    res["total"] = res["fixed"] + res["broke"] + res["changed"] + res["added"] + res["gone"]
    return res
