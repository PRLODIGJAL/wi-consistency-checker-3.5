"""本地状态的落盘与读取：上次选择、代码戳、路径解析。

从 app.py 拆出来的（2026-11-03 第二轮）。这一组东西的共同点：
**它们碰磁盘**（读/写 json、看文件 mtime），而 `ui_fmt.py` 里的函数全是
纯计算。分开之后，想给格式化逻辑写单测时不必碰文件系统。
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
import sys
from datetime import datetime

from paths import data_path



# ---- 代码版本戳 -------------------------------------------------------------
# Python 只在**启动那一刻**读一次源码。窗口开着的时候改了 wi_core.py，
# 界面里跑的仍然是启动时的旧代码 —— 而界面上完全看不出来：同样的 bug 照旧复现，
# 人只会得出「改了这么多遍怎么还没修好」的结论（2026-09-21 实际发生过一次）。
# 对策两步：① 标题栏亮出启动时的引擎时间戳，一眼能看出窗口有多旧；
#          ② 每次点「开始比对」时核一遍磁盘，源码已更新就当场提示重启。
# 2026-11-03 起不再列文件名，改成**扫目录**（app/engine 与 app/ui 下的全部 .py）。
# 原来只写死 ("wi_core.py", "app.py")，两个毛病：
#   ① 目录改四层后这里算的是 app/ui/，wi_core.py 其实在 app/engine/，
#      所以引擎那一项永远读不到，等于没盯；
#   ② 界面拆成 13 个文件后，改 ui_table.py 不会刷新时间戳，窗口照旧说
#      「源码没更新」，人会以为改动无效而反复重跑。
# 扫目录以后新增文件自动纳入。排除 __pycache__，不含 setup_env.py（GUI 不 import 它）。
_CODE_DIRS = ("engine", "ui")




def _code_stamp():
    """(最新的源码 mtime, 对应的文件名)；全都读不到时返回 (0.0, "未知")。"""
    app = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # app/
    newest, who = 0.0, "未知"
    for sub in _CODE_DIRS:
        for dp, dn, fn in os.walk(os.path.join(app, sub)):
            dn[:] = [d for d in dn if d != "__pycache__"]
            for f in fn:
                if not f.endswith(".py"):
                    continue
                try:
                    m = os.path.getmtime(os.path.join(dp, f))
                except OSError:
                    continue
                if m > newest:
                    newest, who = m, f
    return newest, who





# 目录里最新的那份代码是什么时候改的（窗口启动时算一次，后面都读这个）
_CODE_STAMP, _CODE_STAMP_FILE = _code_stamp()

def _stamp_text(ts=None):
    ts = _CODE_STAMP if ts is None else ts
    return "未知" if not ts else datetime.fromtimestamp(ts).strftime("%m-%d %H:%M")




def code_is_stale():
    """磁盘上的源码是否比本窗口启动时更新 —— 返回提示语，没更新返回 ""。"""
    newest, who = _code_stamp()
    if newest <= _CODE_STAMP + 1:       # 容差 1 秒：同一次保存的多个文件
        return ""
    return ("源码已更新（%s %s），但当前窗口里跑的还是启动时的旧代码 —— "
            "请关闭本窗口、重新双击启动器后再看结果。" % (who, _stamp_text(newest)))




# 上次选择记忆：把上一次用过的 PDF / Excel / 工作表写在这，下次打开自动带回；
# 若该文件已被删除或改名，对应栏位留空，由用户重新选择。
# 2026-10-30：落盘位置统一交给 paths.py（程序文件只读，写出来的东西进 数据\）。
CONFIG_PATH = data_path("上次选择.json")




def load_last():
    """读取上次选择；文件缺失、损坏或不是 JSON 对象时一律返回空 dict。"""
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}




def _is_real_session():
    """当前这个进程，是不是「用户真的双击打开的那一个界面」。

    为什么要这一道：`数据\上次选择.json` 是用户真实的记忆（上次用的哪对文件、
    左栏拖到多宽、选的哪张表）。可是所有测试脚本都是 `import app` 之后自己
    `App(root)` 一个出来 —— 它们摆假路径、假宽度、还拖分栏线量像素，
    只要哪一处走到 `_save_last()`，假东西就顶掉了真记忆。
    （2026-11-02 一天里踩了三次：`split: 777` 让左栏一开就顶到最宽、
      假路径 `//192.168.0.1/…/w/…` 顶掉了用户那对真实的 WI-0042 + DWG-0000014。）

    靠「给每个测试脚本加环境变量」治不住 —— 工作区里有 20 个脚本会建 App，
    补一个漏一个（这次就漏了写 `from app import App` 的 s0 / s7）。
    所以改成认**进程身份**：正常启动是 `setup_env.py` → `python app/ui/app.py`，
    此时 `__main__` 就是 app.py 自己；任何被 import 进测试进程的 app 模块，
    `__main__` 都是那个测试脚本，一律不落盘。
    判断不出身份时**按「不是真实会话」处理** —— 宁可不记，也不能覆盖用户的记忆。
    """
    try:
        main = sys.modules.get("__main__")
        f = os.path.abspath(getattr(main, "__file__", "") or "")
    except Exception:
        return False
    return os.path.basename(f).lower() == "app.py"




def save_last(d):
    """写回上次选择；写失败不影响主流程（只丢记忆，不影响比对）。

    两道门都过才落盘：
      ① 环境变量 `WI_NO_SAVE_LAST` —— 测试脚本显式声明「我在测试」；
      ② `_is_real_session()` —— 进程本身得是用户双击起来的那个界面。
    ② 是兜底：即使某个测试脚本忘了设 ①，也不会动用户的文件。
    正常双击启动两道门都放行，行为一字不变。
    """
    if os.environ.get("WI_NO_SAVE_LAST"):
        return
    if not _is_real_session():
        return
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass




def _pick_dir(*paths):
    """取第一个存在的目录作为文件对话框的起始目录，都没有则返回 None。"""
    for p in paths:
        if p and os.path.isdir(p):
            return p
    return None




def _resolve(last, key, fallback):
    """决定启动时预填哪个路径。

    - 有这一项记忆且文件还在 → 用它；
    - 有记忆但文件已被删除/改名 → 留空，让用户重新选（不偷偷回落到默认）；
    - 完全没有这一项记忆（首次运行）→ 用内置默认路径（存在才用）。
    """
    if key in last:
        p = str(last.get(key) or "").strip()
        return p if os.path.exists(p) else ""
    return fallback if os.path.exists(fallback) else ""
