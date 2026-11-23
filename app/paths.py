# -*- coding: utf-8 -*-
r"""paths.py —— **全项目唯一的路径来源**。

2026-11-03 目录改造时新建（对齐 `项目i5统一UI美学/统一架构_目录规范.md`）。
以前 `dirname(__file__)` 在 app.py / anchors.py / name_lex.py 里各拼一遍，
「词库写哪儿」「快照写哪儿」每处一套说法，搬一次目录就散一地。
**凡是「某个东西在哪儿」的问题，只能问这里。**

四层结构

    app/ui/      界面层    只画界面、收事件、把活派下去
    app/judge/   判定层    ⚠️ 本版**暂未启用**，见下面「为什么 judge 是空的」
    app/engine/  引擎层    读 PDF / Excel、锚点与表级规则、参数名词库
    app/shell/   启动壳    环境自检、装依赖、启动过程的那几行字

    data/        运行期产物（词库、锚点、上次选择、历史快照），可删
    dev/         开发资产（回归件、代码结构总表、操作详解），**交付可整个删掉**
    .venv/       本地环境，不随包分发

为什么 judge 是空的
------------------
`wi_core.py`（14664 行）目前是「引擎 + 判定 + 导出 + 批量」一体，而且
`pdf_tables`（引擎）与 `wi_core` 之间是**双向懒导入**。硬拆成
`engine/` + `judge/rules/` 会让「依赖只能向下」立刻破功，也会把判定链
动一遍 —— 那是独立一轮工作，不是搬目录该干的事。

同 `PDF转换拼接 5.0` 的处理方式：judge/ 先建目录留占位，业务规则暂并进
engine/wi_core.py。拆分路线已写在 `dev/代码结构与规则总表.md` 第八节
（common → memory → engine → rules → ui，八步，每步过四道闸）。
拆到第 5 步时，`rules/` 的内容搬进这里的 judge/，界面一行都不用改
（界面只认 engine 对外暴露的 31 个名字）。

⚠️ 模块名保持**扁平**（`import wi_core`，不写 `from app.engine import ...`）。
四层是**目录**分层，不是 Python 包分层 —— 引擎内部互相 import 有二十多处，
改成包路径收益为零、风险极高。`install()` 把四层一起塞进 `sys.path`。
"""

import os
import sys

# …\交付根\app
HERE = os.path.dirname(os.path.abspath(__file__))
# …\交付根
ROOT = os.path.dirname(HERE)

APP = HERE
UI = os.path.join(APP, "ui")
JUDGE = os.path.join(APP, "judge")
ENGINE = os.path.join(APP, "engine")
SHELL = os.path.join(APP, "shell")

# 四层一次性加进搜索路径：顺序即优先级（目前无重名模块）
LAYERS = (UI, JUDGE, ENGINE, SHELL)

DATA_DIR = os.path.join(ROOT, "data")
DEV_DIR = os.path.join(ROOT, "dev")
VENV = os.path.join(ROOT, ".venv")

# 旧名保留：`DOC_DIR` 在 2.0 原包里指「文档\」，现在统一并进 dev/
DOC_DIR = DEV_DIR


def install():
    """把四层加进 `sys.path`（幂等），返回交付根目录。

    入口只要两行自举：

        sys.path.insert(0, os.path.join(_ROOT, "app"))
        import paths; paths.install()
    """
    for p in reversed(LAYERS):          # 反向插入，最终 UI 在最前
        if os.path.isdir(p) and p not in sys.path:
            sys.path.insert(0, p)
    return ROOT


def data_dir():
    """数据目录，不存在就建。"""
    os.makedirs(DATA_DIR, exist_ok=True)
    return DATA_DIR


def data_path(name):
    """数据目录下的某个文件（父目录不存在会自动建）。"""
    data_dir()
    return os.path.join(DATA_DIR, name)


def describe():
    """一行说明这份代码在哪、四层齐不齐 —— 给自检和报障用。"""
    miss = [os.path.basename(p) for p in LAYERS if not os.path.isdir(p)]
    return "根目录 %s｜四层 %s" % (ROOT, "齐全" if not miss else "缺 " + "、".join(miss))
