这个目录是「判定层」，按 i5《统一架构_目录规范.md》第 2 节的定义：

    判定层（judge）＝ 业务规则：什么算对、怎么分类、要不要先问用户

现在它是空的，**不是漏建**，跟 PDF转换拼接 5.0 一样是刻意的：

  wi_core.py（约 1.47 万行）目前是「引擎 + 判定 + 导出 + 批量」一体，
  而且 pdf_tables（读表）与 wi_core 之间是双向懒导入。硬拆成
  engine/ + judge/ 会让「依赖只能向下（ui → judge → engine）」当场破功，
  也会把判定链整条动一遍 —— 那是独立一轮工作，不是搬目录该干的事。

拆分路线已写在 dev/代码结构与规则总表.md 第八节（common → memory →
engine → rules → compare → ui，八步，每步过四道闸）。拆到第 5 步时，
rules/ 的内容搬进这里，界面一行都不用改 —— 因为界面只认 engine 对外
暴露的 31 个名字。

在那之前：判定逻辑住在 app/engine/wi_core.py，结构自检住在
app/engine/wic/audit.py。
