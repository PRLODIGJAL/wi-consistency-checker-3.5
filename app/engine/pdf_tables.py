# -*- coding: utf-8 -*-
"""pdf_tables.py —— PDF 的「表目录」：第几页、第几张表、这张表叫什么

为什么要这个模块（用户 2026-10-30 的原话）：
  「这一次就遇到了个整一大片都是没有这个关键词的……那这个东西在这个情况当中，
    它所体现的是表，那个表的表头或者表的名称，而不是表内部的某一个某一行字段」
  「是不是应该在这个中间再加一步，就是先选中是 P、D、F 当中的哪一个表，
    然后再提取这个表中的某哪一个字段的信息」
  「可以在比对的时候加一个它是第几页的表几，那这样子既方便人工去核对检验，
    也方便那个机器去提高准确率」

一个必须先说清的坑（否则做出来就是错的）：
  pdfplumber 在 WI-0043 第 6 页排出 5 张表，编号和文档自己印的完全错位 ——

    文档印的                           pdfplumber 的
    表1 NEO Pleat Head 开口尺寸设定  →  第 2 张（它的表1 只是页眉那一行）
    表2 NEO Fold Head 闭口尺寸设定   →  第 4 张
    表3 球囊折叠工艺参数表           →  第 5 张
    表4 球囊尺寸与转换工具对应表     →  第 3 张

  所以「第几页表几」的「表几」**必须用文档印的那个号**。报 pdfplumber 的序号，
  人工拿着去翻，翻到的还是另一张表 —— 等于没解决「行号对不上」那个老问题。

  文档号从表格**上方最近的**「表N，标题」那行抓（按 x 区间归属，不按行序），
  抓不到就老实写「第 N 张表」，绝不按位置猜一个号出来。

三条边界（与 xl_locate.py 一致）：
  · 只读；不参与任何判定，不加分、不筛候选、不改结论；
  · 认不出表号 → 如实说「第 N 张表」，不编；
  · 结果按 (路径, mtime) 缓存，一次运行只算一遍。
"""
import os
import re
import sys as _sys

# ── 目录四层自举（2026-11-03 目录改造）：`_open()` 会懒导入 wi_core，
#    四层不挂上时那条路会静默退回裸 pdfplumber（丢缓存、变慢但不报错）。
_APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _APP_DIR not in _sys.path:
    _sys.path.insert(0, _APP_DIR)
try:
    import paths as _paths
    _paths.install()
except Exception:       # pragma: no cover
    pass

_CACHE = {}

# 文档印的表号：「表1」「表2-a」「附表3」「Table 4」
_NUM = r"(?:[0-9]{1,2}|[一二三四五六七八九十]{1,3})"
_TAB_HEAD = re.compile(
    r"^(?:附表|表|Table)\s*(%s)\s*((?:[-－—–]\s*[A-Za-z0-9]{1,2})?)\s*"
    r"[:：,，、.．]?\s*(.*)$" % _NUM, re.I)

# 页眉/页脚那种「一张表只有一行」的条带，不是数据表
_STRIP_H = 22.0

_CJK = re.compile(r"[\u4e00-\u9fff]")


# ------------------------------------------------------------------ 小工具
def _open(path):
    """优先借引擎的 PDF 缓存（同一份文件不重复解析）。"""
    try:
        import wi_core
        return wi_core.open_pdf(path)
    except Exception:
        import pdfplumber
        return pdfplumber.open(path)


def _cell(v):
    return re.sub(r"\s+", " ", str(v or "")).strip()


def _join(tokens):
    """拼一行文字：两边都是中文就直接粘，否则补一个空格。

    pdfplumber 按空白切词，「表1，NEO Pleat Head 开口尺寸设定」会切成
    ['表1，NEO','Pleat','Head','开口尺寸设定'] —— 补空格才拼得回去；
    而纯中文行可能被切成多块，补空格会变成「球 囊 直 径」。
    """
    s = ""
    for t in tokens:
        if not s:
            s = t
            continue
        if _CJK.search(s[-1:]) and _CJK.search(t[:1]):
            s += t
        else:
            s += " " + t
    return s


def _lines(page, tol=3.5):
    """把一页的词按 y 聚成行 → [{text, top, bottom, x0, x1}]（自带坐标）。"""
    try:
        ws = page.extract_words(use_text_flow=False, keep_blank_chars=False) or []
    except Exception:
        return []
    ws = [w for w in ws if str(w.get("text") or "").strip()]
    if not ws:
        return []
    ws.sort(key=lambda w: (round(float(w["top"]), 1), float(w["x0"])))
    out = []
    for w in ws:
        top, bot = float(w["top"]), float(w["bottom"])
        if out and abs(top - out[-1]["_top"]) <= tol:
            out[-1]["ws"].append(w)
            out[-1]["_top"] = min(out[-1]["_top"], top)
            out[-1]["_bot"] = max(out[-1]["_bot"], bot)
        else:
            out.append({"_top": top, "_bot": bot, "ws": [w]})
    for ln in out:
        ln["ws"].sort(key=lambda w: float(w["x0"]))
        ln["top"], ln["bottom"] = ln["_top"], ln["_bot"]
        ln["x0"] = min(float(w["x0"]) for w in ln["ws"])
        ln["x1"] = max(float(w["x1"]) for w in ln["ws"])
        ln["text"] = _join([str(w["text"]) for w in ln["ws"]])
    return out


def _in_box(ln, box, pad=1.5):
    return (ln["x0"] >= box[0] - pad and ln["x1"] <= box[2] + pad
            and ln["top"] >= box[1] - pad and ln["bottom"] <= box[3] + pad)


def _segments(ln, gap=20.0):
    """一行里并排摆了好几个「表N」标题时，按词切成若干段（各自带 x 范围）。

    WI-0039 第 6 页是 4 行 × 3 列共 12 张表，每一**行**表的上方只有一行标题，
    一行里塞着三个：「表2-a, 外管拉伸参数表（DWG-0000013-01） 表2-e, …
    表3-c, …」。整行只取第一个匹配的话，横着三张表会全叫「表2-a」——
    人工照这个号去翻，翻到的是最左边那张，等于又回到「对不上」。

    切法：哪个词自己以「表N」开头，就从它这里起一段。要求「表」后必须跟着数字，
    所以「表面处理参数」这类正文不会被误切。

    第二道断句：横向间距超过 `gap` 也算断。同一行里左右两栏的表标题之间隔着几十点，
    不断开的话右栏那张表的第一格会被吞进左栏表名里 —— WI-0043 p6 的「表3，
    球囊折叠工艺参数表」就吞成了「…工艺参数表 φ5.0mm×200mm T-2712-26 SPC2000010-19」
    （那串是右栏表4 某一行的内容，只是恰好排在同一高度）。
    """
    segs = []
    for w in ln["ws"]:
        t = str(w.get("text") or "")
        new = False
        if _TAB_HEAD.match(t.strip()):
            new = True
        elif segs and (float(w["x0"]) - float(segs[-1][-1]["x1"])) > gap:
            new = True
        if new or not segs:
            segs.append([w])
        else:
            segs[-1].append(w)
    out = []
    for ws in segs:
        text = _join([str(w.get("text")) for w in ws]).strip()
        m = _TAB_HEAD.match(text)
        if not m:
            continue
        out.append({"text": text, "no": m.group(1),
                    "suffix": (m.group(2) or "").replace(" ", ""),
                    "name": (m.group(3) or "").strip(),
                    "x0": min(float(w["x0"]) for w in ws),
                    "x1": max(float(w["x1"]) for w in ws),
                    "top": ln["top"], "bottom": ln["bottom"]})
    return out


def _doc_titles(page, boxes):
    """这一页印着的「表N，标题」→ [{no, suffix, name, text, top, bottom, x0, x1}]。

    落在**某张表的框里**的一律不算（那是表格里的「表1」交叉引用，
    比如 WI-0043 p6 表3 第 8 行「Pleat Head Open Stop Micrometer Setting | 表1」）。
    """
    out = []
    for ln in _lines(page):
        if not _TAB_HEAD.match(ln["text"].strip()):
            continue              # 整行都不以「表N」开头，各段也不用试
        if any(_in_box(ln, b, pad=0.0) for b in boxes):
            continue              # 表内的「表N」引用，不是标题
        for ti in _segments(ln):
            if not ti["name"] and not ti["suffix"]:
                continue          # 光一个「表1」孤零零的，不当标题（无从归属）
            out.append(ti)
    return out


def _title_for(box, titles, max_gap=70.0):
    """框正上方、x 区间有重叠、离得最近的那个标题（不按行序，按坐标）。

    为什么按 x 归属：第 6 页左边「表1/表2/表3」竖着排、右边「表4」另起一栏，
    按行序会给右边的表派左栏的标题。

    max_gap 的作用：标题离表太远就不算它的。没有这道门，一张自己没印表号的表
    会把上方很远的「表1」抢过来 —— 报出去的表号比不报更糟。
    """
    bx0, btop, bx1, _bbot = box
    best = None
    for ti in titles:
        if ti["bottom"] > btop + 2.0:
            continue                              # 必须在表上方
        if btop - ti["bottom"] > max_gap:
            continue                              # 隔太远：不是这张表的标题
        if min(ti["x1"], bx1) - max(ti["x0"], bx0) <= 0:
            continue                              # x 不重叠 = 不是这张表的标题
        if best is None or ti["top"] > best["top"]:
            best = ti
    return best


def _subtitle(lines, ti, max_gap=16.0):
    """标题下面那行英文/括号补充 —— Excel 常直接抄它当采集项名。

    「表1，NEO Pleat Head 开口尺寸设定」下一行是
    「（Pleat Head Open Stop Micrometer Setting）」，而 Excel 里的采集项正是
    「Pleat Head Open Stop Micrometer Setting」：抓出来，人才对得上。
    """
    for ln in lines:
        if not (ti["bottom"] - 1.0 <= ln["top"] <= ti["bottom"] + max_gap):
            continue
        s = ln["text"].strip()
        if len(s) > 2 and s[0] in "(（" and s[-1] in ")）":
            return s.strip("()（）").strip()
    return ""


def _headers(rows):
    """表头字段（前两行按列拼，「›」表示上下两级）。只给人看，不参与判定。

    整段正文不算字段名：有些表的头一格是合并格，塞着几百字的作业说明
    （WI-0039 p4 实测），照抄出来「字段」就成了一篇散文，反而看不出这张表是什么。
    判据取「长度 > 60 字」—— 最长的正常字段名
    （`Fold Head Close Stop Micrometer Setting`，38 字）离它还远。
    """
    if not rows:
        return []
    n = max(len(r or ()) for r in rows[:2])
    out = []
    for c in range(n):
        parts = []
        for r in rows[:2]:
            v = _cell(r[c]) if r and c < len(r) else ""
            if len(v) > 60:
                continue
            if v and v not in parts:
                parts.append(v)
        if parts:
            out.append(" › ".join(parts))
    return out[:8]


def _spec_in_table(rows):
    """表**内部**写在「参数名称」那格括号里的适用规格 → 去重后的字符串列表。

    为什么补这一路（2026-11-09 实测）：
      `page_tier.spec_text` 抓的是**表格框外面**那一行「（For NEO RX-90cm，
      Balloon Length 15-80mm）」，只有 WI-0042 那种「12 张附表」的版式才有。
      真实语料 15 份 PDF / 157 张表实测 **0 张**——不是它坏了，是那种页一张都没有。

      而这些文档的规格其实**写在表里面**：每张小节的头一行是

          序号 | 参 数名 称 (产品线A/Aperi SC/… Ø1.5-2.25mm,产品线B Ø1.5-3.0mm) | 参数值或范围

      括号里就是「这张小节适用哪些规格」。同一张表里堆了 4~7 个这样的小节，
      所以一张表可能同时适用好几种规格 —— 全收下来，按出现顺序去重。

    只给人看（选表、核对时知道这页说的是哪一档），**不参与判定**。
    """
    out = []
    for r in rows or ():
        for c in (r or ()):
            s = _cell(c)
            if len(s) > 200 or "(" not in s and "（" not in s:
                continue
            # 真实 PDF 里这几个字常被拆开排：「参 数名 称 (…)」「序 号」。
            # 判关键词前先把空白去掉，否则一半的表头认不出来（实测 6/157 → 修后见下）。
            key = re.sub(r"\s+", "", s)
            if not re.search(r"(参数名称|参数名|Parameter|规格)", key, re.I):
                continue
            for m in re.finditer(r"[(（]([^()（）]{3,160})[)）]", s):
                g = m.group(1).strip()
                # 得像个规格：带直径符号 / mm / RX / OTW 才算，别把「见表2」收进来
                if re.search(r"[ØøΦφФ]|mm|RX|OTW|cm", g, re.I):
                    if g not in out:
                        out.append(g)
    return out[:6]


# ------------------------------------------------------------------ 主入口
def catalog(pdf_path, recs=None):
    """整份 PDF 的表目录（只读、带缓存）。

    recs: 引擎抽出的竖表记录（`extract_pdf_vertical` 的结果）。给了就能算出
          「这张表里引擎抽到几条」，用来发现「整张表一条都没抽出来」的漏抽。
          不给则 nrec=None。

    每条：
      page    页码（1 起）        tno    pdfplumber 序号（0 起，与引擎记录同源）
      no      文档印的表号        suffix 「-a」这种分表后缀
      name    表名（文档印的）    en     表名下的英文/括号副标题
      label   「第6页 表2」/「第6页 第4张表」
      headers 表头字段            rows   表里行数
      nrec    引擎从这张表抽了几条  first_col 第一列前几个值
      kind    "table" / "header"（页眉条）
    """
    try:
        key = (pdf_path, os.path.getmtime(pdf_path))
    except OSError:
        return []
    ck = ("cat", key)
    if ck in _CACHE:
        got = _CACHE[ck]
    else:
        got = _build(pdf_path)
        _attach_spec(pdf_path, got)
        _CACHE[ck] = got
    if recs is None:
        return got
    cnt = {}
    for r in recs:
        k = (r.get("page"), r.get("table"))
        cnt[k] = cnt.get(k, 0) + 1
    out = []
    for it in got:
        d = dict(it)
        d["nrec"] = cnt.get((it["page"], it["tno"]), 0)
        out.append(d)
    return out


def _attach_spec(pdf_path, cat):
    """给表目录补一条「表外规格」（`spec`）—— 只给人看，不参与判定。

    为什么要有（2026-11-02 用户原话）：
      「我在选这个表的时候……它根本没有表以外的数据信息。那我根本不知道这个
        规格参数是哪一页，哪张表，我还得去翻原 PDF 里面的。」
    WI-0042 那 12 张附表的表名一样（页眉只写「设备参数附表1…12」），
    真正分得开的是标题那行括号里的 `For NEO RX-90cm，Balloon Length 15-80mm`。
    解析交给 `page_tier.spec_text`（它已经在解析同一行，不另写一套）。

    2026-11-09 实测补正：这条「页级」规格在真实语料 157 张表里命中 0 张 ——
    那种附表页一张都没有，所以**不是它坏了**。真正有规格的是表内那行
    `参数名称(Ø1.5-2.25mm/…)`，由 `_spec_in_table` 在 `_build` 里先填好。
    两者关系：页级（附表身份）优先，没有就保留表内那份，显示层不区分来源。
    """
    try:
        import page_tier
        specs = page_tier.spec_text(pdf_path)
    except Exception:
        return
    for it in cat or ():
        s = specs.get(it.get("page"))
        if s:
            it["spec"] = s


def _build(pdf_path):
    out = []
    try:
        pdf = _open(pdf_path)
    except Exception:
        return out
    try:
        for pno, page in enumerate(pdf.pages, 1):
            try:
                tabs = page.find_tables() or []
            except Exception:
                continue
            if not tabs:
                continue
            boxes = [tuple(t.bbox) for t in tabs]
            try:
                lines = _lines(page)
            except Exception:
                lines = []
            titles = _doc_titles(page, boxes)
            n_data = 0
            for tno, t in enumerate(tabs):
                box = tuple(t.bbox)
                try:
                    rows = t.extract() or []
                except Exception:
                    rows = []
                is_hdr = (len(rows) <= 1) or ((box[3] - box[1]) < _STRIP_H)
                ti = None if is_hdr else _title_for(box, titles)
                if is_hdr:
                    n_data += 0
                    label = "第%d页 页眉条" % pno
                else:
                    n_data += 1
                    if ti and ti["no"]:
                        label = "第%d页 表%s%s" % (pno, ti["no"], ti["suffix"])
                    else:
                        label = "第%d页 第%d张表" % (pno, n_data)
                first = []
                for r in rows[1:]:
                    v = _cell(r[0]) if r else ""
                    if v and v not in first:
                        first.append(v)
                    if len(first) >= 3:
                        break
                # 非空格子数 / 有内容的行数：用来分辨「整页空白/纯图页」与「真参数表」。
                # WI-0023/2106 第 2 页是 35 行 × 10 列的表壳，除页眉外全空 ——
                # 只看行数会当成参数表。为什么两个数都要：
                # 光看「非空格子占比」，窄竖表会被误伤（WI-0024 第 6 页的参数只在
                # 前两列，10 列里 8 列空，占比天然很低，但它是真参数表）。
                # 「有内容的行数」不受列宽影响：空白页只有页眉那 1 行。
                nonempty = 0
                filled = 0
                for r in rows:
                    hit = False
                    for v in (r or ()):
                        if _cell(v):
                            nonempty += 1
                            hit = True
                    if hit:
                        filled += 1
                out.append({
                    "page": pno, "tno": tno,
                    "no": (ti or {}).get("no", ""),
                    "suffix": (ti or {}).get("suffix", ""),
                    "name": (ti or {}).get("name", ""),
                    "en": _subtitle(lines, ti) if ti else "",
                    "title": (ti or {}).get("text", ""),
                    "label": label,
                    "spec": "；".join(_spec_in_table(rows)) if not is_hdr else "",
                    "headers": _headers(rows),
                    "rows": len(rows),
                    "cols": max((len(r) for r in rows), default=0),
                    "nonempty": nonempty,
                    "filled_rows": filled,
                    "first_col": first,
                    "kind": "header" if is_hdr else "table",
                    "bbox": box,
                })
    finally:
        try:
            pdf.close()
        except Exception:
            pass
    return out


# ------------------------------------------------------------------ 表格网格
def npages(pdf_path):
    """这份 PDF 一共几页。

    给「② 页」那张页码表用（2026-11-04 第十二轮）：原来那里是「表目录」，
    现在改成「所有页码」，就得知道总页数 —— `catalog()` 只返回**有表的**页，
    封皮、说明页不在里面，拿它的最大页号当总数会少算。
    走 `_open()`（借 wi_core 的缓存），不会为了一次页数再解析一遍文件。
    """
    try:
        return int(len(_open(pdf_path).pages))
    except Exception:          # noqa: BLE001 —— 坏文件退回 0，由调用方兜
        return 0


def grid(pdf_path, page, tno):
    """某一张表的**原始格子**—— rows[r][c] 是字符串（空串表示空/合并格）。

    为什么要这个入口（2026-11-02）：
      竖表通道只认「一列参数名 + 右边一列值」那种版式。WI-0039 第 7 页的
      「表4-1 DWG-0000013白管切定长」是**横着**排的（列是「拉伸前/拉伸后/拉伸段」），
      引擎从这张表里一条都没抽出来（实测 nrec=0）。可人眼一看就知道哪一格对哪一格。
      所以需要一条「按坐标直接读格子」的通路 —— 由人指定位置，工具照着读，
      不再要求它自己看懂这张表是什么版式。

    只读；结果按 (路径, mtime, 页, 表号) 缓存（一次比对里同一张表会被问很多遍）。
    """
    try:
        key = (pdf_path, os.path.getmtime(pdf_path))
    except OSError:
        return []
    ck = ("grid", key, int(page), int(tno))
    if ck in _CACHE:
        return _CACHE[ck]
    rows = []
    try:
        pdf = _open(pdf_path)
    except Exception:
        _CACHE[ck] = []
        return []
    try:
        try:
            p = pdf.pages[int(page) - 1]
            tabs = p.find_tables() or []
            if 0 <= int(tno) < len(tabs):
                rows = [[_cell(v) for v in (r or [])]
                        for r in (tabs[int(tno)].extract() or [])]
        except Exception:
            rows = []
    finally:
        try:
            pdf.close()
        except Exception:
            pass
    _CACHE[ck] = rows
    return rows


def cell_boxes(pdf_path, page, tno):
    """这张表**每个格子的 bbox** —— `rows[r][c]` 是 `(x0, top, x1, bottom)`，缺给 None。

    为什么有它（2026-11-04 第十一轮，用户原话）：
      「它提取错了，我也没办法去勾选对，而且甚至我也没法发现……有没有可能做到，
        就是基于原本的 PDF 表格，我给它划分可点击的区域。然后我鼠标拖到相对应的
        位置，它就进行高亮标记……那我就可以直接点击原生的 PDF 图像。」
      也就是：不再让人对着**重排过的**提取结果点，而是对着 PDF 原图上点。

    行列顺序与 `grid()` **完全一致**（同一个 `find_tables()[tno]`、同一套 Table
    对象），所以「在原图上点中一格」能直接换算回 (r, c)，喂给现有的钉住逻辑，
    判定引擎一行都不用改。

    可靠性（2026-11-04 实测）：6 份 WI、57 张表，`rows` 数 vs `extract()` 行数、
    每行 `cells` 数 vs 列数 —— **100% 对齐**。但界面那边仍然要求「点完在 ③ 里
    同步高亮」，让人自己看一眼：真遇到对不齐的表，那一眼就能发现，
    比闷着学错强（用户原话：「那整个机器就会导致越学习越笨，越学习越乱」）。

    只读；按 (路径, mtime, 页, 表号) 缓存 —— 拖一下窗口就会问几十遍。
    """
    try:
        key = (pdf_path, os.path.getmtime(pdf_path))
    except OSError:
        return []
    ck = ("cells", key, int(page), int(tno))
    if ck in _CACHE:
        return _CACHE[ck]
    out = []
    try:
        pdf = _open(pdf_path)
    except Exception:
        _CACHE[ck] = []
        return []
    try:
        try:
            p = pdf.pages[int(page) - 1]
            tabs = p.find_tables() or []
            if 0 <= int(tno) < len(tabs):
                for row in (tabs[int(tno)].rows or []):
                    cs = []
                    # 老版 pdfplumber 的 TableRow 没有 .cells —— 给 None 占位，
                    # 调用方按「这一格没坐标」处理（不画热区，别猜）。
                    for c in (getattr(row, "cells", None) or []):
                        cs.append(None if not c else
                                  (float(c[0]), float(c[1]),
                                   float(c[2]), float(c[3])))
                    out.append(cs)
        except Exception:
            out = []
    finally:
        try:
            pdf.close()
        except Exception:
            pass
    _CACHE[ck] = out
    return out


# ------------------------------------------------------------------ 模板指纹
# 「表号 + 表名 + 前几个字段名」= 这张表的模板指纹。
#
# 为什么不用参数名做键（2026-11-02 用户点破的方向）：
#   参数名的写法每个 WI 都在变（「Fold 头关闭时针规设定值 不过」vs
#   `Fold Head Close Stop Micrometer Setting`），而**表的形状**在同一族模板里
#   几乎不动 —— 表叫什么、有哪几列、列头怎么写，都稳定。
#   所以「记住位置」比「记住名字」耐用：换个文档、换个页码，只要还是这张表，
#   位置照样对得上。
_FP_NOISE = re.compile(
    r"^(?:页码|版本|文件编号|修订|生效日期|第\s*\d+\s*页|rev\.?\s*\d+|wi-?\s*\d+)",
    re.I)


def _fp_fields(headers):
    """表头字段里的「稳定部分」—— 页码 / 版本 / 文件编号这类每页都变的去掉。"""
    out = []
    for h in headers or ():
        s = _cell(h)
        s = re.sub(r"\s+", "", s)
        if not s or _FP_NOISE.match(s):
            continue
        # 「工序名称 › 管件切割」取后半截：前半截是合并格的上层标头，各页都一样
        if "›" in s:
            s = s.split("›")[-1].strip()
        if s and s not in out:
            out.append(s)
    return out


def fingerprint(item):
    """一张表的模板指纹（人也能看懂的字符串，便于排查）。

    表号 + 表名 + 前 4 个字段名。表号/表名缺一个就少一个，用剩下的拼 ——
    拼出来太短（比如只有一个「第1张表」）就当没指纹，返回空串，
    由调用方决定「认不出就不套用」，绝不硬猜。
    """
    if not item or item.get("kind") == "header":
        return ""
    no = (_cell(item.get("no") or "") + _cell(item.get("suffix") or "")).strip()
    name = re.sub(r"\s+", "", _cell(item.get("name") or ""))
    flds = _fp_fields(item.get("headers"))[:4]
    parts = [x for x in (no, name) if x]
    if not parts:
        # 没有表号也没有表名：只能靠字段名认，而且字段得有 3 个以上才敢用，
        # 否则「序号|参数名称|值」这种万能头会把所有表都认成同一张。
        if len(flds) < 3:
            return ""
        return "…|" + "|".join(flds)
    return "|".join(parts + flds)


def match_tables(cat, fp):
    """指纹对得上的**所有**表（精确档；精确档一张都没有时退回宽松档，最多给一张）。

    为什么需要「所有」而不是「第一张」（2026-11-02 实案 WI-0030）：
      第 4 页有 5 张表，其中「第 2 张表」（锻打，`1 Die 0.0210″`）和
      「第 4 张表」（焊接，`1 Temperature 175-181/181-185`）都是
      `序号 | 参数名称 | 参数值或范围` 三列 —— 表号和表名都没抽到，指纹只能
      靠字段名，于是两张表的指纹**一模一样**：

          …|序号|参数名称|参数值或范围      ← 第2张表
          …|序号|参数名称|参数值或范围      ← 第4张表

      `match_table()` 按目录顺序返回第一张 → 用户钉在「第4张表第2行第2列」
      （温度），解析时却跑到「第2张表第2行第2列」去取，读到 `0.0210″` →
      界面显示 0.021。**同一份文件里串表**，比认不出更糟：它造出来的是
      一条看起来很像真的假差异。

      所以这里把「有几张」如实交出去，由 `anchors` 用「钉的时候在哪一页哪张表」
      和「那一格的原文」去认，认不出就 stale —— 不猜。
    """
    fp = (fp or "").strip()
    if not fp:
        return []
    tabs = [c for c in (cat or ()) if c.get("kind") != "header"]
    exact = [c for c in tabs if fingerprint(c) == fp]
    if exact:
        return exact
    # 宽松档：**并列最高分的表全给出去**，不挑一张。
    # 2026-11-11 实案（WI-0028 的焊机温度被读成 0.021）：老锚点的指纹是
    # `…|序号|参数名称|参数值或范围`（当年没抽到表号表名），这一页 6 张表全是
    # 这个表头，宽松档算出来清一色 0.6 分。以前只返回**第一张**（表1，Die 锻打），
    # 于是「按你钉的位置读到 0.0210″」。多给几张不会更糟 —— `anchors.pick_tab`
    # 后面还有「那一格的原文」「钉的时候在第几页第几张表」两道判据可用来消歧，
    # 实在认不出就 stale 不套用（宁可不套，也不能凭空造一条假差异）。
    return _loose_tables(cat, fp, tabs)


def _loose_tables(cat, fp, tabs):
    """宽松档：表名/表号对得上 + 字段重合度最高的一批（可多张）。"""
    head = [x for x in fp.split("|") if x]
    if not head:
        return []
    ident = [x for x in head if not x.startswith("…")]
    best_s, best = 0.0, []
    for c in tabs:
        cfp = fingerprint(c)
        if not cfp:
            continue
        ch = [x for x in cfp.split("|") if x]
        inter = len(set(head) & set(ch))
        union = len(set(head) | set(ch)) or 1
        s = inter / union
        # 表名（带表号的那一段）必须对得上，否则一律不算
        if ident and not (set(ident) & set(ch)):
            continue
        if s > best_s:
            best, best_s = [c], s
        elif s == best_s and s > 0:
            best.append(c)
    return best if best_s >= 0.5 else []


def match_table(cat, fp):
    """在这份 PDF 的表目录里，找指纹对得上的那张表（精确 → 宽松）。

    返回目录项 dict；找不到返回 None。

    ⚠️ 同指纹可能不止一张表（见 `match_tables` 的说明）。这个函数只返回**第一张**，
    只适合「不关心是哪一张、只要知道有没有」的场合；**按锚点定位必须用
    `match_tables`**，否则会同页串表。

    宽松那一档的必要性：同一族模板升版时常在表名后加个括号、字段名改个标点
    （「适用范围（mm)」半角括号变全角）。要求逐字相同等于一次小改版就全废。
    但宽松也只放宽到「表名相同 + 字段重合 ≥ 一半」，不认识就返回 None ——
    套错表比不套更糟（会把别人的值安到这一行上，制造假一致）。
    """
    fp = (fp or "").strip()
    if not fp:
        return None
    tabs = [c for c in (cat or ()) if c.get("kind") != "header"]
    for c in tabs:
        if fingerprint(c) == fp:
            return c
    # 宽松：拆成「ident（表号/表名）」与「字段」两段比
    head = [x for x in fp.split("|") if x]
    if not head:
        return None
    ident = [x for x in head if not x.startswith("…")]
    best, best_s = None, 0.0
    for c in tabs:
        cfp = fingerprint(c)
        if not cfp:
            continue
        ch = [x for x in cfp.split("|") if x]
        inter = len(set(head) & set(ch))
        union = len(set(head) | set(ch)) or 1
        s = inter / union
        # 表名（带表号的那一段）必须对得上，否则一律不算
        if ident and not (set(ident) & set(ch)):
            continue
        if s > best_s:
            best, best_s = c, s
    return best if best_s >= 0.5 else None


# ------------------------------------------------------------------ 查询
def tables_only(cat):
    """去掉页眉条，只留真正的数据表。"""
    return [c for c in (cat or []) if c.get("kind") == "table"]


def by_key(cat):
    """{(页, tno): 目录项} —— 用来把引擎的一条记录对回它所在的表。"""
    return {(c["page"], c["tno"]): c for c in (cat or [])}


def label_of(cat, page, tno):
    """「第6页 表2 NEO Fold Head 闭口尺寸设定」；查不到就空串。"""
    it = by_key(cat).get((page, tno))
    if not it:
        return ""
    if it.get("name"):
        return "%s %s" % (it["label"], it["name"])
    return it["label"]


def name_where(recs=None):
    """{PDF 参数名: (页, tno)} —— 它主要出现在哪张表里。

    一个名字横跨多张表时取出现次数最多的那张（并列取页码/表号靠前的）。
    只用于「按表筛选」和显示，不参与判定。
    """
    cnt = {}
    for r in recs or ():
        p = str(r.get("param") or "").strip()
        if not p:
            continue
        k = (r.get("page"), r.get("table"))
        cnt.setdefault(p, {})
        cnt[p][k] = cnt[p].get(k, 0) + 1
    return {p: max(ks, key=lambda x: (ks[x], -(x[0] or 0), -(x[1] or 0)))
            for p, ks in cnt.items()}


def name_index(cat, recs=None):
    """{PDF 参数名: 它出现最多的那张表的标签}（「第6页 表2 NEO Fold Head 闭口尺寸设定」）。

    只用于显示「这个名字出自哪张表」，不参与判定。
    """
    return {p: label_of(cat, k[0], k[1]) for p, k in name_where(recs).items()}


def text(pdf_path, recs=None):
    """给人看的纯文本目录（日志/说明里用）。"""
    cat = catalog(pdf_path, recs)
    if not cat:
        return "（这份 PDF 没有识别出表格）"
    L = []
    for c in cat:
        if c["kind"] == "header":
            continue
        L.append("%-16s %-46s %2d行  %s"
                 % (c["label"], (c.get("name") or "—")[:46], c["rows"],
                    ("引擎抽到 %d 条" % c["nrec"]) if c.get("nrec") is not None else ""))
        if c.get("en"):
            L.append("%-16s   （%s）" % ("", c["en"][:64]))
        if c.get("headers"):
            L.append("%-16s   字段：%s" % ("", " | ".join(c["headers"])))
    return "\n".join(L)
