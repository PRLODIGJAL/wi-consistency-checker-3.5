# -*- coding: utf-8 -*-
"""把 PDF 的一页（或页里的一块）**画成图片** —— 给「指定位置」窗看原样用的。

为什么需要它（2026-11-04 第七轮，用户原话）：
  「我双击展开之后，它不是会展示出表格吗？但是那个表格也是提取出来的嘛毕竟，
    它不一定完全的准确。所以我觉得……在右边附属上这个 PDF 的展示图片，就是这个
    表格在 PDF 中原来是什么样子的。」

也就是说：右边那张格子表是 pdfplumber **认**出来的，认错了（列错位、合并格拆开）
只有对着 PDF 原图才能一眼看出来。这里只负责「把原图取出来」，不参与任何判定。

实现选择
  · 用 pypdfium2 直接渲染（pdfplumber 的硬依赖，双平台装得上，不用再加包）；
  · 不走 pdfplumber 的 page.to_image()：那条路每次都要重新 open 一份文档、
    而且它按 cropbox 反推缩放，遇到 /Rotate 页会错位（见 _visual_box 的注释）；
  · 只渲染需要的**一块**：按目标像素宽度反推 scale，再裁切 —— 整页渲染再缩到
    500px 宽，表里的字就糊成一团，等于白给。

只读、无副作用、结果按 (路径, mtime, 页, 区域, 宽度) 缓存。
"""
import os
import sys as _sys

_APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _APP_DIR not in _sys.path:
    _sys.path.insert(0, _APP_DIR)
try:
    import paths as _paths
    _paths.install()
except Exception:       # pragma: no cover
    pass

# 渲染结果的缓存。键里带 mtime：同一路径的文件被换掉之后不能拿旧图。
_CACHE = {}
_CACHE_MAX = 24          # 一个窗最多来回看十几张表，24 张够用，再多就丢最旧的


def available():
    """能不能渲染（pypdfium2 在不在）。缺了就由界面降级成一句说明。"""
    try:
        import pypdfium2      # noqa: F401
        return True
    except Exception:         # noqa: BLE001
        return False


def _to_display(box, pw, ph, rot):
    """pdfplumber 的 bbox → **眼睛看到的那一页**上的 bbox。

    两个坐标系不是一回事：pdfplumber 给的是 PDF 用户空间（未旋转、原点左下、
    bbox 的 top/bottom 是从上往下数的距离），而 pypdfium2 默认渲染时**已经把
    /Rotate 转正了** —— 人打开 PDF 看到什么，它就画什么。

    这套语料里只有封面页带 /Rotate 180（没有表），但书签/封面 180 度在真实
    文件里很常见，不处理的话一旦表出现在那种页上，图就会裁到对角去。

    `pw/ph` 必须是**未旋转**的页面尺寸（pdfplumber 的 page.width/height）。
    返回值落在转正后的坐标系里：rot 为 90/270 时那个坐标系是 (ph, pw)。

    2026-11-04 第十一轮：从 `_visual_box(page, bbox)` 里抽出来，改成纯函数 ——
    因为「在原图上点格子」那条路拿不到 pdfplumber 的 page 对象，只有
    (宽, 高, 旋转角) 三个数。公式只此一份，两个入口共用，免得哪天改了一边。
    """
    x0, top, x1, bottom = [float(v) for v in box]
    r = int(rot or 0) % 360
    if r == 180:
        return (pw - x1, ph - bottom, pw - x0, ph - top)
    if r == 90:
        return (ph - bottom, x0, ph - top, x1)
    if r == 270:
        return (top, pw - x1, bottom, pw - x0)
    return (x0, top, x1, bottom)


def _visual_box(page, bbox):
    """（旧入口，保留）从 pdfplumber 的 page 对象取显示坐标。"""
    return _to_display(bbox, float(page.width), float(page.height),
                       int(getattr(page, "rotation", 0) or 0) % 360)


def to_px(box, meta):
    """pdfplumber 的 bbox → **这张图上的像素坐标** (x0,y0,x1,y1)。

    `meta` 是 `render_ex()` 返回的第二个值。三件事依次做：

      ① 转成显示坐标（`_to_display`，管 /Rotate）；
      ② 减掉**页面盒原点** `org`（管 MediaBox 不从 0,0 起的 PDF，见 `_page_org`）；
      ③ 乘倍率、减裁切原点。

    第 ② 步是 2026-11-11 补的：以前只有 ①③ ，遇到负数起点的 MediaBox，
    所有框会整体偏出去 ~12pt —— 一页上正好一行，看着特别像「行号错了一位」，
    实际是坐标系原点没对齐。

    拿不到 meta 就返回 None：调用方据此**不画热区**。宁可点不动，
    也不要画一个偏的框 —— 那比不画更让人误判。
    """
    if not meta:
        return None
    try:
        pw, ph = meta["page_wh"]
        d = _to_display(box, pw, ph, meta.get("rot") or 0)
        s = float(meta["scale"])
        ox, oy = meta["crop"]
        gx, gy = meta.get("org") or (0.0, 0.0)
        return ((d[0] - gx) * s - ox, (d[1] - gy) * s - oy,
                (d[2] - gx) * s - ox, (d[3] - gy) * s - oy)
    except (KeyError, TypeError, ValueError):
        return None


def _page_org(pg, rot):
    """页面盒原点 → pdfplumber 那套坐标要**加/减**多少 pt。只处理 rot=0。

    为什么需要（2026-11-11，用户截图里「选中的框整体偏了一大截」）：

    pdfplumber 的格子坐标是相对**页面盒左上角**算的；pdfium 渲染时却把整页
    画进 `(0, 0, 宽, 高)`。盒子的起点只要不是 (0,0)，两边就差一个固定量。
    实测 WI-0028 第 4 页：MediaBox = (-11.99, 11.96, 829.69, 607.16)，
    差 11.96pt —— 那张表行高正好 10.7pt，所以看起来像「整块往上一行」。

    正常 PDF（绝大多数）盒起点就是 (0,0)，这里返回 (0,0)，行为一个字不变。
    旋转页（90/180/270）的映射要另算一套，这里先返回 (0,0)：
    宁可保持旧行为，也不要在没验证过的情况下乱动。
    """
    if int(rot or 0) % 360:
        return (0.0, 0.0)
    try:
        bb = pg.get_bbox()
        x0, y0 = float(bb[0]), float(bb[1])
    except Exception:          # noqa: BLE001 —— 拿不到就当它是正常的
        return (0.0, 0.0)
    if abs(x0) < 0.01 and abs(y0) < 0.01:
        return (0.0, 0.0)
    # pdfplumber 的上边距 = -y0（盒子高 = 上-下，减一下正好是 -y0）
    return (x0, -y0)


def page_size(pdf_path, page_no):
    """这一页**显示时**的宽高（pt）。取不到就给 None。"""
    try:
        import pypdfium2 as pdfium
        doc = pdfium.PdfDocument(pdf_path)
        try:
            return doc[max(0, int(page_no) - 1)].get_size()
        finally:
            doc.close()
    except Exception:         # noqa: BLE001
        return None


def visual_box(pdf_path, page_no, bbox):
    """pdfplumber 的 bbox → 显示坐标系（人看到的那一页）的 (x0,y0,x1,y1)。

    返回 (box, (页宽, 页高))；取不到给 (None, None)。
    「整页」视图要在原图上用红框圈出这张表，用的就是它 —— 圈错位置比不圈更糟，
    所以调用方拿到 None 就直接不画框。
    """
    try:
        import pdfplumber
        with pdfplumber.open(str(pdf_path)) as pdf:
            p = pdf.pages[max(0, int(page_no) - 1)]
            return _visual_box(p, bbox), (float(p.width), float(p.height))
    except Exception:         # noqa: BLE001
        return None, None


def render_ex(pdf_path, page_no, bbox=None, target_w=520, pad_pt=12.0,
              max_scale=4.0, whole=False):
    """渲染一页（或页里 bbox 那一块），返回 `(PIL.Image, meta)`；失败给 `(None, None)`。

    参数
      page_no   页码，1 起（与 pdfplumber / 界面上的「第 N 页」一致）
      bbox      pdfplumber 的 (x0, top, x1, bottom)；None 或 whole=True → 整页
      target_w  想要多宽（px）。scale 由它反推，所以窗口拉宽后图更清楚，
                而不是把同一张图拉伸变糊
      pad_pt    裁切时四周留多少 pt 余量（不留的话表框贴着图边，看不出上下文）

    **meta 是 2026-11-04 第十一轮加的**（用户要「直接在 PDF 原图上点格子」）：
      图出来时是**裁过、缩过**的，外面只拿到一张 PIL.Image 时，算不出「PDF 上
      那一点落在图像的第几个像素」。把那几个数一并吐出来，`to_px()` 拿它换算。
        {"scale": 每 pt 多少 px, "crop": 裁切原点(px), "page_wh": 未旋转页宽高(pt),
         "rot": 旋转角, "disp": 转正后的页宽高(pt)}
    """
    key = (os.path.abspath(str(pdf_path)), page_no, tuple(bbox or ()),
           int(target_w), float(pad_pt), bool(whole))
    try:
        key = key + (os.path.getmtime(pdf_path),)
    except Exception:           # noqa: BLE001
        pass
    hit = _CACHE.get(key)
    if hit is not None:
        return hit

    try:
        import pypdfium2 as pdfium
        import pdfplumber
    except Exception:           # noqa: BLE001
        return None, None

    img = None
    meta = None
    try:
        doc = pdfium.PdfDocument(str(pdf_path))
        try:
            idx = max(0, int(page_no) - 1)
            if idx >= len(doc):
                return None, None
            pg = doc[idx]
            dw, dh = pg.get_size()              # 已转正，就是人看到的那一页
            rot = int(getattr(pg, "get_rotation", lambda: 0)() or 0) % 360
            # pdfium 给的是**转正后**的尺寸；90/270 时未旋转尺寸要换回来，
            # 否则 _to_display() 里的 pw/ph 会左右颠倒，热区整体错位。
            pw, ph = (dh, dw) if rot in (90, 270) else (dw, dh)
            # 页面盒原点：MediaBox 不从 (0,0) 起时，图上的内容和 pdfplumber 的
            # 格子坐标差一个固定量（WI-0028 差 11.96pt，看着像「差一行」）。
            org = _page_org(pg, rot)
            box = (0.0, 0.0, float(dw), float(dh))
            if bbox and not whole:
                try:
                    with pdfplumber.open(str(pdf_path)) as pdf:
                        p0 = pdf.pages[idx]
                        # 一律以 pdfplumber 的 page 为准（_visual_box 一直是这么
                        # 用的，且宽高是**未旋转**的，正是 _to_display 要的口径）
                        pw, ph = float(p0.width), float(p0.height)
                        rot = int(getattr(p0, "rotation", 0) or 0) % 360
                        box = _to_display(bbox, pw, ph, rot)
                except Exception:               # noqa: BLE001
                    box = (0.0, 0.0, float(dw), float(dh))
                x0, y0, x1, y1 = box
                # 越界 / 反了就退回整页 —— 宁可给一张大的，也别裁出一片空白
                if not (x1 > x0 and y1 > y0) or x0 < -2 or y0 < -2 \
                        or x1 > dw + 4 or y1 > dh + 4:
                    box = (0.0, 0.0, float(dw), float(dh))
                else:
                    box = (max(0.0, x0 - pad_pt), max(0.0, y0 - pad_pt),
                           min(float(dw), x1 + pad_pt), min(float(dh), y1 + pad_pt))
            bw, bh = box[2] - box[0], box[3] - box[1]
            scale = max(0.6, min(float(max_scale), float(target_w) / max(1.0, bw)))
            full = pg.render(scale=scale)
            img = full.to_pil()
            crop = (0, 0)
            if box != (0.0, 0.0, float(dw), float(dh)):
                # 先减页面盒原点再乘倍率 —— 裁切窗口必须落在**图上真实的位置**，
                # 否则放大视图裁到的也是偏的那一块（和热区同一个毛病）。
                bx = (int(round((box[0] - org[0]) * scale)),
                      int(round((box[1] - org[1]) * scale)),
                      int(round((box[2] - org[0]) * scale)),
                      int(round((box[3] - org[1]) * scale)))
                bx = (max(0, bx[0]), max(0, bx[1]),
                      min(img.size[0], bx[2]), min(img.size[1], bx[3]))
                if bx[2] > bx[0] and bx[3] > bx[1]:
                    img = img.crop(bx)
                    # 裁掉的这些像素要在换算时减回去 —— 少减一下，整片热区
                    # 就整体偏出去一个 pad（默认 12pt，肉眼很明显）
                    crop = (bx[0], bx[1])
            img = img.convert("RGB")
            meta = {"scale": float(scale), "crop": crop, "org": org,
                    "page_wh": (float(pw), float(ph)), "rot": int(rot),
                    "disp": (float(dw), float(dh))}
        finally:
            doc.close()
    except Exception:           # noqa: BLE001 —— 坏 PDF / 缺库都不该把窗带崩
        return None, None

    if img is not None:
        if len(_CACHE) >= _CACHE_MAX:
            try:
                _CACHE.pop(next(iter(_CACHE)))
            except Exception:   # noqa: BLE001
                _CACHE.clear()
        _CACHE[key] = (img, meta)
    return img, meta


def render(pdf_path, page_no, bbox=None, target_w=520, pad_pt=12.0,
           max_scale=4.0, whole=False):
    """同 `render_ex()`，只返回图片 —— 老调用方用这个，签名一个字都没变。

    要坐标换算（在原图上叠可点区域）就用 `render_ex()`。
    """
    return render_ex(pdf_path, page_no, bbox=bbox, target_w=target_w,
                     pad_pt=pad_pt, max_scale=max_scale, whole=whole)[0]


def clear_cache():
    _CACHE.clear()
