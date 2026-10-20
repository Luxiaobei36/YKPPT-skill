# -*- coding: utf-8 -*-
"""
verify_deck.py — 生成后静态校验（不需要 PowerPoint）

检查项
------
1. 越界      任何形状超出画布或侵入页眉带
2. 字号      出现 style.md 字号阶梯以外的值
3. 字体      出现微软雅黑 / Arial 以外的字体（含 east-asian 字形）
4. 标题      横幅标题超过 26 字（会溢出）
5. 结论条    文案超过 52 字
6. 溢出      文本框内容高度超过框高（按 PIL 实测字宽估算）
7. 页码      正文页缺页码或格式不是 -N-

用法
----
    python verify_deck.py out.pptx
    python verify_deck.py out.pptx --strict     # 有問題即非零退出
"""

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Emu

import ykrender as R


def _inch100(v):
    return None if v is None else Emu(int(v)).inches * 100.0


def _walk(shapes, sx=1.0, sy=1.0):
    """递归遍历，同时累乘组合的缩放系数。

    为什么要算缩放：组合内子形状的 left/width 是**子坐标系**的值，
    和屏幕上看到的尺寸差一个系数。封面「汇报人」文本框在母版里只有 2.31in 宽，
    但所在组合横向缩放 2.87 → 实际画出来 6.78in 宽。不乘这个系数，
    会把本来放得下的文字误报成溢出。
    """
    from pptx.oxml.ns import qn
    for sh in shapes:
        yield sh, sx, sy
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            gsx, gsy = sx, sy
            xfrm = sh._element.find(".//" + qn("a:xfrm"))
            if xfrm is not None:
                ext = xfrm.find(qn("a:ext"))
                ch = xfrm.find(qn("a:chExt"))
                if ext is not None and ch is not None:
                    cx, cy = int(ch.get("cx")), int(ch.get("cy"))
                    if cx:
                        gsx = sx * int(ext.get("cx")) / cx
                    if cy:
                        gsy = sy * int(ext.get("cy")) / cy
            yield from _walk(sh.shapes, gsx, gsy)


def _line_height_units(text, size_pt, box_w, spacing):
    """估算文本块高度（1/100 inch）。

    单行文本的高度由字体自身行高决定（约 1.2em），与段落 line_spacing 无关；
    多行时才按 line_spacing 累加。否则 54pt「目录」这种单行会被按 1.5 倍误判溢出。
    """
    n = R.wrap_lines(text, size_pt, box_w)
    eff = spacing if n > 1 else 1.2
    return n * size_pt * eff / 72.0 * 100.0


def _fonts_of(shape):
    """取形状内所有 run 的字体信息：(latin, ea, size_pt, text)。"""
    out = []
    if not shape.has_text_frame:
        return out
    from pptx.oxml.ns import qn
    for p in shape.text_frame.paragraphs:
        for run in p.runs:
            if not run.text.strip():
                continue
            rPr = run._r.find(qn("a:rPr"))
            ea = None
            if rPr is not None:
                el = rPr.find(qn("a:ea"))
                if el is not None:
                    ea = el.get("typeface")
            out.append((run.font.name, ea,
                        run.font.size.pt if run.font.size else None,
                        run.text))
    return out


def check(pptx_path):
    prs = Presentation(str(pptx_path))
    problems = []
    for idx, slide in enumerate(prs.slides, 1):
        tag = f"第 {idx} 页"
        # --- 标题 / 结论条 ---
        title_sh = R._find(slide, "yk_title")
        if title_sh is not None:
            t = title_sh.text_frame.text.strip()
            if len(t) > 26:
                problems.append(f"{tag} 横幅标题 {len(t)} 字（上限 26）：{t}")
        concl_sh = R._find(slide, "yk_concl_text")
        if concl_sh is not None:
            c = concl_sh.text_frame.text.strip()
            if len(c) > 52:
                problems.append(f"{tag} 结论条 {len(c)} 字（上限 52）：{c[:20]}…")
            elif not c:
                problems.append(f"{tag} 结论条为空")

        is_body = title_sh is not None
        if is_body:
            n = None
            for ph in slide.placeholders:
                if ph.placeholder_format.type == 13:
                    n = ph.text_frame.text.strip()
            if not n:
                problems.append(f"{tag} 正文页缺少页码（应来自版式占位符）")
            elif not (n.startswith("-") and n.endswith("-")):
                problems.append(f"{tag} 页码格式应为 -N-，实为 {n!r}")

        # --- 逐个形状 ---
        for sh, sx, sy in _walk(slide.shapes):
            l, t = _inch100(sh.left), _inch100(sh.top)
            w, h = _inch100(sh.width), _inch100(sh.height)
            if None in (l, t, w, h):
                continue
            # 换算到屏幕坐标（组合内子形状要乘缩放系数）
            l, w = l * sx, w * sx
            t, h = t * sy, h * sy
            # 越界
            if l < -1 or t < -1 or l + w > R.CANVAS_W + 1 or t + h > R.CANVAS_H + 1:
                # 母版来的装饰形状允许出血（校徽打底、照片）
                if not sh.name.startswith(("校徽打底", "图书馆照片", "海天苑", "背景色块",
                                           "五边形", "矩形: 剪去对角", "矩形 3", "矩形 5")):
                    problems.append(
                        f"{tag} 形状 {sh.name!r} 越界："
                        f"({l:.0f},{t:.0f},{w:.0f},{h:.0f})")
            # 侵入页眉带
            if is_body and t > 0 and t < R.HEAD_BOTTOM - 12 and h < 500 \
                    and not sh.name.startswith(("yk_section", "yk_title", "矩形", "占位符")):
                problems.append(f"{tag} 形状 {sh.name!r} 侵入页眉带 y={t:.0f}")
            # 字体与字号
            for latin, ea, size, text in _fonts_of(sh):
                if size is not None and int(round(size)) not in R.SIZE_LADDER:
                    problems.append(
                        f"{tag} 字号 {size}pt 不在阶梯内（{sorted(R.SIZE_LADDER)}）：{text[:16]}")
                if latin and latin not in R.ALLOWED_FONTS:
                    problems.append(f"{tag} 西文字体 {latin!r} 不合规：{text[:16]}")
                if ea and ea not in R.ALLOWED_FONTS:
                    problems.append(f"{tag} 中文字体 {ea!r} 不合规：{text[:16]}")
            # 文本溢出
            if sh.has_text_frame and sh.text_frame.text.strip():
                paras = sh.text_frame.paragraphs
                sizes = [r.font.size.pt for p in paras for r in p.runs if r.font.size]
                if sizes:
                    sp = paras[0].line_spacing or 1.5
                    if not isinstance(sp, float):
                        sp = 1.5
                    need = _line_height_units(sh.text_frame.text, max(sizes), w, sp)
                    # 组合内的文字形状可能 h=0（几何由组合缩放决定），不据此判定
                    if h > 1 and need > h + 6 \
                            and not sh.name.startswith(("yk_title", "yk_section")):
                        problems.append(
                            f"{tag} 文字可能溢出 {sh.name!r}：需 {need:.0f} 有 {h:.0f}")
    return problems


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description="静态校验生成的 PPTX")
    ap.add_argument("pptx")
    ap.add_argument("--strict", action="store_true")
    args = ap.parse_args(argv)

    probs = check(args.pptx)
    if not probs:
        print("✓ 静态校验通过")
        return 0
    print(f"发现 {len(probs)} 处问题：")
    for p in probs:
        print("  ! " + p)
    return 1 if args.strict else 0


if __name__ == "__main__":
    sys.exit(main())
