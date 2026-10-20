# -*- coding: utf-8 -*-
"""
ykrender.py — 西工大组会/答辩 PPT 渲染引擎

设计原则
--------
1. **不画页眉**。五边形、校徽、斜切深蓝横幅、页脚渐变条都在 slideLayout 里，
   由母版提供；本引擎只克隆母版的 5 张「页型样板页」再往正文区填内容。
2. **坐标全部来自 references/style.md 与 archetypes.md**，单位 1/100 inch。
3. **一页一论点**：每页必须有 section / title / lead / conclusion。

依赖：python-pptx >= 0.6.21。Pillow 可选（用于精确测字宽，缺了退回启发式估算）。
"""

import copy
import re
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

# ============================================================ 度量
UNIT = 9144                      # 1/100 inch 对应的 EMU 数
CANVAS_W, CANVAS_H = 1333, 750   # 画布（1/100 inch）

BODY_L, BODY_R = 35, 1298        # 正文区左右边界
HEAD_BOTTOM = 103                # 页眉横线
SEP_Y = 180                      # 定性句下的分隔线
CONTENT_TOP = 190                # 正文内容上界
CONCL_Y = 642                    # 结论条顶
CONTENT_BOTTOM_CONCL = 615       # 有结论条时的内容下界
CONTENT_BOTTOM = 642             # 无结论条时的内容下界

# ============================================================ 色板（见 style.md §2）
NAVY = "142C76"
DEEP = "0E1E52"
BRAND = "014DA1"
SKY = "5BB0E6"
PALE = "EAF3FA"
PALE2 = "DDEEFA"
WHITE = "FFFFFF"
RED = "C00000"
RED2 = "FF0000"
TIP_YELLOW = "FFFF00"
DASH_GRAY = "7F7F7F"
DASH_AMBER = "FFC000"
NODE_RED = "F6645F"
NODE_TEAL = "3CB6B9"
NODE_BLUE = "123987"
TEXT_SUB = "4E4E4E"
TABLE_HEAD = "0071C1"

# ============================================================ 字体与字号（style.md §3/§4）
LATIN = "Arial"
CJK = "微软雅黑"
S_COVER, S_TOC_HEAD, S_TITLE, S_LEAD = 48, 54, 28, 20
S_BODY, S_DENSE, S_SMALL, S_TINY = 16, 14, 12, 10
# 目录条目 24pt：原 4 份文件全部如此，是这套模板唯一的 24pt 用法，
# 所以收进阶梯而不是改字号（改小了目录页会明显发空）。
S_TOC_ENTRY = 24
SIZE_LADDER = {S_COVER, S_TOC_HEAD, S_TITLE, S_LEAD, S_BODY, S_DENSE,
               S_SMALL, S_TINY, S_TOC_ENTRY}

# 正文字体：中文微软雅黑，西文 Arial（style.md §3 唯一允许的两款）
ALLOWED_FONTS = {CJK, LATIN}

_rgb = RGBColor.from_string


# ============================================================ 文字度量
_FONT_CACHE = {}


def _pil_font(size_pt):
    """按字号缓存 PIL 字体对象；不可用时返回 None。"""
    if size_pt in _FONT_CACHE:
        return _FONT_CACHE[size_pt]
    font = None
    try:
        from PIL import ImageFont
        for cand in (r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\msyh.ttf"):
            if Path(cand).exists():
                # 放大 4 倍量测再除，减少取整误差
                font = ImageFont.truetype(cand, int(size_pt * 4))
                break
    except Exception:
        font = None
    _FONT_CACHE[size_pt] = font
    return font


def text_width(text, size_pt):
    """返回文本单行宽度，单位 1/100 inch。"""
    if not text:
        return 0.0
    font = _pil_font(size_pt)
    if font is not None:
        try:
            return font.getlength(text) / 4.0 / 72.0 * 100.0
        except Exception:
            pass
    # 启发式：CJK 全角 = 1em，ASCII ≈ 0.52em
    w = 0.0
    for ch in text:
        w += 1.0 if ord(ch) > 0x2E80 else 0.52
    return w * size_pt / 72.0 * 100.0


MARK_RE = re.compile(r"(〔.+?〕|\*\*.+?\*\*|==.+?==)")


def plain(text):
    """去掉强调标记，得到纯文本。"""
    def _strip(m):
        s = m.group(0)
        if s.startswith("〔"):
            return s[1:-1]
        return s[2:-2]
    return MARK_RE.sub(_strip, text or "")


def wrap_lines(text, size_pt, box_w):
    """估算折行后的行数。遵循 CJK 逐字断行 + 西文按词断行。"""
    if not text:
        return 1
    total = 0
    for para in text.split("\n"):
        para = plain(para)
        if not para:
            total += 1
            continue
        cur = 0.0
        lines = 1
        for token in re.findall(r"[A-Za-z0-9_\-\./]+|\s|.", para):
            tw = text_width(token, size_pt)
            if cur + tw > box_w and cur > 0:
                lines += 1
                cur = tw if token.strip() else 0.0
            else:
                cur += tw
        total += lines
    return total


def text_height(text, size_pt, box_w, spacing=1.5):
    """返回文本块所需高度，单位 1/100 inch。"""
    n = wrap_lines(text, size_pt, box_w)
    return n * size_pt * spacing / 72.0 * 100.0


# ============================================================ 基础写入
def _style_run(run, size, bold, color, latin=LATIN, cjk=CJK):
    """设一个 run 的字体。中文必须显式写 <a:ea>，否则回退到主题字体。

    DrawingML 要求 rPr 子元素有序：... a:latin, a:ea, a:cs ...
    所以 a:ea 必须紧跟 a:latin，追加到末尾会被 PowerPoint 判为非法。

    `charset="-122"` 不能省：它是 GB2312 的中文字符集标识。缺了它，
    PowerPoint 在中文 Windows 上会挑到另一款更宽的 CJK 字形，
    同一段 48pt 文字就会多折一行（实测结尾页「…批评指正！」的叹号被挤到第二行）。
    """
    f = run.font
    f.size = Pt(size)
    f.bold = bold
    if color:
        f.color.rgb = _rgb(color)
    f.name = latin                      # 会写入 a:latin（并按 schema 插到正确位置）
    rPr = run._r.get_or_add_rPr()
    # lang 决定断行规则：标为 zh-CN 时 PowerPoint 才启用中文标点悬挂/压缩，
    # 否则「…批评指正！」的！会被当成普通字符单独折到下一行。
    rPr.set("lang", "zh-CN")
    rPr.set("altLang", "en-US")
    old = rPr.find(qn("a:ea"))
    if old is not None:
        rPr.remove(old)
    ea = etree.Element(qn("a:ea"))
    ea.set("typeface", cjk)
    ea.set("charset", "-122")
    latin_el = rPr.find(qn("a:latin"))
    if latin_el is not None:
        latin_el.addnext(ea)
    else:
        rPr.append(ea)


def _fill_text(tf, text, size, bold, color, spacing, align, space_after=0):
    """把 text 写进 tf：支持 **加粗红** / ==高亮红== 与 \\n 分段。"""
    tf.word_wrap = True
    first = True
    for para_text in str(text).split("\n"):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if spacing is not None:
            p.line_spacing = spacing
        if space_after:
            p.space_after = Pt(space_after)
        for piece in MARK_RE.split(para_text):
            if not piece:
                continue
            if piece.startswith("〔") and piece.endswith("〕"):
                body, b, c = piece[1:-1], True, RED
            elif piece.startswith("**") and piece.endswith("**"):
                body, b, c = piece[2:-2], True, RED
            elif piece.startswith("==") and piece.endswith("=="):
                body, b, c = piece[2:-2], True, RED2
            else:
                body, b, c = piece, bold, color
            run = p.add_run()
            run.text = body
            _style_run(run, size, b, c)


def add_text(slide, l, t, w, h, text, size=S_BODY, bold=False, color="000000",
             spacing=1.5, align=PP_ALIGN.JUSTIFY, anchor=MSO_ANCHOR.TOP,
             name=None, margins=(0, 0, 0, 0)):
    box = slide.shapes.add_textbox(Emu(int(l * UNIT)), Emu(int(t * UNIT)),
                                   Emu(int(w * UNIT)), Emu(int(h * UNIT)))
    tf = box.text_frame
    tf.vertical_anchor = anchor
    tf.margin_left, tf.margin_right = Emu(int(margins[0] * UNIT)), Emu(int(margins[1] * UNIT))
    tf.margin_top, tf.margin_bottom = Emu(int(margins[2] * UNIT)), Emu(int(margins[3] * UNIT))
    _fill_text(tf, text, size, bold, color, spacing, align)
    if name:
        box.name = name
    return box


def add_round_rect(slide, l, t, w, h, fill=None, line=None, line_w=0.75,
                   radius=0.055, name=None, dash=None):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                Emu(int(l * UNIT)), Emu(int(t * UNIT)),
                                Emu(int(w * UNIT)), Emu(int(h * UNIT)))
    try:
        sh.adjustments[0] = radius
    except Exception:
        pass
    if fill:
        sh.fill.solid()
        sh.fill.fore_color.rgb = _rgb(fill)
    else:
        sh.fill.background()
    if line:
        sh.line.color.rgb = _rgb(line)
        sh.line.width = Pt(line_w)
        if dash:
            sh.line.dash_style = dash
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    if name:
        sh.name = name
    return sh


def add_rect(slide, l, t, w, h, fill, name=None, line=None, line_w=0.75):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                Emu(int(l * UNIT)), Emu(int(t * UNIT)),
                                Emu(int(w * UNIT)), Emu(int(h * UNIT)))
    sh.fill.solid()
    sh.fill.fore_color.rgb = _rgb(fill)
    if line:
        sh.line.color.rgb = _rgb(line)
        sh.line.width = Pt(line_w)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    if name:
        sh.name = name
    return sh


def add_line(slide, l, t, w, h, color=SKY, weight=1.0, dash=None, name=None):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,
                                    Emu(int(l * UNIT)), Emu(int(t * UNIT)),
                                    Emu(int((l + w) * UNIT)), Emu(int((t + h) * UNIT)))
    ln.line.color.rgb = _rgb(color)
    ln.line.width = Pt(weight)
    if dash:
        ln.line.dash_style = dash
    if name:
        ln.name = name
    return ln


def add_picture_fit(slide, path, l, t, w, h, name=None):
    """等比缩放进 (l,t,w,h) 并居中，返回实际占位。"""
    path = str(path)
    if not Path(path).exists():
        raise FileNotFoundError(f"图片不存在：{path}")
    pic = slide.shapes.add_picture(path, Emu(0), Emu(0))
    nw, nh = pic.width, pic.height
    k = min(w * UNIT / nw, h * UNIT / nh)
    pw, ph = int(nw * k), int(nh * k)
    pic.width, pic.height = Emu(pw), Emu(ph)
    pic.left = Emu(int(l * UNIT + (w * UNIT - pw) / 2))
    pic.top = Emu(int(t * UNIT + (h * UNIT - ph) / 2))
    if name:
        pic.name = name
    return pic


def _cell_border(cell, edges, color, weight_pt):
    """给表格单元格画边线。python-pptx 无 API，必须手写 a:lnL/lnR/lnT/lnB。"""
    tcPr = cell._tc.get_or_add_tcPr()
    tag_of = {"L": "a:lnL", "R": "a:lnR", "T": "a:lnT", "B": "a:lnB"}
    for edge in edges:
        tag = tag_of[edge]
        old = tcPr.find(qn(tag))
        if old is not None:
            tcPr.remove(old)
        ln = etree.SubElement(tcPr, qn(tag))
        ln.set("w", str(int(weight_pt * 12700)))
        ln.set("cap", "flat")
        ln.set("cmpd", "sng")
        ln.set("algn", "ctr")
        fill = etree.SubElement(ln, qn("a:solidFill"))
        clr = etree.SubElement(fill, qn("a:srgbClr"))
        clr.set("val", color)
        etree.SubElement(ln, qn("a:prstDash")).set("val", "solid")


# ============================================================ 母版克隆
_ARCH_ORDER = ["cover", "toc", "body", "body_conclusion", "close"]

R_NS_URI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _strip_shape_tags(el):
    """剥离 `p:nvPr/p:custDataLst/p:tags` —— 形状标签（tag part）元数据。

    为什么必须剥：标签用的是 `r:id`，而新页的关系表里 rId 是**重新分配**的。
    拷贝过来的 `r:id="rId1"` 可能正好撞上「版式」关系，于是校验 rId 存在
    也查不出来，但 `p:tags` 指向的是 slideLayout 而不是 tags part，
    PowerPoint 直接判文件损坏（只弹一句 could not open，不给任何细节）。

    这些标签只是「形状分组/标记」的元数据，对渲染无影响，整块删掉最稳。
    """
    for cd in list(el.iter(qn("p:custDataLst"))):
        parent = cd.getparent()
        if parent is not None:
            parent.remove(cd)


def _drop_dangling_rids(el, slide_part):
    """删掉元素树里指向「新页不存在的 rId」的属性与容器。

    页型样板页里的形状常带 `p:nvPr/p:custDataLst/p:tags r:id="rId1"` 这类
    形状标签引用（原 PPT 里每个形状都绑了 tag）。深拷贝到新页后这些 rId
    在新页的关系表里不存在 → PowerPoint 判定文件损坏、拒绝打开，
    而且没有任何可读报错（只弹「PowerPoint could not open the file」）。

    实测：封面/结束页没有 tag 关系所以能开，目录页与全部正文页都带，
    因此表现得像是"只有某些页型坏"。
    """
    rels = slide_part.part.rels if hasattr(slide_part, "part") else slide_part.rels
    valid = set(rels.keys())
    KILL_TAGS = {"tags", "custDataLst"}

    for node in list(el.iter()):
        bad = [a for a in node.attrib
               if a.startswith("{%s}" % R_NS_URI) and node.attrib[a] not in valid]
        if not bad:
            continue
        # p:tags 缺 r:id 是非法 XML，所以整块容器一起删
        target = node
        if node.tag.split("}")[-1] in KILL_TAGS:
            aim = node
            while aim is not None and aim.tag.split("}")[-1] in KILL_TAGS:
                target = aim
                aim = aim.getparent()
        for a in bad:
            node.attrib.pop(a, None)
        if target.getparent() is not None and target is not el:
            target.getparent().remove(target)


class Template:
    """载入母版，按角色名持有 5 张样板页。"""

    def __init__(self, path):
        self.prs = Presentation(str(path))
        self.arch = {}
        for name, slide in zip(_ARCH_ORDER, self.prs.slides):
            self.arch[name] = slide
        self.layout_body = self._layout("1_标题和内容")
        self.layout_blank = self._layout("空白")

    def _layout(self, name):
        for lay in self.prs.slide_layouts:
            if lay.name == name:
                return lay
        return self.prs.slide_layouts[6]

    def shape(self, arch, role):
        for sh in self.arch[arch].shapes:
            if sh.name == role:
                return sh
        return None

    def clone(self, arch):
        """在末尾新建一页并深拷贝样板页的形状（含图片关系重挂）。"""
        layout = self.layout_blank if arch in ("cover", "toc", "close") else self.layout_body
        new = self.prs.slides.add_slide(layout)
        src, dst = self.arch[arch].shapes._spTree, new.shapes._spTree
        R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        for el in src.iterchildren():
            if el.tag.split("}")[-1] in ("nvGrpSpPr", "grpSpPr"):
                continue
            ph = el.find(".//" + qn("p:ph"))
            if ph is not None and ph.get("type") != "sldNum":
                # 内容/标题占位符不复制：add_slide 已经从版式带出对应占位符，
                # 再复制一份会撞 idx。但**页码占位符要复制**——版式里的那个
                # 不会被 add_slide 实例化到 shape 树，不复制就没有页码。
                continue
            dup = copy.deepcopy(el)
            for node in dup.iter():
                for attr in ("embed", "link"):
                    key = "{%s}%s" % (R_NS, attr)
                    if key in node.attrib:
                        target = self.arch[arch].part.related_part(node.attrib[key])
                        if target is not None:
                            node.attrib[key] = new.part.relate_to(
                                target, "http://schemas.openxmlformats.org/"
                                        "officeDocument/2006/relationships/image")
            _strip_shape_tags(dup)
            _drop_dangling_rids(dup, new)
            dst.append(dup)
        # 清掉版式带出的空占位符，只留页码
        for sh in list(new.shapes):
            if not sh.is_placeholder:
                continue
            try:
                if sh.placeholder_format.type == 13:      # SLIDE_NUMBER
                    continue
            except Exception:
                pass
            sh._element.getparent().remove(sh._element)
        return new

    def drop_archetypes(self):
        """删掉 5 张样板页，并同步清理 p14:sectionLst 里的悬空 slide id。

        坑：<p:sldId id="634"> 的 id 是**字符串**，而 Slide.slide_id 是 **int**。
        直接 `if slide_id in {s.get("id") ...}` 永远为 False，样板页一个都删不掉，
        还会让 p14:sectionLst 保持完整、文件看起来"正常"。必须显式转 str/int 对齐。
        """
        arch_ids = {str(self.arch[n].slide_id) for n in _ARCH_ORDER}
        for sldId in list(self.prs.slides._sldIdLst):
            if str(sldId.get("id")) in arch_ids:
                self.prs.part.drop_rel(sldId.rId)
                self.prs.slides._sldIdLst.remove(sldId)
        P14 = "{http://schemas.microsoft.com/office/powerpoint/2010/main}"
        live = {str(s.get("id")) for s in self.prs.slides._sldIdLst}
        for section in self.prs.part._element.iter(P14 + "section"):
            lst = section.find(P14 + "sldIdLst")
            if lst is None:
                continue
            for sid in list(lst):
                if str(sid.get("id")) not in live:
                    lst.remove(sid)

    def save(self, out):
        out = Path(out)
        out.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(out))


# ============================================================ 页型渲染器
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


def strip_timing(slide):
    for t in slide._element.findall(P_NS + "timing"):
        slide._element.remove(t)


def set_page_number(slide, n):
    for ph in slide.placeholders:
        if ph.placeholder_format.type == 13:      # SLIDE_NUMBER
            for p in ph.text_frame.paragraphs:
                for r in p.runs:
                    r.text = ""
            tf = ph.text_frame
            tf.text = f"-{n}-"
            for p in tf.paragraphs:
                for r in p.runs:
                    _style_run(r, S_SMALL, False, TEXT_SUB)
            return


def _find(slide, name):
    """按名字找形状，递归穿透组合。"""
    return _find_in(slide.shapes, name)


def _find_in(shapes, name):
    for sh in shapes:
        if sh.name == name:
            return sh
        if sh.shape_type == 6:      # GROUP
            hit = _find_in(sh.shapes, name)
            if hit is not None:
                return hit
    return None


def _shrink_to_fit(shape, text, size, box_w, height_cap, min_size=S_DENSE):
    """单行放不下就逐级降字号；降到下限还不行返回 None 表示该换行。

    用于封面标题这类"版面高度固定"的位置：宁可字小一点，也不要撑破图片区。
    """
    s = size
    while s >= min_size:
        if wrap_lines(text, s, box_w) * s * 1.15 / 72.0 * 100.0 <= height_cap:
            return s
        s -= 2
    return None


def _set_shape_text(shape, text, size=None, bold=None, color=None,
                    spacing=None, align=None):
    """覆写一个既有形状的文字，保留第一个 run 的位置但重设字体。"""
    tf = shape.text_frame
    for p in list(tf.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    p0 = tf.paragraphs[0]
    for r in list(p0.runs):
        r._r.getparent().remove(r._r)
    if align is not None:
        p0.alignment = align
    if spacing is not None:
        p0.line_spacing = spacing
    run = p0.add_run()
    run.text = text
    if size is not None:
        _style_run(run, size, bool(bold) if bold is not None else False,
                   color if color is not None else "000000")


def _require(slide, role, warn, where=""):
    """按角色名取形状；找不到就报警。

    静默返回 None 是最危险的失败模式：生成"成功"、页数正确、没有异常，
    但那一块内容其实是模板默认文案。宁可吵，也不要安静地交一份错稿。
    """
    sh = _find(slide, role)
    if sh is None:
        warn.append(f"{where}找不到母版角色形状 {role!r}，该处内容未被写入")
    return sh


def render_cover(tpl, page, n, warn=None):
    warn = warn if warn is not None else []
    slide = tpl.clone("cover")
    title = page["title"]
    t = _require(slide, "yk_cover_title", warn, "封面")
    if t is not None:
        # 封面标题框高 91（1/100 inch），下面是汇报人条，撑破就会压上去。
        # 原文件 48pt 单行；超过 22 字就逐级降到能放下为止。
        box_w = Emu(t.width).inches * 100
        size = _shrink_to_fit(t, title, S_COVER, box_w, 91) or S_DENSE
        _set_shape_text(t, title, size=size, bold=True, color=NAVY,
                        align=PP_ALIGN.LEFT, spacing=1.0)
    m = _require(slide, "yk_cover_meta", warn, "封面")
    if m is not None:
        meta = (f" 汇报人：{page.get('presenter', '')}     "
                f"导师：{page.get('advisor', '')}       {page.get('date', '')} ")
        _set_shape_text(m, meta, size=S_BODY, bold=False, color=WHITE, spacing=1.0)
    strip_timing(slide)
    return slide


TOC_ENTRIES = ["研究背景", "相关工作", "研究内容", "总结与展望", "质询与回答"]


def render_toc(tpl, page, n):
    slide = tpl.clone("toc")
    entries = page.get("entries") or TOC_ENTRIES
    current = page.get("current")
    slots = []
    for i in range(1, len(TOC_ENTRIES) + 1):
        g = _find(slide, f"yk_toc_{i}")
        if g is not None:
            slots.append((i, g))
    # 超出母版条目数的直接删掉；不足的重新均分纵向位置
    for i, g in slots:
        if i > len(entries):
            g._element.getparent().remove(g._element)
    slots = [(i, g) for i, g in slots if i <= len(entries)]
    if slots:
        # 母版 5 条位于 y 151 / 245 / 339 / 433 / 528；条目变少时均分这段纵向区间
        top0, y_last = 151, 528
        step = (y_last - top0) / (len(slots) - 1) if len(slots) > 1 else 0
        for k, (i, g) in enumerate(slots):
            g.top = Emu(int((top0 + k * step) * UNIT))
        for k, (i, g) in enumerate(slots):
            txt = _find(slide, f"yk_toc_{i}_text")
            if txt is None:
                continue
            label = entries[k]
            hot = current is not None and label == current
            _set_shape_text(txt, label, size=S_TOC_ENTRY, bold=True,
                            color=RED2 if hot else NAVY, spacing=1.0)
    strip_timing(slide)
    return slide


def _body_base(tpl, page, n, with_conclusion):
    arch = "body_conclusion" if with_conclusion else "body"
    slide = tpl.clone(arch)
    sec = _find(slide, "yk_section")
    if sec is not None:
        _set_shape_text(sec, page.get("section", ""), size=S_TITLE, bold=True, color=NAVY,
                        align=PP_ALIGN.LEFT)
    ttl = _find(slide, "yk_title")
    if ttl is not None:
        _set_shape_text(ttl, page["title"], size=S_TITLE, bold=True, color=WHITE,
                        align=PP_ALIGN.CENTER)
    if with_conclusion:
        ct = _find(slide, "yk_concl_text")
        if ct is not None:
            _set_shape_text(ct, page.get("conclusion", ""), size=S_BODY, bold=True,
                            color="000000", align=PP_ALIGN.CENTER)
    set_page_number(slide, n)
    strip_timing(slide)
    return slide


def _lead_block(slide, page, warn):
    """画定性句 + 分隔线。返回内容区上界。"""
    lead = page.get("lead", "")
    if not lead:
        warn.append("缺少 lead（定性句）：每页应用一句话说明本页结论")
        return CONTENT_TOP
    lines = wrap_lines(lead, S_LEAD, BODY_R - BODY_L - 35)
    if lines > 2:
        warn.append(f"lead 折成 {lines} 行，应在 60 字内")
    add_text(slide, 53, 108, BODY_R - BODY_L - 35, 78, lead,
             size=S_LEAD, bold=True, color="000000", spacing=1.45,
             align=PP_ALIGN.JUSTIFY, name="yk_lead")
    add_line(slide, 60, SEP_Y, 1203, 0, color=SKY, weight=1.0, name="yk_sep")
    return CONTENT_TOP


def render_lead(tpl, page, n, warn):
    slide = _body_base(tpl, page, n, bool(page.get("conclusion")))
    _lead_block(slide, page, warn)
    return slide


def render_bullets(tpl, page, n, warn):
    slide = _body_base(tpl, page, n, bool(page.get("conclusion")))
    top = _lead_block(slide, page, warn)
    y = top + 6
    bottom = CONTENT_BOTTOM_CONCL if page.get("conclusion") else CONTENT_BOTTOM
    items = page.get("bullets") or []
    for it in items:
        h = text_height(it, S_BODY, 1170, 1.5)
        if y + h + 14 > bottom:
            warn.append(f"bullets 条目溢出：{it[:14]}…")
            break
        add_rect(slide, 60, y + 9, 14, 14, BRAND)
        add_text(slide, 92, y, 1170, h + 8, it, size=S_BODY, bold=False,
                 spacing=1.5, align=PP_ALIGN.JUSTIFY)
        y += h + 20
    return slide


def render_cards(tpl, page, n, warn):
    slide = _body_base(tpl, page, n, bool(page.get("conclusion")))
    top = _lead_block(slide, page, warn)
    cards = page.get("cards") or []
    k = len(cards)
    if k == 0:
        return slide
    gap = 32 if k == 2 else 28
    w = (BODY_R - BODY_L - gap * (k - 1)) / k
    bottom = CONTENT_BOTTOM_CONCL if page.get("conclusion") else CONTENT_BOTTOM
    has_img = any(c.get("image") for c in cards)

    if has_img:
        # 有图：文字在上（y top..top+112），图卡在下
        txt_h = 112
        for i, c in enumerate(cards):
            l = BODY_L + i * (w + gap)
            add_text(slide, l + 8, top, w - 16, 30, c.get("head", ""),
                     size=S_LEAD, bold=True, color=NAVY, spacing=1.2, align=PP_ALIGN.LEFT)
            add_text(slide, l + 8, top + 34, w - 16, txt_h - 34, c.get("text", ""),
                     size=S_DENSE, bold=False, spacing=1.5, align=PP_ALIGN.JUSTIFY)
            card_t = top + txt_h + 8
            card_h = bottom - card_t
            add_round_rect(slide, l, card_t, w, card_h, fill=WHITE, line=PALE2, line_w=0.25)
            if c.get("image"):
                add_picture_fit(slide, c["image"], l + 12, card_t + 10, w - 24, card_h - 20)
    else:
        card_h = bottom - top - 8
        for i, c in enumerate(cards):
            l = BODY_L + i * (w + gap)
            add_round_rect(slide, l, top, w, card_h, fill=PALE, line=PALE2, line_w=0.25)
            add_text(slide, l + 18, top + 14, w - 36, 28, c.get("head", ""),
                     size=S_LEAD, bold=True, color=NAVY, spacing=1.2, align=PP_ALIGN.LEFT)
            body = c.get("text", "")
            avail = card_h - 60
            if text_height(body, S_BODY, w - 36, 1.5) > avail:
                warn.append(f"卡片正文溢出（{w - 36:.0f} 宽）：{plain(body)[:16]}…")
            add_text(slide, l + 18, top + 50, w - 36, avail, body, size=S_BODY,
                     bold=False, spacing=1.5, align=PP_ALIGN.JUSTIFY)
    return slide


def render_flow(tpl, page, n, warn):
    """竖向流程，最多 3 步（4 步会超出画布）。"""
    slide = _body_base(tpl, page, n, bool(page.get("conclusion")))
    top = _lead_block(slide, page, warn)
    steps = page.get("steps") or []
    if len(steps) > 3:
        warn.append(f"flow 竖向最多 3 步，当前 {len(steps)} 步；请改用 flow_h")
        steps = steps[:3]
    box_w, box_h = 835, 68
    step = 178
    for i, s in enumerate(steps):
        y = top + i * step
        add_rect(slide, 227, y, box_w, box_h, WHITE, line="142E79", line_w=1.5)
        txt = s.get("text", "") if isinstance(s, dict) else str(s)
        add_text(slide, 237, y, box_w - 20, box_h, txt, size=S_LEAD, bold=True,
                 color="000000", spacing=1.2, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
        tag = s.get("tag") if isinstance(s, dict) else None
        if tag:
            add_text(slide, 1082, y, 200, box_h, tag, size=S_BODY, bold=True,
                     color=NAVY, spacing=1.2, align=PP_ALIGN.LEFT,
                     anchor=MSO_ANCHOR.MIDDLE)
        if i < len(steps) - 1:
            ar = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW,
                                        Emu(int(570 * UNIT)), Emu(int((y + box_h + 4) * UNIT)),
                                        Emu(int(134 * UNIT)), Emu(int(50 * UNIT)))
            ar.fill.solid()
            ar.fill.fore_color.rgb = _rgb("142E79")
            ar.line.color.rgb = _rgb("002060")
            ar.line.width = Pt(1.25)
            ar.shadow.inherit = False
    return slide


def render_flow_h(tpl, page, n, warn):
    """横向流程，4–5 步。"""
    slide = _body_base(tpl, page, n, bool(page.get("conclusion")))
    top = _lead_block(slide, page, warn)
    steps = page.get("steps") or []
    k = len(steps)
    if k == 0:
        return slide
    gap = 30
    w = (BODY_R - BODY_L - gap * (k - 1)) / k
    box_h = 120
    y = top + 60
    for i, s in enumerate(steps):
        l = BODY_L + i * (w + gap)
        add_round_rect(slide, l, y, w, box_h, fill=PALE, line=BRAND, line_w=1.0)
        txt = s.get("text", "") if isinstance(s, dict) else str(s)
        num = i + 1
        o = slide.shapes.add_shape(MSO_SHAPE.OVAL, Emu(int((l + 8) * UNIT)),
                                   Emu(int((y - 20) * UNIT)),
                                   Emu(int(40 * UNIT)), Emu(int(40 * UNIT)))
        o.fill.solid()
        o.fill.fore_color.rgb = _rgb(BRAND)
        o.line.fill.background()
        o.shadow.inherit = False
        tf = o.text_frame
        tf.text = str(num)
        for p in tf.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for r in p.runs:
                _style_run(r, S_SMALL, True, WHITE)
        add_text(slide, l + 10, y + 8, w - 20, box_h - 16, txt, size=S_BODY, bold=True,
                 color="000000", spacing=1.3, align=PP_ALIGN.CENTER,
                 anchor=MSO_ANCHOR.MIDDLE)
        if i < k - 1:
            ax = l + w + 4
            ar = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                        Emu(int(ax * UNIT)), Emu(int((y + box_h / 2 - 12) * UNIT)),
                                        Emu(int((gap - 8) * UNIT)), Emu(int(24 * UNIT)))
            ar.fill.solid()
            ar.fill.fore_color.rgb = _rgb(BRAND)
            ar.line.fill.background()
            ar.shadow.inherit = False
    return slide


def render_table(tpl, page, n, warn):
    slide = _body_base(tpl, page, n, bool(page.get("conclusion")))
    top = _lead_block(slide, page, warn)
    spec = page.get("table") or {}
    head = spec.get("head") or []
    rows = spec.get("rows") or []
    if not head or not rows:
        warn.append("table 页缺少 head / rows")
        return slide
    ncol = len(head)
    width = spec.get("width") or 900
    left = spec.get("x")
    if left is None:
        left = (CANVAS_W - width) / 2
    top_y = spec.get("y") or (top + 30)
    head_h, row_h = 46, 41
    total_h = head_h + row_h * len(rows)
    gf = slide.shapes.add_table(len(rows) + 1, ncol,
                                Emu(int(left * UNIT)), Emu(int(top_y * UNIT)),
                                Emu(int(width * UNIT)), Emu(int(total_h * UNIT)))
    gf.name = "yk_table"
    tbl = gf.table
    tbl.first_row = True
    tbl.horz_banding = False
    col_w = width / ncol
    for c in tbl.columns:
        c.width = Emu(int(col_w * UNIT))
    tbl.rows[0].height = Emu(int(head_h * UNIT))
    for r in list(tbl.rows)[1:]:
        r.height = Emu(int(row_h * UNIT))

    def put(cell, text, size, bold, color, fill):
        cell.fill.solid()
        cell.fill.fore_color.rgb = _rgb(fill)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.margin_left = cell.margin_right = Emu(int(9 * UNIT))
        cell.margin_top = cell.margin_bottom = Emu(int(4.5 * UNIT))
        tf = cell.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        p.line_spacing = 1.2
        for r_ in list(p.runs):
            r_._r.getparent().remove(r_._r)
        run = p.add_run()
        run.text = plain(str(text))
        _style_run(run, size, bold, color)
        # 只画上下细线，横向分隔；不画左右框线（与原模板一致）
        _cell_border(cell, ["T", "B"], PALE2, 0.75)

    for j, h in enumerate(head):
        put(tbl.cell(0, j), h, S_BODY, True, WHITE, TABLE_HEAD)
    for i, row in enumerate(rows):
        fill = PALE if i % 2 == 0 else WHITE
        for j in range(ncol):
            val = row[j] if j < len(row) else ""
            put(tbl.cell(i + 1, j), val, S_BODY, False, "000000", fill)
    return slide


def render_image(tpl, page, n, warn):
    slide = _body_base(tpl, page, n, bool(page.get("conclusion")))
    top = _lead_block(slide, page, warn)
    bottom = CONTENT_BOTTOM_CONCL if page.get("conclusion") else CONTENT_BOTTOM
    cap = page.get("caption")
    cap_h = 24 if cap else 0
    add_picture_fit(slide, page["image"], BODY_L, top, BODY_R - BODY_L,
                    bottom - top - cap_h, name="yk_image")
    if cap:
        add_text(slide, BODY_L, bottom - cap_h + 4, BODY_R - BODY_L, cap_h, cap,
                 size=S_SMALL, bold=False, color=TEXT_SUB, spacing=1.2,
                 align=PP_ALIGN.CENTER, name="yk_caption")
    return slide


def render_close(tpl, page, n):
    slide = tpl.clone("close")
    t = _find(slide, "yk_close_text")
    if t is not None:
        _set_shape_text(t, "请各位老师同学批评指正！", size=S_COVER, bold=True,
                        color=NAVY, align=PP_ALIGN.CENTER, spacing=1.0)
    strip_timing(slide)
    return slide


RENDERERS = {
    "cover": lambda tpl, p, n, w: render_cover(tpl, p, n),
    "toc": lambda tpl, p, n, w: render_toc(tpl, p, n),
    "lead": render_lead,
    "bullets": render_bullets,
    "cards": render_cards,
    "flow": render_flow,
    "flow_h": render_flow_h,
    "table": render_table,
    "image": render_image,
    "close": lambda tpl, p, n, w: render_close(tpl, p, n),
}


# ============================================================ 编排入口
def render_deck(template_path, pages, out_path):
    """
    pages: [{type, section, title, lead, conclusion, ...}, ...]
    返回 (输出路径, 警告列表)。
    """
    tpl = Template(template_path)
    warnings = []
    n_content = 0
    for page in pages:
        kind = page.get("type", "lead")
        if kind not in RENDERERS:
            raise ValueError(f"未知页型：{kind}（可用：{', '.join(RENDERERS)}）")
        page_warn = []
        RENDERERS[kind](tpl, page, 0, page_warn)
        # 页码：只给正文页编号
        if kind not in ("cover", "toc", "close"):
            n_content += 1
            set_page_number(tpl.prs.slides[-1], n_content)
        for w in page_warn:
            warnings.append(f"[第 {len(tpl.prs.slides)} 页] {w}")
    tpl.drop_archetypes()
    tpl.save(out_path)
    return out_path, warnings
