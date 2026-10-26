#!/usr/bin/env python
"""按《学术论文 Word 排版与生成规范》对 docx 做自动质检。

用法：
    python qa_check.py out/report.docx
    python qa_check.py out/report.docx --pdf out/report.pdf   # 附加 PDF 页数/体检
    python qa_check.py out/report.docx --strict               # 警告也算失败

输出：分组的检查结果 + 失败/警告统计；有 FAIL 时退出码 1。
"""

from __future__ import annotations

import argparse
import os
import re
import sys

import _console
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Emu

CM_TWIPS = 567.0
# 段落开头出现任何一个都算"用了项目符号"
BULLETS = "•●○▪◆■◦‣∙·"
# 这些符号出现在任何位置都是问题（不会作为正常标点出现）
BULLETS_ANYWHERE = "•●○▪◆■"

RESULTS = []   # (level, group, name, detail)
COUNTS = {"PASS": 0, "FAIL": 0, "WARN": 0}

# 期望值（默认取规范原值；用 --spec 传覆盖时可以校验换肤后的文档）
EXPECT = {
    "navy": "1F3A5F",
    "blue": "0072B2",
    "body_pt": 12,
    "body_first": "480",   # 2 字符 = 2 × 字号 × 20 twips
    "h1_sz": "32",
    "h2_sz": "26",
    "caption_pt": 9.5,
    "reference_pt": 9,
    "header_pt": 8.5,
    "footer_pt": 8,
    "doc_colors": None,   # --spec 覆盖后的文档色（换肤时校验用）
}


def apply_spec(path):
    """按 spec 覆盖文件（与 build_docx.py --spec 同一份）校准期望值。"""
    import json
    with open(os.path.abspath(path), encoding="utf-8-sig") as fh:
        spec = json.load(fh)
    colors = spec.get("colors") or {}
    sizes = spec.get("sizes") or {}
    if colors.get("navy"):
        EXPECT["navy"] = str(colors["navy"]).upper()
    if colors.get("blue"):
        EXPECT["blue"] = str(colors["blue"]).upper()
    for key, field in (("body", "body_pt"), ("h1", "h1_sz"), ("h2", "h2_sz")):
        if key in sizes and field != "body_pt":
            EXPECT[field] = str(int(round(float(sizes[key]) * 2)))
    if "body" in sizes:
        EXPECT["body_pt"] = float(sizes["body"])
        EXPECT["body_first"] = str(int(round(EXPECT["body_pt"] * 40)))
    for key in ("caption", "reference", "header", "footer"):
        if key in sizes:
            EXPECT[key + "_pt"] = float(sizes[key])

    # 换肤后：被覆盖的文档色也算"批准色"，否则配色检查会误判 FAIL
    import palette

    overrides = set()
    role_map = {"navy": "heading", "blue": "accent", "orange": "warn",
                "green": "ok", "ink": "ink", "gray": "muted",
                "pale": "pale", "pale_line": "rule"}
    for key, value in colors.items():
        if key in role_map and value:
            overrides.add(palette.normalize(str(value)))
    EXPECT["doc_colors"] = overrides or None


def check(group, name, ok, detail="", warn=False):
    if ok:
        level = "PASS"
    elif warn:
        level = "WARN"
    else:
        level = "FAIL"
    RESULTS.append((level, group, name, detail))
    COUNTS[level] += 1
    return ok


def warn_if(group, name, bad, detail=""):
    if bad:
        RESULTS.append(("WARN", group, name, detail))
        COUNTS["WARN"] += 1
    else:
        RESULTS.append(("PASS", group, name, ""))
        COUNTS["PASS"] += 1


def a(el, name):
    return None if el is None else el.get(qn("w:" + name))


def twips(value):
    """Length/twips 属性 → 数字 twips。"""
    if value is None:
        return None
    try:
        return int(round(Emu(int(value)).twips))
    except Exception:
        return int(value)


def style_of(doc, name):
    for st in doc.styles:
        if st.name == name:
            return st
    return None


def rpr(run_el):
    return run_el.find(qn("w:rPr"))


def _doc_defaults_rfonts(doc):
    styles_el = doc.styles.element
    dd = styles_el.find(qn("w:docDefaults"))
    if dd is None:
        return None
    rpr_default = dd.find(qn("w:rPrDefault"))
    rPr = rpr_default.find(qn("w:rPr")) if rpr_default is not None else None
    return rPr.find(qn("w:rFonts")) if rPr is not None else None


def effective_size(p):
    """run 上的字号优先；Word 保存后可能把与样式相同的直接格式收进样式，所以要回退。"""
    for r in p.runs:
        if r.font.size is not None:
            return r.font.size.pt
    st = p.style
    while st is not None:
        if st.font.size is not None:
            return st.font.size.pt
        st = st.base_style
    return None


def effective_color(p):
    for r in p.runs:
        if r.font.color is not None and r.font.color.rgb is not None:
            return str(r.font.color.rgb)
    st = p.style
    while st is not None:
        if st.font.color is not None and st.font.color.rgb is not None:
            return str(st.font.color.rgb)
        st = st.base_style
    return None


# --------------------------------------------------------------- 检查实现

def check_page(doc):
    sec = doc.sections[0]
    sp = sec._sectPr
    pg = sp.find(qn("w:pgSz"))
    mar = sp.find(qn("w:pgMar"))
    check("页面", "纸张 A4 21×29.7 cm",
          pg is not None and abs(int(a(pg, "w")) - 11906) <= 30
          and abs(int(a(pg, "h")) - 16838) <= 40,
          "pgSz=%s×%s" % (a(pg, "w"), a(pg, "h")) if pg is not None else "缺少 pgSz")
    want = {"top": 2.4, "bottom": 2.2, "left": 2.6, "right": 2.4,
            "header": 1.2, "footer": 1.0, "gutter": 0.0}
    bad = []
    for key, cm in want.items():
        got = a(mar, key) if mar is not None else None
        if got is None:
            bad.append("%s 缺失" % key)
            continue
        if abs(int(got) - cm * CM_TWIPS) > 8:
            bad.append("%s=%.2fcm" % (key, int(got) / CM_TWIPS))
    check("页面", "页边距 / 页眉页脚距离", not bad, ", ".join(bad))
    title_pg = sp.find(qn("w:titlePg"))
    check("页面", "首页不同（封面无页眉页脚）",
          title_pg is not None and a(title_pg, "val") not in ("false", "0"),
          "缺少 w:titlePg" if title_pg is None else "w:titlePg 被关闭")


def check_settings(doc):
    settings = doc.settings.element
    view = settings.find(qn("w:view"))
    zoom = settings.find(qn("w:zoom"))
    check("视图", "默认页面视图 w:view=print",
          a(view, "val") == "print", "w:view=%s" % a(view, "val"))
    check("视图", "缩放 100%",
          a(zoom, "percent") == "100" or a(zoom, "val") == "100",
          "w:zoom=%s" % (a(zoom, "percent") or a(zoom, "val")))


def check_styles(doc):
    normal = style_of(doc, "Normal")
    defaults = _doc_defaults_rfonts(doc)
    ok = normal is not None
    if ok:
        rPr = normal.element.find(qn("w:rPr"))
        rFonts = rPr.find(qn("w:rFonts")) if rPr is not None else None
        pPr = normal.element.find(qn("w:pPr"))
        sz = rPr.find(qn("w:sz")) if rPr is not None else None
        spacing = pPr.find(qn("w:spacing")) if pPr is not None else None
        ind = pPr.find(qn("w:ind")) if pPr is not None else None
        jc = pPr.find(qn("w:jc")) if pPr is not None else None
        cn = a(rFonts, "eastAsia") or a(defaults, "eastAsia")
        en = a(rFonts, "ascii") or a(defaults, "ascii")
        ok = (cn == "宋体" and en == "Times New Roman"
              and a(sz, "val") == str(int(EXPECT["body_pt"] * 2))
              and a(spacing, "line") == "360"
              and a(ind, "firstLine") == EXPECT["body_first"]
              and a(jc, "val") == "both")
        detail = "eastAsia=%s ascii=%s sz=%s line=%s firstLine=%s jc=%s" % (
            cn, en, a(sz, "val"), a(spacing, "line"),
            a(ind, "firstLine"), a(jc, "val"))
    else:
        detail = "缺少 Normal 样式"
    check("样式", "正文：宋体+Times 12pt / 1.5 倍 / 首行 24pt / 两端对齐", ok, detail)

    h1 = style_of(doc, "Heading 1")
    detail = ""
    ok = h1 is not None
    if ok:
        el = h1.element
        rPr = el.find(qn("w:rPr"))
        pPr = el.find(qn("w:pPr"))
        sz = rPr.find(qn("w:sz"))
        col = rPr.find(qn("w:color"))
        b = rPr.find(qn("w:b"))
        pBdr = pPr.find(qn("w:pBdr"))
        bottom = pBdr.find(qn("w:bottom")) if pBdr is not None else None
        keep = pPr.find(qn("w:keepNext"))
        ok = (a(sz, "val") == EXPECT["h1_sz"] and a(col, "val") == EXPECT["navy"]
              and b is not None and a(b, "val") not in ("false", "0")
              and a(bottom, "sz") == "6" and a(bottom, "color") == EXPECT["blue"]
              and keep is not None)
        detail = "sz=%s color=%s border=%s keepNext=%s" % (
            a(sz, "val"), a(col, "val"),
            (a(bottom, "sz"), a(bottom, "color")) if bottom is not None else None,
            keep is not None)
    check("样式", "一级标题：黑体 16pt 加粗 #1F3A5F + 蓝色下边框", ok, detail)

    h2 = style_of(doc, "Heading 2")
    ok = h2 is not None
    detail = ""
    if ok:
        el = h2.element
        rPr = el.find(qn("w:rPr"))
        pPr = el.find(qn("w:pPr"))
        sz = rPr.find(qn("w:sz"))
        col = rPr.find(qn("w:color"))
        keep = pPr.find(qn("w:keepNext"))
        ok = (a(sz, "val") == EXPECT["h2_sz"] and a(col, "val") == EXPECT["blue"]
              and keep is not None)
        detail = "sz=%s color=%s keepNext=%s" % (a(sz, "val"), a(col, "val"),
                                                 keep is not None)
    check("样式", "二级标题：黑体 13pt 加粗 #0072B2", ok, detail)


def _all_paragraphs(doc):
    for p in doc.paragraphs:
        yield p
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p


_EXCLUDE = (
    r"^\[\d+\]",        # 参考文献
    r"^（\d+）",         # （1）（2）列表
    r"^注：",            # 表注 / 图注
    r"^关键词", r"^Keywords",
    r"^图\s*\d", r"^表\s*\d",   # 图表题
)


def _prose(doc, min_len=40):
    """识别"正文段落"：Normal 样式、12 pt、非题注/参考文献/列表/英文摘要等特殊段。"""
    out = []
    for p in doc.paragraphs:
        text = p.text.strip()
        if len(text) < min_len:
            continue
        if p.style is None or p.style.name != "Normal":
            continue
        if any(re.match(pat, text) for pat in _EXCLUDE):
            continue
        size = effective_size(p)
        if size is not None and abs(size - 12) > 0.01:
            continue          # 封面说明行 11 pt、英文摘要 11 pt、图注 8.5 pt…
        out.append(p)
    return out


def _p_indent(p):
    pPr = p._p.find(qn("w:pPr"))
    return pPr.find(qn("w:ind")) if pPr is not None else None


def _p_spacing(p):
    pPr = p._p.find(qn("w:pPr"))
    return pPr.find(qn("w:spacing")) if pPr is not None else None


def check_body_paragraphs(doc):
    prose = _prose(doc)
    if not prose:
        warn_if("正文", "未识别到 12 pt 正文段落（若文档只有封面/摘要可忽略）", True)
        return
    check("正文", "识别到 %d 个 12 pt 正文段落" % len(prose), True)
    bad_indent = [p.text.strip()[:16] for p in prose
                  if a(_p_indent(p), "firstLineChars") != "200"]
    ratio = 1 - len(bad_indent) / len(prose)
    check("正文", "正文段落首行缩进 2 字符（≥90% 覆盖率）", ratio >= 0.9,
          "%d/%d 段落不合规，例如：%s" % (len(bad_indent), len(prose),
                                     "、".join(bad_indent[:3])))

    bad_line = [p.text.strip()[:16] for p in prose
                if a(_p_spacing(p), "line") not in ("360", None)]
    ratio2 = 1 - len(bad_line) / len(prose)
    check("正文", "正文 1.5 倍行距（≥90% 覆盖率）", ratio2 >= 0.9,
          "%d/%d 段落不合规，例如：%s" % (len(bad_line), len(prose),
                                     "、".join(bad_line[:3])))

    bad_jc = [p.text.strip()[:16] for p in prose
              if a(p._p.find(qn("w:pPr")).find(qn("w:jc"))
                   if p._p.find(qn("w:pPr")) is not None else None, "val")
              not in ("both", None)]
    check("正文", "正文两端对齐", not bad_jc, "、".join(bad_jc[:3]))


def check_lists(doc):
    bad = []
    for p in _all_paragraphs(doc):
        text = p.text.lstrip()
        if text[:1] and text[0] in BULLETS:
            bad.append(text[:14])
        elif any(ch in text for ch in BULLETS_ANYWHERE):
            bad.append(text[:14])
        if p.style is not None and "List" in (p.style.name or ""):
            bad.append("样式 %s: %s" % (p.style.name, text[:10]))
        pPr = p._p.find(qn("w:pPr"))
        if pPr is not None and pPr.find(qn("w:numPr")) is not None:
            bad.append("numPr: %s" % text[:10])
    check("列表", "无圆点符号 / 无 List 样式 / 无自动编号", not bad,
          "、".join(bad[:3]))


def check_headers_footers(doc):
    sec = doc.sections[0]
    hdr = sec.header
    p = hdr.paragraphs[0]
    pPr = p._p.find(qn("w:pPr"))
    jc = a(pPr.find(qn("w:jc")) if pPr is not None else None, "val")
    size = effective_size(p)
    check("页眉页脚", "页眉右对齐、8.5 pt、有文字",
          jc == "right" and size == EXPECT["header_pt"] and p.text.strip() != "",
          "jc=%s size=%s text=%r" % (jc, size, p.text[:16]))

    ftr = sec.footer
    fp = ftr.paragraphs[0]
    instrs = [el.text or "" for el in fp._p.iter(qn("w:instrText"))]
    joined = " ".join(instrs)
    fsize = effective_size(fp)
    pPr = fp._p.find(qn("w:pPr"))
    fjc = a(pPr.find(qn("w:jc")) if pPr is not None else None, "val")
    check("页眉页脚", "页脚居中、8 pt、PAGE + NUMPAGES 域",
          fjc == "center" and fsize == EXPECT["footer_pt"] and "PAGE" in joined
          and "NUMPAGES" in joined,
          "jc=%s size=%s fields=%s" % (fjc, fsize, joined.strip()[:40]))

    try:
        first_hdr = sec.first_page_header
        first_ftr = sec.first_page_footer
        empty = (all(not x.text.strip() for x in first_hdr.paragraphs)
                 and all(not x.text.strip() for x in first_ftr.paragraphs))
    except Exception:
        empty = True
    check("页眉页脚", "封面页不显示页眉页脚（首页页眉页脚为空）", empty)


def _caption_kind(p):
    """识别"图表题"：居中 + 9.5 pt 的 `图N/表N` 段落。

    正文里也常出现"图6给出…"这类以图号开头的句子，只靠正则会把它们误判成题注
    （进而误报字号/加粗，并把编号算重）。所以题注必须同时满足：居中、字号 ≤10 pt。
    """
    if p._p.tag != qn("w:p"):        # 传进来的可能是表格节点
        return None, None
    text = (p.text or "").strip()
    m = re.match(r"^(图|表)\s*(\d+)", text)
    if not m:
        return None, None
    pPr = p._p.find(qn("w:pPr"))
    jc = a(pPr.find(qn("w:jc")) if pPr is not None else None, "val")
    if jc != "center":
        return None, None
    size = None
    color = None
    bold = False
    for r in p.runs:
        if r.font.size is not None:
            size = r.font.size.pt
        if r.font.color is not None and r.font.color.rgb is not None:
            color = str(r.font.color.rgb)
        bold = bold or bool(r.bold)
    if size is None or size > 10.0:
        return None, None
    return m.group(1), (int(m.group(2)), size, color, bold)


def check_tables_and_captions(doc):
    body = doc.element.body
    tables = [el for el in body.iter(qn("w:tbl"))]
    three_line = 0
    borderless = 0
    problems = []

    for tbl in tables:
        tblPr = tbl.find(qn("w:tblPr"))
        borders = tblPr.find(qn("w:tblBorders")) if tblPr is not None else None
        if borders is None:
            problems.append("表格缺少 tblBorders")
            continue
        spec = {e.tag.split("}")[-1]: a(e, "val") for e in borders}
        szs = {e.tag.split("}")[-1]: a(e, "sz") for e in borders}

        if all(v in (None, "nil", "none") for v in spec.values()):
            borderless += 1
            continue

        three_line += 1
        ok = (spec.get("top") == "single" and szs.get("top") == "12"
              and spec.get("bottom") == "single" and szs.get("bottom") == "12"
              and spec.get("left") in (None, "nil")
              and spec.get("right") in (None, "nil")
              and spec.get("insideV") in (None, "nil")
              and spec.get("insideH") in (None, "nil"))
        if not ok:
            problems.append("三线表边框不合规：%s" % spec)

        rows = tbl.findall(qn("w:tr"))
        if rows:
            trPr = rows[0].find(qn("w:trPr"))
            if trPr is None or trPr.find(qn("w:tblHeader")) is None:
                problems.append("首行未设置重复表头 w:tblHeader")
            first_cell = rows[0].find(qn("w:tc"))
            tcPr = first_cell.find(qn("w:tcPr"))
            tcB = tcPr.find(qn("w:tcBorders")) if tcPr is not None else None
            bottom = tcB.find(qn("w:bottom")) if tcB is not None else None
            if bottom is None or a(bottom, "sz") != "6":
                problems.append("栏目线不是 0.75 pt（sz=6）")
        for tr in rows:
            trPr = tr.find(qn("w:trPr"))
            if trPr is None or trPr.find(qn("w:cantSplit")) is None:
                problems.append("存在允许跨页拆分的表格行")
                break

    check("表格", "数据表为三线表（顶/底线 1.5pt、栏目线 0.75pt、无竖线）",
          not problems, "；".join(problems[:3]) + ("…" if len(problems) > 3 else ""))
    check("表格", "表格统计：三线表 %d 个 / 无边框表 %d 个"
          % (three_line, borderless), True)

    # 表题 / 图题
    captions_bad = []
    seen = {"图": [], "表": []}
    for p in doc.paragraphs:
        kind, info = _caption_kind(p)
        if not kind:
            continue
        num, size, color, bold = info
        seen[kind].append(num)
        if size != EXPECT["caption_pt"] or not bold or color != EXPECT["navy"]:
            captions_bad.append("%s%d(size=%s,bold=%s,color=%s)"
                                % (kind, num, size, bold, color))
    check("图表", "图题 / 表题：9.5 pt 加粗 #1F3A5F", not captions_bad,
          "、".join(captions_bad[:3]))

    for kind, nums in seen.items():
        if nums:
            expect = list(range(1, len(nums) + 1))
            check("图表", "%s编号从 1 连续不跳号" % kind, nums == expect,
                  "实际编号：%s" % nums)


def check_figures(doc):
    shapes = doc.inline_shapes
    if not shapes:
        warn_if("图片", "文档中没有图片（如不需要可忽略）", True)
        return
    bad = []
    for shape in shapes:
        w = Emu(shape.width).cm
        h = Emu(shape.height).cm
        if w > 15.5 + 0.01:
            bad.append("图片宽 %.2f cm 超过 15.5 cm" % w)
        # 等比校验（用包内原图比例）；缺 Pillow 时跳过，不算失败
        try:
            blob = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
            part = doc.part.related_parts[blob]
            import io

            from PIL import Image
            img = Image.open(io.BytesIO(part.blob))
            native = img.width / img.height
            if abs((w / h) - native) / native > 0.01:
                bad.append("图片比例被改变（%.3f vs %.3f）" % (w / h, native))
        except ImportError:
            pass          # 未装 Pillow：跳过比例校验，宽度检查仍然生效
        except Exception as exc:  # pragma: no cover
            bad.append("无法校验图片比例：%s" % exc)
    check("图片", "图片等比缩放且不超版心宽度", not bad, "；".join(bad[:3]))

    # 图片段落 keep_with_next + 图题在下方
    figs = [p for p in doc.paragraphs if p._p.findall(".//" + qn("w:drawing"))]
    bad2 = []
    for p in figs:
        pPr = p._p.find(qn("w:pPr"))
        keep = pPr.find(qn("w:keepNext")) if pPr is not None else None
        if keep is None:
            bad2.append("图片段落未设置 keep_with_next")
    check("图片", "图片与图题同页（keep_with_next）", not bad2, "；".join(bad2[:3]))


def check_equations(doc):
    maths = list(doc.element.body.iter(
        "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"))
    if not maths:
        warn_if("公式", "文档中没有 OMML 公式（如不需要可忽略）", True)
        return
    check("公式", "公式为 Word 原生 OMML（可编辑）", len(maths) > 0,
          "共 %d 个 oMath" % len(maths))
    nums = []
    for tbl in doc.tables:
        if len(tbl.rows) == 1 and len(tbl.columns) == 3:
            cells = tbl.rows[0].cells
            if cells[1]._tc.find(".//" + qn("m:oMath")) is not None or \
                    cells[1]._tc.findall(".//{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath"):
                m = re.search(r"[（(](\d+)[）)]", cells[2].text)
                if m:
                    nums.append(int(m.group(1)))
    if nums:
        check("公式", "公式编号连续、位于右侧单元格",
              nums == list(range(1, len(nums) + 1)), "编号：%s" % nums)


def check_heading_levels(doc):
    """GB/T 7713.2—2022 5.2.2：正文一般不宜超过 4 级；章节标题一般不宜超过 15 字。"""
    levels, too_long, deep = {}, [], []
    for para in doc.paragraphs:
        style = (para.style.name or "") if para.style is not None else ""
        if not style.startswith("Heading "):
            continue
        try:
            lv = int(style.split()[-1])
        except ValueError:
            continue
        text = para.text.strip()
        levels[lv] = levels.get(lv, 0) + 1
        if lv > 4:
            deep.append(text[:14])
        if len(text) > 15:
            too_long.append("%d 级「%s」%d 字" % (lv, text[:12], len(text)))
    if not levels:
        return
    check("版式", "标题层级 ≤4 级（5.2.2）", not deep, "；".join(deep[:3]))
    warn_if("版式", "章节标题 ≤15 字（5.2.2）", bool(too_long), "；".join(too_long[:3]))
    detail = "、".join("%d 级 %d 个" % (k, v) for k, v in sorted(levels.items()))
    check("版式", "标题层级分布：%s" % detail, True)


def check_toc(doc):
    instrs = [el.text or "" for el in doc.element.body.iter(qn("w:instrText"))]
    toc = [x for x in instrs if "TOC" in x]
    check("目录", "存在自动目录域（TOC）", bool(toc),
          ("域代码：%s" % toc[0].strip()) if toc else "未找到 TOC 域")
    if toc:
        import re as _re
        m = _re.search(r'\\o\s+"?(\d+)-(\d+)"?', toc[0])
        top, bottom = (int(m.group(1)), int(m.group(2))) if m else (1, 1)
        detail = '当前 \\o "%d-%d"' % (top, bottom)
        # 只列一级标题基本没有导航价值；GB/T 7713.1—2025 附录 E 的目次示例含"章 + 节"两级
        # 只列一级基本没有导航价值；判 WARN，允许小型报告显式用 1-1
        warn_if("目录", "目录含二级及以上条目（\\o \"1-N\"，N≥2）",
                not (top == 1 and bottom >= 2), detail)

    toc_items = [p for p in doc.paragraphs
                 if p.style is not None and (p.style.name or "").lower().startswith("toc")]
    topics = [p.text for p in doc.paragraphs if "目录域已插入" in p.text]
    check("目录", "目录已由 Word 展开（不是占位提示）",
          bool(toc_items) and not topics,
          "目录条目 %d 条%s" % (len(toc_items),
                              "；仍为占位文本，请更新域" if topics else ""),
          warn=bool(topics))


def check_references(doc):
    refs = []
    for p in doc.paragraphs:
        if re.match(r"^\[\d+\]", p.text.strip()):
            refs.append(p)
    if not refs:
        warn_if("参考文献", "未找到 [n] 形式的参考文献条目", True)
        return
    bad = []
    for p in refs:
        pPr = p._p.find(qn("w:pPr"))
        ind = pPr.find(qn("w:ind")) if pPr is not None else None
        hanging = a(ind, "hanging")
        left = a(ind, "left")
        if not hanging or int(hanging) < 300:
            bad.append("%s 缺少悬挂缩进" % p.text[:10])
        if not left:
            bad.append("%s 缺少左缩进" % p.text[:10])
        size = None
        for r in p.runs:
            if r.font.size is not None:
                size = r.font.size.pt
                break
        if size != EXPECT["reference_pt"]:
            bad.append("%s 字号 %s" % (p.text[:10], size))
    check("参考文献", "9 pt + 悬挂缩进 0.74 cm", not bad, "；".join(bad[:3]))
    nums = [int(re.match(r"^\[(\d+)\]", p.text.strip()).group(1)) for p in refs]
    check("参考文献", "编号连续", nums == list(range(1, len(nums) + 1)),
          "编号：%s" % nums)


def check_crossrefs(doc):
    text = "\n".join(p.text for p in doc.paragraphs)
    caps = {}
    for p in doc.paragraphs:
        m = re.match(r"^(图|表)\s*(\d+)", p.text.strip())
        if m:
            caps.setdefault(m.group(1), set()).add(int(m.group(2)))
    missing = []
    # 题注本身算 1 次命中，正文里应至少再引用 1 次（共 ≥2 次）
    for kind, nums in caps.items():
        for n in sorted(nums):
            hits = len(re.findall(r"%s\s*%d(?!\d)" % (kind, n), text))
            if hits < 2:
                missing.append("%s%d" % (kind, n))
    warn_if("交叉引用", "每张图/表都在正文中被引用", bool(missing),
            "未被引用：%s" % "、".join(missing))


def check_placeholders(doc):
    bad = []
    patterns = [r"XXX", r"TODO", r"Lorem", r"\[待填写\]", r"xxx",
                r"占位符", r"此处填写"]
    for p in doc.paragraphs:
        for pat in patterns:
            if re.search(pat, p.text):
                bad.append("%s → %r" % (pat, p.text.strip()[:20]))
    warn_if("内容", "无明显占位符残留（XXX / TODO / Lorem）", bool(bad),
            "；".join(bad[:3]))

    pending = sum(1 for p in doc.paragraphs if "待补充" in p.text or "待核实" in p.text)
    if pending:
        RESULTS.append(("WARN", "内容", "存在待补充 / 待核实标记",
                        "共 %d 处（规范要求如实标注，交付前请确认）" % pending))
        COUNTS["WARN"] += 1


def check_table_precision(doc):
    """同一列小数位数必须一致（GB/T 7713.2）；空缺允许用破折号。"""
    problems = []
    number = re.compile(r"^-?\d+(?:\.\d+)?$")
    for ti, tbl in enumerate(doc.tables, start=1):
        borders = tbl._tbl.find(qn("w:tblPr")).find(qn("w:tblBorders"))
        if borders is None:
            continue
        spec = {e.tag.split("}")[-1]: a(e, "val") for e in borders}
        if spec.get("top") != "single":
            continue        # 只看三线表（封面/公式表不算）
        grid = [[c.text.strip() for c in row.cells] for row in tbl.rows[1:]]
        if not grid:
            continue
        ncol = max(len(r) for r in grid)
        for col in range(ncol):
            decimals = set()
            for row in grid:
                if col >= len(row):
                    continue
                cell = row[col].replace("±", " ").split()[0] if row[col] else ""
                if cell in ("", "—", "-", "–"):
                    continue
                if not number.match(cell):
                    continue
                decimals.add(len(cell.split(".")[1]) if "." in cell else 0)
            if len(decimals) > 1:
                problems.append("表%d 第 %d 列小数位不一致：%s"
                                % (ti, col + 1, sorted(decimals)))
    check("表格", "同一列数字小数位一致（空缺用“—”）", not problems,
          "；".join(problems[:3]))


def check_palette(doc):
    """文档用色必须来自 palette.py：全部在批准色板内，且彩色 ≤5。"""
    import palette

    allowed = {palette.normalize(v) for v in palette.DOC.values()}
    allowed |= {palette.normalize(c) for c in palette.SERIES}
    allowed |= {palette.normalize(c) for c in palette.GREY_RAMP}
    allowed.add("#FFFFFF")
    if EXPECT.get("doc_colors"):        # --spec 换肤后的颜色同样视为批准
        allowed |= set(EXPECT["doc_colors"])
    unknown = {}
    used = set()
    for para in _all_paragraphs(doc):
        for run in para.runs:
            color = run.font.color
            if color is None or color.rgb is None:
                continue
            value = palette.normalize(str(color.rgb))
            used.add(value)
            if value not in allowed:
                unknown[value] = unknown.get(value, 0) + 1
    check("配色", "文档用色均来自 palette.py 的批准色板", not unknown,
          "未批准：" + "、".join("%s×%d" % kv for kv in list(unknown.items())[:4]))
    chromatic = palette.chromatic_count(sorted(used))
    check("配色", "彩色数量 ≤5（实测 %d 种）" % chromatic, chromatic <= 5)


def check_figure_quality(doc):
    """插图分辨率与版面宽度：线条图 ≥600 dpi（目标 1000），宽度 ≤160 mm。"""
    shapes = doc.inline_shapes
    if not shapes:
        warn_if("图片", "文档中没有图片（如不需要可忽略）", True)
        return
    low, missing = [], []
    for shape in shapes:
        width_mm = shape.width / 36000.0
        try:
            blob = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
            part = doc.part.related_parts[blob]
            import io

            from PIL import Image
            img = Image.open(io.BytesIO(part.blob))
            dpi = img.info.get("dpi")
            if not dpi:
                missing.append("%.1f mm 无 dpi 信息" % width_mm)
            elif float(dpi[0]) < 600:
                low.append("%.0f dpi" % float(dpi[0]))
        except Exception as exc:  # pragma: no cover
            missing.append(str(exc))
    check("图片", "插图分辨率 ≥600 dpi（线条图目标 1000）", not low, "；".join(low[:3]))
    warn_if("图片", "插图带 dpi 元数据（便于期刊核对）", bool(missing),
            "；".join(missing[:3]))



def check_standard_figures(doc):
    """GB/T 7713.2—2022 5.4：插图尺寸、字号、图注、先文后图。"""
    import palette

    sec = doc.sections[0]
    text_w_cm = (sec.page_width - sec.left_margin - sec.right_margin) / 360000.0
    usable_h_cm = (sec.page_height - sec.top_margin - sec.bottom_margin) / 360000.0

    shapes = list(doc.inline_shapes)
    if not shapes:
        return
    too_wide, too_tall, tiny = [], [], []
    for idx, shape in enumerate(shapes, start=1):
        w_cm = shape.width / 360000.0
        h_cm = shape.height / 360000.0
        if w_cm > text_w_cm + 0.05:
            too_wide.append("图%d %.2f cm > 版心 %.2f cm" % (idx, w_cm, text_w_cm))
        if h_cm > usable_h_cm - 3.0:      # 留出图题 + 图注 + 上下间距
            too_tall.append("图%d 高 %.1f cm" % (idx, h_cm))
        try:
            blob = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
            part = doc.part.related_parts[blob]
            import io as _io

            from PIL import Image
            img = Image.open(_io.BytesIO(part.blob))
            info = dict(img.info)
            nat_w = img.width
            software = info.get("Software") or ""
            design_w = design_pt = None
            for token in software.split():
                if token.startswith("width_mm="):
                    design_w = float(token.split("=")[1])
                elif token.startswith("font_pt="):
                    design_pt = float(token.split("=")[1])
            if design_w and design_pt and nat_w:
                effective = design_pt * (w_cm * 10.0 / design_w)
                if effective < palette.MIN_FIGURE_FONT_PT:
                    tiny.append("图%d 约 %.1f pt" % (idx, effective))
        except ImportError:
            pass                      # 缺 Pillow：跳过字号核算

    check("图片", "插图宽度 ≤ 版心宽（%.1f cm，GB/T 7713.2—2022 5.4.1）" % text_w_cm,
          not too_wide, "；".join(too_wide[:3]))
    check("图片", "插图高度不超一页可用高（%.1f cm）" % (usable_h_cm - 3.0),
          not too_tall, "；".join(too_tall[:3]))
    warn_if("图片", "图内文字插入后 ≥%.1f pt（图中文字清晰可见）"
            % palette.MIN_FIGURE_FONT_PT, bool(tiny), "；".join(tiny[:3]))


def check_standard_tables(doc):
    """GB/T 7713.2—2022 5.5.4：表头不得有斜线、不得用“同上/同左”代替重复。"""
    diagonal, ditto = [], []
    for ti, tbl in enumerate(doc.tables, start=1):
        for row in tbl.rows:
            for cell in row.cells:
                tcPr = cell._tc.find(qn("w:tcPr"))
                if tcPr is None:
                    continue
                borders = tcPr.find(qn("w:tcBorders"))
                if borders is not None:
                    for edge in ("tl2br", "tr2bl"):
                        node = borders.find(qn("w:" + edge))
                        if node is not None and a(node, "val") not in (None, "nil", "none"):
                            diagonal.append("表%d 表头有斜线" % ti)
                text = cell.text.strip()
                if text in ("同上", "同左", "同右", "〃", "‰"):
                    ditto.append("表%d 单元格用了“%s”" % (ti, text))
    check("表格", "表头不准许使用斜线（5.5.4）", not diagonal, "；".join(diagonal[:3]))
    check("表格", "内容相同的栏应重复写出，不用“同上/同左”（5.5.4）",
          not ditto, "；".join(ditto[:3]))


def check_note_period(doc):
    """5.4.3 / 5.5.3：图注、表注的末尾应加“。”。"""
    bad = []
    for p in _all_paragraphs(doc):
        text = p.text.strip()
        if not text.startswith("注："):
            continue
        if not text.endswith(("。", ".")):
            bad.append(text[:18] + "…")
    check("图表", "图注 / 表注末尾有句号（5.4.3、5.5.3）", not bad,
          "；".join(bad[:3]))


def check_caption_before_content(doc):
    """5.4.5 / 5.5.5：插图、表格宜紧置首次提及处之后——先见文字后见图表。

    题注行用 `_caption_kind`（居中 + ≤10 pt）识别，不用"以图号开头"判断：
    正文里"图6给出……""表1 给出一个最小示例"这类句子同样以编号开头，
    按文字判断会把它们误当题注，反而漏掉真正的提及。
    """
    from docx.text.paragraph import Paragraph

    body = doc.element.body
    nodes = list(body.iterchildren())

    def para_at(node):
        return Paragraph(node, doc._body)

    def text_at(node):
        if node.tag.split("}")[-1] != "p":
            return ""
        return "".join(t.text or "" for t in node.findall(".//" + qn("w:t"))).strip()

    # 1) 图表本体 + 配对题注
    items = []
    caption_idx = set()
    for idx, node in enumerate(nodes):
        tag = node.tag.split("}")[-1]
        if tag == "p" and node.findall(".//" + qn("w:drawing")):
            for j in range(idx + 1, min(idx + 3, len(nodes))):
                kind, info = _caption_kind(para_at(nodes[j]))
                if kind == "图":
                    items.append(("图", str(info[0]), idx))
                    caption_idx.add(j)
                    break
        elif tag == "tbl":
            for j in range(idx - 1, max(idx - 4, -1), -1):
                kind, info = _caption_kind(para_at(nodes[j]))
                if kind == "表":
                    items.append(("表", str(info[0]), idx))
                    caption_idx.add(j)
                    break

    # 2) 正文提及（排除题注行）
    mention = {}
    for idx, node in enumerate(nodes):
        if idx in caption_idx:
            continue
        text = text_at(node)
        if not text:
            continue
        for kind, num in re.findall(r"(图|表)\s*(\d+)", text):
            mention.setdefault(kind + num, idx)

    problems = []
    for kind, number, pos in items:
        key = kind + number
        first = mention.get(key)
        if first is None:
            problems.append("%s 未在正文提及" % key)
        elif first >= pos:
            problems.append("%s 出现在首次提及之前" % key)
    # 标准用"宜"（推荐性条款），故判 WARN；加 --strict 时会被提升为 FAIL
    warn_if("图表", "先见文字后见图表（5.4.5、5.5.5）", bool(problems),
            "；".join(problems[:3]))


def check_front_matter_order(doc):
    """GB/T 7713.1—2025 附录 A 表 A.1：摘要页在目次页之前。

    前置部分顺序为：封面 → 题名页 → 原创性声明 → 使用授权 → 致谢 →
    **摘要页 → 序言/前言 → 目次页** → 插图和附表清单 …
    少数学校模板要求目录在前，故判 WARN 而非 FAIL。
    """
    def norm(text):
        return text.replace("\u3000", "").replace(" ", "").strip()

    abs_idx = toc_idx = None
    for i, para in enumerate(_all_paragraphs(doc)):
        key = norm(para.text)
        if abs_idx is None and key in ("摘要", "中文摘要", "摘要页"):
            abs_idx = i
        elif toc_idx is None and key in ("目录", "目次", "目次页"):
            toc_idx = i
    if toc_idx is None or abs_idx is None:
        return
    wrong = toc_idx < abs_idx
    msg = "目录出现在摘要之前" if wrong else ""
    warn_if("版式", "前置部分顺序：摘要页在目次页之前（GB/T 7713.1—2025 表 A.1）",
            wrong, msg + "（与国标相反；若学校模板确有此要求可忽略）" if wrong else "")


def check_para_shading(doc):
    """论文正文不惯用彩色底框；出现段底纹时提示（模板默认已不加框）。"""
    shaded = []
    for para in _all_paragraphs(doc):
        pPr = para._p.find(qn("w:pPr"))
        if pPr is None:
            continue
        shd = pPr.find(qn("w:shd"))
        if shd is None:
            continue
        fill = a(shd, "fill")
        if fill and fill not in ("auto", "FFFFFF"):
            shaded.append((para.text or "").strip()[:16] or "（空段）")
    warn_if("版式", "正文段落无彩色底纹（论文不惯用提示框）", bool(shaded),
            "；".join(shaded[:3]))


_CITE_RE = re.compile(r"[\[［]\s*(\d+(?:\s*[,，\-–]\s*\d+)*)\s*[\]］]")
_REF_ENTRY = re.compile(r"^\s*[\[［]\s*\d+\s*[\]］]")
_REF_TYPE = re.compile(r"\[(?:J|M|C|D|S|P|R|N|G|Z|A|EB/OL|DB/OL|DB/CD|CP/DK|M/OL|J/OL)\]")


def check_citations(doc):
    """GB/T 7714—2015 顺序编码制：正文标注用上标方括号数字，编号与文后条目对应。"""
    refs = [p for p in _all_paragraphs(doc) if _REF_ENTRY.match(p.text or "")]
    ref_nums = {}
    for para in refs:
        m = re.match(r"^\s*[\[［]\s*(\d+)\s*[\]］]", para.text)
        if m:
            ref_nums[int(m.group(1))] = para.text
    n_refs = len(refs)

    cited: set[int] = set()
    not_superscript, out_of_range = [], []
    for para in _all_paragraphs(doc):
        if _REF_ENTRY.match(para.text or ""):
            continue                     # 参考文献条目本身
        sups = [r.text.strip() for r in para.runs if r.font.superscript]
        for m in _CITE_RE.finditer(para.text or ""):
            token = m.group(0)
            nums = [int(x) for x in re.findall(r"\d+", m.group(1))]
            in_sup = any(token.strip(" ") in s or m.group(1) in s for s in sups)
            if in_sup:
                cited.update(nums)
                if n_refs and max(nums) > n_refs:
                    out_of_range.append("%s（共 %d 条）" % (token, n_refs))
            else:
                not_superscript.append(token)

    check("参考文献", "正文引文标注为上标方括号（GB/T 7714—2015 顺序编码制）",
          not not_superscript, "未上标：%s" % "、".join(dict.fromkeys(not_superscript))[:40])
    check("参考文献", "引文编号不超出文后条目数", not out_of_range,
          "；".join(out_of_range[:3]))
    if ref_nums:
        missing = sorted(set(ref_nums) - cited)
        warn_if("参考文献", "每条文献都在正文被引用（顺序编码制应按引用顺序编号）",
                bool(missing), "未引用：%s" % "、".join(str(n) for n in missing[:6]))
        no_type = [str(n) for n, text in ref_nums.items() if not _REF_TYPE.search(text)]
        warn_if("参考文献", "条目含文献类型标识（[J]/[M]/[C]/[EB/OL] 等）",
                bool(no_type), "缺标识：%s" % "、".join(no_type[:6]))


def check_ai_tone(path):
    """AI 味只做提示：分数是启发式的，不判断作者身份，也绝不判 FAIL。"""
    try:
        import ai_tone_check
    except Exception as exc:  # pragma: no cover - 模块缺失时不阻断质检
        warn_if("文风", "AI 味检测跳过", True, str(exc))
        return
    try:
        paragraphs, name = ai_tone_check.collect_paragraphs_docx(path)
        report = ai_tone_check.analyze(paragraphs, name)
    except Exception as exc:  # pragma: no cover
        warn_if("文风", "AI 味检测跳过", True, str(exc))
        return

    score = report.score
    top = []
    for hit in report.hits:
        if hit.rule_id not in top:
            top.append(hit.rule_id)
        if len(top) >= 3:
            break
    reg = report.register_score
    detail = "AI 味 %d/100（%s）｜ 语体 %d/100（%s）%s；详见 references/anti-ai-academic-zh.md" % (
        score, ai_tone_check.grade(score),
        reg, ai_tone_check.grade_register(reg),
        "，主要命中：%s" % "/".join(top) if top else "")
    check("文风", "AI 味检测（提示项，不判失败）", score < 40, detail, warn=True)


def check_runs_fonts(doc):
    bad = []
    for p in doc.paragraphs:
        style_name = (p.style.name if p.style is not None else "") or ""
        for r in p.runs:
            rPr = r._r.find(qn("w:rPr"))
            rFonts = rPr.find(qn("w:rFonts")) if rPr is not None else None
            if rFonts is None:
                continue
            cn = a(rFonts, "eastAsia")
            if cn and cn not in ("宋体", "黑体"):
                bad.append("%s 使用了中文字体 %s" % (style_name or "正文", cn))
            en = a(rFonts, "ascii")
            if en and en not in ("Times New Roman", "Arial"):
                bad.append("%s 使用了西文字体 %s" % (style_name or "正文", en))
    check("字体", "只使用 宋体/黑体 + Times New Roman/Arial", not bad,
          "；".join(sorted(set(bad))[:3]))


def check_blank_pages(doc):
    runs = 0
    worst = 0
    for p in doc.paragraphs:
        if not p.text.strip() and p._p.findall(".//" + qn("w:drawing")) == []:
            runs += 1
            worst = max(worst, runs)
        else:
            runs = 0
    warn_if("版式", "无连续空段落（避免空白页）", worst >= 3,
            "最多连续 %d 个空段落" % worst)


# --------------------------------------------------------------- 主流程

GROUPS = ["页面", "视图", "样式", "正文", "列表", "表格", "图表", "图片",
          "公式", "目录", "页眉页脚", "参考文献", "交叉引用", "字体", "版式",
          "内容", "文风"]


def main(argv=None):
    _console.setup()
    ap = argparse.ArgumentParser(description="按规范质检 docx")
    ap.add_argument("docx")
    ap.add_argument("--pdf", help="同时体检导出的 PDF")
    ap.add_argument("--spec", help="与 build_docx.py 同一份 spec 覆盖（整体换配色/字号后用它校准期望值）")
    ap.add_argument("--strict", action="store_true", help="警告也视为失败")
    args = ap.parse_args(argv)

    if args.spec:
        apply_spec(args.spec)

    path = os.path.abspath(args.docx)
    if not os.path.exists(path):
        print("找不到文件：%s" % path)
        return 2

    # 允许在同一进程里多次调用（测试、批处理），每次清零
    RESULTS.clear()
    for key in COUNTS:
        COUNTS[key] = 0

    doc = Document(path)
    check_page(doc)
    check_settings(doc)
    check_styles(doc)
    check_body_paragraphs(doc)
    check_lists(doc)
    check_headers_footers(doc)
    check_tables_and_captions(doc)
    check_figures(doc)
    check_equations(doc)
    check_toc(doc)
    check_heading_levels(doc)
    check_references(doc)
    check_crossrefs(doc)
    check_runs_fonts(doc)
    check_blank_pages(doc)
    check_placeholders(doc)
    check_ai_tone(path)
    check_standard_figures(doc)
    check_standard_tables(doc)
    check_note_period(doc)
    check_caption_before_content(doc)
    check_front_matter_order(doc)
    check_para_shading(doc)
    check_citations(doc)
    check_table_precision(doc)
    check_palette(doc)
    check_figure_quality(doc)

    if args.pdf:
        try:
            try:
                import pymupdf as fitz
            except ImportError:  # pragma: no cover
                import fitz
            pdf = fitz.open(os.path.abspath(args.pdf))
            text = "\n".join(pdf.load_page(i).get_text() for i in range(pdf.page_count))
            pdf.close()
            bad = [b for b in ("错误！未定义书签", "Error! Bookmark not defined",
                               "TOC \\o") if b in text]
            check("PDF", "无未定义书签 / 无原始域代码", not bad, "；".join(bad))
        except Exception as exc:
            warn_if("PDF", "PDF 体检跳过", True, str(exc))

    print("\n%s  质检报告" % os.path.basename(path))
    print("=" * 78)
    cur = None
    for level, group, name, detail in RESULTS:
        if group != cur:
            print("\n[%s]" % group)
            cur = group
        mark = {"PASS": "  OK  ", "WARN": " WARN ", "FAIL": " FAIL "}[level]
        print("  %s %s" % (mark, name))
        if detail and level != "PASS":
            print("         %s" % detail)
    print("\n" + "=" * 78)
    print("通过 %d 项，警告 %d 项，失败 %d 项"
          % (COUNTS["PASS"], COUNTS["WARN"], COUNTS["FAIL"]))

    failed = COUNTS["FAIL"] + (COUNTS["WARN"] if args.strict else 0)
    if failed:
        print("\n结论：不通过。请按上面的 FAIL%s 项修复后重跑。"
              % ("/WARN" if args.strict else ""))
        return 1
    print("\n结论：符合规范。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
