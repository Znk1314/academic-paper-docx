#!/usr/bin/env python
"""自动排版：消掉正文里的大片留白（结尾页除外）。

为什么需要它：Word 的"图与图题同页"（keep_with_next）会让一整块图+题被推到下一页，
上一页底部就空出半页；表格不让跨页时同理。这类留白在期刊稿和学位论文里都会被编辑
指出。

做法（迭代收敛，不猜）：
    1. 构建 docx → 用 Word/LibreOffice 导出 PDF
    2. 用 PyMuPDF 量每一页的"内容底部到版心底部"的空白
    3. 对空白超阈值且**不是最后一页**的页，找出被推到下一页的图（按图题文本定位），
       把该图宽度缩到能塞进上一页剩余空间的比例（有下限），重构建
    4. 最多 MAX_PASSES 轮；仍不达标就如实报告，不硬凑

只调"图宽"与"段落间距"，不改文字、不改字号、不删内容。
"""

from __future__ import annotations

import copy
import os
import re

CM_TO_PT = 72.0 / 2.54        # 1 cm = 28.35 pt
PT_TO_CM = 2.54 / 72.0        # 1 pt = 0.0353 cm

MAX_PASSES = 3
MIN_FIGURE_CM = 6.0
DEFAULT_MAX_BLANK_CM = 5.0
SHRINK_STEP = 0.92


def measure_pdf(pdf_path, bottom_margin_cm=2.0, header_zone_cm=3.0,
                footer_zone_cm=1.6):
    """量每页的填充情况。

    返回 [{"page": n, "blank_cm": x, "fill": 0-1}]，blank_cm 是内容底部到
    版心底部的距离（含页脚区）。最后一页不参与判定（结尾允许留白）。
    """
    try:
        import pymupdf as fitz
    except ImportError:  # pragma: no cover
        try:
            import fitz
        except ImportError:
            return None

    doc = fitz.open(pdf_path)
    pages = []
    for index in range(doc.page_count):
        page = doc.load_page(index)
        rect = page.rect
        height_mm = rect.height * PT_TO_CM * 10
        top_limit = rect.y0 + header_zone_cm * CM_TO_PT
        bottom_limit = rect.y1 - bottom_margin_cm * CM_TO_PT
        # 页脚（页码）不算正文内容，否则每页都会被判成"填满"
        footer_top = rect.y1 - footer_zone_cm * CM_TO_PT

        lowest = top_limit
        for block in page.get_text("blocks"):
            y1 = block[3]
            if y1 <= footer_top and y1 > lowest:
                lowest = y1
        for image in page.get_images(full=True):
            try:
                bbox = page.get_image_bbox(image)
            except Exception:                    # pragma: no cover
                continue
            if bbox.y1 <= footer_top and bbox.y1 > lowest:
                lowest = bbox.y1

        blank_pt = max(0.0, bottom_limit - lowest)
        span = max(1.0, bottom_limit - top_limit)
        pages.append({
            "page": index + 1,
            "blank_cm": round(blank_pt * PT_TO_CM, 2),
            "fill": round(min(1.0, max(0.0, (lowest - top_limit) / span)), 2),
            "height_mm": round(height_mm, 1),
        })
    doc.close()
    return pages


def page_texts(pdf_path):
    """每页文本，用于把"被推走的图题"定位到页。"""
    try:
        import pymupdf as fitz
    except ImportError:  # pragma: no cover
        try:
            import fitz
        except ImportError:
            return []
    doc = fitz.open(pdf_path)
    texts = [doc.load_page(i).get_text() for i in range(doc.page_count)]
    doc.close()
    return texts


BODY_HEADING_RE = re.compile(r"^\s*[0-9１-９]+\s*[　 ]+\S", re.M)

FRONT_MATTER_MARKS = ("摘　要", "摘要", "Abstract", "目　录", "目录",
                      "参考文献", "致　谢", "致谢", "……", "....")


def body_start_page(texts, headings=None):
    """正文起始页（1-based）。前置部分各页按规范本就另起页，允许留白。

    优先用一级标题文本定位；定位不到时跳过含摘要/目录字样的页。
    """
    headings = [h for h in (headings or []) if h and len(h) >= 2]
    front_marks = [m.replace(" ", "") for m in FRONT_MATTER_MARKS]
    for idx, text in enumerate(texts, start=1):
        flat = text.replace(" ", "")
        # 目录页会列出全部标题，必须先按前置部分标记排除
        if any(mark in flat for mark in front_marks):
            continue
        if headings and any(h.replace(" ", "") in flat for h in headings):
            return idx
        if not headings and BODY_HEADING_RE.search(text):
            return idx
    return 1


def violations(pages, max_blank_cm=DEFAULT_MAX_BLANK_CM, body_start=1):
    """正文页（非末页）里留白超阈值的页；前置部分与末页不算。"""
    if not pages:
        return []
    body = [p for p in pages[:-1] if p["page"] >= body_start]
    return [p for p in body if p["blank_cm"] > max_blank_cm]


def _iter_figures(model):
    """遍历内容模型里所有图块，返回 [(item, ref)]；ref 用于写回宽度。"""
    for item in model.get("sections", []) or []:
        if isinstance(item, dict) and item.get("type") == "figure":
            yield item, item
        if isinstance(item, dict) and item.get("type") == "figure_grid":
            for sub in item.get("figures", []) or []:
                yield sub, sub


def _caption_of(item):
    return str(item.get("caption") or item.get("caption_zh") or "").strip()


def auto_fit(model, build_docx, to_pdf, out_path, work_dir=None,
             max_blank_cm=DEFAULT_MAX_BLANK_CM, passes=MAX_PASSES,
             log=print):
    """反复构建→量页→缩图，直到正文页没有大片留白。

    build_docx(model, path) 负责生成 docx；to_pdf(docx_path) 返回 pdf 路径或 None。
    返回 (final_model, report)；report 里记录每轮的违规页与调整过的图。
    改了字/字号/内容以外的任何东西都要在 report 里说清楚。
    """
    current = copy.deepcopy(model)
    report = {"passes": [], "adjusted": [], "ok": False,
              "max_blank_cm": max_blank_cm}
    best = {"score": None, "model": None, "pdf": None}

    for pass_no in range(1, passes + 1):
        build_docx(current, out_path)
        pdf_path = to_pdf(out_path)
        if not pdf_path or not os.path.exists(pdf_path):
            report["passes"].append({"pass": pass_no, "note": "无法导出 PDF，跳过自动排版"})
            return current, report

        pages = measure_pdf(pdf_path)
        if pages is None:
            report["passes"].append({"pass": pass_no, "note": "缺 PyMuPDF，跳过"})
            return current, report

        texts = page_texts(pdf_path)
        start = body_start_page(texts, model.get("_level1_headings"))
        bad = violations(pages, max_blank_cm, body_start=start)
        report["passes"].append({
            "pass": pass_no,
            "pages": len(pages),
            "body_start": start,
            "violations": [(p["page"], p["blank_cm"]) for p in bad],
        })
        score = sum(p["blank_cm"] for p in bad)
        if best["score"] is None or score < best["score"]:
            best = {"score": score, "model": copy.deepcopy(current),
                    "pdf": pdf_path}
        if not bad:
            report["ok"] = True
            return current, report
        if pass_no > 1 and score >= report.get("_last_score", 1e9):
            # 上一轮缩图没有改善（甚至更糟）：回滚到最好的一轮，停止折腾
            report["passes"].append({"pass": pass_no,
                                     "note": "无改善，回滚到第 %d 轮的结果"
                                             % (pass_no - 1)})
            if best["model"] is not None:
                return best["model"], report
            return current, report
        report["_last_score"] = score
        if pass_no == passes:
            break

        # 把被推到下一页的图缩一点：下一页码里有该图题，就缩这个图
        adjusted_this_pass = 0
        for bad_page in bad:
            next_no = bad_page["page"] + 1
            next_text = texts[next_no - 1] if next_no - 1 < len(texts) else ""
            shrunk = False
            for item, ref in _iter_figures(current):
                cap = _caption_of(item)
                if shrunk or not cap or cap not in next_text:
                    continue
                width = float(ref.get("width_cm") or 14.0)
                new_width = max(MIN_FIGURE_CM, round(width * SHRINK_STEP, 2))
                if new_width >= width:
                    continue
                ref["width_cm"] = new_width
                report["adjusted"].append({
                    "caption": cap[:24], "from_cm": width, "to_cm": new_width,
                    "for_page": bad_page["page"],
                })
                adjusted_this_pass += 1
                shrunk = True
            # 缩图没命中时，收一下该页实际段落间距（全局小步）
            if not adjusted_this_pass:
                for key in ("space_after_pt", "para_space_after"):
                    if isinstance(current.get(key), (int, float)):
                        current[key] = max(0, round(current[key] * 0.7, 1))
                current["_compact"] = int(current.get("_compact", 0)) + 1
                report["adjusted"].append({"note": "收紧段后间距", "for_page": bad_page["page"]})

        log("自动排版 %d/%d：留白页 %s → 调整 %d 处"
            % (pass_no, passes,
               "、".join("第%d页 %.1fcm" % (p["page"], p["blank_cm"]) for p in bad),
               adjusted_this_pass))
        if not adjusted_this_pass and not bad:
            break

    report["ok"] = False
    return current, report


def format_report(report):
    lines = []
    for entry in report.get("passes", []):
        if entry.get("note"):
            lines.append("  第%d轮：%s" % (entry["pass"], entry["note"]))
            continue
        bad = entry.get("violations") or []
        lines.append("  第%d轮：%d 页，留白超标 %d 页%s"
                     % (entry["pass"], entry.get("pages", 0), len(bad),
                        ("（" + "、".join("%d 页 %.1f cm" % (p, b) for p, b in bad) + "）")
                        if bad else ""))
    for item in report.get("adjusted", []):
        if item.get("note"):
            lines.append("  · %s" % item["note"])
        else:
            lines.append("  · 缩图「%s」%.1f→%.1f cm（为第 %d 页让位）"
                         % (item["caption"], item["from_cm"], item["to_cm"],
                            item["for_page"]))
    lines.append("  结论：%s" % ("正文页无大片留白" if report.get("ok")
                                else "仍有留白，请人工调版"))
    return "\n".join(lines)
