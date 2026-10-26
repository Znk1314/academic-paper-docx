#!/usr/bin/env python
"""把 PDF 逐页渲染成 PNG，并顺手做几项"肉眼看才知道"的自动检查。

用法：
    python render_preview.py out/report.pdf -o preview --dpi 110

检查项：
    * 空白页（无文字、无图）
    * "错误！未定义书签" / "Error! Bookmark not defined"
    * 目录仍是原始域代码（出现 "TOC \\o" 字样）
    * 页面是否 A4（按页面尺寸判断）
"""

from __future__ import annotations

import argparse
import os
import sys

import _console

try:
    import pymupdf as fitz  # PyMuPDF
except ImportError:  # pragma: no cover
    try:
        import fitz
    except ImportError:  # 依赖缺失时给一句话提示，而不是抛栈
        print("缺少 PyMuPDF，无法做 PDF 预览与体检。")
        print("安装：python -m pip install PyMuPDF（不影响生成 docx）")
        raise SystemExit(2) from None

BAD_TEXT = [
    "错误！未定义书签", "错误!未定义书签", "Error! Bookmark not defined",
    "错误！找不到引用源", "Error! Reference source not found",
]
RAW_FIELD_MARKS = ["TOC \\o", "TOC\\o", "PAGEREF", "MERGEFORMAT", "HYPERLINK \\l"]


def main(argv=None):
    _console.setup()
    ap = argparse.ArgumentParser(description="PDF 逐页预览 + 基础体检")
    ap.add_argument("pdf")
    ap.add_argument("-o", "--outdir", default=None)
    ap.add_argument("--dpi", type=int, default=110)
    ap.add_argument("--pages", default=None, help="例如 1-12，默认全部")
    args = ap.parse_args(argv)

    pdf = os.path.abspath(args.pdf)
    outdir = args.outdir or os.path.join(os.path.dirname(pdf), "preview")
    os.makedirs(outdir, exist_ok=True)

    doc = fitz.open(pdf)
    total = doc.page_count
    first, last = 1, total
    if args.pages:
        a, _, b = args.pages.partition("-")
        first = int(a)
        last = int(b or a)

    print("PDF：%s（共 %d 页）" % (pdf, total))
    problems = []
    for idx in range(total):
        page = doc.load_page(idx)
        no = idx + 1
        if first <= no <= last:
            pix = page.get_pixmap(dpi=args.dpi)
            path = os.path.join(outdir, "page-%02d.png" % no)
            pix.save(path)

        text = page.get_text().strip()
        images = page.get_images(full=True)
        rect = page.rect
        if not text and not images:
            problems.append("第 %d 页是空白页" % no)
        for bad in BAD_TEXT:
            if bad in page.get_text():
                problems.append("第 %d 页出现“%s”" % (no, bad))
        for mark in RAW_FIELD_MARKS:
            if mark in page.get_text():
                problems.append("第 %d 页出现原始域代码“%s”" % (no, mark))
        w_cm = rect.width / 72 * 2.54
        h_cm = rect.height / 72 * 2.54
        if abs(w_cm - 21.0) > 0.2 or abs(h_cm - 29.7) > 0.3:
            problems.append("第 %d 页不是 A4（%.1f × %.1f cm）" % (no, w_cm, h_cm))

    doc.close()

    print("预览图：%s" % outdir)
    if problems:
        print("\n需要人工确认：")
        for p in problems:
            print("  - " + p)
        return 1
    print("未发现空白页 / 未定义书签 / 原始域代码 / 非 A4 页面")
    return 0


if __name__ == "__main__":
    sys.exit(main())
