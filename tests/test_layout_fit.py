"""自动排版测试：页填充测量、正文起点、违规判定、回滚报告。"""

from __future__ import annotations

import os
import sys

import pytest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))), "scripts"))

import layout_fit as L  # noqa: E402


def _pages(*blanks):
    return [{"page": i + 1, "blank_cm": b, "fill": round(1 - b / 24, 2),
             "height_mm": 297.0} for i, b in enumerate(blanks)]


def test_violations_ignores_last_page():
    pages = _pages(1.0, 9.0, 12.0)          # 第 3 页是末页，允许留白
    bad = L.violations(pages, 5.0)
    assert [p["page"] for p in bad] == [2]


def test_violations_skips_front_matter():
    pages = _pages(12.0, 12.0, 9.0, 1.0, 1.0)
    bad = L.violations(pages, 5.0, body_start=3)
    assert [p["page"] for p in bad] == [3]  # 前置两页不算，正文第 3 页算


def test_body_start_prefers_headings_but_skips_toc():
    texts = [
        "本科毕业设计（论文）",                       # 封面
        "Abstract  Indoor parameters ...",          # 英文摘要
        "目 录\n1 引言 .......... 1",                  # 目录（含标题，必须跳过）
        "1 引言\n实验室环境 ...",                      # 正文起点
    ]
    assert L.body_start_page(texts, ["1 引言"]) == 4


def test_body_start_falls_back_to_numbered_heading():
    texts = ["封面", "摘　要\n本文...", "1　引　言\n正文开始"]
    assert L.body_start_page(texts) == 3


def test_body_start_defaults_to_first_page_without_marks():
    assert L.body_start_page(["封面文字", "Abstract text"]) == 1


def test_format_report_states_conclusion():
    report = {"passes": [{"pass": 1, "pages": 11, "body_start": 5,
                          "violations": [(8, 9.5)]}],
              "adjusted": [{"caption": "某图", "from_cm": 12.0, "to_cm": 11.0,
                            "for_page": 8}],
              "ok": False}
    text = L.format_report(report)
    assert "第1轮" in text and "缩图" in text and "人工调版" in text


def test_measure_real_pdf(tmp_path):
    """用 matplotlib 生成两页 PDF，确认能测出页数与空白（缺 PyMuPDF 时跳过）。"""
    pytest.importorskip("pymupdf")
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages

    pdf = tmp_path / "two.pdf"
    with PdfPages(pdf) as book:
        for label in ("page one", "page two"):
            fig, ax = plt.subplots(figsize=(8.27, 11.69))
            ax.text(0.05, 0.95, label, va="top")
            ax.set_axis_off()
            book.savefig(fig)
            plt.close(fig)

    pages = L.measure_pdf(str(pdf))
    assert pages and len(pages) == 2
    assert all(0 <= p["blank_cm"] <= 40 for p in pages)
    assert pages[0]["height_mm"] == pytest.approx(297, abs=3)
