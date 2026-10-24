"""pytest 公共配置与共享 fixture。"""

from __future__ import annotations

import json
import os
import sys

import pytest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(REPO, "scripts")
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)

MINIMAL_MODEL = os.path.join(REPO, "examples", "minimal", "content.json")
DEMO_MODEL = os.path.join(REPO, "examples", "demo", "content_model.json")
DEMO_FIGURE = os.path.join(REPO, "examples", "demo", "figs", "fig1_duty.png")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


def w(tag: str) -> str:
    return "{%s}%s" % (W_NS, tag)


def m(tag: str) -> str:
    return "{%s}%s" % (M_NS, tag)


def localname(el) -> str:
    return el.tag.split("}")[-1]


def load_model(path: str = MINIMAL_MODEL) -> dict:
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def demo_model() -> dict:
    return load_model(DEMO_MODEL)


@pytest.fixture(scope="session")
def repo() -> str:
    return REPO


@pytest.fixture(scope="session")
def scripts_dir() -> str:
    return SCRIPTS


@pytest.fixture()
def basic_model() -> dict:
    """一份小而易读的内容模型：封面 + 摘要 + 目录 + 正文 + 表 + 公式 + 参考文献。"""
    return {
        "meta": {
            "org": "测试单位",
            "doc_type": "测试报告",
            "title": "测试标题",
            "subtitle": "测试副标题",
            "note": "测试说明行",
            "info": [["编号", "T-001"], ["作者", "测试者"]],
            "college": "测试学院",
            "date": "2026 年 3 月",
            "header_type": "测试报告",
        },
        "abstract": {
            "heading": "摘　要",
            "paragraphs": ["这是一段用于测试的中文摘要。" * 4],
            "keywords": ["关键词甲", "关键词乙"],
        },
        "toc": True,
        "sections": [
            {"level": 1, "text": "1 引言"},
            {"type": "p", "text": "这是第一段正文内容，用于检查首行缩进与行距。" * 3},
            {"type": "list", "items": ["要点一", "要点二"]},
            {
                "type": "table",
                "caption": "示例表",
                "headers": ["列A", "列B"],
                "rows": [["1", "2"], ["3", "4"]],
                "widths": [7.0, 8.5],
                "note": "注：示例数据。",
            },
            {"level": 2, "text": "1.1 小节"},
            {"type": "equation", "latex": r"y = \frac{a}{b}", "number": "（1）"},
            {"type": "p", "text": "式（1）用于测试 OMML 公式渲染。"},
        ],
        "references": ["作者. 题名[J]. 刊名, 2020, 1(2): 3-4."],
    }


@pytest.fixture()
def built(basic_model, tmp_path):
    """用内容模型生成 docx，返回 (builder, 路径)。"""
    import docx_builder

    out = tmp_path / "out" / "report.docx"
    builder, path = docx_builder.build_from_model(basic_model, str(out))
    return builder, path


@pytest.fixture()
def built_doc(built):
    from docx import Document

    return Document(built[1])
