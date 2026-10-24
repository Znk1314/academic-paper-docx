"""qa_check.py 的自检：合格文档要过，故意破坏的文档要 FAIL。"""

from __future__ import annotations

import json

import docx_builder as db
import pytest
import qa_check
from conftest import w
from docx import Document


def run_qa(path, capsys, extra=None):
    argv = [str(path)] + (extra or [])
    code = qa_check.main(argv)
    return code, capsys.readouterr().out


def test_clean_document_passes(built, capsys):
    code, out = run_qa(built[1], capsys)
    # 未跑 Word 时"目录未展开"只是 WARN，不影响通过
    assert code == 0, out
    assert "失败 0 项" in out


def test_web_view_is_detected(built, capsys):
    doc = Document(built[1])
    db._set_setting(doc, "view", {"val": "web"})
    doc.save(built[1])
    code, out = run_qa(built[1], capsys)
    assert code == 1
    assert "默认页面视图" in out


def test_missing_title_pg_is_detected(built, capsys):
    doc = Document(built[1])
    sectPr = doc.sections[0]._sectPr
    titlePg = sectPr.find(w("titlePg"))
    if titlePg is not None:
        sectPr.remove(titlePg)
    doc.save(built[1])
    code, out = run_qa(built[1], capsys)
    assert code == 1
    assert "首页不同" in out


def test_bullet_symbol_is_detected(basic_model, tmp_path, capsys):
    basic_model["sections"] = [
        {"level": 1, "text": "1 引言"},
        {"type": "p", "text": "这里用了圆点：\u2022 第一项"},
    ]
    _, path = db.build_from_model(basic_model, str(tmp_path / "b.docx"))
    code, out = run_qa(path, capsys)
    assert code == 1
    assert "圆点符号" in out


def test_missing_cross_reference_only_warns(basic_model, tmp_path, capsys):
    """图表没在正文里被引用 → WARN，默认不算失败。"""
    basic_model["sections"] = [
        {"level": 1, "text": "1 引言"},
        {"type": "table", "caption": "孤立表", "headers": ["A"], "rows": [["1"]]},
    ]
    _, path = db.build_from_model(basic_model, str(tmp_path / "c.docx"))
    code, out = run_qa(path, capsys)
    assert code == 0
    assert "未被引用" in out


def test_strict_mode_turns_warnings_into_failure(basic_model, tmp_path, capsys):
    basic_model["sections"] = [
        {"level": 1, "text": "1 引言"},
        {"type": "table", "caption": "孤立表", "headers": ["A"], "rows": [["1"]]},
    ]
    _, path = db.build_from_model(basic_model, str(tmp_path / "d.docx"))
    code, out = run_qa(path, capsys, ["--strict"])
    assert code == 1


def test_spec_calibration(basic_model, tmp_path, capsys):
    """换配色后带 --spec 质检应通过，不带则 FAIL。"""
    spec = {"colors": {"navy": "2F4858", "blue": "1F6F8B"}}
    spec_path = tmp_path / "spec.json"
    spec_path.write_text(json.dumps(spec), encoding="utf-8")
    _, path = db.build_from_model(basic_model, str(tmp_path / "s.docx"), spec)

    code_plain, _ = run_qa(path, capsys)
    assert code_plain == 1

    code_spec, out_spec = run_qa(path, capsys, ["--spec", str(spec_path)])
    assert code_spec == 0, out_spec


def test_apply_spec_updates_expectations(tmp_path):
    spec_path = tmp_path / "s.json"
    spec_path.write_text(json.dumps({
        "colors": {"navy": "112233"},
        "sizes": {"body": 10.5, "caption": 9, "header": 8, "footer": 7.5,
                  "reference": 8.5, "h1": 15, "h2": 12},
    }), encoding="utf-8")
    qa_check.apply_spec(str(spec_path))
    try:
        assert qa_check.EXPECT["navy"] == "112233"
        assert qa_check.EXPECT["body_pt"] == 10.5
        assert qa_check.EXPECT["body_first"] == str(int(round(10.5 * 40)))
        assert qa_check.EXPECT["h1_sz"] == "30"
        assert qa_check.EXPECT["h2_sz"] == "24"
        assert qa_check.EXPECT["caption_pt"] == 9
        assert qa_check.EXPECT["footer_pt"] == 7.5
    finally:
        qa_check.EXPECT.update({
            "navy": "1F3A5F", "blue": "0072B2", "body_pt": 12,
            "body_first": "480", "h1_sz": "32", "h2_sz": "26",
            "caption_pt": 9.5, "reference_pt": 9, "header_pt": 8.5,
            "footer_pt": 8,
        })


def test_missing_file_returns_2(tmp_path, capsys):
    assert qa_check.main([str(tmp_path / "nope.docx")]) == 2


@pytest.mark.parametrize("example", ["examples/minimal/content.json",
                                     "examples/demo/content_model.json"])
def test_repo_examples_pass_qa(example, tmp_path, capsys, repo):
    import os

    model_path = os.path.join(repo, example)
    with open(model_path, encoding="utf-8") as fh:
        model = json.load(fh)
    builder, path = db.build_from_model(
        model, str(tmp_path / "ex.docx"), base_dir=os.path.dirname(model_path))
    code, out = run_qa(path, capsys)
    assert code == 0, out
    assert "失败 0 项" in out
