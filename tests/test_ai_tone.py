"""AI 味检测器的测试：能抓到、不误伤、规则与文档一一对应。"""

from __future__ import annotations

import json
import os

import ai_tone_check as tone
import pytest
from conftest import DEMO_MODEL


def paras(*texts):
    return [(t, False) for t in texts]


def ids(report):
    return {h.rule_id for h in report.hits}


# ---------------------------------------------------------------- 判定方向

def test_clean_academic_text_scores_low():
    report = tone.analyze(paras(
        "在 12 个节点、60 s 采样周期下，平均功耗为 7.4 mW。",
        "节点数增加到 20 时，误差从 0.21 ℃ 升到 0.31 ℃。",
        "原因是人员进出使温度的时间尺度接近采样间隔，滑动平均无法抑制阶跃扰动。",
    ), "clean")
    assert report.score < 20, tone.render(report)
    assert "S2.1" not in ids(report)


def test_ai_dense_text_scores_high():
    report = tone.analyze(paras(
        "随着物联网技术的快速发展，实验室环境监测受到了广泛关注。",
        "这不仅是一次技术改进，更是对传统思路的深刻变革。",
        "实验结果表明，该方法显著提升了监测精度，取得了良好的效果。",
        "众所周知，该框架具有广阔的应用前景。",
        "综上所述，本文为后续工作奠定了坚实的基础。",
    ), "dense")
    assert report.score >= 60, tone.render(report)
    assert {"S2.1", "S2.3"} <= ids(report)


def test_score_grows_with_more_tells():
    base = tone.analyze(paras("在 12 个节点上测得平均功耗 7.4 mW。"), "a")
    noisy = tone.analyze(paras(
        "在 12 个节点上测得平均功耗 7.4 mW。",
        "该方法具有重要的理论意义和现实价值。",
        "专家认为它具有广阔的应用前景。",
    ), "b")
    assert noisy.score > base.score


def test_very_short_text_does_not_score():
    assert tone.analyze(paras("测试。"), "x").score == 0


def test_grade_thresholds():
    assert tone.grade(0) == "清爽"
    assert tone.grade(25) == "轻微"
    assert tone.grade(45) == "明显"
    assert tone.grade(80) == "重"


# ---------------------------------------------------------------- 豁免与边界

def test_inflation_exempt_when_sentence_has_data():
    """S2.2 只在"没有数字"时报警。"""
    report = tone.analyze(paras("改进方法的 F1 从 0.71 显著提升到 0.79，样本 120 组。"), "t")
    assert "S2.2" not in ids(report)

    report2 = tone.analyze(paras("改进方法显著提升了检测精度。"), "t")
    assert "S2.2" in ids(report2)


def test_quoted_and_cited_content_is_stripped():
    report = tone.analyze(paras(
        '作者原话是“具有广阔的应用前景”，但正文没有给出任何数据。',
        "文献[3]报告该现象，未说明原因。",
    ), "t")
    assert "S2.1" not in ids(report), tone.render(report)


def test_numeric_range_dash_is_not_decoration():
    report = tone.analyze(paras("采样周期取 60—90 s 时，误差保持在 0.31 ℃ 以内。"), "t")
    assert "B2.13" not in ids(report)

    report2 = tone.analyze(paras("结果终于出现——文件无法导出。"), "t")
    assert "B2.13" in ids(report2)


def test_method_passive_is_exempt():
    """方法段的被动/无主句是学术惯例，B2.10 不应命中。"""
    report = tone.analyze(paras("样本在 4 ℃ 下保存 24 h，随后以 3000 r/min 离心 10 min。"), "t")
    assert "B2.10" not in ids(report)


def test_repeat_closer_only_at_paragraph_end():
    """S2.4 只认"段末那句"，段中的转折/推进行为不算复述收尾。"""
    mid = tone.analyze(paras("由此可见该方法有效。后续补充了三组对照，误差降到 0.21 ℃。"), "t")
    assert "S2.4" not in ids(mid), tone.render(mid)

    end = tone.analyze(paras("三组对照的误差分别是 0.21、0.24、0.31 ℃。由此可见该方法有效。"), "t")
    assert "S2.4" in ids(end)


def test_heading_repeat_detected():
    report = tone.analyze([("3.1 实验设置", True), ("本节介绍实验设置。", False)], "t")
    assert "A3.1" in ids(report)


# ---------------------------------------------------------------- 规则表自洽

def test_rule_ids_unique():
    rule_ids = [r.id for r in tone.RULES]
    assert len(rule_ids) == len(set(rule_ids))


def test_every_rule_documented(repo):
    """规则编号必须能在 references/anti-ai-academic-zh.md 里查到，避免代码与文档脱节。"""
    path = os.path.join(repo, "references", "anti-ai-academic-zh.md")
    with open(path, encoding="utf-8") as fh:
        doc = fh.read()
    missing = [r.id for r in tone.RULES if r.id not in doc]
    assert not missing, "文档缺少规则说明：%s" % missing


# ---------------------------------------------------------------- 输入与 CLI

def test_collect_from_repo_demo_model():
    paragraphs, name = tone.collect_paragraphs_model(DEMO_MODEL)
    assert paragraphs and name.endswith(".json")
    report = tone.analyze(paragraphs, name)
    assert report.chars > 1000
    assert report.score < 40


def test_cli_text_and_threshold(tmp_path, capsys):
    clean = tmp_path / "clean.md"
    clean.write_text("在 12 个节点上测得平均功耗 7.4 mW，采样周期 60 s。\n", encoding="utf-8")
    assert tone.main(["--text", str(clean), "--max-score", "40"]) == 0

    dirty = tmp_path / "dirty.md"
    dirty.write_text(
        "随着技术的发展，该问题受到了广泛关注。\n\n"
        "这不仅是一次改进，更是深刻的变革。\n\n"
        "综上所述，本文具有重要的理论意义和现实价值。\n", encoding="utf-8")
    assert tone.main(["--text", str(dirty), "--max-score", "40"]) == 1


def test_cli_json_output(tmp_path, capsys):
    path = tmp_path / "dirty.md"
    path.write_text("该方法具有重要的理论意义和现实价值。\n", encoding="utf-8")
    tone.main(["--text", str(path), "--json-out"])
    payload = json.loads(capsys.readouterr().out)
    assert payload["score"] >= 0
    assert payload["grade"] in {"清爽", "轻微", "明显", "重"}
    assert payload["hits"]


def test_cli_missing_paragraphs_exit_2(tmp_path):
    empty = tmp_path / "empty.md"
    empty.write_text("\n\n", encoding="utf-8")
    assert tone.main(["--text", str(empty)]) == 2


@pytest.mark.parametrize("kind", ["--docx", "--json", "--text"])
def test_mutually_exclusive_source_requires_one(kind):
    with pytest.raises(SystemExit):
        tone.main([])
