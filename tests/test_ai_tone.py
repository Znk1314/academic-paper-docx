"""学术文风审计测试：规则来自 academic-humanizer（MIT）。

覆盖三类：Layer 1 通用痕迹、Layer 2 学术痕迹、Layer 4 主张-证据；
外加 Layer 3 的"保留清单"（有证据的写法不许误报）。
"""

from __future__ import annotations

import ai_tone_check as t


def _analyze(text):
    return t.analyze([(text, False)], "inline")


# ---------------------------------------------------------------- Layer 1

def test_significance_hype_detected():
    r = _analyze("该方法具有重要的理论意义和应用价值。")
    assert {"T1.1"} & {h.rule_id for h in r.hits}
    assert r.tells_score > 0


def test_promotional_words_detected():
    r = _analyze("本研究赋能产业升级，助力乡村振兴，形成闭环。")
    assert "T1.3" in {h.rule_id for h in r.hits}


def test_unsourced_authority_detected():
    r = _analyze("专家认为该方法适合该场景。")
    assert "T1.4" in {h.rule_id for h in r.hits}


def test_filler_detected():
    r = _analyze("值得注意的是，该结果稳定。")
    assert "T1.10" in {h.rule_id for h in r.hits}


def test_em_dash_detected():
    r = _analyze("该流程——从采集到上传——均可本地完成。")
    assert "T1.12" in {h.rule_id for h in r.hits}


def test_long_sentence_detected():
    r = _analyze("现有方法虽然具有一定的可解释性，但其参数依赖人工设定，"
                 "在样本分布发生变化时容易出现偏移，从而限制了进一步推广。")
    assert {"T1.11", "T2.11"} & {h.rule_id for h in r.hits}


# ---------------------------------------------------------------- Layer 2

def test_overclaim_verb_detected():
    r = _analyze("本文方法证明了该假设，并彻底解决了误差累积问题。")
    ids = {h.rule_id for h in r.hits}
    assert "T2.1" in ids


def test_overclaim_verb_not_flagged_with_numbers():
    """有数字时视为有证据（Layer 3），不报。"""
    r = _analyze("在 3 个数据集上，本文方法的准确率由 0.86 提升到 0.91。")
    assert "T2.1" not in {h.rule_id for h in r.hits}


def test_novelty_padding_detected():
    r = _analyze("据我们所知，这是首次研究该问题。")
    assert "T2.4" in {h.rule_id for h in r.hits}


def test_formulaic_opener_detected():
    r = _analyze("随着深度学习技术的快速发展，该问题受到广泛关注。")
    assert "T2.5" in {h.rule_id for h in r.hits}


def test_contribution_cliche_detected():
    r = _analyze("本文的主要贡献如下：首先提出一个框架。")
    assert "T2.7" in {h.rule_id for h in r.hits}


def test_vague_hedging_detected():
    r = _analyze("该方法较为稳定，效果比较好。")
    assert "T2.9" in {h.rule_id for h in r.hits}


# ---------------------------------------------------------------- Layer 4

def test_unbacked_claim_detected():
    r = _analyze("本文方法优于现有方法。")
    assert "C1" in {h.rule_id for h in r.hits}
    assert r.claims_score > 0


def test_vague_magnitude_detected():
    r = _analyze("该设计将功耗降低了不少。")
    assert "C2" in {h.rule_id for h in r.hits}


def test_overgeneralized_claim_detected():
    r = _analyze("该模型能够广泛应用于各类场景。")
    assert "C4" in {h.rule_id for h in r.hits}


# ---------------------------------------------------------------- Layer 3 保留

def test_evidence_tied_hedging_is_preserved():
    """证据型模糊限制是正确写法，不许报。"""
    r = _analyze("结果提示该趋势可能与采样间隔有关（图 3）。")
    assert not r.hits, [h.rule_id for h in r.hits]


def test_passive_and_we_are_not_flagged():
    r = _analyze("样本在 4 ℃ 下离心 10 min。We compare three baselines.")
    assert not r.hits


def test_clean_academic_text_scores_zero():
    r = _analyze("在节点数为 12、采样周期为 60 s 的条件下，平均功耗由 15.9 mW 降至 10.4 mW，"
                 "降幅 34.6%。结果表明通信占空比是该场景下功耗的主要来源。")
    assert r.score == 0, [h.rule_id for h in r.hits]


def test_metrics_and_grades():
    r = _analyze("随着技术的发展，该问题受到广泛关注。")
    assert r.metrics["句长均值"] > 0
    assert t.grade(0) == "干净"
    assert t.grade(80) == "重"
