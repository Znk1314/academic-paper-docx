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


# ---------------------------------------------------------------- Layer 6 本子

def _proposal(text):
    return t.analyze_proposal([(text, False)], None)


BAD_AIMS = """早期脓毒症的检测是一个非常重要且及时的问题。本项目将探索是否可能用机器学习。
我们的中心假设是可能可以改善预警效果。
Aim 1：应用我们的模型到 ICU 数据。
Aim 2：把 Aim 1 的方法扩展到多中心数据。
Aim 3：若 Aim 1 与 Aim 2 成功，开展一次试点。
这项工作将具有变革性，有望彻底改变临床预警范式，填补国内空白。
"""

GOOD_AIMS = """Specific Aims
床旁评分报得晚且报警过多，临床难以区分真实告警与噪声。
本申请的目标是交付一个告警经过校准、并且可跨中心迁移的早期预警模型。
我们的中心假设是：融合连续生命体征与常规化验可提前数小时检出恶化。该假设基于单中心前期结果（图 1，n=812，AUC 0.88）。
Aim 1：确定回顾性数据上的检出—误报边界并给出逼近该边界的模型。
Aim 2：判断加入常规化验能否改善早期检出，以及哪些亚组受益。
Aim 3：量化告警校准的跨中心迁移性。
三个 aim 相互独立，各自都能产出可用结果。已与两家合作医院签署协议（支持信随附）。
"""


def test_proposal_detects_weak_moves():
    r = _proposal(BAD_AIMS)
    ids = {i[0] for i in r["issues"]}
    assert {"P3", "P5", "P6"} <= ids, ids
    assert r["aims"] == 3


def test_proposal_clean_when_well_formed():
    r = _proposal(GOOD_AIMS)
    assert not r["issues"], r["issues"]
    assert r["kind"] == "nih"


def test_proposal_nsf_structure_missing():
    r = t.analyze_proposal([("我们做一个很酷的项目。", False)], "nsf")
    ids = {i[0] for i in r["issues"]}
    assert "P1" in ids
    assert any("Project Summary" in i[2] for i in r["issues"] if i[0] == "P1")


def test_proposal_domino_detected():
    text = ("Aim 1：建立基线。\nAim 2：若 Aim 1 与 Aim 2 成功，开展试点。\n"
            "Great idea with 前期 数据 in 图 1.")
    r = t.analyze_proposal([(text, False)], "nih")
    assert "P4" in {i[0] for i in r["issues"]}
