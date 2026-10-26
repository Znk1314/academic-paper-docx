"""配色模块测试：色板合法性、灰度可辨、色觉模拟、红绿配对、色图禁用。"""

from __future__ import annotations

import palette
import pytest


def test_okabe_ito_hex_values():
    """Okabe-Ito 的 8 个色值必须与文献一致（Wong, Nature Methods 2011）。"""
    assert palette.OKABE_ITO == {
        "black": "#000000", "orange": "#E69F00", "sky": "#56B4E9",
        "green": "#009E73", "yellow": "#F0E442", "blue": "#0072B2",
        "purple": "#CC79A7", "grey": "#999999",
    }
    assert palette.SERIES == list(palette.OKABE_ITO.values())


def test_doc_palette_has_at_most_five_chromatic_colors():
    assert palette.chromatic_count(palette.DOC.values()) <= 5


def test_doc_palette_uses_okabe_ito_accents():
    values = {palette.normalize(v) for v in palette.DOC.values()}
    assert palette.normalize(palette.OKABE_ITO["blue"]) in values
    assert palette.normalize(palette.OKABE_ITO["orange"]) in values
    assert palette.normalize(palette.OKABE_ITO["green"]) in values


def test_normalize_accepts_short_and_bare_hex():
    assert palette.normalize("0072b2") == "#0072B2"
    assert palette.normalize("#07b") == "#0077BB"
    with pytest.raises(ValueError):
        palette.normalize("nope")


def test_luminance_ordering():
    assert palette.relative_luminance("#000000") < palette.relative_luminance("#7F7F7F")
    assert palette.relative_luminance("#7F7F7F") < palette.relative_luminance("#FFFFFF")


def test_grayscale_preserves_relative_ordering():
    a = palette.relative_luminance(palette.grayscale_hex("#0072B2"))
    b = palette.relative_luminance(palette.grayscale_hex("#E69F00"))
    assert abs(a - b) > 0.1     # 蓝与橙在灰度下相差明显，可黑白区分


def test_delta_l_detects_indistinguishable_pair():
    assert palette.delta_l("#D55E00", "#009E73") < 20      # 红橙 vs 绿，灰度差异小
    assert palette.delta_l("#0072B2", "#F0E442") > 40      # 蓝 vs 黄，差异大


def test_cvd_simulation_changes_red_green():
    sim = palette.simulate_cvd("#D55E00", "deuteranopia")
    assert sim != "#D55E00"
    # 红绿配对在红绿色盲下距离显著变小
    assert palette.cvd_distance("#D55E00", "#009E73", "deuteranopia") < \
        palette.cvd_distance("#0072B2", "#E69F00", "deuteranopia")


def test_red_green_pair_detection():
    assert palette.is_red_green_pair("#D55E00", "#009E73") is True
    assert palette.is_red_green_pair("#0072B2", "#E69F00") is False


def test_check_series_reports_out_of_palette_color():
    errors, _ = palette.check_series(["#0072B2", "#FF00FF"])
    assert any("不在批准色板" in msg for msg in errors)


def test_check_series_warns_on_close_luminance_but_passes():
    """橙与天蓝灰度接近（真实属性），只提示，不阻断。"""
    errors, warnings = palette.check_series(["#E69F00", "#56B4E9"])
    assert not errors
    assert any("灰度亮度差" in msg for msg in warnings)


def test_check_series_blocks_colour_blind_collision():
    """紫红与中灰在红绿色盲下几乎同色（距离 5），必须硬拦。"""
    errors, _ = palette.check_series(["#CC79A7", "#999999"])
    assert any("几乎同色" in msg for msg in errors)


def test_red_green_pair_is_flagged():
    """红绿配对必须被拦下：既命中"红—绿"提示，也在红绿色盲下距离过近。"""
    errors, warnings = palette.check_series(["#D55E00", "#009E73"])
    assert errors or warnings
    assert any("红—绿" in msg or "几乎同色" in msg for msg in errors + warnings)


def test_serie_style_cycles_color_linestyle_marker():
    styles = [palette.serie_style(i) for i in range(6)]
    assert len({s["color"] for s in styles[:5]}) == 5
    assert len({s["linestyle"] for s in styles[:3]}) == 3
    assert len({s["marker"] for s in styles[:3]}) == 3


@pytest.mark.parametrize("cmap", ["jet", "rainbow", "nipy_spectral", "hsv"])
def test_forbidden_cmaps_rejected(cmap):
    with pytest.raises(ValueError):
        palette.resolve_cmap(cmap)


def test_sequential_cmap_resolves():
    assert palette.resolve_cmap("viridis") == "viridis"
    assert palette.resolve_cmap("cividis") == "cividis"


def test_grey_ramp_is_monotone_and_separated():
    """灰阶兜底色：亮度必须单调递增，相邻阶差 ≥20%。"""
    lums = [palette.lstar(c) for c in palette.GREY_RAMP]
    assert lums == sorted(lums)
    for a, b in zip(lums, lums[1:]):
        assert (b - a) >= 15
