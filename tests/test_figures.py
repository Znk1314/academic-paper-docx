"""插图渲染测试：毫米版面、dpi、坐标轴标目、冗余编码、色板拦截、灰度输出。"""

from __future__ import annotations

import json
import os

import figures
import make_chart
import palette
import pytest
from PIL import Image


def _spec(**over):
    spec = {
        "kind": "line",
        "xlabel": {"label": "采样周期", "symbol": "t", "unit": "s"},
        "ylabel": {"label": "平均功耗", "symbol": "P", "unit": "mW"},
        "x": [30, 60, 90, 120],
        "series": [{"name": "本文", "values": [15.9, 10.4, 8.6, 7.4]}],
        "width_mm": 40, "dpi": 100, "style": "cn", "formats": ["png"],
    }
    spec.update(over)
    return spec


def _render(tmp_path, spec, name="f.png", grayscale=False):
    out = tmp_path / name
    figures.render(spec, str(out), grayscale=grayscale)
    return out


def test_width_is_exact_millimetres(tmp_path):
    out = _render(tmp_path, _spec(width_mm=75, dpi=100))
    with Image.open(out) as img:
        dpi = img.info["dpi"][0]
        assert img.size[0] / dpi * 25.4 == pytest.approx(75, abs=0.5)


def test_half_and_full_column_presets(tmp_path):
    for layout, expected in (("half", 75), ("full", 155)):
        out = _render(tmp_path, _spec(width_mm=layout), name="%s.png" % layout)
        with Image.open(out) as img:
            dpi = img.info["dpi"][0]
            got = img.size[0] / dpi * 25.4
        assert got == pytest.approx(expected, abs=1.0)


def test_default_dpi_is_1000_for_line_art(tmp_path):
    spec = _spec()
    spec.pop("dpi")
    out = _render(tmp_path, spec)
    with Image.open(out) as img:
        assert img.info["dpi"][0] >= 999


def test_axis_label_uses_symbol_over_unit():
    axis = figures.parse_axis({"label": "平均功耗", "symbol": "P", "unit": "mW"})
    text = figures.axis_text(axis)
    assert text.startswith("$") and "/mW" in text   # 符号斜体、单位正体


def test_axis_label_parses_legacy_string():
    axis = figures.parse_axis("时间占比 / %")
    assert axis["label"] == "时间占比" and axis["unit"] == "%"


def test_series_get_redundant_encoding(tmp_path):
    spec = _spec(series=[{"name": "A", "values": [1, 2, 3]},
                         {"name": "B", "values": [3, 2, 1]}])
    out = _render(tmp_path, spec)
    assert os.path.getsize(out) > 3000
    assert palette.serie_style(0)["linestyle"] != palette.serie_style(1)["linestyle"]


def test_red_green_pair_rejected(tmp_path):
    spec = _spec(series=[{"name": "A", "values": [1, 2], "color": "#D55E00"},
                         {"name": "B", "values": [2, 1], "color": "#009E73"}])
    with pytest.raises(ValueError) as err:
        _render(tmp_path, spec)
    assert "配色不合规" in str(err.value)


def test_out_of_palette_color_rejected(tmp_path):
    spec = _spec(series=[{"name": "A", "values": [1, 2], "color": "#FF00FF"},
                         {"name": "B", "values": [2, 1]}])
    with pytest.raises(ValueError):
        _render(tmp_path, spec)


def test_jet_heatmap_rejected(tmp_path):
    spec = {"kind": "heatmap", "cmap": "jet", "rows": ["a", "b"],
            "categories": ["x", "y", "z"], "width_mm": 40, "dpi": 100,
            "series": [{"name": "v", "values": [1, 2, 3, 4, 5, 6]}]}
    with pytest.raises(ValueError) as err:
        _render(tmp_path, spec)
    assert "禁止使用" in str(err.value)


def test_heatmap_viridis_ok(tmp_path):
    spec = {"kind": "heatmap", "cmap": "cividis", "rows": ["a", "b"],
            "categories": ["x", "y", "z"],
            "series": [{"name": "v", "values": [1, 2, 3, 4, 5, 6]}],
            "width_mm": 40, "dpi": 100}
    out = _render(tmp_path, spec)
    assert out.exists()


def test_grayscale_output(tmp_path):
    out = _render(tmp_path, _spec(), grayscale=True)
    grey = tmp_path / (out.stem + "-grey.png")
    assert out.exists() and grey.exists()


def test_bar_zero_base(tmp_path):
    spec = {"kind": "bar", "categories": ["a", "b"], "width_mm": 40, "dpi": 100,
            "series": [{"name": "v", "values": [3, 5]}]}
    out = _render(tmp_path, spec)
    assert out.exists()


def test_errorbar_and_significance(tmp_path):
    spec = _spec(series=[{"name": "A", "values": [1, 2, 3], "error": [0.1, 0.2, 0.1]}])
    spec["brackets"] = [{"x1": 0, "x2": 2, "label": "**"}]
    out = _render(tmp_path, spec)
    assert out.exists()


def test_unknown_kind_rejected(tmp_path):
    with pytest.raises(ValueError):
        _render(tmp_path, _spec(kind="pie3d"))


def test_make_chart_compat_maps_px_to_mm():
    legacy = {"type": "bar", "categories": ["a"], "series": [{"name": "v", "values": [1]}],
              "width_px": 1672, "y_min": 0, "y_max": 10}
    converted = make_chart.convert(legacy)
    assert converted["kind"] == "bar"
    assert converted["width_mm"] == pytest.approx(141.6, abs=0.5)
    assert converted["y_range"] == [0, 10]


def test_cli_renders_file(tmp_path, capsys):
    spec_path = tmp_path / "spec.json"
    spec_path.write_text(json.dumps(_spec(), ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "cli.png"
    assert figures.main([str(spec_path), "-o", str(out)]) == 0
    assert out.exists()

def test_theme_palette_drives_series_colors(tmp_path):
    """spec["palette"] 用色卡的安全子集，不再用 Okabe-Ito。"""
    spec = _spec(palette="rose6", series=[{"name": "A", "values": [1, 2, 3]},
                                         {"name": "B", "values": [3, 2, 1]},
                                         {"name": "C", "values": [2, 3, 1]}])
    out = _render(tmp_path, spec)
    assert out.exists()
    colors = [s["color"] for s in figures._series(spec)]
    assert set(colors) <= set(palette.THEMES["rose6"]["colors"])


def test_explicit_color_beats_palette(tmp_path):
    spec = _spec(palette="rose6", series=[{"name": "A", "values": [1, 2], "color": "#0072B2"},
                                          {"name": "B", "values": [2, 1]}])
    colors = [s["color"] for s in figures._series(spec)]
    assert colors[0] == "#0072B2"


def test_unknown_palette_rejected(tmp_path):
    spec = _spec(palette="不存在的色卡", series=[{"name": "A", "values": [1, 2]}])
    with pytest.raises(ValueError):
        figures._series(spec)
