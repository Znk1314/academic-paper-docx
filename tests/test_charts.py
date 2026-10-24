"""图表生成测试：尺寸、配色、边界情况。"""

from __future__ import annotations

import json
import os

import make_chart
import pytest
from PIL import Image


def _write(tmp_path, spec):
    path = tmp_path / "spec.json"
    path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
    return path


def test_bar_chart_size(tmp_path):
    spec = {
        "type": "bar", "xlabel": "类别", "ylabel": "占比 / %",
        "categories": ["甲", "乙", "丙"],
        "series": [{"name": "占比", "values": [40, 35, 25]}],
    }
    out = tmp_path / "bar.png"
    make_chart.render(spec, str(out))
    with Image.open(out) as img:
        assert img.size == (1672, 941)
        assert img.format == "PNG"


def test_line_chart_with_explicit_x(tmp_path):
    spec = {
        "type": "line", "xlabel": "周期 / s", "ylabel": "功耗 / mW",
        "x": [30, 60, 90, 120],
        "series": [{"name": "功耗", "values": [15.9, 10.4, 8.6, 7.4]}],
    }
    out = tmp_path / "line.png"
    make_chart.render(spec, str(out))
    assert out.stat().st_size > 5000


def test_grouped_bar_legend_and_palette(tmp_path):
    spec = {
        "type": "grouped_bar",
        "categories": ["基线", "改进"],
        "series": [
            {"name": "精度", "values": [0.71, 0.79]},
            {"name": "召回", "values": [0.68, 0.75]},
        ],
        "y_min": 0, "y_max": 1,
    }
    out = tmp_path / "grouped.png"
    make_chart.render(spec, str(out))
    with Image.open(out) as img:
        colors = img.convert("RGB").getcolors(maxcolors=100000)
    # Okabe-Ito 蓝 #0072B2 与橙 #E69F00 都应出现
    rgb = {c for _, c in colors}
    assert (0, 114, 178) in rgb
    assert (230, 159, 0) in rgb


def test_empty_series_raises(tmp_path):
    with pytest.raises(ValueError):
        make_chart.render({"type": "bar", "categories": [], "series": []},
                          str(tmp_path / "x.png"))


def test_custom_size(tmp_path):
    out = tmp_path / "small.png"
    make_chart.render({"type": "bar", "categories": ["a"], "series": [
        {"name": "s", "values": [1]}], "width_px": 800, "height_px": 450}, str(out))
    with Image.open(out) as img:
        assert img.size == (800, 450)


def test_nice_ticks_cover_maximum():
    ticks = make_chart._nice_ticks(0, 17.8, 5)
    assert ticks[0] <= 0
    assert ticks[-1] >= 17.8
    assert len(ticks) >= 3


def test_value_formatter_is_consistent():
    fmt = make_chart._value_formatter([92.0, 3.5, 1.5, 3.0])
    assert fmt(92.0) == "92.0"
    assert fmt(3) == "3.0"
    assert fmt(3.5) == "3.5"


def test_cli_writes_file(tmp_path):
    spec_path = _write(tmp_path, {
        "type": "bar", "categories": ["甲"], "series": [{"name": "值", "values": [1]}]})
    out = tmp_path / "cli.png"
    assert make_chart.main([str(spec_path), "-o", str(out)]) == 0
    assert os.path.exists(out)
