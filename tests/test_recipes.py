"""配方出图测试：目录、数据契约、源码拷出、跑图与元数据。"""

from __future__ import annotations

import os
import subprocess
import sys

import pytest
from PIL import Image

import recipes

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RES = os.path.join(REPO, "recipes")
DATA = os.path.join(REPO, "examples", "demo", "data")

# Windows CI 默认 cp1252 控制台，子进程必须强制 UTF-8（否则中文打印会抛异常）
_ENV = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8")

EXPECTED = {"line-multiseries", "bar-single", "bar-grouped",
            "scatter-fit", "boxgroup", "heatmap-seq"}


def test_catalog_has_expected_recipes():
    assert EXPECTED <= set(recipes._ids())


@pytest.mark.parametrize("rid", sorted(EXPECTED))
def test_meta_is_complete(rid):
    meta = recipes._meta(rid)
    assert meta["id"] == rid
    assert meta["title"] and meta["when"]
    assert meta["data_contract"]["required_columns"]
    assert meta["fidelity"] and meta["verify"]
    assert os.path.exists(os.path.join(RES, rid, "recipe.py"))
    assert os.path.exists(os.path.join(RES, rid, "preview.png"))


@pytest.mark.parametrize("rid", sorted(EXPECTED))
def test_recipe_runs_and_writes_metadata(rid, tmp_path):
    csv = {"line-multiseries": "fig2_power", "bar-single": "fig1_duty",
           "bar-grouped": "fig3_nodes", "scatter-fit": "fig4_scatter",
           "boxgroup": "fig5_box", "heatmap-seq": "fig6_heatmap"}[rid]
    out = tmp_path / "fig.png"
    rc = subprocess.call([sys.executable, os.path.join(RES, rid, "recipe.py"),
                          os.path.join(DATA, csv + ".csv"), str(out)],
                         env=_ENV)
    assert rc == 0
    with Image.open(out) as img:
        dpi = img.info["dpi"][0]
        assert dpi > 900                       # 默认 1000 dpi
        assert img.size[0] / dpi * 25.4 == pytest.approx(140, abs=1)
        assert "academic-paper-docx" in (img.info.get("Software") or "")


def test_new_copies_full_source(tmp_path):
    rc = recipes.main(["new", "bar-single", "--out", str(tmp_path)])
    assert rc == 0
    script = tmp_path / "gen_fig_bar_single.py"
    assert script.exists()
    assert (tmp_path / "_kit.py").exists()
    body = script.read_text(encoding="utf-8")
    assert "def main()" in body and "ax.bar(" in body   # 完整源码，不是占位


def test_check_validates_data_contract(tmp_path, capsys):
    bad = tmp_path / "bad.csv"
    bad.write_text("wrong,columns\n1,2\n", encoding="utf-8")
    rc = recipes.main(["check", os.path.join(RES, "bar-single", "recipe.py"),
                       "--data", str(bad)])
    assert rc == 1
    assert "缺列" in capsys.readouterr().out


def test_check_reports_design_changes(tmp_path, capsys):
    rc = recipes.main(["new", "line-multiseries", "--out", str(tmp_path)])
    assert rc == 0
    script = tmp_path / "gen_fig_line_multiseries.py"
    text = script.read_text(encoding="utf-8").replace('width_mm=140.0', 'width_mm=75.0')
    script.write_text(text, encoding="utf-8")
    rc = recipes.main(["check", str(script),
                       "--data", os.path.join(DATA, "fig2_power.csv")])
    assert rc == 0
    out = capsys.readouterr().out
    assert "数据契约通过" in out and "设计参数" in out


def test_heatmap_rejects_rainbow(tmp_path):
    csv = tmp_path / "h.csv"
    csv.write_text("row,a,b\nr1,1,2\n", encoding="utf-8")
    script = tmp_path / "gen.py"
    body = open(os.path.join(RES, "heatmap-seq", "recipe.py"), encoding="utf-8").read()
    body = body.replace('cmap="cividis"', 'cmap="jet"')
    import shutil
    shutil.copyfile(os.path.join(RES, "_kit.py"), tmp_path / "_kit.py")
    script.write_text(body, encoding="utf-8")
    rc = subprocess.call([sys.executable, str(script), str(csv),
                          str(tmp_path / "o.png")], env=_ENV)
    assert rc != 0                          # 拒绝彩虹色谱
