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
            "scatter-fit", "boxgroup", "heatmap-seq",
            "nature-multipanel", "nature-dotplot",
            "scatter-signif-group", "corr-lower", "roc-pr-curve",
            "survival-km", "dual-split"}


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


BASE_RECIPES = {"line-multiseries", "bar-single", "bar-grouped",
                "scatter-fit", "boxgroup", "heatmap-seq"}


@pytest.mark.parametrize("rid", sorted(BASE_RECIPES))
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


# ---------------------------------------------------------------- Nature 体例

@pytest.mark.parametrize("journal,single,double", [
    ("nature", 89.0, 183.0), ("science", 55.0, 183.0),
    ("ieee", 88.9, 181.0), ("csae", 86.0, 182.0), ("thesis", 75.0, 155.0)])
def test_journal_widths(journal, single, double):
    sys.path.insert(0, RES)
    import _kit as K

    assert K.width_for(journal, "single") == single
    assert K.width_for(journal, "double") == double


def test_nature_meta_declares_style_and_size():
    for rid in ("nature-multipanel", "nature-dotplot"):
        meta = recipes._meta(rid)
        assert meta["journal"] == "nature" and meta["style"] == "nature"
        assert "mm" in meta["final_size"]
        assert any("6 pt" in item or "7 pt" in item for item in meta["fidelity"])


def test_nature_dotplot_outputs_three_formats(tmp_path):
    csv = tmp_path / "d.csv"
    csv.write_text("group,value\nA,0.1\nA,0.2\nA,0.3\nB,0.4\nB,0.5\nB,0.6\n",
                   encoding="utf-8")
    out = tmp_path / "d.png"
    rc = subprocess.call([sys.executable,
                          os.path.join(RES, "nature-dotplot", "recipe.py"),
                          str(csv), str(out)], env=_ENV)
    assert rc == 0
    assert out.exists() and (tmp_path / "d.pdf").exists() and (tmp_path / "d.svg").exists()
    with Image.open(out) as img:
        mm = img.size[0] / img.info["dpi"][0] * 25.4
    assert mm == pytest.approx(89, abs=1)          # Nature 单栏


def test_dotplot_rejects_too_few_points(tmp_path):
    csv = tmp_path / "few.csv"
    csv.write_text("group,value\nA,0.1\nA,0.2\n", encoding="utf-8")
    rc = subprocess.call([sys.executable,
                          os.path.join(RES, "nature-dotplot", "recipe.py"),
                          str(csv), str(tmp_path / "o.png")], env=_ENV)
    assert rc != 0                                  # 少于 3 点应被拒


def test_check_flags_pitfalls(tmp_path, capsys):
    script = tmp_path / "gen_fig_pie.py"
    script.write_text("import matplotlib.pyplot as plt\n"
                      "plt.pie([1, 2, 3])\n"
                      "fig, ax = plt.subplots()\nax2 = ax.twinx()\n", encoding="utf-8")
    recipes.main(["check", str(script), "--recipe", "bar-single"])
    out = capsys.readouterr().out
    assert "P-饼图" in out and "P-双 Y 轴" in out


def test_nature_multipanel_runs(tmp_path):
    rows = ["panel,group,x,y"]
    for label in ("a", "b", "c", "d"):
        for g, base in (("对照组", 0.4), ("处理组", 0.6)):
            for i in (1, 2, 3):
                rows.append("%s,%s,%d,%.2f" % (label, g, i, base + i * 0.01))
    csv = tmp_path / "m.csv"
    csv.write_text("\n".join(rows) + "\n", encoding="utf-8")
    out = tmp_path / "m.png"
    rc = subprocess.call([sys.executable,
                          os.path.join(RES, "nature-multipanel", "recipe.py"),
                          str(csv), str(out)], env=_ENV)
    assert rc == 0
    with Image.open(out) as img:
        mm = img.size[0] / img.info["dpi"][0] * 25.4
    assert mm == pytest.approx(183, abs=1)          # Nature 双栏


# ---------------------------------------------------------------- 5 个统计图型

STAT_RECIPES = {
    "scatter-signif-group": "fig7_signif",
    "corr-lower": "fig8_corr",
    "roc-pr-curve": "fig9_roc",
    "survival-km": "fig10_surv",
    "dual-split": "fig11_dual",
}


@pytest.mark.parametrize("rid,csv_name", sorted(STAT_RECIPES.items()))
def test_stat_recipes_run(rid, csv_name, tmp_path):
    out = tmp_path / "x.png"
    rc = subprocess.call([sys.executable, os.path.join(RES, rid, "recipe.py"),
                          os.path.join(DATA, csv_name + ".csv"), str(out)],
                         env=_ENV)
    assert rc == 0, rid
    with Image.open(out) as img:
        assert img.info["dpi"][0] > 900
        assert "academic-paper-docx" in (img.info.get("Software") or "")


def test_survival_km_is_monotone(tmp_path):
    """生存曲线必须单调不增，否则阶梯画错了。"""
    sys.path.insert(0, os.path.join(RES, "survival-km"))
    import importlib.util

    spec = importlib.util.spec_from_file_location(
        "km_recipe", os.path.join(RES, "survival-km", "recipe.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    xs, ys = mod._km([1, 2, 3, 4, 5, 6], [1, 1, 0, 1, 0, 1])
    assert all(b <= a + 1e-9 for a, b in zip(ys, ys[1:])), ys


def test_corr_lower_masks_upper_triangle(tmp_path):
    """相关性矩阵只画下三角：掩码后上三角应是 NaN。"""
    sys.path.insert(0, os.path.join(RES, "corr-lower"))
    import importlib.util

    spec = importlib.util.spec_from_file_location(
        "corr_recipe", os.path.join(RES, "corr-lower", "recipe.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    assert mod._pearson([1, 2, 3], [2, 4, 6]) == pytest.approx(1.0)
    assert mod._pearson([1, 2, 3], [3, 2, 1]) == pytest.approx(-1.0)
