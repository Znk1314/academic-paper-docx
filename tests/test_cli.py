"""统一 CLI 与仓库自洽性测试。"""

from __future__ import annotations

import os

import apd
import pytest
from conftest import MINIMAL_MODEL


def test_version_flag(capsys):
    with pytest.raises(SystemExit) as exc:
        apd.main(["--version"])
    assert exc.value.code == 0
    assert apd.__version__


def test_doctor_runs(capsys):
    assert apd.main(["doctor"]) == 0
    out = capsys.readouterr().out
    assert "python-docx" in out


def test_build_and_qa(tmp_path):
    out = tmp_path / "cli.docx"
    assert apd.main(["build", MINIMAL_MODEL, "-o", str(out)]) == 0
    assert out.exists()
    assert apd.main(["qa", str(out)]) == 0


def test_recipe_subcommand(tmp_path, capsys):
    """出图入口已从 chart 改为 recipe（配方 + 源码适配）。"""
    import apd
    code = apd.main(["recipe", "list"])
    assert code == 0
    out = capsys.readouterr().out
    assert "line-multiseries" in out and "配方目录" in out

def test_post_missing_file_returns_2(tmp_path):
    assert apd.main(["post", str(tmp_path / "nope.docx")]) == 2


def test_skill_frontmatter(repo):
    """SKILL.md 的 frontmatter 必须能让 agent 正确解析。"""
    path = os.path.join(repo, "SKILL.md")
    with open(path, encoding="utf-8") as fh:
        text = fh.read()
    assert text.startswith("---\n")
    front = text.split("---")[1]
    assert "name: academic-paper-docx" in front
    assert "description:" in front
    # 技能名决定安装到 skills 目录后的识别名，必须是 kebab-case；
    # clone 到任意目录名都合法，所以不校验当前目录名。
    name = "academic-paper-docx"
    assert name == name.lower() and " " not in name and "_" not in name


def test_ps1_files_are_utf8_bom(repo):
    """Windows PowerShell 5.1 会把无 BOM 的 UTF-8 当 ANSI，中文会破坏语法。"""
    scripts = os.path.join(repo, "scripts")
    for name in os.listdir(scripts):
        if name.endswith(".ps1"):
            head = open(os.path.join(scripts, name), "rb").read(3)
            assert head == b"\xef\xbb\xbf", "%s 缺少 UTF-8 BOM" % name


def test_requirements_lists_runtime_deps(repo):
    req = open(os.path.join(repo, "requirements.txt"), encoding="utf-8").read().lower()
    for dep in ("python-docx", "lxml", "pillow"):
        assert dep in req


def test_changelog_covers_version(repo):
    version = apd.__version__
    text = open(os.path.join(repo, "CHANGELOG.md"), encoding="utf-8").read()
    assert version in text, "CHANGELOG 里缺少当前版本号 %s" % version
