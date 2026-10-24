"""回归测试：非 UTF-8 控制台下 CLI 不能因为打印中文而崩。

历史 bug：GitHub Actions 的 Windows runner 里 stdout 是管道，编码取系统代码页
（cp1252）。`python scripts/apd.py doctor` 打印"环境自检"直接抛
UnicodeEncodeError，退出码 1，整条 CI 红掉。修复方式见 scripts/_console.py。
"""

from __future__ import annotations

import os
import subprocess
import sys


def _env_with_encoding(encoding: str) -> dict:
    env = dict(os.environ)
    env["PYTHONIOENCODING"] = encoding
    env.pop("PYTHONUTF8", None)
    return env


def _run(repo: str, args: list, encoding: str = "cp1252") -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, os.path.join(repo, "scripts", "apd.py")] + args,
        capture_output=True,
        env=_env_with_encoding(encoding),
        cwd=repo,
    )


def test_doctor_survives_non_utf8_stdout(repo):
    result = _run(repo, ["doctor"])
    assert result.returncode == 0, result.stderr.decode("utf-8", "replace")
    # 输出必须是 UTF-8（而不是降级成 '?' 或抛异常）
    text = result.stdout.decode("utf-8", "strict")
    assert "python-docx" in text
    assert "环境自检" in text


def test_qa_survives_non_utf8_stdout(repo, tmp_path):
    docx = tmp_path / "enc.docx"
    build = subprocess.run(
        [sys.executable, os.path.join(repo, "scripts", "apd.py"), "build",
         os.path.join(repo, "examples", "minimal", "content.json"), "-o", str(docx)],
        capture_output=True, cwd=repo,
    )
    assert build.returncode == 0, build.stderr.decode("utf-8", "replace")

    result = _run(repo, ["qa", str(docx)])
    assert result.returncode == 0, result.stderr.decode("utf-8", "replace")
    text = result.stdout.decode("utf-8", "strict")
    assert "质检报告" in text
    assert "失败 0 项" in text


def test_help_survives_non_utf8_stdout(repo):
    """argparse 的 --help 里也有中文，同样不能崩。"""
    result = _run(repo, ["--help"])
    assert result.returncode == 0, result.stderr.decode("utf-8", "replace")
    assert "环境自检" in result.stdout.decode("utf-8", "strict")


def test_python_api_example_survives_non_utf8_stdout(repo):
    """CI 会直接跑这个示例脚本，它也必须不崩。"""
    result = subprocess.run(
        [sys.executable, os.path.join(repo, "examples", "python_api.py")],
        capture_output=True,
        env=_env_with_encoding("cp1252"),
        cwd=repo,
    )
    assert result.returncode == 0, result.stderr.decode("utf-8", "replace")
    assert "生成完成" in result.stdout.decode("utf-8", "strict")
