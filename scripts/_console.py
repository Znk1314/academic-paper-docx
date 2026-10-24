"""让命令行脚本在"非 UTF-8 控制台"下也能正常输出中文。

为什么需要：
    Windows 上 `sys.stdout` 被重定向成管道时（GitHub Actions、CI、`| more`、被别的
    程序调用），编码取的是系统代码页（如 cp1252）。此时 print 任何中文都会抛
    UnicodeEncodeError → 脚本退出码 1。本地手动跑通常因为终端是 UTF-8 或设置了
    PYTHONIOENCODING 而看不出来。

做法：
    在 main() 开头调用 setup()，把 stdout/stderr 重配成 UTF-8，
    并用 errors="replace" 兜底（哪怕重配失败也不会因为编码问题崩掉）。
"""

from __future__ import annotations

import sys


def setup() -> None:
    """把 stdout/stderr 重配为 UTF-8。可重复调用，失败也不抛异常。"""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # pragma: no cover - 只可能在非常规流上发生
            pass
