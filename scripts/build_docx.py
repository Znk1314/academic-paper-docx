#!/usr/bin/env python
"""从 JSON 内容模型生成符合《学术论文 Word 排版与生成规范》的 .docx。

用法：
    python build_docx.py content.json -o out/report.docx
    python build_docx.py content.json -o out/report.docx --spec my_spec.json

内容模型字段见 templates/content_model.example.json 与 SKILL.md。
图片路径可以是相对 JSON 文件所在目录的相对路径。
"""

from __future__ import annotations

import argparse
import json
import os
import sys

import _console

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import docx_builder
from docx_builder import build_from_model  # noqa: E402


def _resolve_figures(model, base_dir):
    for item in model.get("sections", []):
        if isinstance(item, dict) and item.get("type") == "figure":
            path = item.get("path")
            if path and not os.path.isabs(path):
                item["path"] = os.path.normpath(os.path.join(base_dir, path))
    return model


def main(argv=None):
    _console.setup()
    ap = argparse.ArgumentParser(description="生成规范格式的学术论文 / 研究报告 docx")
    ap.add_argument("model", help="内容模型 JSON 路径")
    ap.add_argument("-o", "--output", help="输出 docx 路径（默认与 JSON 同名）")
    ap.add_argument("--spec", help="覆盖格式参数的 JSON（配色 / 字号等）")
    ap.add_argument("--profile", default=None,
                    choices=["thesis", "csae"],
                    help="排版模板：thesis=学位论文（默认）；csae=农业工程学报期刊稿")
    args = ap.parse_args(argv)

    model_path = os.path.abspath(args.model)
    with open(model_path, encoding="utf-8-sig") as fh:
        model = json.load(fh)
    model = _resolve_figures(model, os.path.dirname(model_path))

    overrides = None
    if args.spec:
        with open(os.path.abspath(args.spec), encoding="utf-8-sig") as fh:
            overrides = json.load(fh)

    out = args.output or os.path.splitext(model_path)[0] + ".docx"
    profile, reason = docx_builder.resolve_profile(model, args.profile)
    print("模板：%s（%s）" % (profile, reason))
    if profile == "csae":
        builder, path = docx_builder.build_journal_from_model(
            model, out, overrides, base_dir=os.path.dirname(model_path))
    else:
        builder, path = build_from_model(model, out, overrides,
                                         base_dir=os.path.dirname(model_path))

    print("生成完成：%s" % path)
    print("统计：一级标题 %d / 表 %d / 图 %d / 公式 %d"
          % (builder.h1_no, builder.tbl_no, builder.fig_no, builder.eq_no))
    for w in builder.warnings:
        print("警告：%s" % w)
    return 0


if __name__ == "__main__":
    sys.exit(main())
