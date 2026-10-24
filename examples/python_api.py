#!/usr/bin/env python
"""用 Python API 直接生成一份规范格式的论文骨架。

适合不想写 JSON 内容模型的场景；对比 examples/demo/content_model.json 的 JSON 驱动方式。

运行：
    python examples/python_api.py            # 生成 examples/demo/build/api_demo.docx
"""

from __future__ import annotations

import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SCRIPTS = os.path.join(os.path.dirname(HERE), "scripts")
sys.path.insert(0, SCRIPTS)


import _console  # noqa: E402
from docx_builder import PaperBuilder, enable_update_fields, set_print_view  # noqa: E402

_console.setup()  # 控制台非 UTF-8（如 Windows 管道 cp1252）时也能打印中文

META = {
    "org": "某某大学",
    "doc_type": "课程报告",
    "title": "示例课程报告标题",
    "subtitle": "副标题写研究方法或范围",
    "note": "本文件由 examples/python_api.py 生成，内容为占位示例，请替换",
    "info": [
        ["课程名称", "示例课程"],
        ["姓名", "（示例）"],
        ["学号", "（示例）"],
        ["提交日期", "2026 年 3 月"],
    ],
    "college": "某某学院",
    "date": "2026 年 3 月",
    "header_type": "课程报告",
}


def build(out_path: str) -> None:
    b = PaperBuilder()
    b.setup()

    b.add_cover(META)
    b.add_abstract_zh(
        "摘　要",
        ["这里写 200—350 字的中文摘要：背景、方法、阶段结果、下一步。"
         "不要虚构指标，未完成的部分如实写明。"],
        ["关键词1", "关键词2", "关键词3"],
        page_break=True,
    )
    b.add_abstract_en(
        "Abstract",
        ["Write the English abstract here, corresponding to the Chinese one."],
        ["keyword1", "keyword2"],
    )
    b.add_toc()

    b.add_heading1("1 研究背景", page_break=True)
    b.add_body("正文段落：宋体 12 pt、1.5 倍行距、首行缩进 2 字符、两端对齐。"
               "正文里引用图表公式要写清楚，例如见图1、见表1、见式（1）。")
    b.add_heading2("1.1 问题描述")
    b.add_body("二级标题下继续正文。列举时用编号段落，不要用圆点：")
    b.add_list(["第一条要点，写完整句子。", "第二条要点，与第一条保持同一层级。"])

    b.add_heading1("2 方法与结果")
    b.add_table(
        caption="示例三线表",
        headers=["项目", "取值", "说明"],
        rows=[["样本数", "120", "示例数据"], ["周期", "60 s", "示例数据"]],
        widths=[4.0, 3.0, 8.5],
        aligns=["left", "center", "left"],
        note="注：表中为占位示例数据，不代表真实测量结果。",
    )
    b.add_body("由表1可见……（正文需要引用该表编号）")

    fig = os.path.join(HERE, "demo", "figs", "fig1_duty.png")
    if os.path.exists(fig):
        b.add_figure(fig, caption="示例图", note="注：示例配图，请替换为真实图片。",
                     width_cm=14.0)
        b.add_body("如图1所示……")

    b.add_equation(r"\bar{x} = \frac{1}{n}\sum_{i=1}^{n} x_i", number="（1）")
    b.add_body("式（1）给出了样本均值的计算方式，其中 n 为样本数。")

    b.add_heading1("3 结论")
    b.add_body("总结本文工作，并说明局限性。")

    b.add_references([
        "作者. 题名[J]. 刊名, 年, 卷(期): 起止页码. [待补充/待核实]",
        "作者. 书名[M]. 出版地: 出版社, 年: 页码. [待补充/待核实]",
    ])

    b.add_header_footer(META)
    enable_update_fields(b.doc)
    set_print_view(b.doc)

    out_path = os.path.abspath(out_path)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    b.doc.save(out_path)
    print("生成完成：%s" % out_path)
    print("下一步：postprocess.ps1 更新目录 → qa_check.py 质检")


if __name__ == "__main__":
    build(os.path.join(HERE, "demo", "build", "api_demo.docx"))
