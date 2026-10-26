#!/usr/bin/env python
"""academic-paper-docx 统一命令行。

    python scripts/apd.py doctor                       环境自检
    python scripts/apd.py build content.json -o out.docx   内容模型 → docx
    python scripts/apd.py post out.docx                更新域/目录 + 导 PDF + 固定页面视图
    python scripts/apd.py qa out.docx --pdf out.pdf    规范自动质检
    python scripts/apd.py preview out.pdf              逐页预览 PNG
    python scripts/apd.py chart chart.json -o fig.png  生成规范配色图表
    python scripts/apd.py demo -o build/demo           一键跑完整流程（含示例内容）

各子命令等价于单独调用 scripts/ 下对应脚本；这里是给不想记路径的人用的。
"""

from __future__ import annotations

import argparse
import glob
import json
import os
import shutil
import subprocess
import sys

import _console

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from _version import __version__  # noqa: E402

# ------------------------------------------------------------------ 子命令

def cmd_build(args):
    import build_docx
    argv = [args.model]
    if args.output:
        argv += ["-o", args.output]
    if args.spec:
        argv += ["--spec", args.spec]
    return build_docx.main(argv)


def cmd_qa(args):
    import qa_check
    argv = [args.docx]
    if args.pdf:
        argv += ["--pdf", args.pdf]
    if args.spec:
        argv += ["--spec", args.spec]
    if args.strict:
        argv += ["--strict"]
    return qa_check.main(argv)


def cmd_preview(args):
    import render_preview
    argv = [args.pdf, "-o", args.out, "--dpi", str(args.dpi)]
    if args.pages:
        argv += ["--pages", args.pages]
    return render_preview.main(argv)


def cmd_chart(args):
    import make_chart
    return make_chart.main([args.spec, "-o", args.output])


def cmd_figure(args):
    import figures
    argv = [args.spec, "-o", args.output]
    if args.grayscale:
        argv.append("--grayscale")
    return figures.main(argv)


def cmd_tone(args):
    import ai_tone_check
    argv = []
    if args.docx:
        argv += ["--docx", args.docx]
    elif args.json:
        argv += ["--json", args.json]
    else:
        argv += ["--text", args.text]
    if args.max_score is not None:
        argv += ["--max-score", str(args.max_score)]
    if args.json_out:
        argv += ["--json-out"]
    return ai_tone_check.main(argv)


def cmd_post(args):
    docx = os.path.abspath(args.docx)
    if not os.path.exists(docx):
        print("找不到 docx：%s" % docx, file=sys.stderr)
        return 2
    pdf = os.path.abspath(args.pdf) if args.pdf else os.path.splitext(docx)[0] + ".pdf"

    if os.name == "nt":
        exe = shutil.which("pwsh") or shutil.which("powershell")
        if not exe:
            print("找不到 PowerShell，无法使用 Word COM 更新目录。", file=sys.stderr)
            return 2
        script = os.path.join(HERE, "postprocess.ps1")
        cmd = [exe, "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script,
               "-Docx", docx]
        if not args.no_pdf:
            cmd += ["-Pdf", pdf]
        if args.no_pdf:
            cmd += ["-NoPdf"]
        if args.no_view_reset:
            cmd += ["-NoViewReset"]
        return subprocess.call(cmd)

    bash = shutil.which("bash")
    if not bash:
        print("非 Windows 且找不到 bash，请手动导出 PDF 后运行 set_view.py。",
              file=sys.stderr)
        return 2
    script = os.path.join(HERE, "postprocess.sh")
    return subprocess.call([bash, script, docx, pdf])


def cmd_doctor(args):
    checks = []

    def probe(name, fn, critical=False):
        try:
            detail = fn()
            checks.append(("OK", name, str(detail or ""), critical))
        except Exception as exc:
            checks.append(("MISSING", name, str(exc), critical))

    def python_docx():
        import docx
        return "python-docx %s" % getattr(docx, "__version__", "?")

    def lxml():
        import lxml.etree
        return "lxml %s" % lxml.etree.__version__

    def pillow():
        import PIL
        return "Pillow %s" % PIL.__version__

    def pymupdf():
        try:
            import pymupdf
            return "PyMuPDF %s" % pymupdf.__doc__.split()[-1]
        except Exception:
            return "PyMuPDF (fitz)"

    def word():
        if os.name != "nt":
            raise RuntimeError("非 Windows，无 Word COM")
        for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
            if not base:
                continue
            cand = os.path.join(base, "Microsoft Office", "Root", "Office16", "WINWORD.EXE")
            if os.path.exists(cand):
                return cand
        raise RuntimeError("未找到 WINWORD.EXE（可手写 postprocess.ps1 的路径）")

    def soffice():
        for name in ("soffice", "libreoffice"):
            path = shutil.which(name)
            if path:
                return path
        raise RuntimeError("未安装 LibreOffice（非 Windows 导出 PDF 用）")

    def powershell():
        for name in ("pwsh", "powershell"):
            path = shutil.which(name)
            if path:
                return path
        raise RuntimeError("未找到 PowerShell")

    def cjk_font():
        for path in (r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\simhei.ttf",
                     "/System/Library/Fonts/PingFang.ttc",
                     "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
                     "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"):
            if os.path.exists(path):
                return path
        raise RuntimeError("无中文字体，make_chart.py 的中文标签会显示方框")

    probe("python-docx", python_docx, critical=True)
    probe("lxml", lxml, critical=True)
    probe("Pillow（生成图表）", pillow)
    probe("PyMuPDF（PDF 预览/体检）", pymupdf)
    probe("中文字体（图表标签）", cjk_font)
    probe("PowerShell", powershell)
    probe("Microsoft Word（更新目录/导 PDF）", word)
    probe("LibreOffice（无 Word 时的兜底）", soffice)

    print("academic-paper-docx %s  环境自检" % __version__)
    print("=" * 62)
    missing_critical = False
    for state, name, detail, critical in checks:
        print("  %-8s %-28s %s" % (state, name, detail))
        if state != "OK" and critical:
            missing_critical = True
    print("=" * 62)
    if missing_critical:
        print("缺少必需依赖：python -m pip install -r requirements.txt")
        return 1
    print("必需依赖齐全。Windows + Word 可获得完整的目录更新与 PDF 导出能力。")
    return 0


def cmd_demo(args):
    import build_docx
    import make_chart

    out = os.path.abspath(args.out)
    demo_dir = os.path.join(REPO, "examples", "demo")
    os.makedirs(out, exist_ok=True)

    charts = sorted(glob.glob(os.path.join(demo_dir, "charts", "*.json")))
    for spec in charts:
        with open(spec, encoding="utf-8") as fh:
            name = json.load(fh).get("output") or os.path.splitext(os.path.basename(spec))[0]
        target = os.path.join(demo_dir, "figs", name + ".png")
        print("图表 → %s" % os.path.relpath(target, REPO))
        make_chart.main([spec, "-o", target])

    docx = os.path.join(out, "demo_report.docx")
    print("文档 → %s" % docx)
    build_docx.main([os.path.join(demo_dir, "content_model.json"), "-o", docx])

    pdf = os.path.splitext(docx)[0] + ".pdf"
    print("更新域/目录并导出 PDF（无 Word 时会跳过）…")
    rc = cmd_post(argparse.Namespace(docx=docx, pdf=pdf, no_pdf=False,
                                     no_view_reset=False))
    if rc != 0:
        print("  跳过：未取得 Word/LibreOffice，docx 仍可正常使用。")

    print("质检：")
    qa_rc = cmd_qa(argparse.Namespace(docx=docx, pdf=pdf if os.path.exists(pdf) else None,
                                      spec=None, strict=False))
    if os.path.exists(pdf):
        cmd_preview(argparse.Namespace(pdf=pdf, out=os.path.join(out, "preview"),
                                       dpi=110, pages=None))
    return qa_rc


# ------------------------------------------------------------------ 入口

def build_parser():
    ap = argparse.ArgumentParser(
        prog="apd", description="academic-paper-docx 命令行",
        formatter_class=argparse.RawDescriptionHelpFormatter, epilog=__doc__)
    ap.add_argument("--version", action="version", version="%(prog)s " + __version__)
    sub = ap.add_subparsers(dest="command", required=True)

    p = sub.add_parser("build", help="内容模型 JSON → docx")
    p.add_argument("model")
    p.add_argument("-o", "--output")
    p.add_argument("--spec", help="配色/字号覆盖 JSON")
    p.set_defaults(func=cmd_build)

    p = sub.add_parser("post", help="更新域/目录 + 导出 PDF + 固定页面视图")
    p.add_argument("docx")
    p.add_argument("-p", "--pdf")
    p.add_argument("--no-pdf", action="store_true")
    p.add_argument("--no-view-reset", action="store_true")
    p.set_defaults(func=cmd_post)

    p = sub.add_parser("qa", help="按规范自动质检")
    p.add_argument("docx")
    p.add_argument("--pdf")
    p.add_argument("--spec")
    p.add_argument("--strict", action="store_true")
    p.set_defaults(func=cmd_qa)

    p = sub.add_parser("preview", help="PDF 逐页转 PNG + 体检")
    p.add_argument("pdf")
    p.add_argument("-o", "--out", default=None)
    p.add_argument("--dpi", type=int, default=110)
    p.add_argument("--pages")
    p.set_defaults(func=cmd_preview)


    p = sub.add_parser("figure", help="按国标/期刊惯例渲染插图（毫米版面 + 色板校验）")
    p.add_argument("spec")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--grayscale", action="store_true", help="同时输出灰度版自检")
    p.set_defaults(func=cmd_figure)

    p = sub.add_parser("chart", help="旧名，等价于 figure")
    p.add_argument("spec")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--grayscale", action="store_true")
    p.set_defaults(func=cmd_figure)

    p = sub.add_parser("tone", help="检测正文的「AI 味」（只报警，不改稿）")
    src = p.add_mutually_exclusive_group(required=True)
    src.add_argument("--docx", help="已生成的 .docx")
    src.add_argument("--json", help="内容模型 content.json")
    src.add_argument("--text", help="纯文本 / Markdown")
    p.add_argument("--max-score", type=int, default=None,
                   help="超过该分数时退出码 1（CI 门禁用）")
    p.add_argument("--json-out", action="store_true", help="机器可读输出")
    p.set_defaults(func=cmd_tone)

    p = sub.add_parser("doctor", help="环境自检")
    p.set_defaults(func=cmd_doctor)

    p = sub.add_parser("demo", help="一键生成示例（图表 → docx → PDF → 质检 → 预览）")
    p.add_argument("-o", "--out", default=os.path.join("build", "demo"))
    p.set_defaults(func=cmd_demo)

    return ap


def main(argv=None):
    _console.setup()
    args = build_parser().parse_args(argv)
    try:
        return args.func(args)
    except KeyboardInterrupt:
        print("\n已中断", file=sys.stderr)
        return 130


if __name__ == "__main__":
    sys.exit(main())
