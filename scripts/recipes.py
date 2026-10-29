#!/usr/bin/env python
"""配方工具的 CLI：列出配方、拷出完整源码、跑图、核对数据契约与保真改动。

工作方式：配方 = 完整可运行源码；拷到工作区后适配，再回检数据契约与设计改动。

  1. **选择** `apd recipe list` / `apd recipe info <id>`：按数据语义选配方，看数据契约与保真要点
  2. **适配** `apd recipe new <id> --out work/figs`：把**完整源码**拷成
     `work/figs/gen_fig_<id>.py`（连 `_kit.py` 一起），在真实文件上改数据、标签与布局
  3. **跑图** `apd recipe run work/figs/gen_fig_<id>.py data.csv out.png [width_mm]`
  4. **核对** `apd recipe check work/figs/gen_fig_<id>.py --data data.csv`
     —— 校验数据契约 + 报告脚本相对原配方改了哪些行（设计块改动会单独提示）

配方脚本自带 PNG 元数据（设计宽度 / 基准字号），docx 端据此核算"插入后图内文字还剩几 pt"。
"""

from __future__ import annotations

import argparse
import difflib
import json
import os
import re
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
RECIPES = os.path.join(REPO, "recipes")
KIT = os.path.join(RECIPES, "_kit.py")

# 设计块：这些行被改动时要提示"先确认是有意的"
# 科研画图的经典错误：check 时静态扫描，命中就明确提示（不阻断）
PITFALLS = [
    (r"\bplt\.pie\(|\bax\.pie\(", "P-饼图",
     "饼图判角度差长度约 3 倍：改用横向柱状并按值排序"),
    (r"\btwinx\(\)|\btwiny\(\)|secondary_yaxis", "P-双 Y 轴",
     "双 Y 轴会制造视觉相关：拆上下子图共享 x，或标准化共轴"),
    (r"cmap\s*=\s*['\"](?:jet|rainbow|nipy_spectral|hsv)['\"]", "P-彩虹色谱",
     "jet/rainbow 感知不均匀会凭空造峰：改 cividis/viridis/RdBu_r"),
    (r"projection\s*=\s*['\"]3d['\"]", "P-3D 图",
     "3D 视角扭曲数值：改 2D 柱或热图"),
]
MEANBAR_RE = re.compile(r"ax\.bar\(|\.bar\(")
SCATTER_RE = re.compile(r"\.plot\([^)]*linestyle\s*=\s*['\"]none|scatter\(|stripplot")

DESIGN_MARKERS = ("SERIES", "LINESTYLES", "MARKERS", "STYLE", "rcParams",
                  "fontsize", "linewidth", "cmap", "set_ylim", "set_xlim",
                  "figsize", "dpi", "savefig")


def _ids():
    if not os.path.isdir(RECIPES):
        return []
    return sorted(d for d in os.listdir(RECIPES)
                  if os.path.isfile(os.path.join(RECIPES, d, "recipe.py")))


def _meta(rid):
    path = os.path.join(RECIPES, rid, "meta.json")
    if not os.path.exists(path):
        raise SystemExit("配方不存在：%s（可用：%s）" % (rid, "、".join(_ids())))
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def cmd_list(args):
    rows = [(_meta(r)["id"], _meta(r)["title"], _meta(r)["when"]) for r in _ids()]
    width = max(len(r[0]) for r in rows) + 2
    print("配方目录（%d 个）" % len(rows))
    print("-" * 78)
    for rid, title, when in rows:
        print("%-*s %s" % (width, rid, title))
        print("%-*s %s" % (width, "", when))
    print("-" * 78)
    print("用法：apd recipe info <id> ｜ apd recipe new <id> --out <目录>")
    return 0


def cmd_info(args):
    meta = _meta(args.id)
    print("配方：%s —— %s" % (meta["id"], meta["title"]))
    print("适用：%s" % meta["when"])
    print("数据契约：")
    for key, val in meta["data_contract"].items():
        print("  %-18s %s" % (key, val))
    print("保真要点（适配时不要丢）：")
    for item in meta.get("fidelity", []):
        print("  · %s" % item)
    print("交付前自查：")
    for item in meta.get("verify", []):
        print("  - %s" % item)
    print("源码：%s" % os.path.relpath(
        os.path.join(RECIPES, meta["id"], "recipe.py"), REPO))
    return 0


def cmd_new(args):
    rid = args.id
    _meta(rid)
    out_dir = os.path.abspath(args.out)
    os.makedirs(out_dir, exist_ok=True)
    script = os.path.join(out_dir, "gen_fig_%s.py" % rid.replace("-", "_"))
    if os.path.exists(script) and not args.force:
        print("已存在：%s（加 --force 覆盖）" % script, file=sys.stderr)
        return 2
    shutil.copyfile(os.path.join(RECIPES, rid, "recipe.py"), script)
    shutil.copyfile(KIT, os.path.join(out_dir, "_kit.py"))
    print("已拷出完整源码：%s" % script)
    print("公共件：%s" % os.path.join(out_dir, "_kit.py"))
    print("下一步：在 %s 里把 config() 与数据列名改成你的；然后" % os.path.basename(script))
    print("  apd recipe run %s <data.csv> <out.png> [width_mm]" % script)
    return 0


def _child_env():
    """子进程统一 UTF-8：Windows 默认 cp1252 时中文打印会抛 UnicodeEncodeError。"""
    env = dict(os.environ)
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def cmd_run(args):
    script = os.path.abspath(args.script)
    if not os.path.exists(script):
        print("找不到脚本：%s" % script, file=sys.stderr)
        return 2
    argv = [sys.executable, script, os.path.abspath(args.data),
            os.path.abspath(args.output)]
    if args.width_mm:
        argv.append(str(args.width_mm))
    return subprocess.call(argv, env=_child_env())


def cmd_check(args):
    rid = args.recipe or _guess_recipe(os.path.abspath(args.script))
    meta = _meta(rid)
    ok = True

    if args.data:
        try:
            header, rows = _read_csv(args.data)
        except Exception as exc:                     # pragma: no cover
            print("读数据失败：%s" % exc, file=sys.stderr)
            return 2
        need = [c for c in meta["data_contract"]["required_columns"]
                if not c.startswith("<")]
        missing = [c for c in need if c not in header]
        if missing:
            ok = False
            print("✗ 数据缺列：%s" % "、".join(missing))
        min_rows = int(meta["data_contract"].get("min_rows") or 0)
        if len(rows) < min_rows:
            ok = False
            print("✗ 行数不足：%d < %d" % (len(rows), min_rows))
        if ok:
            print("✓ 数据契约通过（%d 行 × %d 列）" % (len(rows), len(header)))

    if os.path.exists(args.script):
        _scan_pitfalls(args.script, rid)

    base_path = os.path.join(RECIPES, rid, "recipe.py")
    if os.path.exists(args.script):
        changed, design = _diff(base_path, args.script)
        if not changed:
            print("· 源码与原配方一致（可能还没适配数据）")
        else:
            print("· 相对原配方改了 %d 行；其中涉及设计参数 %d 行"
                  % (changed, len(design)))
            for line in design[:6]:
                print("    %s" % line)
            if design:
                print("  提示：改字号/线宽/配色/图幅前先确认是有意的适配，"
                      "保真要点见 apd recipe info %s" % rid)
    return 0 if ok else 1


def _scan_pitfalls(script, rid):
    """静态扫描经典错误图型；命中只提示，不阻断。"""
    with open(script, encoding="utf-8") as fh:
        text = fh.read()
    hits = []
    for pattern, name, advice in PITFALLS:
        if re.search(pattern, text):
            hits.append((name, advice))
    if MEANBAR_RE.search(text) and not SCATTER_RE.search(text):
        hits.append(("P-均值柱掩盖分布",
                     "柱状只有高度与一根误差棒时，n<10 请改 dotplot/箱线并叠加每个点"))
    for name, advice in hits:
        print("! [%s] %s" % (name, advice))
    if hits:
        print("  （以上为建议，不阻断；见 references/nature-style.md §3）")


def _guess_recipe(script):
    # 仓库内直接指向 recipes/<id>/recipe.py 时，用父目录名当配方名
    parent = os.path.basename(os.path.dirname(os.path.abspath(script)))
    if parent in _ids():
        return parent
    name = os.path.basename(script)
    m = re.match(r"gen_fig_(.+)\.py$", name)
    if m:
        cand = m.group(1).replace("_", "-")
        if cand in _ids():
            return cand
    raise SystemExit("无法从文件名推断配方；用 --recipe 指定（可用：%s）"
                     % "、".join(_ids()))


def _read_csv(path):
    import csv

    with open(path, encoding="utf-8-sig", newline="") as fh:
        rows = [r for r in csv.reader(fh) if r]
    return [h.strip() for h in rows[0]], rows[1:]


def _diff(base_path, script_path):
    with open(base_path, encoding="utf-8") as fh:
        base = fh.readlines()
    with open(script_path, encoding="utf-8") as fh:
        cur = fh.readlines()
    changed = 0
    design = []
    for line in difflib.unified_diff(base, cur, n=0):
        if not line.startswith(("+", "-")) or line.startswith(("+++", "---")):
            continue
        changed += 1
        body = line[1:].strip()
        if any(marker in body for marker in DESIGN_MARKERS):
            design.append(body[:72])
    return changed, design


def main(argv=None):
    ap = argparse.ArgumentParser(description="配方工具：列配方 / 拷源码 / 跑图 / 核对")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("list", help="列出全部配方")
    p.set_defaults(func=cmd_list)

    p = sub.add_parser("info", help="看配方说明、数据契约与保真要点")
    p.add_argument("id")
    p.set_defaults(func=cmd_info)

    p = sub.add_parser("new", help="把配方完整源码拷到工作区")
    p.add_argument("id")
    p.add_argument("--out", required=True, help="工作目录（脚本会放这里）")
    p.add_argument("--force", action="store_true")
    p.set_defaults(func=cmd_new)

    p = sub.add_parser("run", help="跑适配后的脚本")
    p.add_argument("script")
    p.add_argument("data")
    p.add_argument("output")
    p.add_argument("width_mm", nargs="?", type=float, default=None)
    p.set_defaults(func=cmd_run)

    p = sub.add_parser("check", help="核对数据契约与相对原配方的改动")
    p.add_argument("script")
    p.add_argument("--data")
    p.add_argument("--recipe")
    p.set_defaults(func=cmd_check)

    args = ap.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
