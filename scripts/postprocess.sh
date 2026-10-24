#!/usr/bin/env bash
# 没有 Word COM 时的兜底：用 LibreOffice 更新目录（尽力）+ 导出 PDF。
#
# 用法：
#   bash postprocess.sh report.docx [report.pdf]
#
# 说明：
#   * LibreOffice 不会执行 Word 的 Fields.Update()，目录域只能靠用户打开时更新
#     （文档已写 w:updateFields / TOC dirty，Word 打开会提示或自动更新）。
#   * 导出 PDF 后仍会跑 set_view.py 固定页面视图。
set -euo pipefail

DOCX="${1:?用法: postprocess.sh <report.docx> [report.pdf]}"
PDF="${2:-${DOCX%.docx}.pdf}"
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PY="${PYTHON:-python3}"

SOFFICE=""
for cand in soffice libreoffice "/Applications/LibreOffice.app/Contents/MacOS/soffice"; do
  if command -v "$cand" >/dev/null 2>&1; then SOFFICE="$cand"; break; fi
  if [ -x "$cand" ]; then SOFFICE="$cand"; break; fi
done

if [ -z "$SOFFICE" ]; then
  echo "找不到 LibreOffice（soffice）。请手动用 Word/WPS 打开 $DOCX，"
  echo "Ctrl+A → F9 更新目录，然后另存为 PDF。" >&2
  exit 2
fi

OUTDIR="$(cd "$(dirname "$PDF")" && pwd)"
mkdir -p "$OUTDIR"
echo "导出 PDF：$PDF"
"$SOFFICE" --headless --norestore --convert-to pdf --outdir "$OUTDIR" "$DOCX" >/dev/null
BASE="$(basename "${DOCX%.docx}").pdf"
if [ "$OUTDIR/$BASE" != "$PDF" ]; then
  mv "$OUTDIR/$BASE" "$PDF"
fi

echo "固定页面视图：$DOCX"
"$PY" "$HERE/set_view.py" "$DOCX"
echo "完成。提示：目录若未展开，用 Word 打开后 Ctrl+A → F9。"
