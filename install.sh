#!/usr/bin/env bash
# 把 academic-paper-docx 安装到 agent 的 skills 目录。
#
# 用法：
#   ./install.sh                 # 安装到 ~/.dsh/skills
#   ./install.sh codex           # 安装到 ~/.codex/skills
#   ./install.sh claude          # 安装到 ~/.claude/skills
#   ./install.sh project         # 安装到 ./.dsh/skills（项目级）
#   ./install.sh custom /path/to/skills
#   LINK=1 ./install.sh          # 用符号链接（改仓库代码立即生效）
#   FORCE=1 ./install.sh         # 覆盖已存在的目标
set -euo pipefail

NAME="academic-paper-docx"
SRC="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
TARGET="${1:-dsh}"
CUSTOM="${2:-}"

[ -f "$SRC/SKILL.md" ] || { echo "install.sh 必须放在仓库根目录（找不到 SKILL.md）" >&2; exit 1; }

case "$TARGET" in
  dsh)     ROOT="$HOME/.dsh/skills" ;;
  codex)   ROOT="$HOME/.codex/skills" ;;
  claude)  ROOT="$HOME/.claude/skills" ;;
  project) ROOT="$(pwd)/.dsh/skills" ;;
  custom)
    [ -n "$CUSTOM" ] || { echo "custom 需要第二个参数：目标 skills 根目录" >&2; exit 2; }
    ROOT="$CUSTOM" ;;
  *) echo "未知目标：$TARGET（dsh|codex|claude|project|custom）" >&2; exit 2 ;;
esac

DEST="$ROOT/$NAME"
echo "源目录：$SRC"
echo "目标：  $DEST"

mkdir -p "$ROOT"
if [ -e "$DEST" ]; then
  if [ "${FORCE:-0}" = "1" ]; then
    rm -rf "$DEST"
  else
    echo "目标已存在。用 FORCE=1 ./install.sh 覆盖，或手动删除：$DEST" >&2
    exit 3
  fi
fi

if [ "${LINK:-0}" = "1" ]; then
  ln -s "$SRC" "$DEST"
  echo "已创建符号链接（改仓库代码立即生效）。"
else
  if command -v rsync >/dev/null 2>&1; then
    rsync -a --exclude .git --exclude __pycache__ --exclude .pytest_cache \
      --exclude .ruff_cache --exclude .github --exclude build --exclude '*.pyc' \
      "$SRC/" "$DEST/"
  else
    mkdir -p "$DEST"
    (cd "$SRC" && tar --exclude .git --exclude __pycache__ --exclude .pytest_cache \
      --exclude .ruff_cache --exclude .github --exclude build -cf - .) | (cd "$DEST" && tar -xf -)
  fi
  echo "已复制。"
fi

cat <<EOF

下一步：
  1) 安装依赖： python3 -m pip install -r "$DEST/requirements.txt"
  2) 自检：     python3 "$DEST/scripts/apd.py" doctor
  3) 生成示例： python3 "$DEST/scripts/apd.py" demo -o build/demo

重启 agent（或新开会话）后技能即可被发现。
EOF
