#!/bin/bash
# ============ 标书导出工作流 ============
# 流程: MD → pandoc → raw.docx → minimax-docx(套模板) → 表格式化 → post-fix → 图格式化 → 成品.docx
#
# 用法: TPL=/path/模板.docx ./bid_export.sh input.md [output.docx]
#
# 环境变量:
#   TPL             docx 模板路径（必填，缺省报错；用户自备的标书模板）
#   MINIMAX_DOCX_CLI  minimax-docx CLI 的 dll 路径（可选）。缺省时 ensure_env.sh
#                     自动解析本 skill vendor/minimax-docx/ 下的编译产物，
#                     首次使用且有 dotnet 时会自动 dotnet build -c Release
#
# 其他依赖: pandoc、dotnet runtime、python3 + python-docx + lxml
# ==========================================
set -e
DIR="$(cd "$(dirname "$0")" && pwd)"
FMTS="$DIR/bid_format_tables.py"
FIG="$DIR/bid_format_figures.py"
POSTFIX="$DIR/fix_docx_post.py"

IN="${1:-}"; OUT="${2:-${IN%.md}.docx}"
[ -f "$IN"  ] || { echo "用法: TPL=/path/模板.docx $0 input.md [output.docx]"; exit 1; }

# 环境自检: 解析/编译 minimax-docx CLI（导出 $CLI）、校验 TPL（导出 $TPL）
source "$DIR/ensure_env.sh"

RAW="/tmp/bid_raw_$$.docx"
TMP="/tmp/bid_tmp_$$.docx"

# Step 1: pandoc MD → DOCX
echo "[1/5] pandoc..."
pandoc "$IN" -o "$RAW" --from markdown --to docx --resource-path="$(dirname "$IN")"

# Step 2: minimax-docx 套模板
echo "[2/5] minimax-docx apply-template..."
dotnet "$CLI" apply-template --input "$RAW" --template "$TPL" --output "$TMP" 2>/dev/null

# Step 3: 表格式化 + SEQ编号
echo "[3/5] table format..."
python3 "$FMTS" "$TMP" "$TMP"

# Step 4: 后处理 (styles/numbering 重命名 + H6 编号 + rels 修复)
echo "[4/5] post-fix..."
python3 "$POSTFIX" "$TMP" "$TMP"

# Step 5: 插图格式化
echo "[5/5] figure format..."
python3 "$FIG" "$TMP" "$OUT"

rm -f "$RAW" "$TMP"
echo "✅ $OUT"
