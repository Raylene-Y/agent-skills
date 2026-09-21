#!/bin/bash
# ============ 标书配图管线: pptx → png ============
# 流程: input.pptx → soffice --headless 转 pdf → pdftoppm -r 150 转 png → PIL 裁白边
#
# 用法: ./fig_pptx2png.sh input.pptx output.png
#
# 依赖:
#   - libreoffice (soffice)
#   - poppler-utils (pdftoppm)
#   - python3-pil (Pillow, 裁白边)
# =================================================
set -e

IN="${1:-}"
OUT="${2:-}"
[ -f "$IN" ] && [ -n "$OUT" ] || { echo "用法: $0 input.pptx output.png"; exit 1; }

WORKDIR="$(mktemp -d /tmp/fig_pptx2png_XXXXXX)"
trap 'rm -rf "$WORKDIR"' EXIT

# Step 1: pptx → pdf
echo "[1/3] soffice pptx → pdf..."
soffice --headless --convert-to pdf --outdir "$WORKDIR" "$IN" >/dev/null
PDF="$WORKDIR/$(basename "${IN%.pptx}").pdf"
[ -f "$PDF" ] || { echo "错误: pdf 转换失败: $PDF"; exit 1; }

# Step 2: pdf → png (150 dpi)
echo "[2/3] pdftoppm pdf → png (150dpi)..."
pdftoppm -r 150 -png -singlefile "$PDF" "$WORKDIR/raw"
[ -f "$WORKDIR/raw.png" ] || { echo "错误: png 转换失败"; exit 1; }

# Step 3: PIL 裁白边
echo "[3/3] 裁白边..."
python3 - "$WORKDIR/raw.png" "$OUT" <<'EOF'
import sys
from PIL import Image, ImageChops

src, dst = sys.argv[1], sys.argv[2]
img = Image.open(src).convert('RGB')
bg = Image.new('RGB', img.size, (255, 255, 255))
diff = ImageChops.difference(img, bg)
bbox = diff.getbbox()
if bbox:
    pad = 10  # 留白边距(px)，避免贴边
    l, t, r, b = bbox
    bbox = (max(0, l - pad), max(0, t - pad), min(img.size[0], r + pad), min(img.size[1], b + pad))
    img.crop(bbox).save(dst)
else:
    img.save(dst)
print(f"✅ {dst} ({img.size[0]}x{img.size[1]} -> {bbox})")
EOF
