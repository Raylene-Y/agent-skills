#!/bin/bash
# ============ 导出前置环境自检 ============
# 被 bid_export.sh source 调用。负责:
#   1. 解析 minimax-docx CLI dll 路径并导出为 $CLI
#      优先级: $MINIMAX_DOCX_CLI → vendor 编译产物
#   2. vendor 有源码无产物且有 dotnet → 自动 dotnet build -c Release
#   3. 无 dotnet → 人话报错并非零退出
#   4. 校验 $TPL（用户自备的标书模板，无法自动安装）
# ==========================================

BID_SKILL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
VENDOR_DIR="$BID_SKILL_DIR/vendor/minimax-docx"
VENDOR_PROJ="$VENDOR_DIR/dotnet/MiniMaxAIDocx.Cli"

_bid_err() {
    echo "错误: $1" >&2
    echo "  它是什么: $2" >&2
    echo "  怎么补: $3" >&2
    exit 1
}

# --- 1. CLI dll 解析 ---
CLI=""
if [ -n "${MINIMAX_DOCX_CLI:-}" ]; then
    [ -f "$MINIMAX_DOCX_CLI" ] || _bid_err \
        "MINIMAX_DOCX_CLI 指向的文件不存在: $MINIMAX_DOCX_CLI" \
        "minimax-docx 的 CLI 编译产物（.NET dll），用于把 pandoc 生成的 docx 套进标书模板" \
        "检查路径；或 unset MINIMAX_DOCX_CLI 让脚本自动使用内置 vendor 版本"
    CLI="$MINIMAX_DOCX_CLI"
else
    # vendor 编译产物（net8.0/net9.0 均可）
    CLI="$(ls "$VENDOR_PROJ"/bin/Release/net*/MiniMaxAIDocx.Cli.dll 2>/dev/null | head -1)"
fi

# --- 2. vendor 源码在、产物不在 → 自动编译 ---
if [ -z "$CLI" ] && [ -f "$VENDOR_PROJ/MiniMaxAIDocx.Cli.csproj" ]; then
    if command -v dotnet >/dev/null 2>&1; then
        echo "[ensure_env] 首次使用，自动编译内置 minimax-docx ..."
        if dotnet build "$VENDOR_PROJ" -c Release -v q --nologo 2>&1 | tail -3; then
            CLI="$(ls "$VENDOR_PROJ"/bin/Release/net*/MiniMaxAIDocx.Cli.dll 2>/dev/null | head -1)"
        fi
        [ -n "$CLI" ] || _bid_err \
            "minimax-docx 自动编译失败" \
            "skill 内置的 docx 模板套用工具（源码在 vendor/minimax-docx/）" \
            "手动执行: dotnet build $VENDOR_PROJ -c Release 查看完整错误"
        echo "[ensure_env] 编译完成: $CLI"
    else
        _bid_err \
            "未找到 dotnet，且 minimax-docx 尚无编译产物" \
            "minimax-docx 是本 skill 内置的 docx 模板套用工具（MIT，源码在 vendor/minimax-docx/），编译和运行它都需要 .NET 9+ SDK" \
            "安装 .NET SDK: Ubuntu/Debian 可用 apt install dotnet-sdk-9.0，其他系统见 https://dotnet.microsoft.com/download；装好后重跑本命令即可自动编译"
    fi
fi

[ -n "$CLI" ] || _bid_err \
    "找不到 minimax-docx CLI" \
    "docx 模板套用工具，本 skill 已将其源码 vendor 到 vendor/minimax-docx/" \
    "确认 vendor/minimax-docx 存在且有 dotnet，或手动设置 MINIMAX_DOCX_CLI=<编译产物 dll 路径>"
export CLI

# --- 3. TPL 校验（用户自备，无法自动安装） ---
TPL="${TPL:-}"
[ -n "$TPL" ] || _bid_err \
    "未设置模板环境变量 TPL" \
    "贵方标书的 docx 模板（含页眉页脚/样式定义），pandoc 转出的 docx 会套进这个模板" \
    "用法: TPL=/path/模板.docx $0 input.md [output.docx]"
[ -f "$TPL" ] || _bid_err \
    "模板文件不存在: $TPL" \
    "贵方标书的 docx 模板，需用户自备" \
    "确认路径拼写；模板通常来自招标文件附件或既往中标文档"
export TPL
