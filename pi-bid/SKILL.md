---
name: pi-bid
description: 标书/投标技术方案写作全流程——评分标准拆解、撰写逻辑管理、八步选型论证、口径一致性迭代、md转docx导出、标书配图管线（pptx→png、AI生图还原为可编辑PPT）
argument-hint: "<标书部分/章节主题>"
---

# pi-bid

标书（投标技术方案）写作工作流。四层组织：

```
pi-bid/
├── SKILL.md            # 流程层（本文件）：四阶段流程
├── checklist.md        # 规则层：行文检查清单
├── templates/          # 撰写逻辑模板（口径决策表/结构决策记录）
├── scripts/            # 工具层：检查与导出脚本
└── examples/           # 知识层：虚构项目实例参考（脱敏样例）
```

核心思想：**口径先于文字**。多轮修改最大的风险是口径漂移（指标值、模块名、结构前后不一致），所以所有口径落在"撰写逻辑"文件里，改口径先改它再改正文。

## 前置环境（阶段四导出用）

- **minimax-docx 已 vendor 内置**（`vendor/minimax-docx/`，MIT 协议，含 LICENSE），并附带预编译产物——无需外部安装、无需设 `MINIMAX_DOCX_CLI`；首次使用且产物缺失时 `bid_export.sh` 会自动 `dotnet build -c Release`。
- **需要 .NET 9+ SDK**（编译/运行 minimax-docx 用；缺装时脚本会给出安装指引）。
- **TPL 需用户自备**：贵方标书 docx 模板（含页眉页脚/样式），导出时以 `TPL=/path/模板.docx` 环境变量指定。
- 其他依赖：pandoc、python3 + python-docx + lxml；配图管线另需 soffice、pdftoppm。

## 阶段一：定盘

1. 拆解评分标准与需求指标：把评分原话逐条摘出，每条标注"用什么章节/什么证据承接"；指标数值整理成表，作为口径唯一来源。
2. 初始化撰写逻辑：复制 `templates/撰写逻辑模板.md` 到项目目录（如 `<章节>_撰写逻辑.md`），填口径决策表（指标口径/命名口径/结构口径）与总览表。
3. 通读 `checklist.md` 立规则，尤其记住：禁前向引用、指标加粗、落选方案写"场景不适配"、禁人造评分表。

## 阶段二：撰写

- **技术路线选型**：按八步逻辑写（约束先行→维度可溯→调研全景→定性对比→约束筛除→推荐+理由→指标挂钩→局限化解），详见 `examples/技术路线选型写作逻辑.md`。
- **创新小节**："总—分—合"范式——总起预告（本部分创新点 N 项）→ 创新点逐条独立段【是什么/突破 + 做法 + 效果挂指标】→ 前沿分析段收口。
- **工作原理**：三段式——原理文字 → 原理图（"如下图所示"引导）→ 工作流程（阶段链或时序）。
- **每个模块四件套**：简述（定位+职责 2-3 句）→ 组成图 → 需求总结（逐项承接需求）→ 子模块表（子模块|用途|研发状态）。
- 行文规则全程对照 `checklist.md`；完整实例见 `examples/总体架构_撰写逻辑.md`。

## 阶段三：迭代

1. 按 `checklist.md` 逐条自查。
2. 跑检查脚本：
   - `python3 scripts/check_sensitive_words.py <目录> -r [--wordlist 项目词表.txt]` —— 敏感词扫描
   - `python3 scripts/scan_numbered_items.py <文件.md> [--verbose]` —— 编号列表项分布
   - `python3 scripts/merge_figure_captions.py <文件.md>` —— 图标题合并进 alt text（导出前）
3. 口径库更新：任何指标/命名/结构变化，先更新撰写逻辑文件的口径决策表与结构决策记录（写明日期+理由+推翻的旧决策），再改正文。
4. 领导意见处理：**先定位口径再动笔**——判断意见动的是指标口径、命名口径还是结构口径，改口径表 → 评估影响面（哪些章节引用该口径）→ 统一改正文。

## 阶段四：交付

1. 导出 docx：
   ```bash
   TPL=/path/模板.docx scripts/bid_export.sh input.md [output.docx]
   ```
   流程：pandoc → minimax-docx 套模板 → 表格式化+SEQ编号 → post-fix → 图格式化。若表题较多，可用 `BID_TABLE_NAMES_FILE` 指定表名清单（每行一个，按文档顺序）。
2. 配图管线：
   - 常规：python-pptx 脚本生成 pptx → `scripts/fig_pptx2png.sh in.pptx out.png`（soffice 转 pdf → pdftoppm 150dpi → PIL 裁白边）。
   - **AI 生图必须还原为可编辑 PPT**：生图（ChatGPT 等）→ python-pptx 按图重建（图标用语义近似库替代、纯色代渐变）→ `fig_pptx2png.sh` 渲染自检 → 与 AI 原图比对布局 → 按同文件名部署覆盖（正文引用不动）。
3. 交付前最后一遍敏感词扫描 + 口径抽查（抽 3 处指标核对撰写逻辑文件）。

## 脚本一览

| 脚本 | 功能 |
|---|---|
| `scripts/bid_export.sh` | md → docx 五步导出管线 |
| `scripts/ensure_env.sh` | 导出前置自检：解析/自动编译 minimax-docx CLI、校验 TPL |
| `scripts/fig_pptx2png.sh` | pptx → pdf → 150dpi png → 裁白边 |
| `scripts/bid_format_tables.py` | 表格格式化 + SEQ 自动编号 |
| `scripts/bid_format_figures.py` | 插图格式化 + 图题编号 |
| `scripts/fix_docx_post.py` | 修复 minimax-docx 导出残留（styles/numbering 重命名、H6 编号、rels） |
| `scripts/merge_figure_captions.py` | 图标题合并进 markdown alt text |
| `scripts/check_sensitive_words.py` | 敏感词扫描/替换（支持自定义词表） |
| `scripts/scan_numbered_items.py` | 编号列表项扫描统计 |
