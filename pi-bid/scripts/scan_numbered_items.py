#!/usr/bin/env python3
"""
扫描 markdown 中的编号列表项（如 `1.xxx`、`2.xxx`），按 ##### 小节分组统计。
用于把控"编号列表项是否过多/是否需要表格化"，以及导出前自查。
用法: python3 scan_numbered_items.py <文件路径> [--verbose]
"""
import re, sys

if len(sys.argv) < 2 or sys.argv[1].startswith('-'):
    print("用法: python3 scan_numbered_items.py <文件路径> [--verbose]")
    sys.exit(1)
MD_PATH = sys.argv[1]
VERBOSE = '--verbose' in sys.argv

with open(MD_PATH) as f:
    lines = f.readlines()

# Match: line starts with digit+dot followed by non-digit text (with or without space)
PATTERN = re.compile(r'^\d+\.[^\d]')

current_module = ''
module_counts = {}
all_items = []

for i, l in enumerate(lines):
    if l.startswith('##### '):
        current_module = l.strip()
        module_counts.setdefault(current_module, 0)
    if PATTERN.match(l):
        module_counts[current_module] = module_counts.get(current_module, 0) + 1
        all_items.append((i+1, current_module, l.strip()))

total = 0
print(f"=== 编号列表项扫描: {MD_PATH} ===\n")
for mod, count in sorted(module_counts.items(), key=lambda x: -x[1]):
    if count > 0:
        print(f"  {count:4d}  {mod}")
        total += count

print(f"\n  总计: {total} 处\n")

if VERBOSE and total > 0:
    print("=== 详细列表 (前50条) ===")
    for lineno, mod, txt in all_items[:50]:
        print(f"  L{lineno:5d} [{mod[:30]}]: {txt[:100]}")
