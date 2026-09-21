#!/usr/bin/env python3
"""
标书敏感词扫描脚本 - 检查真实单位名/联系方式/涉密词泄漏。
用法:
  python3 check_sensitive_words.py <目录或文件> [--wordlist 自定义词表.txt] [--recursive] [--fix]

自定义词表格式（每行一条）: 正则|替换词    （|替换词 可省略，默认为【敏感词】）
内置规则只含通用条目；项目相关的真实单位名/人名务必通过 --wordlist 补充。
"""
import os
import sys
import re
import argparse

SENSITIVE_WORDS = [
    # 间接暴露词
    (r'采购人', '甲方'),
    (r'投标人', '乙方'),
    (r'我司', '【我方】'),
    (r'我方', '【我方】'),
    (r'本公司', '【我方】'),
    # 联系人/地址/电话
    (r'(手机|电话|联系).*?[：:]\s*1[3-9]\d{9}', '【联系方式】'),
    (r'(邮箱|mail|E-?mail).*?[：:]\s*\S+@\S+', '【邮箱】'),
    (r'地址.*?[：:]\s*\S+', '【地址】'),
    # 军事涉密词（根据项目实际情况用 --wordlist 调整）
    (r'(某)?部队\S{0,4}(编号|番号|代号)[：:]\s*\S+', '【军事编号】'),
    (r'涉密', '内部'),
    (r'机密', '内部'),
    (r'绝密', '内部'),
]

def load_custom_wordlist(path):
    """加载自定义敏感词表"""
    extra = []
    if path and os.path.exists(path):
        with open(path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    parts = line.split('|')
                    if len(parts) >= 2:
                        extra.append((parts[0], parts[1]))
                    else:
                        extra.append((line, '【敏感词】'))
    return extra

def scan_file(filepath, wordlist, fix=False):
    """扫描单个文件"""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
    except:
        try:
            with open(filepath, 'r', encoding='gbk') as f:
                content = f.read()
        except:
            return [], 0

    findings = []
    modified = content
    for pattern, replacement in wordlist:
        matches = list(re.finditer(pattern, content, re.IGNORECASE))
        if matches:
            findings.extend((m.start(), m.group(), pattern) for m in matches)
            if fix:
                modified = re.sub(pattern, replacement, modified, flags=re.IGNORECASE)

    if fix and modified != content:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(modified)

    return findings, len(content)

def scan_dir(dirpath, wordlist, recursive=False, fix=False):
    """扫描目录"""
    total_findings = []
    for root, dirs, files in os.walk(dirpath):
        if not recursive and root != dirpath:
            continue
        for f in files:
            if f.endswith('.md') or f.endswith('.txt'):
                path = os.path.join(root, f)
                findings, _ = scan_file(path, wordlist, fix)
                for pos, match, pat in findings:
                    total_findings.append((path, pos, match, pat))
    return total_findings

def main():
    parser = argparse.ArgumentParser(description='标书敏感词扫描')
    parser.add_argument('target', help='文件或目录')
    parser.add_argument('--wordlist', '-w', help='自定义敏感词表')
    parser.add_argument('--recursive', '-r', action='store_true', help='递归扫描')
    parser.add_argument('--fix', action='store_true', help='自动替换敏感词')
    args = parser.parse_args()

    wordlist = SENSITIVE_WORDS
    if args.wordlist:
        wordlist += load_custom_wordlist(args.wordlist)

    if os.path.isfile(args.target):
        findings, _ = scan_file(args.target, wordlist, args.fix)
        results = [(args.target, pos, match, pat) for pos, match, pat in findings]
    else:
        results = scan_dir(args.target, wordlist, args.recursive, args.fix)

    if not results:
        print("✅ 未发现敏感词")
        return

    print(f"⚠️  发现 {len(results)} 处敏感词:")
    for path, pos, match, pat in results:
        print('  {}:{} → "{}" (匹配规则: {})'.format(path, pos, match, pat))

    if args.fix:
        print(f"\n已自动修复 {len(results)} 处")


if __name__ == '__main__':
    main()
