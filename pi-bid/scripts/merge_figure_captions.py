#!/usr/bin/env python3
"""
合并图标题到 markdown alt text
将:  ![](docs/imageN.png)\n\n标题文本
改为: ![标题文本](docs/imageN.png)

用法: python3 merge_figure_captions.py <input.md> [output.md]   # 缺省原地修改
"""

import re, sys

def is_caption_line(text):
    """判断是否是图片标题"""
    if not text or len(text) > 100:
        return False
    # 排除表格标题
    if re.match(r'^表\s', text):
        return False
    # 排除 heading
    if re.match(r'^#{1,6}\s', text):
        return False
    # 排除格式说明
    if re.match(r'^(编号格式|页面设置|正文格式|一级标题|二级标题|标题编号|页码|图.*编号|表格|成文后|其他要求|正式成文)', text):
        return False
    # 排除代码块/引用的内容
    if text.startswith('```') or text.startswith('>'):
        return False
    # 排除纯文件名（imageN.png 类）
    if re.match(r'^image\d+\.(png|jpg|jpeg)$', text, re.I):
        return False
    return True

def merge_captions(text):
    """处理整个 markdown 文本"""
    lines = text.split('\n')
    result = []
    i = 0
    merged = 0

    while i < len(lines):
        line = lines[i]

        # 检测空 alt 的图片
        m = re.match(r'^!\[\]\((docs/.*?)\)$', line)
        if m:
            img_path = m.group(1)
            # 跳过后续空行
            j = i + 1
            while j < len(lines) and not lines[j].strip():
                j += 1

            # 检查下一行非空内容是否是标题
            if j < len(lines):
                caption = lines[j].strip()
                if is_caption_line(caption):
                    # 合并为带 alt 的图片
                    result.append(f'![{caption}]({img_path})')
                    merged += 1
                    i = j + 1
                    continue

        result.append(line)
        i += 1

    return '\n'.join(result), merged

def main():
    if len(sys.argv) < 2:
        print("用法: python3 merge_figure_captions.py <input.md> [output.md]")
        sys.exit(1)
    input_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else input_path

    with open(input_path, 'r', encoding='utf-8') as f:
        text = f.read()

    new_text, merged = merge_captions(text)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(new_text)

    print(f"✅ Merged {merged} figure captions → alt text")
    print(f"   {input_path} → {output_path}")


if __name__ == '__main__':
    main()
