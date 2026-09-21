#!/usr/bin/env python3
"""
后处理: 修复 minimax-docx 导出的两个问题
1. 重命名 styles2.xml / numbering2.xml → styles.xml / numbering.xml
2. 修复 Heading6 numPr 编号
3. 更新 Content_Types.xml

用法: python3 fix_docx_post.py 输入.docx [输出.docx]
"""

import sys, os, zipfile, shutil
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'

def fix_docx(input_path, output_path=None):
    if output_path is None:
        output_path = input_path

    tmp = output_path + '.ztmp'
    with zipfile.ZipFile(input_path, 'r') as zin:
        items = {}
        for name in zin.namelist():
            items[name] = zin.read(name)

        changes = []

        # 1. Rename styles2.xml → styles.xml
        if 'word/styles2.xml' in items and 'word/styles.xml' not in items:
            items['word/styles.xml'] = items.pop('word/styles2.xml')
            changes.append('styles2.xml → styles.xml')

        # 2. Rename numbering2.xml → numbering.xml
        if 'word/numbering2.xml' in items and 'word/numbering.xml' not in items:
            items['word/numbering.xml'] = items.pop('word/numbering2.xml')
            changes.append('numbering2.xml → numbering.xml')

        # 3. Rename theme2.xml → theme.xml (if theme.xml missing)
        if 'word/theme/theme2.xml' in items:
            theme1_exists = any(k.startswith('word/theme/theme') and not k.endswith('2.xml') for k in items)
            if not theme1_exists:
                items['word/theme/theme.xml'] = items.pop('word/theme/theme2.xml')
                changes.append('theme2.xml → theme.xml')

        # 4. Update Content_Types.xml
        ct_key = '[Content_Types].xml'
        if ct_key in items:
            ct_root = etree.fromstring(items[ct_key])
            ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'

            # Remove any Override for styles2/numbering2
            for ov in list(ct_root.findall(f'{{{ct_ns}}}Override')):
                pn = ov.get('PartName', '')
                if 'styles2' in pn or 'numbering2' in pn or 'theme2' in pn:
                    ct_root.remove(ov)
                    changes.append(f'removed Override for {pn}')

            # Add correct Overrides if missing
            existing_parts = {ov.get('PartName') for ov in ct_root.findall(f'{{{ct_ns}}}Override')}
            needed = {
                '/word/styles.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
                '/word/numbering.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml',
            }
            for part_name, content_type in needed.items():
                if part_name in existing_parts:
                    continue
                ov = etree.SubElement(ct_root, f'{{{ct_ns}}}Override')
                ov.set('PartName', part_name)
                ov.set('ContentType', content_type)
                changes.append(f'added Override: {part_name}')

            items[ct_key] = etree.tostring(ct_root, xml_declaration=True, encoding='UTF-8', standalone=True)

        # 5. Fix document.xml.rels — rename styles2 → styles, numbering2 → numbering
        rels_key = 'word/_rels/document.xml.rels'
        if rels_key in items:
            rels_root = etree.fromstring(items[rels_key])
            rels_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'

            for rel in list(rels_root):
                target = rel.get('Target', '')
                if target == 'styles2.xml':
                    rel.set('Target', 'styles.xml')
                    changes.append('rels: styles2.xml → styles.xml')
                elif target == 'numbering2.xml':
                    rel.set('Target', 'numbering.xml')
                    changes.append('rels: numbering2.xml → numbering.xml')
                elif target == 'theme/theme2.xml':
                    rel.set('Target', 'theme/theme.xml')
                    changes.append('rels: theme2.xml → theme.xml')

            items[rels_key] = etree.tostring(rels_root, xml_declaration=True, encoding='UTF-8', standalone=True)

        # 6. Ensure H6 headings have numPr if style doesn't provide it
        doc_key = 'word/document.xml'
        if doc_key in items:
            doc_root = etree.fromstring(items[doc_key])
            body = doc_root.find(f'{{{W}}}body')
            h6_fixed = 0
            for p in body.iter(f'{{{W}}}p'):
                pPr = p.find(f'{{{W}}}pPr')
                if pPr is None:
                    continue
                pStyle = pPr.find(f'{{{W}}}pStyle')
                if pStyle is None:
                    continue
                if pStyle.get(f'{{{W}}}val') == 'Heading6':
                    numPr = pPr.find(f'{{{W}}}numPr')
                    if numPr is None:
                        numPr = etree.Element(f'{{{W}}}numPr')
                        # Insert numPr right after pStyle
                        pStyle_idx = list(pPr).index(pStyle)
                        pPr.insert(pStyle_idx + 1, numPr)
                        h6_fixed += 1
            if h6_fixed > 0:
                items[doc_key] = etree.tostring(doc_root, xml_declaration=True, encoding='UTF-8', standalone=True)
                changes.append(f'fixed {h6_fixed} H6 numPr')

    # Write output
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in items.items():
            zout.writestr(name, data)
    os.replace(tmp, output_path)

    print(f"✅ Post-fix: {len(changes)} changes")
    for c in changes:
        print(f"   {c}")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python3 fix_docx_post.py 输入.docx [输出.docx]")
        sys.exit(1)
    fix_docx(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
