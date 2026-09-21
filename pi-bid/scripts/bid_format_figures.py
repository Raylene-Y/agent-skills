#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标书插图格式化 — 直接操作docx XML, 不经过python-docx save
用法: python3 scripts/bid_format_figures.py 输入.docx [输出.docx]

v2 改进:
  - 检测图片后紧跟的段落作为图标题
  - ImageCaption + 普通段落都处理
"""

import re, sys, os, zipfile
from lxml import etree
from copy import deepcopy
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'

def _pPr(e):
    pPr = e.find(f'{{{W}}}pPr')
    if pPr is None:
        pPr = etree.SubElement(e, f'{{{W}}}pPr')
        e.insert(0, pPr)
    return pPr

def _is_image_para(child):
    """Check if paragraph contains an image."""
    blips = child.findall(f'.//{{{A}}}blip')
    return len(blips) > 0

def _is_table_caption(text):
    """Check if text looks like a table caption."""
    return bool(re.match(r'^表\s*\d*.*', text))

def _is_heading(child):
    """Check if paragraph is a heading."""
    pPr = child.find(f'{{{W}}}pPr')
    if pPr is not None:
        pStyle = pPr.find(f'{{{W}}}pStyle')
        if pStyle is not None:
            sv = pStyle.get(f'{{{W}}}val', '')
            # BodyText 的 styleId 是 '15'，不能误判为标题
            if 'Heading' in sv or (sv.isdigit() and sv != '15'):
                return True
    return False

def _get_para_text(child):
    """Extract text from paragraph."""
    texts = [t.text or '' for t in child.iter(f'{{{W}}}t')]
    return ''.join(texts).strip()

def _font_rPr():
    """Create a run properties element with 黑体 font."""
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), '黑体')
    rFonts.set(qn('w:eastAsia'), '黑体')
    rPr.append(rFonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '21')
    rPr.append(sz)
    return rPr


def _add_seq_field(parent, seq_name, font_rPr, display_value=1):
    """Add Word SEQ auto-numbering field to parent element.
    OOXML field structure: begin → instrText → separate → display text → end
    """
    # 1. begin
    run_begin = OxmlElement('w:r')
    if font_rPr is not None:
        run_begin.append(deepcopy(font_rPr))
    fldChar_begin = OxmlElement('w:fldChar')
    fldChar_begin.set(qn('w:fldCharType'), 'begin')
    run_begin.append(fldChar_begin)
    parent.append(run_begin)

    # 2. instrText (field code)
    run_instr = OxmlElement('w:r')
    if font_rPr is not None:
        run_instr.append(deepcopy(font_rPr))
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = f' SEQ {seq_name} \\* ARABIC '
    run_instr.append(instr)
    parent.append(run_instr)

    # 3. separate
    run_separate = OxmlElement('w:r')
    if font_rPr is not None:
        run_separate.append(deepcopy(font_rPr))
    fldChar_separate = OxmlElement('w:fldChar')
    fldChar_separate.set(qn('w:fldCharType'), 'separate')
    run_separate.append(fldChar_separate)
    parent.append(run_separate)

    # 4. display text
    run_display = OxmlElement('w:r')
    if font_rPr is not None:
        run_display.append(deepcopy(font_rPr))
    t_display = OxmlElement('w:t')
    t_display.text = str(display_value)
    run_display.append(t_display)
    parent.append(run_display)

    # 5. end
    run_end = OxmlElement('w:r')
    if font_rPr is not None:
        run_end.append(deepcopy(font_rPr))
    fldChar_end = OxmlElement('w:fldChar')
    fldChar_end.set(qn('w:fldCharType'), 'end')
    run_end.append(fldChar_end)
    parent.append(run_end)


def _make_caption_run(child, title_text, fig_number):
    """Replace paragraph content with SEQ-based caption, using 黑体 font."""
    # Remove existing runs
    for r in list(child.findall(f'{{{W}}}r')):
        child.remove(r)

    # Set style to 表标题(82) for consistency
    pPr = _pPr(child)
    pStyle = pPr.find(f'{{{W}}}pStyle')
    if pStyle is None:
        pStyle = etree.SubElement(pPr, f'{{{W}}}pStyle')
    pStyle.set(f'{{{W}}}val', '82')

    # Clear indent
    ind = pPr.find(f'{{{W}}}ind')
    if ind is None:
        ind = etree.SubElement(pPr, f'{{{W}}}ind')
    ind.set(f'{{{W}}}firstLine', '0')
    ind.set('{%s}firstLineChars' % W, '0')

    # Center
    jc = pPr.find(f'{{{W}}}jc')
    if jc is None:
        jc = etree.SubElement(pPr, f'{{{W}}}jc')
    jc.set(f'{{{W}}}val', 'center')

    # Add "图 " text
    r_prefix = OxmlElement('w:r')
    r_prefix.append(_font_rPr())
    t_prefix = OxmlElement('w:t')
    t_prefix.set(qn('xml:space'), 'preserve')
    t_prefix.text = '图 '
    r_prefix.append(t_prefix)
    child.append(r_prefix)

    # Add SEQ field
    _add_seq_field(child, '插图', _font_rPr(), fig_number)

    # Add name text after field (title_text is already just the name, no '图N' prefix)
    name = title_text.strip() if title_text else ''
    if name:
        r_name = OxmlElement('w:r')
        r_name.append(_font_rPr())
        t_name = OxmlElement('w:t')
        t_name.set(qn('xml:space'), 'preserve')
        t_name.text = f' {name}'
        r_name.append(t_name)
        child.append(r_name)


def format_figures(input_path, output_path=None):
    if output_path is None:
        output_path = input_path

    with zipfile.ZipFile(input_path, 'r') as zin:
        items = {}
        for name in zin.namelist():
            items[name] = zin.read(name)

        doc_xml = items.get('word/document.xml')
        if doc_xml is None:
            print("[ERROR] word/document.xml not found")
            return

        root = etree.fromstring(doc_xml)
        body = root.find(f'{{{W}}}body')
        if body is None:
            return

        # Get all body children (paragraphs + tables) in order
        children = list(body)
        img_count = 0
        fig_count = 0

        for i, child in enumerate(children):
            if child.tag != f'{{{W}}}p':
                continue

            # ── Image paragraphs: center + spacing + no indent ──
            if _is_image_para(child):
                img_count += 1
                pPr = _pPr(child)

                # ── Constrain image size to fit A4 page ──
                MAX_IMG_HEIGHT_EMU = 7560000  # ~14cm, fits A4 with margins
                inline = child.find(f'.//{{{WP}}}inline')
                if inline is not None:
                    extent = inline.find(f'{{{WP}}}extent')
                    if extent is not None:
                        cx = int(extent.get('cx', '0'))
                        cy = int(extent.get('cy', '0'))
                        if cy > MAX_IMG_HEIGHT_EMU and cx > 0:
                            scale = MAX_IMG_HEIGHT_EMU / cy
                            extent.set('cx', str(int(cx * scale)))
                            extent.set('cy', str(MAX_IMG_HEIGHT_EMU))

                jc = pPr.find(f'{{{W}}}jc')
                if jc is None:
                    jc = etree.SubElement(pPr, f'{{{W}}}jc')
                jc.set(f'{{{W}}}val', 'center')

                ind = pPr.find(f'{{{W}}}ind')
                if ind is None:
                    ind = etree.SubElement(pPr, f'{{{W}}}ind')
                ind.set(f'{{{W}}}firstLine', '0')
                ind.set('{%s}firstLineChars' % W, '0')

                sp = pPr.find(f'{{{W}}}spacing')
                if sp is None:
                    sp = etree.SubElement(pPr, f'{{{W}}}spacing')
                sp.set(f'{{{W}}}line', '240')
                sp.set('{%s}lineRule' % W, 'auto')
                sp.set(f'{{{W}}}before', '0')
                sp.set(f'{{{W}}}after', '0')

                # Check if THIS paragraph already has ImageCaption style (pandoc alt text)
                pPr2 = child.find(f'{{{W}}}pPr')
                if pPr2 is not None:
                    ps = pPr2.find(f'{{{W}}}pStyle')
                    if ps is not None and ps.get(f'{{{W}}}val') == 'ImageCaption':
                        current_name = _get_para_text(child)
                        if re.match(r'^image\d+\.(png|jpg|jpeg)$', current_name, re.I):
                            caption_name = ''
                        else:
                            caption_name = re.sub(r'^图\s*\d+\s*', '', current_name).strip()
                        fig_count += 1
                        _make_caption_run(child, caption_name, fig_count)
                        continue

                # Check NEXT paragraph for figure caption
                next_idx = i + 1
                while next_idx < len(children):
                    next_child = children[next_idx]
                    if next_child.tag != f'{{{W}}}p':
                        next_idx += 1
                        continue  # 跳过 bookmark 等非段落元素
                    next_text = _get_para_text(next_child)
                    if not next_text:
                        next_idx += 1
                        continue
                    if _is_table_caption(next_text):
                        break
                    if _is_heading(next_child):
                        break
                    if re.match(r'^编号格式|^页面设置|^正文格式|^一级标题|^二级标题|^标题编号', next_text):
                        break
                    # Strip any existing "图N" prefix from the text
                    caption_name = re.sub(r'^图\s*\d+\s*', '', next_text).strip()
                    fig_count += 1
                    _make_caption_run(next_child, caption_name, fig_count)
                    break

        # Write back
        items['word/document.xml'] = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)

        # Ensure updateFields in settings
        settings_key = None
        for k in items:
            if k.endswith('settings.xml'):
                settings_key = k
                break
        if settings_key:
            sroot = etree.fromstring(items[settings_key])
            uf = sroot.find(f'{{{W}}}updateFields')
            if uf is None:
                uf = etree.SubElement(sroot, f'{{{W}}}updateFields')
            uf.set(f'{{{W}}}val', 'true')
            items[settings_key] = etree.tostring(sroot, xml_declaration=True, encoding='UTF-8', standalone=True)

    # Write output
    tmp = output_path + '.ztmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for name, data in items.items():
            zout.writestr(name, data)
    os.replace(tmp, output_path)

    print(f"✅ 插图: {img_count}张图居中, {fig_count}个图标题")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python3 scripts/bid_format_figures.py 输入.docx [输出.docx]")
        sys.exit(1)
    format_figures(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
