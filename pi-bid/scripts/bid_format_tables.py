#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
标书表格格式化 + SEQ 自动编号 — 直接操作 docx XML。
用法: python3 bid_format_tables.py 输入.docx [输出.docx]

表标题名来源（按优先级）:
  1. 环境变量 BID_TABLE_NAMES_FILE 指向的文本文件（每行一个表名，按文档顺序）
  2. 表格前一个段落已有的"表N 名称"题注（markdown 里写好的）
  3. 以上都没有时，只生成"表N"编号不带名称

功能:
  1. 所有表格页面对齐: 居中
  2. 表格文字: 黑体五号, 左右/上下居中, 无首行缩进
  3. 表标题: 黑体五号, SEQ自动编号 (表 1, 表 2, ...)
  4. 表头: 加粗
  5. 文档设置: updateFields=true (打开时自动刷新编号)

依赖: python-docx, lxml
"""
import os

def _load_table_names():
    path = os.environ.get('BID_TABLE_NAMES_FILE', '')
    if path and os.path.exists(path):
        with open(path, encoding='utf-8') as f:
            return [l.strip() for l in f if l.strip()]
    return []

_TABLE_NAMES = _load_table_names()

import re
import sys
import zipfile
from copy import deepcopy
from lxml import etree
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


# ─── helpers ──────────────────────────────────────────────────────────

def _font_rPr(name='黑体', size_half_pt=21):
    """创建字体属性元素 (rPr)"""
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), name)
    rFonts.set(qn('w:eastAsia'), name)
    rPr.append(rFonts)
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(size_half_pt))
    rPr.append(sz)
    return rPr


def _add_seq_field(parent, seq_name='表格', display_value=1, font_rPr=None):
    """
    添加 Word SEQ 自动编号字段到 parent 元素。
    生成结构: begin → instrText → separate → 显示值 → end
    """
    if font_rPr is None:
        font_rPr = _font_rPr()

    def _fld(t):
        e = OxmlElement('w:fldChar')
        e.set(qn('w:fldCharType'), t)
        return e

    # begin
    r1 = OxmlElement('w:r')
    r1.append(deepcopy(font_rPr))
    r1.append(_fld('begin'))
    parent.append(r1)

    # instrText
    r2 = OxmlElement('w:r')
    r2.append(deepcopy(font_rPr))
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = f' SEQ {seq_name} \\* ARABIC \\s 1 '
    r2.append(instr)
    parent.append(r2)

    # separate
    r3 = OxmlElement('w:r')
    r3.append(deepcopy(font_rPr))
    r3.append(_fld('separate'))
    parent.append(r3)

    # 显示值 (初始序号)
    r4 = OxmlElement('w:r')
    r4.append(deepcopy(font_rPr))
    t = OxmlElement('w:t')
    t.text = str(display_value)
    r4.append(t)
    parent.append(r4)

    # end
    r5 = OxmlElement('w:r')
    r5.append(deepcopy(font_rPr))
    r5.append(_fld('end'))
    parent.append(r5)


# ─── main ────────────────────────────────────────────────────────────

def format_tables(input_path: str, output_path: str = None):
    """格式化所有表格: 居中 + 黑体五号 + SEQ 自动编号"""

    if output_path is None:
        output_path = input_path

    # ── 0. 正文样式: Normal → BodyText (首行缩进) ──
    doc = Document(input_path)

    # pandoc 输出的正文段落都是 Normal 样式, 模板的 BodyText 才有首行缩进
    # 注意: python-docx 的 style.name 对 pandoc 非标题样式全返回 'Normal'
    # (ImageCaption, BodyText, FirstParagraph, Compact 等均被误判)
    # 所以必须检查 XML w:pStyle 来跳过图相关段落
    try:
        bodytext = doc.styles['Body Text']
        ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        for para in doc.paragraphs:
            if para.style.name != 'Normal':
                continue
            # ── 检查 XML 实际样式 ID ──
            # python-docx 把 Heading6/Heading7/ImageCaption/BodyText/... 全报为 'Normal'
            # 标题和图相关段落必须跳过
            pPr = para._element.find(qn('w:pPr'))
            if pPr is not None:
                ps = pPr.find(qn('w:pStyle'))
                if ps is not None:
                    style_id = ps.get(qn('w:val'))
                    if style_id.startswith('Heading') or style_id in ('ImageCaption', 'CaptionedFigure', 'BlockText'):
                        continue  # 标题/图/引用块不转 Body Text
            # ──
            t = para.text.strip()
            if not t or len(t) < 8:
                continue
            # 跳过图片行
            if para._element.findall(f'.//{{{ns_a}}}blip'):
                continue
            # 跳过列表项和编号项
            if t[0].isdigit() and '. ' in t[:5]:
                continue
            if t.startswith('-') or t.startswith('•'):
                continue
            para.style = bodytext
            # ── 列表项(w:numPr)会覆盖样式级首行缩进, 加显式段落级缩进 ──
            # pandoc 的编号列表项(1. xxx)/定义项都带 w:numPr
            # Word 中列表缩进优先级高于样式缩进, 必须显式添加
            pPr = para._element.find(qn('w:pPr'))
            if pPr is not None and pPr.find(qn('w:numPr')) is not None:
                ind = pPr.find(qn('w:ind'))
                if ind is None:
                    ind = OxmlElement('w:ind')
                    pPr.append(ind)
                # 与样式 15 定义保持一致: firstLineChars=200
                if not ind.get(qn('w:firstLine')):
                    ind.set(qn('w:firstLine'), '480')
                    ind.set(qn('w:firstLineChars'), '200')
            # ──
    except KeyError:
        pass  # BodyText 样式不存在时不处理

    # ── 1. 表格居中 + 黑体五号 + 清除缩进 + Table Grid ──
    try:
        tbl_grid_style = doc.styles['Table Grid']
    except KeyError:
        tbl_grid_style = None

    for table in doc.tables:
        # 表格页面居中
        tbl = table._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)
        jc = tblPr.find(qn('w:jc'))
        if jc is None:
            jc = OxmlElement('w:jc')
            tblPr.append(jc)
        jc.set(qn('w:val'), 'center')

        # 应用表格样式 (Table Grid: 有框线)
        if tbl_grid_style is not None:
            table.style = tbl_grid_style

        # 添加显式表格边框（确保渲染）
        tblBorders = tblPr.find(qn('w:tblBorders'))
        if tblBorders is None:
            tblBorders = OxmlElement('w:tblBorders')
            tblPr.append(tblBorders)
            for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                be = OxmlElement(f'w:{edge}')
                be.set(qn('w:val'), 'single')
                be.set(qn('w:sz'), '4')
                be.set(qn('w:space'), '0')
                be.set(qn('w:color'), '000000')
                tblBorders.append(be)

        for ri, row in enumerate(table.rows):
            for cell in row.cells:
                for para in cell.paragraphs:
                    # 首行缩进清除
                    pPr = para._element.find(qn('w:pPr'))
                    if pPr is None:
                        pPr = OxmlElement('w:pPr')
                        para._element.insert(0, pPr)

                    # 行距: 单倍不固定
                    spacing = pPr.find(qn('w:spacing'))
                    if spacing is None:
                        spacing = OxmlElement('w:spacing')
                        pPr.append(spacing)
                    spacing.set(qn('w:line'), '240')
                    spacing.set(qn('w:lineRule'), 'auto')
                    spacing.set(qn('w:before'), '0')
                    spacing.set(qn('w:after'), '0')

                    ind = pPr.find(qn('w:ind'))
                    if ind is None:
                        ind = OxmlElement('w:ind')
                        pPr.append(ind)
                    ind.set(qn('w:firstLine'), '0')
                    ind.set(qn('w:firstLineChars'), '0')
                    ind.set(qn('w:left'), '0')

                    # 左右居中
                    para.alignment = 1  # CENTER

                    # 垂直居中
                    tc = cell._tc
                    tcPr = tc.find(qn('w:tcPr'))
                    if tcPr is None:
                        tcPr = OxmlElement('w:tcPr')
                        tc.insert(0, tcPr)
                    va = tcPr.find(qn('w:vAlign'))
                    if va is None:
                        va = OxmlElement('w:vAlign')
                        tcPr.append(va)
                    va.set(qn('w:val'), 'center')

                    # 字体
                    for run in para.runs:
                        run.font.name = '黑体'
                        run.font.size = Pt(10.5)
                        rPr = run._element.find(qn('w:rPr'))
                        if rPr is None:
                            rPr = OxmlElement('w:rPr')
                            run._element.insert(0, rPr)
                        rFonts = rPr.find(qn('w:rFonts'))
                        if rFonts is None:
                            rFonts = OxmlElement('w:rFonts')
                            rPr.append(rFonts)
                        rFonts.set(qn('w:ascii'), '黑体')
                        rFonts.set(qn('w:eastAsia'), '黑体')
                        rFonts.set(qn('w:hAnsi'), '黑体')

                    # 表头不加粗 (取消之前加粗逻辑)
                    # 根据内容调整行高
                    row_height = row._tr.find(qn('w:trPr'))
                    if row_height is None:
                        row_height = OxmlElement('w:trPr')
                        row._tr.insert(0, row_height)
                    # 设置自动行高 (最小高度为0)
                    trHeight = row_height.find(qn('w:trHeight'))
                    if trHeight is None:
                        trHeight = OxmlElement('w:trHeight')
                        row_height.append(trHeight)
                    trHeight.set(qn('w:val'), '0')
                    trHeight.set(qn('w:hRule'), 'auto')

    # ── 2. 表标题 SEQ 自动编号 ──
    seq_counter = 0

    # 遍历所有表格, 检查其前面是否有题注段落
    body = doc.element.body
    prev_para = None
    prev_comment = None
    for child in body:
        tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

        if tag == 'p':
            prev_para = child
            prev_comment = None  # reset

        elif tag == 'r':
            # Check for comment reference (not used here)
            pass

        elif tag == 'tbl':
            seq_counter += 1
            caption_text = None
            if seq_counter <= len(_TABLE_NAMES):
                caption_text = _TABLE_NAMES[seq_counter - 1]

            # Check if previous paragraph is already a table caption (from markdown)
            reuse_prev = False
            if prev_para is not None:
                prev_texts = [t.text or '' for t in prev_para.iter(f'{{{W}}}t')]
                prev_text = ''.join(prev_texts).strip()
                if prev_text.startswith('表'):
                    reuse_prev = True
                    # If prev_text has descriptive name beyond just "表N", use it
                    prev_name = re.sub(r'^表[\s\d\-]*', '', prev_text).strip()
                    if prev_name:
                        caption_text = prev_name
                    # Reuse the existing paragraph as caption
                    cap_para = prev_para
                    # Clear existing content
                    for child_elem in list(cap_para):
                        cap_para.remove(child_elem)
                    # Set style to 82
                    pPr = cap_para.find(f'{{{W}}}pPr')
                    if pPr is None:
                        pPr = OxmlElement('w:pPr')
                        cap_para.insert(0, pPr)
                    pStyle = pPr.find(f'{{{W}}}pStyle')
                    if pStyle is None:
                        pStyle = OxmlElement('w:pStyle')
                        pPr.append(pStyle)
                    pStyle.set(qn('w:val'), '82')
                    # Remove indent
                    ind = pPr.find(f'{{{W}}}ind')
                    if ind is not None:
                        pPr.remove(ind)
                    # Add center alignment
                    jc = pPr.find(f'{{{W}}}jc')
                    if jc is None:
                        jc = OxmlElement('w:jc')
                        pPr.append(jc)
                    jc.set(qn('w:val'), 'center')

            if not reuse_prev:
                # Build new caption paragraph
                cap_para = OxmlElement('w:p')
                pPr = OxmlElement('w:pPr')
                pStyle = OxmlElement('w:pStyle')
                pStyle.set(qn('w:val'), '82')
                pPr.append(pStyle)
                cap_para.append(pPr)
                # Insert before table
                child.addprevious(cap_para)

            # "表 " 前缀
            r0 = OxmlElement('w:r')
            r0.append(_font_rPr())
            t0 = OxmlElement('w:t')
            t0.set(qn('xml:space'), 'preserve')
            t0.text = '表 '
            r0.append(t0)
            cap_para.append(r0)

            # SEQ 字段
            _add_seq_field(cap_para, '表格', seq_counter, _font_rPr())

            # Table name text
            if caption_text:
                cname = re.sub(r'^表[\s\d]*', '', caption_text).strip()
                if cname:
                    r6 = OxmlElement('w:r')
                    r6.append(_font_rPr())
                    t2 = OxmlElement('w:t')
                    t2.set(qn('xml:space'), 'preserve')
                    t2.text = f' {cname}'
                    r6.append(t2)
                    cap_para.append(r6)

            prev_para = None

    doc.save(output_path)

    # ── 3. 文档设置 updateFields ──
    tmp = output_path + '.tmp'
    with zipfile.ZipFile(output_path, 'r') as zin:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item == 'word/settings.xml':
                    root = etree.fromstring(data)
                    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                    uf = root.find('.//w:updateFields', ns)
                    if uf is None:
                        uf = etree.SubElement(root, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}updateFields')
                    uf.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'true')
                    data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
                zout.writestr(item, data)
    os.replace(tmp, output_path)

    print(f"✅ 格式化完成: {seq_counter} 张表已编号, 全部居中/黑体五号")
    return True


# ─── CLI ─────────────────────────────────────────────────────────────

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python3 scripts/bid_format_tables.py 输入.docx [输出.docx]")
        sys.exit(1)
    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    if not os.path.exists(inp):
        print(f"[ERROR] 文件不存在: {inp}")
        sys.exit(1)
    format_tables(inp, out)
