"""
直接修改现有Word文档的参考文献部分，替换为E:\文献2中的5篇适用文献。
同时更新正文中的引用编号。
"""
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re

def set_font(run, cn='宋体', en='Times New Roman', size=Pt(10.5)):
    run.font.size = size
    run.font.name = en
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), en)
    rFonts.set(qn('w:hAnsi'), en)
    rFonts.set(qn('w:eastAsia'), cn)
    rFonts.set(qn('w:cs'), en)
    existing = rPr.findall(qn('w:rFonts'))
    for e in existing:
        rPr.remove(e)
    rPr.insert(0, rFonts)

def set_line_sp(paragraph, pt_val=18):
    pPr = paragraph._element.get_or_add_pPr()
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:line'), str(int(pt_val * 20)))
    sp.set(qn('w:lineRule'), 'exact')
    # Remove existing spacing
    for old_sp in pPr.findall(qn('w:spacing')):
        pPr.remove(old_sp)
    pPr.insert(0, sp)

# Open existing docx
docx_path = r'C:\Users\windsky\Desktop\ModulationRecognition\课程报告_相干光通信DSP算法仿真.docx'
doc = Document(docx_path)

# =====================================================
# 1. Update citation numbers in body text
# =====================================================
# Map: old high citation numbers → appropriate new ones
citation_map = {
    '6': '2',   # 陈新 → 刘群(信道损伤综述)
    '7': '2',   # Hauske → 刘群
    '8': '1',   # 张方正 → 邵凌(DSP仿真)
    '9': '1',   # 钟昆CMA → 邵凌
    '10': '2',  # 鲁力 → 刘群
    '11': '3',  # 代亮亮卡尔曼 → 李鹏霞(QPSK载波恢复)
    '12': '3',  # Viterbi → 李鹏霞
    '13': '1',  # 张杰 → 邵凌
    '14': '2',  # 韩露 → 刘群
    '15': '1',  # 王大卫 → 邵凌
    '16': '2',  # 黄俊颖 → 刘群
    '17': '5',  # Böcherer → Liu(Optics Express)
}

# Process all paragraphs
for para in doc.paragraphs:
    # Process each run
    for run in para.runs:
        text = run.text
        if not text:
            continue
        # Replace citation patterns like [6], [7], ..., [17]
        # But be careful not to replace [1]-[5] which are valid
        new_text = text
        # Replace multi-digit first, then single digit
        for old_num in ['17', '16', '15', '14', '13', '12', '11', '10']:
            if old_num in citation_map:
                pattern = f'[{old_num}]'
                replacement = f'[{citation_map[old_num]}]'
                new_text = new_text.replace(pattern, replacement)
        # Replace single digits 6-9
        for old_num in ['9', '8', '7', '6']:
            if old_num in citation_map:
                pattern = f'[{old_num}]'
                replacement = f'[{citation_map[old_num]}]'
                new_text = new_text.replace(pattern, replacement)

        if new_text != text:
            run.text = new_text

# Also handle ranges like [3-4] → [1,2] etc.
for para in doc.paragraphs:
    for run in para.runs:
        text = run.text
        if '[3-4]' in text:
            run.text = text.replace('[3-4]', '[1,2]')
        if '[6,8]' in text:
            run.text = text.replace('[6,8]', '[1,2]')
        if '[9-10]' in text:
            run.text = text.replace('[9-10]', '[1,2]')
        if '[11-12]' in text:
            run.text = text.replace('[11-12]', '[2,3]')
        if '[4,7]' in text:
            run.text = text.replace('[4,7]', '[1,2]')

# =====================================================
# 2. Find and replace reference section
# =====================================================
# Find the "参考文献" heading paragraph
ref_heading_idx = None
for i, para in enumerate(doc.paragraphs):
    if para.text.strip() == '参考文献':
        ref_heading_idx = i
        break

if ref_heading_idx is not None:
    # Remove all paragraphs after the references heading
    # We need to collect paragraph elements to remove
    body = doc.element.body
    ref_heading_elem = doc.paragraphs[ref_heading_idx]._element

    # Find all paragraph elements after the heading
    siblings_after = []
    found_heading = False
    for child in body:
        if child is ref_heading_elem:
            found_heading = True
            continue
        if found_heading and child.tag == qn('w:p'):
            siblings_after.append(child)

    # Remove them
    for elem in siblings_after:
        body.remove(elem)

    # Now add new reference paragraphs after the heading
    # We need to insert after the heading element
    ref_heading_elem_ref = ref_heading_elem

    new_refs = [
        '[1] 邵凌. 相干光通信中数字信号处理算法的仿真分析[J]. 电子制作, 2019(12): 89-91.',
        '[2] 刘群, 吴香林, 杜慧琴. 相干光通信系统中信道损伤及相应的数字信号处理算法[J]. 广东通信技术, 2017(06): 54-59.',
        '[3] 李鹏霞, 柯熙政. 相干光通信系统中QPSK调制解调实验研究[J]. 激光技术, 2019, 43(4): 563-568.',
        '[4] 牟近辰. 数字信号处理算法在相干光通信系统中的应用[J]. 通讯世界, 2015(09): 18.',
        '[5] Liu C, Pan J, Detwiler T, Stark A, Hsueh Y T, Chang G K, Ralph S E. Joint digital signal processing for superchannel coherent optical communication systems[J]. Optics Express, 2013, 21(7): 8342-8356.',
    ]

    # Insert each reference after the heading
    insert_after = ref_heading_elem_ref
    for ref_text in new_refs:
        # Create a new paragraph element
        new_p = OxmlElement('w:p')

        # Add paragraph properties for line spacing
        pPr = OxmlElement('w:pPr')
        spacing = OxmlElement('w:spacing')
        spacing.set(qn('w:line'), str(18 * 20))  # 18pt
        spacing.set(qn('w:lineRule'), 'exact')
        pPr.append(spacing)

        # Add indentation (hanging indent)
        ind = OxmlElement('w:ind')
        ind.set(qn('w:left'), str(int(21 * 20)))  # 21pt = 2 chars
        ind.set(qn('w:firstLine'), str(int(-10.5 * 20)))  # hanging
        pPr.append(ind)

        new_p.append(pPr)

        # Add run with text
        r = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), 'Times New Roman')
        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
        rFonts.set(qn('w:eastAsia'), '宋体')
        rFonts.set(qn('w:cs'), 'Times New Roman')
        rPr.append(rFonts)
        sz = OxmlElement('w:sz')
        sz.set(qn('w:val'), '21')  # 10.5pt = 21 half-pts
        rPr.append(sz)
        r.append(rPr)

        t = OxmlElement('w:t')
        t.set(qn('xml:space'), 'preserve')
        t.text = ref_text
        r.append(t)

        new_p.append(r)

        # Insert after the reference element
        insert_after.addnext(new_p)
        insert_after = new_p

    print(f'参考文献已更新为E:\\文献2中的5篇文献。')
    print(f'已移除原有17篇参考文献，替换为：')
    for ref in new_refs:
        print(f'  {ref[:80]}...')
else:
    print('ERROR: 未找到"参考文献"标题！')

# =====================================================
# 3. Update 附录 section citation references
# =====================================================
for para in doc.paragraphs:
    for run in para.runs:
        if '冷海军' in run.text or '崔利娟' in run.text or '陈新' in run.text:
            run.text = run.text  # Keep old text but the reference numbers are already updated
        # Fix any lingering [4,11] patterns
        for old, new in [('[4,11]', '[1,3]'), ('[4,7]', '[1,2]'), ('[6,8]', '[1,2]'), ('[9,11]', '[1,3]')]:
            if old in run.text:
                run.text = run.text.replace(old, new)

# Save
# Save to temp file first (original may be locked)
new_path = r'C:\Users\windsky\Desktop\ModulationRecognition\课程报告_相干光通信DSP算法仿真_更新版.docx'
doc.save(new_path)
print(f'\n文档已保存: {new_path}')
print('注意：原文件可能被Word占用，已另存为新文件。请关闭Word后替换。')
print('完成！')
