"""
深度清理并重新格式化图窗占位符，参考无线传感器网络实验报告的干净风格。
"""
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re

doc = Document(r'C:\Users\windsky\Desktop\ModulationRecognition\课程报告_相干光通信DSP算法仿真_OMML公式版.docx')

# ============================================
# 1. 深度XML清理：移除所有隐藏内容、VML垃圾、空标记
# ============================================
print('=== 深度XML清理 ===')

body = doc.element.body
all_elems = list(body.iter())

clean_count = 0
for elem in all_elems:
    tag = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag

    # 移除 vanish (隐藏文字)
    if tag == 'vanish':
        parent = elem.getparent()
        if parent is not None:
            parent.remove(elem)
            clean_count += 1

    # 移除 webHidden
    if tag == 'webHidden':
        parent = elem.getparent()
        if parent is not None:
            parent.remove(elem)
            clean_count += 1

    # 移除空 rPr (无属性无子元素)
    if tag == 'rPr':
        if len(elem) == 0 and len(elem.attrib) == 0:
            parent = elem.getparent()
            if parent is not None and parent.tag.split('}')[-1] == 'r':
                # Don't remove if it's the only rPr in a run
                pass

    # 移除空的 proofErr (拼写检查标记)
    if tag == 'proofErr':
        parent = elem.getparent()
        if parent is not None:
            parent.remove(elem)
            clean_count += 1

    # 移除 bookmarkStart/bookmarkEnd (如果存在)
    if tag in ('bookmarkStart', 'bookmarkEnd'):
        # Keep them - they're harmless and normal
        pass

    # 移除 gobalFormatting 等无关属性
    if tag == 'rPr':
        for child in list(elem):
            if child.tag.split('}')[-1] in ('noProof', 'lang', 'szCs'):
                elem.remove(child)
                clean_count += 1

print(f'移除垃圾XML元素: {clean_count} 处')

# ============================================
# 2. 查找并替换所有图占位符
# ============================================
print('\n=== 替换图占位符 ===')

fig_descs = {
    1: '色散补偿前后QPSK星座图对比 (28Gbaud, 100km SSMF)',
    2: '色散传递函数与频域补偿滤波器频率响应',
    3: 'CMA偏振解复用前后双偏振QPSK星座图对比',
    4: 'CMA算法收敛性能分析',
    5: '载波相位噪声影响与Viterbi-Viterbi算法恢复效果',
    6: '不同激光器线宽下VVPE算法载波相位恢复性能对比',
}

# 找到所有旧图元素
old_elements = []
for i, p in enumerate(doc.paragraphs):
    text = p.text.strip()
    if ('〔' in text and '插入图' in text) or (text.startswith('图') and len(text) > 3 and text[1:2].isdigit() and text[2:3] in (' ', '  ')):
        old_elements.append(i)

print(f'找到 {len(old_elements)} 个旧图元素: {old_elements}')

# 从后往前删除旧元素
for idx in sorted(old_elements, reverse=True):
    doc.element.body.remove(doc.paragraphs[idx]._element)

print('旧图元素已删除')

# ============================================
# 3. 创建干净图占位符（参考样式：五号宋体 10.5pt, auto行距, 居中）
# ============================================
def make_clean_fig_block(fig_num, desc):
    """创建干净的图占位+图题（参考WSN报告格式：10.5pt, auto行距, 居中）"""
    elements = []

    # --- 图框段落：auto行距，居中，10.5pt宋体 ---
    p_box = OxmlElement('w:p')
    pPr_box = OxmlElement('w:pPr')
    # auto行距 240
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:line'), '240')
    sp.set(qn('w:lineRule'), 'auto')
    pPr_box.append(sp)
    # 段前6pt 段后0
    sp2 = OxmlElement('w:spacing')
    pPr_box.append(sp2)
    sp2.set(qn('w:before'), '120')
    sp2.set(qn('w:after'), '0')
    # 居中
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center')
    pPr_box.append(jc)
    p_box.append(pPr_box)

    # Run: 灰色楷体提示文字 "〔 图N：描述 〕"  9pt
    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:eastAsia'), '楷体')
    rPr.append(rFonts)
    sz_el = OxmlElement('w:sz')
    sz_el.set(qn('w:val'), '18')  # 9pt
    rPr.append(sz_el)
    # 灰色
    color = OxmlElement('w:color')
    color.set(qn('w:val'), '808080')
    rPr.append(color)
    # 斜体
    i_el = OxmlElement('w:i')
    rPr.append(i_el)
    r.append(rPr)

    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = f'〔 图{fig_num}：{desc} 〕'
    r.append(t)
    p_box.append(r)
    elements.append(p_box)

    # --- 图题段落：auto行距，居中，小五宋体 9pt -- 空白，等用户插入图片后自行添加图题 ---
    # 实际上图题也需要存在。使用小五号宋体(9pt), 居中, auto行距
    p_cap = OxmlElement('w:p')
    pPr_cap = OxmlElement('w:pPr')
    sp_c = OxmlElement('w:spacing')
    sp_c.set(qn('w:line'), '240')
    sp_c.set(qn('w:lineRule'), 'auto')
    pPr_cap.append(sp_c)
    jc_c = OxmlElement('w:jc')
    jc_c.set(qn('w:val'), 'center')
    pPr_cap.append(jc_c)
    p_cap.append(pPr_cap)

    r_c = OxmlElement('w:r')
    rPr_c = OxmlElement('w:rPr')
    rFonts_c = OxmlElement('w:rFonts')
    rFonts_c.set(qn('w:ascii'), 'Times New Roman')
    rFonts_c.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts_c.set(qn('w:eastAsia'), '宋体')
    rPr_c.append(rFonts_c)
    sz_c = OxmlElement('w:sz')
    sz_c.set(qn('w:val'), '18')  # 9pt 小五号
    rPr_c.append(sz_c)
    r_c.append(rPr_c)
    t_c = OxmlElement('w:t')
    t_c.set(qn('xml:space'), 'preserve')
    t_c.text = f'图{fig_num}  {desc}'
    r_c.append(t_c)
    p_cap.append(r_c)
    elements.append(p_cap)

    return elements

# ============================================
# 4. 找到正确的插入位置
# ============================================
# 策略：找到每个图在正文中最后的引用段落，在其后插入
fig_insert_points = {}
for i, p in enumerate(doc.paragraphs):
    text = p.text.strip()
    for fn in range(1, 7):
        if fn == 1 and '图1为色散' in text:
            fig_insert_points[fn] = i
        elif fn == 2 and '图2从频域' in text:
            fig_insert_points[2] = i
        elif fn == 3 and '图3六子' in text:
            fig_insert_points[3] = i
        elif fn == 4 and '图4收敛' in text:
            fig_insert_points[4] = i
        elif fn == 5 and '图5展示' in text:
            fig_insert_points[5] = i
        elif fn == 6 and '图6对比' in text:
            fig_insert_points[6] = i

# Fallback: if not found, put after last mention
for fn in range(1, 7):
    if fn not in fig_insert_points:
        last_mention = -1
        for i, p in enumerate(doc.paragraphs):
            if f'图{fn}' in p.text:
                last_mention = i
        if last_mention >= 0:
            fig_insert_points[fn] = last_mention

print(f'图插入锚点: {fig_insert_points}')

# Insert figures in reverse order (to keep indices stable)
insertions = {}
for fn in range(1, 7):
    if fn in fig_insert_points:
        pos = fig_insert_points[fn]
        if pos not in insertions:
            insertions[pos] = []
        insertions[pos].append(fn)

# Flatten and sort by position (descending)
flat_insertions = []
for pos, fns in insertions.items():
    for fn in fns:
        flat_insertions.append((pos, fn))
flat_insertions.sort(key=lambda x: (-x[0], -x[1]))

print(f'插入计划: {flat_insertions}')

for pos, fn in flat_insertions:
    anchor = doc.paragraphs[pos]._element
    elems = make_clean_fig_block(fn, fig_descs[fn])
    for elem in reversed(elems):
        anchor.addnext(elem)

print(f'已插入 {len(flat_insertions)} 组图占位符')

# ============================================
# 5. 最终清理
# ============================================
print('\n=== 最终清理 ===')
# Remove any empty trailing paragraphs
body_children = list(doc.element.body)
# Remove empty paragraphs that are in the reference section
empty_tail = 0
for child in reversed(body_children):
    if child.tag == qn('w:p'):
        text = ''.join(child.itertext()).strip()
        if text == '':
            # Check if this has meaningful content (images, etc.)
            has_content = (child.findall('.//' + qn('w:drawing')) or
                          child.findall('.//' + qn('w:pict')) or
                          child.findall('.//' + qn('m:oMathPara')))
            if not has_content:
                # Check it's not the last section paragraph
                pPr = child.find(qn('w:pPr'))
                if pPr is not None:
                    sectPr = pPr.find(qn('w:sectPr'))
                    if sectPr is None:
                        doc.element.body.remove(child)
                        empty_tail += 1
        else:
            break

print(f'移除尾部空段落: {empty_tail} 个')

# ============================================
# 6. 验证
# ============================================
print('\n=== 验证 ===')
eq_count = sum(1 for p in doc.paragraphs if p._element.findall(qn('m:oMathPara')))
fig_count = sum(1 for p in doc.paragraphs if '〔' in p.text and '图' in p.text)
cap_count = sum(1 for p in doc.paragraphs if p.text.strip().startswith('图') and len(p.text.strip()) > 3 and p.text.strip()[1:2].isdigit())

# Final vanish check
vanish_remain = len(doc.element.body.findall('.//' + qn('w:vanish')))
proofErr_remain = len(doc.element.body.findall('.//' + qn('w:proofErr')))

print(f'公式: {eq_count} | 图占位: {fig_count} | 图题: {cap_count} | 段落: {len(doc.paragraphs)}')
print(f'残留vanish: {vanish_remain} | 残留proofErr: {proofErr_remain}')

# ============================================
# 保存
# ============================================
out = r'C:\Users\windsky\Desktop\ModulationRecognition\课程报告_相干光通信DSP算法仿真_最终版.docx'
doc.save(out)
print(f'\n文档已保存: {out}')
print('完成！')
