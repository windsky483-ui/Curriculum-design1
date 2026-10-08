"""生成交付说明的Word文档"""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_run_font(run, font_name_cn='宋体', font_name_en='Times New Roman', size=Pt(12)):
    run.font.size = size
    run.font.name = font_name_en
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), font_name_en)
    rFonts.set(qn('w:hAnsi'), font_name_en)
    rFonts.set(qn('w:eastAsia'), font_name_cn)
    rFonts.set(qn('w:cs'), font_name_en)
    rPr.insert(0, rFonts)

def set_line_spacing(paragraph, spacing_pt=20):
    pPr = paragraph._element.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:line'), str(int(spacing_pt * 20)))
    spacing.set(qn('w:lineRule'), 'exact')
    pPr.insert(0, spacing)

def add_p(doc, text, bold=False, size=Pt(12), align=None, cn_font='宋体'):
    p = doc.add_paragraph()
    set_line_spacing(p, 20)
    if align is not None:
        p.alignment = align
    run = p.add_run(text)
    if bold:
        run.bold = True
    set_run_font(run, font_name_cn=cn_font, size=size)
    return p

def add_h(doc, text, level=1):
    sizes = {1: Pt(18), 2: Pt(14), 3: Pt(12)}
    p = doc.add_paragraph()
    set_line_spacing(p, 24 if level == 1 else 22)
    run = p.add_run(text)
    run.bold = True
    set_run_font(run, font_name_cn='黑体', size=sizes.get(level, Pt(12)))
    return p

doc = Document()
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(12)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

section = doc.sections[0]
section.page_width = Cm(21)
section.page_height = Cm(29.7)
section.top_margin = Cm(2.54)
section.bottom_margin = Cm(2.54)
section.left_margin = Cm(2.5)
section.right_margin = Cm(2.5)

add_h(doc, '相干光通信课程报告——交付说明', level=1)

add_p(doc, '选题：相干光通信系统中数字信号处理算法的MATLAB仿真研究')
add_p(doc, '生成日期：2026年6月22日')
add_p(doc, '')

add_h(doc, '一、交付文件清单', level=2)

items = [
    ('课程报告_相干光通信DSP算法仿真.docx', 'Word文档', '完整课程报告，已按格式要求排版'),
    ('sim1_cd_compensation.m', 'MATLAB代码', '仿真1：频域色散补偿（生成图1、图2）'),
    ('sim2_cma_polarization.m', 'MATLAB代码', '仿真2：CMA偏振解复用（生成图3、图4）'),
    ('sim3_carrier_recovery.m', 'MATLAB代码', '仿真3：Viterbi-Viterbi载波相位恢复（生成图5、图6）'),
    ('generate_docx.py', 'Python脚本', 'Word文档生成脚本（可重新生成docx）'),
]
for name, typ, desc in items:
    add_p(doc, f'• {name}  [{typ}]  {desc}', size=Pt(11))

add_h(doc, '二、报告结构概览', level=2)

chapters = [
    ('摘要', '全文概述，关键词：相干光通信；DSP；色散补偿；CMA；载波相位恢复'),
    ('第1章 引言', '光纤通信发展背景，相干光通信技术的复兴与意义'),
    ('第2章 系统基本原理', '发送端/链路/接收端架构，相干检测的数学模型'),
    ('第3章 DSP算法研究', '6个算法模块：IQ补偿→色散补偿→时钟恢复→CMA→频偏估计→载波相位恢复'),
    ('第4章 MATLAB仿真', '3组仿真（各含参数、流程、结果分析），6张图占位符'),
    ('第5章 发展趋势', 'PCS整形、AI/深度学习DSP、空分复用SDM、低功耗芯片、DCI下沉'),
    ('第6章 总结与体会', '仿真结论、课程学习体会、不足与展望'),
    ('参考文献', '15篇，GB/T 7714格式（知网中文学位/期刊论文 + IEEE英文期刊论文）'),
    ('附录', 'MATLAB仿真代码说明与运行指南'),
]
for title, desc in chapters:
    add_p(doc, f'【{title}】{desc}', size=Pt(11))

add_h(doc, '三、MATLAB仿真输出对照', level=2)
add_p(doc, '以下6张图需要运行MATLAB代码生成后插入Word文档：')

figs = [
    ('图1', 'sim1', '色散补偿前后QPSK星座图对比（3个子图：发送/损伤/补偿）'),
    ('图2', 'sim1', '色散传递函数与补偿滤波器频域响应（相位+幅度）'),
    ('图3', 'sim2', 'CMA偏振解复用前后双偏振星座图（6个子图：X/Y偏振）'),
    ('图4', 'sim2', 'CMA收敛性能分析（误差曲线+抽头演化+学习曲线）'),
    ('图5', 'sim3', '相位噪声与VV恢复效果（5个子图：噪声/估计/星座/误差）'),
    ('图6', 'sim3', '4种线宽下VVPE性能对比（4星座+EVM柱状图）'),
]
for fid, src, desc in figs:
    add_p(doc, f'  {fid}（{src}）: {desc}', size=Pt(11))

add_h(doc, '四、操作步骤', level=2)

steps = [
    '第一步：在MATLAB中依次运行 sim1_cd_compensation.m、sim2_cma_polarization.m、sim3_carrier_recovery.m',
    '第二步：保存生成的6张图片（建议PNG格式，分辨率300dpi）',
    '第三步：打开 课程报告_相干光通信DSP算法仿真.docx',
    '第四步：在文档中标注 [此处插入图X] 的位置插入对应图片，图片居中排列',
    '第五步：确认图片下方图例编号正确（图1至图6），格式：居中、五号字',
    '第六步：（可选）在第二章和第三章开头分别插入系统架构框图和DSP流程图',
    '第七步：检查全文格式，确认总页数≥6页，双面黑白打印',
]
for i, step in enumerate(steps):
    add_p(doc, f'{step}', size=Pt(11))

add_h(doc, '五、格式设置确认清单', level=2)
checks = [
    '正文：小四宋体（12pt），英文 Times New Roman',
    '行间距：固定值20磅',
    '一级标题：小三黑体（15-16pt），加粗',
    '二级标题：四号黑体（14pt），加粗',
    '页面：A4，上下2.54cm，左右3.18cm',
    '图片≥3张：共6张MATLAB仿真图',
    '图片居中 + 图例居中 + 阿拉伯数字编号（图1-图6）',
    '参考文献：15篇，GB/T 7714格式',
    '双面黑白打印：Word打印设置中选择',
]
for c in checks:
    add_p(doc, f'  ☐ {c}', size=Pt(11))

add_p(doc, '')
add_p(doc, '— 以上为全部交付内容，祝报告顺利完成 —', size=Pt(12), align=WD_ALIGN_PARAGRAPH.CENTER, cn_font='楷体')

output_path = r'c:\Users\windsky\Desktop\课程报告交付说明.docx'
doc.save(output_path)
print(f'交付说明已保存至: {output_path}')
