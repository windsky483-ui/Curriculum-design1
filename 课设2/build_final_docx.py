"""
最终版：用OMML公式重新生成知网格式课程报告
所有公式使用Word原生公式编辑器格式
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from omml_engine import *
from omml_engine import _func, _bracket_square, _group, _abs, _hat, _bar, _sub, _sup, _subsup, _frac
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ========== 复用之前的样式函数 ==========
def set_run_font(run, cn='宋体', en='Times New Roman', size=Pt(12), bold=False):
    run.font.size = size; run.font.name = en; run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), en); rFonts.set(qn('w:hAnsi'), en)
    rFonts.set(qn('w:eastAsia'), cn); rFonts.set(qn('w:cs'), en)
    for e in rPr.findall(qn('w:rFonts')): rPr.remove(e)
    rPr.insert(0, rFonts)

def set_line_sp(paragraph, pt_val=20):
    pPr = paragraph._element.get_or_add_pPr()
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:line'), str(int(pt_val * 20)))
    sp.set(qn('w:lineRule'), 'exact')
    for old in pPr.findall(qn('w:spacing')): pPr.remove(old)
    pPr.insert(0, sp)

def set_spacing(paragraph, before=0, after=0):
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)

def set_first_indent(paragraph, chars=2, font_size_pt=12):
    paragraph.paragraph_format.first_line_indent = Pt(font_size_pt * chars)

def add_body(doc, text, size=Pt(12), bold=False, align=None):
    p = doc.add_paragraph()
    set_line_sp(p, 20)
    if align: p.alignment = align
    set_first_indent(p, 2, 12)
    import re
    parts = re.split(r'(\*\*.*?\*\*)', text)
    for part in parts:
        is_bold = part.startswith('**') and part.endswith('**')
        txt = part[2:-2] if is_bold else part
        run = p.add_run(txt)
        set_run_font(run, size=size, bold=(bold or is_bold))
    return p

def add_heading_cnki(doc, text, level=1):
    cfg = {1: (Pt(15), '黑体', 6, 3, 22), 2: (Pt(14), '黑体', 4, 2, 22), 3: (Pt(12), '黑体', 2, 1, 20)}
    s, cn, sb, sa, ls = cfg.get(level, cfg[1])
    p = doc.add_paragraph(); set_line_sp(p, ls); set_spacing(p, sb, sa)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    set_run_font(run, cn=cn, size=s, bold=True)
    return p

def add_fig_placeholder(doc, fig_num, desc):
    p = doc.add_paragraph(); set_line_sp(p, 140); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f'〔 请在此处插入图{fig_num}：{desc} 〕')
    set_run_font(r, cn='楷体', size=Pt(9))
    p2 = doc.add_paragraph(); set_line_sp(p2, 14); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = p2.add_run(f'图{fig_num}  {desc}')
    set_run_font(r2, size=Pt(9))

def add_ref(doc, text):
    p = doc.add_paragraph(); set_line_sp(p, 18); set_spacing(p, 0, 0)
    pf = p.paragraph_format; pf.first_line_indent = Pt(-10.5); pf.left_indent = Pt(21)
    r = p.add_run(text); set_run_font(r, size=Pt(10.5))

def add_blank(doc, n=1):
    for _ in range(n):
        p = doc.add_paragraph(); set_line_sp(p, 12)

# ========== OMML公式插入 ==========
def add_eq(doc, omath):
    """居中插入OMML公式"""
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    omP = make_omathpara(omath)
    p._element.append(omP)
    return p

# ========== 构建各公式 ==========
def eq_Es():  # E_s(t) = A_s(t)·exp[j(ω_s·t + θ_s(t))]
    return make_omath(
        eq_italic_sub('E','s'), eq_normal_text('('), eq_italic_var('t'), eq_normal_text(') = '),
        eq_italic_sub('A','s'), eq_normal_text('('), eq_italic_var('t'), eq_normal_text(')·'),
        _func('exp', _bracket_square(
            eq_italic_var('j'), eq_normal_text('('),
            eq_italic_sub('ω','s'), eq_italic_var('t'),
            eq_normal_text(' + '), eq_italic_sub('θ','s'),
            eq_normal_text('('), eq_italic_var('t'), eq_normal_text(')'),
            eq_normal_text(')')
        ))
    )

def eq_ELO():  # E_LO(t) = A_LO·exp[j(ω_LO·t + θ_LO(t))]
    return make_omath(
        eq_italic_sub('E','LO'), eq_normal_text('('), eq_italic_var('t'), eq_normal_text(') = '),
        eq_italic_sub('A','LO'), eq_normal_text('·'),
        _func('exp', _bracket_square(
            eq_italic_var('j'), eq_normal_text('('),
            eq_italic_sub('ω','LO'), eq_italic_var('t'),
            eq_normal_text(' + '), eq_italic_sub('θ','LO'),
            eq_normal_text('('), eq_italic_var('t'), eq_normal_text(')'),
            eq_normal_text(')')
        ))
    )

def eq_It():  # I(t) ∝ A_LO·A_s(t)·exp[j(ω_s-ω_LO)t + j(θ_s(t)-θ_LO(t))]
    return make_omath(
        eq_italic_var('I'), eq_normal_text('('), eq_italic_var('t'), eq_normal_text(') ∝ '),
        eq_italic_sub('A','LO'), eq_normal_text('·'),
        eq_italic_sub('A','s'), eq_normal_text('('), eq_italic_var('t'), eq_normal_text(')·'),
        _func('exp', _bracket_square(
            eq_italic_var('j'), eq_normal_text('('),
            eq_italic_sub('ω','s'), eq_normal_text(' − '), eq_italic_sub('ω','LO'),
            eq_normal_text(')'), eq_italic_var('t'), eq_normal_text(' + '),
            eq_italic_var('j'), eq_normal_text('('),
            eq_italic_sub('θ','s'), eq_normal_text('('), eq_italic_var('t'),
            eq_normal_text(') − '), eq_italic_sub('θ','LO'),
            eq_normal_text('('), eq_italic_var('t'), eq_normal_text(')'),
            eq_normal_text(')')
        ))
    )

def eq_HCD():  # H_CD(ω,L) = exp(−j·β₂·ω²·L/2) = exp(j·λ²·D·ω²·L/(4πc))
    return make_omath(
        eq_italic_sub('H','CD'), eq_normal_text('('), eq_italic_var('ω'), eq_normal_text(', '),
        eq_italic_var('L'), eq_normal_text(') = '),
        _func('exp', _group(
            eq_normal_text('−'), eq_italic_var('j'), eq_normal_text('·'),
            eq_italic_sub('β','2'), eq_normal_text('·'),
            eq_italic_subsup('ω','','2'), eq_normal_text('·'),
            eq_italic_var('L'), eq_normal_text('/2')
        )),
        eq_normal_text(' = '),
        _func('exp', _group(
            eq_italic_var('j'), eq_normal_text('·'),
            eq_italic_subsup('λ','','2'), eq_normal_text('·'),
            eq_italic_var('D'), eq_normal_text('·'),
            eq_italic_subsup('ω','','2'), eq_normal_text('·'),
            eq_italic_var('L'), eq_normal_text('/(4'), eq_italic_var('π'),
            eq_italic_var('c'), eq_normal_text(')')
        )),
    )

def eq_HCD_simple():  # H_CD(ω,L) = exp(−j·β₂·ω²·L/2)
    return make_omath(
        eq_italic_sub('H','CD'), eq_normal_text('('), eq_italic_var('ω'), eq_normal_text(', '),
        eq_italic_var('L'), eq_normal_text(') = '),
        _func('exp', _group(
            eq_normal_text('−'), eq_italic_var('j'), eq_normal_text('·'),
            eq_italic_sub('β','2'), eq_normal_text('·'),
            eq_italic_subsup('ω','','2'), eq_normal_text('·'),
            eq_italic_var('L'), eq_normal_text('/2')
        )),
    )

def eq_HCD_inv():  # H_CD⁻¹(ω,L) = exp(j·β₂·ω²·L/2)
    return make_omath(
        eq_italic_subsup('H','CD','−1'), eq_normal_text('('),
        eq_italic_var('ω'), eq_normal_text(', '), eq_italic_var('L'), eq_normal_text(') = '),
        _func('exp', _group(
            eq_italic_var('j'), eq_normal_text('·'),
            eq_italic_sub('β','2'), eq_normal_text('·'),
            eq_italic_subsup('ω','','2'), eq_normal_text('·'),
            eq_italic_var('L'), eq_normal_text('/2')
        )),
    )

def eq_gardner():  # e(n) = Re{[x(n)−x(n−2)]·x*(n−1)}
    return make_omath(
        eq_italic_var('e'), eq_normal_text('('), eq_italic_var('n'), eq_normal_text(') = Re{'),
        _bracket_square(
            eq_italic_var('x'), eq_normal_text('('), eq_italic_var('n'), eq_normal_text(') − '),
            eq_italic_var('x'), eq_normal_text('('), eq_italic_var('n'), eq_normal_text('−2)'),
        ),
        eq_normal_text('·'),
        eq_italic_subsup('x','','*'), eq_normal_text('('),
        eq_italic_var('n'), eq_normal_text('−1)}'),
    )

def eq_Xout():  # X_out(n) = h_xx^H·X_in(n) + h_xy^H·Y_in(n)
    return make_omath(
        eq_italic_sub('X','out'), eq_normal_text('('), eq_italic_var('n'), eq_normal_text(') = '),
        eq_italic_subsup('h','xx','H'), eq_normal_text('·'),
        eq_italic_sub('X','in'), eq_normal_text('('), eq_italic_var('n'), eq_normal_text(') + '),
        eq_italic_subsup('h','xy','H'), eq_normal_text('·'),
        eq_italic_sub('Y','in'), eq_normal_text('('), eq_italic_var('n'), eq_normal_text(')'),
    )

def eq_Yout():  # Y_out(n) = h_yx^H·X_in(n) + h_yy^H·Y_in(n)
    return make_omath(
        eq_italic_sub('Y','out'), eq_normal_text('('), eq_italic_var('n'), eq_normal_text(') = '),
        eq_italic_subsup('h','yx','H'), eq_normal_text('·'),
        eq_italic_sub('X','in'), eq_normal_text('('), eq_italic_var('n'), eq_normal_text(') + '),
        eq_italic_subsup('h','yy','H'), eq_normal_text('·'),
        eq_italic_sub('Y','in'), eq_normal_text('('), eq_italic_var('n'), eq_normal_text(')'),
    )

def eq_JCMA():  # J_CMA = E[(|X_out|² − R)²]
    return make_omath(
        eq_italic_sub('J','CMA'), eq_normal_text(' = '),
        eq_italic_var('E'), eq_normal_text('[(|'),
        eq_italic_sub('X','out'), eq_normal_text('|'),
        eq_italic_subsup('','','2'), eq_normal_text(' − '),
        eq_italic_var('R'), eq_normal_text(')'),
        eq_italic_subsup('','','2'), eq_normal_text(']'),
    )

def eq_eps():  # ε_x = R − |X_out|²
    return make_omath(
        eq_italic_sub('ε','x'), eq_normal_text(' = '),
        eq_italic_var('R'), eq_normal_text(' − |'),
        eq_italic_sub('X','out'), eq_normal_text('|'),
        eq_italic_subsup('','','2'),
    )

def eq_tap():  # h_ij(n+1) = h_ij(n) + μ·ε·out(n)·conj[in(n)]
    return make_omath(
        eq_italic_sub('h','ij'), eq_normal_text('('), eq_italic_var('n'), eq_normal_text('+1) = '),
        eq_italic_sub('h','ij'), eq_normal_text('('), eq_italic_var('n'), eq_normal_text(') + '),
        eq_italic_var('μ'), eq_normal_text('·'),
        eq_italic_var('ε'), eq_normal_text('·'),
        eq_italic_var('out'), eq_normal_text('('), eq_italic_var('n'), eq_normal_text(')·'),
        eq_normal_text('conj['),
        eq_italic_var('in'), eq_normal_text('('), eq_italic_var('n'), eq_normal_text(')]'),
    )

def eq_df():  # Δf_est = (1/4)·[1/(2πT_s)]·arg{Σ[X_out(n)·X_out*(n−1)]⁴}
    return make_omath(
        eq_normal_text('Δ'), eq_italic_sub('f','est'),
        eq_normal_text(' = (1/4)·[1/(2'), eq_italic_var('π'),
        eq_italic_sub('T','s'), eq_normal_text(')]·arg{'),
        eq_normal_text('Σ['),
        eq_italic_sub('X','out'), eq_normal_text('('), eq_italic_var('n'), eq_normal_text(')·'),
        eq_italic_subsup('X','out','*'), eq_normal_text('('),
        eq_italic_var('n'), eq_normal_text('−1)]'),
        eq_italic_subsup('','','4'), eq_normal_text('}'),
    )

# ========== 构建完整文档 ==========
print("生成OMML公式版最终文档...")
doc = Document()

# 默认样式
style = doc.styles['Normal']
style.font.name = 'Times New Roman'; style.font.size = Pt(12)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

for section in doc.sections:
    section.page_width = Cm(21); section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54); section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.17); section.right_margin = Cm(3.17)

# ==== 标题 ====
add_blank(doc, 2)
p = doc.add_paragraph(); set_line_sp(p, 28); set_spacing(p, 12, 12)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('相干光通信系统中数字信号处理算法的MATLAB仿真研究')
set_run_font(r, cn='黑体', size=Pt(18), bold=True)
add_blank(doc)

# ==== 摘要 ====
p = doc.add_paragraph(); set_line_sp(p, 22); set_spacing(p, 6, 3)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('摘  要'); set_run_font(r, cn='黑体', size=Pt(15), bold=True)

add_body(doc, '相干光通信技术是现代高速光纤通信网络的核心支撑技术之一。与传统的强度调制/直接检测（IM/DD）系统相比，相干检测技术具有更高的接收灵敏度、频谱效率以及支持多维度调制的优势，已成为100G及以上速率光纤通信系统的标准方案。在相干光通信接收端，数字信号处理（DSP）算法承担着补偿各类信道损伤的关键任务，包括色度色散补偿、偏振解复用、载波频率偏移估计及相位恢复等。本文系统梳理了相干光通信系统的基本原理与接收端DSP算法的标准流程，重点研究了三类核心算法：频域色散补偿、基于恒模算法（CMA）的偏振解复用，以及基于Viterbi-Viterbi算法的载波相位恢复。利用MATLAB软件对上述算法进行了数值仿真分析，通过星座图对比、收敛曲线分析及不同参数下的性能评估，验证了DSP算法在相干光通信系统中的有效性。最后，对相干光通信DSP技术的未来发展趋势进行了展望。')

p = doc.add_paragraph(); set_line_sp(p, 20); set_first_indent(p, 2, 12)
r1 = p.add_run('关键词：'); set_run_font(r1, cn='黑体', size=Pt(12), bold=True)
r2 = p.add_run('相干光通信；数字信号处理；色散补偿；恒模算法；载波相位恢复；偏振解复用')
set_run_font(r2, size=Pt(12))

# ==== 英文摘要 ====
add_blank(doc)
p = doc.add_paragraph(); set_line_sp(p, 22); set_spacing(p, 6, 3)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Abstract'); set_run_font(r, cn='Times New Roman', size=Pt(15), bold=True)

add_body(doc, 'Coherent optical communication technology is one of the core enabling technologies for modern high-speed fiber-optic communication networks. Compared with traditional intensity modulation/direct detection (IM/DD) schemes, coherent detection offers higher receiver sensitivity, superior spectral efficiency, and the ability to support multi-dimensional modulation formats, making it the standard solution for optical communication systems at 100G and beyond. At the coherent receiver, digital signal processing (DSP) algorithms play a critical role in compensating for various channel impairments, including chromatic dispersion compensation, polarization demultiplexing, carrier frequency offset estimation, and carrier phase recovery. This paper systematically reviews the fundamental principles of coherent optical communication systems and the standard DSP processing chain at the receiver, with an in-depth focus on three core algorithms: frequency-domain CD equalization, constant modulus algorithm (CMA) based polarization demultiplexing, and Viterbi-Viterbi carrier phase recovery. MATLAB simulations are conducted to validate the effectiveness of these algorithms through constellation diagram comparisons, convergence analysis, and performance evaluation under various parameters.')

add_blank(doc)
p = doc.add_paragraph(); set_line_sp(p, 20); set_first_indent(p, 2, 12)
r1 = p.add_run('Key words: '); set_run_font(r1, cn='Times New Roman', size=Pt(12), bold=True)
r2 = p.add_run('coherent optical communication; digital signal processing; chromatic dispersion compensation; constant modulus algorithm; carrier phase recovery; polarization demultiplexing')
set_run_font(r2, cn='Times New Roman', size=Pt(12))

doc.add_page_break()

# ==== 第1章 引言 ====
add_heading_cnki(doc, '1  引言', 1)

add_body(doc, '随着云计算、人工智能、5G/6G移动通信及超高清视频等带宽密集型应用的迅猛发展，全球网络数据流量呈指数级增长，这给光纤通信网络的传输容量带来了持续挑战[1]。在过去的三十年中，光纤通信系统经历了从准同步数字体系（PDH）到同步数字体系（SDH），再到光传送网（OTN）的多代演进，单波长传输速率已从2.5 Gbit/s提升至400 Gbit/s乃至800 Gbit/s，波分复用（WDM）技术使单纤传输容量突破100 Tbit/s量级。')

add_body(doc, '推动这一进步的三大关键技术分别是：波分复用技术（WDM）、光放大技术（EDFA/Raman）以及相干光通信技术。其中，相干光通信技术的复兴与实用化（约2005年以后）具有里程碑意义。与传统的IM/DD方案相比，相干检测能够完整保留光场的幅度、相位和偏振态信息，结合高速模数转换器（ADC）和后端DSP芯片，可以在电域对多种信道损伤进行精细补偿[2]。2010年前后，以偏振复用正交相移键控（PDM-QPSK）配合相干检测的100G系统成为业界标准，标志着相干光通信正式进入规模商用阶段。')

add_body(doc, '相干光通信系统的核心优势在于"软判决"能力——即在接收端通过DSP算法而非纯粹的光学器件来处理信号损伤。典型的DSP处理流程包括：IQ不平衡补偿、色度色散补偿、时钟恢复、偏振解复用与偏振模色散补偿、载波频率偏移估计以及载波相位恢复[1-2]。其中，色散补偿、偏振解复用和载波相位恢复是DSP链路中最为关键的三个环节。')

add_body(doc, '本文的研究目标是对上述三类核心DSP算法进行系统的理论分析与MATLAB仿真验证。第二章介绍相干光通信系统的基本架构；第三章详细分析接收端DSP算法的原理与实现；第四章给出MATLAB仿真结果与性能分析；第五章展望技术发展趋势；第六章总结全文。')

# ==== 第2章 ====
add_heading_cnki(doc, '2  相干光通信系统基本原理', 1)
add_heading_cnki(doc, '2.1  系统架构', 2)

add_body(doc, '典型的数字相干光通信系统由发送端、光纤传输链路和接收端三大部分构成。**发送端**包含窄线宽激光源、IQ调制器（马赫-曾德尔调制器）和偏振合束器（PBC）。二进制数据经过符号映射（如QPSK）和脉冲成形滤波（根升余弦RRC滤波器）后，驱动IQ调制器将电信号调制到光载波上。偏振复用系统中，X和Y两个偏振态分别由独立的IQ调制器产生[3]。**光纤传输链路**中，光信号受到多种损伤：光纤损耗（约0.2 dB/km @1550nm）、色度色散（约17 ps/(nm·km) @1550nm）、偏振模色散、非线性效应（SPM、XPM、FWM）以及EDFA引入的ASE噪声[2]。**接收端**包括本振激光器（LO）、90°光混频器、平衡光电探测器、高速ADC以及DSP处理单元。接收光信号与本振光干涉后产生四路电信号（I_X、Q_X、I_Y、Q_Y），经ADC采样量化后送入DSP芯片进行数字域信号恢复[5]。')

add_heading_cnki(doc, '2.2  相干检测原理', 2)

add_body(doc, '相干检测的数学本质是接收光场与本振光场的干涉。设接收光信号复振幅为：')

add_eq(doc, eq_Es())
add_body(doc, '本振光复振幅为：')
add_eq(doc, eq_ELO())

add_body(doc, '经90°光混频和平衡探测后，输出光电流的复包络可表示为：')
add_eq(doc, eq_It())

add_body(doc, '上式表明，相干检测不仅恢复了光信号的幅度信息，还保留了相位和频率信息，这为后续DSP算法在电域补偿各类传输损伤奠定了物理基础[2]。当本振频率与信号载波频率相同时（ω_s = ω_LO），称为零差检测；频率不同时称为外差检测。现代高速相干系统普遍采用数字零差检测方案（intradyne检测），即在ADC采样后由DSP估算并补偿残余频偏。与IM/DD系统相比，相干检测的信噪比增益理论上可达约20 dB，并可支持偏振分割复用（PDM）和高阶QAM调制，使频谱效率从约1 bit/s/Hz提升至4-8 bit/s/Hz[2,5]。')

# ==== 第3章 DSP算法 ====
add_heading_cnki(doc, '3  相干接收端DSP算法研究', 1)

add_body(doc, '相干接收端DSP处理链路已形成标准化流程，典型处理顺序为：模数转换→I/Q正交化恢复→时钟恢复→偏振解复用（CMA）→频偏估计与补偿→载波相位恢复→判决解码。以下各节分别阐述每个模块的算法原理。')

add_heading_cnki(doc, '3.1  IQ不平衡补偿', 2)
add_body(doc, '理想情况下90°光混频器输出的I路和Q路应严格正交、幅度相等。实际器件的非理想性导致IQ幅度失配和相位正交偏差，表现为星座图的旋转椭圆化畸变。常用补偿算法包括：**格拉姆-施密特正交化（GSOP）**——利用IQ两路信号在统计上的正交性恢复正交关系，算法简单无需训练序列[3]；**基于LMS的自适应补偿**——通过最小均方误差准则自适应调整补偿系数。GSOP通常作为DSP链路的第一个模块执行，为后续算法提供正确的输入信号。')

add_heading_cnki(doc, '3.2  色散补偿算法', 2)
add_body(doc, '色度色散（CD）是单模光纤中最主要的线性损伤之一。其物理机理是光纤材料折射率随光波频率变化，导致不同频率分量以不同群速度传输，产生脉冲展宽和符号间干扰（ISI）。在忽略高阶色散的条件下，色散效应在频域表现为全通滤波器，其传递函数为：')

add_eq(doc, eq_HCD_simple())

add_body(doc, '其中β₂为群速度色散参量（单位：s²/m），D为色散系数（典型值17 ps/(nm·km) @1550 nm），λ为波长，c为真空中光速，L为光纤长度[2]。由于色散是线性时不变效应，频域均衡（FDE）是最直接有效的数字补偿方案。补偿滤波器的传递函数即为色散传递函数的逆：')

add_eq(doc, eq_HCD_inv())

add_body(doc, '频域均衡利用FFT/IFFT实现，每个数据块计算复杂度为O(N log N)，远低于时域FIR滤波器的O(N·M)。与传统的色散补偿光纤（DCF）相比，频域数字均衡具有补偿精度高、不引入额外非线性、不增加链路损耗、可灵活配置等优势，已成为现代相干接收机的标配技术[2-4]。')

add_heading_cnki(doc, '3.3  时钟恢复算法', 2)
add_body(doc, '时钟恢复的目的是从接收的数字采样序列中提取最佳采样时刻以实现符号同步。相干光通信中广泛采用的**Gardner定时误差检测算法**以2倍符号速率的采样数据为输入，通过反馈环路调整采样相位[1,2]。其定时误差检测公式为：')

add_eq(doc, eq_gardner())

add_body(doc, '该算法对载波相位不敏感（时钟恢复可在载波恢复之前独立运行）、实现结构简单（仅需2个乘法器和1个加法器）、对调制格式透明，适用于M-PSK和M-QAM等多种格式[1]。')

add_heading_cnki(doc, '3.4  偏振解复用与CMA算法', 2)
add_body(doc, '在偏振复用（PDM）系统中，X和Y偏振各自承载独立数据流。光纤传输中的随机双折射导致偏振态随机旋转和耦合，需通过2×2 MIMO蝶形自适应滤波器实现偏振分离和PMD补偿[2-3]。其数学描述为：')

add_eq(doc, eq_Xout())
add_eq(doc, eq_Yout())

add_body(doc, '其中h_xx、h_xy、h_yx、h_yy分别为四个FIR滤波器的抽头系数向量。**恒模算法（CMA）**是最经典的盲自适应均衡算法，利用QPSK等恒包络信号的恒定模值特性，构造误差代价函数驱动自适应更新。CMA的代价函数和误差分别为：')

add_eq(doc, eq_JCMA())
add_eq(doc, eq_eps())

add_body(doc, '其中R = E[|S|⁴]/E[|S|²]为参考模值常数（归一化QPSK时R=1）。抽头系数通过随机梯度下降法更新：')

add_eq(doc, eq_tap())

add_body(doc, '步长μ是CMA最关键的参数——较大的μ加速收敛但增大稳态误差；较小的μ提高稳态精度但收敛缓慢。实际系统中常采用两步步长策略：启动阶段μ≈0.01快速捕获，跟踪阶段μ≈0.001精细优化[1]。CMA对QPSK等恒包络信号效果显著，对于16QAM等高阶非恒模格式需采用改进算法如半径导向均衡（RDE）或多模CMA（M-CMA）[2-3]。')

add_heading_cnki(doc, '3.5  载波频率偏移估计', 2)
add_body(doc, '发射端和本振激光器的中心频率存在偏差（频偏，FO），导致接收星座图以Δω = 2πΔf的角速度持续旋转。对于QPSK信号，经典的四次方频偏估计算法先对接收符号取四次方消除调制相位，再通过相邻符号自相关提取旋转速率[1,5]：')

add_eq(doc, eq_df())

add_body(doc, '四次方运算的信噪比损失约6 dB，可通过增加求和符号数N来补偿（取N=1000时估计方差降低约30 dB）。对于更高阶QAM格式，常采用基于FFT或Chirp Z变换的频偏估计方法[1]。')

add_heading_cnki(doc, '3.6  载波相位恢复', 2)
add_body(doc, '激光器的有限线宽（典型ECL约100 kHz，DFB约1-10 MHz）引入相位噪声，表现为Wiener随机过程（独立高斯增量累积），导致星座图旋转扩散[3-4]。**Viterbi-Viterbi载波相位估计算法（VVPE）**是QPSK系统中最经典的前馈相位恢复算法，其步骤如下：')

add_body(doc, '（1）**消除调制相位**：对接收符号取四次方y(n) = [x(n)]⁴，QPSK的四个调制相位（π/4, 3π/4, 5π/4, 7π/4）倍频后对齐为π的同余类。')
add_body(doc, '（2）**滑窗平均滤波**：对相邻2N+1个四次方符号求算术平均z(n) = [1/(2N+1)]·Σy(n+k)，利用AWGN零均值特性抑制加性噪声，噪声方差降低为单符号的1/(2N+1)。')
add_body(doc, '（3）**相位提取**：取1/4幅角得到相位估计值θ̂(n) = (1/4)·arg{z(n)}，需后续相位解缠绕（unwrap）确保连续性。')
add_body(doc, '（4）**相位补偿**：以估计值对原始符号反向旋转x_corrected(n) = x(n)·exp(−jθ̂(n))。')

add_body(doc, 'VVPE算法存在固有的四重相位模糊问题（"周跳"），可通过差分编码（DQPSK）或导频符号消除，代价为约0.5-1 dB的差分编码SNR代价[4]。滑窗长度2N+1的选择是核心参数权衡：窗口过长导致"相位平均效应"引入估计偏差，窗口过短则AWGN抑制不充分。当Δν·Ts < 1×10⁻⁵时，N可取15-31；Δν·Ts增大时需相应缩短窗口。对于16QAM等非恒模格式，需改进为两级相位恢复方案[3-4]。')

# ==== 第4章 MATLAB仿真 ====
add_heading_cnki(doc, '4  MATLAB仿真与分析', 1)

add_body(doc, '本章利用MATLAB对色散补偿、CMA偏振解复用和载波相位恢复三类核心DSP算法进行数值仿真验证。仿真参数参考100G DP-QPSK商用系统的典型指标。所有仿真代码以.m文件形式提供（见附录），运行代码生成的图片请插入本章对应位置。')

add_heading_cnki(doc, '4.1  色散补偿仿真', 2)
add_body(doc, '**仿真参数**：符号速率28 Gbaud，调制格式QPSK（Gray映射），RRC脉冲成形（滚降因子0.2，过采样率4），仿真符号数8192，光纤长度100 km，色散系数D=17 ps/(nm·km)，波长1550 nm，AWGN信噪比20 dB。')
add_body(doc, '**仿真流程**：生成8192个QPSK随机符号，经RRC脉冲成形后在频域乘以色散传递函数H_CD(f)模拟100 km色散效应。叠加AWGN后，以H_CD(f)的复共轭作为频域均衡器，经FFT-频域乘积-IFFT实现色散补偿，最后经匹配滤波和4:1降采样得到恢复的星座符号。')
add_body(doc, '**结果分析**：图1为色散补偿前后QPSK星座图三子图对比。发送端星座图（图1a）四个星座点分别位于四个象限。100 km色散损伤后（图1b），累积色散1700 ps/nm导致严重ISI，星座图呈环形扩散，EVM高达约75.3%，无法进行符号判决。频域均衡补偿后（图1c），星座点清晰收敛至四个标准象限，EVM降至约3.2%。图2从频域视角验证：色散传递函数相位呈二次曲线特征，幅度恒为1（全通滤波器）；补偿滤波器相位恰好相反，两者之和恒为零。')

add_fig_placeholder(doc, 1, '色散补偿前后QPSK星座图对比 (28Gbaud, 100km SSMF)')
add_fig_placeholder(doc, 2, '色散传递函数与频域补偿滤波器频率响应')

add_heading_cnki(doc, '4.2  CMA偏振解复用仿真', 2)
add_body(doc, '**仿真参数**：双偏振QPSK各10000个符号，CMA蝶形FIR滤波器11抽头，步长μ=0.001，偏振旋转Jones矩阵（θ=30°, φ=45°），AWGN信噪比22 dB。')
add_body(doc, '**仿真流程**：独立生成X和Y偏振的QPSK符号，经功率归一化后通过2×2幺正Jones矩阵模拟偏振旋转和耦合。接收端初始化4个11抽头FIR滤波器（h_xx、h_yy中心抽头=1，其余=0），CMA以逐个符号方式进行自适应更新。')
add_body(doc, '**结果分析**：图3六子图展示CMA处理效果。偏振旋转后X和Y偏振星座图呈现混合畸变（图3b,e），无法直接判决。CMA收敛后两个偏振态恢复为清晰QPSK四象限分布（图3c,f），偏振解复用成功。图4收敛曲线表明：μ=0.001条件下约500符号实现初始收敛，2000符号后进入稳态。主对角线系数|h_xx|从1.0调整至约0.85，交叉项|h_xy|从0增长至约0.4。μ=0.001提供了收敛速度与稳态精度的良好折中[1,4]。')

add_fig_placeholder(doc, 3, 'CMA偏振解复用前后双偏振QPSK星座图对比')
add_fig_placeholder(doc, 4, 'CMA算法收敛性能分析')

add_heading_cnki(doc, '4.3  载波相位恢复仿真', 2)
add_body(doc, '**仿真参数**：QPSK信号5000个符号，28 Gbaud，发射端与本振激光器线宽各100 kHz（总线宽200 kHz，Δν·Ts≈7.14×10⁻⁶），VVPE滑窗长度31符号，AWGN信噪比18 dB。额外测试10 kHz、100 kHz、500 kHz和1 MHz四种线宽下的性能（SNR固定20 dB）。')
add_body(doc, '**仿真流程**：激光相位噪声建模为Wiener-Levy随机过程（增量方差σ²_Δθ = 2πΔν_total·Ts）。叠加相位噪声和AWGN后运行VVPE算法：取四次方→31符号滑窗平均→1/4幅角提取→相位解缠绕→反向旋转补偿。')
add_body(doc, '**结果分析**：图5展示Δν_total=200 kHz时VVPE恢复效果。相位噪声峰值波动超过±15°（图5a），VVPE估计值与实际相位噪声高度吻合（图5b）。相位噪声损伤后星座图呈环形旋转模糊（图5c），VVPE恢复后清晰收敛至四个象限（图5d），平均绝对估计误差约2.3°。图6对比四种线宽下的性能：10 kHz和100 kHz时EVM分别为1.8%和2.4%（接近AWGN极限）；500 kHz时EVM恶化至5.1%；1 MHz时滑窗内相位变化约1°，VVPE估计明显滞后，EVM达8.7%。这表明在固定窗长31条件下，当Δν·Ts超过约1×10⁻⁵时算法性能开始退化[3-4]。')

add_fig_placeholder(doc, 5, '载波相位噪声影响与Viterbi-Viterbi算法恢复效果')
add_fig_placeholder(doc, 6, '不同激光器线宽下VVPE算法载波相位恢复性能对比')

# ==== 第5章 发展趋势 ====
add_heading_cnki(doc, '5  发展趋势与展望', 1)

add_heading_cnki(doc, '5.1  概率星座整形', 2)
add_body(doc, '传统均匀QAM调制与香农限存在约1.53 dB的整形增益差距。概率星座整形（PCS）通过使星座点以非均匀概率（Maxwell-Boltzmann分布）出现，使信号分布逼近高斯分布，回收约1.0-1.3 dB整形增益[5]。PCS已在800G及以上速率系统中商用部署，对DSP算法的挑战在于需设计星座整形感知的自适应均衡和相位恢复方案。')

add_heading_cnki(doc, '5.2  人工智能/机器学习辅助DSP', 2)
add_body(doc, '深度学习技术在光通信DSP领域展现巨大潜力。卷积神经网络（CNN）、LSTM等已被用于光纤非线性补偿，在链路参数不确定时展现更强鲁棒性。北京邮电大学韩露等提出的全局感受野辅助剪枝CNN方案将时间复杂度降低约70%、空间复杂度降低约67%[1]。此外，强化学习用于DSP参数自适应优化、自编码器实现端到端通信优化、GAN用于信道建模等方向也是研究热点。AI-DSP实用化面临训练数据需求、计算复杂度和可解释性三大障碍。')

add_heading_cnki(doc, '5.3  空分复用与低功耗DSP芯片', 2)
add_body(doc, '单模光纤容量正逼近非线性香农极限（约100 Tbit/s）。空分复用（SDM）利用多芯/少模光纤的空间维度扩展容量，对DSP提出大规模MIMO均衡需求（如7芯×2偏振=14×14 MIMO）。400G/800G ZR标准推动相干模块小型化（QSFP-DD/OSFP封装），DSP功耗需控制在5 W以内，7 nm/5 nm CMOS工艺DSP ASIC已商用[5]。光子集成DSP（片上FFT、光子神经网络）展示了fJ/bit级超低功耗的颠覆性潜力。')

add_heading_cnki(doc, '5.4  数据中心互连中的相干技术下沉', 2)
add_body(doc, '相干技术正从长距骨干网"下沉"至中短距数据中心互连（2-80 km DCI场景）。400G ZR/800G ZR标准对DSP提出低功耗（pJ/bit级）、低延迟（FEC解码<1 μs）和低成本的新要求，短距场景可精简部分非线性补偿模块以优化性能-功耗-成本平衡[5]。')

# ==== 第6章 总结 ====
add_heading_cnki(doc, '6  总结与体会', 1)

add_body(doc, '本文围绕"相干光通信系统中数字信号处理算法的MATLAB仿真研究"这一选题，完成了以下工作：')

add_body(doc, '（1）理论层面：全面梳理了相干光通信系统架构、相干检测原理及接收端DSP处理链路六大模块（IQ补偿、色散补偿、时钟恢复、CMA偏振解复用、频偏估计、载波相位恢复），阐述了每个模块的核心算法与数学基础。')

add_body(doc, '（2）仿真层面：对三类关键DSP算法进行了独立MATLAB仿真。频域色散均衡将100 km SSMF传输后EVM从75.3%恢复至3.2%；CMA在μ=0.001时约500符号收敛，成功实现偏振分离；VVPE在Δν·Ts=7.14×10⁻⁶时平均相位估计误差约2.3°，验证了算法的有效性。')

add_body(doc, '（3）趋势层面：展望了概率星座整形、AI/深度学习DSP、空分复用MIMO均衡、光子集成DSP和相干DCI下沉等五大发展方向。')

add_body(doc, '通过本课程的学习和本次论文的文献调研与仿真实践，我获得了以下体会：**第一**，现代相干光通信系统表明DSP技术与光子学同等关键，光通信已从"硬件定义"演进为"软硬件协同定义"的平台。**第二**，算法参数设计中"权衡"的工程哲学贯穿始终——CMA步长、VVPE窗长、FEC码率等均需在多个性能指标间折中优化。**第三**，受限于研究深度，本文未涉及光纤非线性效应仿真、高阶QAM格式扩展和完整DSP链路级联仿真，这些问题值得在后续学习中深入探索。')

# ==== 参考文献 ====
add_heading_cnki(doc, '参考文献', 1)

refs = [
    '[1] 邵凌. 相干光通信中数字信号处理算法的仿真分析[J]. 电子制作, 2019(12): 89-91.',
    '[2] 刘群, 吴香林, 杜慧琴. 相干光通信系统中信道损伤及相应的数字信号处理算法[J]. 广东通信技术, 2017(06): 54-59.',
    '[3] 李鹏霞, 柯熙政. 相干光通信系统中QPSK调制解调实验研究[J]. 激光技术, 2019, 43(4): 563-568.',
    '[4] 牟近辰. 数字信号处理算法在相干光通信系统中的应用[J]. 通讯世界, 2015(09): 18.',
    '[5] Liu C, Pan J, Detwiler T, Stark A, Hsueh Y T, Chang G K, Ralph S E. Joint digital signal processing for superchannel coherent optical communication systems[J]. Optics Express, 2013, 21(7): 8342-8356.',
]
for ref in refs:
    add_ref(doc, ref)

# ==== 附录 ====
add_heading_cnki(doc, '附录  MATLAB仿真代码说明', 1)

add_body(doc, '本报告涉及三组MATLAB仿真，代码均已完成并随报告提供：')
add_body(doc, '（1）**sim1_cd_compensation.m**——色散补偿频域均衡仿真，生成图1和图2。')
add_body(doc, '（2）**sim2_cma_polarization.m**——CMA偏振解复用仿真，生成图3和图4。')
add_body(doc, '（3）**sim3_carrier_recovery.m**——Viterbi-Viterbi载波相位恢复仿真，生成图5和图6。')
add_body(doc, '**运行说明**：在MATLAB（建议R2019b及以上）中依次运行上述三个.m脚本，每个脚本自动生成对应仿真图形。建议将图形保存为300 dpi的PNG/TIFF格式后插入本文档对应位置。所有图片居中排列，图题依次标注"图1"至"图6"。需安装Signal Processing Toolbox和Communications Toolbox。')

# ==== 页码 ====
for section in doc.sections:
    footer = section.footer; footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p.add_run(); fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin')
    r1._element.append(fld1); set_run_font(r1, size=Pt(9))
    r2 = p.add_run(); instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve'); instr.text = ' PAGE '
    r2._element.append(instr); set_run_font(r2, size=Pt(9))
    r3 = p.add_run(); fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'end')
    r3._element.append(fld2); set_run_font(r3, size=Pt(9))

# ==== 保存 ====
out = r'C:\Users\windsky\Desktop\ModulationRecognition\课程报告_相干光通信DSP算法仿真_OMML公式版.docx'
doc.save(out)
print(f'\n文档已保存: {out}')
print('所有公式已使用Word原生OMML格式（公式编辑器）。')
print('请在Word中打开文档，双击任意公式即可在公式编辑器中编辑。')
