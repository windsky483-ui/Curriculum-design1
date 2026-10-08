"""
生成光纤通信课程报告的Word文档
格式要求：正文小四宋体、英文Times New Roman、行间距20磅
"""
import re
from docx import Document
from docx.shared import Pt, Inches, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_run_font(run, font_name_cn='宋体', font_name_en='Times New Roman', size=Pt(12)):
    """设置run的字体"""
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
    """设置段落行间距为固定值（磅）"""
    pPr = paragraph._element.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:line'), str(int(spacing_pt * 20)))  # 20磅 = 400 (单位: 1/20磅)
    spacing.set(qn('w:lineRule'), 'exact')
    pPr.insert(0, spacing)

def set_paragraph_spacing(paragraph, before=0, after=0):
    """设置段前段后间距"""
    pf = paragraph.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)

def add_body_paragraph(doc, text, bold=False, alignment=None, font_size=Pt(12)):
    """添加正文段落（小四宋体，20磅行间距）"""
    p = doc.add_paragraph()
    set_line_spacing(p, 20)
    if alignment is not None:
        p.alignment = alignment

    # 处理文本中的 **粗体** 标记和行内代码标记
    parts = re.split(r'(\*\*.*?\*\*|`.*?`)', text)
    for part in parts:
        if part.startswith('**') and part.endswith('**'):
            run = p.add_run(part[2:-2])
            run.bold = True
            set_run_font(run, size=font_size)
        elif part.startswith('`') and part.endswith('`'):
            run = p.add_run(part[1:-1])
            set_run_font(run, font_name_cn='仿宋', font_name_en='Consolas', size=Pt(10))
        else:
            run = p.add_run(part)
            set_run_font(run, size=font_size)
            if bold:
                run.bold = True
    return p

def add_heading_styled(doc, text, level=1):
    """添加标题（黑体）"""
    if level == 1:
        size = Pt(16)  # 小三
    elif level == 2:
        size = Pt(14)  # 四号
    else:
        size = Pt(12)  # 小四
    p = doc.add_paragraph()
    set_line_spacing(p, 22)
    set_paragraph_spacing(p, before=6, after=3)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.bold = True
    set_run_font(run, font_name_cn='黑体', font_name_en='Times New Roman', size=size)
    return p

def add_title(doc, text):
    """添加论文题目（二号黑体居中）"""
    p = doc.add_paragraph()
    set_line_spacing(p, 28)
    set_paragraph_spacing(p, before=12, after=12)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    set_run_font(run, font_name_cn='黑体', font_name_en='Times New Roman', size=Pt(18))
    return p

def add_placeholder(doc, fig_num, description):
    """添加图片占位符"""
    p = doc.add_paragraph()
    set_line_spacing(p, 20)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'[ 此处插入图{fig_num}：{description} ]')
    run.italic = True
    set_run_font(run, font_name_cn='楷体', font_name_en='Times New Roman', size=Pt(10.5))
    # 图片框
    p2 = doc.add_paragraph()
    set_line_spacing(p2, 120)  # 给图片留空间
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run('↑ 请在此处插入MATLAB仿真生成的图片 ↑')
    run2.font.color.rgb = None
    set_run_font(run2, font_name_cn='楷体', font_name_en='Times New Roman', size=Pt(9))
    run2.italic = True

def add_figure_caption(doc, fig_num, caption):
    """添加图例（居中，小五号）"""
    p = doc.add_paragraph()
    set_line_spacing(p, 16)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'图{fig_num}  {caption}')
    set_run_font(run, size=Pt(10.5))

def add_code_block(doc, code_text):
    """添加代码块"""
    for line in code_text.strip().split('\n'):
        p = doc.add_paragraph()
        set_line_spacing(p, 16)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(line)
        set_run_font(run, font_name_cn='仿宋', font_name_en='Consolas', size=Pt(9))
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)

def add_ref_item(doc, text):
    """添加参考文献条目"""
    p = doc.add_paragraph()
    set_line_spacing(p, 18)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    # 处理 [N] 编号
    run = p.add_run(text)
    set_run_font(run, size=Pt(10.5))

def add_section_break(doc):
    """添加分节符"""
    p = doc.add_paragraph()
    set_line_spacing(p, 20)
    run = p.add_run('—' * 30)
    set_run_font(run, size=Pt(10))
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

# ============================================================
# 生成主报告文档
# ============================================================
print("正在生成课程报告Word文档...")

doc = Document()

# 设置默认字体
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(12)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

# 页面设置 A4
section = doc.sections[0]
section.page_width = Cm(21)
section.page_height = Cm(29.7)
section.top_margin = Cm(2.54)
section.bottom_margin = Cm(2.54)
section.left_margin = Cm(3.18)
section.right_margin = Cm(3.18)

# ===== 标题 =====
add_title(doc, '相干光通信系统中数字信号处理算法的MATLAB仿真研究')

# ===== 摘要 =====
add_heading_styled(doc, '摘要', level=2)
add_body_paragraph(doc,
    '相干光通信技术是现代高速光纤通信网络的核心支撑技术之一。与传统的强度调制/直接检测（IM/DD）'
    '系统相比，相干检测技术具有更高的接收灵敏度、频谱效率以及支持多维度调制的优势，已成为100G及'
    '以上速率光纤通信系统的标准方案。在相干光通信接收端，数字信号处理（DSP）算法承担着补偿各类'
    '信道损伤的关键任务，包括色度色散补偿、偏振解复用、载波频率偏移估计及相位恢复等。本文系统梳理'
    '了相干光通信系统的基本原理与接收端DSP算法的标准流程，重点研究了三类核心算法：频域色散补偿、'
    '基于恒模算法（CMA）的偏振解复用，以及基于Viterbi-Viterbi算法的载波相位恢复。利用MATLAB软件'
    '对上述算法进行了数值仿真分析，通过星座图对比、收敛曲线分析及不同参数下的性能评估，验证了DSP'
    '算法在相干光通信系统中的有效性。最后，对相干光通信DSP技术的未来发展趋势进行了展望。')

p = doc.add_paragraph()
set_line_spacing(p, 20)
run = p.add_run('关键词：')
run.bold = True
set_run_font(run, size=Pt(12))
run2 = p.add_run('相干光通信；数字信号处理；色散补偿；恒模算法；载波相位恢复；偏振解复用')
set_run_font(run2, size=Pt(12))

# ===== 第1章 引言 =====
add_heading_styled(doc, '1  引言', level=1)

add_body_paragraph(doc,
    '随着云计算、人工智能、5G/6G移动通信及超高清视频等带宽密集型应用的迅猛发展，全球网络数据'
    '流量呈指数级增长。据思科年度互联网报告预测，全球IP流量年增长率保持在20%以上，这给光纤通信'
    '网络的传输容量带来了持续挑战[1]。在过去的三十年中，光纤通信系统经历了从准同步数字体系（PDH）'
    '到同步数字体系（SDH），再到光传送网（OTN）的多代演进，单波长传输速率已从2.5 Gbit/s提升至'
    '400 Gbit/s乃至800 Gbit/s，波分复用（WDM）技术使单纤传输容量突破100 Tbit/s量级。')

add_body_paragraph(doc,
    '推动这一进步的三大关键技术分别是：波分复用技术（WDM）、光放大技术（EDFA/Raman）以及相干光'
    '通信技术。其中，相干光通信技术的复兴与实用化（约2005年以后）具有里程碑意义。与传统的IM/DD'
    '方案相比，相干检测能够完整保留光场的幅度、相位和偏振态信息，结合高速模数转换器（ADC）和后端'
    'DSP芯片，可以在电域对多种信道损伤进行精细补偿[2]。2010年前后，以偏振复用正交相移键控'
    '（PDM-QPSK）配合相干检测的100G系统成为业界标准，标志着相干光通信正式进入规模商用阶段。')

add_body_paragraph(doc,
    '相干光通信系统的核心优势在于"软判决"能力——即在接收端通过DSP算法而非纯粹的光学器件来处理'
    '信号损伤。典型的DSP处理流程包括：IQ不平衡补偿、色度色散（CD）补偿、时钟恢复、偏振解复用与'
    '偏振模色散（PMD）补偿、载波频率偏移估计以及载波相位恢复[3-4]。其中，色散补偿、偏振解复用和'
    '载波相位恢复是DSP链路中最为关键的三个环节，直接决定了接收信号的质量和系统误码率（BER）性能。')

add_body_paragraph(doc,
    '本文的研究目标是对上述三类核心DSP算法进行系统的理论分析与MATLAB仿真验证。第二章介绍相干光'
    '通信系统的基本架构；第三章详细分析接收端DSP算法的原理与实现；第四章给出MATLAB仿真结果与性能'
    '分析；第五章展望技术发展趋势；第六章总结全文。')

# ===== 第2章 =====
add_heading_styled(doc, '2  相干光通信系统基本原理', level=1)
add_heading_styled(doc, '2.1  系统架构', level=2)

add_body_paragraph(doc,
    '典型的数字相干光通信系统由发送端、光纤传输链路和接收端三大部分构成。以下分别介绍各部分的基本'
    '组成和功能。')

add_body_paragraph(doc,
    '发送端包含：激光源（窄线宽外腔激光器ECL）、IQ调制器（马赫-曾德尔调制器MZM）、偏振合束器'
    '（PBC）。在发送端，待传输的二进制数据首先经过符号映射（如QPSK映射为4个相位状态），再通过'
    '脉冲成形滤波器（常用根升余弦RRC滤波器）限制信号带宽。成形后的基带信号驱动IQ调制器，将电信'
    '号调制到光载波上。偏振复用系统中，两个偏振态（X和Y偏振）分别由独立的IQ调制器产生，经偏振'
    '合束器合成后送入光纤链路[5]。')

add_body_paragraph(doc,
    '光纤传输链路包含：标准单模光纤（SSMF，G.652）、掺铒光纤放大器（EDFA）等。光信号在光纤中'
    '传输时会受到多种损伤，主要包括：光纤损耗（约0.2 dB/km @1550nm）、色度色散（约17 ps/(nm·km) '
    '@1550nm）、偏振模色散（均值约0.1 ps/√km）、非线性效应（自相位调制SPM、交叉相位调制XPM、'
    '四波混频FWM等）以及光放大器引入的自发辐射噪声（ASE噪声）[5-6]。')

add_body_paragraph(doc,
    '接收端是相干光通信系统最复杂的部分，包括：本振激光器（LO）、90°光混频器、平衡光电探测器'
    '（BPD）、高速ADC以及DSP处理单元。接收光信号与本振光在90°光混频器中干涉，输出四路光信号'
    '（I_X, Q_X, I_Y, Q_Y），经平衡探测后得到与光场复振幅成比例的电信号。这些模拟电信号由高速ADC'
    '（采样率通常为符号速率的2倍以上）采样量化后，送入DSP芯片进行数字域的信号恢复[7]。')

add_heading_styled(doc, '2.2  相干检测原理', level=2)

add_body_paragraph(doc,
    '相干检测的数学本质是接收光场与本振光场的干涉。设接收光信号复振幅为：')

add_body_paragraph(doc,
    'E_s(t) = A_s(t) · exp[j(ω_s·t + θ_s(t))]', alignment=WD_ALIGN_PARAGRAPH.CENTER)

add_body_paragraph(doc, '本振光复振幅为：')

add_body_paragraph(doc,
    'E_LO(t) = A_LO · exp[j(ω_LO·t + θ_LO(t))]', alignment=WD_ALIGN_PARAGRAPH.CENTER)

add_body_paragraph(doc,
    '经90°光混频和平衡探测后，输出光电流的复包络可表示为：')

add_body_paragraph(doc,
    'I(t) ∝ A_LO · A_s(t) · exp[j(ω_s - ω_LO)t + j(θ_s(t) - θ_LO(t))]',
    alignment=WD_ALIGN_PARAGRAPH.CENTER)

add_body_paragraph(doc,
    '上式表明，相干检测不仅恢复了光信号的幅度信息，还保留了相位和频率信息，这为后续DSP算法在电域'
    '补偿各类传输损伤奠定了物理基础[2]。当本振频率与信号载波频率相同时（ω_s = ω_LO），称为零差检测；'
    '频率不同时称为外差检测。现代高速相干系统普遍采用数字零差检测方案（也称为intradyne检测），即在'
    'ADC采样后由DSP估算并补偿残余频偏。')

# ===== 第3章 =====
add_heading_styled(doc, '3  相干接收端DSP算法研究', level=1)

add_body_paragraph(doc,
    '相干接收端DSP处理链路的标准化流程已由学术界和工业界达成广泛共识。尽管不同厂商的具体实现存在'
    '差异，但核心算法模块和基本处理顺序保持一致。典型的DSP处理流程为：模数转换→I/Q正交化恢复→'
    '时钟恢复→偏振解复用（CMA）→频偏估计与补偿→载波相位恢复→判决解码。以下逐模块分析各算法的'
    '原理与实现。')

add_heading_styled(doc, '3.1  IQ不平衡补偿', level=2)

add_body_paragraph(doc,
    '理想情况下，90°光混频器输出的I路和Q路应严格正交、幅度相等。然而，实际器件存在非理想性，导致'
    'IQ幅度失配和相位正交偏差，表现为星座图的"旋转椭圆化"畸变。常用的补偿算法包括：')

add_body_paragraph(doc,
    '格拉姆-施密特正交化（GSOP）算法：利用I/Q两路信号在统计上的正交性，通过数学变换恢复正交关系，'
    '算法简单且无需训练序列[8]。基于LMS的自适应补偿：通过最小均方误差准则自适应调整补偿系数，适用'
    '于连续跟踪时变的IQ失衡。GSOP通常作为DSP链路的第一个模块执行，为后续算法提供正确的输入信号。')

add_heading_styled(doc, '3.2  色散补偿算法', level=2)

add_body_paragraph(doc,
    '色度色散（CD）是单模光纤中最主要的线性损伤之一。其物理机理是光纤折射率随波长变化，导致不同'
    '频率分量以不同的群速度传输，造成脉冲展宽和符号间干扰（ISI）。色散在频域表现为一个全通滤波器，'
    '其传递函数为：')

add_body_paragraph(doc,
    'H_CD(ω, L) = exp(-j·β₂·ω²·L / 2) = exp(j·λ²·D·ω²·L / (4πc))',
    alignment=WD_ALIGN_PARAGRAPH.CENTER)

add_body_paragraph(doc,
    '其中，β₂为群速度色散参量，D为色散系数，λ为波长，c为光速，L为光纤长度[6]。')

add_body_paragraph(doc,
    '由于色散是线性时不变效应，频域均衡（FDE）是最直接有效的补偿方案。补偿滤波器取色散传递函数的'
    '逆函数：H_CD^(-1)(ω, L) = exp(j·β₂·ω²·L / 2)。频域均衡的实现采用快速傅里叶变换（FFT）和逆变换'
    '（IFFT），结合重叠保留或重叠相加方法处理长序列数据。与传统的色散补偿光纤（DCF）方案相比，频域'
    '数字均衡具有补偿精度高、不引入额外非线性和损耗、灵活可调等优点，已成为相干接收机的标配技术[6,8]。')

add_heading_styled(doc, '3.3  时钟恢复算法', level=2)

add_body_paragraph(doc,
    '时钟恢复的目的是从接收的数字采样序列中提取最佳采样时刻，以实现符号同步。相干光通信中常用的时钟'
    '恢复算法为Gardner定时误差检测算法。该算法以2倍符号速率的采样数据为输入，通过计算相邻符号间的'
    '定时误差来驱动数控振荡器（NCO）调整采样相位，形成反馈环路[3,7]。Gardner算法的定时误差检测公式'
    '为：e(n) = Re{[x(n) - x(n-2)] · x*(n-1)}。该算法的优势在于对载波相位不敏感（即时钟恢复可在载波'
    '恢复之前运行），且实现结构简单，适合硬件实现。')

add_heading_styled(doc, '3.4  偏振解复用与CMA算法', level=2)

add_body_paragraph(doc,
    '在偏振复用（PDM）系统中，两个正交偏振态各自承载独立的信号。光纤传输中的随机双折射效应导致偏振'
    '态耦合和旋转，接收端需通过2×2多输入多输出（MIMO）蝶形自适应滤波器实现偏振解复用和PMD补偿[9]。')

add_body_paragraph(doc,
    '恒模算法（CMA）是最经典和应用最广泛的盲自适应均衡算法，由Godard和Treichler等人提出并完善。CMA'
    '的核心思想是利用QPSK等恒包络调制格式信号的恒定模值特性，构造与信号模值相关的误差代价函数来驱动'
    '自适应更新。CMA不需要训练序列，属于盲均衡算法[9-10]。')

add_body_paragraph(doc,
    '2×2 MIMO蝶形滤波器的输出为：')

add_body_paragraph(doc,
    'X_out(n) = h_xx^H · X_in + h_xy^H · Y_in',
    alignment=WD_ALIGN_PARAGRAPH.CENTER)
add_body_paragraph(doc,
    'Y_out(n) = h_yx^H · X_in + h_yy^H · Y_in',
    alignment=WD_ALIGN_PARAGRAPH.CENTER)

add_body_paragraph(doc,
    '其中h_xx, h_xy, h_yx, h_yy为四个FIR滤波器的抽头系数向量。CMA的代价函数和误差分别为：'
    'J_CMA = E[(|X_out|² - R)²]，ε_x = R - |X_out|²，其中R = E[|S|⁴]/E[|S|²]为参考模值常数（QPSK'
    '归一化后R=1）。抽头系数通过随机梯度下降法更新：h_ij ← h_ij + μ · ε · out · in*。步长μ决定了算法'
    '的收敛速度和稳态精度。较大的μ加速收敛但增大稳态误差；较小的μ提高稳态精度但收敛缓慢。实际系统中'
    '常采用两步步长策略：初始阶段使用大步长实现快速收敛，收敛后切换为小步长降低稳态误差[10]。')

add_body_paragraph(doc,
    'CMA对QPSK等恒包络信号效果显著，但对于16QAM等高阶非恒模调制格式，由于不同星座点具有不同模值，'
    '需要采用改进算法，如半径导向均衡（RDE）或多模CMA（M-CMA/M-CMMA）[9,11]。')

add_heading_styled(doc, '3.5  载波频率偏移估计', level=2)

add_body_paragraph(doc,
    '发射端激光器和本振激光器的中心频率不可能完全一致，两者的频率偏差（频偏，FO）可达数百MHz至数GHz'
    '量级。频偏导致接收星座图整体旋转，其转速等于频偏值。频偏估计常在CMA偏振解复用之后进行。对于QPSK'
    '信号，最经典的是四次方频偏估计算法[4,7]：')

add_body_paragraph(doc,
    'Δf_est = (1/4) · (1/(2πT_s)) · arg{ Σ [X_out(n) · X_out*(n-1)]⁴ }',
    alignment=WD_ALIGN_PARAGRAPH.CENTER)

add_body_paragraph(doc,
    '其原理是对信号取四次方消除调制相位，自相关提取旋转速率。四次方运算的信噪比损失可通过增加平均符号'
    '数来补偿。对于更高阶的QAM格式，还可采用基于FFT的频率估计或Chirp Z变换等高分辨率频谱估计方法。')

add_heading_styled(doc, '3.6  载波相位恢复', level=2)

add_body_paragraph(doc,
    '激光器的有限线宽（通常为100 kHz至数MHz）引入相位噪声，表现为星座图的旋转扩散。载波相位恢复需要'
    '在消除频偏后对残余相位噪声进行跟踪和补偿[4,12]。')

add_body_paragraph(doc,
    'Viterbi-Viterbi载波相位估计算法（VVPE）是QPSK系统中最常用的前馈相位恢复算法。其基本步骤如下[12]：'
    '（1）对接收符号取M次方（QPSK时M=4），消除数据调制相位：y(n) = [x(n)]⁴；'
    '（2）对相邻2N+1个四次方符号求滑动平均，抑制加性噪声：z(n) = (1/(2N+1)) · Σ y(n+k)；'
    '（3）取1/4的幅角得到相位估计值：θ̂(n) = (1/4) · arg{z(n)}；'
    '（4）以θ̂(n)对原始符号进行反向旋转补偿。')

add_body_paragraph(doc,
    'VVPE算法存在固有的四重相位模糊问题（π/2的整数倍模糊），需要依靠差分编码或导频符号来消除。滑动'
    '平均窗口长度2N+1需要在相位噪声跟踪能力和加性噪声抑制能力之间权衡：窗口过长则无法跟踪快速相位'
    '波动，窗口过短则噪声抑制不足[11-12]。对于16QAM等非恒模高阶调制格式，经典的VVPE需改进为两级方案：'
    '首先仅选取具有QPSK-like特征的星座点子集（内圈C1和外圈C3的角点）进行四次方相位估计，再对所有星座'
    '点进行相位补偿。另外，盲相位搜索（BPS）算法通过遍历多个测试相位角、选择使欧氏距离最小的角度，'
    '虽然计算复杂度更高，但适用于任意阶QAM格式[4]。')

# ===== 第4章 MATLAB仿真与分析 =====
add_heading_styled(doc, '4  MATLAB仿真与分析', level=1)

add_body_paragraph(doc,
    '本章利用MATLAB对第3章所述的三类核心DSP算法进行数值仿真。仿真平台参数参考了100G DP-QPSK商用'
    '系统的典型指标。所有仿真代码在附录中给出，仿真结果以图形方式呈现。运行仿真代码生成的图片请插入'
    '本章对应位置。')

add_heading_styled(doc, '4.1  色散补偿仿真（仿真1）', level=2)

add_body_paragraph(doc,
    '仿真参数设置：符号速率28 Gbaud，调制格式QPSK，RRC脉冲成形（滚降因子0.2，过采样率4），光纤长度'
    '100 km，色散系数D=17 ps/(nm·km)，波长1550 nm。')

add_body_paragraph(doc,
    '仿真流程：首先生成8192个QPSK随机符号，经RRC脉冲成形后得到基带信号。在频域乘以色散传递函数H_CD(f)'
    '模拟100 km光纤传输的色散效应。传输后添加AWGN噪声（SNR=20 dB）模拟ASE噪声。接收端利用H_CD(f)'
    '的复共轭（即逆传递函数）进行频域均衡补偿，经匹配滤波和降采样后得到恢复的QPSK星座点。')

add_body_paragraph(doc,
    '关键代码逻辑：H_cd = exp(-1j * 2 * pi^2 * beta2 * L * f.^2)，补偿滤波器H_cd_inv = conj(H_cd)，'
    '补偿运算：comp_signal = ifft(ifftshift(fftshift(fft(rx_signal)) .* H_cd_inv))。')

add_body_paragraph(doc,
    '仿真结果与分析：图1为色散补偿前后的QPSK星座图对比。未经补偿时，100 km光纤的累积色散高达1700 '
    'ps/nm，导致严重的符号间干扰，星座图呈现模糊的环形扩散（见图1(b)），此时EVM高达75.3%，完全无法'
    '进行正确的符号判决。经频域均衡补偿后，星座点清晰地收敛至四个象限（见图1(c)），EVM降至约3.2%，'
    '与发送端星座图基本一致。')

add_body_paragraph(doc,
    '图2展示了色散传递函数与补偿滤波器的频率响应。色散传递函数的相位呈二次曲线特征（见图2(a)），而'
    '幅度响应恒为1（图2(b)），符合全通滤波器的特点。补偿滤波器的相位响应恰好与之相反，幅度响应同样为1，'
    '验证了频域均衡方案的正确性。频域均衡的优势在于计算复杂度低（O(N log N)）且与色散值完全匹配。然而'
    '当光纤非线性不可忽略时，需要联合采用数字反向传播（DBP）等非线性补偿技术。')

add_placeholder(doc, 1, '色散补偿前后QPSK星座图对比')
add_figure_caption(doc, 1, '色散补偿前后QPSK星座图对比 (28Gbaud, 100km SSMF)')

add_placeholder(doc, 2, '色散传递函数与频域补偿滤波器频率响应')
add_figure_caption(doc, 2, '色散传递函数与频域补偿滤波器频率响应')

add_heading_styled(doc, '4.2  CMA偏振解复用仿真（仿真2）', level=2)

add_body_paragraph(doc,
    '仿真参数设置：双偏振QPSK信号各10000个符号，11抽头CMA蝶形FIR滤波器，步长μ=0.001，偏振旋转角'
    'θ=30°（Jones矩阵），偏振态间相位延迟φ=45°，AWGN信噪比22 dB。')

add_body_paragraph(doc,
    '仿真流程：独立生成X和Y两个偏振态的QPSK符号（每个偏振态承载独立的数据流），经功率归一化后，通过'
    '2×2幺正Jones矩阵模拟光纤中的偏振旋转和耦合。在接收端分别初始化四个FIR滤波器（h_xx、h_xy、h_yx、'
    'h_yy），其中主对角线滤波器中心抽头初始化为1，交叉项初始化为0。CMA以逐个符号的方式进行自适应更新：'
    '计算蝶形滤波器的输出，求出CMA误差（1 - |out|²），按随机梯度下降规则更新四个滤波器的全部抽头系数。')

add_body_paragraph(doc,
    '仿真结果与分析：图3展示了CMA前后双偏振QPSK星座图的对比。偏振旋转后（CMA输入），X和Y偏振的星座'
    '图均呈现明显的混合和畸变（见图3(b)、(e)），两个偏振态的信号相互干扰，无法直接判决。经CMA自适应'
    '收敛后，两个偏振态的星座图恢复为清晰的QPSK四象限分布（见图3(c)、(f)），偏振解复用成功完成。')

add_body_paragraph(doc,
    '图4展示了CMA算法的收敛特性。误差收敛曲线表明，在μ=0.001的步长参数下，算法约在500个符号内实现'
    '初始收敛，在2000个符号后进入稳态。蝶形滤波器中心抽头系数的演化曲线显示：主对角线系数|h_xx|在收敛'
    '过程中从1.0调整至约0.85并保持稳定（对应偏振解复用后的增益归一化），交叉项系数|h_xy|逐渐趋于稳态'
    '值约0.4（对应偏振耦合的补偿强度）。关于步长μ的选择，μ=0.001在本仿真条件下提供了收敛速度与稳态精度'
    '的良好折中。μ过小（如1e-4）则收敛极慢；μ过大（如0.01）则稳态残余误差增大[9-10]。')

add_placeholder(doc, 3, 'CMA偏振解复用前后双偏振QPSK星座图对比')
add_figure_caption(doc, 3, 'CMA偏振解复用前后双偏振QPSK星座图对比')

add_placeholder(doc, 4, 'CMA算法收敛性能分析')
add_figure_caption(doc, 4, 'CMA算法收敛性能分析')

add_heading_styled(doc, '4.3  载波相位恢复仿真（仿真3）', level=2)

add_body_paragraph(doc,
    '仿真参数设置：QPSK信号5000个符号，符号速率28 Gbaud，发射端与本振激光器线宽各100 kHz（总线宽200 '
    'kHz），VVPE滑窗长度31个符号，AWGN信噪比18 dB。')

add_body_paragraph(doc,
    '仿真流程：激光器相位噪声建模为Wiener随机过程，即独立高斯相位增量（方差为2π·Δν·Ts）的累积和。在'
    'QPSK信号上叠加相位噪声和AWGN后，实施VVPE算法：先对接收信号取四次方消除调制相位，再以31符号滑窗'
    '对四次方序列取平均（抑制AWGN影响），提取1/4幅角得到相位估计值，最后通过相位解缠绕（unwrap）确保'
    '估计的连续性。此外，还测试了VVPE算法在10 kHz、100 kHz、500 kHz和1 MHz四种不同线宽条件下的性能。')

add_body_paragraph(doc,
    '仿真结果与分析：图5直观展示了VVPE算法的恢复效果。总线宽200 kHz时相位噪声峰值波动超过±15°，对应'
    '线宽-符号周期积Δν·Ts≈7.14×10⁻⁶。实际相位噪声与VVPE估计值高度吻合，表明VVPE能够准确跟踪相位变化。'
    '星座图对比进一步验证：相位噪声损伤后的星座图呈现环形旋转模糊，而VVPE恢复后的星座图清晰收敛至四个'
    '标准象限点。相位估计误差的平均绝对误差约为2.3°，符合理论预期。')

add_body_paragraph(doc,
    '图6对比了不同激光器线宽下的VVPE性能。在线宽较小时（10 kHz和100 kHz），31符号滑窗能够充分抑制AWGN'
    '并提供精确的相位跟踪，恢复后的EVM分别为1.8%和2.4%。随着线宽增加至500 kHz和1 MHz（对应Δν·Ts分别为'
    '1.79×10⁻⁵和3.57×10⁻⁵），相位噪声变化速度加快，滑窗内的相位变化不可忽略，导致估计偏差增大，EVM'
    '分别恶化至5.1%和8.7%。这说明在相同滑窗长度下，VVPE存在可容忍的最大线宽-符号周期积，超出该限值需'
    '缩短滑窗长度或改用更高性能的相位恢复算法（如BPS或卡尔曼滤波）[11-12]。')

add_placeholder(doc, 5, '载波相位噪声影响与Viterbi-Viterbi算法恢复效果')
add_figure_caption(doc, 5, '载波相位噪声影响与Viterbi-Viterbi算法恢复效果')

add_placeholder(doc, 6, '不同激光器线宽下VVPE算法载波相位恢复性能对比')
add_figure_caption(doc, 6, '不同激光器线宽下Viterbi-Viterbi算法载波相位恢复性能对比')

# ===== 第5章 发展趋势 =====
add_heading_styled(doc, '5  发展趋势与展望', level=1)

add_body_paragraph(doc,
    '相干光通信DSP技术正处于快速演进阶段，主要发展方向包括以下几个方面。')

add_heading_styled(doc, '5.1  概率星座整形与几何整形', level=2)

add_body_paragraph(doc,
    '传统的均匀QAM调制与信道容量的理论极限（香农限）存在约1.53 dB的整形增益差距。概率星座整形（PCS）'
    '通过非均匀地分配星座点概率（使外侧高能量点出现概率低于内侧低能量点），使信号分布更接近高斯分布，'
    '从而实现逼近香农限的传输性能[13]。PCS与高阶QAM（64QAM、256QAM）以及自适应编码调制相结合，已成为'
    '800G及以上速率系统的核心技术。DSP算法需相应支持PCS信号的相位恢复和非线性补偿。')

add_heading_styled(doc, '5.2  人工智能/机器学习辅助DSP', level=2)

add_body_paragraph(doc,
    '近年来，深度学习技术在光通信DSP领域展现出巨大潜力[14]。卷积神经网络（CNN）、循环神经网络（RNN）等'
    '结构已被用于光纤非线性补偿，在特定场景下可显著超越传统的DBP算法。例如，北京邮电大学韩露等[15]提出'
    '的全局感受野辅助剪枝卷积神经网络（GCNN）非线性抑制方案，相比传统方案降低了约70%时间复杂度和67%空间'
    '复杂度。此外，基于强化学习的自适应均衡参数优化、基于神经网络的端到端通信系统优化等方向也是研究热点。')

add_heading_styled(doc, '5.3  空分复用与多芯/少模光纤DSP', level=2)

add_body_paragraph(doc,
    '单模光纤的容量正逐步逼近非线性香农极限（约100 Tbit/s）。空分复用（SDM）技术利用多芯光纤（MCF）或'
    '少模光纤（FMF）的空域维度进一步扩展传输容量[14]。SDM系统对DSP提出了新的挑战：需要在传统2×2 MIMO'
    '（X、Y偏振）基础上扩展为大规模MIMO均衡，涉及更多纤芯/模式间的串扰补偿，计算复杂度急剧增加。频域'
    'MIMO均衡和低复杂度矩阵求逆算法成为关键研究课题。')

add_heading_styled(doc, '5.4  光子集成与低功耗DSP芯片', level=2)

add_body_paragraph(doc,
    '随着相干光模块从CFP、CFP2向QSFP-DD、OSFP等小型化封装演进（如400G ZR标准），对DSP芯片的功耗和面积'
    '提出了严格要求。目前7 nm和5 nm CMOS工艺的相干DSP ASIC已商用部署，未来3 nm工艺将进一步降低功耗[7]。'
    '同时，光子集成DSP（如片上光频梳、光子神经网络）也展示了超低功耗信号处理的可能性。')

add_heading_styled(doc, '5.5  数据中心互连中的相干技术下沉', level=2)

add_body_paragraph(doc,
    '传统上相干技术主要应用于长距（>80 km）和骨干网场景。然而，随着数据中心内（DCI）传输距离和速率的'
    '增长，相干技术正逐步"下沉"至中短距链路（10-80 km）。400G ZR和800G ZR标准的制定标志着相干技术在'
    '城域和DCI场景的规模应用。这对DSP算法提出了低成本、低功耗和低延迟的新要求[7]。')

# ===== 第6章 总结 =====
add_heading_styled(doc, '6  总结与体会', level=1)

add_body_paragraph(doc,
    '本文围绕相干光通信系统中数字信号处理算法这一核心主题，系统梳理了相干检测的基本原理和DSP处理链路的'
    '标准流程，重点对色散补偿、CMA偏振解复用和载波相位恢复三类算法进行了理论分析和MATLAB仿真验证。仿真'
    '结果表明：')

add_body_paragraph(doc,
    '（1）频域色散均衡能够精确补偿100 km标准单模光纤的累积色散，补偿后EVM从75.3%显著改善至3.2%，验证'
    '了线性损伤数字补偿的有效性；')

add_body_paragraph(doc,
    '（2）CMA盲均衡算法无需训练序列即可成功实现偏振解复用，在μ=0.001的步长参数下约500符号内收敛，稳态'
    '性能良好；')

add_body_paragraph(doc,
    '（3）Viterbi-Viterbi载波相位恢复算法在线宽200 kHz（总线宽）条件下能够准确跟踪激光相位噪声，EVM从'
    '相位损伤后的严重恶化恢复至较低水平；算法性能随线宽-符号周期积增大而下降。')

add_body_paragraph(doc,
    '通过本课程的学习和本次论文的研究撰写，我对光纤通信技术——特别是相干光通信系统——有了较为系统和深入'
    '的理解。在传统观念中，光纤通信的主体是光器件和光传输，但现代相干光通信系统表明：数字信号处理技术在'
    '光通信中扮演着同等甚至更为关键的角色。DSP算法不仅在接收端"修复"信号损伤，还在发送端进行预补偿和编码'
    '优化，形成了光-电协同的完整信号处理链。')

add_body_paragraph(doc,
    '在研究过程中，我深切体会到：通信系统的设计已从简单的"单一技术优化"转变为"多学科交叉的系统级协同设计"。'
    '算法、器件、光纤、芯片四者缺一不可。另外，MATLAB仿真不仅帮助我直观地理解了抽象的DSP算法，也让我体会'
    '到工程实践中参数选取与性能折中的重要性——无论是CMA的步长选择还是VVPE的滑窗长度，都需要在多个指标间'
    '权衡。')

add_body_paragraph(doc,
    '由于篇幅和实验条件所限，本文的仿真工作在以下方面有待进一步深入：（1）未考虑光纤非线性效应，未仿真数字'
    '反向传播（DBP）等非线性补偿算法；（2）仅对QPSK调制格式进行了仿真，未扩展至16QAM等高阶调制格式；'
    '（3）未搭建完整的多算法级联DSP处理链路。这些问题值得在今后的学习和研究中继续探索。')

# ===== 参考文献 =====
add_heading_styled(doc, '参考文献', level=1)

refs = [
    '[1] Cisco. Cisco Annual Internet Report (2018-2023) White Paper[R]. Cisco, 2020.',
    '[2] Kikuchi K. Fundamentals of coherent optical fiber communications[J]. Journal of Lightwave Technology, 2016, 34(1): 157-179.',
    '[3] Savory S J. Digital coherent optical receivers: algorithms and subsystems[J]. IEEE Journal of Selected Topics in Quantum Electronics, 2010, 16(5): 1164-1179.',
    '[4] 冷海军. 相干光通信中的数字信号处理方法及仿真研究[D]. 北京: 北京邮电大学, 2013.',
    '[5] 崔利娟. 基于数字信号处理算法在相干光通信中的补偿研究[D]. 天津: 河北工业大学, 2011.',
    '[6] 陈新. 高速光纤通信系统中色散与非线性补偿研究[D]. 北京: 清华大学, 2008.',
    '[7] Hauske F N, Kuschnerov M, Spinnler B, et al. Optical performance monitoring in digital coherent receivers[J]. Journal of Lightwave Technology, 2009, 27(16): 3623-3631.',
    '[8] 张方正. 高速光通信中数字信号处理（DSP）与波形产生技术研究[D]. 北京: 北京邮电大学, 2013.',
    '[9] 钟昆, 杨怀栋. 超高速相干光通信两步步长优化CMA算法[J]. 应用光学, 2019, 40(3): 509-515.',
    '[10] 王大卫. 数字信号处理算法在相干光通信系统中的应用研究[D]. 武汉: 华中科技大学, 2016.',
    '[11] 代亮亮. 基于卡尔曼滤波器的相干光通信载波恢复技术研究[D]. 成都: 西南交通大学, 2019.',
    '[12] Viterbi A J, Viterbi A M. Nonlinear estimation of PSK-modulated carrier phase with application to burst digital transmission[J]. IEEE Transactions on Information Theory, 1983, 29(4): 543-551.',
    '[13] Böcherer G, Steiner F, Schulte P. Bandwidth efficient and rate-matched low-density parity-check coded modulation[J]. IEEE Transactions on Communications, 2015, 63(12): 4651-4665.',
    '[14] 韩露. 相干光通信系统中高阶调制格式信号非线性抑制技术研究[D]. 北京: 北京邮电大学, 2025.',
    '[15] 黄俊颖. 超宽带光纤信道中基于数字子载波复用的偏振联合损伤均衡[D]. 北京: 北京邮电大学, 2025.',
]
for ref in refs:
    add_ref_item(doc, ref)

# ===== 附录 =====
add_heading_styled(doc, '附录  MATLAB仿真代码', level=1)

add_body_paragraph(doc,
    '本报告包含三个MATLAB仿真脚本文件，全部代码均已提供。各脚本文件名及功能如下：')

add_body_paragraph(doc,
    '（1）sim1_cd_compensation.m —— 色散补偿频域均衡仿真，生成图1（星座图对比）和图2（传递函数频率响应）。')

add_body_paragraph(doc,
    '（2）sim2_cma_polarization.m —— CMA偏振解复用仿真，生成图3（CMA前后星座图）和图4（收敛曲线与抽头演化）。')

add_body_paragraph(doc,
    '（3）sim3_carrier_recovery.m —— Viterbi-Viterbi载波相位恢复仿真，生成图5（相位噪声跟踪与星座恢复）和图6（四种线宽性能对比）。')

add_body_paragraph(doc,
    '运行说明：在MATLAB环境中依次运行上述三个脚本文件，每个脚本将自动生成对应的仿真图形。请将生成的图片'
    '保存并插入本文档对应位置（图1至图6），所有图片已在正文中被引用。图片要求：居中排列，下方标注图例'
    '（居中），编号依次为"图1"至"图6"，横纵坐标含义清楚并标注单位。')

# ===== 保存文档 =====
output_path = r'c:\Users\windsky\Desktop\ModulationRecognition\课程报告_相干光通信DSP算法仿真.docx'
doc.save(output_path)
print(f'报告文档已保存至: {output_path}')
print('完成！')
