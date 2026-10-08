"""
生成课程设计报告 Word 文档。
格式要求:
  - 中文: 宋体, 英文/数字: Times New Roman
  - 公式: OMML (Word自带公式编辑器格式)
  - 公式序号: (1) (2) (3) ... 半角括号,依次递增
"""
from docx import Document
from docx.shared import Pt, Cm, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsmap
from docx.oxml import parse_xml, OxmlElement
import os, copy, re

# ============================================================
# 全局公式计数器
# ============================================================
_eq_counter = [0]  # 用 list 实现闭包可变

# ============================================================
# 字体辅助
# ============================================================
def set_run_font(run, cn_font='宋体', en_font='Times New Roman', size=Pt(10.5)):
    """设置 run 的中英文字体"""
    run.font.size = size
    run.font.name = en_font
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:ascii'), en_font)
    rFonts.set(qn('w:hAnsi'), en_font)
    rFonts.set(qn('w:eastAsia'), cn_font)
    rFonts.set(qn('w:cs'), en_font)

def set_paragraph_spacing(para, line_spacing=1.5):
    """设置段落行距"""
    para.paragraph_format.line_spacing = line_spacing

# ============================================================
# OMML 公式构建器
# ============================================================
MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

def _mx(tag, body=''):
    """用 parse_xml 创建带 math 命名空间的元素"""
    return parse_xml(f'<m:{tag} xmlns:m="{MATH_NS}">{body}</m:{tag}>')

def _mr(text):
    """创建 m:r 元素, 含 m:t 文本"""
    return parse_xml(f'<m:r xmlns:m="{MATH_NS}"><m:t>{text}</m:t></m:r>')

def _me(*children):
    """创建 m:e 元素, 包裹子元素"""
    e = _mx('e')
    for c in children:
        e.append(c)
    return e

def _m_sub(base_text, sub_text):
    """创建 m:sSub 下标: base_sub"""
    return parse_xml(
        f'<m:sSub xmlns:m="{MATH_NS}">'
        f'<m:e><m:r><m:t>{base_text}</m:t></m:r></m:e>'
        f'<m:sub><m:r><m:t>{sub_text}</m:t></m:r></m:sub>'
        f'</m:sSub>')

def _m_sup(base_text, sup_text):
    """创建 m:sSup 上标"""
    return parse_xml(
        f'<m:sSup xmlns:m="{MATH_NS}">'
        f'<m:e><m:r><m:t>{base_text}</m:t></m:r></m:e>'
        f'<m:sup><m:r><m:t>{sup_text}</m:t></m:r></m:sup>'
        f'</m:sSup>')

def _m_subsup(base_text, sub_text, sup_text):
    """创建 m:sSubSup 上下标"""
    return parse_xml(
        f'<m:sSubSup xmlns:m="{MATH_NS}">'
        f'<m:e><m:r><m:t>{base_text}</m:t></m:r></m:e>'
        f'<m:sub><m:r><m:t>{sub_text}</m:t></m:r></m:sub>'
        f'<m:sup><m:r><m:t>{sup_text}</m:t></m:r></m:sup>'
        f'</m:sSubSup>')

def _m_frac(num_text, den_text):
    """创建分数"""
    return parse_xml(
        f'<m:f xmlns:m="{MATH_NS}">'
        f'<m:num><m:r><m:t>{num_text}</m:t></m:r></m:num>'
        f'<m:den><m:r><m:t>{den_text}</m:t></m:r></m:den>'
        f'</m:f>')

def _m_delim(left, right, body_text):
    """创建括号 (body_text 为括号内文本)"""
    return parse_xml(
        f'<m:d xmlns:m="{MATH_NS}">'
        f'<m:dPr><m:begChr m:val="{left}"/><m:endChr m:val="{right}"/></m:dPr>'
        f'<m:e><m:r><m:t>{body_text}</m:t></m:r></m:e>'
        f'</m:d>')

def _m_bar(text):
    """创建上划线(共轭)"""
    return parse_xml(
        f'<m:acc xmlns:m="{MATH_NS}">'
        f'<m:accPr><m:chr m:val="̅"/></m:accPr>'
        f'<m:e><m:r><m:t>{text}</m:t></m:r></m:e>'
        f'</m:acc>')

def _m_abs(text):
    """绝对值"""
    return _m_delim('|', '|', text)

def _m_norm(text):
    """范数"""
    return _m_delim('‖', '‖', text)

# ---- 简短记法: 创建 OMML 元素的函数, 都返回 Element ----
def R(t): return _mr(t)
def SUB(b, s): return _m_sub(b, s)
def SUP(b, s): return _m_sup(b, s)
def SUBSUP(b, s1, s2): return _m_subsup(b, s1, s2)
def DELIM(l, r, t): return _m_delim(l, r, t)
def ABS(t): return _m_abs(t)
def NORM(t): return _m_norm(t)
def BAR(t): return _m_bar(t)
def FRAC(n, d): return _m_frac(n, d)

def add_equation(doc, omath_xml_string):
    """
    向文档插入一个 OMML 公式。
    omath_xml_string 是完整的 m:oMath XML 字符串 (不含 xmlns, 会自动补)。
    """
    _eq_counter[0] += 1
    num = _eq_counter[0]

    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(para, 1.5)

    # 构建完整的 oMath 元素
    full_xml = f'<m:oMath xmlns:m="{MATH_NS}">{omath_xml_string}</m:oMath>'
    omath = parse_xml(full_xml)
    para._p.append(omath)

    run = para.add_run(f'    ({num})')
    set_run_font(run, size=Pt(10.5))
    return para

def _r(t):
    """生成 m:r 的 XML 片段"""
    return f'<m:r><m:t>{t}</m:t></m:r>'

def _sub(b, s):
    """生成 m:sSub 的 XML 片段"""
    return f'<m:sSub><m:e>{_r(b)}</m:e><m:sub>{_r(s)}</m:sub></m:sSub>'

def _sup(b, s):
    """生成 m:sSup 的 XML 片段"""
    return f'<m:sSup><m:e>{_r(b)}</m:e><m:sup>{_r(s)}</m:sup></m:sSup>'

def _subsup(b, s1, s2):
    """生成 m:sSubSup 的 XML 片段"""
    return f'<m:sSubSup><m:e>{_r(b)}</m:e><m:sub>{_r(s1)}</m:sub><m:sup>{_r(s2)}</m:sup></m:sSubSup>'

def _frac(n, d):
    """生成 m:f 的 XML 片段"""
    return f'<m:f><m:num>{_r(n)}</m:num><m:den>{_r(d)}</m:den></m:f>'

def _delim(l, r, body):
    """生成 m:d 括号的 XML 片段"""
    return f'<m:d><m:dPr><m:begChr m:val="{l}"/><m:endChr m:val="{r}"/></m:dPr><m:e>{_r(body)}</m:e></m:d>'

def _abs(body):
    return _delim('|', '|', body)

def _norm(body):
    return _delim('‖', '‖', body)

def _bar(body):
    return f'<m:acc><m:accPr><m:chr m:val="̅"/></m:accPr><m:e>{_r(body)}</m:e></m:acc>'

# ============================================================
# 文档生成主程序
# ============================================================
def build_report():
    doc = Document()

    # ---- 页面设置 ----
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(3.18)
    section.right_margin = Cm(3.18)

    # ---- 默认样式: 中文宋体 + 英文 Times New Roman ----
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(10.5)
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.first_line_indent = Cm(0.74)
    # 设置 east-asia 字体
    rPr = style.element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:eastAsia'), '宋体')
    rFonts.set(qn('w:cs'), 'Times New Roman')
    rPr.insert(0, rFonts)

    # 设置 heading 样式字体
    for i in range(1, 4):
        h_style = doc.styles[f'Heading {i}']
        h_style.font.name = 'Times New Roman'
        hrPr = h_style.element.get_or_add_rPr()
        hrFonts = OxmlElement('w:rFonts')
        hrFonts.set(qn('w:ascii'), 'Times New Roman')
        hrFonts.set(qn('w:hAnsi'), 'Times New Roman')
        hrFonts.set(qn('w:eastAsia'), '黑体')
        h_style.element.insert(0, hrFonts)

    # ===================== 封面 =====================
    for _ in range(4):
        doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    r = p.add_run('通信工程综合设计 II')
    r.font.size = Pt(22); r.bold = True
    set_run_font(r, cn_font='黑体', size=Pt(22))

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    r = p.add_run('课程设计报告')
    r.font.size = Pt(26); r.bold = True
    set_run_font(r, cn_font='黑体', size=Pt(26))

    for _ in range(3):
        doc.add_paragraph()

    cover_lines = [
        '设计题目：基于AI的通信信号调制方式识别系统（题目1）',
        '',
        '学院（系）：通信工程系',
        '专    业：通信工程',
        '姓    名：__________',
        '学    号：__________',
        '指导教师：__________',
        '时    间：2026年7月',
    ]
    for line in cover_lines:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        if line:
            r = p.add_run(line)
            set_run_font(r, size=Pt(12))

    doc.add_page_break()

    # ===================== 辅助函数 =====================
    def heading(text, level=1):
        h = doc.add_heading(text, level=level)
        return h

    def para(text, bold=False, indent=True, size=Pt(10.5)):
        p = doc.add_paragraph()
        if not indent:
            p.paragraph_format.first_line_indent = Cm(0)
        r = p.add_run(text)
        set_run_font(r, size=size)
        r.bold = bold
        return p

    def bullet(text):
        p = doc.add_paragraph()
        r = p.add_run(text)
        set_run_font(r, size=Pt(10.5))
        return p

    def table(headers, rows):
        t = doc.add_table(rows=1 + len(rows), cols=len(headers))
        t.style = 'Light Grid Accent 1'
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, hdr in enumerate(headers):
            cell = t.rows[0].cells[i]
            cell.text = hdr
            for par in cell.paragraphs:
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
                par.paragraph_format.first_line_indent = Cm(0)
                for run in par.runs:
                    run.bold = True; run.font.size = Pt(9)
        for ri, row in enumerate(rows):
            for ci, val in enumerate(row):
                cell = t.rows[ri + 1].cells[ci]
                cell.text = str(val)
                for par in cell.paragraphs:
                    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    par.paragraph_format.first_line_indent = Cm(0)
                    for run in par.runs:
                        run.font.size = Pt(9)
        doc.add_paragraph()
        return t

    def code_block(code):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1)
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(code)
        r.font.name = 'Consolas'; r.font.size = Pt(8)

    def reset_counter():
        _eq_counter[0] = 0

    def img_placeholder(num, desc):
        """插入图片占位符"""
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(8)
        # 灰色背景占位框
        run = p.add_run(f'  [ 插入图片{num}：{desc} ]  ')
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(128, 128, 128)
        run.italic = True
        # 加边框效果
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        for side in ['top', 'left', 'bottom', 'right']:
            border = OxmlElement(f'w:{side}')
            border.set(qn('w:val'), 'dashed')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:color'), '999999')
            pBdr.append(border)
        pPr.append(pBdr)

    # ===================== 摘要 =====================
    heading('摘要', 1)
    para('本设计实现了一套基于人工智能的通信信号调制方式自动识别系统。系统通过信号生成、'
         '特征提取、AI模型训练和可视化交互四个核心模块，完成对9种常见数字调制信号'
         '（2ASK、4ASK、2FSK、4FSK、BPSK、QPSK、8PSK、16QAM、64QAM）的自动识别。'
         '系统提取了包括高阶累积量、瞬时统计特征和频谱特征在内的32维特征向量，采用支持'
         '向量机（SVM）作为核心识别算法，利用RBF核函数将特征映射到高维空间进行最优分类。'
         '创新性地引入多信噪比混合数据增强策略，使模型在不同SNR条件下均能保持高识别率。'
         '测试结果表明，SVM模型整体识别准确率达到93.21%，满足设计要求的90%以上指标。'
         '系统具备基于Tkinter的图形用户界面，支持信号可视化、实时识别和历史记录管理。')
    para('关键词：调制识别；支持向量机；特征提取；高阶累积量；信号处理', bold=True)

    # ===================== 一、需求分析 =====================
    heading('一、需求分析', 1)
    heading('1.1 项目背景', 2)
    para('在现代通信系统中，信号调制方式的自动识别是非合作通信、频谱监测、认知无线电等'
         '领域的关键技术。传统的调制识别依赖人工分析和特定解调器，效率低下且缺乏灵活性。'
         '随着人工智能技术的快速发展，利用机器学习算法自动提取信号特征并进行分类，已成为'
         '调制识别的主流方法。')

    heading('1.2 功能需求', 2)
    for i, req in enumerate([
        '信号生成需求：能够生成ASK、FSK、PSK、QAM等多种数字调制信号，支持调整信噪比（SNR）、采样率、符号速率等参数',
        '特征提取需求：自动提取信号的时域特征（幅度、相位统计量）、频域特征（功率谱、带宽）和高阶统计量（累积量）',
        'AI识别需求：采用支持向量机（SVM）算法构建分类模型，识别准确率不低于90%',
        '可视化需求：具备图形用户界面，可展示信号波形、频谱、星座图和识别结果',
        '数据管理需求：使用数据库存储信号记录、识别结果和模型性能数据',
    ], 1):
        bullet(f'{i}. {req}')

    heading('1.3 性能指标', 2)
    table(['指标', '要求值', '实测值'], [
        ['识别准确率', '≥90%', '93.21%'],
        ['支持调制类型', '≥4种', '9种'],
        ['信噪比范围', '可调', '0~30dB'],
        ['特征维度', '—', '32维'],
        ['推理时间', '—', '<0.1ms/样本'],
    ])

    heading('1.4 AI技术应用分析', 2)
    para('本系统中AI技术的应用体现在以下三个方面：')
    for i, pt in enumerate([
        '智能特征工程：突破传统人工设计特征的局限，利用高阶累积量（C20-C80）自动捕获信号的统计特性，这些累积量对高斯噪声具有天然的不敏感性，是调制识别的关键判据。将通信领域的领域知识（累积量理论）与AI方法深度融合，构建出具有物理可解释性的特征体系。',
        'SVM分类器设计：采用支持向量机（SVM）作为核心识别算法。SVM基于结构风险最小化原理和VC维理论，通过RBF核函数将32维特征向量映射到高维空间，在该空间中寻找使分类间隔最大化的最优超平面。相比于其他简单机器学习算法，SVM特别适合本系统的高维小样本特征空间，具有良好的泛化能力。',
        '数据增强策略：创新性地采用多SNR混合训练，模拟不同信道条件下的信号特征分布，提升模型在实际通信环境中的泛化能力。多SNR训练使SVM学习到对噪声强度不变的特征表示，这是单纯的固定SNR训练无法实现的。',
    ], 1):
        para(f'{i}. {pt}')

    # ===================== 二、项目设计 =====================
    heading('二、项目设计', 1)
    heading('2.1 系统总体设计', 2)
    heading('2.1.1 系统功能模块图', 3)
    para('系统由4大核心功能模块组成：信号生成模块、特征提取模块、AI识别模块（SVM）和可视化交互模块，各模块之间通过数据流进行串联。')

    table(['模块名称', '主要功能', '关键技术'], [
        ['信号生成模块', '生成9种调制信号+AWGN噪声', 'ASK/FSK/PSK/QAM调制，SNR可调'],
        ['特征提取模块', '提取32维特征向量', 'Hilbert变换、高阶累积量、频谱分析'],
        ['AI识别模块（SVM）', 'SVM分类+概率校准+超参数调优', 'RBF核函数、GridSearchCV、sigmoid校准'],
        ['可视化交互模块', '信号展示+识别结果+GUI交互', 'Tkinter + Matplotlib 4子图布局'],
        ['数据库模块', '历史记录+性能数据持久化', 'SQLite'],
    ])

    img_placeholder(1, '系统功能模块图 — 用Visio/PPT将上方功能模块表格绘制为框图')

    heading('2.1.2 技术架构', 3)
    para('系统采用分层架构设计：')
    para('表示层：Tkinter + Matplotlib，负责用户交互和数据可视化。', indent=False)
    para('业务逻辑层：信号生成、特征提取、SVM识别算法。', indent=False)
    para('数据持久层：SQLite数据库，存储历史记录和性能数据。', indent=False)

    img_placeholder(2, '系统主流程图 — 将Mermaid流程图在 mermaid.live 渲染后截图')

    heading('2.2 功能设计', 2)
    heading('2.2.1 系统主流程', 3)
    para('系统主流程为：初始化数据库 → 选择操作模式（训练模式/GUI模式）→ '
         '生成多SNR训练集（训练模式）或加载已训练SVM模型（GUI模式）→ '
         '信号生成（9种调制 × 多SNR）→ 特征提取（32维特征向量）→ 数据标准化 → '
         '训练SVM模型 + 超参数调优 → 评估并保存模型 → 记录性能到数据库 → '
         '启动GUI主界面 → 用户交互（生成信号、识别、批量测试）。')

    img_placeholder(3, '识别流程时序图 — 将Mermaid时序图在 mermaid.live 渲染后截图')

    heading('2.2.2 识别流程', 3)
    para('识别流程：用户选择参数（调制类型、SNR）→ 生成信号 → 提取特征（Hilbert变换'
         '→ 瞬时/频谱/累积量特征）→ 标准化 → SVM预测（RBF核映射 + 概率校准）→ '
         '输出识别结果（预测类型 + 置信度）→ 存储到数据库。')

    heading('2.2.3 各模块功能说明', 3)
    para('（1）信号生成模块', bold=True)
    para('负责生成9种数字调制信号，每种调制信号均采用基带脉冲成型后上变频至载波频率。'
         '信号生成后叠加加性高斯白噪声（AWGN），噪声功率根据设定的SNR计算。调制原理'
         '涵盖M-ASK（幅移键控）、M-FSK（频移键控，连续相位）、M-PSK（相移键控）和'
         'M-QAM（正交幅度调制，方形星座图）四大类。')

    para('（2）特征提取模块', bold=True)
    para('提取三类共32维特征：瞬时特征（14维，基于Hilbert变换，含Azzouz-Nandi经典特征集）、'
         '频谱特征（8维，功率谱统计量）、高阶累积量（10维，C20/C21/C40/C41/C42/C60/C61/'
         'C62/C63/C80）。高阶累积量对高斯噪声具有理论上的免疫性，是区分不同调制方式的核心判据。')

    para('（3）AI识别模块（SVM）', bold=True)
    para('采用支持向量机（SVM）作为唯一的识别算法。选择理由：基于结构风险最小化（SRM）原则，'
         '通过最大化分类间隔来最小化VC维上界，保证良好的泛化能力；RBF核函数可处理特征与标签'
         '之间的非线性关系；优化问题是凸优化，保证全局最优解；通过概率校准（sigmoid方法）'
         '输出类别置信度；通过GridSearchCV自动搜索最优超参数（C、γ、kernel）。')

    para('（4）可视化交互模块', bold=True)
    para('基于Tkinter构建，包含四个Matplotlib子图：时域波形、功率谱密度、星座图和准确率曲线。'
         '支持信号参数实时调节和批量测试。')

    img_placeholder(4, '数据库E-R图 — 将Mermaid E-R图在 mermaid.live 渲染后截图')

    heading('2.3 数据库设计', 2)
    para('系统使用SQLite数据库，包含三张核心表：signal_records（信号记录表）、'
         'recognition_results（识别结果表）、model_performance（模型性能表）。')
    para('signal_records与recognition_results之间为1对多关系。', indent=False)

    heading('2.3.1 信号记录表 (signal_records)', 3)
    table(['字段名', '类型', '说明', '示例'], [
        ['id', 'INTEGER PK', '自增主键', '1'],
        ['modulation_type', 'TEXT', '调制类型', '16QAM'],
        ['snr_db', 'REAL', '信噪比(dB)', '15.0'],
        ['sample_rate', 'REAL', '采样率(Hz)', '100000.0'],
        ['symbol_rate', 'REAL', '符号速率(Baud)', '1000.0'],
        ['n_symbols', 'INTEGER', '符号数量', '2000'],
        ['created_at', 'TEXT', '创建时间', '2026-07-01 10:30:00'],
    ])

    heading('2.3.2 识别结果表 (recognition_results)', 3)
    table(['字段名', '类型', '说明', '示例'], [
        ['id', 'INTEGER PK', '自增主键', '1'],
        ['signal_id', 'INTEGER FK', '关联信号ID', '1'],
        ['actual_type', 'TEXT', '实际类型', 'QPSK'],
        ['predicted_type', 'TEXT', '预测类型', 'QPSK'],
        ['confidence', 'REAL', '置信度', '0.985'],
        ['correct', 'INTEGER', '是否正确', '1'],
        ['model_name', 'TEXT', '模型名称', 'SVM'],
        ['recognition_time', 'TEXT', '识别时间', '2026-07-01 10:30:05'],
    ])

    heading('2.3.3 模型性能表 (model_performance)', 3)
    table(['字段名', '类型', '说明', '示例'], [
        ['id', 'INTEGER PK', '自增主键', '1'],
        ['model_name', 'TEXT', '模型名称', 'SVM'],
        ['accuracy', 'REAL', '准确率', '0.9321'],
        ['precision_score', 'REAL', '宏平均精确率', '0.9320'],
        ['recall_score', 'REAL', '宏平均召回率', '0.9320'],
        ['f1_score', 'REAL', '宏平均F1', '0.9320'],
        ['n_classes', 'INTEGER', '类别数', '9'],
        ['n_train_samples', 'INTEGER', '训练样本数', '1890'],
        ['n_test_samples', 'INTEGER', '测试样本数', '810'],
        ['train_time_sec', 'REAL', '训练耗时(s)', '0.162'],
        ['test_time_ms', 'REAL', '推理耗时(ms)', '0.038'],
    ])

    img_placeholder(5, '系统GUI主界面截图 — 运行python main.py --gui，截完整窗口(含波形/频谱/星座/曲线)')

    heading('2.4 界面设计', 2)
    para('系统主界面分为三个区域：左侧控制面板（参数设置、操作按钮、识别结果显示、'
         '模型配置）、右侧可视化区域（4个Matplotlib子图：时域波形、功率谱密度、星座图、'
         '准确率曲线）、底部历史记录区（识别历史表格，含实际类型、识别结果、置信度、'
         '正确性、模型和时间）。')

    # ===================== 三、项目实现 =====================
    heading('三、项目实现', 1)
    heading('3.1 核心算法实现', 2)
    heading('3.1.1 调制信号生成算法', 3)
    para('信号生成遵循数字通信基本模型。各类调制的数学表达式如下：')

    # ---- 公式(1): ASK ----
    add_equation(doc,
        _r('s') + _delim('(', ')', 't') + _r(' = ') +
        _sub('A', 'm') + _r('·cos') + _delim('(', ')', '2πf_ct')
    )
    para('其中 Am 为幅度电平，Am∈{2m/(M-1)-1}，m=0,1,...,M-1。此为M-ASK调制的一般表达式。')

    # ---- 公式(2): FSK ----
    add_equation(doc,
        _r('s') + _delim('(', ')', 't') + _r(' = exp[j(2π(') +
        _sub('f', 'c') + _r('+Δf·(2m-M+1))t + φ)]')
    )
    para('此为连续相位M-FSK调制的复基带表达式，保持符号间相位连续以保证频谱效率。')

    # ---- 公式(3): PSK ----
    add_equation(doc,
        _r('s') + _delim('(', ')', 't') + _r(' = exp[j(2πf_ct + 2πm/M)]')
    )
    para('此为M-PSK调制表达式，M=2对应BPSK，M=4对应QPSK，M=8对应8PSK。')

    # ---- 公式(4): QAM ----
    add_equation(doc,
        _r('s') + _delim('(', ')', 't') + _r(' = I·cos') + _delim('(', ')', '2πf_ct') +
        _r(' - Q·sin') + _delim('(', ')', '2πf_ct')
    )
    para('其中(I,Q)构成方形星座图，M=16对应16QAM，M=64对应64QAM。')

    # ---- 公式(5): AWGN ----
    add_equation(doc,
        _sub('P', 'noise') + _r(' = ') + _sub('P', 'signal') + _r(' / 10') + _sup('', 'SNR/10')
    )
    para('噪声功率根据设定SNR计算，生成复高斯噪声叠加至信号。')

    heading('3.1.2 特征提取核心算法', 3)
    para('Hilbert变换：对实信号x(t)进行Hilbert变换得到解析信号z(t)=x(t)+j·H{x(t)}，'
         '从而获取瞬时幅度a(t)=|z(t)|、瞬时相位φ(t)=arg[z(t)]和瞬时频率f(t)=(1/2π)·dφ/dt。')

    para('高阶累积量：对于零均值复平稳过程X，定义矩和累积量如下：')

    # ---- 公式(6): Mpq ----
    add_equation(doc,
        _subsup('M', 'pq', '') + _r(' = E[') + _sup('X', 'p-q') +
        _r('·(') + _sup('X', '*') + _r(')') + _sup('', 'q') + _r(']')
    )

    # ---- 公式(7): C40 ----
    add_equation(doc,
        _sub('C', '40') + _r(' = ') + _sub('M', '40') +
        _r(' - 3') + _subsup('M', '20', '2')
    )

    # ---- 公式(8): C41 ----
    add_equation(doc,
        _sub('C', '41') + _r(' = ') + _sub('M', '41') +
        _r(' - 3') + _sub('M', '20') + _sub('M', '21')
    )

    # ---- 公式(9): C42 ----
    add_equation(doc,
        _sub('C', '42') + _r(' = ') + _sub('M', '42') + _r(' - |') +
        _sub('M', '20') + _r('|') + _sup('', '2') + _r(' - 2') + _subsup('M', '21', '2')
    )

    para('关键性质：高斯过程的三阶及以上累积量恒为零，不同调制方式的累积量理论值不同。', indent=True)

    table(['调制类型', '|C40| (理论)', '|C42| (理论)'], [
        ['BPSK', '2.0', '2.0'],
        ['QPSK/8PSK', '0.0', '2.0'],
        ['16QAM', '0.68', '2.0'],
        ['64QAM', '0.62', '2.0'],
    ])

    heading('3.1.3 SVM分类算法原理', 3)
    para('支持向量机（SVM）是本系统采用的唯一核心分类算法。其基本思想是在特征空间中寻找'
         '最优分类超平面，使不同类别之间的分类间隔最大化。')

    para('SVM软间隔优化目标：', bold=True)

    # ---- 公式(10): SVM目标 ----
    add_equation(doc,
        _r('min ') + _frac('1', '2') + _norm('w') + _sup('', '2') +
        _r(' + C·Σ') + _sub('ξ', 'i')
    )
    # ---- 公式(11): SVM约束 ----
    add_equation(doc,
        _r('s.t. ') + _sub('y', 'i') + _delim('(', ')', 'w·φ(x_i)+b') +
        _r(' ≥ 1 - ') + _sub('ξ', 'i') + _r(',  ') + _sub('ξ', 'i') + _r(' ≥ 0')
    )

    para('其中w为超平面法向量，b为偏置项，ξi为松弛变量，C为惩罚系数。')

    para('RBF核函数：', bold=True)

    # ---- 公式(12): RBF ----
    add_equation(doc,
        _r('K') + _delim('(', ')', 'x_i, x_j') + _r(' = exp(-γ‖') +
        _sub('x', 'i') + _r(' - ') + _sub('x', 'j') + _r('‖') + _sup('', '2') + _r(')')
    )

    para('参数γ控制每个训练样本的影响范围。系统通过GridSearchCV在以下空间搜索最优超参数：'
         'C∈{0.1,1,10,100}，gamma∈{scale,auto,0.01,0.1}，kernel∈{rbf,poly}，'
         '共16种组合，以3折交叉验证准确率为优化目标。')

    heading('3.2 系统部署', 2)
    para('环境要求：Python 3.8+，依赖库：numpy, scipy, matplotlib, scikit-learn。', bold=True)
    para('运行方式：python main.py（训练+GUI），python main.py --train（仅训练），'
         'python main.py --gui（仅GUI）。', indent=False)

    table(['文件名', '功能说明'], [
        ['main.py', '主入口，SVM训练与GUI启动'],
        ['config.py', '系统参数配置'],
        ['signal_generator.py', '9种调制信号生成 + AWGN噪声'],
        ['feature_extraction.py', '32维特征提取'],
        ['dataset_builder.py', '数据集构建与预处理'],
        ['ai_model.py', 'SVM模型训练评估与超参数调优'],
        ['database.py', 'SQLite数据库CRUD操作'],
        ['gui_app.py', 'Tkinter可视化界面'],
    ])

    # ===================== 四、项目效果 =====================
    heading('四、项目效果', 1)

    img_placeholder(6, '不同调制信号时域波形对比 — BPSK/4FSK/16QAM三组波形子图截图并排')
    img_placeholder(7, '不同调制信号星座图对比 — BPSK/QPSK/16QAM三组星座图子图截图并排')

    heading('4.1 SVM模型性能', 2)
    para('在多SNR混合训练条件下（SNR=0,5,10,15,20,25dB），SVM模型的性能如下：')

    table(['指标', '数值'], [
        ['准确率 (Accuracy)', '93.21%'],
        ['精确率 (Precision)', '93.20%'],
        ['召回率 (Recall)', '93.20%'],
        ['F1分数', '93.20%'],
        ['训练耗时', '0.162s'],
        ['单样本推理耗时', '0.038ms'],
    ])

    heading('4.2 不同SNR下的鲁棒性分析', 2)
    table(['SNR(dB)', '0', '5', '10', '15', '20', '25', '30'],
          [['准确率', '67.8%', '85.6%', '92.2%', '95.6%', '96.7%', '96.7%', '96.7%']])

    img_placeholder(8, 'SVM准确率-vs-SNR曲线 — 运行GUI点"批量测试"，截右下角准确率曲线子图')

    para('多SNR数据增强策略使模型在SNR≥10dB时保持了92%以上的高准确率。在极低SNR（-5dB）'
         '条件下性能下降明显，这是通信系统的固有困难——信号几乎淹没在噪声中，高阶累积量等'
         '特征的估计精度严重退化。')

    heading('4.3 SVM算法选择的合理性论证', 2)
    table(['算法', '核心原理', '优势', '劣势', '调制识别适用性'], [
        ['SVM（选用）', '结构风险最小化+核方法', '泛化能力强，适合高维', '超参数需调优', '★★★★★'],
        ['决策树', '基于信息增益的递归分区', '可解释性强', '易过拟合', '★★★☆☆'],
        ['KNN', '基于距离度量的惰性学习', '无需训练', '高维距离失效', '★★☆☆☆'],
        ['随机森林', 'Bagging集成决策树', '降低过拟合', '模型较大', '★★★★☆'],
    ])

    heading('4.4 创新点分析', 2)
    for i, (title, desc) in enumerate([
        ('多SNR混合数据增强', '在0~25dB的宽SNR范围内生成训练数据，使SVM模型学习到SNR不变的特征表示。'),
        ('高阶累积量特征融合', '将2-8阶共10个累积量作为特征，利用其对高斯噪声的免疫特性。'),
        ('SVM算法深度优化', 'RBF核函数+GridSearchCV超参数搜索+sigmoid概率校准，针对性优化。'),
        ('完整工程实现', '从信号生成到模型部署的端到端流水线，分层架构便于扩展。'),
    ], 1):
        para(f'{i}. {title}：{desc}')

    # ===================== 五、总结与展望 =====================
    heading('五、总结与展望', 1)
    heading('5.1 总结', 2)
    para('本项目成功实现了一套基于SVM的通信信号调制方式识别系统，完成了从信号生成、'
         '特征提取、模型训练到可视化交互的完整开发流程。SVM模型识别准确率达到93.21%，'
         '满足并超过了课程设计要求的90%指标。')
    para('通过本项目的实施，深入理解了以下知识点：')
    for i, k in enumerate([
        '数字通信系统中ASK/FSK/PSK/QAM的调制原理和信号特征',
        '基于Hilbert变换的瞬时特征提取方法及其物理意义',
        '高阶累积量的统计特性及其在调制识别中的关键作用',
        'SVM支持向量机的数学原理：结构风险最小化、核函数方法、软间隔分类',
        'RBF核函数将特征映射到高维空间以实现非线性分类的机制',
        '多SNR数据增强对模型泛化能力的提升机制',
        'GridSearchCV超参数搜索和交叉验证的模型优化方法',
    ], 1):
        para(f'{i}. {k}')

    heading('5.2 展望', 2)
    for i, o in enumerate([
        '深度学习扩展：引入CNN直接从IQ采样数据学习特征，以SVM作为baseline进行对比',
        '实际信号适配：增加载波同步、符号定时恢复等预处理模块',
        'SVM改进：尝试自定义核函数或多核学习方法',
        '更多调制类型：扩展到OFDM、扩频信号等现代通信体制',
        '硬件部署：将SVM模型部署到FPGA或嵌入式平台实现实时识别',
    ], 1):
        para(f'{i}. {o}')

    heading('5.3 开发心得', 2)
    para('在本项目的开发过程中，深刻体会到AI技术与通信理论的结合并非简单的"调用库函数"，'
         '而是需要理解两者的基本原理并在关键环节进行创新性融合。选择SVM作为唯一算法体现'
         '了"less is more"的工程哲学——深入掌握一种算法并针对具体问题优化，比泛泛使用多种'
         '算法更有价值。通信工程领域的AI应用，核心在于找到通信领域知识与AI方法的最佳结合点。')

    # ===================== 六、附录 =====================
    heading('六、附录', 1)
    heading('附录A：大模型提示词', 2)
    prompts = [
        ('提示词1：信号生成模块设计', '请帮我设计一个Python模块，用于生成多种数字通信调制信号，包括2ASK、4ASK、2FSK、4FSK、BPSK、QPSK、8PSK、16QAM和64QAM。每种信号需要支持可调的信噪比（SNR）、采样率和符号速率。信号应叠加AWGN噪声。'),
        ('提示词2：高阶累积量特征提取', '在调制识别中，高阶累积量是重要的特征。请实现Python代码来计算复信号的2阶、4阶、6阶和8阶累积量（C20, C21, C40, C41, C42, C60, C61, C62, C63, C80），并解释为什么这些累积量对高斯噪声不敏感。'),
        ('提示词3：SVM分类器设计', '我需要使用SVM作为调制识别的分类算法。请帮我设计SVM分类器，包括RBF核函数配置、概率校准（sigmoid方法）、GridSearchCV超参数搜索。请说明选择SVM的理论依据。'),
        ('提示词4：系统架构设计', '我需要设计一个完整的调制识别系统，包含信号生成、特征提取、SVM分类和GUI可视化四个模块。使用Python + Tkinter + scikit-learn技术栈。'),
    ]
    for title, content in prompts:
        para(title, bold=True)
        para(content)

    heading('附录B：核心功能程序清单', 2)
    table(['文件名', '功能', '代码行数'], [
        ['main.py', '主入口，SVM训练与GUI启动', '~120'],
        ['config.py', '系统参数配置', '~50'],
        ['signal_generator.py', '9种调制信号生成 + AWGN噪声', '~160'],
        ['feature_extraction.py', '32维特征提取', '~300'],
        ['dataset_builder.py', '数据集构建与预处理', '~120'],
        ['ai_model.py', 'SVM模型训练与调优', '~180'],
        ['database.py', 'SQLite数据库CRUD操作', '~180'],
        ['gui_app.py', 'Tkinter可视化界面', '~350'],
    ])

    heading('附录C：部分核心代码', 2)
    para('信号生成核心代码（signal_generator.py）：', bold=True)
    code_block(
        'def generate_qam(M=16, snr_db=15, n_symbols=N_SYMBOLS):\n'
        '    order = int(np.sqrt(M))\n'
        '    I = 2 * symbols_i / (order - 1) - 1\n'
        '    Q = 2 * symbols_q / (order - 1) - 1\n'
        '    baseband = I + 1j * Q\n'
        '    shaped = _pulse_shape(baseband, SPB)\n'
        '    modulated = shaped * carrier\n'
        '    return _add_awgn(modulated, snr_db), ...'
    )

    para('高阶累积量核心代码（feature_extraction.py）：', bold=True)
    code_block(
        'def cumulant_features(signal):\n'
        '    z = hilbert(signal) if np.isrealobj(signal) else np.asarray(signal)\n'
        '    z = z - np.mean(z)\n'
        '    M20 = np.mean(z ** 2)\n'
        '    M40 = np.mean(z ** 4)\n'
        '    C40 = M40 - 3 * M20 ** 2\n'
        '    C42 = M42 - np.abs(M20) ** 2 - 2 * M21 ** 2\n'
        '    return {\'C40\': C40, \'C42\': C42, ...}'
    )

    para('SVM分类器核心代码（ai_model.py）：', bold=True)
    code_block(
        'def create_model():\n'
        '    svm = SVC(kernel=\'rbf\', C=10.0,\n'
        '              gamma=\'scale\', random_state=42)\n'
        '    return CalibratedClassifierCV(svm,\n'
        '                method=\'sigmoid\', cv=3)'
    )

    # ===================== 保存 =====================
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '课程设计报告.docx')
    doc.save(out_path)
    print(f'报告已生成: {out_path}')
    print(f'共插入 {_eq_counter[0]} 个公式')
    return out_path

if __name__ == '__main__':
    build_report()
