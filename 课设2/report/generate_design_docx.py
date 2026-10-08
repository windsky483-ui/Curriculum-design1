"""生成设计方案 Word 文档（本科学生汇报版）
格式: 中文宋体, 英文/数字 Times New Roman
"""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

doc = Document()
section = doc.sections[0]
section.page_width = Cm(21); section.page_height = Cm(29.7)
section.top_margin = Cm(2.0); section.bottom_margin = Cm(2.0)
section.left_margin = Cm(2.5); section.right_margin = Cm(2.5)

# ---- 默认样式: 中文宋体 + 英文 Times New Roman ----
style = doc.styles['Normal']
style.font.name = 'Times New Roman'
style.font.size = Pt(11)
style.paragraph_format.line_spacing = 1.4
style.paragraph_format.first_line_indent = Cm(0.74)
rPr = style.element.get_or_add_rPr()
rFonts = OxmlElement('w:rFonts')
rFonts.set(qn('w:ascii'), 'Times New Roman')
rFonts.set(qn('w:hAnsi'), 'Times New Roman')
rFonts.set(qn('w:eastAsia'), '宋体')
rFonts.set(qn('w:cs'), 'Times New Roman')
rPr.insert(0, rFonts)

# 标题样式: 黑体
for i in range(1, 4):
    hs = doc.styles[f'Heading {i}']
    hs.font.name = 'Times New Roman'
    h_rPr = hs.element.get_or_add_rPr()
    h_rFonts = OxmlElement('w:rFonts')
    h_rFonts.set(qn('w:ascii'), 'Times New Roman')
    h_rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    h_rFonts.set(qn('w:eastAsia'), '黑体')
    hs.element.insert(0, h_rFonts)

def set_font(run, cn='宋体', en='Times New Roman', size=Pt(11)):
    """设置 run 的中英文字体"""
    run.font.name = en
    run.font.size = size
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), en)
    rFonts.set(qn('w:hAnsi'), en)
    rFonts.set(qn('w:eastAsia'), cn)
    rFonts.set(qn('w:cs'), en)
    rPr.insert(0, rFonts)

def h(text, level=1):
    doc.add_heading(text, level=level)

def p(text, bold=False, indent=True):
    para = doc.add_paragraph()
    if indent: para.paragraph_format.first_line_indent = Cm(0.74)
    else: para.paragraph_format.first_line_indent = Cm(0)
    run = para.add_run(text)
    set_font(run, size=Pt(11))
    run.bold = bold

def tbl(headers, rows):
    t = doc.add_table(rows=1+len(rows), cols=len(headers))
    t.style = 'Light Grid Accent 1'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, hdr in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = ''
        r = cell.paragraphs[0]
        r.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = r.add_run(hdr)
        set_font(run, size=Pt(9)); run.bold = True
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = t.rows[ri+1].cells[ci]
            cell.text = ''
            r = cell.paragraphs[0]
            r.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = r.add_run(str(val))
            set_font(run, size=Pt(9))
    doc.add_paragraph()

def ai_note(text):
    """AI辅助说明（灰色小字）"""
    para = doc.add_paragraph()
    para.paragraph_format.left_indent = Cm(1.5)
    para.paragraph_format.first_line_indent = Cm(0)
    run = para.add_run('💡 AI辅助说明：' + text)
    set_font(run, size=Pt(9))
    run.italic = True

# === 封面 ===
for _ in range(4):
    _p = doc.add_paragraph()
    _p.paragraph_format.first_line_indent = Cm(0)
tp = doc.add_paragraph(); tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
tp.paragraph_format.first_line_indent = Cm(0)
r = tp.add_run('基于AI的通信信号调制方式识别系统')
set_font(r, cn='黑体', size=Pt(20)); r.bold = True

tp2 = doc.add_paragraph(); tp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
tp2.paragraph_format.first_line_indent = Cm(0)
r2 = tp2.add_run('设计方案')
set_font(r2, cn='黑体', size=Pt(16))

doc.add_paragraph()
for line in ['题目：题目1 — 基于AI的通信信号调制方式识别系统',
             '姓名：__________    学号：__________',
             '日期：2026年7月']:
    ip = doc.add_paragraph(); ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ip.paragraph_format.first_line_indent = Cm(0)
    r3 = ip.add_run(line)
    set_font(r3, size=Pt(12))
doc.add_page_break()

# === 一、题目理解 ===
h('一、题目理解', 1)
p('这个题目要求做一个能自动识别通信信号调制方式的系统。简单说就是：给它一段调制过的无线电信号，它告诉你这是BPSK还是16QAM。')
p('老师要求用"简单机器学习算法"（SVM、决策树、KNN里选），不用深度学习。我选了SVM（支持向量机），后面会解释为什么。')

# === 二、关键技术词 ===
h('二、用到的主要技术词（方便快速理解）', 1)
tbl(['关键词', '大白话解释'], [
    ['调制方式', '把数字信息"装"到载波上的方法。好比寄快递：换箱子大小=ASK，换箱子颜色=FSK，换贴纸角度=PSK，同时换大小和角度=QAM'],
    ['信噪比 SNR', '信号和噪声的比值。SNR=20dB像安静教室说话，SNR=0dB像KTV里说话，噪声和信号一样大'],
    ['特征提取', '从信号里"提炼"数学指标，比如幅度波动多大、频率怎么分布。这些指标就是AI的"眼睛"'],
    ['高阶累积量', '一种统计量，最大的好处是高斯噪声对它没影响（理论上噪声的累积量=0）。这是通信原理课上学过的概念'],
    ['SVM', '支持向量机，一种分类算法。思路是在特征空间里画一条"最宽的马路"把不同类别分开'],
    ['RBF核函数', 'SVM的参数选项，能把数据"弯"到高维空间去分割，处理非线性分类问题'],
    ['GridSearchCV', '自动帮你试不同参数组合，找出最好的那个。相当于自动调参工具'],
])

# === 三、为什么选SVM ===
h('三、为什么选SVM（不选决策树和KNN）', 1)
p('指导书给了三个选项：SVM、决策树、KNN。我做了对比：')
tbl(['对比项', 'SVM', '决策树', 'KNN'], [
    ['简单说', '画"最宽分界线"', '像做选择题一层层判断', '看最近的邻居是啥'],
    ['处理32维特征', '✅ 很好（核函数帮忙）', '⚠️ 还行', '❌ 高维距离不准确'],
    ['会不会过拟合', '不容易（自带防止机制）', '容易（树太深就会）', '不太会'],
    ['处理非线性分类', '✅ RBF核可以', '⚠️ 需要很多层', '❌ 只能画直线'],
    ['需要多少数据', '中等就行', '需要比较多', '需要很多'],
])
p('结论：32维特征 + 9种调制类型 + 调制信号之间有明显非线性边界 → SVM + RBF核最合适。')
ai_note('这个对比分析是我先列出三种算法的特点，然后用AI帮我整理成表格形式，具体判断是我自己根据通信原理和机器学习的知识做出的。')

# === 四、系统架构 ===
h('四、系统怎么搭（4个模块流水线）', 1)
p('整体思路：生成信号 → 提取特征 → SVM分类 → 显示结果。下面逐个模块说。')

h('模块1：信号生成', 2)
p('仿照通信原理课上学的方法，用Python生成9种调制信号：')
p('• 2ASK / 4ASK — 改载波幅度。2ASK只有"有"和"没有"，4ASK有4档幅度')
p('• 2FSK / 4FSK — 改载波频率。用了连续相位，避免频率切换时相位突变（频谱更干净）')
p('• BPSK / QPSK / 8PSK — 改载波相位。BPSK只有0°和180°，QPSK有4个，8PSK有8个')
p('• 16QAM / 64QAM — 同时改幅度和相位，星座图是方形的，16QAM=16个星座点，64QAM=64个')
p('每个信号生成后加上AWGN噪声（加性高斯白噪声），SNR范围-5到30dB可调。')
ai_note('信号生成的Python函数是AI帮我写的框架，我根据通信原理课本核实了FSK连续相位的实现和QAM星座图映射公式。')

h('模块2：特征提取（32维）', 2)
p('从原始信号里提取32个数学特征，分成3类：')
p('① 瞬时特征（14个）：用Hilbert变换得到解析信号，算瞬时幅度、相位、频率的各种统计量。包括经典的Azzouz-Nandi特征（γ_max、σ_ap、σ_dp等），ASK幅度跳变→γ_max大，FSK频率跳变→σ_af大。还有幅度峰度（"尖不尖"）和偏度（"偏不偏"）。', bold=False)
p('② 频谱特征（8个）：做FFT得到频谱，算中心频率、带宽、对称性、峰均比等。不同调制的频谱形状确实不一样。', bold=False)
p('③ 高阶累积量（10个）——最关键的：算了C20到C80共10个值。核心原理：高斯噪声的三阶及以上累积量恒等于0。这意味着累积量值理论上只跟信号本身有关，不受噪声大小影响。不同调制理论值不同，比如BPSK的|C40|≈2，QPSK的|C40|≈0——这就成了区分的核心依据。', bold=False)
ai_note('累积量计算公式比较复杂，我先理解了"累积量对噪声免疫"这个关键性质（通信原理课上学过），然后让AI辅助写Python代码，最后核实了BPSK/QPSK理论值是否和代码输出一致。')

h('模块3：SVM识别', 2)
p('核心AI部分：32维特征 → 标准化 → SVM分类 → 输出"这是XX调制（置信度XX%）"。')
p('SVM的原理（大白话版）：在特征空间里找一条"最宽的马路"把不同类别分开。遇到分不开的数据（比如BPSK和QPSK特征可能搅在一起），用RBF核函数把数据"映射到高维空间"，高维里就能分开了。')
p('具体配置：核函数=RBF，用GridSearchCV自动搜最优C和γ（试16种组合×3折交叉验证），加了sigmoid概率校准输出置信度。')
p('训练策略：多SNR混合训练（0/5/10/15/20/25dB共6个等级），每个等级200个样本，9种调制×6×200=10,800个样本。让模型学会"不管噪声多大，核心特征不变"。')
ai_note('SVM原理我是通过课本和网上资料学的，GridSearchCV的搜索范围是让AI建议的，然后我根据调制识别特点自己确定了训练策略（多SNR混合）。')

h('模块4：GUI可视化界面', 2)
p('用Tkinter搭了一个界面，4个子图：左上=时域波形（8个符号，轮廓清晰），右上=幅度谱（红色虚线标载波位置），左下=星座图（能看见BPSK=2点、QPSK=4点、16QAM=16点），右下=准确率vs SNR曲线（含90%参考线）。')
p('按钮有"生成信号"、"识别信号"、"批量测试"，底部表格显示识别历史。数据存SQLite。')
ai_note('GUI框架是AI辅助搭建的，我调整了显示参数（8个符号避免太密、频谱显示幅度不显示dB、准确率画曲线不画柱状图）。')

# === 五、预期效果 ===
h('五、预期准确率', 1)
tbl(['SNR条件', '预期准确率', '原因'], [
    ['≥ 20dB', '96%+', '信号很干净，特征提取准，几乎不错'],
    ['15dB', '95%+', '正常通信环境，少量噪声不影响累积量'],
    ['10dB', '90%+', '刚好达到老师要求线'],
    ['5dB', '~85%', '噪声明显，部分特征开始模糊'],
    ['≤ 0dB', '~65-70%', '信号快被噪声淹了，物理限制'],
    ['整体平均', '~92%', '综合所有SNR，超过90%要求'],
])

# === 六、技术栈 ===
h('六、用到的软硬件', 1)
tbl(['项目', '具体内容'], [
    ['编程语言', 'Python 3.8+（机房有装）'],
    ['信号处理', 'NumPy + SciPy（Hilbert变换、FFT）'],
    ['机器学习', 'Scikit-learn（SVM、GridSearchCV）'],
    ['画图', 'Matplotlib'],
    ['界面', 'Tkinter（Python自带）'],
    ['数据库', 'SQLite（Python自带）'],
    ['开发工具', 'VS Code'],
    ['操作系统', 'Windows 11'],
])

# === 七、AI辅助说明 ===
h('七、AI辅助开发说明', 1)
p('老师说可以借助AI，这里如实写清楚哪些是AI帮的、哪些是自己做的。')
tbl(['部分', 'AI帮了什么', '我自己做了什么'], [
    ['信号生成代码', '帮忙写了9种调制的Python函数框架', '核实了数学公式（ASK/FSK/PSK/QAM表达式），修正了FSK连续相位实现'],
    ['特征提取代码', '帮忙写了Hilbert变换、累积量计算代码', '理解累积量对噪声免疫的原理（通信原理课上学过），核实BPSK/QPSK理论值'],
    ['SVM训练代码', '建议GridSearchCV搜索范围', '理解SVM原理（结构风险最小化、RBF核），对比后决定选SVM不用决策树/KNN'],
    ['GUI界面代码', '帮忙搭了Tkinter框架+Matplotlib布局', '调整显示参数（8符号波形、幅度谱、准确率曲线）'],
    ['设计方案文档', '帮忙整理架构描述', '所有技术判断（为什么选SVM、特征怎么设计、参数怎么选）都是自己决定的'],
])
p('最大体会：AI可以帮你写代码，但代替不了你对通信原理的理解。比如累积量特征，AI能写计算代码，但"为什么噪声对它没影响"、"BPSK的C40理论值是多少"这些问题，必须靠课上学过的知识。AI辅助的核心是：用专业知识指导方向，AI帮忙实现细节。', bold=False)

# === 八、参考 ===
h('八、参考资料', 1)
p('1. 通信原理教材 — ASK/FSK/PSK/QAM调制原理')
p('2. Nandi & Azzouz, "Algorithms for automatic modulation recognition," IEEE Trans. Commun., 1998 — 经典特征集')
p('3. Swami & Sadler, "Hierarchical digital modulation classification using cumulants," IEEE Trans. Commun., 2000 — 高阶累积量方法')
p('4. Scikit-learn官方文档 — SVM和GridSearchCV使用方法')

out = os.path.join(os.path.dirname(os.path.abspath(__file__)), '设计方案.docx')
doc.save(out)
print(f'设计方案已生成: {out}')
