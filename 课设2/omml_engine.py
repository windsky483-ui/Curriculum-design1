"""
Word OMML公式引擎 —— 生成Word原生公式对象
支持：上下标、分式、希腊字母、积分、求和、括号、根号等
"""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

def m_tag(tag):
    """创建OMML命名空间下的元素"""
    return etree.SubElement

def _m(tag_name, parent=None, **attrs):
    """创建m:命名空间元素"""
    elem = OxmlElement(f'm:{tag_name}')
    for k, v in attrs.items():
        elem.set(qn(f'm:{k}'), str(v))
    if parent is not None:
        parent.append(elem)
    return elem

def _mr(parent=None):
    """创建m:r (run) 元素"""
    return _m('r', parent)

def _mt(text, parent=None):
    """创建m:t (text) 元素"""
    t = _m('t', parent)
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    return t

def _run_with_text(text, parent=None):
    """创建包含文本的m:r"""
    r = _mr(parent)
    _mt(text, r)
    return r

def _sub(arg_elem, subscript_text):
    """将元素包装为下标结构 m:sSub"""
    sSub = _m('sSub')
    # e (base)
    e = _m('e', sSub)
    e.append(arg_elem)
    # sub
    sub = _m('sub', sSub)
    _run_with_text(subscript_text, sub)
    return sSub

def _sup(arg_elem, superscript_text):
    """将元素包装为上标结构 m:sSup"""
    sSup = _m('sSup')
    e = _m('e', sSup)
    e.append(arg_elem)
    sup = _m('sup', sSup)
    _run_with_text(superscript_text, sup)
    return sSup

def _subsup(arg_elem, sub_text, sup_text):
    """将元素包装为上下标 m:sSubSup"""
    sSubSup = _m('sSubSup')
    e = _m('e', sSubSup)
    e.append(arg_elem)
    sub = _m('sub', sSubSup)
    _run_with_text(sub_text, sub)
    sup = _m('sup', sSubSup)
    _run_with_text(sup_text, sup)
    return sSubSup

def _frac(num_text, den_text):
    """分数 m:f"""
    f = _m('f')
    num = _m('num', f)
    _run_with_text(num_text, num)
    den = _m('den', f)
    _run_with_text(den_text, den)
    return f

def _group(*elements):
    """分组（用括号包裹）"""
    d = _m('d')
    dPr = _m('dPr', d)
    # Left parenthesis
    begChr = _m('begChr', dPr)
    begChr.set(qn('m:val'), '(')
    # Right parenthesis
    endChr = _m('endChr', dPr)
    endChr.set(qn('m:val'), ')')
    for elem in elements:
        e = _m('e', d)
        if isinstance(elem, str):
            _run_with_text(elem, e)
        else:
            e.append(elem)
    return d

def _bracket_square(*elements):
    """方括号"""
    d = _m('d')
    dPr = _m('dPr', d)
    begChr = _m('begChr', dPr)
    begChr.set(qn('m:val'), '[')
    endChr = _m('endChr', dPr)
    endChr.set(qn('m:val'), ']')
    for elem in elements:
        e = _m('e', d)
        if isinstance(elem, str):
            _run_with_text(elem, e)
        else:
            e.append(elem)
    return d

def _abs(*elements):
    """绝对值/模"""
    d = _m('d')
    dPr = _m('dPr', d)
    begChr = _m('begChr', dPr)
    begChr.set(qn('m:val'), '|')
    endChr = _m('endChr', dPr)
    endChr.set(qn('m:val'), '|')
    for elem in elements:
        e = _m('e', d)
        if isinstance(elem, str):
            _run_with_text(elem, e)
        else:
            e.append(elem)
    return d

def _bar(elem):
    """上划线（共轭）"""
    acc = _m('acc')
    accPr = _m('accPr', acc)
    chr_elem = _m('chr', accPr)
    chr_elem.set(qn('m:val'), '̅')  # combining overline
    e = _m('e', acc)
    if isinstance(elem, str):
        _run_with_text(elem, e)
    else:
        e.append(elem)
    return acc

def _hat(elem):
    """帽（估计值）"""
    acc = _m('acc')
    accPr = _m('accPr', acc)
    chr_elem = _m('chr', accPr)
    chr_elem.set(qn('m:val'), '̂')  # combining circumflex
    e = _m('e', acc)
    if isinstance(elem, str):
        _run_with_text(elem, e)
    else:
        e.append(elem)
    return acc

def _func(func_name, arg):
    """函数如 exp(arg)"""
    func = _m('func')
    fName = _m('fName', func)
    _run_with_text(func_name, fName)
    e = _m('e', func)
    if isinstance(arg, str):
        _run_with_text(arg, e)
    else:
        e.append(arg)
    return func

def _text_run(s):
    """纯文本run（用于公式中的普通字符）"""
    r = _mr()
    _mt(s, r)
    return r

def _greek(letter_name):
    """希腊字母"""
    greek_map = {
        'alpha': 'α', 'beta': 'β', 'gamma': 'γ',
        'delta': 'δ', 'epsilon': 'ε', 'theta': 'θ',
        'lambda': 'λ', 'mu': 'μ', 'pi': 'π',
        'sigma': 'σ', 'tau': 'τ', 'phi': 'φ',
        'omega': 'ω', 'Delta': 'Δ', 'Pi': 'Π',
        'Sigma': 'Σ', 'Omega': 'Ω', 'Phi': 'Φ',
        'Gamma': 'Γ', 'Theta': 'Θ', 'Lambda': 'Λ',
        'varepsilon': 'ε', 'varphi': 'φ',
    }
    return _text_run(greek_map.get(letter_name, letter_name))

def make_omath(*elements):
    """创建完整的m:oMath元素"""
    oMath = _m('oMath')
    for elem in elements:
        if isinstance(elem, str):
            _run_with_text(elem, oMath)
        else:
            oMath.append(elem)
    return oMath

def make_omathpara(omath):
    """创建m:oMathPara包装"""
    oMathPara = _m('oMathPara')
    oMathPara.append(omath)
    return oMathPara

# ========== 高级构建函数 ==========

def eq_italic_var(name):
    """斜体变量如 E_s, A_s, h_xx"""
    r = _mr()
    rPr = _m('rPr', r)
    # 设置斜体
    italic_elem = _m('sty', rPr)
    italic_elem.set(qn('m:val'), 'i')
    _mt(name, r)
    return r

def eq_italic_sub(var, sub):
    """斜体变量带下标: E_s → Eₛ"""
    var_r = eq_italic_var(var)
    return _sub(var_r, sub)

def eq_italic_subsup(var, sub_t, sup_t):
    """斜体变量带上下标"""
    var_r = eq_italic_var(var)
    return _subsup(var_r, sub_t, sup_t)

def eq_normal_text(s):
    """普通文本（非斜体）"""
    return _text_run(s)

# ========== 便捷构建 ==========

def eq_exp(exp_arg):
    """exp(参数) 函数"""
    return _func('exp', exp_arg)

def eq_complex_exp(omega, t, theta):
    """exp[j(ωt + θ)] 类型"""
    inner = _group(
        eq_italic_var('j'),
        eq_normal_text('('),
        eq_italic_sub(omega, 's'),
        eq_italic_var('t'),
        eq_normal_text(' + '),
        eq_italic_sub(theta, 's'),
        eq_normal_text('(t)'),
        eq_normal_text(')'),
    )
    return _func('exp', inner)

# ========== 插入Word文档 ==========

def add_omml_equation(doc, omathpara, alignment=WD_ALIGN_PARAGRAPH.CENTER):
    """向文档插入OMML公式段落"""
    p = doc.add_paragraph()
    p.alignment = alignment

    # 创建段落属性
    pPr = p._element.get_or_add_pPr()

    # 添加公式
    p._element.append(omathpara)
    return p

def set_paragraph_line_sp(p, pt_val=20):
    """设置行间距"""
    pPr = p._element.get_or_add_pPr()
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:line'), str(int(pt_val * 20)))
    spacing.set(qn('w:lineRule'), 'exact')
    # Remove existing
    for old in pPr.findall(qn('w:spacing')):
        pPr.remove(old)
    pPr.insert(0, spacing)
