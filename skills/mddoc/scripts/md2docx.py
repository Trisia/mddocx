#!/usr/bin/env python3
"""
mddoc - Markdown 转学术格式 DOCX

用法:
    python md2docx.py <input.md>          # 转换 Markdown 文件
    python md2docx.py <input.md> -o out.docx  # 指定输出路径
    python md2docx.py --text "markdown..."  # 直接转换文本

依赖:
    pip install python-docx Pillow requests mistune
"""

import argparse
import os
import re
import sys
import tempfile
import urllib.parse
from io import BytesIO

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from lxml import etree
from PIL import Image


# ============================================================
# OMML 公式工具 (LaTeX → OMML XML)
# ============================================================

# OMML 命名空间
NSM = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

def _m_qn(local):
    """构造 m 命名空间的 Clark 格式属性名，如 _m_qn('val') → '{NSM}val'"""
    return f'{{{NSM}}}{local}'

# LaTeX 希腊字母 → Unicode 映射
GREEK = {
    'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ',
    'epsilon': 'ε', 'zeta': 'ζ', 'eta': 'η', 'theta': 'θ',
    'iota': 'ι', 'kappa': 'κ', 'lambda': 'λ', 'mu': 'μ',
    'nu': 'ν', 'xi': 'ξ', 'omicron': 'ο', 'pi': 'π',
    'rho': 'ρ', 'sigma': 'σ', 'tau': 'τ', 'upsilon': 'υ',
    'phi': 'φ', 'chi': 'χ', 'psi': 'ψ', 'omega': 'ω',
    'Gamma': 'Γ', 'Delta': 'Δ', 'Theta': 'Θ', 'Lambda': 'Λ',
    'Xi': 'Ξ', 'Pi': 'Π', 'Sigma': 'Σ', 'Upsilon': 'Υ',
    'Phi': 'Φ', 'Psi': 'Ψ', 'Omega': 'Ω',
    'varepsilon': 'ϵ', 'varphi': 'ϕ', 'vartheta': 'ϑ',
}

# LaTeX 符号 → Unicode
SYMBOLS = {
    'infty': '∞', 'cdot': '·', 'times': '×', 'pm': '±', 'mp': '∓',
    'leq': '≤', 'geq': '≥', 'neq': '≠', 'approx': '≈', 'equiv': '≡',
    'propto': '∝', 'sim': '∼', 'simeq': '≃',
    'partial': '∂', 'nabla': '∇', 'forall': '∀', 'exists': '∃',
    'in': '∈', 'notin': '∉', 'subset': '⊂', 'supset': '⊃',
    'subseteq': '⊆', 'supseteq': '⊇', 'cap': '∩', 'cup': '∪',
    'rightarrow': '→', 'leftarrow': '←', 'Rightarrow': '⇒', 'Leftarrow': '⇐',
    'leftrightarrow': '↔', 'mapsto': '↦',
    'ldots': '…', 'cdots': '⋯', 'vdots': '⋮', 'ddots': '⋱',
    'circ': '∘', 'bullet': '•', 'oplus': '⊕', 'otimes': '⊗',
    'angle': '∠', 'triangle': '△', 'square': '□',
    'mid': '|', 'parallel': '∥', 'perp': '⊥',
    'aleph': 'ℵ', 'hbar': 'ℏ', 'imath': 'ı', 'jmath': 'ȷ',
    'ell': 'ℓ', 'wp': '℘', 'Re': 'ℜ', 'Im': 'ℑ',
    'prime': '′', 'emptyset': '∅', 'varnothing': '∅',
    'langle': '⟨', 'rangle': '⟩', 'lceil': '⌈', 'rceil': '⌉',
    'lfloor': '⌊', 'rfloor': '⌋',
}

# 需要特殊处理的函数名
FUNCTIONS = {
    'sin', 'cos', 'tan', 'cot', 'sec', 'csc',
    'arcsin', 'arccos', 'arctan',
    'sinh', 'cosh', 'tanh', 'coth',
    'log', 'lg', 'ln', 'exp',
    'lim', 'max', 'min', 'sup', 'inf',
    'det', 'dim', 'ker', 'gcd', 'deg', 'hom',
    'arg', 'Pr',
}

BIG_OPS = {
    'sum': '∑', 'prod': '∏', 'coprod': '∐',
    'int': '∫', 'iint': '∬', 'iiint': '∭', 'oint': '∮',
    'bigcup': '⋃', 'bigcap': '⋂', 'bigvee': '⋁', 'bigwedge': '⋀',
    'bigoplus': '⨁', 'bigotimes': '⨂', 'bigodot': '⨀',
}

ACCENTS = {
    'hat': '̂', 'bar': '̅', 'vec': '⃗', 'dot': '̇',
    'ddot': '̈', 'tilde': '̃', 'breve': '̆', 'check': '̌',
    'acute': '́', 'grave': '̀',
}


def _m_elem(tag):
    """创建带 OMML 命名空间的 XML 元素"""
    return etree.Element(f'{{{NSM}}}{tag}', nsmap={'m': NSM})


def _m_run(text, italic=True):
    """创建 m:r 元素，包含 m:t 文本子元素"""
    r = _m_elem('r')
    if not italic:
        rPr = _m_elem('rPr')
        nor = _m_elem('nor')
        nor.set(_m_qn('val'), '1')
        rPr.append(nor)
        r.append(rPr)
    t_elem = _m_elem('t')
    t_elem.set(qn('xml:space'), 'preserve')
    t_elem.text = text
    r.append(t_elem)
    return r


def _m_text_run(text):
    """创建非斜体文本 run（用于 \text{}、函数名等）"""
    return _m_run(text, italic=False)


def _m_wrap_d(children):
    """将子元素列表包装在 m:d（分隔符）元素中"""
    d = _m_elem('d')
    for c in children:
        if isinstance(c, str):
            d.append(_m_run(c))
        elif c is not None:
            d.append(c)
    return d


class _LatexParser:
    """LaTeX 数学公式递归下降解析器，生成 OMML XML"""

    def __init__(self, latex):
        self.s = latex
        self.pos = 0

    def peek(self):
        if self.pos < len(self.s):
            return self.s[self.pos]
        return None

    def consume(self):
        c = self.peek()
        if c is not None:
            self.pos += 1
        return c

    def expect(self, expected):
        if self.peek() != expected:
            return False
        self.consume()
        return True

    def skip_spaces(self):
        while self.peek() in (' ', '\t', '\n'):
            self.consume()

    def read_name(self):
        """读取反斜杠后的命令名，如 \frac → 'frac'"""
        name = []
        while self.peek() and self.peek().isalpha():
            name.append(self.consume())
        return ''.join(name)

    def read_required_arg(self):
        """读取花括号中的必选参数 { ... }，注意嵌套"""
        self.skip_spaces()
        if not self.expect('{'):
            return None
        depth = 1
        buf = []
        while self.pos < len(self.s) and depth > 0:
            c = self.consume()
            if c == '{':
                depth += 1
            elif c == '}':
                depth -= 1
                if depth == 0:
                    break
            buf.append(c)
        return ''.join(buf)

    def parse_expr(self):
        """解析表达式：term*"""
        children = []
        while self.pos < len(self.s):
            c = self.peek()
            if c is None or c == '}':
                break
            if c == '&':  # 矩阵列分隔
                self.consume()
                children.append(_m_run('&'))
                continue
            if c == '\\\\':  # 矩阵行分隔
                self.consume()
                self.consume()
                children.append(_m_run('\\\\'))
                continue
            term = self.parse_term()
            if term is not None:
                children.append(term)
            if self.peek() == '}':
                break
        return children

    def parse_term(self):
        """解析单个 term：atom 后可选 sub/sup"""
        c = self.peek()
        if c is None or c == '}':
            return None

        # 上标 ^
        if c == '^':
            self.consume()
            sup_elem = self.parse_atom()
            if sup_elem is not None:
                sSup = _m_elem('sSup')
                sSup_e = _m_elem('e')
                sSup_e.append(_m_run(''))  # 空基数
                sSup.append(sSup_e)
                sSup_sup = _m_elem('sup')
                sSup_sup.append(sup_elem if not isinstance(sup_elem, str) else _m_run(sup_elem))
                sSup.append(sSup_sup)
                return sSup  # Don't try to attach sub/sup further
            else:
                return _m_run('^')

        # 下标 _
        if c == '_':
            self.consume()
            sub_elem = self.parse_atom()
            if sub_elem is not None:
                sSub = _m_elem('sSub')
                sSub_e = _m_elem('e')
                sSub_e.append(_m_run(''))
                sSub.append(sSub_e)
                sSub_sub = _m_elem('sub')
                sSub_sub.append(sub_elem if not isinstance(sub_elem, str) else _m_run(sub_elem))
                sSub.append(sSub_sub)
                return sSub
            else:
                return _m_run('_')

        # atom
        atom = self.parse_atom()
        if atom is None:
            return None

        # 检查后续的 sub/sup
        c = self.peek()

        # _ 后接 ^ (如 x_i^2 或 \sum_{i=1}^{n})
        if c == '_':
            self.consume()
            sub_arg = self.parse_atom()
            sup_arg = None
            if self.peek() == '^':
                self.consume()
                sup_arg = self.parse_atom()

            if sup_arg is not None:
                sSubSup = _m_elem('sSubSup')
                e = _m_elem('e')
                e.append(atom if not isinstance(atom, str) else _m_run(atom))
                sSubSup.append(e)
                s = _m_elem('sub')
                s.append(sub_arg if not isinstance(sub_arg, str) else _m_run(sub_arg))
                sSubSup.append(s)
                sup = _m_elem('sup')
                sup.append(sup_arg if not isinstance(sup_arg, str) else _m_run(sup_arg))
                sSubSup.append(sup)
                return sSubSup
            else:
                sSub = _m_elem('sSub')
                e = _m_elem('e')
                e.append(atom if not isinstance(atom, str) else _m_run(atom))
                sSub.append(e)
                s = _m_elem('sub')
                s.append(sub_arg if not isinstance(sub_arg, str) else _m_run(sub_arg))
                sSub.append(s)
                return sSub

        elif c == '^':
            self.consume()
            sup_arg = self.parse_atom()
            if sup_arg is not None:
                sSup = _m_elem('sSup')
                e = _m_elem('e')
                e.append(atom if not isinstance(atom, str) else _m_run(atom))
                sSup.append(e)
                sup = _m_elem('sup')
                sup.append(sup_arg if not isinstance(sup_arg, str) else _m_run(sup_arg))
                sSup.append(sup)
                return sSup

        return atom

    def parse_atom(self):
        """解析最小单元：字母/数字、命令、花括号组"""
        c = self.peek()
        if c is None or c == '}' or c == '&':
            return None

        # 花括号组 { ... }
        if c == '{':
            self.consume()
            children = self.parse_expr()
            self.expect('}')  # consume closing brace if present
            if len(children) == 1:
                return children[0]
            return _m_wrap_d(children)

        # LaTeX 命令 \
        if c == '\\':
            self.consume()
            if self.peek() is None:
                return _m_run('\\')

            # 检查是否为空
            if self.peek() in (' ', '\t', '\n'):
                return _m_run(' ')

            name = self.read_name()
            if not name:
                # 转义字符如 \$ \# \% \& \_ \{ \}
                esc = self.consume()
                if esc:
                    return _m_run(esc)
                return _m_run('\\')

            # n-ary 大运算符（sum, prod, int 等）
            if name in BIG_OPS:
                nary = _m_elem('nary')
                naryPr = _m_elem('naryPr')
                chr_elem = _m_elem('chr')
                chr_elem.set(_m_qn('val'), BIG_OPS[name])
                naryPr.append(chr_elem)
                limLoc = _m_elem('limLoc')
                limLoc.set(_m_qn('val'), 'undOvr')  # limits above and below
                naryPr.append(limLoc)
                nary.append(naryPr)

                # 读取下上限
                if self.peek() == '_':
                    self.consume()
                    sub_arg = self.parse_atom()
                    if sub_arg is not None:
                        sub = _m_elem('sub')
                        sub.append(sub_arg if not isinstance(sub_arg, str) else _m_run(sub_arg))
                        nary.append(sub)

                if self.peek() == '^':
                    self.consume()
                    sup_arg = self.parse_atom()
                    if sup_arg is not None:
                        sup = _m_elem('sup')
                        sup.append(sup_arg if not isinstance(sup_arg, str) else _m_run(sup_arg))
                        nary.append(sup)

                e = _m_elem('e')
                # 读取 integrand / summand
                rest = self.parse_expr()
                if rest:
                    for r in rest:
                        e.append(r if not isinstance(r, str) else _m_run(r))
                nary.append(e)
                return nary

            # 分数 \frac{num}{den}
            if name == 'frac':
                num_text = self.read_required_arg()
                den_text = self.read_required_arg()
                if num_text is not None and den_text is not None:
                    f = _m_elem('f')
                    num = _m_elem('num')
                    num_parsed = _LatexParser(num_text).parse_expr()
                    for item in num_parsed:
                        num.append(item if not isinstance(item, str) else _m_run(item))
                    f.append(num)
                    den = _m_elem('den')
                    den_parsed = _LatexParser(den_text).parse_expr()
                    for item in den_parsed:
                        den.append(item if not isinstance(item, str) else _m_run(item))
                    f.append(den)
                    return f
                return _m_run(f'\\frac{{{num_text}}}{{{den_text}}}')

            # 平方根 \sqrt{...} 或 \sqrt[n]{...}
            if name == 'sqrt':
                rad = _m_elem('rad')
                deg = None
                if self.peek() == '[':
                    self.consume()
                    deg_text = []
                    while self.peek() and self.peek() != ']':
                        deg_text.append(self.consume())
                    self.expect(']')
                    deg_text = ''.join(deg_text)
                    deg = _m_elem('deg')
                    deg.append(_m_run(deg_text))
                if deg is not None:
                    rad.append(deg)
                e = _m_elem('e')
                body = self.read_required_arg()
                if body is not None:
                    body_parsed = _LatexParser(body).parse_expr()
                    for item in body_parsed:
                        e.append(item if not isinstance(item, str) else _m_run(item))
                rad.append(e)
                return rad

            # 函数名 \sin, \cos, \log 等
            if name in FUNCTIONS:
                return _m_run(name)

            # 文本 \text{...}
            if name == 'text':
                txt = self.read_required_arg()
                return _m_text_run(txt) if txt else _m_run('\\text{}')

            # 重音符 \hat, \bar, \vec 等
            if name in ACCENTS:
                acc = _m_elem('acc')
                accPr = _m_elem('accPr')
                chr_elem = _m_elem('chr')
                chr_elem.set(_m_qn('val'), ACCENTS[name])
                accPr.append(chr_elem)
                acc.append(accPr)
                e = _m_elem('e')
                # 重音符可以跟花括号参数或单个字符
                base = self.parse_atom()
                if base is not None:
                    e.append(base if not isinstance(base, str) else _m_run(base))
                acc.append(e)
                return acc

            # 定界符 \left( \right) \left[ \right] 等
            if name == 'left':
                delim_char = self.consume()
                # 暂时用 d 包装，标记左定界符
                d = _m_elem('d')
                dPr = _m_elem('dPr')
                begChr = _m_elem('begChr')
                begChr.set(_m_qn('val'), delim_char or '(')
                dPr.append(begChr)
                d.append(dPr)
                # 读取直到 \right
                body = []
                while self.pos < len(self.s):
                    if self.peek() == '\\':
                        saved = self.pos
                        self.consume()
                        name2 = self.read_name()
                        if name2 == 'right':
                            right_delim = self.consume()
                            endChr = _m_elem('endChr')
                            endChr.set(_m_qn('val'), right_delim or ')')
                            dPr.append(endChr)
                            break
                        else:
                            # 不是 \right，回退
                            self.pos = saved
                            body.append(self.parse_term())
                    else:
                        body.append(self.parse_term())
                for b in body:
                    if b is not None:
                        d.append(b if not isinstance(b, str) else _m_run(b))
                return d

            # 花括号（literal braces via \{ \}）
            if name == '{':
                return _m_run('{')
            if name == '}':
                return _m_run('}')

            # 希腊字母
            if name in GREEK:
                return _m_run(GREEK[name])

            # 符号
            if name in SYMBOLS:
                return _m_run(SYMBOLS[name])

            # 未知命令：保留原始文本
            result = ['\\' + name]
            if self.peek() == '{':
                result.append('{' + (self.read_required_arg() or '') + '}')
            return _m_run(''.join(result))

        # 普通字符
        ch = self.consume()
        return _m_run(ch)

    def parse(self):
        """主入口：解析整个 LaTeX 字符串并返回 OMML 元素列表"""
        self.skip_spaces()
        return self.parse_expr()


def latex_to_omml(latex, display=False):
    """将 LaTeX 数学公式转换为 OMML XML 元素

    参数:
        latex: LaTeX 公式字符串（不含 $ 定界符）
        display: True=行间公式(m:oMathPara), False=行内公式(m:oMath)
    返回:
        OMML 顶级元素（m:oMath 或 m:oMathPara），可直接插入段落 XML
    """
    parser = _LatexParser(latex)
    children = parser.parse()

    omml = _m_elem('oMath')
    for c in children:
        omml.append(c if not isinstance(c, str) else _m_run(c))

    if display:
        omp = _m_elem('oMathPara')
        omp.append(omml)
        return omp

    return omml


# ============================================================
# 字体工具
# ============================================================

def set_run_font(run, cn_font, en_font="Times New Roman", size_pt=10.5, bold=False):
    """设置 run 的中英文字体、字号、加粗"""
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.name = en_font
    run.font.color.rgb = None  # 黑色
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), cn_font)
    rFonts.set(qn('w:ascii'), en_font)
    rFonts.set(qn('w:hAnsi'), en_font)


def add_empty_para(doc):
    """添加五号空行"""
    p = doc.add_paragraph()
    run = p.add_run('')
    set_run_font(run, '宋体', size_pt=10.5)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    return p


# ============================================================
# 图片工具
# ============================================================

def download_image(url):
    """下载或复制图片到临时文件，返回路径。失败返回 None。

    支持：
    - HTTP/HTTPS URL：requests 下载
    - 本地文件路径：直接复制到临时文件
    - base64 data URI：解码保存
    """
    import requests
    import shutil

    # 1. 本地文件路径（相对于工作目录的路径）
    if os.path.exists(url):
        suffix = os.path.splitext(url)[1] or '.png'
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        shutil.copy2(url, tmp.name)
        tmp.close()
        return tmp.name

    # 2. base64 data URI
    if url.startswith('data:'):
        import base64
        # 格式: data:image/png;base64,xxxx 或 data:image/png,xxxx
        header, data = url.split(',', 1)
        is_base64 = ';base64' in header
        suffix = '.png'
        mime_match = re.match(r'data:image/(\w+)', header)
        if mime_match:
            suffix = '.' + mime_match.group(1)
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(base64.b64decode(data) if is_base64 else data.encode())
        tmp.close()
        return tmp.name

    # 3. HTTP/HTTPS URL
    headers = {'User-Agent': 'Mozilla/5.0 (compatible; mddoc-converter/1.0)'}
    try:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        suffix = os.path.splitext(urllib.parse.urlparse(url).path)[1] or '.png'
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(resp.content)
        tmp.close()
        return tmp.name
    except Exception:
        return None


def calc_image_size(img_path):
    """计算图片在 docx 中的显示尺寸：8-12cm 规则，等比缩放"""
    img = Image.open(img_path)
    dpi = img.info.get('dpi', (96, 96))
    dpi_x = dpi[0] if dpi and dpi[0] > 0 else 96
    dpi_y = dpi[1] if dpi and dpi[1] > 0 else 96

    width_cm = img.width / dpi_x * 2.54
    height_cm = img.height / dpi_y * 2.54

    if width_cm > 12:
        target_w = 12.0
    elif width_cm < 8:
        target_w = 8.0
    else:
        target_w = width_cm

    ratio = target_w / width_cm
    return Cm(target_w), Cm(height_cm * ratio)


def add_placeholder_image(doc, text):
    """生成占位图片：灰色矩形 + 文字描述"""
    from PIL import ImageDraw, ImageFont
    img = Image.new('RGB', (600, 200), (220, 220, 220))
    draw = ImageDraw.Draw(img)
    draw.text((10, 80), text, fill=(80, 80, 80))
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
    img.save(tmp, format='PNG')
    tmp.close()
    return tmp.name


# ============================================================
# 表格工具
# ============================================================

def set_three_line_table(table):
    """三线表：仅顶线粗+底线粗，表头格底线细，数据行之间无线，无竖线"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)

    borders = OxmlElement('w:tblBorders')

    # 顶线：粗 1.5pt
    top = OxmlElement('w:top')
    top.set(qn('w:val'), 'single')
    top.set(qn('w:sz'), '12')
    top.set(qn('w:space'), '0')
    top.set(qn('w:color'), '000000')
    borders.append(top)

    # 底线：粗 1.5pt
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '12')
    bottom.set(qn('w:space'), '0')
    bottom.set(qn('w:color'), '000000')
    borders.append(bottom)

    # 所有内部线+竖线全部关闭（不用insideH！数据行之间无线）
    for edge in ('left', 'right', 'insideH', 'insideV'):
        elem = OxmlElement(f'w:{edge}')
        elem.set(qn('w:val'), 'none')
        elem.set(qn('w:sz'), '0')
        elem.set(qn('w:space'), '0')
        elem.set(qn('w:color'), 'auto')
        borders.append(elem)

    tblPr.append(borders)

    # 表头行每格底部单独加细线 0.75pt（cell级border，非table级insideH）
    if table.rows:
        for cell in table.rows[0].cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = OxmlElement('w:tcBorders')
            cell_bottom = OxmlElement('w:bottom')
            cell_bottom.set(qn('w:val'), 'single')
            cell_bottom.set(qn('w:sz'), '6')
            cell_bottom.set(qn('w:space'), '0')
            cell_bottom.set(qn('w:color'), '000000')
            tcBorders.append(cell_bottom)
            tcPr.append(tcBorders)


def add_page_break_before(paragraph):
    """在段落前插入分页符"""
    pPr = paragraph._element.get_or_add_pPr()
    pb = OxmlElement('w:pageBreakBefore')
    pPr.append(pb)


def set_outline_level(paragraph, level):
    """设置 outline level（XML方式，兼容所有python-docx版本）"""
    pPr = paragraph._element.get_or_add_pPr()
    ol = OxmlElement('w:outlineLvl')
    ol.set(qn('w:val'), str(level))
    pPr.append(ol)


# ============================================================
# Markdown 解析
# ============================================================

def split_inline_math(text):
    r"""分割段落文本中的行内公式 $...$，返回 segment 列表。
    每个 segment: {'type': 'text'|'math', 'content': str}
    支持 \$ 转义，不支持嵌套。
    """
    if not text:
        return [{'type': 'text', 'content': text}]
    segments = []
    # 匹配 $...$ 但不匹配 $$
    pattern = re.compile(r'(?<!\\)\$([^$]+?)(?<!\\)\$')
    last_end = 0
    for m in pattern.finditer(text):
        if m.start() > last_end:
            segments.append({'type': 'text', 'content': text[last_end:m.start()]})
        segments.append({'type': 'math', 'content': m.group(1)})
        last_end = m.end()
    if last_end < len(text):
        segments.append({'type': 'text', 'content': text[last_end:]})
    return segments if segments else [{'type': 'text', 'content': text}]


def parse_markdown(text):
    """解析 Markdown 文本，返回节点列表。
    每个节点: {'type': 'title'|'heading'|'para'|'image'|'table'|'code'|'list',
               'level': int (仅heading),
               'text': str,
               'children': [...] (仅list),
               'alt': str (仅image),
               'url': str (仅image),
               'header': list (仅table),
               'rows': list (仅table)}
    """
    import mistune

    nodes = []
    lines = text.split('\n')
    i = 0
    heading_count = 0  # 跟踪 # 标题数量，第一个是题目

    while i < len(lines):
        line = lines[i]

        # 代码块
        if line.strip().startswith('```'):
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            i += 1
            nodes.append({'type': 'code', 'text': '\n'.join(code_lines)})
            continue

        # 行间公式 $$ ... $$（同行或跨行）
        if line.strip().startswith('$$'):
            stripped = line.strip()
            # 同行闭合：$$ ... $$
            if stripped.count('$$') >= 2:
                inner = re.search(r'\$\$(.+?)\$\$', stripped)
                if inner:
                    nodes.append({'type': 'display_math', 'text': inner.group(1).strip()})
                    i += 1
                    continue
            # 跨行：$$ 独占一行开始
            math_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith('$$'):
                math_lines.append(lines[i])
                i += 1
            i += 1  # skip closing $$
            nodes.append({'type': 'display_math', 'text': '\n'.join(math_lines).strip()})
            continue

        # 标题 (# ...)
        m = re.match(r'^(#{1,6})\s+(.*)', line)
        if m:
            level = len(m.group(1))
            title_text = m.group(2).strip()
            heading_count += 1
            if level == 1 and heading_count == 1:
                nodes.append({'type': 'title', 'text': title_text, 'level': 1})
            else:
                nodes.append({'type': 'heading', 'text': title_text, 'level': level})
            i += 1
            continue

        # 图片 (![alt](url))
        m = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)', line.strip())
        if m:
            nodes.append({'type': 'image', 'alt': m.group(1), 'url': m.group(2)})
            i += 1
            continue

        # 表格（检测 | 开头的行）
        if line.strip().startswith('|'):
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i].strip())
                i += 1
            # 解析表格：跳过第二行（分隔符行）
            header = [c.strip() for c in table_lines[0].split('|')[1:-1]]
            rows = []
            for tl in table_lines[2:]:
                rows.append([c.strip() for c in tl.split('|')[1:-1]])
            nodes.append({'type': 'table', 'header': header, 'rows': rows})
            continue

        # 空行
        if not line.strip():
            nodes.append({'type': 'empty'})
            i += 1
            continue

        # 无序列表
        m = re.match(r'^(\s*)[-*+]\s+(.*)', line)
        if m:
            list_items = []
            indent = len(m.group(1))
            while i < len(lines):
                lm = re.match(r'^(\s*)[-*+]\s+(.*)', lines[i])
                if not lm:
                    break
                if abs(len(lm.group(1)) - indent) > 2 and list_items:
                    break
                list_items.append(lm.group(2).strip())
                indent = len(lm.group(1))
                i += 1
            nodes.append({'type': 'list', 'children': list_items})
            continue

        # 普通段落
        text = line.strip()
        nodes.append({'type': 'para', 'text': text, 'segments': split_inline_math(text)})
        i += 1

    return nodes


# ============================================================
# DOCX 生成
# ============================================================

def extract_title_text(title_node):
    """从题目节点提取纯标题文字（去除'第一章'等前缀）"""
    text = title_node['text']
    # 去除中文序号前缀如 "第一章 "、"第1章 "
    text = re.sub(r'^第[一二三四五六七八九十\d]+章\s*', '', text)
    return text.strip()


def generate_docx(nodes, output_path, title_text=None):
    """根据节点列表生成 docx 文件"""
    doc = Document()

    # --- 默认样式 ---
    style = doc.styles['Normal']
    style.font.size = Pt(10.5)
    style.font.name = 'Times New Roman'
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    style.paragraph_format.line_spacing = 1.3
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(0)

    # --- 节/页边距设置 ---
    section = doc.sections[0]
    # 页边距：左3cm 右2cm 上2cm 下2cm
    section.left_margin = Cm(3)
    section.right_margin = Cm(2)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    header = section.header
    header.is_linked_to_previous = False
    # 左顶格 xxxxx
    hp_left = header.paragraphs[0]
    hp_left.paragraph_format.space_after = Pt(0)
    run_l = hp_left.add_run('xxxxx')
    set_run_font(run_l, '黑体', size_pt=9)
    # 右顶格 题目
    hp_right = header.add_paragraph()
    hp_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hp_right.paragraph_format.space_before = Pt(0)
    run_r = hp_right.add_run(title_text or '未命名文档')
    set_run_font(run_r, '黑体', size_pt=9)

    # --- 页脚/页码设置 ---
    section.footer_distance = Cm(1)
    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # "第"
    r1 = fp.add_run('第')
    set_run_font(r1, '宋体', size_pt=10.5)
    # PAGE 字段
    r_page = fp.add_run()
    fc1 = OxmlElement('w:fldChar'); fc1.set(qn('w:fldCharType'), 'begin')
    r_page._r.append(fc1)
    it1 = OxmlElement('w:instrText'); it1.set(qn('xml:space'), 'preserve'); it1.text = ' PAGE '
    r_page._r.append(it1)
    fc2 = OxmlElement('w:fldChar'); fc2.set(qn('w:fldCharType'), 'end')
    r_page._r.append(fc2)
    # "页  共"
    r2 = fp.add_run('页  共')
    set_run_font(r2, '宋体', size_pt=10.5)
    # NUMPAGES 字段
    r_total = fp.add_run()
    fc3 = OxmlElement('w:fldChar'); fc3.set(qn('w:fldCharType'), 'begin')
    r_total._r.append(fc3)
    it2 = OxmlElement('w:instrText'); it2.set(qn('xml:space'), 'preserve'); it2.text = ' NUMPAGES '
    r_total._r.append(it2)
    fc4 = OxmlElement('w:fldChar'); fc4.set(qn('w:fldCharType'), 'end')
    r_total._r.append(fc4)
    # "页"
    r3 = fp.add_run('页')
    set_run_font(r3, '宋体', size_pt=10.5)

    # --- 计数器 ---
    chapter_path = [1]  # 一级标题序号
    has_chapter = False  # 是否遇到章标题（# 一级标题）
    fig_counter = {}  # key: chapter index tuple, value: counter
    tab_counter = {}
    eq_counter = {}

    def get_chapter_key():
        return tuple(chapter_path[:1])  # 只用一级序号作为章key

    def incr_fig():
        key = get_chapter_key()
        fig_counter[key] = fig_counter.get(key, 0) + 1
        return fig_counter[key]

    def incr_tab():
        key = get_chapter_key()
        tab_counter[key] = tab_counter.get(key, 0) + 1
        return tab_counter[key]

    def incr_eq():
        key = get_chapter_key()
        eq_counter[key] = eq_counter.get(key, 0) + 1
        return eq_counter[key]

    def make_fig_label():
        if has_chapter:
            return f"图{chapter_path[0]}-{incr_fig()}"
        return f"图{incr_fig()}"

    def make_tab_label():
        if has_chapter:
            return f"表{chapter_path[0]}-{incr_tab()}"
        return f"表{incr_tab()}"

    # --- 遍历节点生成内容 ---
    for node in nodes:
        t = node['type']

        if t == 'title':
            add_empty_para(doc)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(node['text'])
            set_run_font(run, '黑体', size_pt=16)
            add_empty_para(doc)

        elif t == 'heading':
            lv = node['level']
            if lv == 1:
                # 一级标题：新页 + 三号黑体居中
                has_chapter = True
                chapter_path[0] += 1
                chapter_path[1:] = [0]  # reset sub-levels

                add_empty_para(doc)
                p = doc.add_paragraph()
                add_page_break_before(p)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_outline_level(p, 1)
                run = p.add_run(node['text'])
                set_run_font(run, '黑体', size_pt=16)
                add_empty_para(doc)

            elif lv == 2:
                # 二级标题：四号黑体顶格不加粗，上下不空行
                p = doc.add_paragraph()
                set_outline_level(p, 2)
                p.paragraph_format.first_line_indent = Pt(0)
                run = p.add_run(node['text'])
                set_run_font(run, '黑体', size_pt=14, bold=False)

            elif lv >= 3:
                # 三级标题：小四宋体不加粗，首行缩进同正文，上下不空行
                p = doc.add_paragraph()
                set_outline_level(p, min(lv, 3))
                p.paragraph_format.first_line_indent = Pt(21)
                run = p.add_run(node['text'])
                set_run_font(run, '宋体', size_pt=12, bold=False)

        elif t == 'para':
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Pt(21)
            p.paragraph_format.line_spacing = 1.3
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            segments = node.get('segments', [{'type': 'text', 'content': node['text']}])
            for seg in segments:
                if seg['type'] == 'text':
                    if seg['content']:  # 跳过空字符串
                        run = p.add_run(seg['content'])
                        set_run_font(run, '宋体', size_pt=10.5)
                else:  # math — 插入 OMML 公式
                    try:
                        omml = latex_to_omml(seg['content'], display=False)
                        p._element.append(omml)
                    except Exception:
                        # 解析失败时回退到斜体文本
                        run = p.add_run(seg['content'])
                        set_run_font(run, '宋体', en_font='Times New Roman', size_pt=10.5)
                        run.font.italic = True

        elif t == 'image':
            url = node['url']
            alt = node['alt']
            img_path = download_image(url)
            if img_path is None:
                img_path = add_placeholder_image(doc, alt or url)

            w, h = calc_image_size(img_path)

            add_empty_para(doc)
            # 图片居中
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_img = p_img.add_run()
            run_img.add_picture(img_path, width=w)

            # 图题：小五宋体加粗居中
            # alt 为空时，使用图片文件名（不含扩展名）作为默认图题
            if not alt:
                alt = os.path.splitext(os.path.basename(url))[0]
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_cap = p_cap.add_run(f'{make_fig_label()} {alt}')
            set_run_font(run_cap, '宋体', size_pt=9, bold=True)
            add_empty_para(doc)

            # 清理临时图片
            try:
                os.unlink(img_path)
            except Exception:
                pass

        elif t == 'table':
            header_cells = node['header']
            rows = node['rows']
            ncols = len(header_cells)

            # 检查前一个节点是否为表题段落
            # （在 parse 阶段已处理，这里简化：表题由前一个 para 节点承担，此处检测）
            # 实际简化处理：如果有 header text，自动加表题

            # 表题
            add_empty_para(doc)
            p_tab_cap = doc.add_paragraph()
            p_tab_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run_tc = p_tab_cap.add_run(f'{make_tab_label()} {header_cells[0] if header_cells else ""}')
            set_run_font(run_tc, '宋体', size_pt=10.5, bold=True)

            # 表格
            table = doc.add_table(rows=len(rows) + 1, cols=ncols)
            table.alignment = WD_ALIGN_PARAGRAPH.CENTER
            set_three_line_table(table)

            # 表头行设为重复标题行（跨页自动出现）
            tblHeader = OxmlElement('w:tblHeader')
            trPr = table.rows[0]._tr.get_or_add_trPr()
            trPr.append(tblHeader)

            # 表头行
            for j, cell_text in enumerate(header_cells):
                cell = table.rows[0].cells[j]
                cell.paragraphs[0].clear()
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = cell.paragraphs[0].add_run(cell_text)
                set_run_font(run, '宋体', size_pt=9)
                cell.paragraphs[0].paragraph_format.first_line_indent = Pt(0)

            # 数据行
            for i, row_data in enumerate(rows):
                for j, cell_text in enumerate(row_data):
                    cell = table.rows[i + 1].cells[j]
                    cell.paragraphs[0].clear()
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
                    run = cell.paragraphs[0].add_run(cell_text)
                    set_run_font(run, '宋体', size_pt=9)
                    cell.paragraphs[0].paragraph_format.first_line_indent = Pt(0)

            add_empty_para(doc)

        elif t == 'code':
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(1)
            run = p.add_run(node['text'])
            set_run_font(run, '宋体', en_font='Courier New', size_pt=10.5)

        elif t == 'display_math':
            # 行间公式：上下各空一行，公式居中，编号右对齐
            add_empty_para(doc)
            p = doc.add_paragraph()
            # 设置段落 tab stops：居中 + 右对齐
            pPr = p._element.get_or_add_pPr()
            tabs = OxmlElement('w:tabs')
            # 居中 tab：左边距3cm右边距2cm，可用宽度16cm，中心8cm → 4536 twips
            center_tab = OxmlElement('w:tab')
            center_tab.set(qn('w:val'), 'center')
            center_tab.set(qn('w:pos'), '4536')
            tabs.append(center_tab)
            # 右对齐 tab：16cm → 9072 twips
            right_tab = OxmlElement('w:tab')
            right_tab.set(qn('w:val'), 'right')
            right_tab.set(qn('w:pos'), '9072')
            tabs.append(right_tab)
            pPr.append(tabs)
            # tab → 居中位置
            run_t1 = p.add_run()
            tab1 = OxmlElement('w:tab')
            run_t1._r.append(tab1)
            # 插入 OMML 行内公式（m:oMath 在 w:r 同级参与排版）
            try:
                omml = latex_to_omml(node['text'], display=False)
                p._element.append(omml)
            except Exception:
                run_fb = p.add_run(node['text'])
                set_run_font(run_fb, '宋体', en_font='Times New Roman', size_pt=10.5)
                run_fb.font.italic = True
            # tab → 右对齐位置 → 公式编号
            run_t2 = p.add_run()
            tab2 = OxmlElement('w:tab')
            run_t2._r.append(tab2)
            ch_num = chapter_path[0]
            eq_num = incr_eq()
            # 编号格式：括号用宋体，数字用TNR
            run_lp = p.add_run('(')
            set_run_font(run_lp, '宋体', en_font='宋体', size_pt=10.5)
            run_lnum = p.add_run(f'{ch_num}-{eq_num}')
            set_run_font(run_lnum, '宋体', en_font='Times New Roman', size_pt=10.5)
            run_rp = p.add_run(')')
            set_run_font(run_rp, '宋体', en_font='宋体', size_pt=10.5)
            add_empty_para(doc)

        elif t == 'list':
            for idx, item in enumerate(node['children'], 1):
                p = doc.add_paragraph()
                p.paragraph_format.first_line_indent = Pt(21)
                p.paragraph_format.line_spacing = 1.3
                run = p.add_run(f'（{idx}）{item}')
                set_run_font(run, '宋体', size_pt=10.5)

        elif t == 'empty':
            pass  # 段落间不空行，空行由标题/图表handler显式添加

    doc.save(output_path)
    return output_path


# ============================================================
# 主入口
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description='mddoc - Markdown 转学术格式 DOCX',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python md2docx.py paper.md                    # 转换文件，输出到同目录
  python md2docx.py paper.md -o output.docx     # 指定输出路径
  python md2docx.py --text "# 标题\\n\\n正文内容" -o out.docx  # 直接转换文本
        """.strip())
    parser.add_argument('input', nargs='?', help='输入 Markdown 文件路径')
    parser.add_argument('-o', '--output', help='输出 DOCX 文件路径')
    parser.add_argument('--text', help='直接传入 Markdown 文本（不需要文件）')
    args = parser.parse_args()

    if not args.input and not args.text:
        parser.print_help()
        sys.exit(1)

    # 读取内容
    if args.text:
        md_content = args.text
        source_dir = os.getcwd()
    else:
        with open(args.input, 'r', encoding='utf-8') as f:
            md_content = f.read()
        source_dir = os.path.dirname(os.path.abspath(args.input))

    # 解析
    nodes = parse_markdown(md_content)

    # 提取题目（可选）
    title_node = next((n for n in nodes if n['type'] == 'title'), None)
    title_text = extract_title_text(title_node) if title_node else None

    # 确定输出路径
    if args.output:
        output_path = args.output
    elif title_text:
        safe_name = re.sub(r'[\\/*?:"<>|]', '', title_text)
        output_path = os.path.join(source_dir, f'{safe_name}.docx')
    elif args.text:
        output_path = os.path.join(source_dir, '未命名文档.docx')
    else:
        stem = os.path.splitext(os.path.basename(args.input))[0]
        output_path = os.path.join(source_dir, f'{stem}.docx')

    # 生成
    generate_docx(nodes, output_path, title_text)
    print(f'已生成: {output_path}')


if __name__ == '__main__':
    main()
