"""LaTeX 子集 → OMML（Word 原生可编辑公式）转换器。

只依赖 python-docx 自带的 lxml，不引入额外数学库。
支持的结构：
    x_i  x^2  x_i^2           上下标（可叠加）
    \\frac{a}{b}               分式
    \\sqrt{x}  \\sqrt[n]{x}     根式
    \\sum_{i=1}^{n} \\prod \\int \\iint \\oint   求和/积分（自动带上下限）
    \\left( ... \\right)       自动括号
    \\text{...} \\mathrm{...}   正体文字（中文也用这个）
    \\alpha \\beta \\Delta ...  希腊字母
    \\times \\le \\approx ...   运算符

用法：
    from omml_latex import add_latex_to_paragraph
    add_latex_to_paragraph(paragraph, r"\\theta_k = \\theta_{k-1} + \\omega_k \\Delta t")
"""

from __future__ import annotations

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "varepsilon": "ε", "zeta": "ζ", "eta": "η", "theta": "θ", "vartheta": "ϑ",
    "iota": "ι", "kappa": "κ", "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ",
    "omicron": "ο", "pi": "π", "rho": "ρ", "sigma": "σ", "tau": "τ",
    "upsilon": "υ", "phi": "φ", "varphi": "φ", "chi": "χ", "psi": "ψ",
    "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Xi": "Ξ",
    "Pi": "Π", "Sigma": "Σ", "Upsilon": "Υ", "Phi": "Φ", "Psi": "Ψ",
    "Omega": "Ω",
}

SYMBOLS = {
    "times": "×", "cdot": "⋅", "div": "÷", "pm": "±", "mp": "∓",
    "le": "≤", "leq": "≤", "ge": "≥", "geq": "≥", "ne": "≠", "neq": "≠",
    "approx": "≈", "equiv": "≡", "propto": "∝", "sim": "∼", "simeq": "≃",
    "ll": "≪", "gg": "≫", "to": "→", "rightarrow": "→", "leftarrow": "←",
    "leftrightarrow": "↔", "Rightarrow": "⇒", "Leftarrow": "⇐",
    "Leftrightarrow": "⇔", "infty": "∞", "partial": "∂", "nabla": "∇",
    "in": "∈", "notin": "∉", "ni": "∋", "subset": "⊂", "subseteq": "⊆",
    "supset": "⊃", "cup": "∪", "cap": "∩", "emptyset": "∅", "forall": "∀",
    "exists": "∃", "neg": "¬", "wedge": "∧", "vee": "∨", "oplus": "⊕",
    "otimes": "⊗", "perp": "⊥", "parallel": "∥", "angle": "∠",
    "triangle": "△", "circ": "∘", "ast": "∗", "star": "⋆", "bullet": "∙",
    "ldots": "…", "cdots": "⋯", "vdots": "⋮", "ddots": "⋱", "dots": "…",
    "prime": "′", "degree": "°", "percent": "%", "quad": "\u2003",
    "qquad": "\u2003\u2003", "space": "\u2002", "enspace": "\u2002",
    "lbrace": "{", "rbrace": "}", "langle": "⟨", "rangle": "⟩",
    "lceil": "⌈", "rceil": "⌉", "lfloor": "⌊", "rfloor": "⌋",
    "vert": "|", "Vert": "‖", "backslash": "\\",
    "{": "{", "}": "}", "%": "%", "_": "_", "#": "#", "&": "&",
    "|": "‖", "\\": " ", ",": "\u2009", ";": "\u2005", ":": "\u2005",
    "!": "",
}

NARY = {
    "sum": ("∑", "undOvr"), "prod": ("∏", "undOvr"), "coprod": ("∐", "undOvr"),
    "bigcup": ("⋃", "undOvr"), "bigcap": ("⋂", "undOvr"),
    "int": ("∫", "subSup"), "iint": ("∬", "subSup"), "iiint": ("∭", "subSup"),
    "oint": ("∮", "subSup"), "lim": ("lim", "undOvr"),
}

TEXT_CMDS = {"text", "mathrm", "operatorname", "mathbf", "mathsf", "mathit",
             "textrm", "mbox", "unit"}

ACCENTS = {
    "hat": "\u0302", "widehat": "\u0302", "tilde": "\u0303",
    "widetilde": "\u0303", "dot": "\u0307", "ddot": "\u0308",
    "check": "\u030c", "breve": "\u0306", "acute": "\u0301",
    "grave": "\u0300", "vec": "\u20d7",
}


def _local(el) -> str:
    return el.tag.split("}")[-1] if isinstance(el.tag, str) else ""


def _mk(tag, **attrs):
    el = OxmlElement("m:" + tag)
    for k, v in attrs.items():
        el.set(qn("m:" + k), v)
    return el


def m_r(text: str, upright: bool = False):
    """一个数学 run。默认字体样式（变量自动斜体），upright=True 强制正体。"""
    r = _mk("r")
    if upright:
        rpr = _mk("rPr")
        sty = _mk("sty")
        sty.set(qn("m:val"), "p")
        rpr.append(sty)
        r.append(rpr)
    t = _mk("t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    r.append(t)
    return r


def _take(parent, tag: str):
    """取走 parent 下名为 m:<tag> 的子元素内容列表并从父节点移除。"""
    node = parent.find(qn("m:" + tag))
    if node is None:
        return []
    children = list(node)
    parent.remove(node)
    return children


def _fill(parent, tag: str, children):
    if not children:
        return
    holder = _mk(tag)
    for c in children:
        holder.append(c)
    parent.append(holder)


def _script_node(base_els, sub=None, sup=None):
    base_els = base_els or []
    if sub and sup:
        node = _mk("sSubSup")
        _fill(node, "e", base_els)
        _fill(node, "sub", sub)
        _fill(node, "sup", sup)
        return node
    if sub:
        node = _mk("sSub")
        _fill(node, "e", base_els)
        _fill(node, "sub", sub)
        return node
    if sup:
        node = _mk("sSup")
        _fill(node, "e", base_els)
        _fill(node, "sup", sup)
        return node
    if len(base_els) == 1:
        return base_els[0]
    box = _mk("box")
    _fill(box, "e", base_els)
    return box


class _Parser:
    def __init__(self, src: str):
        self.s = src or ""
        self.n = len(self.s)
        self.i = 0

    # ---------- 基础读取 ----------
    def eof(self) -> bool:
        return self.i >= self.n

    def peek(self) -> str:
        return self.s[self.i] if self.i < self.n else ""

    def skip_ws(self):
        while not self.eof() and self.s[self.i] in " \t\r\n":
            self.i += 1

    def parse_arg_group(self):
        """解析一个 {..} 或单个 token 作为参数。"""
        self.skip_ws()
        if self.peek() == "{":
            self.i += 1
            els = self.parse_seq("}")
            if self.peek() == "}":
                self.i += 1
            return els
        return self.parse_atom(None, single=True)

    # ---------- 主循环 ----------
    def parse_seq(self, stop=None):
        nodes = []
        buf = []

        def flush():
            if buf:
                nodes.append(m_r("".join(buf)))
                buf.clear()

        while not self.eof():
            self.skip_ws()
            if self.eof():
                break
            ch = self.peek()
            if stop and ch in stop:
                break
            if ch in "_^":
                flush()
                self.i += 1
                script = self.parse_arg_group()
                last = nodes.pop() if nodes else None
                if last is None:
                    nodes.append(_script_node([], script if ch == "_" else None,
                                              script if ch == "^" else None))
                    continue
                lt = _local(last)
                if lt in ("sSub", "sSup", "sSubSup"):
                    base = _take(last, "e")
                    sub = _take(last, "sub")
                    sup = _take(last, "sup")
                    if ch == "_":
                        sub = script
                    else:
                        sup = script
                    nodes.append(_script_node(base, sub, sup))
                else:
                    nodes.append(_script_node(
                        [last], script if ch == "_" else None,
                        script if ch == "^" else None))
                continue
            if ch in "{}":
                # 裸括号（非分组）：当普通符号输出
                flush()
                self.i += 1
                nodes.append(m_r(ch))
                continue
            if ch == "\\" or ch == "~":
                flush()
                nodes.extend(self.parse_atom(stop))
                continue
            buf.append(ch)
            self.i += 1
        flush()
        return nodes

    def parse_atom(self, stop, single=False):
        self.skip_ws()
        if self.eof():
            return []
        ch = self.peek()
        if ch == "{":
            self.i += 1
            els = self.parse_seq("}")
            if self.peek() == "}":
                self.i += 1
            return els
        if ch == "\\":
            return self.parse_command(stop)
        if ch == "~":
            self.i += 1
            return [m_r("\u00a0")]
        self.i += 1
        return [m_r(ch)]

    # ---------- 命令 ----------
    def parse_command(self, stop=None):
        self.i += 1  # 跳过反斜杠
        if self.eof():
            return []
        if not self.peek().isalpha():
            ch = self.peek()
            self.i += 1
            return [m_r(SYMBOLS.get(ch, ch))]
        name = ""
        while not self.eof() and self.peek().isalpha():
            name += self.peek()
            self.i += 1

        if name == "frac":
            num = self.parse_arg_group()
            den = self.parse_arg_group()
            node = _mk("f")
            fpr = _mk("fPr")
            typ = _mk("type")
            typ.set(qn("m:val"), "bar")
            fpr.append(typ)
            node.append(fpr)
            _fill(node, "num", num)
            _fill(node, "den", den)
            return [node]

        if name == "sqrt":
            self.skip_ws()
            deg = []
            if self.peek() == "[":
                self.i += 1
                deg = self.parse_seq("]")
                if self.peek() == "]":
                    self.i += 1
            body = self.parse_arg_group()
            node = _mk("rad")
            rpr = _mk("radPr")
            if not deg:
                hide = _mk("degHide")
                hide.set(qn("m:val"), "1")
                rpr.append(hide)
            node.append(rpr)
            _fill(node, "deg", deg)
            _fill(node, "e", body)
            return [node]

        if name in ACCENTS:
            body = self.parse_arg_group()
            node = _mk("acc")
            apr = _mk("accPr")
            chr_el = _mk("chr")
            chr_el.set(qn("m:val"), ACCENTS[name])
            apr.append(chr_el)
            node.append(apr)
            _fill(node, "e", body)
            return [node]

        if name in ("overline", "bar", "underline"):
            body = self.parse_arg_group()
            node = _mk("bar")
            bpr = _mk("barPr")
            pos = _mk("pos")
            pos.set(qn("m:val"), "bot" if name == "underline" else "top")
            bpr.append(pos)
            node.append(bpr)
            _fill(node, "e", body)
            return [node]

        if name in NARY:
            return [self._parse_nary(name, stop)]

        if name == "left":
            return [self._parse_delimited()]

        if name == "right":
            # 孤立 \right：忽略
            self._read_delim()
            return []

        if name in TEXT_CMDS:
            self.skip_ws()
            if self.peek() == "{":
                self.i += 1
                start = self.i
                depth = 1
                while not self.eof() and depth:
                    c = self.peek()
                    if c == "{":
                        depth += 1
                    elif c == "}":
                        depth -= 1
                        if depth == 0:
                            break
                    self.i += 1
                raw = self.s[start:self.i]
                if self.peek() == "}":
                    self.i += 1
                raw = raw.replace("\\,", " ").replace("\\ ", " ")
                return [m_r(raw, upright=True)]
            atom = self.parse_atom(stop, single=True)
            for el in atom:
                if _local(el) == "r":
                    rpr = _mk("rPr")
                    sty = _mk("sty")
                    sty.set(qn("m:val"), "p")
                    rpr.append(sty)
                    el.insert(0, rpr)
            return atom

        if name in GREEK:
            return [m_r(GREEK[name])]

        if name in SYMBOLS:
            return [m_r(SYMBOLS[name])]

        # 未知命令：退化为正体字母串
        return [m_r(name, upright=True)]

    def _parse_nary(self, name, stop):
        ch, limloc = NARY[name]
        sub = sup = None
        self.skip_ws()
        if self.peek() == "_":
            self.i += 1
            sub = self.parse_arg_group()
        self.skip_ws()
        if self.peek() == "^":
            self.i += 1
            sup = self.parse_arg_group()
        body = self.parse_seq(stop)

        node = _mk("nary")
        npr = _mk("naryPr")
        chr_el = _mk("chr")
        chr_el.set(qn("m:val"), ch)
        npr.append(chr_el)
        if name == "lim":
            npr.append(_mk("limLoc", val="undOvr"))
        else:
            npr.append(_mk("limLoc", val=limloc))
        if not sub:
            hide = _mk("subHide")
            hide.set(qn("m:val"), "1")
            npr.append(hide)
        if not sup:
            hide = _mk("supHide")
            hide.set(qn("m:val"), "1")
            npr.append(hide)
        node.append(npr)
        _fill(node, "sub", sub)
        _fill(node, "sup", sup)
        _fill(node, "e", body)
        return node

    def _read_delim(self):
        self.skip_ws()
        if self.eof():
            return ""
        if self.peek() == "\\":
            self.i += 1
            if self.eof():
                return ""
            if not self.peek().isalpha():
                ch = self.peek()
                self.i += 1
                return {"|": "‖", "{": "{", "}": "}", "\\": " "}.get(ch, ch)
            name = ""
            while not self.eof() and self.peek().isalpha():
                name += self.peek()
                self.i += 1
            return SYMBOLS.get(name, "")
        ch = self.peek()
        self.i += 1
        return "" if ch == "." else ch

    def _parse_delimited(self):
        beg = self._read_delim()
        start = self.i
        depth = 1
        right_at = None
        while not self.eof():
            if self.s.startswith("\\left", self.i):
                depth += 1
                self.i += 5
            elif self.s.startswith("\\right", self.i):
                depth -= 1
                if depth == 0:
                    right_at = self.i
                    self.i += 6
                    break
                self.i += 6
            else:
                self.i += 1
        end = self._read_delim() if right_at is not None else ""
        inner_src = self.s[start:(right_at if right_at is not None else self.i)]

        node = _mk("d")
        dpr = _mk("dPr")
        b = _mk("begChr")
        b.set(qn("m:val"), beg or "")
        e = _mk("endChr")
        e.set(qn("m:val"), end or "")
        dpr.append(b)
        dpr.append(e)
        node.append(dpr)
        _fill(node, "e", _Parser(inner_src).parse_seq(None))
        return node


def latex_to_omml(src: str):
    """返回可放进 <m:oMath> 的元素列表。"""
    return _Parser(src).parse_seq(None)


def add_latex_to_paragraph(paragraph, src: str):
    """把 LaTeX 子集公式追加到段落里（作为 Word 原生 OMML 公式）。"""
    omath = OxmlElement("m:oMath")
    for el in latex_to_omml(src):
        omath.append(el)
    paragraph._p.append(omath)
    return omath


if __name__ == "__main__":
    from lxml import etree

    for demo in [
        r"\theta_k = \theta_{k-1} + \omega_k \Delta t",
        r"E = \frac{1}{N}\sum_{i=1}^{N}(y_i - \hat{y}_i)^2",
        r"\sigma = \sqrt{\frac{1}{n}\sum_{i=1}^{n}(x_i - \bar{x})^2}",
        r"\left( a + b \right)^2 = a^2 + 2ab + b^2",
    ]:
        els = latex_to_omml(demo)
        print(demo)
        print("  ", etree.tostring(els[0], pretty_print=False).decode()[:160], "...")
        print("   elements:", len(els))
