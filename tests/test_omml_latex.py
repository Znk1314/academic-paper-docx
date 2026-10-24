"""OMML 转换器测试：LaTeX 子集 → Word 原生公式。"""

from __future__ import annotations

import pytest
from conftest import localname, m
from docx import Document
from omml_latex import add_latex_to_paragraph, latex_to_omml


def only(src):
    els = latex_to_omml(src)
    assert els, "没有产生任何 OMML 元素: %r" % src
    return els


def text_of(el):
    return "".join(t.text or "" for t in el.iter(m("t")))


def test_greek_and_subscript():
    root = only(r"\theta_k")[0]
    assert localname(root) == "sSub"
    assert text_of(root.find(m("e"))) == "θ"
    assert text_of(root.find(m("sub"))) == "k"


def test_superscript_and_combined_scripts():
    root = only(r"x_i^2")[0]
    assert localname(root) == "sSubSup"
    assert text_of(root.find(m("sub"))) == "i"
    assert text_of(root.find(m("sup"))) == "2"


def test_fraction():
    root = only(r"\frac{a}{b}")[0]
    assert localname(root) == "f"
    assert text_of(root.find(m("num"))) == "a"
    assert text_of(root.find(m("den"))) == "b"


def test_sqrt_without_degree_hides_placeholder():
    root = only(r"\sqrt{x}")[0]
    assert localname(root) == "rad"
    hide = root.find(m("radPr") + "/" + m("degHide"))
    assert hide is not None and hide.get(m("val")) == "1"


def test_sqrt_with_degree():
    root = only(r"\sqrt[n]{x}")[0]
    assert localname(root) == "rad"
    assert text_of(root.find(m("deg"))) == "n"


def test_nary_sum_has_bounds_and_body():
    root = only(r"\sum_{i=1}^{n} x_i")[0]
    assert localname(root) == "nary"
    chr_el = root.find(m("naryPr") + "/" + m("chr"))
    assert chr_el.get(m("val")) == "∑"
    assert text_of(root.find(m("sub"))) == "i=1"
    assert text_of(root.find(m("sup"))) == "n"
    body = text_of(root.find(m("e")))
    assert "x" in body


def test_delimiters():
    root = only(r"\left( a \right)")[0]
    assert localname(root) == "d"
    assert root.find(m("dPr") + "/" + m("begChr")).get(m("val")) == "("
    assert root.find(m("dPr") + "/" + m("endChr")).get(m("val")) == ")"


def test_text_is_upright():
    root = only(r"\text{中文单位}")[0]
    assert localname(root) == "r"
    sty = root.find(m("rPr") + "/" + m("sty"))
    assert sty is not None and sty.get(m("val")) == "p"
    assert text_of(root) == "中文单位"


def test_accents():
    root = only(r"\hat{x}")[0]
    assert localname(root) == "acc"
    assert root.find(m("accPr") + "/" + m("chr")).get(m("val")) == "\u0302"


def test_unknown_command_does_not_raise():
    els = only(r"\foobar + 1")
    assert any(text_of(e) for e in els)


def test_unbalanced_group_does_not_raise():
    els = latex_to_omml(r"\frac{a}{b")
    assert els


def test_empty_source():
    assert latex_to_omml("") == []


def test_add_latex_to_paragraph_appends_omath():
    doc = Document()
    p = doc.add_paragraph()
    omath = add_latex_to_paragraph(p, r"E = mc^2")
    assert localname(omath) == "oMath"
    assert p._p.find(m("oMath")) is not None


@pytest.mark.parametrize("src", [
    r"\mathrm{RMSE} = \sqrt{\frac{1}{M}\sum_{i=1}^{M}(x_i - \bar{x})^2}",
    r"\theta_k = \theta_{k-1} + \omega_k \Delta t",
    r"\left[ \begin{matrix} a \end{matrix} \right]",
    r"\int_0^{T} f(t)\,dt",
])
def test_spec_style_formulas_parse(src):
    assert latex_to_omml(src)
