# -*- coding: utf-8 -*-
"""Builders for OMML, the Office Math Markup Language that Word uses for native
equations.

All builders return a flat list of XML elements so expressions compose with +,
for example:

    V("R") + P("=") + frac(V("λ"), P("Δ") + V("λ")) + P("=") + V("K") + V("N")

* V(text) - italic run, for variables (x, N, λ, ...)
* P(text) - upright run, for digits, operators, units and function names
* frac(num, den), sub(base, sub), subsup(base, sub, sup)
* delim(inner, "[", "]"), func("sin", arg)

Word renders these as real, editable equations using Cambria Math, so they stay
crisp at any print size - unlike a screenshot of a formula.
"""
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
MATH_FONT = "Cambria Math"


def _flat(items):
    out = []
    for it in items:
        if isinstance(it, (list, tuple)):
            out.extend(_flat(it))
        else:
            out.append(it)
    return out


def _fill(parent, items):
    for el in _flat(items):
        parent.append(el)
    return parent


def mrun(text, italic=False):
    r = OxmlElement("m:r")
    rpr = OxmlElement("m:rPr")
    sty = OxmlElement("m:sty")
    sty.set(qn("m:val"), "i" if italic else "p")
    rpr.append(sty)
    r.append(rpr)
    wrpr = OxmlElement("w:rPr")
    rf = OxmlElement("w:rFonts")
    rf.set(qn("w:ascii"), MATH_FONT)
    rf.set(qn("w:hAnsi"), MATH_FONT)
    rf.set(qn("w:cs"), MATH_FONT)
    wrpr.append(rf)
    r.append(wrpr)
    t = OxmlElement("m:t")
    t.text = text
    if text != text.strip():
        t.set(XML_SPACE, "preserve")
    r.append(t)
    return [r]


def V(text):
    """Italic variable run."""
    return mrun(text, italic=True)


def P(text):
    """Upright run for numbers, operators, units and function names."""
    return mrun(text, italic=False)


def frac(num, den):
    f = OxmlElement("m:f")
    n = OxmlElement("m:num")
    _fill(n, num)
    d = OxmlElement("m:den")
    _fill(d, den)
    f.append(n)
    f.append(d)
    return [f]


def sub(base, sb):
    s = OxmlElement("m:sSub")
    e = OxmlElement("m:e")
    _fill(e, base)
    s.append(e)
    s2 = OxmlElement("m:sub")
    _fill(s2, sb)
    s.append(s2)
    return [s]


def subsup(base, sb, sp):
    s = OxmlElement("m:sSubSup")
    e = OxmlElement("m:e")
    _fill(e, base)
    s.append(e)
    s2 = OxmlElement("m:sub")
    _fill(s2, sb)
    s.append(s2)
    s3 = OxmlElement("m:sup")
    _fill(s3, sp)
    s.append(s3)
    return [s]


def delim(inner, beg="(", end=")"):
    """Auto-sized delimiters, e.g. delim([...], "[", "]")."""
    d = OxmlElement("m:d")
    dpr = OxmlElement("m:dPr")
    b = OxmlElement("m:begChr")
    b.set(qn("m:val"), beg)
    e = OxmlElement("m:endChr")
    e.set(qn("m:val"), end)
    dpr.append(b)
    dpr.append(e)
    d.append(dpr)
    el = OxmlElement("m:e")
    _fill(el, inner)
    d.append(el)
    return [d]


def func(name, arg):
    """Upright function name plus its argument, e.g. func("sin", V("θ"))."""
    f = OxmlElement("m:func")
    fn = OxmlElement("m:fName")
    fn.append(mrun(name, italic=False)[0])
    f.append(fn)
    e = OxmlElement("m:e")
    _fill(e, arg)
    f.append(e)
    return [f]


def omml(children):
    om = OxmlElement("m:oMath")
    _fill(om, children)
    return om
